/**
 * 立替回収トラッカー（GAS側）v2
 * スプレッドシートに紐付けて使う。初回だけ setup() を実行する。
 */

const SHEET_NAMES = { members: 'メンバー', advances: '立替', shares: '回収明細' };
const HEADERS = {
  members:  ['名前', '区分', '性別'],
  advances: ['ID', '日付', '内容', '総額', '割る人数(自分含む)', 'メモ', '登録日時'],
  shares:   ['ID', '立替ID', '相手', '金額', '回収済み', '回収日'],
};

/** 初回に1回だけ実行：3つのシートを作る（v1から使っている場合は列を足す） */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(SHEET_NAMES).forEach(key => {
    let sh = ss.getSheetByName(SHEET_NAMES[key]);
    if (!sh) sh = ss.insertSheet(SHEET_NAMES[key]);
    if (sh.getLastRow() === 0) {
      sh.appendRow(HEADERS[key]);
      sh.setFrozenRows(1);
    } else if (sh.getLastColumn() < HEADERS[key].length) {
      // 旧バージョンのシートに不足分の見出しを追加
      sh.getRange(1, 1, 1, HEADERS[key].length).setValues([HEADERS[key]]);
    }
  });
}

/** Webアプリとして開いたときに画面を返す */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('立替回収')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ---------- 内部で使う小道具 ----------
function sheet_(key) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAMES[key]);
  if (!sh || sh.getLastColumn() < HEADERS[key].length) { setup(); sh = ss.getSheetByName(SHEET_NAMES[key]); }
  return sh;
}
function rows_(key) {
  const values = sheet_(key).getDataRange().getValues();
  values.shift(); // 見出し行を除く
  return values.filter(r => r[0] !== '');
}
function date_(d) {
  return d instanceof Date ? Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy-MM-dd') : String(d || '');
}
function newId_() {
  return Utilities.getUuid().slice(0, 8);
}
function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try { return fn(); } finally { lock.releaseLock(); }
}
function memberRowIndex_(name) {
  const values = sheet_('members').getDataRange().getValues();
  for (let i = 1; i < values.length; i++) if (String(values[i][0]) === name) return i + 1;
  return 0;
}

// ---------- 画面から呼ばれる関数 ----------

/** 全データをまとめて返す */
function getData() {
  const members = rows_('members').map(r => ({
    name: String(r[0]), status: String(r[1] || ''), gender: String(r[2] || ''),
  }));
  const shares = rows_('shares').map(r => ({
    id: String(r[0]), advanceId: String(r[1]), person: String(r[2]),
    amount: Number(r[3]), paid: r[4] === true, paidDate: date_(r[5]),
  }));
  const advances = rows_('advances').map(r => ({
    id: String(r[0]), date: date_(r[1]), title: String(r[2]),
    total: r[3] === '' ? '' : Number(r[3]), headcount: Number(r[4]), memo: String(r[5] || ''),
    shares: shares.filter(s => s.advanceId === String(r[0])),
  })).sort((a, b) => b.date.localeCompare(a.date));
  return { members, advances };
}

/** メンバー追加　m = {name, status, gender} */
function addMember(m) {
  const name = String(m.name || '').trim();
  if (!name) throw new Error('名前を入力してください');
  return withLock_(() => {
    if (memberRowIndex_(name)) throw new Error(name + ' はすでに登録されています');
    sheet_('members').appendRow([name, m.status || '', m.gender || '']);
    return getData();
  });
}

/** メンバー編集　m = {oldName, name, status, gender}（名前を変えると過去の明細も追随） */
function updateMember(m) {
  const oldName = String(m.oldName || '');
  const name = String(m.name || '').trim();
  if (!name) throw new Error('名前を入力してください');
  return withLock_(() => {
    const row = memberRowIndex_(oldName);
    if (!row) throw new Error(oldName + ' が見つかりません');
    if (name !== oldName && memberRowIndex_(name)) throw new Error(name + ' はすでに登録されています');
    sheet_('members').getRange(row, 1, 1, 3).setValues([[name, m.status || '', m.gender || '']]);
    if (name !== oldName) {
      const sh = sheet_('shares');
      const values = sh.getDataRange().getValues();
      for (let i = 1; i < values.length; i++) {
        if (String(values[i][2]) === oldName) sh.getRange(i + 1, 3).setValue(name);
      }
    }
    return getData();
  });
}

/** メンバーの並び順を入れ替える　names = 新しい順の名前の配列 */
function reorderMembers(names) {
  return withLock_(() => {
    const sh = sheet_('members');
    const body = sh.getDataRange().getValues().slice(1).filter(r => r[0] !== '');
    const byName = {};
    body.forEach(r => byName[String(r[0])] = r);
    const ordered = [];
    names.forEach(n => { if (byName[n]) { ordered.push(byName[n]); delete byName[n]; } });
    body.forEach(r => { if (byName[String(r[0])]) ordered.push(r); }); // 指定漏れは末尾へ
    if (ordered.length) {
      sh.getRange(2, 1, ordered.length, 3).setValues(ordered.map(r => [r[0], r[1] || '', r[2] || '']));
    }
    return getData();
  });
}

/** メンバー削除（過去の立替の記録はそのまま残す） */
function deleteMember(name) {
  return withLock_(() => {
    const row = memberRowIndex_(String(name));
    if (!row) throw new Error(name + ' が見つかりません');
    sheet_('members').deleteRow(row);
    return getData();
  });
}

/**
 * 立替を登録
 * p = { date, title, total, headcount, memo, shares: [{person, amount}] }
 */
function addAdvance(p) {
  if (!p.title) throw new Error('内容を入力してください');
  if (!p.shares || !p.shares.length) throw new Error('回収する相手と金額を入力してください');
  return withLock_(() => {
    const id = newId_();
    sheet_('advances').appendRow([id, p.date, p.title, p.total, p.headcount, p.memo || '', new Date()]);
    const rows = p.shares.map(s => [newId_(), id, s.person, Number(s.amount), false, '']);
    const sh = sheet_('shares');
    sh.getRange(sh.getLastRow() + 1, 1, rows.length, 6).setValues(rows);
    return getData();
  });
}

/**
 * 立替を編集（回収済みの状態は、同じ相手ならそのまま引き継ぐ）
 * p = { id, date, title, total, headcount, memo, shares: [{person, amount}] }
 */
function updateAdvance(p) {
  if (!p.id) throw new Error('編集する立替が指定されていません');
  if (!p.title) throw new Error('内容を入力してください');
  if (!p.shares || !p.shares.length) throw new Error('回収する相手と金額を入力してください');
  return withLock_(() => {
    const adSh = sheet_('advances');
    const av = adSh.getDataRange().getValues();
    let row = 0;
    for (let i = 1; i < av.length; i++) if (String(av[i][0]) === String(p.id)) { row = i + 1; break; }
    if (!row) throw new Error('編集する立替が見つかりません');
    // ID(1列目)と登録日時(7列目)は触らず、2〜6列目だけ更新
    adSh.getRange(row, 2, 1, 5).setValues([[p.date, p.title, p.total, p.headcount, p.memo || '']]);

    // 旧い回収明細をいったん外し、相手が同じなら 明細ID・回収済み・回収日 を引き継ぐ
    const shSh = sheet_('shares');
    const sv = shSh.getDataRange().getValues();
    const kept = {};
    for (let i = sv.length - 1; i >= 1; i--) {
      if (String(sv[i][1]) === String(p.id)) {
        kept[String(sv[i][2])] = [String(sv[i][0]), sv[i][4], sv[i][5]];
        shSh.deleteRow(i + 1);
      }
    }
    const rows = p.shares.map(s => {
      const old = kept[s.person];
      return old ? [old[0], p.id, s.person, Number(s.amount), old[1], old[2]]
                 : [newId_(), p.id, s.person, Number(s.amount), false, ''];
    });
    shSh.getRange(shSh.getLastRow() + 1, 1, rows.length, 6).setValues(rows);
    return getData();
  });
}

/** 回収済み／未回収の切り替え（複数まとめてOK） */
function setPaid(shareIds, paid) {
  return withLock_(() => {
    const target = {};
    shareIds.forEach(id => target[id] = true);
    const sh = sheet_('shares');
    const values = sh.getDataRange().getValues();
    for (let i = 1; i < values.length; i++) {
      if (target[String(values[i][0])]) {
        sh.getRange(i + 1, 5, 1, 2).setValues([[paid, paid ? new Date() : '']]);
      }
    }
    return getData();
  });
}

/** 立替を削除（紐づく回収明細も消す） */
function deleteAdvance(advanceId) {
  return withLock_(() => {
    const shSh = sheet_('shares');
    const sv = shSh.getDataRange().getValues();
    for (let i = sv.length - 1; i >= 1; i--) {
      if (String(sv[i][1]) === advanceId) shSh.deleteRow(i + 1);
    }
    const adSh = sheet_('advances');
    const av = adSh.getDataRange().getValues();
    for (let i = av.length - 1; i >= 1; i--) {
      if (String(av[i][0]) === advanceId) adSh.deleteRow(i + 1);
    }
    return getData();
  });
}
