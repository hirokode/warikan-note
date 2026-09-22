/**
 * 立替回収トラッカー（GAS側）v3 — 複数ユーザー対応
 *
 * 画面は GitHub Pages に置き、ここは API（データの受け渡し窓口）として動く。
 * ユーザーごとに「メンバー_u1」「立替_u1」「回収明細_u1」のようにシートを分けて持つ。
 *
 * 初回に1回だけ setupMultiUser() を実行すること。
 */

const SHEET_NAMES = { members: 'メンバー', advances: '立替', shares: '回収明細' };
const HEADERS = {
  members:  ['名前', '区分', '性別'],
  advances: ['ID', '日付', '内容', '総額', '割る人数(自分含む)', 'メモ', '登録日時'],
  shares:   ['ID', '立替ID', '相手', '金額', '回収済み', '回収日'],
};

// ログイン情報を入れる管理用シート
const SYS_NAMES = { users: 'ユーザー', sessions: 'セッション' };
const SYS_HEADERS = {
  users:    ['ユーザーID', '名前', 'メール', 'ソルト', 'パスワードハッシュ', '反復回数', '登録日時', '最終ログイン'],
  sessions: ['トークンハッシュ', 'ユーザーID', '発行日時', '有効期限'],
};

const HASH_ITERATIONS = 2000;   // パスワードを混ぜ返す回数。多いほど破られにくいが、ログインが遅くなる
const SESSION_DAYS    = 90;     // ログイン状態を保つ日数
const MIN_PASSWORD    = 8;      // パスワードの最低文字数

// ===================================================================
// 初回セットアップ（エディタから1回だけ手で実行する）
// ===================================================================

/**
 * 既存の3シートを u1（＝最初に登録するアカウント）用に付け替え、
 * 管理用シートとパスワード用の秘密の値を用意する。
 * セルの中身と列は一切触らない。シート名を変えるだけ。
 */
function setupMultiUser() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const props = PropertiesService.getScriptProperties();
  const log = [];

  ensureSystemSheets_();
  log.push('管理用シート（ユーザー／セッション）を用意しました');

  if (!props.getProperty('PEPPER')) {
    props.setProperty('PEPPER', Utilities.getUuid() + Utilities.getUuid());
    log.push('パスワード用の秘密の値（PEPPER）を自動生成しました');
  } else {
    log.push('PEPPER はすでに設定済みです');
  }

  if (props.getProperty('UID_SEQ') === null) {
    props.setProperty('UID_SEQ', '0');
  }

  Object.keys(SHEET_NAMES).forEach(key => {
    const oldName = SHEET_NAMES[key];
    const newName = oldName + '_u1';
    const old = ss.getSheetByName(oldName);
    if (old && !ss.getSheetByName(newName)) {
      old.setName(newName);
      log.push(oldName + ' → ' + newName + ' に名前を変えました（中身はそのまま）');
    }
  });

  const invite = props.getProperty('INVITE_CODE');
  log.push(invite ? '招待コードは設定済みです' : '⚠ 招待コード（INVITE_CODE）がまだ未設定です。スクリプト プロパティから設定してください');

  const msg = log.join('\n');
  Logger.log(msg);
  return msg;
}

function ensureSystemSheets_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(SYS_NAMES).forEach(key => {
    let sh = ss.getSheetByName(SYS_NAMES[key]);
    if (!sh) {
      sh = ss.insertSheet(SYS_NAMES[key]);
      sh.appendRow(SYS_HEADERS[key]);
      sh.setFrozenRows(1);
    }
  });
}

/** そのユーザー用の3シートを（なければ）作る */
function ensureUserSheets_(uid) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(SHEET_NAMES).forEach(key => {
    const name = SHEET_NAMES[key] + '_' + uid;
    let sh = ss.getSheetByName(name);
    if (!sh) {
      sh = ss.insertSheet(name);
      sh.appendRow(HEADERS[key]);
      sh.setFrozenRows(1);
    } else if (sh.getLastRow() === 0) {
      sh.appendRow(HEADERS[key]);
      sh.setFrozenRows(1);
    }
  });
}

// ===================================================================
// 通信の入口
// ===================================================================

/** ブラウザから直接開かれたとき（アプリ本体は GitHub Pages 側にある） */
function doGet() {
  return ContentService.createTextOutput(
    'このURLは立替回収トラッカーのデータ窓口です。アプリの画面ではありません。'
  );
}

/** 画面からの問い合わせは、すべてここを通る */
function doPost(e) {
  let out;
  try {
    const req = JSON.parse(e.postData.contents);
    out = { ok: true, data: route_(req) };
  } catch (err) {
    out = { ok: false, error: (err && err.message) ? err.message : String(err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out))
    .setMimeType(ContentService.MimeType.JSON);
}

function route_(req) {
  const action = String(req.action || '');
  const args = req.args || [];

  // ログインしなくても呼べるもの
  if (action === 'signup') return signup_(req.name, req.email, req.password, req.invite);
  if (action === 'login')  return login_(req.email, req.password);
  if (action === 'logout') { logout_(req.token); return { ok: true }; }

  // ここから先はログインが必要。トークンから「誰か」を確定させる
  const uid = uidFromToken_(req.token);

  switch (action) {
    case 'getData':        return getData(uid);
    case 'addMember':      return addMember(uid, args[0]);
    case 'updateMember':   return updateMember(uid, args[0]);
    case 'reorderMembers': return reorderMembers(uid, args[0]);
    case 'deleteMember':   return deleteMember(uid, args[0]);
    case 'addAdvance':     return addAdvance(uid, args[0]);
    case 'updateAdvance':  return updateAdvance(uid, args[0]);
    case 'setPaid':        return setPaid(uid, args[0], args[1]);
    case 'deleteAdvance':  return deleteAdvance(uid, args[0]);
    case 'me':             return profile_(uid);
    default: throw new Error('知らない操作です: ' + action);
  }
}

// ===================================================================
// ログインまわり
// ===================================================================

/** パスワードに混ぜる、サーバーだけが知っている秘密の値 */
function pepper_() {
  const props = PropertiesService.getScriptProperties();
  let p = props.getProperty('PEPPER');
  if (!p) { p = Utilities.getUuid() + Utilities.getUuid(); props.setProperty('PEPPER', p); }
  return p;
}

/**
 * パスワードを元に戻せない形に変換する。
 * ソルト（利用者ごとの飾り）＋ペッパー（サーバーだけが持つ秘密）を混ぜ、
 * 何千回も掛け直すことで、総当たりに時間がかかるようにしている。
 */
function hashPassword_(password, salt, iterations) {
  let bytes = Utilities.newBlob(salt + '|' + password + '|' + pepper_()).getBytes();
  for (let i = 0; i < iterations; i++) {
    bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, bytes);
  }
  return Utilities.base64Encode(bytes);
}

function normalizeEmail_(email) {
  return String(email || '').trim().toLowerCase();
}

function sysSheet_(key) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SYS_NAMES[key]);
  if (!sh) { ensureSystemSheets_(); sh = ss.getSheetByName(SYS_NAMES[key]); }
  return sh;
}

function findUserRow_(email) {
  const sh = sysSheet_('users');
  const v = sh.getDataRange().getValues();
  const target = normalizeEmail_(email);
  for (let i = 1; i < v.length; i++) {
    if (normalizeEmail_(v[i][2]) === target) return { row: i + 1, values: v[i] };
  }
  return null;
}

function findUserByUid_(uid) {
  const sh = sysSheet_('users');
  const v = sh.getDataRange().getValues();
  for (let i = 1; i < v.length; i++) {
    if (String(v[i][0]) === String(uid)) return { row: i + 1, values: v[i] };
  }
  return null;
}

function nextUid_() {
  const props = PropertiesService.getScriptProperties();
  const n = Number(props.getProperty('UID_SEQ') || '0') + 1;
  props.setProperty('UID_SEQ', String(n));
  return 'u' + n;
}

function newToken_() {
  return (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, '');
}

function tokenHash_(token) {
  return Utilities.base64Encode(
    Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(token))
  );
}

function issueSession_(uid) {
  const token = newToken_();
  const now = new Date();
  const until = new Date(now.getTime() + SESSION_DAYS * 24 * 60 * 60 * 1000);
  sysSheet_('sessions').appendRow([tokenHash_(token), uid, now, until]);
  return { token: token, expires: until.toISOString() };
}

/** トークンから利用者を割り出す。ここを通らないとデータには触れない */
function uidFromToken_(token) {
  if (!token) throw new Error('ログインしてください');
  const sh = sysSheet_('sessions');
  const v = sh.getDataRange().getValues();
  const h = tokenHash_(token);
  const now = new Date();
  for (let i = 1; i < v.length; i++) {
    if (String(v[i][0]) === h) {
      if (new Date(v[i][3]) < now) {
        sh.deleteRow(i + 1);
        throw new Error('ログインの期限が切れました。もう一度ログインしてください');
      }
      return String(v[i][1]);
    }
  }
  throw new Error('ログインしてください');
}

function profile_(uid) {
  const u = findUserByUid_(uid);
  if (!u) throw new Error('アカウントが見つかりません');
  return { uid: String(u.values[0]), name: String(u.values[1]), email: String(u.values[2]) };
}

function signup_(name, email, password, invite) {
  const props = PropertiesService.getScriptProperties();
  const code = props.getProperty('INVITE_CODE');
  if (!code) throw new Error('招待コードが未設定です。管理者に連絡してください');
  if (String(invite || '').trim() !== code) throw new Error('招待コードが違います');

  const nm = String(name || '').trim();
  const em = normalizeEmail_(email);
  const pw = String(password || '');
  if (!nm) throw new Error('名前を入力してください');
  if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(em)) throw new Error('メールアドレスの形式が正しくありません');
  if (pw.length < MIN_PASSWORD) throw new Error('パスワードは' + MIN_PASSWORD + '文字以上にしてください');

  return withLock_(() => {
    if (findUserRow_(em)) throw new Error('このメールアドレスは登録済みです');
    const uid = nextUid_();
    const salt = Utilities.getUuid();
    const hash = hashPassword_(pw, salt, HASH_ITERATIONS);
    const now = new Date();
    sysSheet_('users').appendRow([uid, nm, em, salt, hash, HASH_ITERATIONS, now, now]);
    ensureUserSheets_(uid);
    const s = issueSession_(uid);
    return { token: s.token, expires: s.expires, profile: { uid: uid, name: nm, email: em } };
  });
}

function login_(email, password) {
  const em = normalizeEmail_(email);
  const pw = String(password || '');
  const found = findUserRow_(em);
  // 「メールが無い」と「パスワードが違う」を区別しない（どのメールが登録済みか探られないため）
  if (!found) throw new Error('メールアドレスかパスワードが違います');

  const salt = String(found.values[3]);
  const stored = String(found.values[4]);
  const iter = Number(found.values[5]) || HASH_ITERATIONS;
  if (hashPassword_(pw, salt, iter) !== stored) throw new Error('メールアドレスかパスワードが違います');

  const uid = String(found.values[0]);
  ensureUserSheets_(uid);
  sysSheet_('users').getRange(found.row, 8).setValue(new Date());
  const s = issueSession_(uid);
  return {
    token: s.token, expires: s.expires,
    profile: { uid: uid, name: String(found.values[1]), email: String(found.values[2]) },
  };
}

function logout_(token) {
  if (!token) return;
  const sh = sysSheet_('sessions');
  const v = sh.getDataRange().getValues();
  const h = tokenHash_(token);
  for (let i = v.length - 1; i >= 1; i--) {
    if (String(v[i][0]) === h) sh.deleteRow(i + 1);
  }
}

/** 期限切れのセッションを掃除する（時間主導トリガーに入れておくとよい） */
function cleanupSessions() {
  const sh = sysSheet_('sessions');
  const v = sh.getDataRange().getValues();
  const now = new Date();
  let removed = 0;
  for (let i = v.length - 1; i >= 1; i--) {
    if (v[i][3] && new Date(v[i][3]) < now) { sh.deleteRow(i + 1); removed++; }
  }
  return removed + '件の期限切れセッションを削除しました';
}

// ===================================================================
// 内部で使う小道具
// ===================================================================

function sheet_(key, uid) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const name = SHEET_NAMES[key] + '_' + uid;
  let sh = ss.getSheetByName(name);
  if (!sh || sh.getLastColumn() < HEADERS[key].length) {
    ensureUserSheets_(uid);
    sh = ss.getSheetByName(name);
  }
  return sh;
}
function rows_(key, uid) {
  const values = sheet_(key, uid).getDataRange().getValues();
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
function memberRowIndex_(uid, name) {
  const values = sheet_('members', uid).getDataRange().getValues();
  for (let i = 1; i < values.length; i++) if (String(values[i][0]) === name) return i + 1;
  return 0;
}

// ===================================================================
// 画面から呼ばれる処理（中身は v2 と同じ。対象シートが利用者ごとになっただけ）
// ===================================================================

/** 全データをまとめて返す */
function getData(uid) {
  const members = rows_('members', uid).map(r => ({
    name: String(r[0]), status: String(r[1] || ''), gender: String(r[2] || ''),
  }));
  const shares = rows_('shares', uid).map(r => ({
    id: String(r[0]), advanceId: String(r[1]), person: String(r[2]),
    amount: Number(r[3]), paid: r[4] === true, paidDate: date_(r[5]),
  }));
  const advances = rows_('advances', uid).map(r => ({
    id: String(r[0]), date: date_(r[1]), title: String(r[2]),
    total: r[3] === '' ? '' : Number(r[3]), headcount: Number(r[4]), memo: String(r[5] || ''),
    shares: shares.filter(s => s.advanceId === String(r[0])),
  })).sort((a, b) => b.date.localeCompare(a.date));
  return { members, advances };
}

/** メンバー追加　m = {name, status, gender} */
function addMember(uid, m) {
  const name = String(m.name || '').trim();
  if (!name) throw new Error('名前を入力してください');
  return withLock_(() => {
    if (memberRowIndex_(uid, name)) throw new Error(name + ' はすでに登録されています');
    sheet_('members', uid).appendRow([name, m.status || '', m.gender || '']);
    return getData(uid);
  });
}

/** メンバー編集　m = {oldName, name, status, gender}（名前を変えると過去の明細も追随） */
function updateMember(uid, m) {
  const oldName = String(m.oldName || '');
  const name = String(m.name || '').trim();
  if (!name) throw new Error('名前を入力してください');
  return withLock_(() => {
    const row = memberRowIndex_(uid, oldName);
    if (!row) throw new Error(oldName + ' が見つかりません');
    if (name !== oldName && memberRowIndex_(uid, name)) throw new Error(name + ' はすでに登録されています');
    sheet_('members', uid).getRange(row, 1, 1, 3).setValues([[name, m.status || '', m.gender || '']]);
    if (name !== oldName) {
      const sh = sheet_('shares', uid);
      const values = sh.getDataRange().getValues();
      for (let i = 1; i < values.length; i++) {
        if (String(values[i][2]) === oldName) sh.getRange(i + 1, 3).setValue(name);
      }
    }
    return getData(uid);
  });
}

/** メンバーの並び順を入れ替える　names = 新しい順の名前の配列 */
function reorderMembers(uid, names) {
  return withLock_(() => {
    const sh = sheet_('members', uid);
    const body = sh.getDataRange().getValues().slice(1).filter(r => r[0] !== '');
    const byName = {};
    body.forEach(r => byName[String(r[0])] = r);
    const ordered = [];
    names.forEach(n => { if (byName[n]) { ordered.push(byName[n]); delete byName[n]; } });
    body.forEach(r => { if (byName[String(r[0])]) ordered.push(r); }); // 指定漏れは末尾へ
    if (ordered.length) {
      sh.getRange(2, 1, ordered.length, 3).setValues(ordered.map(r => [r[0], r[1] || '', r[2] || '']));
    }
    return getData(uid);
  });
}

/** メンバー削除（過去の立替の記録はそのまま残す） */
function deleteMember(uid, name) {
  return withLock_(() => {
    const row = memberRowIndex_(uid, String(name));
    if (!row) throw new Error(name + ' が見つかりません');
    sheet_('members', uid).deleteRow(row);
    return getData(uid);
  });
}

/**
 * 立替を登録
 * p = { date, title, total, headcount, memo, shares: [{person, amount}] }
 */
function addAdvance(uid, p) {
  if (!p.title) throw new Error('内容を入力してください');
  if (!p.shares || !p.shares.length) throw new Error('回収する相手と金額を入力してください');
  return withLock_(() => {
    const id = newId_();
    sheet_('advances', uid).appendRow([id, p.date, p.title, p.total, p.headcount, p.memo || '', new Date()]);
    const rows = p.shares.map(s => [newId_(), id, s.person, Number(s.amount), false, '']);
    const sh = sheet_('shares', uid);
    sh.getRange(sh.getLastRow() + 1, 1, rows.length, 6).setValues(rows);
    return getData(uid);
  });
}

/**
 * 立替を編集（回収済みの状態は、同じ相手ならそのまま引き継ぐ）
 * p = { id, date, title, total, headcount, memo, shares: [{person, amount}] }
 */
function updateAdvance(uid, p) {
  if (!p.id) throw new Error('編集する立替が指定されていません');
  if (!p.title) throw new Error('内容を入力してください');
  if (!p.shares || !p.shares.length) throw new Error('回収する相手と金額を入力してください');
  return withLock_(() => {
    const adSh = sheet_('advances', uid);
    const av = adSh.getDataRange().getValues();
    let row = 0;
    for (let i = 1; i < av.length; i++) if (String(av[i][0]) === String(p.id)) { row = i + 1; break; }
    if (!row) throw new Error('編集する立替が見つかりません');
    // ID(1列目)と登録日時(7列目)は触らず、2〜6列目だけ更新
    adSh.getRange(row, 2, 1, 5).setValues([[p.date, p.title, p.total, p.headcount, p.memo || '']]);

    // 旧い回収明細をいったん外し、相手が同じなら 明細ID・回収済み・回収日 を引き継ぐ
    const shSh = sheet_('shares', uid);
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
    return getData(uid);
  });
}

/** 回収済み／未回収の切り替え（複数まとめてOK） */
function setPaid(uid, shareIds, paid) {
  return withLock_(() => {
    const target = {};
    shareIds.forEach(id => target[id] = true);
    const sh = sheet_('shares', uid);
    const values = sh.getDataRange().getValues();
    for (let i = 1; i < values.length; i++) {
      if (target[String(values[i][0])]) {
        sh.getRange(i + 1, 5, 1, 2).setValues([[paid, paid ? new Date() : '']]);
      }
    }
    return getData(uid);
  });
}

/** 立替を削除（紐づく回収明細も消す） */
function deleteAdvance(uid, advanceId) {
  return withLock_(() => {
    const shSh = sheet_('shares', uid);
    const sv = shSh.getDataRange().getValues();
    for (let i = sv.length - 1; i >= 1; i--) {
      if (String(sv[i][1]) === advanceId) shSh.deleteRow(i + 1);
    }
    const adSh = sheet_('advances', uid);
    const av = adSh.getDataRange().getValues();
    for (let i = av.length - 1; i >= 1; i--) {
      if (String(av[i][0]) === advanceId) adSh.deleteRow(i + 1);
    }
    return getData(uid);
  });
}
