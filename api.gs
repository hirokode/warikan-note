/**
 * 立替回収トラッカー（GAS側）v3 — 複数ユーザー対応
 *
 * 画面は GitHub Pages に置き、ここは API（データの受け渡し窓口）として動く。
 * ユーザーごとに「メンバー_u1」「立替_u1」「回収明細_u1」「グループ_u1」のようにシートを分けて持つ。
 *
 * 初回に1回だけ setupMultiUser() を実行すること。
 */

const SHEET_NAMES = { members: 'メンバー', advances: '立替', shares: '回収明細', groups: 'グループ' };
// 別アプリがこのシートを列名で読むため、既存の列名・並びは変えない。新しい列は必ず末尾に足す
const HEADERS = {
  members:  ['名前', '区分', '性別', '所属グループ'],
  advances: ['ID', '日付', '内容', '総額', '割る人数(自分含む)', 'メモ', '登録日時', '支払手段'],
  shares:   ['ID', '立替ID', '相手', '金額', '回収済み', '回収日', '入金手段'],
  groups:   ['グループ名'],
};
const PAY_IN_METHODS  = ['PayPay', '銀行', '現金', '楽天ペイ', 'その他'];   // 回収したときの入金手段
const PAY_OUT_METHODS = ['カード', 'PayPay', '現金', 'その他'];             // 立て替えたときの支払手段

// ログイン情報を入れる管理用シート
const SYS_NAMES = { users: 'ユーザー', sessions: 'セッション' };
const SYS_HEADERS = {
  users:    ['ユーザーID', '名前', 'メール', 'ソルト', 'パスワードハッシュ', '反復回数', '登録日時', '最終ログイン'],
  sessions: ['トークンハッシュ', 'ユーザーID', '発行日時', '有効期限'],
};

const HASH_ITERATIONS = 200;    // パスワードを混ぜ返す回数。多いほど破られにくいが、ログインが遅くなる
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

/**
 * そのユーザー用のシートを（なければ）作る。
 * 既存シートに新しい列の見出しが無ければ、末尾に見出しだけを足す（既存の列・データには触れない）
 */
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
    } else {
      const width = sh.getLastColumn();
      const head = sh.getRange(1, 1, 1, width).getValues()[0].map(String);
      const missing = HEADERS[key].filter(h => head.indexOf(h) < 0);
      if (missing.length) sh.getRange(1, width + 1, 1, missing.length).setValues([missing]);
    }
    ensured_[name] = true;
  });
}
const ensured_ = {};   // この実行中に見出しを点検済みのシート名

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
    case 'setPaid':        return setPaid(uid, args[0], args[1], args[2]);
    case 'deleteAdvance':  return deleteAdvance(uid, args[0]);
    case 'addGroup':       return addGroup(uid, args[0]);
    case 'renameGroup':    return renameGroup(uid, args[0], args[1]);
    case 'deleteGroup':    return deleteGroup(uid, args[0]);
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
  if (!String(code).trim()) throw new Error('招待コードが未設定です。管理者に連絡してください');
  // 保存された値の前後に空白が紛れ込んでいても一致するようにする（コピペ事故対策）
  if (String(invite || '').trim() !== String(code).trim()) throw new Error('招待コードが違います');

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
  const sh = sysSheet_('users');

  // 反復回数が変わっていたら、新しい回数で再ハッシュして保存する（次回以降のログインが速くなる）
  if (iter !== HASH_ITERATIONS) {
    const newSalt = Utilities.getUuid();
    const newHash = hashPassword_(pw, newSalt, HASH_ITERATIONS);
    sh.getRange(found.row, 4, 1, 3).setValues([[newSalt, newHash, HASH_ITERATIONS]]);
  }

  sh.getRange(found.row, 8).setValue(new Date());
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
  if (!sh || !ensured_[name]) {
    ensureUserSheets_(uid);
    sh = ss.getSheetByName(name);
  }
  return sh;
}
function rows_(key, uid) {
  return table_(key, uid).rows;
}
/** データ行と「見出し名 → 列の位置(0始まり)」を返す。末尾に足した列は位置を決め打ちせず見出し名で探す */
function table_(key, uid) {
  const values = sheet_(key, uid).getDataRange().getValues();
  const head = values.shift() || []; // 見出し行を除く
  const idx = {};
  head.forEach((h, i) => { if (h !== '' && !(h in idx)) idx[String(h)] = i; });
  return { rows: values.filter(r => r[0] !== ''), idx: idx, width: head.length };
}
/** 見出し名をキーにした値から、シートの列幅ぶんの1行を作る */
function rowFor_(t, obj) {
  const row = new Array(t.width).fill('');
  Object.keys(obj).forEach(k => { if (k in t.idx) row[t.idx[k]] = obj[k]; });
  return row;
}
function cell_(r, t, name) {
  return (name in t.idx) ? r[t.idx[name]] : '';
}
function checkMethod_(v, list, label) {
  const s = String(v || '').trim();
  if (s && list.indexOf(s) < 0) throw new Error(label + 'が正しくありません: ' + s);
  return s;
}
/** 所属グループ（カンマ区切り）⇔ 配列 */
function splitGroups_(v) {
  return String(v || '').split(',').map(s => s.trim()).filter(Boolean);
}
function joinGroups_(arr) {
  const seen = {};
  return (arr || []).map(s => String(s).trim()).filter(s => s && !seen[s] && (seen[s] = true)).join(',');
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
  const mt = table_('members', uid);
  const members = mt.rows.map(r => ({
    name: String(r[0]), status: String(r[1] || ''), gender: String(r[2] || ''),
    groups: splitGroups_(cell_(r, mt, '所属グループ')),
  }));
  const st = table_('shares', uid);
  const shares = st.rows.map(r => ({
    id: String(r[0]), advanceId: String(r[1]), person: String(r[2]),
    amount: Number(r[3]), paid: r[4] === true, paidDate: date_(r[5]),
    method: String(cell_(r, st, '入金手段') || ''),
  }));
  const at = table_('advances', uid);
  const advances = at.rows.map(r => ({
    id: String(r[0]), date: date_(r[1]), title: String(r[2]),
    total: r[3] === '' ? '' : Number(r[3]), headcount: Number(r[4]), memo: String(r[5] || ''),
    payMethod: String(cell_(r, at, '支払手段') || ''),
    shares: shares.filter(s => s.advanceId === String(r[0])),
  })).sort((a, b) => b.date.localeCompare(a.date));
  const groups = rows_('groups', uid).map(r => String(r[0]));
  return { members, advances, groups };
}

/** メンバー追加　m = {name, status, gender, groups?} */
function addMember(uid, m) {
  const name = String(m.name || '').trim();
  if (!name) throw new Error('名前を入力してください');
  return withLock_(() => {
    if (memberRowIndex_(uid, name)) throw new Error(name + ' はすでに登録されています');
    const t = table_('members', uid);
    sheet_('members', uid).appendRow(rowFor_(t, {
      '名前': name, '区分': m.status || '', '性別': m.gender || '', '所属グループ': joinGroups_(m.groups),
    }));
    return getData(uid);
  });
}

/** メンバー編集　m = {oldName, name, status, gender, groups?}（名前を変えると過去の明細も追随） */
function updateMember(uid, m) {
  const oldName = String(m.oldName || '');
  const name = String(m.name || '').trim();
  if (!name) throw new Error('名前を入力してください');
  return withLock_(() => {
    const row = memberRowIndex_(uid, oldName);
    if (!row) throw new Error(oldName + ' が見つかりません');
    if (name !== oldName && memberRowIndex_(uid, name)) throw new Error(name + ' はすでに登録されています');
    const msh = sheet_('members', uid);
    msh.getRange(row, 1, 1, 3).setValues([[name, m.status || '', m.gender || '']]);
    if (Array.isArray(m.groups)) {
      const t = table_('members', uid);
      msh.getRange(row, t.idx['所属グループ'] + 1).setValue(joinGroups_(m.groups));
    }
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
      // 所属グループなど末尾の列も一緒に動かすため、行の全列を書き戻す
      sh.getRange(2, 1, ordered.length, ordered[0].length).setValues(ordered);
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
 * p = { date, title, total, headcount, memo, payMethod, shares: [{person, amount}] }
 */
function addAdvance(uid, p) {
  if (!p.title) throw new Error('内容を入力してください');
  if (!p.shares || !p.shares.length) throw new Error('回収する相手と金額を入力してください');
  const payMethod = checkMethod_(p.payMethod, PAY_OUT_METHODS, '支払手段');
  return withLock_(() => {
    const id = newId_();
    const at = table_('advances', uid);
    sheet_('advances', uid).appendRow(rowFor_(at, {
      'ID': id, '日付': p.date, '内容': p.title, '総額': p.total, '割る人数(自分含む)': p.headcount,
      'メモ': p.memo || '', '登録日時': new Date(), '支払手段': payMethod,
    }));
    const st = table_('shares', uid);
    const rows = p.shares.map(s => rowFor_(st, {
      'ID': newId_(), '立替ID': id, '相手': s.person, '金額': Number(s.amount), '回収済み': false, '回収日': '', '入金手段': '',
    }));
    const sh = sheet_('shares', uid);
    sh.getRange(sh.getLastRow() + 1, 1, rows.length, st.width).setValues(rows);
    return getData(uid);
  });
}

/**
 * 立替を編集（回収済みの状態は、同じ相手ならそのまま引き継ぐ）
 * p = { id, date, title, total, headcount, memo, payMethod, shares: [{person, amount}] }
 */
function updateAdvance(uid, p) {
  if (!p.id) throw new Error('編集する立替が指定されていません');
  if (!p.title) throw new Error('内容を入力してください');
  if (!p.shares || !p.shares.length) throw new Error('回収する相手と金額を入力してください');
  const payMethod = checkMethod_(p.payMethod, PAY_OUT_METHODS, '支払手段');
  return withLock_(() => {
    const adSh = sheet_('advances', uid);
    const av = adSh.getDataRange().getValues();
    let row = 0;
    for (let i = 1; i < av.length; i++) if (String(av[i][0]) === String(p.id)) { row = i + 1; break; }
    if (!row) throw new Error('編集する立替が見つかりません');
    // ID(1列目)と登録日時(7列目)は触らず、2〜6列目と支払手段だけ更新
    adSh.getRange(row, 2, 1, 5).setValues([[p.date, p.title, p.total, p.headcount, p.memo || '']]);
    const at = table_('advances', uid);
    adSh.getRange(row, at.idx['支払手段'] + 1).setValue(payMethod);

    // 旧い回収明細をいったん外し、相手が同じなら 明細ID・回収済み・回収日・入金手段 を引き継ぐ
    const shSh = sheet_('shares', uid);
    const st = table_('shares', uid);
    const sv = shSh.getDataRange().getValues();
    const kept = {};
    for (let i = sv.length - 1; i >= 1; i--) {
      if (String(sv[i][1]) === String(p.id)) {
        kept[String(sv[i][2])] = { id: String(sv[i][0]), paid: sv[i][4], date: sv[i][5], method: cell_(sv[i], st, '入金手段') };
        shSh.deleteRow(i + 1);
      }
    }
    const rows = p.shares.map(s => {
      const old = kept[s.person];
      return rowFor_(st, old
        ? { 'ID': old.id, '立替ID': p.id, '相手': s.person, '金額': Number(s.amount), '回収済み': old.paid, '回収日': old.date, '入金手段': old.method }
        : { 'ID': newId_(), '立替ID': p.id, '相手': s.person, '金額': Number(s.amount), '回収済み': false, '回収日': '', '入金手段': '' });
    });
    shSh.getRange(shSh.getLastRow() + 1, 1, rows.length, st.width).setValues(rows);
    return getData(uid);
  });
}

/**
 * 回収済み／未回収の切り替え（複数まとめてOK）
 * method … 回収済みにするときの入金手段。未回収に戻すときは入金手段も消す
 */
function setPaid(uid, shareIds, paid, method) {
  const m = paid ? checkMethod_(method, PAY_IN_METHODS, '入金手段') : '';
  return withLock_(() => {
    const target = {};
    shareIds.forEach(id => target[id] = true);
    const sh = sheet_('shares', uid);
    const methodCol = table_('shares', uid).idx['入金手段'] + 1;
    const values = sh.getDataRange().getValues();
    for (let i = 1; i < values.length; i++) {
      if (target[String(values[i][0])]) {
        sh.getRange(i + 1, 5, 1, 2).setValues([[paid, paid ? new Date() : '']]);
        sh.getRange(i + 1, methodCol).setValue(m);
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

// ---------- グループ（メンバーの集まり。登録時に参加者を一括で選ぶためのもの） ----------

function groupName_(name) {
  const nm = String(name || '').trim();
  if (!nm) throw new Error('グループ名を入力してください');
  if (nm.indexOf(',') >= 0) throw new Error('グループ名に「,」は使えません');
  return nm;
}
function groupRowIndex_(uid, name) {
  const values = sheet_('groups', uid).getDataRange().getValues();
  for (let i = 1; i < values.length; i++) if (String(values[i][0]) === name) return i + 1;
  return 0;
}
/** 全メンバーの所属グループを書き換える　fn(配列) → 新しい配列 */
function rewriteMemberGroups_(uid, fn) {
  const sh = sheet_('members', uid);
  const t = table_('members', uid);
  const col = t.idx['所属グループ'];
  const values = sh.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    const before = String(values[i][col] || '');
    const after = joinGroups_(fn(splitGroups_(before)));
    if (after !== before) sh.getRange(i + 1, col + 1).setValue(after);
  }
}

/** グループ追加 */
function addGroup(uid, name) {
  const nm = groupName_(name);
  return withLock_(() => {
    if (groupRowIndex_(uid, nm)) throw new Error(nm + ' はすでにあります');
    sheet_('groups', uid).appendRow([nm]);
    return getData(uid);
  });
}

/** グループ名の変更（メンバーの所属グループも追随） */
function renameGroup(uid, oldName, name) {
  const nm = groupName_(name);
  return withLock_(() => {
    const row = groupRowIndex_(uid, String(oldName));
    if (!row) throw new Error(oldName + ' が見つかりません');
    if (nm !== oldName && groupRowIndex_(uid, nm)) throw new Error(nm + ' はすでにあります');
    sheet_('groups', uid).getRange(row, 1).setValue(nm);
    rewriteMemberGroups_(uid, gs => gs.map(g => g === oldName ? nm : g));
    return getData(uid);
  });
}

/** グループ削除（メンバーの所属からも外す。メンバー自体は消さない） */
function deleteGroup(uid, name) {
  return withLock_(() => {
    const row = groupRowIndex_(uid, String(name));
    if (!row) throw new Error(name + ' が見つかりません');
    sheet_('groups', uid).deleteRow(row);
    rewriteMemberGroups_(uid, gs => gs.filter(g => g !== name));
    return getData(uid);
  });
}

// ===================================================================
// 調査と修復（エディタから手で実行する。画面からは呼べない）
// ===================================================================

/**
 * いま何がどうなっているかを一覧で出す。
 * パスワードやトークンは表示しない。
 */
function diagnose() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const props = PropertiesService.getScriptProperties();
  const out = [];

  out.push('■ シート一覧（行数は見出しを除いたデータ件数）');
  ss.getSheets().forEach(sh => {
    const rows = Math.max(0, sh.getLastRow() - 1);
    out.push('  ' + sh.getName() + '  … ' + rows + '件');
  });

  out.push('');
  out.push('■ 登録済みのアカウント');
  const uv = sysSheet_('users').getDataRange().getValues();
  if (uv.length < 2) {
    out.push('  （まだありません）');
  } else {
    for (let i = 1; i < uv.length; i++) {
      out.push('  ' + uv[i][0] + ' : ' + uv[i][1] + ' <' + uv[i][2] + '>');
    }
  }

  out.push('');
  out.push('■ 採番カウンタ UID_SEQ = ' + props.getProperty('UID_SEQ'));
  out.push('■ PEPPER 設定済み = ' + (props.getProperty('PEPPER') ? 'はい' : 'いいえ'));
  out.push('■ INVITE_CODE 設定済み = ' + (props.getProperty('INVITE_CODE') ? 'はい' : 'いいえ'));

  const msg = out.join('\n');
  Logger.log(msg);
  return msg;
}

/**
 * データの入ったシートを、指定したアカウントのものとして引き継ぐ。
 *
 *   from … 引き継ぎ元。旧バージョンのシート（メンバー／立替／回収明細）なら '' を渡す。
 *          すでに 'u1' が付いているなら 'u1' を渡す。
 *   to   … 引き継ぎ先のユーザーID（例 'u2'）
 *
 * 例：moveDataTo('', 'u1')     旧シートを u1 のものにする
 *     moveDataTo('u1', 'u2')   u1 のシートを u2 のものにする
 *
 * 引き継ぎ先にデータが入っている場合は、安全のため何もせず中止する。
 */
function moveDataTo(from, to) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!to) throw new Error('引き継ぎ先のユーザーIDを指定してください（例 "u1"）');
  const log = [];

  // 先に全部を点検してから動かす（途中で止まって半端な状態にしないため）
  const plan = [];
  Object.keys(SHEET_NAMES).forEach(key => {
    const base = SHEET_NAMES[key];
    const srcName = from ? base + '_' + from : base;
    const dstName = base + '_' + to;
    const src = ss.getSheetByName(srcName);
    const dst = ss.getSheetByName(dstName);

    if (!src) { log.push('× ' + srcName + ' が見つかりません（このシートは飛ばします）'); return; }
    if (src.getName() === dstName) { log.push('・' + dstName + ' はすでに引き継ぎ済みです'); return; }
    if (dst && dst.getLastRow() > 1) {
      throw new Error(dstName + ' にすでに ' + (dst.getLastRow() - 1) + '件のデータがあります。'
        + '上書きを避けるため中止しました。内容を確認してください。');
    }
    plan.push({ src: src, dst: dst, srcName: srcName, dstName: dstName });
  });

  if (!plan.length) { const m = log.join('\n') || '動かすものがありませんでした'; Logger.log(m); return m; }

  plan.forEach(item => {
    if (item.dst) { ss.deleteSheet(item.dst); log.push('・空だった ' + item.dstName + ' を削除しました'); }
    item.src.setName(item.dstName);
    log.push('○ ' + item.srcName + ' → ' + item.dstName + '（中身はそのまま）');
  });

  const msg = log.join('\n');
  Logger.log(msg);
  return msg;
}

/**
 * パスワードの再設定（管理者用）。引数なしで実行できる。
 *
 * 手順：
 *   1. スクリプト プロパティに次の2つを追加する
 *        RESET_EMAIL … 対象のメールアドレス
 *        RESET_PW    … 新しいパスワード（8文字以上）
 *   2. この関数を ▶ で実行する
 *   3. 2つのプロパティは自動的に消える
 *
 * パスワードをコードに書かずに済むので、git にも実行履歴にも残らない。
 */
function resetPassword() {
  const props = PropertiesService.getScriptProperties();
  const email = props.getProperty('RESET_EMAIL');
  const pw = props.getProperty('RESET_PW');

  if (!email || !pw) {
    throw new Error('先にスクリプト プロパティへ RESET_EMAIL と RESET_PW を設定してください');
  }
  if (String(pw).length < MIN_PASSWORD) {
    throw new Error('新しいパスワードは' + MIN_PASSWORD + '文字以上にしてください');
  }

  const found = findUserRow_(email);
  if (!found) {
    throw new Error(email + ' のアカウントが見つかりません');
  }

  const uid = String(found.values[0]);
  const salt = Utilities.getUuid();
  const hash = hashPassword_(String(pw), salt, HASH_ITERATIONS);
  const sh = sysSheet_('users');
  sh.getRange(found.row, 4, 1, 3).setValues([[salt, hash, HASH_ITERATIONS]]);

  // 念のため、このアカウントの古いログイン状態はすべて無効にする
  const ses = sysSheet_('sessions');
  const sv = ses.getDataRange().getValues();
  let killed = 0;
  for (let i = sv.length - 1; i >= 1; i--) {
    if (String(sv[i][1]) === uid) { ses.deleteRow(i + 1); killed++; }
  }

  // 使い終わったら必ず消す（パスワードを残さないため）
  props.deleteProperty('RESET_EMAIL');
  props.deleteProperty('RESET_PW');

  const msg = uid + '（' + email + '）のパスワードを再設定しました。\n'
            + '古いログイン状態 ' + killed + '件を無効にしました。\n'
            + 'RESET_EMAIL と RESET_PW は削除済みです。新しいパスワードでログインしてください。';
  Logger.log(msg);
  return msg;
}
