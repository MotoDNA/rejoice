/**
 * 리조이스 챌린지 - 서버 (Google Apps Script)
 * 스프레드시트에 글/좋아요/점수를 저장하고, 사진은 구글 드라이브에 저장합니다.
 */

var ADMIN_CODE  = '0914';                 // 임원진 코드 (원하는 숫자로 바꾸세요)
var FOLDER_NAME = '리조이스 챌린지 사진';   // 사진이 저장될 드라이브 폴더 이름

var HEAD = ['id','type','name','text','l1','l2','l3','l4','image','likes','a','b','ts','deleted','cid'];
var CID_COL = HEAD.length;   // 중복 등록 방지용 키가 들어가는 열

function doGet(e)  { return handle_(e); }
function doPost(e) { return handle_(e); }

function handle_(e) {
  var p = {};
  try {
    if (e && e.postData && e.postData.contents) p = JSON.parse(e.postData.contents);
    else p = (e && e.parameter) || {};
  } catch (err) { p = (e && e.parameter) || {}; }

  var out;
  try { out = route_(p); }
  catch (err) { out = { ok: false, error: String(err) }; }

  return ContentService.createTextOutput(JSON.stringify(out))
    .setMimeType(ContentService.MimeType.JSON);
}

function route_(p) {
  var action = p.action || 'list';

  if (action === 'list') return { ok: true, entries: readAll_(), totalMembers: getConfig_('totalMembers') };

  // ---- 투표 (읽기) ----
  if (action === 'vote_get')    return voteGet_(p);
  if (action === 'vote_result') return voteResult_(p);

  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    if (action === 'add')    return addEntry_(p);
    if (action === 'like')   return likeEntry_(p);
    if (action === 'score')  return scoreEntry_(p);
    if (action === 'delete') return deleteEntry_(p);
    if (action === 'config') return setTotal_(p);
    if (action === 'vote_set')  return voteSet_(p);
    if (action === 'vote_open') return voteOpen_(p);
    return { ok: false, error: 'unknown action: ' + action };
  } finally {
    lock.releaseLock();
  }
}

/* ---------- 시트 ---------- */
var SHEET_ID = '1iUKIQJSwh0d5HBXZ5C67OBUUMUTyCpnjnmi89nXizEw';  // 데이터가 저장될 스프레드시트 ID
function book_() { return SpreadsheetApp.openById(SHEET_ID); }

function sheet_(name, header) {
  var ss = book_();
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(header);
    sh.setFrozenRows(1);
  }
  return sh;
}

function entriesSheet_() {
  var sh = sheet_('entries', HEAD);
  // cid 열이 없던 시절에 만들어진 시트를 위한 보정
  if (sh.getLastColumn() < CID_COL) sh.getRange(1, CID_COL).setValue('cid');
  return sh;
}

// 같은 cid로 이미 등록된 행을 찾습니다 (응답이 유실돼 재시도될 때 중복 방지)
function findByCid_(cid) {
  if (!cid) return 0;
  var sh = entriesSheet_();
  var last = sh.getLastRow();
  if (last < 2) return 0;
  var vals = sh.getRange(2, CID_COL, last - 1, 1).getValues();
  for (var i = 0; i < vals.length; i++) if (String(vals[i][0]) === String(cid)) return i + 2;
  return 0;
}
function configSheet_()  { return sheet_('config', ['key', 'value']); }

function readAll_() {
  var sh = entriesSheet_();
  var last = sh.getLastRow();
  if (last < 2) return [];
  var rows = sh.getRange(2, 1, last - 1, HEAD.length).getValues();
  var out = [];
  for (var i = 0; i < rows.length; i++) {
    var r = rows[i];
    if (r[13] === true || r[13] === 'TRUE' || r[13] === 1) continue;
    if (!r[0]) continue;
    out.push({
      id: String(r[0]), type: String(r[1]), name: String(r[2]), text: String(r[3]),
      lines: [String(r[4]), String(r[5]), String(r[6]), String(r[7])],
      image: String(r[8]),
      likes: Number(r[9]) || 0, a: Number(r[10]) || 0, b: Number(r[11]) || 0,
      ts: Number(r[12]) || 0
    });
  }
  return out;
}

function findRow_(id) {
  var sh = entriesSheet_();
  var last = sh.getLastRow();
  if (last < 2) return 0;
  var ids = sh.getRange(2, 1, last - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) if (String(ids[i][0]) === String(id)) return i + 2;
  return 0;
}

/* ---------- 동작 ---------- */
function addEntry_(p) {
  var name = String(p.name || '').slice(0, 12);
  if (!name) return { ok: false, error: '이름이 없습니다' };

  // 이미 같은 cid로 들어온 요청이면 새로 만들지 않고 기존 결과를 돌려줍니다
  var dup = findByCid_(p.cid);
  if (dup) {
    var row = entriesSheet_().getRange(dup, 1, 1, HEAD.length).getValues()[0];
    return { ok: true, id: String(row[0]), image: String(row[8]), duplicate: true };
  }

  var id  = 'e' + new Date().getTime() + Math.floor(Math.random() * 1000);
  var url = '';
  if (p.type === 'photo') {
    if (!p.image) return { ok: false, error: '사진이 없습니다' };
    url = saveImage_(p.image, id);
  }
  var L = p.lines || ['', '', '', ''];

  entriesSheet_().appendRow([
    id, p.type, name, String(p.text || '').slice(0, 2000),
    L[0] || '', L[1] || '', L[2] || '', L[3] || '',
    url, 0, 0, 0, new Date().getTime(), false, String(p.cid || '')
  ]);
  return { ok: true, id: id, image: url };
}

function likeEntry_(p) {
  var row = findRow_(p.id);
  if (!row) return { ok: false, error: '글을 찾지 못했습니다' };
  var cell = entriesSheet_().getRange(row, 10);
  var v = (Number(cell.getValue()) || 0) + 1;
  cell.setValue(v);
  return { ok: true, likes: v };
}

function scoreEntry_(p) {
  if (p.code !== ADMIN_CODE) return { ok: false, error: '임원진 코드가 맞지 않습니다' };
  var row = findRow_(p.id);
  if (!row) return { ok: false, error: '글을 찾지 못했습니다' };
  var col = (p.key === 'a') ? 11 : 12;
  entriesSheet_().getRange(row, col).setValue(Number(p.value) || 0);
  return { ok: true };
}

function deleteEntry_(p) {
  if (p.code !== ADMIN_CODE) return { ok: false, error: '임원진 코드가 맞지 않습니다' };
  var row = findRow_(p.id);
  if (!row) return { ok: false, error: '글을 찾지 못했습니다' };
  entriesSheet_().getRange(row, 14).setValue(true);
  return { ok: true };
}

function setTotal_(p) {
  if (p.code !== ADMIN_CODE) return { ok: false, error: '임원진 코드가 맞지 않습니다' };
  setConfig_('totalMembers', Number(p.value) || 0);
  return { ok: true };
}

/* ---------- 설정값 ---------- */
function getConfig_(key) {
  var sh = configSheet_(), last = sh.getLastRow();
  if (last < 2) return 0;
  var rows = sh.getRange(2, 1, last - 1, 2).getValues();
  for (var i = 0; i < rows.length; i++) if (rows[i][0] === key) return Number(rows[i][1]) || 0;
  return 0;
}

function setConfig_(key, value) {
  var sh = configSheet_(), last = sh.getLastRow();
  if (last >= 2) {
    var rows = sh.getRange(2, 1, last - 1, 1).getValues();
    for (var i = 0; i < rows.length; i++) {
      if (rows[i][0] === key) { sh.getRange(i + 2, 2).setValue(value); return; }
    }
  }
  sh.appendRow([key, value]);
}

/* ---------- 사진 ---------- */
function folder_() {
  var it = DriveApp.getFoldersByName(FOLDER_NAME);
  return it.hasNext() ? it.next() : DriveApp.createFolder(FOLDER_NAME);
}

function saveImage_(dataUrl, id) {
  var m = String(dataUrl).match(/^data:(image\/[a-zA-Z+]+);base64,([\s\S]+)$/);
  if (!m) throw new Error('사진 형식을 읽지 못했습니다');
  var blob = Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], id + '.jpg');
  var file = folder_().createFile(blob);
  try {
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (e) { /* 조직 계정에서 막힐 수 있음 */ }
  return 'https://drive.google.com/thumbnail?id=' + file.getId() + '&sz=w1000';
}


/* =========================================================
   2027 임원 선거 (총무 · 회계)
   - 이름 + 비밀번호 4자리를 해시로만 저장해 익명을 유지합니다
   - 같은 이름·비밀번호로 다시 들어오면 수정됩니다
   ========================================================= */

var VOTE_SALT   = 'rejoice-2027-vote-v1';   // 바꾸면 기존 투표를 못 찾습니다
var CANDIDATES  = ['최진솔','최환준','배우리','장민','이준호'];
var MAX_PICK    = 2;
var VOTE_HEAD   = ['key','a','b','ts'];

function votesSheet_() { return sheet_('votes', VOTE_HEAD); }

// 이름+비밀번호를 되돌릴 수 없는 값으로 바꿉니다 (시트에는 이 값만 남습니다)
function voterKey_(name, pw) {
  var raw = String(name || '').trim() + '|' + String(pw || '').trim() + '|' + VOTE_SALT;
  var bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, raw, Utilities.Charset.UTF_8);
  var hex = '';
  for (var i = 0; i < bytes.length; i++) {
    var b = bytes[i] < 0 ? bytes[i] + 256 : bytes[i];
    var h = b.toString(16);
    hex += (h.length < 2 ? '0' + h : h);
  }
  return hex.slice(0, 32);
}

function findVoteRow_(key) {
  var sh = votesSheet_();
  var last = sh.getLastRow();
  if (last < 2) return 0;
  var keys = sh.getRange(2, 1, last - 1, 1).getValues();
  for (var i = 0; i < keys.length; i++) if (String(keys[i][0]) === key) return i + 2;
  return 0;
}

function parsePicks_(v) {
  var out = [];
  String(v || '').split(',').forEach(function (n) {
    n = n.trim();
    if (n && CANDIDATES.indexOf(n) >= 0 && out.indexOf(n) < 0) out.push(n);
  });
  return out.slice(0, MAX_PICK);
}

function checkLogin_(p) {
  var name = String(p.name || '').trim();
  var pw   = String(p.pw || '').trim();
  if (name.length < 2) return '이름을 정확히 입력해 주세요';
  if (!/^[0-9]{4}$/.test(pw)) return '비밀번호는 숫자 네 자리입니다';
  return '';
}

function voteGet_(p) {
  var bad = checkLogin_(p);
  if (bad) return { ok: false, error: bad };

  var row = findVoteRow_(voterKey_(p.name, p.pw));
  if (!row) return { ok: true, voted: false, a: [], b: [] };

  var v = votesSheet_().getRange(row, 1, 1, VOTE_HEAD.length).getValues()[0];
  return { ok: true, voted: true, a: parsePicks_(v[1]), b: parsePicks_(v[2]) };
}

function voteSet_(p) {
  var bad = checkLogin_(p);
  if (bad) return { ok: false, error: bad };

  var a = parsePicks_((p.a || []).join(','));
  var b = parsePicks_((p.b || []).join(','));
  if (!a.length && !b.length) return { ok: false, error: '한 명 이상 선택해 주세요' };

  var key = voterKey_(p.name, p.pw);
  var sh  = votesSheet_();
  var row = findVoteRow_(key);
  var now = new Date().getTime();

  if (row) sh.getRange(row, 1, 1, VOTE_HEAD.length).setValues([[key, a.join(','), b.join(','), now]]);
  else     sh.appendRow([key, a.join(','), b.join(','), now]);

  return { ok: true, a: a, b: b };
}

function voteResult_(p) {
  var open = getConfig_('voteOpen') === 1;
  if (!open && p.code !== ADMIN_CODE) return { ok: false, error: '아직 결과가 공개되지 않았습니다' };

  var sh = votesSheet_();
  var last = sh.getLastRow();
  var ca = {}, cb = {};
  CANDIDATES.forEach(function (n) { ca[n] = 0; cb[n] = 0; });

  if (last >= 2) {
    var rows = sh.getRange(2, 1, last - 1, VOTE_HEAD.length).getValues();
    for (var i = 0; i < rows.length; i++) {
      parsePicks_(rows[i][1]).forEach(function (n) { ca[n]++; });
      parsePicks_(rows[i][2]).forEach(function (n) { cb[n]++; });
    }
  }
  return { ok: true, voters: Math.max(0, last - 1), a: ca, b: cb, open: open };
}

// 임원진이 결과를 모두에게 공개하거나 다시 숨깁니다
function voteOpen_(p) {
  if (p.code !== ADMIN_CODE) return { ok: false, error: '임원진 코드가 맞지 않습니다' };
  setConfig_('voteOpen', p.value ? 1 : 0);
  return { ok: true, open: !!p.value };
}
