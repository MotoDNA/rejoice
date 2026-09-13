/**
 * 리조이스 챌린지 - 서버 (Google Apps Script)
 * 스프레드시트에 글/좋아요/점수를 저장하고, 사진은 구글 드라이브에 저장합니다.
 */

var ADMIN_CODE  = '0914';                 // 임원진 코드 (원하는 숫자로 바꾸세요)
var FOLDER_NAME = '리조이스 챌린지 사진';   // 사진이 저장될 드라이브 폴더 이름

var HEAD = ['id','type','name','text','l1','l2','l3','l4','image','likes','a','b','ts','deleted'];

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

  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    if (action === 'add')    return addEntry_(p);
    if (action === 'like')   return likeEntry_(p);
    if (action === 'score')  return scoreEntry_(p);
    if (action === 'delete') return deleteEntry_(p);
    if (action === 'config') return setTotal_(p);
    return { ok: false, error: 'unknown action: ' + action };
  } finally {
    lock.releaseLock();
  }
}

/* ---------- 시트 ---------- */
function book_() { return SpreadsheetApp.getActiveSpreadsheet(); }

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

function entriesSheet_() { return sheet_('entries', HEAD); }
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

  var id  = 'e' + new Date().getTime() + Math.floor(Math.random() * 1000);
  var url = '';
  if (p.type === 'photo') {
    if (!p.image) return { ok: false, error: '사진이 없습니다' };
    url = saveImage_(p.image, id);
  }
  var L = p.lines || ['', '', '', ''];

  entriesSheet_().appendRow([
    id, p.type, name, String(p.text || '').slice(0, 100),
    L[0] || '', L[1] || '', L[2] || '', L[3] || '',
    url, 0, 0, 0, new Date().getTime(), false
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
