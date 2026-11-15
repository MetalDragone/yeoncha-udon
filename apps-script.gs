// 연차 2개 쓰고 어떡행: 날짜 투표와 방명록 저장
// 구글 드라이브에 '어떡행 답변' 시트가 생기고 '투표', '방명록' 탭에 쌓입니다. 행을 지우면 페이지에서도 사라집니다.

var MAX_MEMO = 200;
var NONE = '이번엔 어려워요';

function book_() {
  var a = null;
  try { a = SpreadsheetApp.getActiveSpreadsheet(); } catch (e) {}
  if (a) return a;
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('SHEET_ID');
  if (id) return SpreadsheetApp.openById(id);
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    id = props.getProperty('SHEET_ID');
    if (id) return SpreadsheetApp.openById(id);
    var made = SpreadsheetApp.create('어떡행 답변');
    props.setProperty('SHEET_ID', made.getId());
    return made;
  } finally { lock.releaseLock(); }
}

function sheet_(name, header) {
  var ss = book_();
  var sh = ss.getSheetByName(name);
  if (!sh) { sh = ss.insertSheet(name); sh.appendRow(header); sh.setFrozenRows(1); }
  return sh;
}
function voteSheet_() { return sheet_('투표', ['시간', '이름', '인원', '가능한 날짜 (1월)', '하고 싶은 말']); }
function guestSheet_() { return sheet_('방명록', ['시간', '이름', '메시지']); }

function clean_(v, n) {
  var s = String(v == null ? '' : v).replace(/[\u0000-\u0009\u000b-\u001f]/g, ' ').trim().slice(0, n);
  if (/^[=+\-@]/.test(s)) s = "'" + s;
  return s;
}

// 날짜 후보는 '14-17'처럼 1월 날짜 범위로 저장
function plans_(v) {
  if (!Array.isArray(v)) return [];
  var seen = {};
  return v.map(String).filter(function (k) {
    if (!/^\d{1,2}-\d{1,2}$/.test(k) || seen[k]) return false;
    seen[k] = 1; return true;
  }).slice(0, 8);
}

function read_() {
  var votes = voteSheet_().getDataRange().getValues().slice(1)
    .filter(function (x) { return x[1]; })
    .map(function (x) {
      var p = String(x[3] || '');
      return { time: new Date(x[0]).toISOString(), name: String(x[1]), size: Number(x[2]) === 1 ? 1 : 2,
               plans: p === NONE ? [] : plans_(p.split(/\s*,\s*/)), memo: String(x[4] || '') };
    });
  var g = guestSheet_().getDataRange().getValues().slice(1)
    .filter(function (x) { return x[1] && x[2]; })
    .map(function (x) { return { time: new Date(x[0]).toISOString(), name: String(x[1]), msg: String(x[2]) }; })
    .reverse().slice(0, 100);
  return { votes: votes, guestbook: g };
}

function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function doGet() { return out_({ ok: true, data: read_() }); }

function doPost(e) {
  book_();
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var b = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (b.type === 'vote') {
      var name = clean_(b.name, 20);
      if (!name) return out_({ ok: false, error: '이름을 적어 주세요.' });
      var p = plans_(b.plans);
      var row = [new Date(), name, b.size === '1' ? 1 : 2, p.length ? p.join(', ') : NONE, clean_(b.memo, MAX_MEMO)];
      var sh = voteSheet_();
      var last = sh.getLastRow();
      var names = last > 1 ? sh.getRange(2, 2, last - 1, 1).getValues().map(function (x) { return String(x[0]); }) : [];
      var i = names.indexOf(name.replace(/^'/, ''));
      if (i >= 0) sh.getRange(i + 2, 1, 1, 5).setValues([row]); // 같은 이름이면 투표를 바꿈
      else sh.appendRow(row);
    } else if (b.type === 'guest') {
      var gname = clean_(b.name, 20), msg = clean_(b.msg, MAX_MEMO);
      if (!gname || !msg) return out_({ ok: false, error: '이름과 남길 말을 적어 주세요.' });
      guestSheet_().appendRow([new Date(), gname, msg]);
    } else {
      return out_({ ok: false, error: '알 수 없는 요청이에요.' });
    }
    return out_({ ok: true, data: read_() });
  } catch (err) {
    return out_({ ok: false, error: '저장하지 못했어요. 잠시 후 다시 해 주세요.' });
  } finally {
    lock.releaseLock();
  }
}
