/**
 * ตัวเก็บข้อมูลปฏิทินติดตามงานทีม (อวดเมืองโคราช)
 * วางโค้ดนี้ใน Google Sheet > ส่วนขยาย > Apps Script แล้ว Deploy เป็น Web app
 * (Execute as: Me, Who has access: Anyone)
 */
const SHEET = 'tasks';
const HEAD = ['id', 'date', 'end', 'name', 'task', 'status', 'note', 'created', 'updated'];

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET);
  if (!sh) {
    sh = ss.insertSheet(SHEET);
    sh.getRange(1, 1, 1, HEAD.length).setValues([HEAD]).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.getRange('B:C').setNumberFormat('@');
  }
  return sh;
}

function rows_() {
  const sh = sheet_();
  const n = sh.getLastRow() - 1;
  if (n < 1) return [];
  return sh.getRange(2, 1, n, HEAD.length).getDisplayValues()
    .filter(r => r[0])
    .map(r => Object.fromEntries(HEAD.map((h, i) => [h, r[i]])));
}

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function clean_(v, max) { return String(v == null ? '' : v).slice(0, max); }

function doGet() {
  return out_({ ok: true, tasks: rows_() });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const b = JSON.parse(e.postData.contents || '{}');
    const sh = sheet_();
    const now = new Date().toISOString();
    const t = b.task || {};
    const okStatus = ['todo', 'doing', 'done'];
    const rec = {
      date: clean_(t.date, 10), end: clean_(t.end, 10), name: clean_(t.name, 60),
      task: clean_(t.task, 300), status: okStatus.indexOf(t.status) >= 0 ? t.status : 'todo',
      note: clean_(t.note, 500)
    };
    const ids = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().map(r => String(r[0])) : [];
    if (b.action === 'add') {
      const id = Utilities.getUuid().slice(0, 8);
      sh.appendRow([id, rec.date, rec.end, rec.name, rec.task, rec.status, rec.note, now, now]);
    } else if (b.action === 'update' || b.action === 'status') {
      const i = ids.indexOf(String(b.id));
      if (i < 0) return out_({ ok: false, error: 'not_found', tasks: rows_() });
      if (b.action === 'status') {
        sh.getRange(i + 2, 6).setValue(rec.status);
      } else {
        sh.getRange(i + 2, 2, 1, 6).setValues([[rec.date, rec.end, rec.name, rec.task, rec.status, rec.note]]);
      }
      sh.getRange(i + 2, 9).setValue(now);
    } else if (b.action === 'delete') {
      const i = ids.indexOf(String(b.id));
      if (i >= 0) sh.deleteRow(i + 2);
    }
    return out_({ ok: true, tasks: rows_() });
  } finally {
    lock.releaseLock();
  }
}
