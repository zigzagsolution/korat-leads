/**
 * ตัวเก็บข้อมูลปฏิทินติดตามงานทีม (อวดเมืองโคราช)
 * วางโค้ดนี้ใน Google Sheet > ส่วนขยาย > Apps Script แล้ว Deploy เป็น Web app
 * (Execute as: Me, Who has access: Anyone)
 */

// ===== รหัสเข้าปฏิทิน: เปลี่ยนตรงนี้ได้เลย (คนเปิดเว็บมองไม่เห็นโค้ดนี้) =====
// แนะนำให้ยาว 6 ตัวขึ้นไป เปลี่ยนแล้วทุกคนต้องกรอกรหัสใหม่
const PASSCODE = '1111';
// รหัสสำหรับลูกค้า: ดูได้อย่างเดียว แก้ไม่ได้ (ต้องไม่ซ้ำกับรหัสทีม) ไม่เห็นชื่อคนและข้อความหมายเหตุ
const VIEW_PASSCODE = '2222';

const SHEET = 'tasks';
const HEAD = ['id', 'date', 'end', 'name', 'task', 'status', 'note', 'created', 'updated', 'extra'];
// คอลัมน์ extra เก็บข้อมูลเพิ่มของฟังก์ชันใหม่ในอนาคต (JSON) จึงไม่ต้องแก้โค้ดนี้บ่อย

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET);
  if (!sh) {
    sh = ss.insertSheet(SHEET);
    sh.getRange(1, 1, 1, HEAD.length).setValues([HEAD]).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.getRange('B:C').setNumberFormat('@');
  }
  // ถ้าหัวตารางยังไม่มีคอลัมน์ใหม่ ให้เติมเองอัตโนมัติ
  const have = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getValues()[0];
  if (have.length < HEAD.length) sh.getRange(1, 1, 1, HEAD.length).setValues([HEAD]).setFontWeight('bold');
  return sh;
}

function rows_() {
  const sh = sheet_();
  const n = sh.getLastRow() - 1;
  if (n < 1) return [];
  return sh.getRange(2, 1, n, HEAD.length).getDisplayValues()
    .filter(r => r[0])
    .map(r => {
      const o = Object.fromEntries(HEAD.map((h, i) => [h, r[i]]));
      try { o.extra = o.extra ? JSON.parse(o.extra) : {}; } catch (e) { o.extra = {}; }
      return o;
    });
}

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function clean_(v, max) { return String(v == null ? '' : v).slice(0, max); }

const VERSION = 4;
const FOLDER = 'ปฏิทินงานทีม-ไฟล์แนบ';

// เก็บไฟล์แนบไว้ใน Google Drive ของเจ้าของชีต แล้วเปิดให้คนที่มีลิงก์ดูได้
function upload_(b) {
  const it = DriveApp.getFoldersByName(FOLDER);
  const folder = it.hasNext() ? it.next() : DriveApp.createFolder(FOLDER);
  const name = clean_(b.name || 'file', 200);
  const blob = Utilities.newBlob(Utilities.base64Decode(b.data || ''), b.mime || 'application/octet-stream', name);
  const f = folder.createFile(blob);
  try { f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}
  return { ok: true, file: { name: name, url: f.getUrl(), id: f.getId() } };
}

// ตรวจรหัส: ถ้าผิดเกิน 20 ครั้งใน 10 นาที จะล็อกชั่วคราว กันการเดารหัส
function auth_(k) {
  const c = CacheService.getScriptCache();
  const fails = Number(c.get('fails') || 0);
  if (fails >= 20) return false;
  if (String(k || '') === PASSCODE) return true;
  c.put('fails', String(fails + 1), 600);
  return false;
}

function doGet(e) {
  const k = e && e.parameter && e.parameter.k;
  if (VIEW_PASSCODE && VIEW_PASSCODE !== PASSCODE && String(k || '') === VIEW_PASSCODE)
    return out_({ ok: true, auth: true, view: true, version: VERSION, tasks: viewRows_() });
  if (!auth_(k)) return out_({ ok: false, error: 'auth' });
  return out_({ ok: true, auth: true, version: VERSION, tasks: rows_() });
}

// ข้อมูลฉบับลูกค้า: ตัดชื่อคน ตัดข้อความหมายเหตุ (เหลือเฉพาะลิงก์) ตัดงานที่ติ๊กซ่อนจากลูกค้า
function viewRows_() {
  // ซ่อนงานที่ตั้งเป็น private/hidden แต่เก็บแถวกิจกรรมที่อ้าง ref ไว้ (เพื่อให้กิจกรรมตั้งต้นที่ซ่อน/แก้ไว้ แสดงตรงกับปฏิทินทีม)
  return rows_().filter(r => { const x = r.extra || {}; if (x.private) return false; if (x.hidden) return x.kind === 'event' && !!x.ref; return true; }).map(r => {
    const links = String(r.note || '').match(/https?:\/\/[^\s<>"']+/g) || [];
    const x = r.extra || {};
    return { id: r.id, date: r.date, end: r.end, task: r.task, status: r.status, note: links.join(' '),
      extra: { kind: x.kind, ref: x.ref, hidden: x.hidden, type: x.type, tor: x.tor, axis: x.axis, channels: x.channels, files: x.files } };
  });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const b = JSON.parse(e.postData.contents || '{}');
    if (!auth_(b.k)) return out_({ ok: false, error: 'auth' });
    if (b.action === 'upload') return out_(upload_(b));
    const sh = sheet_();
    const now = new Date().toISOString();
    const t = b.task || {};
    const okStatus = ['todo', 'doing', 'done'];
    const rec = {
      date: clean_(t.date, 10), end: clean_(t.end, 10), name: clean_(t.name, 60),
      task: clean_(t.task, 300), status: okStatus.indexOf(t.status) >= 0 ? t.status : 'todo',
      note: clean_(t.note, 500),
      extra: t.extra ? clean_(JSON.stringify(t.extra), 5000) : ''
    };
    const ids = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().map(r => String(r[0])) : [];
    if (b.action === 'add') {
      const id = Utilities.getUuid().slice(0, 8);
      sh.appendRow([id, rec.date, rec.end, rec.name, rec.task, rec.status, rec.note, now, now, rec.extra]);
    } else if (b.action === 'update' || b.action === 'status') {
      const i = ids.indexOf(String(b.id));
      if (i < 0) return out_({ ok: false, error: 'not_found', tasks: rows_() });
      if (b.action === 'status') {
        sh.getRange(i + 2, 6).setValue(rec.status);
      } else {
        sh.getRange(i + 2, 2, 1, 6).setValues([[rec.date, rec.end, rec.name, rec.task, rec.status, rec.note]]);
      }
      if (b.action === 'update' && t.extra !== undefined) sh.getRange(i + 2, 10).setValue(rec.extra);
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
