// ===== ตั้งค่า =====
const SPREADSHEET_ID = '1-RkAp6hQbOa_bbfqddR2n1YxN_KnIIO-Us2TImnB_jw';
const SHEET_NAME = 'แผ่นงาน1';
const LOG_SHEET = 'บันทึกสลิป';        // แท็บเก็บประวัติสลิป (สร้างให้เอง)

const FOLDER_ID = '';                  // ID โฟลเดอร์สลิปใน Drive (ถ้าว่าง จะหา/สร้างโฟลเดอร์ชื่อ SLIP_FOLDER)
const SLIP_FOLDER = 'สลิปการจ่ายเงิน';
const GEMINI_KEY = 'ใส่ Gemini API Key';
const MODEL = 'gemini-2.5-flash';      // เช็กชื่อรุ่นที่ใช้ได้ใน AI Studio

const FIRST_ROW = 2;          // แถวแรกที่เป็นข้อมูล
const ID_COL = 2;             // B = เลขประจำตัว
const NAME_COL = 3;           // C = ชื่อ
const SUBJECT_START_COL = 5;  // E = วิชาแรก
const SUBJECT_COUNT = 10;     // E ถึง N

const SUBJECT_NAMES = ['ฟิสิกส์', 'เคมี', 'อังกฤษ', 'สังคม', 'แมท',
  'วิชาที่ 6', 'วิชาที่ 7', 'วิชาที่ 8', 'วิชาที่ 9', 'วิชาที่ 10'];
const SUBJECT_PRICES = [100, 100, 250, 100, 140, 0, 0, 0, 0, 0];

// คอลัมน์ในชีต (ใช้ถึง S เท่านั้น)
const PAY_COL = 15;      // O = ยอดที่โอนมา (รวมทุกสลิป)
const STATUS_COL = 16;   // P = สถานะการจ่าย
const SLIP_COL = 17;     // Q = ลิงก์สลิป
const DATE_COL = 18;     // R = วันที่จ่ายล่าสุด
const CHECK_COL = 19;    // S = ตรวจสอบสลิป (ชื่อ-สกุลผู้โอน + จ่าย/ไม่จ่าย)

const FULL = 'จ่ายครบ';
const PARTIAL = 'จ่ายบางส่วน';
const UNPAID = 'ยังไม่จ่าย';
const COLOR_ROWS = 300;
const ALLOWED_TYPES = ['image/jpeg', 'image/png'];

// ลำดับคอลัมน์ในแท็บ "บันทึกสลิป"
const LOG_HEADERS = ['เวลาบันทึก', 'เลขประจำตัว', 'ชื่อนักเรียน', 'ชื่อ-สกุลผู้โอน (ตามสลิป)',
  'ยอด', 'วันเวลาที่โอน', 'เลขอ้างอิง', 'ไฟล์สลิป', 'ชื่อตรงกัน', 'สถานะรายการ'];
const L = { TIME: 0, ID: 1, NAME: 2, SENDER: 3, AMOUNT: 4, WHEN: 5, KEY: 6, URL: 7, MATCH: 8, STATE: 9 };
const ACTIVE = 'ใช้งาน', VOID = 'ยกเลิก';

// ===== ยูทิลิตี้ =====
function getSpreadsheet() {
  return SPREADSHEET_ID ? SpreadsheetApp.openById(SPREADSHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
}

function getSheet() {
  const ss = getSpreadsheet();
  const sh = ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];
  if (!sh) throw new Error('ไม่พบแท็บในชีต');
  return sh;
}

function getLog() {
  const ss = getSpreadsheet();
  let log = ss.getSheetByName(LOG_SHEET);
  if (!log) {
    log = ss.insertSheet(LOG_SHEET);
    log.getRange(1, 1, 1, LOG_HEADERS.length).setValues([LOG_HEADERS]).setFontWeight('bold');
    log.setFrozenRows(1);
  }
  return log;
}

function colA1(c) {
  let s = '';
  while (c > 0) {
    const m = (c - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    c = Math.floor((c - 1) / 26);
  }
  return s;
}

function subjectNames(sh) {
  return (sh || getSheet()).getRange(1, SUBJECT_START_COL, 1, SUBJECT_COUNT)
    .getDisplayValues()[0]
    .map(h => h.replace(/\s*\d+(\.\d+)?\s*$/, '').trim());
}

function getSlipFolder() {
  if (FOLDER_ID) return DriveApp.getFolderById(FOLDER_ID);
  const it = DriveApp.getFoldersByName(SLIP_FOLDER);
  return it.hasNext() ? it.next() : DriveApp.createFolder(SLIP_FOLDER);
}

// รหัสของนักเรียนในแท็บบันทึก (ใช้เลขประจำตัว ถ้าไม่มีใช้เลขแถว)
function studentKey(sh, row) {
  return String(sh.getRange(row, ID_COL).getDisplayValue()).trim() || ('row' + row);
}

// ===== เมนู =====
function onOpen() {
  SpreadsheetApp.getUi().createMenu('ระบบจ่ายเงิน')
    .addItem('เปิด Dashboard', 'showDashboard')
    .addItem('เพิ่มสลิป (แถวที่เลือก)', 'showUploadDialog')
    .addSeparator()
    .addItem('คำนวณทั้งหมดใหม่', 'recalcAll')
    .addItem('ระบายสีวิชาใหม่', 'applySubjectColors')
    .addItem('ซิงก์สลิปจาก Drive', 'syncSlipsFromDrive')
    .addItem('รีเซ็ตเป็นยังไม่จ่าย (แถวที่เลือก)', 'markUnpaid')
    .addItem('รีเซ็ตการจ่ายทั้งหมด', 'resetAllPayments')
    .addSeparator()
    .addItem('สรุปสถานะการจ่าย', 'showSummary')
    .addToUi();
}

function showDashboard() {
  const html = HtmlService.createHtmlOutputFromFile('Dashboard').setWidth(1200).setHeight(650);
  SpreadsheetApp.getUi().showModelessDialog(html, 'ระบบจ่ายเงิน');
}

function showUploadDialog() {
  openUploadForRow(SpreadsheetApp.getActiveRange().getRow());
}

// เปิดหน้าอัปโหลดของแถวนั้น (เรียกจากเมนูและจาก Dashboard)
function openUploadForRow(row) {
  row = Number(row);
  const sh = getSheet();
  if (!row || row < FIRST_ROW || !sh.getRange(row, NAME_COL).getValue()) {
    SpreadsheetApp.getUi().alert('กรุณาเลือกแถวของนักเรียนก่อน');
    return;
  }
  const t = HtmlService.createTemplateFromFile('Upload');
  t.row = row;
  t.name = sh.getRange(row, NAME_COL).getDisplayValue();
  SpreadsheetApp.getUi().showModalDialog(t.evaluate().setWidth(480).setHeight(620), 'อัปโหลดสลิป');
}

// ===== ตั้งค่าครั้งแรก =====
function setup() {
  const sh = getSheet();
  const n = Math.max(sh.getLastRow(), FIRST_ROW) - FIRST_ROW + 1;

  const subj = sh.getRange(1, SUBJECT_START_COL, 1, SUBJECT_COUNT);
  subj.setValues([subj.getValues()[0].map((v, i) =>
    v === '' ? `${SUBJECT_NAMES[i]} ${SUBJECT_PRICES[i]}` : v)]).setFontWeight('bold');

  sh.getRange(1, PAY_COL, 1, 5).setValues([[
    'ยอดที่โอนมา', 'สถานะการจ่าย', 'ลิงก์สลิป', 'วันที่จ่ายล่าสุด', 'ตรวจสอบสลิป']]).setFontWeight('bold');

  const prices = sh.getRange(FIRST_ROW, SUBJECT_START_COL, n, SUBJECT_COUNT);
  prices.setValues(prices.getValues().map(r => r.map((v, i) => v === '' ? SUBJECT_PRICES[i] : v)));

  getLog();
  applySubjectColors();
  recalcAll();
}

// ===== คำนวณ =====
// ตัดวิชาตามลำดับ จนกว่าวิชาถัดไปจะไม่พอ
function allocate(prices, names, paid) {
  let remaining = paid, paidSum = 0, stopped = false;
  const flags = [];
  const total = prices.reduce((a, b) => a + (Number(b) || 0), 0);

  prices.forEach(price => {
    const p = Number(price) || 0;
    if (p <= 0) { flags.push('-'); return; }
    if (!stopped && remaining >= p) { remaining -= p; paidSum += p; flags.push('paid'); }
    else { stopped = true; flags.push('owe'); }
  });

  const owed = total - paidSum;
  const status = total <= 0 ? UNPAID : owed <= 0 ? FULL : paidSum > 0 ? PARTIAL : UNPAID;
  const paidList = names.filter((_, i) => flags[i] === 'paid').join(', ') || '-';
  return { total, owed, status, paidList };
}

// ข้อความสถานะในคอลัมน์ P
function statusText(r) {
  if (r.total <= 0) return 'ยังไม่ตั้งราคา';
  if (r.status === FULL) return FULL;
  return `${r.status} ค้าง ${r.owed}`;
}

// คำสั้น ๆ จ่าย/ไม่จ่าย สำหรับสรุป
function payLabel(r) {
  if (r.total <= 0) return 'ยังไม่ตั้งราคา';
  if (r.status === FULL) return 'จ่ายแล้ว';
  if (r.status === PARTIAL) return `จ่ายบางส่วน (ค้าง ${r.owed})`;
  return 'ไม่จ่าย';
}

// ข้อความในคอลัมน์ S: ชื่อ-สกุลผู้โอน · จ่าย/ไม่จ่าย
function checkText(r, info) {
  if (!info) return payLabel(r);
  return `${info.sender || 'ไม่ทราบชื่อ'} · ${payLabel(r)}${info.match === 'ไม่ตรง' ? ' ⚠ชื่อไม่ตรง' : ''}`;
}

// สลิปล่าสุดที่ใช้งานอยู่ของแต่ละคน: { เลขประจำตัว: {sender, match} }
function slipInfoMap() {
  const log = getLog();
  const last = log.getLastRow();
  const map = {};
  if (last < 2) return map;
  log.getRange(2, 1, last - 1, LOG_HEADERS.length).getValues().forEach(r => {
    if (r[L.STATE] === ACTIVE) map[String(r[L.ID])] = { sender: r[L.SENDER], match: r[L.MATCH] };
  });
  return map;
}

function recalcRow(row, sh, infoMap) {
  sh = sh || getSheet();
  infoMap = infoMap || slipInfoMap();
  const prices = sh.getRange(row, SUBJECT_START_COL, 1, SUBJECT_COUNT).getValues()[0];
  const paid = Number(sh.getRange(row, PAY_COL).getValue()) || 0;
  const res = allocate(prices, subjectNames(sh), paid);
  sh.getRange(row, STATUS_COL).setValue(statusText(res));
  sh.getRange(row, CHECK_COL).setValue(checkText(res, infoMap[studentKey(sh, row)]));
  return { status: statusText(res), check: checkText(res, infoMap[studentKey(sh, row)]) };
}

function recalcAll() {
  const sh = getSheet();
  const last = sh.getLastRow();
  if (last < FIRST_ROW) return;
  const n = last - FIRST_ROW + 1;
  const names = subjectNames(sh);
  const infoMap = slipInfoMap();
  const ids = sh.getRange(FIRST_ROW, ID_COL, n, 1).getDisplayValues();
  const who = sh.getRange(FIRST_ROW, NAME_COL, n, 1).getValues();
  const prices = sh.getRange(FIRST_ROW, SUBJECT_START_COL, n, SUBJECT_COUNT).getValues();
  const paid = sh.getRange(FIRST_ROW, PAY_COL, n, 1).getValues();

  const status = [], check = [];
  prices.forEach((p, i) => {
    if (!who[i][0]) { status.push(['']); check.push(['']); return; }
    const res = allocate(p, names, Number(paid[i][0]) || 0);
    const key = String(ids[i][0]).trim() || ('row' + (FIRST_ROW + i));
    status.push([statusText(res)]);
    check.push([checkText(res, infoMap[key])]);
  });
  sh.getRange(FIRST_ROW, STATUS_COL, n, 1).setValues(status);
  sh.getRange(FIRST_ROW, CHECK_COL, n, 1).setValues(check);
}

// แก้ราคาหรือยอดโอนด้วยมือ แล้วคำนวณแถวนั้นใหม่
function onEdit(e) {
  if (!e || !e.range) return;
  const sh = e.range.getSheet();
  const row = e.range.getRow(), col = e.range.getColumn();
  if (sh.getName() !== SHEET_NAME || row < FIRST_ROW) return;
  if (col === PAY_COL || (col >= SUBJECT_START_COL && col < SUBJECT_START_COL + SUBJECT_COUNT)) {
    recalcRow(row, sh);
  }
}

// เขียว = จ่ายแล้ว, แดง = ค้าง
function applySubjectColors() {
  const sh = getSheet();
  const range = sh.getRange(FIRST_ROW, SUBJECT_START_COL, COLOR_ROWS, SUBJECT_COUNT);
  const e = colA1(SUBJECT_START_COL), pay = colA1(PAY_COL), r = FIRST_ROW;
  const rule = (op, color) => SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(`=AND(${e}${r}>0,$${pay}${r}${op}SUM($${e}${r}:${e}${r}))`)
    .setBackground(color).setRanges([range]).build();
  sh.setConditionalFormatRules([rule('>=', '#c6efce'), rule('<', '#f4c7c3')]);
}

// ===== ตรวจชื่อ-สกุลในสลิปกับชื่อนักเรียน =====
// ตัดคำนำหน้า แล้วเทียบชื่อต้น (สลิปมักย่อนามสกุล จึงไม่เทียบนามสกุล)
function firstName(full) {
  const s = String(full || '').trim()
    .replace(/^(นางสาว|นาย|นาง|เด็กชาย|เด็กหญิง|ด\.ช\.|ด\.ญ\.|น\.ส\.|นส\.|mr\.?|mrs\.?|miss|ms\.?)\s*/i, '');
  return (s.split(/\s+/)[0] || '').replace(/[.\-]/g, '').toLowerCase();
}

function nameMatch(sender, studentName) {
  const a = firstName(sender), b = firstName(studentName);
  if (!a) return 'ไม่ทราบ';
  if (a.length < 2 || b.length < 2) return 'ไม่ทราบ';
  return (a.indexOf(b) !== -1 || b.indexOf(a) !== -1) ? 'ตรง' : 'ไม่ตรง';
}

// ===== อ่านสลิปด้วย AI (ขั้นที่ 1: ยังไม่บันทึกอะไร) =====
function readSlip(base64, mimeType) {
  if (ALLOWED_TYPES.indexOf(mimeType) === -1) throw new Error('รองรับเฉพาะไฟล์ .jpg และ .png');

  const prompt =
    'นี่คือสลิปโอนเงินของธนาคารไทยหรือ TrueMoney ให้อ่านข้อมูลตามที่เห็นในรูปเท่านั้น ' +
    'ห้ามเดา ถ้าอ่านไม่ออกให้ใส่ null. sender คือชื่อ-สกุลของผู้โอน (ฝั่งผู้โอน ไม่ใช่ผู้รับ). ' +
    'amount คือยอดโอนเป็นตัวเลข ไม่มีเครื่องหมายจุลภาค. ' +
    'datetime ให้ใช้รูปแบบ YYYY-MM-DD HH:mm (ปี ค.ศ.). ref คือเลขที่รายการ/เลขอ้างอิง.';
  const body = {
    contents: [{ parts: [{ text: prompt }, { inlineData: { mimeType: mimeType, data: base64 } }] }],
    generationConfig: {
      temperature: 0,
      responseMimeType: 'application/json',
      responseSchema: {
        type: 'OBJECT',
        properties: {
          isSlip:   { type: 'BOOLEAN' },
          bank:     { type: 'STRING', nullable: true },
          amount:   { type: 'NUMBER', nullable: true },
          sender:   { type: 'STRING', nullable: true },
          receiver: { type: 'STRING', nullable: true },
          datetime: { type: 'STRING', nullable: true },
          ref:      { type: 'STRING', nullable: true }
        },
        required: ['isSlip']
      }
    }
  };
  const url = 'https://generativelanguage.googleapis.com/v1beta/models/' + MODEL + ':generateContent';

  let res;
  for (let i = 0; i < 3; i++) {          // ลองใหม่ถ้าชน rate limit
    res = UrlFetchApp.fetch(url, {
      method: 'post', contentType: 'application/json',
      headers: { 'x-goog-api-key': GEMINI_KEY },
      payload: JSON.stringify(body), muteHttpExceptions: true
    });
    if (res.getResponseCode() !== 429) break;
    Utilities.sleep(3000 * (i + 1));
  }
  if (res.getResponseCode() !== 200) {
    throw new Error('AI อ่านสลิปไม่สำเร็จ: ' + res.getContentText().slice(0, 300));
  }
  const text = JSON.parse(res.getContentText()).candidates[0].content.parts[0].text;
  return JSON.parse(text);
}

// ตัวอย่างสรุปก่อนบันทึก: ชื่อตรงไหม และสถานะหลังบันทึกจะเป็นอย่างไร
function previewSlip(row, amount, sender) {
  row = Number(row);
  const sh = getSheet();
  const studentName = String(sh.getRange(row, NAME_COL).getDisplayValue());
  const prices = sh.getRange(row, SUBJECT_START_COL, 1, SUBJECT_COUNT).getValues()[0];
  const paid = (Number(sh.getRange(row, PAY_COL).getValue()) || 0) + (Number(amount) || 0);
  const res = allocate(prices, subjectNames(sh), paid);
  return {
    studentName: studentName,
    match: nameMatch(sender, studentName),
    label: payLabel(res),
    status: res.status
  };
}

// ===== บันทึกสลิป (ขั้นที่ 2: หลังคนตรวจและยืนยันแล้ว) =====
function saveSlip(row, data, base64, fileName, mimeType) {
  row = Number(row);
  const amount = Number(data && data.amount);
  if (!row || row < FIRST_ROW) throw new Error('ไม่พบแถวนักเรียนนี้');
  if (!(amount > 0)) throw new Error('ยอดเงินไม่ถูกต้อง');
  const ext = String(fileName).split('.').pop().toLowerCase();
  if (ALLOWED_TYPES.indexOf(mimeType) === -1 || ['jpg', 'jpeg', 'png'].indexOf(ext) === -1) {
    throw new Error('รองรับเฉพาะไฟล์ .jpg และ .png');
  }

  const bytes = Utilities.base64Decode(base64);
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = getSheet();
    const studentName = String(sh.getRange(row, NAME_COL).getDisplayValue());
    if (!studentName) throw new Error('ไม่พบนักเรียนในแถวนี้');

    // กันสลิปซ้ำ: เลขอ้างอิง > ยอด+เวลา > ลายนิ้วมือของไฟล์ (เทียบกับรายการที่ยังใช้งาน)
    const ref = data.ref && String(data.ref).trim();
    const key = ref ||
      (data.datetime ? amount + '|' + data.datetime
        : Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, bytes)));
    const log = getLog();
    const logLast = log.getLastRow();
    if (logLast >= 2) {
      const used = log.getRange(2, 1, logLast - 1, LOG_HEADERS.length).getValues()
        .some(r => r[L.STATE] === ACTIVE && String(r[L.KEY]) === key);
      if (used) throw new Error('สลิปนี้เคยถูกใช้แล้ว');
    }

    // เซฟไฟล์ โดยตั้งชื่อให้มีเลขประจำตัว (ให้ซิงก์จาก Drive หาเจอ)
    const id = studentKey(sh, row);
    const stamp = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyyMMdd-HHmmss');
    const newName = `${id}_${stamp}.${ext === 'jpeg' ? 'jpg' : ext}`;
    const file = getSlipFolder().createFile(Utilities.newBlob(bytes, mimeType, newName));

    // วันที่ตามสลิป ถ้าอ่านไม่ได้ใช้เวลาปัจจุบัน
    let when = data.datetime ? new Date(String(data.datetime).replace(' ', 'T') + ':00+07:00') : new Date();
    if (isNaN(when.getTime())) when = new Date();

    const total = (Number(sh.getRange(row, PAY_COL).getValue()) || 0) + amount;  // สะสมทุกสลิป
    sh.getRange(row, PAY_COL).setValue(total);
    sh.getRange(row, SLIP_COL).setFormula('=HYPERLINK("' + file.getUrl() + '","ดูสลิป")');
    sh.getRange(row, DATE_COL).setValue(when).setNumberFormat('dd/mm/yyyy hh:mm');

    log.appendRow([new Date(), id, studentName, data.sender || '', amount,
      Utilities.formatDate(when, 'Asia/Bangkok', 'yyyy-MM-dd HH:mm'), key, file.getUrl(),
      nameMatch(data.sender, studentName), ACTIVE]);

    return recalcRow(row, sh).check;
  } finally {
    lock.releaseLock();
  }
}

// ===== รีเซ็ต =====
// ยกเลิกรายการในแท็บบันทึก (ไม่ลบ เก็บเป็นประวัติ) เพื่อให้สลิปเดิมอัปโหลดใหม่ได้
function voidLog(idOrNull) {
  const log = getLog();
  const last = log.getLastRow();
  if (last < 2) return;
  const range = log.getRange(2, 1, last - 1, LOG_HEADERS.length);
  const vals = range.getValues();
  vals.forEach(r => {
    if (r[L.STATE] === ACTIVE && (idOrNull === null || String(r[L.ID]) === idOrNull)) r[L.STATE] = VOID;
  });
  range.setValues(vals);
}

function resetRow(row) {
  row = Number(row);
  if (!row || row < FIRST_ROW) throw new Error('ไม่พบแถวนักเรียนนี้');
  const sh = getSheet();
  voidLog(studentKey(sh, row));
  sh.getRange(row, PAY_COL).setValue(0);
  sh.getRange(row, SLIP_COL, 1, 2).clearContent();   // ลิงก์ + วันที่
  recalcRow(row, sh);
}

function markUnpaid() {
  const row = SpreadsheetApp.getActiveRange().getRow();
  if (row >= FIRST_ROW) resetRow(row);
}

function resetAllPayments() {
  const ui = SpreadsheetApp.getUi();
  const sh = getSheet();
  const last = sh.getLastRow();
  if (last < FIRST_ROW) return;
  const ok = ui.alert('รีเซ็ตการจ่ายทั้งหมด',
    'ยอดโอน ลิงก์สลิป และวันที่ของทุกคนจะถูกล้าง\n(ไฟล์สลิปใน Drive และแท็บบันทึกสลิปยังอยู่)\nต้องการดำเนินการต่อหรือไม่?',
    ui.ButtonSet.YES_NO);
  if (ok !== ui.Button.YES) return;
  const n = last - FIRST_ROW + 1;
  voidLog(null);
  sh.getRange(FIRST_ROW, PAY_COL, n, 1).setValue(0);
  sh.getRange(FIRST_ROW, SLIP_COL, n, 2).clearContent();
  recalcAll();
  ui.alert('รีเซ็ตการจ่ายทั้งหมดเรียบร้อยแล้ว');
}

// ===== Dashboard =====
function getDashboardData() {
  const sh = getSheet();
  const last = sh.getLastRow();
  if (last < FIRST_ROW) return { students: [] };

  const n = last - FIRST_ROW + 1;
  const data = sh.getRange(FIRST_ROW, 1, n, DATE_COL).getValues();
  const slipFormulas = sh.getRange(FIRST_ROW, SLIP_COL, n, 1).getFormulas();
  const names = subjectNames(sh);
  const infoMap = slipInfoMap();

  const students = [];
  data.forEach((r, i) => {
    if (!r[NAME_COL - 1]) return;
    const prices = r.slice(SUBJECT_START_COL - 1, SUBJECT_START_COL - 1 + SUBJECT_COUNT);
    const paid = Number(r[PAY_COL - 1]) || 0;
    const res = allocate(prices, names, paid);
    const m = String(slipFormulas[i][0]).match(/HYPERLINK\("([^"]+)"/);
    const date = r[DATE_COL - 1];
    const id = String(r[ID_COL - 1] || '');
    const info = infoMap[id.trim() || ('row' + (FIRST_ROW + i))];
    students.push({
      row: FIRST_ROW + i,
      id: id,
      name: String(r[NAME_COL - 1]),
      total: res.total, status: res.status, paid: paid, owed: res.owed,
      paidList: res.paidList,
      sender: info ? String(info.sender || '') : '',
      match: info ? String(info.match || '') : '',
      slip: m ? m[1] : '',
      date: date instanceof Date ? Utilities.formatDate(date, 'Asia/Bangkok', 'dd/MM/yyyy HH:mm') : ''
    });
  });
  return { students };
}

function showSummary() {
  const students = getDashboardData().students;
  if (!students.length) return;
  let full = 0, partial = 0, owed = 0;
  const unpaid = [];
  students.forEach(s => {
    owed += s.owed;
    if (s.status === FULL) full++;
    else if (s.status === PARTIAL) partial++;
    else unpaid.push(s.name);
  });
  let msg = `จ่ายครบ: ${full} / ${students.length} คน\nจ่ายบางส่วน: ${partial} คน\n` +
    `ยังไม่จ่าย: ${unpaid.length} คน\nยอดค้างชำระรวม: ${owed} บาท\n\n`;
  msg += unpaid.length
    ? 'รายชื่อยังไม่จ่าย:\n' + unpaid.map((n, i) => `${i + 1}. ${n}`).join('\n')
    : 'ไม่มีคนที่ยังไม่จ่าย';
  SpreadsheetApp.getUi().alert('สรุปสถานะการจ่ายเงิน', msg, SpreadsheetApp.getUi().ButtonSet.OK);
}

// ===== ซิงก์สลิปจากโฟลเดอร์ Drive (จับคู่จากเลขประจำตัวในชื่อไฟล์) =====
function syncSlipsFromDrive() {
  const ui = SpreadsheetApp.getUi();
  const sh = getSheet();
  const last = sh.getLastRow();
  if (last < FIRST_ROW) return;

  const files = [];
  const it = getSlipFolder().getFiles();
  while (it.hasNext()) files.push(it.next());

  const n = last - FIRST_ROW + 1;
  const ids = sh.getRange(FIRST_ROW, ID_COL, n, 1).getDisplayValues();
  const names = sh.getRange(FIRST_ROW, NAME_COL, n, 1).getDisplayValues();
  let linked = 0;
  const missing = [];

  for (let i = 0; i < n; i++) {
    if (!names[i][0]) continue;
    const id = String(ids[i][0]).trim();
    const matches = id ? files.filter(f => f.getName().indexOf(id) !== -1) : [];
    if (!matches.length) { missing.push(names[i][0]); continue; }
    matches.sort((a, b) => b.getLastUpdated() - a.getLastUpdated());
    const row = FIRST_ROW + i;
    sh.getRange(row, SLIP_COL).setFormula('=HYPERLINK("' + matches[0].getUrl() + '","ดูสลิป")');
    sh.getRange(row, DATE_COL).setValue(matches[0].getLastUpdated()).setNumberFormat('dd/mm/yyyy hh:mm');
    linked++;
  }

  let msg = `เชื่อมสลิปสำเร็จ: ${linked} คน\nไม่พบสลิป: ${missing.length} คน`;
  if (missing.length) {
    msg += '\n\n' + missing.slice(0, 20).map((m, i) => `${i + 1}. ${m}`).join('\n');
    if (missing.length > 20) msg += `\n... และอีก ${missing.length - 20} คน`;
  }
  ui.alert('ซิงก์สลิปจาก Drive', msg, ui.ButtonSet.OK);
}
