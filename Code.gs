const SPREADSHEET_ID = '1ShJrfy912ab274U4EiDtXgIqXFkuZ4HbLui3181aguE';
const SHEET_NAME = 'MuseumVisits';

// เปลี่ยนรหัสผ่านได้ที่ Apps Script > Project Settings > Script Properties
// ตั้งค่า AUTH_USER และ AUTH_PASS ใน Script Properties เพื่อไม่เปิดเผยรหัสผ่านในหน้าเว็บ
const DEFAULT_USER = 'library';
const DEFAULT_PASS = 'library5413';

const HEADERS = [
  'Timestamp',
  'GroupName',
  'VisitorCount',
  'VisitDateTime',
  'PackageType'
];

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('ระบบบันทึกการเยี่ยมชมพิพิธภัณฑ์ รร.ตท.สปท.')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function setupSheet_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, HEADERS.length)
      .setFontWeight('bold')
      .setBackground('#6b21a8')
      .setFontColor('#ffffff');
  } else {
    const firstRow = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), HEADERS.length))
      .getDisplayValues()[0];
    if (String(firstRow[0]).trim() !== HEADERS[0]) {
      throw new Error('แถวแรกของชีตไม่ตรงกับหัวตารางที่ระบบต้องการ กรุณาสำรองข้อมูลก่อน แล้วตรวจสอบหัวตาราง');
    }
  }
  return sheet;
}

function checkLogin(username, password) {
  const props = PropertiesService.getScriptProperties();
  const savedUser = props.getProperty('AUTH_USER') || DEFAULT_USER;
  const savedPass = props.getProperty('AUTH_PASS') || DEFAULT_PASS;

  if (String(username || '') !== savedUser || String(password || '') !== savedPass) {
    return { success: false, message: 'ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง' };
  }

  // Token มีอายุ 6 ชั่วโมง และตรวจสอบฝั่งเซิร์ฟเวอร์ทุกครั้งที่เรียกข้อมูล
  const token = Utilities.getUuid();
  CacheService.getScriptCache().put('museum_auth_' + token, 'ok', 21600);
  return { success: true, token: token };
}

function requireAuth_(token) {
  if (!token || CacheService.getScriptCache().get('museum_auth_' + token) !== 'ok') {
    throw new Error('เซสชันหมดอายุ กรุณาเข้าสู่ระบบใหม่');
  }
}

function getVisits(token) {
  requireAuth_(token);
  const sheet = setupSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const values = sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues();
  return values.map((r, i) => ({
    id: i + 2,
    timestamp: r[0] instanceof Date ? r[0].toISOString() : String(r[0] || ''),
    groupName: String(r[1] || ''),
    visitorCount: Number(r[2] || 0),
    visitDateTime: r[3] instanceof Date ? r[3].toISOString() : String(r[3] || ''),
    packageType: String(r[4] || '')
  })).sort((a, b) => new Date(b.visitDateTime).getTime() - new Date(a.visitDateTime).getTime());
}

function saveVisit(token, payload) {
  requireAuth_(token);
  if (!payload) throw new Error('ไม่พบข้อมูลที่ต้องการบันทึก');

  const groupName = String(payload.groupName || '').trim();
  const visitorCount = Number(payload.visitorCount);
  const visitDateTime = String(payload.visitDateTime || '').trim();
  const packageType = String(payload.packageType || '').trim();

  if (!groupName) throw new Error('กรุณาระบุชื่อคณะเยี่ยมชม');
  if (!Number.isInteger(visitorCount) || visitorCount < 1) throw new Error('จำนวนผู้เยี่ยมชมต้องเป็นจำนวนเต็มตั้งแต่ 1 คนขึ้นไป');
  if (!visitDateTime || isNaN(new Date(visitDateTime).getTime())) throw new Error('กรุณาระบุวันที่และเวลาให้ถูกต้อง');
  if (!packageType) throw new Error('กรุณาเลือกชุดการต้อนรับ');

  const sheet = setupSheet_();
  const visitDate = new Date(visitDateTime);
  sheet.appendRow([new Date(), groupName, visitorCount, visitDate, packageType]);
  return { success: true, message: 'บันทึกข้อมูลเรียบร้อยแล้ว' };
}

// เรียกฟังก์ชันนี้หนึ่งครั้ง หากต้องการกำหนดชื่อผู้ใช้/รหัสผ่านใหม่ใน Script Properties
function setLoginCredentials() {
  PropertiesService.getScriptProperties().setProperties({
    AUTH_USER: 'library',
    AUTH_PASS: 'library5413'
  }, true);
}
