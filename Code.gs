/**
 * AFAPS Museum API for GitHub Pages
 * Google Apps Script Web App backend.
 *
 * IMPORTANT:
 * - Deploy as Web app, Execute as: Me.
 * - The frontend uses JSONP for read requests and no-cors POST for writes.
 * - Do not publish this backend code with real credentials in a public repository.
 */

const SPREADSHEET_ID = '1ShJrfy912ab274U4EiDtXgIqXFkuZ4HbLui3181aguE';
const SHEET_NAME = 'MuseumVisits';
const DEFAULT_USER = 'library';
const DEFAULT_PASS = 'library5413';
const TOKEN_TTL_SECONDS = 21600;

const HEADERS = ['Timestamp', 'GroupName', 'VisitorCount', 'VisitDateTime', 'PackageType'];

function doGet(e) {
  const p = (e && e.parameter) || {};
  const action = String(p.action || '');
  let result;

  try {
    if (action === 'login') {
      // Login should be POSTed. Reject credentials sent in a GET URL.
      result = { success: false, message: 'ส่งข้อมูลเข้าสู่ระบบด้วย POST เท่านั้น' };
    } else if (action === 'session') {
      result = { success: isValidToken_(p.token), message: isValidToken_(p.token) ? 'OK' : 'เซสชันหมดอายุ' };
    } else if (action === 'list') {
      requireToken_(p.token);
      result = { success: true, data: listVisits_() };
    } else {
      result = { success: false, message: 'ไม่รู้จักคำสั่ง API' };
    }
  } catch (err) {
    result = { success: false, message: err.message || 'เกิดข้อผิดพลาด' };
  }

  return jsonp_(result, p.callback);
}

function doPost(e) {
  let payload = {};
  try {
    const raw = e && e.postData && e.postData.contents ? e.postData.contents : '';
    payload = raw ? JSON.parse(raw) : {};
  } catch (err) {
    return json_({ success: false, message: 'รูปแบบข้อมูลไม่ถูกต้อง' });
  }

  try {
    const action = String(payload.action || '');
    let result;

    if (action === 'login') {
      const props = PropertiesService.getScriptProperties();
      const user = props.getProperty('AUTH_USER') || DEFAULT_USER;
      const pass = props.getProperty('AUTH_PASS') || DEFAULT_PASS;
      if (String(payload.username || '') !== user || String(payload.password || '') !== pass) {
        result = { success: false, message: 'ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง' };
      } else {
        const token = String(payload.clientToken || '');
        if (!/^[a-zA-Z0-9_-]{16,100}$/.test(token)) throw new Error('โทเคนเข้าสู่ระบบไม่ถูกต้อง');
        CacheService.getScriptCache().put('museum_auth_' + token, 'ok', TOKEN_TTL_SECONDS);
        result = { success: true, token: token };
      }
    } else if (action === 'save') {
      requireToken_(payload.token);
      const groupName = String(payload.groupName || '').trim();
      const visitorCount = Number(payload.visitorCount);
      const visitDateTime = String(payload.visitDateTime || '').trim();
      const packageType = String(payload.packageType || '').trim();

      if (!groupName) throw new Error('กรุณาระบุชื่อคณะเยี่ยมชม');
      if (!Number.isInteger(visitorCount) || visitorCount < 1) throw new Error('จำนวนผู้เยี่ยมชมต้องเป็นจำนวนเต็มตั้งแต่ 1 คนขึ้นไป');
      if (!visitDateTime || isNaN(new Date(visitDateTime).getTime())) throw new Error('กรุณาระบุวันที่และเวลาให้ถูกต้อง');
      if (!packageType) throw new Error('กรุณาเลือกชุดการต้อนรับ');

      const sheet = getSheet_();
      sheet.appendRow([new Date(), groupName, visitorCount, new Date(visitDateTime), packageType]);
      result = { success: true, message: 'บันทึกข้อมูลเรียบร้อยแล้ว' };
    } else if (action === 'logout') {
      if (payload.token) CacheService.getScriptCache().remove('museum_auth_' + payload.token);
      result = { success: true };
    } else {
      result = { success: false, message: 'ไม่รู้จักคำสั่ง API' };
    }

    // In no-cors mode the browser cannot read this response; frontend checks state by JSONP.
    return json_(result);
  } catch (err) {
    return json_({ success: false, message: err.message || 'เกิดข้อผิดพลาด' });
  }
}

function isValidToken_(token) {
  return !!token && CacheService.getScriptCache().get('museum_auth_' + token) === 'ok';
}
function requireToken_(token) {
  if (!isValidToken_(token)) throw new Error('เซสชันหมดอายุ กรุณาเข้าสู่ระบบใหม่');
}
function getSheet_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold').setBackground('#6b21a8').setFontColor('#ffffff');
  } else {
    const first = sheet.getRange(1, 1, 1, HEADERS.length).getDisplayValues()[0];
    if (first[0] !== HEADERS[0]) throw new Error('หัวตารางชีตไม่ตรงกับระบบ กรุณาตรวจสอบแท็บ MuseumVisits');
  }
  return sheet;
}
function listVisits_() {
  const sheet = getSheet_();
  if (sheet.getLastRow() < 2) return [];
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, HEADERS.length).getValues().map((r, i) => ({
    id: i + 2,
    timestamp: r[0] instanceof Date ? r[0].toISOString() : String(r[0] || ''),
    groupName: String(r[1] || ''),
    visitorCount: Number(r[2] || 0),
    visitDateTime: r[3] instanceof Date ? r[3].toISOString() : String(r[3] || ''),
    packageType: String(r[4] || '')
  })).sort((a, b) => new Date(b.visitDateTime).getTime() - new Date(a.visitDateTime).getTime());
}
function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
function jsonp_(obj, callback) {
  const cb = String(callback || '');
  if (!/^[a-zA-Z_$][0-9a-zA-Z_$\.]*$/.test(cb)) return json_(obj);
  return ContentService.createTextOutput(cb + '(' + JSON.stringify(obj).replace(/</g, '\\u003c') + ');')
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}

// Run once from the Apps Script editor to set/change credentials.
function setLoginCredentials() {
  PropertiesService.getScriptProperties().setProperties({
    AUTH_USER: 'library',
    AUTH_PASS: 'library5413'
  }, true);
}
