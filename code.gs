// ============================================================
//  GOOGLE APPS SCRIPT — REST API Backend (ISUZU TSH)
//  วิธีใช้: Deploy > New deployment > Web App
//           Execute as: Me | Who has access: Anyone
// ============================================================

function doGet(e) {
  // ป้องกันกรณีเผลอกดปุ่ม "เรียกใช้" (Run) จากตัวแก้ไขสคริปต์ตรงๆ
  if (!e || !e.parameter) {
    return ContentService
      .createTextOutput(JSON.stringify({ 
        status: 'OK', 
        message: 'Google Apps Script API พร้อมทำงาน (Web App ทำงานปกติ)' 
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  const action = e.parameter.action;
  let result;

  if (action === 'getConfig') {
    result = getConfig();
  } else if (action === 'getKnownFaces') {
    result = getKnownFaces();
  } else {
    result = { error: 'Unknown action: ' + action };
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  if (!e || !e.postData || !e.postData.contents) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: 'No post data' }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  let data;
  try {
    data = JSON.parse(e.postData.contents);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: 'Invalid JSON body' }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // ใช้ LockService เพื่อป้องกันพนักงานสแกนพร้อมกันแล้วข้อมูลชนกัน
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000); // รอคิวล็อกสูงสุด 10 วินาที
  } catch (lockErr) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: 'Server busy, please try again' }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  let result;
  try {
    const action = data.action;

    if (action === 'registerUser') {
      result = registerUser(data.name, data.faceDescriptor);
    } else if (action === 'logAttendance') {
      result = logAttendance(data.name, data.lat, data.lng);
    } else if (action === 'saveConfig') {
      result = saveConfig(data.lat, data.lng, data.radius);
    } else {
      result = { success: false, error: 'Unknown action: ' + action };
    }
  } catch (processErr) {
    result = { success: false, error: processErr.message };
  } finally {
    lock.releaseLock();
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// --- ส่วนจัดการใบหน้า (Users) ---
function registerUser(name, faceDescriptor) {
  if (!name || !faceDescriptor) {
    return { success: false, error: 'Missing required fields' };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Users');
  if (!sheet) sheet = ss.insertSheet('Users');

  sheet.appendRow([name, JSON.stringify(faceDescriptor), new Date()]);
  return { success: true, message: 'บันทึกข้อมูลหน้าเรียบร้อย' };
}

function getKnownFaces() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Users');
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  let users = [];
  for (let i = 1; i < data.length; i++) {
    const name = data[i][0];
    const jsonStr = data[i][1];
    if (name && jsonStr) {
      try {
        users.push({ label: name, descriptor: JSON.parse(jsonStr) });
      } catch (e) {}
    }
  }
  return users;
}

// --- ส่วนบันทึกเวลา (Attendance) ---
function logAttendance(name, lat, lng) {
  if (!name) {
    return { success: false, error: 'Missing required fields' };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Attendance');
  if (!sheet) {
    sheet = ss.insertSheet('Attendance');
    sheet.appendRow(['Name', 'Time', 'Date', 'Latitude', 'Longitude', 'Google Map Link']);
  }

  const now = new Date();
  const mapLink = (lat && lng) ? `https://www.google.com/maps?q=${lat},${lng}` : '';
  const dateStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'd/M/yyyy');
  const timeStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'HH:mm:ss');

  sheet.appendRow([
    name,
    timeStr,
    "'" + dateStr,
    lat || '-',
    lng || '-',
    mapLink
  ]);
  return { success: true, message: 'บันทึกเวลาสำเร็จ' };
}

// --- ส่วนจัดการ Config (GPS) ---
function saveConfig(lat, lng, radius) {
  const parsedLat = parseFloat(lat);
  const parsedLng = parseFloat(lng);
  const parsedRadius = parseFloat(radius);

  if (isNaN(parsedLat) || isNaN(parsedLng) || isNaN(parsedRadius)) {
    return { success: false, error: 'Invalid coordinate or radius format' };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Config');

  if (!sheet) {
    sheet = ss.insertSheet('Config');
    sheet.getRange('A1:B1').setValues([['Parameter', 'Value']]);
    sheet.getRange('A2').setValue('Target Latitude');
    sheet.getRange('A3').setValue('Target Longitude');
    sheet.getRange('A4').setValue('Allowed Radius (KM)');
    sheet.setColumnWidth(1, 150);
  }

  sheet.getRange('B2').setValue(parsedLat);
  sheet.getRange('B3').setValue(parsedLng);
  sheet.getRange('B4').setValue(parsedRadius);

  return { success: true, message: 'บันทึกการตั้งค่าลง Google Sheets เรียบร้อย' };
}

function getConfig() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Config');

  let config = { lat: 0, lng: 0, radius: 0.02 };

  if (sheet) {
    const latVal = sheet.getRange('B2').getValue();
    const lngVal = sheet.getRange('B3').getValue();
    const radiusVal = sheet.getRange('B4').getValue();

    if (latVal !== '') config.lat = parseFloat(latVal);
    if (lngVal !== '') config.lng = parseFloat(lngVal);
    if (radiusVal !== '') config.radius = parseFloat(radiusVal);
  }

  return config;
}

// ฟังก์ชันสำหรับทดสอบขอสิทธิ์เข้าถึง Google Sheets ครั้งแรก
function testSetup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Logger.log('เชื่อมต่อ Google Sheets สำเร็จ: ' + ss.getName());
}
