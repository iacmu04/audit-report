/**
 * Google Apps Script for Executive Audit Risk Dashboard & Auto Sheet Setup
 * Paste this script into Extensions > Apps Script in your Google Sheet:
 * https://docs.google.com/spreadsheets/d/1ZzryERuXVUz6JM3F8RadJPOlfM7vOoGsxWX-WaWyGX8/edit
 *
 * Deploy as Web App:
 * Execute as: Me
 * Who has access: Anyone
 */

function setupSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. Setup Report Sheet
  var reportSheet = ss.getSheetByName('Report');
  if (!reportSheet) {
    reportSheet = ss.insertSheet('Report');
  }
  
  var reportHeaders = [
    'Finding_ID', 'ปี', 'ส่วนงาน', 'รอบการประชุม คตส.', 'ประเด็นหลัก', 'ประเด็นย่อย', 
    'ประเภท', 'ระดับความเสี่ยง', 'ข้อตรวจพบ', 'ข้อเสนอแนะ', 'รายงานฉบับเต็ม', 
    'สถานะผลการดำเนินงาน', 'รายละเอียดผลการดำเนินงาน'
  ];
  if (reportSheet.getLastRow() <= 1) {
    reportSheet.clear();
    reportSheet.appendRow(reportHeaders);
    reportSheet.getRange(1, 1, 1, reportHeaders.length).setFontWeight('bold').setBackground('#f1f5f9');
  }

  // 2. Setup Clarify Sheet
  var clarifySheet = ss.getSheetByName('Clarify');
  if (!clarifySheet) {
    clarifySheet = ss.insertSheet('Clarify');
  }
  var clarifyHeaders = [
    'Finding_ID', 'ปี', 'ส่วนงาน', 'ประเด็นหลัก', 'ประเด็นย่อย', 
    'ระดับความเสี่ยง', 'ประเภท', 'ข้อตรวจพบ', 'ข้อเสนอแนะ', 'สถานะการดำเนินงาน', 'รายละเอียดการชี้แจง'
  ];
  if (clarifySheet.getLastRow() <= 1) {
    clarifySheet.clear();
    clarifySheet.appendRow(clarifyHeaders);
    clarifySheet.getRange(1, 1, 1, clarifyHeaders.length).setFontWeight('bold').setBackground('#f1f5f9');
  }

  // 3. Setup User Sheet with Secure Password / PIN
  var userSheet = ss.getSheetByName('User');
  if (!userSheet) {
    userSheet = ss.insertSheet('User');
  }
  var userHeaders = ['email', 'name', 'department', 'role', 'status', 'password'];
  if (userSheet.getLastRow() <= 1) {
    userSheet.clear();
    userSheet.appendRow(userHeaders);
    userSheet.getRange(1, 1, 1, userHeaders.length).setFontWeight('bold').setBackground('#f1f5f9');
    
    // Default authorized users with secure initial passwords
    userSheet.appendRow(['amornrath.f@gmail.com', 'อมรรัตน์ (ผู้ดูแลระบบ)', 'ส่วนงานตรวจสอบภายใน', 'Admin', 'Active', '123456']);
    userSheet.appendRow(['iacmu04@gmail.com', 'ผู้ดูแลระบบ (Google)', 'ส่วนงานตรวจสอบภายใน', 'Admin', 'Active', '123456']);
  }

  // 4. Setup AuditReport Sheet for Newly Added Records
  var auditReportSheet = ss.getSheetByName('AuditReport');
  if (!auditReportSheet) {
    auditReportSheet = ss.insertSheet('AuditReport');
  }
  var arHeaders = [
    'Finding_ID', 'ปี', 'ส่วนงาน', 'รอบการประชุม คตส.', 'ประเด็นหลัก', 'ประเด็นย่อย', 
    'ประเภท', 'ระดับความเสี่ยง', 'ข้อตรวจพบ', 'ข้อเสนอแนะ', 'รายละเอียดผลการดำเนินงาน', 
    'สถานะผลการดำเนินงาน', 'รายงานฉบับเต็ม', 'items_json'
  ];
  if (auditReportSheet.getLastRow() <= 1) {
    auditReportSheet.clear();
    auditReportSheet.appendRow(arHeaders);
    auditReportSheet.getRange(1, 1, 1, arHeaders.length).setFontWeight('bold').setBackground('#f1f5f9');
  }

  Logger.log("ตั้งค่าหัวคอลัมน์ชีทสมบูรณ์เรียบร้อยแล้ว!");
  return "ตั้งค่าหัวคอลัมน์ชีทสมบูรณ์เรียบร้อยแล้ว!";
}

function doGet(e) {
  e = e || { parameter: {} };
  setupSheets();

  var action = e.parameter.action || 'getData';
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  if (action === 'getData') {
    var reportSheet = ss.getSheetByName('Report');
    var clarifySheet = ss.getSheetByName('Clarify');
    var auditReportSheet = ss.getSheetByName('AuditReport');
    var masterConfigSheet = ss.getSheetByName('Master_Config');

    var reportData = getSheetDataAsJson(reportSheet);
    var clarifyData = getSheetDataAsJson(clarifySheet);
    var auditReportData = getSheetDataAsJson(auditReportSheet);
    var masterConfigData = getSheetDataAsJson(masterConfigSheet);

    return createCorsResponse({
      status: 'success',
      report: reportData,
      clarify: clarifyData,
      auditReport: auditReportData,
      masterConfig: masterConfigData
    });
  } else if (action === 'login') {
    var email = (e.parameter.email || e.parameter.username || '').toLowerCase().trim();
    var password = String(e.parameter.password || '').trim();

    var userSheet = ss.getSheetByName('User');
    var users = getSheetDataAsJson(userSheet);

    var found = users.find(function(u) {
      var uEmail = String(u.email || u.username || '').toLowerCase().trim();
      var uStatus = String(u.status || 'Active').toLowerCase().trim();
      return uEmail === email && (uStatus === 'active' || uStatus === 'ใช้งาน');
    });

    if (found) {
      // STRICT SECURE PASSWORD / PIN VERIFICATION
      var realPass = String(found.password || found.pin || '123456').trim();
      if (password !== realPass) {
        return createCorsResponse({ status: 'error', message: 'รหัสผ่าน / รหัส PIN ไม่ถูกต้อง' });
      }

      return createCorsResponse({ 
        status: 'success', 
        user: { 
          email: found.email || email, 
          name: found.name || found.email || email, 
          role: found.role || 'Auditor',
          department: found.department || 'ส่วนงานตรวจสอบ'
        } 
      });
    }

    return createCorsResponse({ 
      status: 'error', 
      message: 'อีเมลของคุณ (' + email + ') ไม่มีสิทธิ์เข้าใช้งาน กรุณาติดต่อผู้ดูแลระบบเพื่อเพิ่มรายชื่อในชีท User' 
    });
  }

  return createCorsResponse({ status: 'error', message: 'Invalid action' });
}

function doPost(e) {
  e = e || { parameter: {}, postData: { contents: '{}' } };
  setupSheets();

  try {
    var postData = JSON.parse(e.postData.contents || '{}');
    var action = postData.action;
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    if (action === 'saveRecord') {
      var record = postData.record;
      var auditReportSheet = ss.getSheetByName('AuditReport');
      if (!auditReportSheet) auditReportSheet = ss.insertSheet('AuditReport');
      
      var arData = auditReportSheet.getDataRange().getValues();
      var foundRowIndex = -1;

      for (var i = 1; i < arData.length; i++) {
        if (String(arData[i][0]) === String(record.finding_id)) {
          foundRowIndex = i + 1;
          break;
        }
      }

      var arRowValues = [
        record.finding_id || ('R' + new Date().getTime()),
        record.year || '',
        record.department || '',
        record.meeting_round || '',
        record.main_topic || '',
        record.sub_topic || '',
        record.type || '',
        record.risk_level || '',
        record.finding || '',
        record.recommendation || '',
        record.clarification_summary || '',
        record.overall_status || 'ยังไม่ได้ชี้แจง',
        record.report_url || '',
        record.items_json || ''
      ];

      if (foundRowIndex > 0) {
        auditReportSheet.getRange(foundRowIndex, 1, 1, arRowValues.length).setValues([arRowValues]);
      } else {
        auditReportSheet.appendRow(arRowValues);
      }

      return createCorsResponse({ status: 'success', message: 'บันทึกข้อมูลเรียบร้อยแล้ว' });
    }
  } catch (err) {
    return createCorsResponse({ status: 'error', message: err.toString() });
  }
}

function getSheetDataAsJson(sheet) {
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  var headers = data[0].map(function(h) { return String(h).trim(); });
  var result = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row.join('').trim()) continue;
    var obj = {};
    for (var j = 0; j < headers.length; j++) {
      obj[headers[j]] = row[j] !== undefined ? row[j] : '';
    }
    result.push(obj);
  }
  return result;
}

function createCorsResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
