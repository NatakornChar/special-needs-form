const SPREADSHEET_ID = '1mr5oYvOON8rS7LqEgbPz6V4gmuigHhPtv5VcDQVY8v4';

const DISABILITY_FOLDER_ID = 'ใส่ Folder ID บัตรพิการ';
const MEDICAL_FOLDER_ID = 'ใส่ Folder ID ใบรับรองแพทย์';

function setupSystem() {
  const ss = getSS_();

  createSheetIfNotExists_(ss, 'Teachers', [
    'แผนกวิชา',
    'ชื่อ-สกุลครู'
  ]);

  createSheetIfNotExists_(ss, 'Students', [
    'แผนกวิชา',
    'ระดับชั้น',
    'ห้องเรียน',
    'รหัสนักเรียน',
    'คำนำหน้า',
    'ชื่อ',
    'นามสกุล',
    'ครูที่ปรึกษา'
  ]);

  createSheetIfNotExists_(ss, 'Responses', getResponseHeaders_());

  SpreadsheetApp.flush();
  Logger.log('setupSystem สำเร็จ');
}

function getResponseHeaders_() {
  return [
    'Timestamp',
    'วันที่กรอก',
    'แผนกครู',
    'ชื่อครูผู้กรอก',
    'บทบาท',
    'วิชาที่สอน/รายวิชาที่พบพฤติกรรม',
    'แผนกนักเรียน',
    'ระดับชั้น',
    'ห้องเรียน',
    'รหัสนักเรียน',
    'ชื่อ-นามสกุลนักเรียน',
    'ครูที่ปรึกษา',
    'สถานะบัตรคนพิการ',
    'เลขบัตรคนพิการ',
    'ประเภทความพิการในบัตร',
    'หน่วยงานที่ออกบัตร',
    'วันหมดอายุบัตร',
    'ใบรับรองแพทย์/เอกสารวินิจฉัย',
    'ความบกพร่องทางการมองเห็น',
    'ความบกพร่องทางการได้ยิน',
    'ความบกพร่องทางร่างกาย/การเคลื่อนไหว',
    'ความบกพร่องทางจิตใจ/พฤติกรรม',
    'ความบกพร่องทางสติปัญญา',
    'ความบกพร่องทางการเรียนรู้ LD',
    'ความพิการทางออทิสติก ASD',
    'ระยะเวลาที่พบ',
    'ผลกระทบต่อการเรียน',
    'ผู้ปกครองรับทราบ',
    'การพูดคุย/ประสานงานเบื้องต้น',
    'ระดับความเร่งด่วน',
    'ข้อสังเกตเพิ่มเติม',
    'ลิงก์บัตรพิการ',
    'ลิงก์ใบรับรองแพทย์'
  ];
}

function doGet(e) {
  const params = e && e.parameter ? e.parameter : {};
  const action = clean_(params.action);
  const callback = clean_(params.callback);

  try {
    let result;

    if (action === 'initial') {
      result = getInitialData_();
    } else if (action === 'teachers') {
      result = getTeachersByDepartment_(params.department || '');
    } else if (action === 'rooms') {
      result = getStudentFilters_(params.department || '', params.level || '');
    } else if (action === 'students') {
      result = getStudents_(params.department || '', params.level || '', params.room || '');
    } else if (action === 'save') {
      const data = JSON.parse(params.data || '{}');
      result = saveObservation_(data);
    } else {
      result = { status: 'success', message: 'API พร้อมใช้งาน' };
    }

    return output_(result, callback);

  } catch (err) {
    return output_({
      status: 'error',
      message: err && err.message ? err.message : String(err)
    }, callback);
  }
}

function getInitialData_() {
  const ss = getSS_();
  const teacherSheet = ss.getSheetByName('Teachers');
  const studentSheet = ss.getSheetByName('Students');

  if (!teacherSheet) throw new Error('ไม่พบชีต Teachers');
  if (!studentSheet) throw new Error('ไม่พบชีต Students');

  const teacherData = teacherSheet.getDataRange().getValues().slice(1);
  const studentData = studentSheet.getDataRange().getValues().slice(1);

  return {
    status: 'success',
    teacherDepartments: unique_(teacherData.map(r => r[0])),
    studentDepartments: unique_(studentData.map(r => r[0]))
  };
}

function getTeachersByDepartment_(department) {
  const ss = getSS_();
  const sh = ss.getSheetByName('Teachers');
  const dep = clean_(department);
  const data = sh.getDataRange().getValues().slice(1);

  return data
    .filter(r => clean_(r[0]) === dep)
    .map(r => ({
      department: clean_(r[0]),
      name: clean_(r[1])
    }))
    .filter(x => x.name);
}

function getStudentFilters_(department, level) {
  const ss = getSS_();
  const sh = ss.getSheetByName('Students');
  const dep = clean_(department);
  const lv = clean_(level);
  const data = sh.getDataRange().getValues().slice(1);

  const rooms = unique_(
    data
      .filter(r => clean_(r[0]) === dep && matchLevel_(clean_(r[1]), lv))
      .map(r => r[2])
  );

  return { rooms };
}

function getStudents_(department, level, room) {
  const ss = getSS_();
  const sh = ss.getSheetByName('Students');
  const dep = clean_(department);
  const lv = clean_(level);
  const rm = clean_(room);
  const data = sh.getDataRange().getValues().slice(1);

  return data
    .filter(r =>
      clean_(r[0]) === dep &&
      matchLevel_(clean_(r[1]), lv) &&
      clean_(r[2]) === rm
    )
    .map(r => {
      const prefix = clean_(r[4]);
      const firstName = clean_(r[5]);
      const lastName = clean_(r[6]);

      return {
        department: clean_(r[0]),
        level: clean_(r[1]),
        room: clean_(r[2]),
        studentId: clean_(r[3]),
        prefix,
        firstName,
        lastName,
        fullName: `${prefix} ${firstName}  ${lastName}`.trim(),
        advisorTeacher: clean_(r[7])
      };
    });
}

function saveObservation_(data) {
  const ss = getSS_();
  const sh = ss.getSheetByName('Responses');
  if (!sh) throw new Error('ไม่พบชีต Responses');

  ensureHeaders_(sh, getResponseHeaders_());

  const c = data.categories || {};

  const disabilityLink = uploadPdf_(
    data.disabilityCardFile,
    data.disabilityCardFileName,
    DISABILITY_FOLDER_ID
  );

  const medicalLink = uploadPdf_(
    data.medicalFile,
    data.medicalFileName,
    MEDICAL_FOLDER_ID
  );

  sh.appendRow([
    new Date(),
    data.observeDate || '',
    data.teacherDepartment || '',
    data.teacherName || '',
    data.teacherRole || '',
    data.subjectName || '',
    data.studentDepartment || '',
    data.studentLevel || '',
    data.studentRoom || '',
    data.studentId || '',
    data.studentFullName || '',
    data.advisorTeacher || '',
    data.disabilityCardStatus || '',
    data.disabilityCardNo || '',
    data.cardDisabilityType || '',
    data.cardIssuer || '',
    data.cardExpireDate || '',
    data.medicalDocumentStatus || '',
    arrayText_(c.vision),
    arrayText_(c.hearing),
    arrayText_(c.physical),
    arrayText_(c.mental),
    arrayText_(c.intellectual),
    arrayText_(c.ld),
    arrayText_(c.autism),
    data.duration || '',
    data.learningImpact || '',
    data.parentAware || '',
    data.initialAction || '',
    data.urgencyLevel || '',
    data.observationNote || '',
    disabilityLink || '',
    medicalLink || ''
  ]);

  return {
    status: 'success',
    message: 'บันทึกข้อมูลเรียบร้อยแล้ว'
  };
}

function uploadPdf_(base64, fileName, folderId) {
  if (!base64) return '';

  if (!folderId || folderId.includes('ใส่ Folder ID')) {
    throw new Error('ยังไม่ได้ใส่ Folder ID สำหรับเก็บไฟล์ PDF');
  }

  const safeName = fileName || ('document_' + new Date().getTime() + '.pdf');

  const blob = Utilities.newBlob(
    Utilities.base64Decode(base64),
    MimeType.PDF,
    safeName
  );

  const file = DriveApp.getFolderById(folderId).createFile(blob);

  file.setSharing(
    DriveApp.Access.ANYONE_WITH_LINK,
    DriveApp.Permission.VIEW
  );

  return file.getUrl();
}

function getSS_() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

function createSheetIfNotExists_(ss, sheetName, headers) {
  let sh = ss.getSheetByName(sheetName);
  if (!sh) {
    sh = ss.insertSheet(sheetName);
  }

  ensureHeaders_(sh, headers);
}

function ensureHeaders_(sh, headers) {
  const lastCol = Math.max(sh.getLastColumn(), 1);
  const oldHeaders = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(clean_);

  if (sh.getLastRow() === 0 || oldHeaders.filter(Boolean).length === 0) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  } else {
    headers.forEach(header => {
      if (!oldHeaders.includes(header)) {
        const newCol = sh.getLastColumn() + 1;
        sh.getRange(1, newCol).setValue(header);
      }
    });
  }

  const newLastCol = sh.getLastColumn();
  sh.getRange(1, 1, 1, newLastCol)
    .setFontWeight('bold')
    .setBackground('#1565c0')
    .setFontColor('#ffffff')
    .setHorizontalAlignment('center');

  sh.setFrozenRows(1);
}

function output_(data, callback) {
  const json = JSON.stringify(data);

  if (callback) {
    return ContentService
      .createTextOutput(callback + '(' + json + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  return ContentService
    .createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

function clean_(value) {
  return String(value || '').trim();
}

function unique_(arr) {
  return [...new Set(arr.map(x => clean_(x)).filter(Boolean))];
}

function matchLevel_(sheetLevel, selectedLevel) {
  if (!selectedLevel) return false;
  return sheetLevel === selectedLevel || sheetLevel.startsWith(selectedLevel);
}

function arrayText_(arr) {
  return Array.isArray(arr) ? arr.join(', ') : '';
}
