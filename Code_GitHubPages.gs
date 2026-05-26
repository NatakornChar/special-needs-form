const SPREADSHEET_NAME = 'แบบสังเกตนักเรียนที่มีความต้องการพิเศษ';

function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) ? e.parameter.action : '';

  try {
    if (action === 'initial') return jsonOutput_(getInitialData());
    if (action === 'teachers') return jsonOutput_(getTeachersByDepartment(e.parameter.department || ''));
    if (action === 'rooms') return jsonOutput_(getStudentFilters(e.parameter.department || '', e.parameter.level || ''));
    if (action === 'students') return jsonOutput_(getStudents(e.parameter.department || '', e.parameter.level || '', e.parameter.room || ''));

    return jsonOutput_({ status: 'ok', message: 'API พร้อมใช้งาน' });
  } catch (err) {
    return jsonOutput_({ status: 'error', message: err.message });
  }
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents || '{}');
    return jsonOutput_(submitObservation(data));
  } catch (err) {
    return jsonOutput_({ status: 'error', message: err.message });
  }
}

function jsonOutput_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function setupSystem() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) ss = SpreadsheetApp.create(SPREADSHEET_NAME);

  PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', ss.getId());

  createSheet_(ss, 'Teachers', ['แผนกวิชา', 'ชื่อ-สกุลครู']);

  createSheet_(ss, 'Students', [
    'แผนกวิชา',
    'ระดับชั้น',
    'ห้องเรียน',
    'รหัสนักเรียน',
    'คำนำหน้า',
    'ชื่อ',
    'นามสกุล',
    'ครูที่ปรึกษา'
  ]);

  createSheet_(ss, 'Responses', [
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
    'ข้อสังเกตเพิ่มเติม'
  ]);

  Logger.log('Sheet URL: ' + ss.getUrl());
}

function createSheet_(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);

  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  }

  sh.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#1565c0')
    .setFontColor('#ffffff')
    .setHorizontalAlignment('center');

  sh.setFrozenRows(1);
  sh.autoResizeColumns(1, headers.length);
}

function getSS_() {
  const id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (id) return SpreadsheetApp.openById(id);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('กรุณารัน setupSystem() ก่อน');

  PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', ss.getId());
  return ss;
}

function getInitialData() {
  const ss = getSS_();
  const teacherSheet = ss.getSheetByName('Teachers');
  const studentSheet = ss.getSheetByName('Students');

  if (!teacherSheet) throw new Error('ไม่พบชีต Teachers');
  if (!studentSheet) throw new Error('ไม่พบชีต Students');

  const teacherData = teacherSheet.getDataRange().getValues().slice(1);
  const studentData = studentSheet.getDataRange().getValues().slice(1);

  return {
    status: 'success',
    teacherDepartments: unique_(teacherData.map(r => clean_(r[0]))),
    studentDepartments: unique_(studentData.map(r => clean_(r[0])))
  };
}

function getTeachersByDepartment(department) {
  const ss = getSS_();
  const sh = ss.getSheetByName('Teachers');
  const data = sh.getDataRange().getValues().slice(1);
  const dep = clean_(department);

  return data
    .filter(r => clean_(r[0]) === dep)
    .map(r => ({
      department: clean_(r[0]),
      name: clean_(r[1])
    }))
    .filter(r => r.name);
}

function getStudentFilters(department, level) {
  const ss = getSS_();
  const sh = ss.getSheetByName('Students');
  const data = sh.getDataRange().getValues().slice(1);
  const dep = clean_(department);
  const lv = clean_(level);

  const rooms = unique_(
    data
      .filter(r => clean_(r[0]) === dep && matchLevel_(clean_(r[1]), lv))
      .map(r => clean_(r[2]))
  );

  return { rooms };
}

function getStudents(department, level, room) {
  const ss = getSS_();
  const sh = ss.getSheetByName('Students');
  const data = sh.getDataRange().getValues().slice(1);
  const dep = clean_(department);
  const lv = clean_(level);
  const rm = clean_(room);

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

function submitObservation(data) {
  const ss = getSS_();
  const sh = ss.getSheetByName('Responses');
  if (!sh) throw new Error('ไม่พบชีต Responses กรุณารัน setupSystem() ก่อน');

  const c = data.categories || {};

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
    (c.vision || []).join(', '),
    (c.hearing || []).join(', '),
    (c.physical || []).join(', '),
    (c.mental || []).join(', '),
    (c.intellectual || []).join(', '),
    (c.ld || []).join(', '),
    (c.autism || []).join(', '),
    data.duration || '',
    data.learningImpact || '',
    data.parentAware || '',
    data.initialAction || '',
    data.urgencyLevel || '',
    data.observationNote || ''
  ]);

  return { status: 'success', message: 'บันทึกข้อมูลเรียบร้อยแล้ว' };
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
