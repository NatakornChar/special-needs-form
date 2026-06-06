const SPREADSHEET_ID = '1mr5oYvOON8rS7LqEgbPz6V4gmuigHhPtv5VcDQVY8v4';
const ROOT_FOLDER_ID = ''; // ใส่ Folder ID หลักได้ ถ้าเว้นว่างระบบจะสร้างโฟลเดอร์เอง

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('แบบสังเกตนักเรียนที่มีความต้องการพิเศษ')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function setupSystem() {
  const ss = getSS_();
  createSheetIfNotExists_(ss, 'Teachers', ['แผนกวิชา', 'ชื่อ-สกุลครู']);
  createSheetIfNotExists_(ss, 'Students', ['แผนกวิชา','ระดับชั้น','ห้องเรียน','รหัสนักเรียน','คำนำหน้า','ชื่อ','นามสกุล','ครูที่ปรึกษา']);
  createSheetIfNotExists_(ss, 'Responses', getResponseHeaders_());
  const f = getOrCreateUploadFolders_();
  SpreadsheetApp.flush();
  Logger.log('setupSystem สำเร็จ: ' + ss.getUrl());
  Logger.log('โฟลเดอร์หลัก: ' + f.root.getUrl());
}

function getResponseHeaders_() {
  return [
    'Timestamp','วันที่กรอก','แผนกครู','ชื่อครูผู้กรอก','บทบาท','วิชาที่สอน/รายวิชาที่พบพฤติกรรม',
    'แผนกนักเรียน','ระดับชั้น','ห้องเรียน','รหัสนักเรียน','ชื่อ-นามสกุลนักเรียน','ครูที่ปรึกษา',
    'สถานะบัตรคนพิการ','เลขบัตรคนพิการ','ประเภทความพิการในบัตร','หน่วยงานที่ออกบัตร','วันหมดอายุบัตร','ใบรับรองแพทย์/เอกสารวินิจฉัย',
    'ความบกพร่องทางการมองเห็น','ความบกพร่องทางการได้ยิน','ความบกพร่องทางร่างกาย/การเคลื่อนไหว','ความบกพร่องทางจิตใจ/พฤติกรรม','ความบกพร่องทางสติปัญญา','ความบกพร่องทางการเรียนรู้ LD','ความพิการทางออทิสติก ASD',
    'ระยะเวลาที่พบ','ผลกระทบต่อการเรียน','ผู้ปกครองรับทราบ','การพูดคุย/ประสานงานเบื้องต้น','ระดับความเร่งด่วน','ข้อสังเกตเพิ่มเติม',
    'ลิงก์บัตรพิการ','ลิงก์ใบรับรองแพทย์'
  ];
}

function getInitialData() {
  const ss = getSS_();
  const teacherSheet = ss.getSheetByName('Teachers');
  const studentSheet = ss.getSheetByName('Students');
  if (!teacherSheet) throw new Error('ไม่พบชีต Teachers กรุณารัน setupSystem()');
  if (!studentSheet) throw new Error('ไม่พบชีต Students กรุณารัน setupSystem()');
  const teacherData = teacherSheet.getDataRange().getValues().slice(1);
  const studentData = studentSheet.getDataRange().getValues().slice(1);
  return { status:'success', teacherDepartments: unique_(teacherData.map(r=>r[0])), studentDepartments: unique_(studentData.map(r=>r[0])) };
}

function getTeachersByDepartment(department) {
  const sh = getSS_().getSheetByName('Teachers');
  if (!sh) throw new Error('ไม่พบชีต Teachers');
  const dep = clean_(department);
  return sh.getDataRange().getValues().slice(1)
    .filter(r => clean_(r[0]) === dep)
    .map(r => ({ department: clean_(r[0]), name: clean_(r[1]) }))
    .filter(x => x.name);
}

function getStudentFilters(department, level) {
  const sh = getSS_().getSheetByName('Students');
  if (!sh) throw new Error('ไม่พบชีต Students');
  const dep = clean_(department), lv = clean_(level);
  const rooms = unique_(sh.getDataRange().getValues().slice(1)
    .filter(r => clean_(r[0]) === dep && matchLevel_(clean_(r[1]), lv))
    .map(r => r[2]));
  return { rooms };
}

function getStudents(department, level, room) {
  const sh = getSS_().getSheetByName('Students');
  if (!sh) throw new Error('ไม่พบชีต Students');
  const dep = clean_(department), lv = clean_(level), rm = clean_(room);
  return sh.getDataRange().getValues().slice(1)
    .filter(r => clean_(r[0]) === dep && matchLevel_(clean_(r[1]), lv) && clean_(r[2]) === rm)
    .map(r => {
      const prefix = clean_(r[4]), firstName = clean_(r[5]), lastName = clean_(r[6]);
      return {
        department: clean_(r[0]), level: clean_(r[1]), room: clean_(r[2]), studentId: clean_(r[3]),
        prefix, firstName, lastName, fullName: `${prefix} ${firstName}  ${lastName}`.trim(), advisorTeacher: clean_(r[7])
      };
    });
}

function saveObservation(data) {
  try {
    const ss = getSS_();
    const sh = ss.getSheetByName('Responses');
    if (!sh) throw new Error('ไม่พบชีต Responses กรุณารัน setupSystem()');
    ensureHeaders_(sh, getResponseHeaders_());
    const c = data.categories || {};
    const folders = getOrCreateUploadFolders_();
    Logger.log('disability file: ' + (data.disabilityCardFileName || '') + ' / ' + (data.disabilityCardFileBase64 ? data.disabilityCardFileBase64.length : 0));
    Logger.log('medical file: ' + (data.medicalFileName || '') + ' / ' + (data.medicalFileBase64 ? data.medicalFileBase64.length : 0));

    const disabilityLink = uploadPdfFromBase64_(
      data.disabilityCardFileBase64,
      makeFileName_('บัตรพิการ', data.studentId, data.studentFullName),
      folders.disabilityFolder
    );
    const medicalLink = uploadPdfFromBase64_(
      data.medicalFileBase64,
      makeFileName_('ใบรับรองแพทย์', data.studentId, data.studentFullName),
      folders.medicalFolder
    );

    sh.appendRow([
      new Date(), data.observeDate || '', data.teacherDepartment || '', data.teacherName || '', data.teacherRole || '', data.subjectName || '',
      data.studentDepartment || '', data.studentLevel || '', data.studentRoom || '', data.studentId || '', data.studentFullName || '', data.advisorTeacher || '',
      data.disabilityCardStatus || '', data.disabilityCardNo || '', data.cardDisabilityType || '', data.cardIssuer || '', data.cardExpireDate || '', data.medicalDocumentStatus || '',
      arrayText_(c.vision), arrayText_(c.hearing), arrayText_(c.physical), arrayText_(c.mental), arrayText_(c.intellectual), arrayText_(c.ld), arrayText_(c.autism),
      data.duration || '', data.learningImpact || '', data.parentAware || '', data.initialAction || '', data.urgencyLevel || '', data.observationNote || '',
      disabilityLink || '', medicalLink || ''
    ]);
    return { status:'success', message:'บันทึกข้อมูลเรียบร้อยแล้ว', disabilityLink: disabilityLink || '', medicalLink: medicalLink || '' };
  } catch (err) {
    Logger.log('saveObservation ERROR: ' + (err && err.stack ? err.stack : err));
    throw new Error(err && err.message ? err.message : String(err));
  }
}

function uploadPdfFromBase64_(base64, fileName, folder) {
  if (!base64) return '';
  if (!folder) throw new Error('ไม่พบโฟลเดอร์สำหรับเก็บไฟล์ PDF');
  const blob = Utilities.newBlob(Utilities.base64Decode(base64), MimeType.PDF, sanitizeFileName_(fileName || ('document_' + Date.now() + '.pdf')));
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return file.getUrl();
}

function getOrCreateUploadFolders_() {
  let root;
  if (ROOT_FOLDER_ID && ROOT_FOLDER_ID.trim()) root = DriveApp.getFolderById(ROOT_FOLDER_ID);
  else root = getOrCreateFolderByName_('เอกสารแบบสังเกตนักเรียนพิเศษ');
  return {
    root,
    disabilityFolder: getOrCreateSubFolder_(root, 'บัตรประจำตัวคนพิการ'),
    medicalFolder: getOrCreateSubFolder_(root, 'ใบรับรองแพทย์และเอกสารวินิจฉัย')
  };
}
function getOrCreateFolderByName_(name) { const folders = DriveApp.getFoldersByName(name); return folders.hasNext() ? folders.next() : DriveApp.createFolder(name); }
function getOrCreateSubFolder_(parent, name) { const folders = parent.getFoldersByName(name); return folders.hasNext() ? folders.next() : parent.createFolder(name); }
function getSS_() { return SpreadsheetApp.openById(SPREADSHEET_ID); }
function createSheetIfNotExists_(ss, sheetName, headers) { let sh = ss.getSheetByName(sheetName); if (!sh) sh = ss.insertSheet(sheetName); ensureHeaders_(sh, headers); }
function ensureHeaders_(sh, headers) {
  const lastCol = Math.max(sh.getLastColumn(), 1);
  const oldHeaders = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(clean_);
  if (sh.getLastRow() === 0 || oldHeaders.filter(Boolean).length === 0) sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  else headers.forEach(header => { if (!oldHeaders.includes(header)) sh.getRange(1, sh.getLastColumn() + 1).setValue(header); });
  const newLastCol = sh.getLastColumn();
  sh.getRange(1, 1, 1, newLastCol).setFontWeight('bold').setBackground('#1565c0').setFontColor('#ffffff').setHorizontalAlignment('center');
  sh.setFrozenRows(1);
}
function clean_(value) { return String(value || '').trim(); }
function unique_(arr) { return [...new Set(arr.map(x => clean_(x)).filter(Boolean))]; }
function matchLevel_(sheetLevel, selectedLevel) { if (!selectedLevel) return false; return sheetLevel === selectedLevel || sheetLevel.startsWith(selectedLevel); }
function arrayText_(arr) { return Array.isArray(arr) ? arr.join(', ') : ''; }
function sanitizeFileName_(name) { return String(name || 'document.pdf').replace(/[\\/:*?"<>|]/g, '_').replace(/\s+/g, ' ').trim(); }
function makeFileName_(prefix, studentId, studentName) {
  const stamp = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyyMMdd_HHmmss');
  return sanitizeFileName_([prefix || 'เอกสาร', studentId || '', studentName || '', stamp].filter(Boolean).join('_') + '.pdf');
}
