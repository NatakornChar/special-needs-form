const ss = SpreadsheetApp.openById('1mr5oYvOON8rS7LqEgbPz6V4gmuigHhPtv5VcDQVY8v4');

function doGet(e) {

  const action = e.parameter.action;

  if (action === 'initial') {

    return output({
      status: 'success',
      teacherDepartments: [
        'การบัญชี',
        'การตลาด',
        'เทคโนโลยีธุรกิจดิจิทัล'
      ],
      studentDepartments: [
        'การบัญชี',
        'การตลาด',
        'เทคโนโลยีธุรกิจดิจิทัล'
      ]
    });
  }

  if (action === 'teachers') {

    return output([
      { name: 'นายสมชาย ใจดี' },
      { name: 'นางสาวสายฝน แสงทอง' }
    ]);
  }

  if (action === 'rooms') {

    return output({
      rooms: ['DBT1.1', 'DBT1.2']
    });
  }

  if (action === 'students') {

    return output([
      {
        studentId: '6921910001',
        fullName: 'นางสาวตัวอย่าง ระบบดี',
        advisorTeacher: 'นายสมชาย ใจดี'
      },
      {
        studentId: '6921910002',
        fullName: 'นายทดสอบ โปรแกรม',
        advisorTeacher: 'นายสมชาย ใจดี'
      }
    ]);
  }

  return output({
    status: 'error',
    message: 'Invalid action'
  });
}

function doPost(e) {

  const data = JSON.parse(e.postData.contents);

  const sheet = ss.getSheetByName('Data');

  sheet.appendRow([
    new Date(),
    data.teacherDepartment,
    data.teacherName,
    data.studentFullName
  ]);

  return output({
    status: 'success',
    message: 'บันทึกข้อมูลเรียบร้อย'
  });
}

function output(data) {

  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
