// ═══════════════════════════════════════════════════
// Plano — Google Apps Script
// این کد رو در Google Sheet خودت paste کن:
// Extensions > Apps Script > پاک کن همه چیز > paste کن > Deploy
// ═══════════════════════════════════════════════════

function doGet(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data  = sheet.getDataRange().getValues();
  const headers = data[0];

  // ایندکس ستون‌ها (0-based)
  // A=0 Submission ID, B=1 Respondent ID, C=2 Submitted at
  // D=3 نام, E=4 نام خانوادگی, F=5 شماره تماس, G=6 ایمیل
  // H=7 نام پروژه, I=8 شهر
  // J=9 درخواست جذب (کلی)
  // K=10 بازیگر, L=11 کارگردان, M=12 دستیار کارگردان
  // N=13 فیلمبردار, O=14 صدابردار, P=15 گریمور
  // Q=16 طراح صحنه, R=17 طراح لباس, S=18 تدوینگر, T=19 سایر
  // U=20 توضیحات پروژه, V=21 دستمزد, W=22 نیاز به سابقه
  // X=23 مایل به عضویت

  const rows = [];

  for (var i = 1; i < data.length; i++) {
    var row = data[i];

    // رد کن اگه نام پروژه خالیه
    if (!row[7]) continue;

    // نقش‌های درخواستی رو جمع کن
    var roles = [];
    var roleMap = {
      10: 'بازیگر',
      11: 'کارگردان',
      12: 'دستیار کارگردان',
      13: 'فیلمبردار',
      14: 'صدابردار',
      15: 'گریمور',
      16: 'طراح صحنه',
      17: 'طراح لباس',
      18: 'تدوینگر',
      19: 'سایر'
    };
    for (var col in roleMap) {
      if (row[col] && row[col].toString().trim() !== '') {
        roles.push(roleMap[col]);
      }
    }

    rows.push({
      id:          row[0],
      project:     row[7]  || '',
      city:        row[8]  || '',
      roles:       roles,
      description: row[20] || '',
      salary:      row[21] || '',
      experience:  row[22] || '',
      date:        row[2]  ? new Date(row[2]).toLocaleDateString('fa-IR') : ''
    });
  }

  // جدیدترین اول
  rows.reverse();

  var output = ContentService
    .createTextOutput(JSON.stringify({ status: 'ok', data: rows }))
    .setMimeType(ContentService.MimeType.JSON);

  // CORS — اجازه دسترسی از getplano.ir
  return output;
}
