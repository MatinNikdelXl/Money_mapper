/**
 * === TRX MoneyMapper — Apps Script Backend ===
 * بهترین راه اتصال به Google Sheets بدون OAuth و بدون فیلترینگ.
 *
 * راه‌اندازی (فقط یک بار، ~۲ دقیقه):
 * 1) شیت خودت رو باز کن → از منوی بالا: Extensions (افزونه‌ها) → Apps Script
 * 2) همه کد داخل فایل Code.gs پیش‌فرض رو پاک کن و محتوای همین فایل رو paste کن
 * 3) بالای این فایل، SPREADSHEET_ID رو خالی بذار (خودکار از شیت فعلی گرفته میشه)
 *    و اگر می‌خوای فقط خودت بتونی وصل شی، یه رمز دلخواه بذار تو AUTH_TOKEN
 * 4) Deploy → New deployment → Web app
 *    - Execute as: Me
 *    - Who has access: Anyone
 * 5) آدرس Web app (به /exec ختم میشه) رو کپی کن
 * 6) تو اپ MoneyMapper، پنل تنظیمات → فیلد «آدرس Apps Script» رو paste کن و ذخیره بزن
 *    دیگه نیازی به OAuth Client ID نیست!
 */

// اگر خالی باشه، از شیتِ کلی‌شده استفاده میشه
const SPREADSHEET_ID = '';
// رمز اختیاری — اگر بذاری، هر کسی آدرس وب‌اپ رو داره فقط با این رمز می‌تونه وصل شی
const AUTH_TOKEN = '';

const DATA_SHEET = 'تراکنش‌ها';
const REPORT_SHEET = 'گزارش';
const HEADERS = ['شناسه','تاریخ','نوع','مبلغ','بابت','بانک','کامنت','منبع','آخرین تغییر'];

function ss_() {
  return SPREADSHEET_ID
    ? SpreadsheetApp.openById(SPREADSHEET_ID)
    : SpreadsheetApp.getActiveSpreadsheet();
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents || '{}');
    if (AUTH_TOKEN && body.token !== AUTH_TOKEN)
      return json_({ ok: false, error: 'رمز اشتباه است (AUTH_TOKEN)' });

    let result;
    switch (body.action) {
      case 'list':   result = listRows_(body); break;
      case 'append': result = appendRow_(body); break;
      case 'update': result = updateRow_(body); break;
      case 'delete': result = deleteRow_(body); break;
      default: return json_({ ok: false, error: 'عملیات ناشناخته: ' + body.action });
    }
    return json_(result);
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// اطمینان از وجود استراکچر شیت‌ها + تمپلیت گزارش
function ensureSetup_(sheet) {
  const ss = ss_();
  const dataSheet = sheet || ss.getSheetByName(DATA_SHEET) || ss.insertSheet(DATA_SHEET);
  dataSheet.setRightToLeft(true);

  // هدر
  if (dataSheet.getLastRow() === 0 || dataSheet.getRange(1, 1, 1, 1).getValue() !== HEADERS[0])
    dataSheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);

  const repSheet = ss.getSheetByName(REPORT_SHEET);
  if (!repSheet) buildReport_(ss.insertSheet(REPORT_SHEET));
  return dataSheet;
}

function listRows_(body) {
  const sheet = ensureSetup_(null);
  if (sheet.getLastRow() < 2) return { ok: true, rows: [] };  // شیت خالیه (فقط هدر یا هیچ)
  const last = sheet.getLastRow();
  const values = sheet.getRange(2, 1, last - 1, HEADERS.length).getValues();
  const rows = [];
  values.forEach((row, i) => {
    if (row[0]) rows.push({ rowNum: i + 2, values: row });  // ردیف واقعی در شیت
  });
  return { ok: true, rows, sheets: [DATA_SHEET, REPORT_SHEET] };
}

function appendRow_(body) {
  const sheet = ensureSetup_(null);
  sheet.appendRow(body.row);
  return { ok: true, rowNum: sheet.getLastRow() };
}

function updateRow_(body) {
  const sheet = ensureSetup_(null);
  sheet.getRange(body.rowNum, 1, 1, body.row.length).setValues([body.row]);
  return { ok: true, rowNum: body.rowNum };
}

function deleteRow_(body) {
  const sheet = ensureSetup_(null);
  sheet.deleteRow(body.rowNum);
  return { ok: true };
}

// === تمپلیت «گزارش» (ماهانه + بازه‌ای) ===
function buildReport_(rep) {
  rep.setRightToLeft(true);
  const T = DATA_SHEET;
  const year = new Date().getFullYear();
  const monthRows = [];
  for (let i = 0; i < 12; i++) {
    const r = 15 + i;
    const d = new Date(year, new Date().getMonth() - (11 - i), 1);
    monthRows.push([
      `=DATE(${d.getFullYear()},${d.getMonth() + 1},1)`,
      `=SUMIFS('${T}'!D:D,'${T}'!C:C,"واریز",'${T}'!B:B,">="&TEXT(A${r},"yyyy-mm-dd"),'${T}'!B:B,"<"&TEXT(EOMONTH(A${r},0)+1,"yyyy-mm-dd"))`,
      `=SUMIFS('${T}'!D:D,'${T}'!C:C,"برداشت",'${T}'!B:B,">="&TEXT(A${r},"yyyy-mm-dd"),'${T}'!B:B,"<"&TEXT(EOMONTH(A${r},0)+1,"yyyy-mm-dd"))`,
      `=B${r}-C${r}`,
    ]);
  }
  const values = [
    ['📊 گزارش حساب', '', '', ''],
    ['', '', '', ''],
    ['از تاریخ (yyyy-mm-dd):', '2026-01-01', '', '↙ بین این دو، بازه را ویرایش کن'],
    ['تا تاریخ (yyyy-mm-dd):', '2026-12-31', '', ''],
    ['', '', '', ''],
    ['مجموع واریز در بازه:', `=SUMIFS('${T}'!D:D,'${T}'!C:C,"واریز",'${T}'!B:B,">="&TEXT(B3,"yyyy-mm-dd"),'${T}'!B:B,"<="&TEXT(B4,"yyyy-mm-dd"))`, '', ''],
    ['مجموع برداشت در بازه:', `=SUMIFS('${T}'!D:D,'${T}'!C:C,"برداشت",'${T}'!B:B,">="&TEXT(B3,"yyyy-mm-dd"),'${T}'!B:B,"<="&TEXT(B4,"yyyy-mm-dd"))`, '', ''],
    ['موجودی بازه:', '=B6-B7', '', '(واریز − برداشت)'],
    ['', '', '', ''],
    ['تعداد تراکنش امروز:', `=COUNTIF('${T}'!B:B,TEXT(TODAY(),"yyyy-mm-dd"))`, '', ''],
    ['برداشت امروز:', `=SUMIFS('${T}'!D:D,'${T}'!C:C,"برداشت",'${T}'!B:B,TEXT(TODAY(),"yyyy-mm-dd"))`, '', ''],
    ['', '', '', ''],
    ['خلاصه ماهانه (۱۲ ماه اخیر)', '', '', ''],
    ['ماه', 'واریز', 'برداشت', 'خالص'],
    ...monthRows,
  ];
  rep.getRange(1, 1, values.length, 4).setValues(values);
}
