/** @OnlyCurrentDoc */
// 練舞人測試名單：接收網站送出的 email，寫進這份試算表的「網站名單」工作表。

const SHEET_NAME = "網站名單";
const EMAIL_PATTERN = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const email = String(JSON.parse(e.postData.contents).email || "").trim().toLowerCase();
    if (!EMAIL_PATTERN.test(email) || email.length > 254) {
      return reply({ ok: false, error: "invalid_email" });
    }

    const sheet = getSheet();
    const existing = sheet.getRange(2, 2, Math.max(sheet.getLastRow() - 1, 1), 1).getValues().flat();
    if (!existing.includes(email)) {
      sheet.appendRow([new Date(), email]);
    }
    return reply({ ok: true });
  } catch (error) {
    return reply({ ok: false, error: "server_error" });
  } finally {
    lock.releaseLock();
  }
}

function getSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(SHEET_NAME);
    sheet.appendRow(["時間", "Email"]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function reply(body) {
  return ContentService.createTextOutput(JSON.stringify(body)).setMimeType(ContentService.MimeType.JSON);
}
