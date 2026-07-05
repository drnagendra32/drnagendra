function doPost(e) {
  try {
    const payload = parsePayload(e);
    const appointment = createAppointment(payload);
    const token = typeof appointment === 'string' ? appointment : appointment.token;
    const notification = typeof appointment === 'string' ? null : appointment.notification;

    return ContentService
      .createTextOutput(JSON.stringify({
        status: true,
        success: true,
        data: {
          token: token,
          notification: notification
        }
      }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({
        status: false,
        success: false,
        message: err.message,
        error: err.message
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet() {
  try {
    const spreadsheet = SpreadsheetApp.openById(SHEET_ID);
    let sheet = spreadsheet.getSheetByName(SHEET_NAME);
    if (!sheet) {
      return ContentService
        .createTextOutput(JSON.stringify({ status: false, error: 'Sheet not found: ' + SHEET_NAME }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    const headers = sheet.getRange(1, 1, 1, Math.max(1, sheet.getLastColumn())).getValues()[0];
    return ContentService
      .createTextOutput(JSON.stringify({
        status: true,
        sheetName: SHEET_NAME,
        headers: headers
      }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
