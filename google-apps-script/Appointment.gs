function createAppointment(payload) {
  const spreadsheet = SpreadsheetApp.openById(SHEET_ID);
  let sheet = spreadsheet.getSheetByName(SHEET_NAME);

  // If the target sheet doesn't exist, create it with standard headers
  if (!sheet) {
    sheet = spreadsheet.insertSheet(SHEET_NAME);
    sheet.appendRow(['Timestamp', 'Name', 'Phone', 'Age', 'Gender', 'Department', 'Problem', 'Date', 'Time']);
  }

  // Validate the incoming parameters
  validateAppointmentPayload(payload);

  // Read the current headers from the sheet to match dynamically
  const lastCol = Math.max(1, sheet.getLastColumn());
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  
  const headerMap = {};
  headers.forEach((header, index) => {
    if (header) {
      headerMap[header.toString().trim().toLowerCase()] = index;
    }
  });

  // Create a row array of empty strings matching the sheet width
  const rowData = new Array(headers.length).fill('');

  // Helper to set values based on list of possible header names
  function setVal(names, val) {
    for (let name of names) {
      const cleanName = name.toLowerCase();
      if (cleanName in headerMap) {
        rowData[headerMap[cleanName]] = val;
        break;
      }
    }
  }

  // Pre-generate token to save it into the spreadsheet
  const token = generateToken(sheet, 1);

  // Set parameters in their correct columns based on headers
  setVal(['token', 'booking token', 'token id'], token);
  setVal(['timestamp', 'time stamp', 'date & time'], new Date());
  setVal(['name', 'patient name', 'full name', 'patient_name'], payload.name || '');
  setVal(['phone', 'mobile', 'mobile number', 'phone number', 'contact', 'mobile_number'], payload.phone || payload.mobile || '');
  setVal(['email', 'patient email', 'email address'], payload.email || '');
  setVal(['age'], payload.age || '');
  setVal(['gender', 'sex'], payload.gender || '');
  setVal(['department', 'treatment', 'service'], payload.department || '');
  setVal(['concern', 'problem', 'complaint', 'chief complaint', 'chief_complaint'], payload.problem || '');
  setVal(['date', 'preferred date', 'appointment date'], payload.date || '');
  setVal(['time', 'preferred time', 'appointment time'], payload.time || '');
  
  // Set default values for administrative columns in sheet
  setVal(['doctor', 'assigned doctor'], 'Dr. Nagendra Chauhan');
  setVal(['status', 'booking status'], 'Pending');
  setVal(['source', 'booking source'], 'Website');
  setVal(['created by', 'created_by'], 'Web Intake');
  setVal(['last updated', 'last_updated'], new Date());

  // Append new row and return the generated token
  sheet.appendRow(rowData);

  const notification = notifyAppointment(payload, token);
  return {
    token: token,
    notification: notification
  };
}
