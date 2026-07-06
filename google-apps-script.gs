function doPost(e) {
  const sheetId = '1cY9edL6177CGhRcXeoFq_0xlbFohqSWHpR-vFP40oQU';
  const sheetName = 'Appointments';
  const clinicMobileNumber = '9415964504';
  const clinicWhatsAppNumber = '919415964504';
  const clinicEmail = 'nagendra366@gmail.com';
  const whatsappGraphApiVersion = 'v20.0';
  const whatsappPhoneNumberId = '';
  const whatsappAccessToken = '';
  const whatsappTemplateLanguage = 'en_US';
  const whatsappPatientTemplateName = 'appointment_confirmation';
  const whatsappDoctorTemplateName = 'new_appointment_alert';
  const fast2SmsApiKey = '';

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000); // Wait up to 15 seconds for other concurrent requests to finish
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({
        status: false,
        success: false,
        message: 'System is busy processing another booking. Please try again in a moment.',
        error: 'System busy'
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  try {
    const payload = parsePayload(e);
    const spreadsheet = SpreadsheetApp.openById(sheetId);
    let sheet = spreadsheet.getSheetByName(sheetName);

    if (!sheet) {
      sheet = spreadsheet.insertSheet(sheetName);
      sheet.appendRow(['Timestamp', 'Name', 'Phone', 'Age', 'Gender', 'Department', 'Problem', 'Date', 'Time']);
    }

    // Read the current headers from the sheet to match dynamically
    const lastCol = Math.max(1, sheet.getLastColumn());
    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    
    const headerMap = {};
    headers.forEach((header, index) => {
      if (header) {
        headerMap[header.toString().trim().toLowerCase()] = index;
      }
    });

    // --- Duplicate Check ---
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      const dataRange = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
      const displayDataRange = sheet.getRange(2, 1, lastRow - 1, lastCol).getDisplayValues();
      const phoneColIdx = headerMap['phone'] ?? headerMap['mobile'] ?? headerMap['mobile number'] ?? headerMap['phone number'] ?? headerMap['contact'] ?? headerMap['mobile_number'];
      const dateColIdx = headerMap['date'] ?? headerMap['preferred date'] ?? headerMap['appointment date'];
      const nameColIdx = headerMap['name'] ?? headerMap['patient name'] ?? headerMap['full name'] ?? headerMap['patient_name'];
      const timeColIdx = headerMap['time'] ?? headerMap['preferred time'] ?? headerMap['appointment time'];
      const tokenColIdx = headerMap['token'] ?? headerMap['booking token'] ?? headerMap['token id'];
      const statusColIdx = headerMap['status'] ?? headerMap['booking status'];
      
      const newPhone = normalizeIndianMobile(payload.phone || payload.mobile || '');
      const newDate = (payload.date || '').trim();
      const newTime = (payload.time || '').trim();
      const newName = (payload.name || '').trim().toLowerCase();

      let mobileActiveCount = 0;
      
      const today = new Date();
      const ty = today.getFullYear();
      const tm = String(today.getMonth() + 1).padStart(2, '0');
      const td = String(today.getDate()).padStart(2, '0');
      const todayStr = `${ty}-${tm}-${td}`;

      if (phoneColIdx !== undefined && dateColIdx !== undefined && nameColIdx !== undefined && timeColIdx !== undefined && newDate) {
        for (let i = 0; i < dataRange.length; i++) {
          // Skip cancelled or rejected appointments if status column exists
          if (statusColIdx !== undefined) {
             const status = (dataRange[i][statusColIdx] || '').toString().toLowerCase();
             if (status === 'cancelled' || status === 'rejected' || status === 'deleted') continue;
          }

          let rowDate = dataRange[i][dateColIdx];
          
          let formattedRowDate = '';
          if (rowDate instanceof Date) {
            const y = rowDate.getFullYear();
            const m = String(rowDate.getMonth() + 1).padStart(2, '0');
            const d = String(rowDate.getDate()).padStart(2, '0');
            formattedRowDate = `${y}-${m}-${d}`;
          } else {
            formattedRowDate = (rowDate || '').toString().trim();
          }

          // Count active upcoming appointments for this mobile number
          if (formattedRowDate >= todayStr) {
            const rowPhone = normalizeIndianMobile(dataRange[i][phoneColIdx] || '');
            if (rowPhone === newPhone) {
              mobileActiveCount++;
            }
          }

          if (formattedRowDate === newDate) {
            
            // 1. Check if the exact time slot is already taken by ANYONE
            if (newTime) {
              let rowTime = displayDataRange[i][timeColIdx];
              rowTime = normalizeTimeString(rowTime);
              if (rowTime === normalizeTimeString(newTime)) {
                return ContentService
                  .createTextOutput(JSON.stringify({
                    status: false,
                    success: false,
                    message: 'This time slot is already booked. Please select a different time.',
                    error: 'This time slot is already booked. Please select a different time.'
                  }))
                  .setMimeType(ContentService.MimeType.JSON);
              }
            }

            const rowPhone = normalizeIndianMobile(dataRange[i][phoneColIdx] || '');
            if (rowPhone === newPhone) {
              const rowName = (dataRange[i][nameColIdx] || '').toString().trim().toLowerCase();
              if (rowName === newName) {
                const existingToken = dataRange[i][tokenColIdx] || 'UNKNOWN';
                return ContentService
                  .createTextOutput(JSON.stringify({
                    status: false,
                    success: false,
                    isDuplicate: true,
                    message: 'Already booked on this date.',
                    data: { token: existingToken }
                  }))
                  .setMimeType(ContentService.MimeType.JSON);
              }
            }
          }
        }
        
        if (mobileActiveCount >= 3) {
          return ContentService
            .createTextOutput(JSON.stringify({
              status: false,
              success: false,
              message: 'Maximum of 3 active appointments allowed per mobile number.',
              error: 'Maximum of 3 active appointments allowed per mobile number.'
            }))
            .setMimeType(ContentService.MimeType.JSON);
        }
      }
    }
    // --- End Duplicate Check ---

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

    // Calculate the next token ID based on existing tokens in the sheet
    let maxTokenNum = 1000;
    const tokColIdx = headerMap['token'] ?? headerMap['booking token'] ?? headerMap['token id'];
    if (lastRow > 1 && tokColIdx !== undefined) {
      const tokenValues = sheet.getRange(2, tokColIdx + 1, lastRow - 1, 1).getValues();
      for (let i = 0; i < tokenValues.length; i++) {
        const t = tokenValues[i][0];
        if (t && typeof t === 'string' && t.startsWith('DDC-')) {
          const num = parseInt(t.substring(4), 10);
          if (!isNaN(num) && num > maxTokenNum) {
            maxTokenNum = num;
          }
        }
      }
    } else if (lastRow > 1) {
      // Fallback: If no token column is found but rows exist, use lastRow to guarantee unique tokens
      maxTokenNum = 1000 + lastRow - 1;
    }
    const token = 'DDC-' + String(maxTokenNum + 1);

    // Set parameters in their correct columns based on headers
    setVal(['token', 'booking token', 'token id'], token);
    setVal(['timestamp', 'time stamp', 'date & time'], new Date());
    setVal(['name', 'patient name', 'full name', 'patient_name'], payload.name || '');
    setVal(['phone', 'mobile', 'mobile number', 'phone number', 'contact', 'mobile_number'], payload.phone || payload.mobile || '');
    setVal(['email', 'patient email', 'email address'], payload.email || '');
    setVal(['age'], payload.age || '');
    setVal(['gender', 'sex'], payload.gender || '');
    setVal(['department', 'treatment', 'service', 'concern', 'complaint', 'chief complaint', 'chief_complaint', 'dental concern'], payload.department || '');
    setVal(['problem', 'additional details', 'notes', 'description'], payload.problem || '');
    setVal(['date', 'preferred date', 'appointment date'], payload.date || '');
    setVal(['time', 'preferred time', 'appointment time'], payload.time ? "'" + payload.time : '');

    // Set default values for administrative columns in sheet
    setVal(['doctor', 'assigned doctor'], 'Dr. Nagendra Chauhan');
    setVal(['status', 'booking status'], 'Pending');
    setVal(['source', 'booking source'], 'Website');
    setVal(['created by', 'created_by'], 'Web Intake');
    setVal(['last updated', 'last_updated'], new Date());

    sheet.appendRow(rowData);
    const notification = notifyAppointment(payload, token, {
      clinicMobileNumber: clinicMobileNumber,
      clinicWhatsAppNumber: clinicWhatsAppNumber,
      clinicEmail: clinicEmail,
      whatsappGraphApiVersion: whatsappGraphApiVersion,
      whatsappPhoneNumberId: whatsappPhoneNumberId,
      whatsappAccessToken: whatsappAccessToken,
      whatsappTemplateLanguage: whatsappTemplateLanguage,
      whatsappPatientTemplateName: whatsappPatientTemplateName,
      whatsappDoctorTemplateName: whatsappDoctorTemplateName,
      fast2SmsApiKey: fast2SmsApiKey
    });

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
  } finally {
    if (lock) {
      lock.releaseLock();
    }
  }
}

function doGet(e) {
  const sheetId = '1cY9edL6177CGhRcXeoFq_0xlbFohqSWHpR-vFP40oQU';
  const sheetName = 'Appointments';

  try {
    const spreadsheet = SpreadsheetApp.openById(sheetId);
    let sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet) {
      return ContentService
        .createTextOutput(JSON.stringify({ status: false, error: 'Sheet not found: ' + sheetName }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    // Check if we are asking for booked slots
    if (e && e.parameter && e.parameter.action === 'getSlots') {
      const lastRow = sheet.getLastRow();
      const lastCol = sheet.getLastColumn();
      
      if (lastRow < 2) {
        return ContentService.createTextOutput(JSON.stringify({ status: true, bookedSlots: [] })).setMimeType(ContentService.MimeType.JSON);
      }
      
      const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
      const headerMap = {};
      headers.forEach((h, i) => { if (h) headerMap[h.toString().trim().toLowerCase()] = i; });
      
      const dateCol = headerMap['date'] ?? headerMap['preferred date'] ?? headerMap['appointment date'];
      const timeCol = headerMap['time'] ?? headerMap['preferred time'] ?? headerMap['appointment time'];
      const statusCol = headerMap['status'] ?? headerMap['booking status'];
      
      const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
      const displayData = sheet.getRange(2, 1, lastRow - 1, lastCol).getDisplayValues();
      const bookedSlots = [];
      
      for (let i = 0; i < data.length; i++) {
         if (statusCol !== undefined) {
            const status = (data[i][statusCol] || '').toString().toLowerCase();
            if (status === 'cancelled' || status === 'rejected' || status === 'deleted') continue;
         }
         
         let rDate = data[i][dateCol];
         let formattedDate = '';
         if (rDate instanceof Date) {
            const y = rDate.getFullYear();
            const m = String(rDate.getMonth() + 1).padStart(2, '0');
            const d = String(rDate.getDate()).padStart(2, '0');
            formattedDate = `${y}-${m}-${d}`;
         } else {
            formattedDate = (rDate || '').toString().trim();
         }
         
         // Safely extract Time from Google Sheet's display value
         let rTime = normalizeTimeString(displayData[i][timeCol]);
         
         if (formattedDate && rTime) {
            bookedSlots.push({ date: formattedDate, time: rTime });
         }
      }
      
      return ContentService.createTextOutput(JSON.stringify({ status: true, bookedSlots: bookedSlots })).setMimeType(ContentService.MimeType.JSON);
    }
    
    // Default GET response
    const headers = sheet.getRange(1, 1, 1, Math.max(1, sheet.getLastColumn())).getValues()[0];
    return ContentService
      .createTextOutput(JSON.stringify({
        status: true,
        sheetName: sheetName,
        headers: headers
      }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function parsePayload(e) {
  if (!e || !e.postData || !e.postData.contents) {
    return {};
  }

  const contents = e.postData.contents;

  // Attempt JSON parsing first, regardless of Content-Type header
  try {
    return JSON.parse(contents);
  } catch (err) {
    // Fall back to form parameters if JSON parsing fails
    return e.parameter || {};
  }
}

function notifyAppointment(payload, token, config) {
  const patientMessage = [
    'Divya Dental Clinic appointment booked.',
    'Token: ' + token + '.',
    'Date: ' + (payload.date || '-') + ', Time: ' + (payload.time || '-') + '.',
    'Call 9415964504 for help.'
  ].join(' ');

  const doctorMessage = [
    'New appointment ' + token + '.',
    'Name: ' + (payload.name || '-'),
    'Mobile: ' + (payload.mobile || payload.phone || '-'),
    'Date: ' + (payload.date || '-'),
    'Time: ' + (payload.time || '-'),
    'Dept: ' + (payload.department || '-')
  ].join(' | ');

  const patientWhatsApp = sendPatientWhatsApp(payload, token, config);
  const doctorWhatsApp = sendDoctorWhatsApp(payload, token, config);
  const patientSms = sendSms(payload.mobile || payload.phone, patientMessage, config.fast2SmsApiKey);
  const doctorSms = sendSms(config.clinicMobileNumber, doctorMessage, config.fast2SmsApiKey);
  const patientEmail = sendPatientEmail(payload, token);
  const doctorEmail = sendDoctorEmail(payload, token, config);

  return {
    patientWhatsApp: patientWhatsApp.sent,
    doctorWhatsApp: doctorWhatsApp.sent,
    patientWhatsAppMessage: patientWhatsApp.message,
    doctorWhatsAppMessage: doctorWhatsApp.message,
    patientSms: patientSms.sent,
    doctorSms: doctorSms.sent,
    patientSmsMessage: patientSms.message,
    doctorSmsMessage: doctorSms.message,
    patientEmail: patientEmail.sent,
    doctorEmail: doctorEmail.sent,
    patientEmailMessage: patientEmail.message,
    doctorEmailMessage: doctorEmail.message
  };
}

function sendPatientEmail(payload, token) {
  if (!payload.email) {
    return { sent: false, message: 'Patient email not provided.' };
  }

  const subject = 'Divya Dental Clinic Appointment Token ' + token;
  const body = [
    'Dear ' + (payload.name || 'Patient') + ',',
    '',
    'Your appointment request has been received.',
    '',
    'Token: ' + token,
    'Date: ' + (payload.date || '-'),
    'Time: ' + (payload.time || '-'),
    'Department: ' + (payload.department || '-'),
    '',
    'Clinic phone: +91 94159 64504',
    'Divya Dental Clinic'
  ].join('\n');

  return sendEmail(payload.email, subject, body);
}

function sendDoctorEmail(payload, token, config) {
  const subject = 'New Appointment ' + token + ' - ' + (payload.name || 'Patient');
  const body = [
    'New appointment received from the website.',
    '',
    'Token: ' + token,
    'Name: ' + (payload.name || '-'),
    'Mobile: ' + (payload.mobile || payload.phone || '-'),
    'Email: ' + (payload.email || '-'),
    'Age: ' + (payload.age || '-'),
    'Gender: ' + (payload.gender || '-'),
    'Department: ' + (payload.department || '-'),
    'Date: ' + (payload.date || '-'),
    'Time: ' + (payload.time || '-'),
    'Problem: ' + (payload.problem || '-')
  ].join('\n');

  return sendEmail(config.clinicEmail, subject, body);
}

function sendEmail(to, subject, body) {
  if (!to) {
    return { sent: false, message: 'Email address is missing.' };
  }

  try {
    MailApp.sendEmail({
      to: to,
      subject: subject,
      body: body,
      name: 'Divya Dental Clinic'
    });
    return { sent: true, message: 'Email sent.' };
  } catch (err) {
    return { sent: false, message: err.message };
  }
}

function sendPatientWhatsApp(payload, token, config) {
  const parameters = [
    payload.name || 'Patient',
    token,
    payload.date || '-',
    payload.time || '-',
    '9415964504'
  ];

  return sendWhatsAppTemplate(payload.mobile || payload.phone, config.whatsappPatientTemplateName, parameters, config);
}

function sendDoctorWhatsApp(payload, token, config) {
  const parameters = [
    token,
    payload.name || '-',
    payload.mobile || payload.phone || '-',
    payload.department || '-',
    payload.date || '-',
    payload.time || '-',
    payload.problem || '-'
  ];

  return sendWhatsAppTemplate(config.clinicWhatsAppNumber, config.whatsappDoctorTemplateName, parameters, config);
}

function sendWhatsAppTemplate(phone, templateName, bodyParameters, config) {
  const number = normalizeWhatsAppNumber(phone);

  if (!number) {
    return { sent: false, message: 'Missing or invalid WhatsApp number.' };
  }

  if (!config.whatsappPhoneNumberId || !config.whatsappAccessToken || !templateName) {
    return { sent: false, message: 'WhatsApp Business API is not configured.' };
  }

  const payload = {
    messaging_product: 'whatsapp',
    to: number,
    type: 'template',
    template: {
      name: templateName,
      language: {
        code: config.whatsappTemplateLanguage
      },
      components: [{
        type: 'body',
        parameters: bodyParameters.map(function(value) {
          return {
            type: 'text',
            text: value.toString()
          };
        })
      }]
    }
  };

  try {
    const response = UrlFetchApp.fetch(
      'https://graph.facebook.com/' + config.whatsappGraphApiVersion + '/' + config.whatsappPhoneNumberId + '/messages',
      {
        method: 'post',
        contentType: 'application/json',
        headers: {
          Authorization: 'Bearer ' + config.whatsappAccessToken
        },
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      }
    );

    const statusCode = response.getResponseCode();
    const body = response.getContentText();
    return {
      sent: statusCode >= 200 && statusCode < 300,
      message: body
    };
  } catch (err) {
    return { sent: false, message: err.message };
  }
}

function sendSms(phone, message, fast2SmsApiKey) {
  const number = normalizeIndianMobile(phone);

  if (!number) {
    return { sent: false, message: 'Missing or invalid mobile number.' };
  }

  if (!fast2SmsApiKey) {
    return { sent: false, message: 'FAST2SMS API key is not configured.' };
  }

  try {
    const response = UrlFetchApp.fetch('https://www.fast2sms.com/dev/bulkV2', {
      method: 'post',
      headers: {
        authorization: fast2SmsApiKey
      },
      payload: {
        route: 'q',
        message: message,
        language: 'english',
        flash: '0',
        numbers: number
      },
      muteHttpExceptions: true
    });

    const statusCode = response.getResponseCode();
    const body = response.getContentText();
    return {
      sent: statusCode >= 200 && statusCode < 300,
      message: body
    };
  } catch (err) {
    return { sent: false, message: err.message };
  }
}

function normalizeIndianMobile(phone) {
  if (!phone) return '';

  const digits = phone.toString().replace(/\D/g, '');
  if (digits.length === 10) return digits;
  if (digits.length === 12 && digits.indexOf('91') === 0) return digits.slice(2);

  return '';
}

function normalizeWhatsAppNumber(phone) {
  if (!phone) return '';

  const digits = phone.toString().replace(/\D/g, '');
  if (digits.length === 10) return '91' + digits;
  if (digits.length === 12 && digits.indexOf('91') === 0) return digits;

  return '';
}

function normalizeTimeString(timeStr) {
  if (!timeStr) return '';
  let normalized = timeStr.toString().trim().toUpperCase();
  // Ensure a space before AM/PM (e.g. 10:30AM -> 10:30 AM)
  normalized = normalized.replace(/(\d)(AM|PM)/, '$1 $2');
  
  // Remove seconds if present (e.g. 10:30:00 AM -> 10:30 AM)
  const parts = normalized.split(':');
  if (parts.length === 3) {
    const ampmMatch = parts[2].match(/(AM|PM)/);
    const ampm = ampmMatch ? ampmMatch[0] : '';
    normalized = parts[0] + ':' + parts[1] + (ampm ? ' ' + ampm : '');
  }
  
  return normalized.trim();
}
