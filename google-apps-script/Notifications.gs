function notifyAppointment(payload, token) {
  const message = buildAppointmentSmsMessage(payload, token);
  const doctorMessage = buildDoctorSmsMessage(payload, token);

  const patientWhatsApp = sendPatientWhatsApp(payload, token);
  const doctorWhatsApp = sendDoctorWhatsApp(payload, token);
  const patientSms = sendSms(payload.mobile || payload.phone, message);
  const doctorSms = sendSms(CLINIC_MOBILE_NUMBER, doctorMessage);
  const patientEmail = sendPatientEmail(payload, token);
  const doctorEmail = sendDoctorEmail(payload, token);

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

function sendDoctorEmail(payload, token) {
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

  return sendEmail(CLINIC_EMAIL, subject, body);
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

function sendPatientWhatsApp(payload, token) {
  const parameters = [
    payload.name || 'Patient',
    token,
    payload.date || '-',
    payload.time || '-',
    '9415964504'
  ];

  return sendWhatsAppTemplate(payload.mobile || payload.phone, WHATSAPP_PATIENT_TEMPLATE_NAME, parameters);
}

function sendDoctorWhatsApp(payload, token) {
  const parameters = [
    token,
    payload.name || '-',
    payload.mobile || payload.phone || '-',
    payload.department || '-',
    payload.date || '-',
    payload.time || '-',
    payload.problem || '-'
  ];

  return sendWhatsAppTemplate(CLINIC_WHATSAPP_NUMBER, WHATSAPP_DOCTOR_TEMPLATE_NAME, parameters);
}

function sendWhatsAppTemplate(phone, templateName, bodyParameters) {
  const number = normalizeWhatsAppNumber(phone);

  if (!number) {
    return { sent: false, message: 'Missing or invalid WhatsApp number.' };
  }

  if (!WHATSAPP_PHONE_NUMBER_ID || !WHATSAPP_ACCESS_TOKEN || !templateName) {
    return { sent: false, message: 'WhatsApp Business API is not configured.' };
  }

  const payload = {
    messaging_product: 'whatsapp',
    to: number,
    type: 'template',
    template: {
      name: templateName,
      language: {
        code: WHATSAPP_TEMPLATE_LANGUAGE
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
      'https://graph.facebook.com/' + WHATSAPP_GRAPH_API_VERSION + '/' + WHATSAPP_PHONE_NUMBER_ID + '/messages',
      {
        method: 'post',
        contentType: 'application/json',
        headers: {
          Authorization: 'Bearer ' + WHATSAPP_ACCESS_TOKEN
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

function buildAppointmentSmsMessage(payload, token) {
  return [
    'Divya Dental Clinic appointment booked.',
    'Token: ' + token + '.',
    'Date: ' + (payload.date || '-') + ', Time: ' + (payload.time || '-') + '.',
    'Call 9415964504 for help.'
  ].join(' ');
}

function buildDoctorSmsMessage(payload, token) {
  return [
    'New appointment ' + token + '.',
    'Name: ' + (payload.name || '-'),
    'Mobile: ' + (payload.mobile || payload.phone || '-'),
    'Date: ' + (payload.date || '-'),
    'Time: ' + (payload.time || '-'),
    'Dept: ' + (payload.department || '-')
  ].join(' | ');
}

function sendSms(phone, message) {
  const number = normalizeIndianMobile(phone);

  if (!number) {
    return { sent: false, message: 'Missing or invalid mobile number.' };
  }

  if (!FAST2SMS_API_KEY) {
    return { sent: false, message: 'FAST2SMS_API_KEY is not configured.' };
  }

  try {
    const response = UrlFetchApp.fetch('https://www.fast2sms.com/dev/bulkV2', {
      method: 'post',
      headers: {
        authorization: FAST2SMS_API_KEY
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
