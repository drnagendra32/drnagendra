function validateAppointmentPayload(payload) {
  if (!payload.name || !payload.name.trim()) {
    throw new Error('Patient name is required.');
  }
  if (!payload.phone && !payload.mobile) {
    throw new Error('Mobile number is required.');
  }
  if (!payload.date) {
    throw new Error('Preferred appointment date is required.');
  }
  if (!payload.time) {
    throw new Error('Preferred appointment time is required.');
  }
  return true;
}
