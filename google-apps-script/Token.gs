function generateToken(sheet, offset = 0) {
  const lastRow = sheet.getLastRow();
  // Generates a sequential appointment token starting at DDC-1001 + offset
  return 'DDC-' + String(lastRow + 1000 + offset);
}
