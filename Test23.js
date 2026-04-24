function auditRow13TAT() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('BPMS VRV');
  
  // Check row 13 from column 60 to 150
  const startCol = 60;
  const endCol = 150;
  const values = sheet.getRange(13, startCol, 1, endCol - startCol + 1).getValues()[0];
  
  Logger.log('=== ROW 13 TAT AUDIT (cols 60-150) ===');
  for (let i = 0; i < values.length; i++) {
    const colNum = startCol + i;
    const colLetter = columnToLetter(colNum);
    const val = values[i];
    if (val !== '' && val !== null) {
      Logger.log('Col ' + colLetter + ' (' + colNum + ') = ' + val);
    }
  }
  Logger.log('=== END ===');
}

function columnToLetter(col) {
  let letter = '';
  while (col > 0) {
    const mod = (col - 1) % 26;
    letter = String.fromCharCode(65 + mod) + letter;
    col = Math.floor((col - 1) / 26);
  }
  return letter;
}