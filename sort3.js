// ============================================================
// sort.gs  —  BPMS VRV
// ============================================================


function fixSortNow() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const bpmsSheet = ss.getSheetByName('BPMS VRV');
  ss.toast('Sorting...', 'BPMS', 3);
  sortBPMSByTimestamp(bpmsSheet);
  ss.toast('Coloring...', 'BPMS', 3);
  colorAllRowsFast(bpmsSheet);
  ss.toast('Done!', 'BPMS', 3);
}


// ============================================================
// Sort rows 15+ by Col A timestamp — newest first
// ============================================================
function sortBPMSByTimestamp(bpmsSheet) {
  const sheet = bpmsSheet ||
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName('BPMS VRV');
  const lastRow = sheet.getLastRow();
  if (lastRow < 15) return;

  const colAData = sheet.getRange(15, 1, lastRow - 14, 1).getValues();
  let actualLastRow = 14;
  for (let i = 0; i < colAData.length; i++) {
    if (colAData[i][0] !== '') actualLastRow = 15 + i;
  }
  if (actualLastRow < 15) return;

  sheet.getRange(15, 1, actualLastRow - 14, sheet.getLastColumn())
       .sort({ column: 1, ascending: false });
  Logger.log('sortBPMSByTimestamp: sorted to row ' + actualLastRow);
}


// ============================================================
// Color rows based on step status
// ============================================================
function colorAllRowsFast(bpmsSheet) {
  const sheet = bpmsSheet ||
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName('BPMS VRV');
  const lastRow = sheet.getLastRow();
  if (lastRow < 15) return;

  const totalRows = lastRow - 14;
  const lastCol   = sheet.getLastColumn();

  // Status columns — verified from your debugPrintColumns log
  const cols = {
    s2:  13,   // M  - Step 2 last status (Status-Architect)
    s3:  18,   // R  - Step 3 Status
    s6:  30,   // AD - Step 6 Status
    s7:  36,   // AJ - Step 7 Status
    s8:  40,   // AN - Step 8 Status
    s9:  46,   // AT - Step 9 Status
    s13: 62,   // BJ - Step 13 Status
    s14: 67,   // BO - Step 14 Status
    s27: 124,  // DT - Step 27 Status
    s28: 130,  // DZ - Step 28 Status
    s29: 135,  // EE - Step 29 Status
  };

  function getCol(colNum) {
    return sheet.getRange(15, colNum, totalRows, 1).getValues()
      .map(r => (r[0] || '').toString().toLowerCase().trim());
  }

  const s2  = getCol(cols.s2);
  const s3  = getCol(cols.s3);
  const s6  = getCol(cols.s6);
  const s7  = getCol(cols.s7);
  const s8  = getCol(cols.s8);
  const s9  = getCol(cols.s9);
  const s13 = getCol(cols.s13);
  const s14 = getCol(cols.s14);
  const s27 = getCol(cols.s27);
  const s28 = getCol(cols.s28);
  const s29 = getCol(cols.s29);
  const tsCol = sheet.getRange(15, 1, totalRows, 1).getValues().map(r => r[0]);

  const C = {
    GREEN:  '#C6EFCE', RED:    '#FF0000', PINK:   '#E6B8C2',
    BLUE:   '#BDD7EE', CREAM:  '#FFF2CC', BROWN:  '#C4A882',
    ORANGE: '#F4B942', WHITE:  '#FFFFFF',
  };

  const done    = s => s === 'done';
  const inProc  = s => s === 'in progress';
  const pending = s => s === '' || s === 'in progress';

  for (let i = 0; i < totalRows; i++) {
    if (!tsCol[i]) continue;
    let color = C.WHITE;
    if      (done(s28[i]) && done(s29[i]))                                            color = C.GREEN;
    else if ((done(s27[i]) && done(s28[i]) && inProc(s29[i])) ||
             (done(s27[i]) && pending(s28[i]) && inProc(s29[i])))                     color = C.BLUE;
    else if (done(s27[i]) && pending(s28[i]) && done(s29[i]))                         color = C.RED;
    else if (done(s27[i]) && s29[i] === '')                                           color = C.PINK;
    else if (done(s13[i]) && pending(s14[i]))                                         color = C.ORANGE;
    else if (done(s6[i])  && (inProc(s7[i]) || inProc(s8[i]) || inProc(s9[i])))      color = C.CREAM;
    else if (done(s2[i])  && inProc(s3[i]))                                           color = C.BROWN;
    sheet.getRange(15 + i, 1, 1, lastCol).setBackground(color);
  }
  Logger.log('colorAllRowsFast: colored ' + totalRows + ' rows');
}


// ============================================================
// PLANNED FORMULA BUILDER
//
// Builds the correct Planned formula for any step and row.
// Uses ONLY verified static references — no patching, no regex.
//
// Parameters:
//   row      - sheet row number (15, 16, 17... or 2 for template)
//   startCol - column letter of the previous step's Actual
//              (e.g. 'A' for Step2 which starts from order timestamp,
//               'I' for Step3 which starts from Step2 Actual, etc.)
//   tatCol   - column letter of this step's TAT cell in row 13
//              (e.g. 'H' for Step2, 'P' for Step3, etc.)
// ============================================================
function buildPlannedFormula(row, startCol, tatCol) {
  const start = startCol + row;         // e.g. A15, I15, Q15
  const tat   = tatCol + '$13';         // e.g. H$13, P$13, AB$13

  return '=IF(OR(ISBLANK(' + start + '),ISBLANK(' + tat + ')),"",LET(' +
    '_start,' + start + ',' +
    '_tat,' + tat + '/24,' +
    '_open,$C$8,_close,$D$8,_wd,$E$8,' +
    '_hol,Holidays!$A$3:$A,' +
    '_sd,INT(_start),_st,MOD(_start,1),' +
    '_iswd,WORKDAY.INTL(_sd-1,1,_wd,_hol)=_sd,' +
    '_first,IF(_iswd,IF(_st<_open,_sd+_open,IF(_st>=_close,' +
    'WORKDAY.INTL(_sd,1,_wd,_hol)+_open,_start)),' +
    'WORKDAY.INTL(_sd,1,_wd,_hol)+_open),' +
    '_ft,MOD(_first,1),_avail,_close-MAX(_ft,_open),' +
    'IF(_tat<=_avail,_first+_tat,' +
    'LET(_rem1,_tat-_avail,_daylen,_close-_open,' +
    '_k,INT(_rem1/_daylen),_rem2,MOD(_rem1,_daylen),' +
    '_base,WORKDAY.INTL(INT(_first),1+_k,_wd,_hol),' +
    'IF(_rem2=0,WORKDAY.INTL(INT(_first),_k,_wd,_hol)+_close,' +
    '_base+_open+_rem2)))))';
}


// ============================================================
// TIME DELAY FORMULA BUILDER
// plannedCol - column letter of this step's Planned
// actualCol  - column letter of this step's Actual
// ============================================================
function buildDelayFormula(row, plannedCol, actualCol) {
  const p = plannedCol + row;
  const a = actualCol  + row;
  return '=IF(' + p + ',IF(' + a + '<>"",' +
    'IF(' + a + '>' + p + ',' + a + '-' + p + ',""),' +
    '$A$8-' + p + '),"")';
}


// ============================================================
// STEP MAP — all 29 steps
// Each entry: { planned, startCol, tat, actual, status, delay }
// All column letters verified from your debugPrintColumns() log
//
// planned   = col letter of Planned cell
// startCol  = col letter of previous step's Actual (input to formula)
// tat       = col letter of TAT cell in row 13
// actual    = col letter of Actual cell (filled by trigger)
// delay     = col letter of Time Delay cell (optional, null to skip)
// ============================================================
const STEP_MAP = [
  // Step 2: starts from A (order timestamp), TAT=H, Planned=H, Actual=I, Delay=N
  { step:2,  planned:'H',  startCol:'A',  tat:'H',  actual:'I',  delay:'N'  },
  // Step 3: starts from I (Step2 Actual), TAT=P, Planned=P, Actual=Q, Delay=S
  { step:3,  planned:'P',  startCol:'I',  tat:'P',  actual:'Q',  delay:'S'  },
  // Step 4: starts from Q, TAT=T, Planned=T, Actual=U, Delay=W
  { step:4,  planned:'T',  startCol:'Q',  tat:'T',  actual:'U',  delay:'W'  },
  // Step 5: starts from U, TAT=X, Planned=X, Actual=Y, Delay=AA
  { step:5,  planned:'X',  startCol:'U',  tat:'X',  actual:'Y',  delay:'AA' },
  // Step 6: starts from Y, TAT=AB, Planned=AB, Actual=AC, Delay=AG
  { step:6,  planned:'AB', startCol:'Y',  tat:'AB', actual:'AC', delay:'AG' },
  // Step 7: starts from AC, TAT=AH, Planned=AH, Actual=AI, Delay=AK
  { step:7,  planned:'AH', startCol:'AC', tat:'AH', actual:'AI', delay:'AK' },
  // Step 8: starts from AI, TAT=AL, Planned=AL, Actual=AM, Delay=AP
  { step:8,  planned:'AL', startCol:'AI', tat:'AL', actual:'AM', delay:'AP' },
  // Step 9: starts from AM, TAT=AR, Planned=AR, Actual=AS, Delay=AU
  { step:9,  planned:'AR', startCol:'AM', tat:'AR', actual:'AS', delay:'AU' },
  // Step 10: starts from AS, TAT=AV, Planned=AV, Actual=AW, Delay=AY
  { step:10, planned:'AV', startCol:'AS', tat:'AV', actual:'AW', delay:'AY' },
  // Step 11: starts from AW, TAT=AZ, Planned=AZ, Actual=BA, Delay=BC
  { step:11, planned:'AZ', startCol:'AW', tat:'AZ', actual:'BA', delay:'BC' },
  // Step 12: starts from BA, TAT=BD, Planned=BD, Actual=BE, Delay=BG
  { step:12, planned:'BD', startCol:'BA', tat:'BD', actual:'BE', delay:'BG' },
  // Step 13: starts from BE, TAT=BH, Planned=BH, Actual=BI, Delay=BL
  { step:13, planned:'BH', startCol:'BE', tat:'BH', actual:'BI', delay:'BL' },
  // Step 14: starts from BI, TAT=BM, Planned=BM, Actual=BN, Delay=BQ
  { step:14, planned:'BM', startCol:'BI', tat:'BM', actual:'BN', delay:'BQ' },
  // Step 15: starts from BN, TAT=BS, Planned=BS, Actual=BT, Delay=BV
 { step:15, planned:'BS', startCol:'BN', tat:'BR', actual:'BT', delay:'BV' },
  // Step 16: starts from BT, TAT=BX, Planned=BX, Actual=BY, Delay=CA
  { step:16, planned:'BX', startCol:'BT', tat:'BX', actual:'BY', delay:'CA' },
  // Step 17: starts from BY, TAT=CB, Planned=CB, Actual=CC, Delay=CE
  { step:17, planned:'CB', startCol:'BY', tat:'CB', actual:'CC', delay:'CE' },
  // Step 18: starts from CC, TAT=CF, Planned=CF, Actual=CG, Delay=CI
  { step:18, planned:'CF', startCol:'CC', tat:'CF', actual:'CG', delay:'CI' },
  // Step 19: starts from CG, TAT=CJ, Planned=CJ, Actual=CK, Delay=CM
  { step:19, planned:'CJ', startCol:'CG', tat:'CJ', actual:'CK', delay:'CM' },
  // Step 20: starts from CK, TAT=CN, Planned=CN, Actual=CO, Delay=CQ
  { step:20, planned:'CN', startCol:'CK', tat:'CN', actual:'CO', delay:'CQ' },
  // Step 21: starts from CO, TAT=CR, Planned=CR, Actual=CS, Delay=CU
  { step:21, planned:'CR', startCol:'CO', tat:'CR', actual:'CS', delay:'CU' },
  // Step 22: starts from CS, TAT=CV, Planned=CV, Actual=CW, Delay=CZ
  { step:22, planned:'CV', startCol:'CS', tat:'CV', actual:'CW', delay:'CZ' },
  // Step 23: starts from CW, TAT=DA, Planned=DA, Actual=DB, Delay=DD
  { step:23, planned:'DA', startCol:'CW', tat:'DA', actual:'DB', delay:'DD' },
  // Step 24: starts from DB, TAT=DE, Planned=DE, Actual=DF, Delay=DH
  { step:24, planned:'DE', startCol:'DB', tat:'DE', actual:'DF', delay:'DH' },
  // Step 25: starts from DF, TAT=DI, Planned=DI, Actual=DJ, Delay=DM
  { step:25, planned:'DI', startCol:'DF', tat:'DI', actual:'DJ', delay:'DM' },
  // Step 26: starts from DJ, TAT=DN, Planned=DN, Actual=DO, Delay=DQ
  { step:26, planned:'DN', startCol:'DJ', tat:'DN', actual:'DO', delay:'DQ' },
  // Step 27: starts from DO, TAT=DR, Planned=DR, Actual=DS, Delay=DW
  { step:27, planned:'DR', startCol:'DO', tat:'DR', actual:'DS', delay:'DW' },
  // Step 28: starts from DS, TAT=DX, Planned=DX, Actual=DY, Delay=EA
  { step:28, planned:'DX', startCol:'DS', tat:'DX', actual:'DY', delay:'EA' },
  // Step 29: starts from DY, TAT=EC, Planned=EC, Actual=ED, Delay=EF
  { step:29, planned:'EC', startCol:'DY', tat:'EC', actual:'ED', delay:'EF' },
];


// ============================================================
// Convert column letter to number (A=1, Z=26, AA=27, AB=28...)
// ============================================================
function colLetterToNum(col) {
  let n = 0;
  for (let i = 0; i < col.length; i++) {
    n = n * 26 + (col.charCodeAt(i) - 64);
  }
  return n;
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


// ============================================================
// REBUILD ALL PLANNED FORMULAS
//
// Writes correct Planned (and Time Delay) formulas for every
// step in every data row AND in the template.
// This is the definitive fix — bypasses all corruption.
// Run once, then new orders will copy clean from template.
// ============================================================
function rebuildAllPlannedFormulas() {
  const ss            = SpreadsheetApp.getActiveSpreadsheet();
  const bpmsSheet     = ss.getSheetByName('BPMS VRV');
  const templateSheet = ss.getSheetByName('BPMS-Template');
  const lastRow       = bpmsSheet.getLastRow();

  ss.toast('Rebuilding all Planned formulas...', 'BPMS', 30);

  // ── Fix template (row 2) ───────────────────────────────────
  for (const s of STEP_MAP) {
    const plannedFormula = buildPlannedFormula(2, s.startCol, s.tat);
    templateSheet.getRange(2, colLetterToNum(s.planned))
                 .setFormula(plannedFormula)
                 .setNumberFormat('dd/MM/yyyy HH:mm:ss');

    if (s.delay) {
      const delayFormula = buildDelayFormula(2, s.planned, s.actual);
      templateSheet.getRange(2, colLetterToNum(s.delay))
                   .setFormula(delayFormula)
                   .setNumberFormat('[h]:mm');
    }
  }
  Logger.log('Template rebuilt');

  // ── Fix every data row ────────────────────────────────────
  let rowsFixed = 0;
  for (let row = 15; row <= lastRow; row++) {
    const ts = bpmsSheet.getRange(row, 1).getValue();
    if (ts === '') continue;

    for (const s of STEP_MAP) {
      const plannedFormula = buildPlannedFormula(row, s.startCol, s.tat);
      bpmsSheet.getRange(row, colLetterToNum(s.planned))
               .setFormula(plannedFormula)
               .setNumberFormat('dd/MM/yyyy HH:mm:ss');

      if (s.delay) {
        const delayFormula = buildDelayFormula(row, s.planned, s.actual);
        bpmsSheet.getRange(row, colLetterToNum(s.delay))
                 .setFormula(delayFormula)
                 .setNumberFormat('[h]:mm');
      }
    }
    rowsFixed++;
    if (rowsFixed % 5 === 0) {
      ss.toast('Fixed ' + rowsFixed + ' rows...', 'BPMS', 5);
    }
  }

  Logger.log('rebuildAllPlannedFormulas: fixed ' + rowsFixed + ' rows');
  ss.toast('Planned formulas rebuilt! Sorting...', 'BPMS', 5);

  sortBPMSByTimestamp(bpmsSheet);
  colorAllRowsFast(bpmsSheet);
  ss.toast('Complete! ' + rowsFixed + ' rows fixed.', 'BPMS', 5);
}


// ============================================================
// Copy a new order from Orders sheet → BPMS VRV
// ============================================================
function copyOrderToBPMS(ordersRow) {
  const ss            = SpreadsheetApp.getActiveSpreadsheet();
  const ordersSheet   = ss.getSheetByName('Orders');
  const bpmsSheet     = ss.getSheetByName('BPMS VRV');
  const templateSheet = ss.getSheetByName('BPMS-Template');

  if (!ordersSheet || !bpmsSheet || !templateSheet) {
    ss.toast('Sheet not found!', 'Error', 5);
    return;
  }

  if (!ordersRow) ordersRow = ordersSheet.getLastRow();

  const orderData  = ordersSheet.getRange(ordersRow, 1, 1, 6).getValues()[0];
  const newOrderId = orderData[1].toString().trim();

  // Duplicate check
  const bpmsLastRow = bpmsSheet.getLastRow();
  if (bpmsLastRow >= 15) {
    const existingIds = bpmsSheet
      .getRange(15, 2, bpmsLastRow - 14, 1)
      .getValues()
      .map(r => r[0].toString().trim());
    if (existingIds.includes(newOrderId)) {
      ss.toast('Order already exists!', 'BPMS', 3);
      return;
    }
  }

  const newRow  = Math.max(bpmsLastRow + 1, 15);
  const lastCol = bpmsSheet.getLastColumn();

  // STEP 1: Copy template (brings validations, dropdowns, formatting)
  templateSheet.getRange(2, 1, 1, lastCol)
    .copyTo(
      bpmsSheet.getRange(newRow, 1, 1, lastCol),
      SpreadsheetApp.CopyPasteType.PASTE_NORMAL,
      false
    );

  // STEP 2: Overwrite ALL Planned + Delay formulas with freshly built ones
  // This replaces whatever corrupted formulas copyTo just pasted
  for (const s of STEP_MAP) {
    const plannedFormula = buildPlannedFormula(newRow, s.startCol, s.tat);
    bpmsSheet.getRange(newRow, colLetterToNum(s.planned))
             .setFormula(plannedFormula)
             .setNumberFormat('dd/MM/yyyy HH:mm:ss');

    if (s.delay) {
      const delayFormula = buildDelayFormula(newRow, s.planned, s.actual);
      bpmsSheet.getRange(newRow, colLetterToNum(s.delay))
               .setFormula(delayFormula)
               .setNumberFormat('[h]:mm');
    }
  }

  // STEP 3: Write order data into A:F
  bpmsSheet.getRange(newRow, 1, 1, 6).setValues([orderData]);
  bpmsSheet.getRange(newRow, 1).setNumberFormat('dd/MM/yyyy HH:mm:ss');

  // STEP 4: Restore dropdown validations from template
  const validations = templateSheet
    .getRange(2, 1, 1, lastCol)
    .getDataValidations()[0];
  for (let col = 1; col <= lastCol; col++) {
    if (validations[col - 1]) {
      bpmsSheet.getRange(newRow, col).setDataValidation(validations[col - 1]);
    }
  }

  // STEP 5: Sort + Color
  sortBPMSByTimestamp(bpmsSheet);
  colorAllRowsFast(bpmsSheet);

  ss.toast('Order added & sorted!', 'BPMS', 4);
  Logger.log('copyOrderToBPMS done: ' + newOrderId + ' row ' + newRow);
}