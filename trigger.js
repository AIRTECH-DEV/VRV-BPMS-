// ============================================================
// trigger.gs  —  BPMS VRV Auto-Actual Timestamp
// Column mapping verified from debugPrintColumns() log
//
// Flow: All statuses = Done → Actual gets NOW() timestamp
//       → Next step Planned formula auto-calculates
//       → Chain continues Step 2 → Step 29
//
// Setup: Run installTrigger() ONCE from Apps Script toolbar
// ============================================================


const STEPS = [
  // STEP 2  — I=Actual(9),  statuses: J(10) K(11) L(12) M(13)
  { step: 2,  actual: 9,   statuses: [10, 11, 12, 13] },

  // STEP 3  — Q=Actual(17), statuses: R(18)
  { step: 3,  actual: 17,  statuses: [18] },

  // STEP 4  — U=Actual(21), statuses: V(22)
  { step: 4,  actual: 21,  statuses: [22] },

  // STEP 5  — Y=Actual(25), statuses: Z(26)
  { step: 5,  actual: 25,  statuses: [26] },

  // STEP 6  — AC=Actual(29), statuses: AD(30)
  // AE=Upload DC(31), AF=Send DC on Mail(32) — not statuses
  { step: 6,  actual: 29,  statuses: [30] },

  // STEP 7  — AI=Actual(35), statuses: AJ(36)
  { step: 7,  actual: 35,  statuses: [36] },

  // STEP 8  — AM=Actual(39), statuses: AN(40)
  // AO=Upload Measurement Report(41) — not a status
  { step: 8,  actual: 39,  statuses: [40] },

  // STEP 9  — AS=Actual(45), statuses: AT(46)
  { step: 9,  actual: 45,  statuses: [46] },

  // STEP 10 — AW=Actual(49), statuses: AX(50)
  { step: 10, actual: 49,  statuses: [50] },

  // STEP 11 — BA=Actual(53), statuses: BB(54)
  { step: 11, actual: 53,  statuses: [54] },

  // STEP 12 — BE=Actual(57), statuses: BF(58)
  { step: 12, actual: 57,  statuses: [58] },

  // STEP 13 — BI=Actual(61), statuses: BJ(62)
  // BK=Type Of Machine(63) — not a status
  { step: 13, actual: 61,  statuses: [62] },

  // STEP 14 — BN=Actual(66), statuses: BO(67)
  // BP=Upload Daikin Invoice(68) — not a status
  { step: 14, actual: 66,  statuses: [67] },

  // STEP 15 — BT=Actual(72), statuses: BU(73)
  // BR=Technician Name(70), BW=Remarks(75) — not statuses
  { step: 15, actual: 72,  statuses: [73] },

  // STEP 16 — BY=Actual(77), statuses: BZ(78)
  { step: 16, actual: 77,  statuses: [78] },

  // STEP 17 — CC=Actual(81), statuses: CD(82)
  { step: 17, actual: 81,  statuses: [82] },

  // STEP 18 — CG=Actual(85), statuses: CH(86)
  { step: 18, actual: 85,  statuses: [86] },

  // STEP 19 — CK=Actual(89), statuses: CL(90)
  { step: 19, actual: 89,  statuses: [90] },

  // STEP 20 — CO=Actual(93), statuses: CP(94)
  { step: 20, actual: 93,  statuses: [94] },

  // STEP 21 — CS=Actual(97), statuses: CT(98)
  { step: 21, actual: 97,  statuses: [98] },

  // STEP 22 — CW=Actual(101), statuses: CX(102)
  // CY=Type of Machine(103) — not a status
  { step: 22, actual: 101, statuses: [102] },

  // STEP 23 — DB=Actual(106), statuses: DC(107)
  { step: 23, actual: 106, statuses: [107] },

  // STEP 24 — DF=Actual(110), statuses: DG(111)
  { step: 24, actual: 110, statuses: [111] },

  // STEP 25 — DJ=Actual(114), statuses: DK(115)
  // DL=Upload Measurement Report(116) — not a status
  { step: 25, actual: 114, statuses: [115] },

  // STEP 26 — DO=Actual(119), statuses: DP(120)
  { step: 26, actual: 119, statuses: [120] },

  // STEP 27 — DS=Actual(123), statuses: DT(124)
  // DU=Upload Final Bill(125), DV=Upload Final As Built(126) — not statuses
  { step: 27, actual: 123, statuses: [124] },

  // STEP 28 — DY=Actual(129), statuses: DZ(130)
  { step: 28, actual: 129, statuses: [130] },

  // STEP 29 — ED=Actual(134), statuses: EE(135)
  { step: 29, actual: 134, statuses: [135] },
];

const DATA_START_ROW = 15;
const DONE_VALUE     = 'done';


// ============================================================
// onEdit trigger
// ============================================================
function onEditBPMS(e) {
  try {
    const sheet = e.range.getSheet();
    if (sheet.getName() !== 'BPMS VRV') return;

    const editedRow = e.range.getRow();
    const editedCol = e.range.getColumn();
    if (editedRow < DATA_START_ROW) return;

    const matchedStep = STEPS.find(s => s.statuses.includes(editedCol));
    if (!matchedStep) return;

    checkAndSetActual(sheet, editedRow, matchedStep);

  } catch (err) {
    Logger.log('onEditBPMS error: ' + err.toString());
  }
}


// ============================================================
// Core logic
// ============================================================
function checkAndSetActual(sheet, row, stepDef) {
  const actualCol     = stepDef.actual;
  const currentActual = sheet.getRange(row, actualCol).getValue();

  const statusValues = stepDef.statuses.map(col =>
    sheet.getRange(row, col).getValue().toString().toLowerCase().trim()
  );

  const allDone = statusValues.every(v => v === DONE_VALUE);

  if (allDone) {
    // Only write if Actual is blank — never overwrite existing timestamp
    if (currentActual === '' || currentActual === null) {
      const now = new Date();
      sheet.getRange(row, actualCol)
           .setValue(now)
           .setNumberFormat('dd/MM/yyyy HH:mm:ss');
      Logger.log('Step ' + stepDef.step + ' row ' + row + ': Actual → ' + now);
    }
  } else {
    // Status un-done — clear Actual so next Planned blanks out too
    if (currentActual !== '' && currentActual !== null) {
      sheet.getRange(row, actualCol).clearContent();
      Logger.log('Step ' + stepDef.step + ' row ' + row + ': Actual cleared');
    }
  }
}


// ============================================================
// Run ONCE to install trigger
// ============================================================
function installTrigger() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'onEditBPMS') {
      ScriptApp.deleteTrigger(t);
    }
  });

  ScriptApp.newTrigger('onEditBPMS')
    .forSpreadsheet(ss)
    .onEdit()
    .create();

  ss.toast('Trigger installed! Auto-Actual active.', 'Setup', 5);
  Logger.log('installTrigger: done');
}


// ============================================================
// Run once to backfill Actual for existing rows where
// all statuses are already Done but Actual is blank
// ============================================================
function backfillActuals() {
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const sheet   = ss.getSheetByName('BPMS VRV');
  const lastRow = sheet.getLastRow();
  let filled    = 0;

  ss.toast('Backfilling actuals...', 'BPMS', 10);

  for (let row = DATA_START_ROW; row <= lastRow; row++) {
    if (sheet.getRange(row, 1).getValue() === '') continue;

    for (const stepDef of STEPS) {
      const currentActual = sheet.getRange(row, stepDef.actual).getValue();
      if (currentActual !== '' && currentActual !== null) continue;

      const statusValues = stepDef.statuses.map(col =>
        sheet.getRange(row, col).getValue().toString().toLowerCase().trim()
      );

      if (statusValues.every(v => v === DONE_VALUE)) {
        const orderTs = sheet.getRange(row, 1).getValue();
        sheet.getRange(row, stepDef.actual)
             .setValue(orderTs)
             .setNumberFormat('dd/MM/yyyy HH:mm:ss');
        filled++;
        Logger.log('Backfilled step ' + stepDef.step + ' row ' + row);
      }
    }
  }

  ss.toast('Backfill done — ' + filled + ' actuals set.', 'BPMS', 5);
}