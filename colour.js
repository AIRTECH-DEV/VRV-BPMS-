// ============================================================
// BPMS VRV - Auto Row Color based on Step Status
// COMPLETE FINAL VERSION v3
// ============================================================

const COLORS = {
  GREEN:      '#C6EFCE',  // All done - Project Complete
  BLUE:       '#BDD7EE',  // Commissioning in process
  BRIGHT_RED: '#FF0000',  // 1-27 done, 28 no status, 29 done
  PINK:       '#E6B8C2',  // 1-27 done, 28 & 29 pending/empty
  ORANGE:     '#F4B942',  // 1-13 done, 14 in process or empty
  CREAM:      '#FFF2CC',  // Step 7/8/9 in process
  BROWN:      '#76ff14',  // Step 3 in progress
  WHITE:      '#FFFFFF',  // Default
};

function colNum(col) {
  if (!col || typeof col !== 'string') return null;
  col = col.toUpperCase().trim();
  let n = 0;
  for (let i = 0; i < col.length; i++) {
    n = n * 26 + (col.charCodeAt(i) - 64);
  }
  return n;
}

// ============================================================
// ALL STATUS COLUMNS - verified
// ============================================================
const STATUS_COLS = {
  step_2:  colNum('M'),   // Working drawing
  step_3:  colNum('R'),   // Delivery challan
  step_6:  colNum('AD'),  // DC delivery
  step_7:  colNum('AJ'),  // Copper piping
  step_8:  colNum('AN'),  // Invoice submission
  step_9:  colNum('AT'),  // CRM report
  step_13: colNum('BJ'),  // Store plan / IDU ODU
  step_14: colNum('BO'),  // CRM confirm delivery
  step_15: colNum('BU'),  // Installation started
  step_16: colNum('BZ'),  // Installation measurements
  step_27: colNum('DT'),  // Final RA bill
  step_28: colNum('DZ'),  // Payment confirmation ✅
  step_29: colNum('EE'),  // Handover ✅
};

// ============================================================
// onEdit - auto triggers on every cell change
// ============================================================
function onEdit(e) {
  const sheet = e.source.getActiveSheet();
  const row = e.range.getRow();
  if (row < 14) return;
  colorRow(sheet, row);
}

// ============================================================
// colorAllRows - run once manually for all existing rows
// ============================================================
function colorAllRows() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const lastRow = sheet.getLastRow();
  SpreadsheetApp.getActiveSpreadsheet().toast('Coloring rows, please wait...', 'BPMS', 5);
  for (let row = 14; row <= lastRow; row++) {
    colorRow(sheet, row);
  }
  SpreadsheetApp.getActiveSpreadsheet().toast('Done! All rows colored.', 'BPMS', 3);
}

// ============================================================
// MAIN COLOR LOGIC
// ============================================================
function colorRow(sheet, row) {
  try {
    const lastCol = sheet.getLastColumn();
    const range = sheet.getRange(row, 1, 1, lastCol);

    const s2  = getStatus(sheet, row, STATUS_COLS.step_2);
    const s3  = getStatus(sheet, row, STATUS_COLS.step_3);
    const s6  = getStatus(sheet, row, STATUS_COLS.step_6);
    const s7  = getStatus(sheet, row, STATUS_COLS.step_7);
    const s8  = getStatus(sheet, row, STATUS_COLS.step_8);
    const s9  = getStatus(sheet, row, STATUS_COLS.step_9);
    const s13 = getStatus(sheet, row, STATUS_COLS.step_13);
    const s14 = getStatus(sheet, row, STATUS_COLS.step_14);
    const s15 = getStatus(sheet, row, STATUS_COLS.step_15);
    const s16 = getStatus(sheet, row, STATUS_COLS.step_16);
    const s27 = getStatus(sheet, row, STATUS_COLS.step_27);
    const s28 = getStatus(sheet, row, STATUS_COLS.step_28);
    const s29 = getStatus(sheet, row, STATUS_COLS.step_29);

    // helper - is step pending (empty or in progress)
    const pending = (s) => s === '' || s === 'in progress';
    const inProc  = (s) => s === 'in progress';
    const done    = (s) => s === 'done';

    let color = COLORS.WHITE;

    // ✅ GREEN — Steps 1-29 all Done
    if (done(s27) && done(s28) && done(s29)) {
      color = COLORS.GREEN;

    // 🔵 BLUE — Steps 1-28 Done, Step 29 In Progress
    // OR Steps 1-27 Done, Step 28 In Progress/empty, Step 29 In Progress
    } else if (
      (done(s27) && done(s28) && inProc(s29)) ||
      (done(s27) && pending(s28) && inProc(s29))
    ) {
      color = COLORS.BLUE;

    // 🔴 BRIGHT RED — Steps 1-27 Done, Step 28 In Progress/empty, Step 29 Done
    } else if (done(s27) && pending(s28) && done(s29)) {
      color = COLORS.BRIGHT_RED;

    // 🌸 PINK — Steps 1-27 Done, Step 28 & 29 both In Progress/empty
 
   } else if (done(s27) && s29 === '') {
  color = COLORS.PINK;

    // 🟠 ORANGE — Steps 1-13 Done, Step 14 In Progress or empty
    } else if (done(s13) && pending(s14)) {
      color = COLORS.ORANGE;

    // 🟡 LIGHT YELLOW — Steps 1-6 Done, Step 7/8/9 any In Progress
    } else if (done(s6) && (inProc(s7) || inProc(s8) || inProc(s9))) {
      color = COLORS.CREAM;

    // 🟤 BROWN — Steps 1-2 Done, Step 3 In Progress
    } else if (done(s2) && inProc(s3)) {
      color = COLORS.BROWN;

    // ⚪ WHITE — nothing started
    } else {
      color = COLORS.WHITE;
    }

    range.setBackground(color);

  } catch(err) {
    Logger.log('Error on row ' + row + ': ' + err.message);
  }
}

// ============================================================
// HELPER - get status as lowercase string
// ============================================================
function getStatus(sheet, row, col) {
  if (!col) return '';
  try {
    const val = sheet.getRange(row, col).getValue();
    return val ? val.toString().toLowerCase().trim() : '';
  } catch(e) {
    return '';
  }
}