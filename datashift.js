// const RESPONSES_SHEET_NAME = 'Orders';
// const ID_HEADER = 'Order ID';
// const ID_PREFIX = 'ORD';

// function handleFormSubmit(e) {
//     const submittedSheet = e.range.getSheet().getName();
//   if (submittedSheet !== "Form Responses 9") return;
//   const ss = SpreadsheetApp.getActiveSpreadsheet();
//   const source = ss.getSheetByName("Form Responses 9");
//   const target = ss.getSheetByName("Orders");

//   if (!source || !target) {
//     Logger.log("Sheet not found!");
//     return;
//   }

//   // STEP 1 — Copy form response to Orders with column mapping
//   const sourceHeaders = source.getRange(1, 1, 1, source.getLastColumn()).getValues()[0];
//   const targetHeaders = target.getRange(1, 1, 1, target.getLastColumn()).getValues()[0];
//   const lastSourceRow = source.getLastRow();
//   const sourceData = source.getRange(lastSourceRow, 1, 1, source.getLastColumn()).getValues()[0];

//   const newRow = [];
//   for (var i = 0; i < targetHeaders.length; i++) {
//     const sourceIndex = sourceHeaders.indexOf(targetHeaders[i]);
//     newRow.push(sourceIndex !== -1 ? sourceData[sourceIndex] : "");
//   }
//   target.appendRow(newRow);

//   // STEP 2 — Generate Order ID on the row just appended
//   const idCol = targetHeaders.indexOf(ID_HEADER) + 1;
//   if (idCol < 1) return;

//   const newRow2 = target.getLastRow(); // row just appended above

//   const props = PropertiesService.getScriptProperties();
//   let counter = Number(props.getProperty('order_counter') || '0');

//   if (counter === 0) {
//     const idValues = target.getRange(2, idCol, Math.max(1, newRow2 - 1)).getValues().flat();
//     const re = /^ORD-\d{6}-(\d{5})$/;
//     let maxSeq = 0;
//     idValues.forEach(v => {
//       if (typeof v === 'string') {
//         const m = v.match(re);
//         if (m) maxSeq = Math.max(maxSeq, Number(m[1]));
//       }
//     });
//     counter = maxSeq;
//   }

//   counter += 1;
//   const tz = Session.getScriptTimeZone() || 'Asia/Kolkata';
//   const datePart = Utilities.formatDate(new Date(), tz, 'yyMMdd');
//   const seqPart = Utilities.formatString('%05d', counter);
//   const id = `ORD-${datePart}-${seqPart}`;

//   target.getRange(newRow2, idCol).setValue(id);
//   props.setProperty('order_counter', String(counter));

// // Copy new order to BPMS VRV + sort + color
// copyOrderToBPMS(target.getLastRow());

// }

// function backfillOrderIds() {
//   const ss = SpreadsheetApp.getActive();
//   const sh = ss.getSheetByName(RESPONSES_SHEET_NAME);
//   const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
//   const idCol = headers.indexOf(ID_HEADER) + 1;
//   if (idCol < 1) throw new Error('Order ID column not found');

//   const lastRow = sh.getLastRow();
//   const props = PropertiesService.getScriptProperties();
//   let counter = Number(props.getProperty('order_counter') || '0');

//   if (counter === 0) {
//     const idValues = sh.getRange(2, idCol, lastRow - 1).getValues().flat();
//     const re = /^ORD-\d{6}-(\d{5})$/;
//     let maxSeq = 0;
//     idValues.forEach(v => {
//       if (typeof v === 'string') {
//         const m = v.match(re);
//         if (m) maxSeq = Math.max(maxSeq, Number(m[1]));
//       }
//     });
//     counter = maxSeq;
//   }

//   for (let r = 2; r <= lastRow; r++) {
//     const existing = sh.getRange(r, idCol).getValue();
//     const hasTimestamp = sh.getRange(r, 1).getValue() !== '';
//     if (!existing && hasTimestamp) {
//       counter += 1;
//       const tz = Session.getScriptTimeZone() || 'Asia/Kolkata';
//       const datePart = Utilities.formatDate(new Date(), tz, 'yyMMdd');
//       const seqPart = Utilities.formatString('%05d', counter);
//       sh.getRange(r, idCol).setValue(`ORD-${datePart}-${seqPart}`);
//     }
//   }
//   props.setProperty('order_counter', String(counter));
//   SpreadsheetApp.getUi().alert('Backfilled Order IDs for existing rows.');
// }



// ============================================================
// datashift.gs  —  Form Response 9 → Orders → BPMS VRV
// ============================================================

const RESPONSES_SHEET_NAME = 'Orders';
const ID_HEADER             = 'Order ID';
const ID_PREFIX             = 'ORD';

// ============================================================
// Main trigger function
// IMPORTANT: Must be installed as a Spreadsheet onFormSubmit
// trigger — NOT a simple onEdit or run manually.
// Run installFormTrigger() once to set it up correctly.
// ============================================================
function handleFormSubmit(e) {
  try {
    // Guard: e must exist and have namedValues (form submit event)
    if (!e || !e.range) {
      Logger.log('handleFormSubmit: no event object — run installFormTrigger() to set up trigger');
      return;
    }

    // Guard: only fire for Form Responses 9
    const submittedSheet = e.range.getSheet().getName();
    if (submittedSheet !== 'Form Responses 9') {
      Logger.log('handleFormSubmit: wrong sheet — ' + submittedSheet);
      return;
    }

    const ss     = SpreadsheetApp.getActiveSpreadsheet();
    const source = ss.getSheetByName('Form Responses 9');
    const target = ss.getSheetByName('Orders');

    if (!source || !target) {
      Logger.log('handleFormSubmit: sheet not found');
      return;
    }

    // ── STEP 1: Copy form response → Orders with column mapping ──
    const sourceHeaders = source.getRange(1, 1, 1, source.getLastColumn()).getValues()[0];
    const targetHeaders = target.getRange(1, 1, 1, target.getLastColumn()).getValues()[0];
    const lastSourceRow = source.getLastRow();
    const sourceData    = source.getRange(lastSourceRow, 1, 1, source.getLastColumn()).getValues()[0];

    const newRow = [];
    for (let i = 0; i < targetHeaders.length; i++) {
      const sourceIndex = sourceHeaders.indexOf(targetHeaders[i]);
      newRow.push(sourceIndex !== -1 ? sourceData[sourceIndex] : '');
    }
    target.appendRow(newRow);
    Logger.log('handleFormSubmit: row appended to Orders');

    // ── STEP 2: Generate Order ID ─────────────────────────────
    const idCol = targetHeaders.indexOf(ID_HEADER) + 1;
    if (idCol < 1) {
      Logger.log('handleFormSubmit: Order ID column not found in Orders headers');
      return;
    }

    const newRowNum = target.getLastRow();

    const props = PropertiesService.getScriptProperties();
    let counter = Number(props.getProperty('order_counter') || '0');

    if (counter === 0) {
      // Bootstrap counter from existing IDs on first run
      const idValues = target.getRange(2, idCol, Math.max(1, newRowNum - 1)).getValues().flat();
      const re = /^ORD-\d{6}-(\d{5})$/;
      let maxSeq = 0;
      idValues.forEach(v => {
        if (typeof v === 'string') {
          const m = v.match(re);
          if (m) maxSeq = Math.max(maxSeq, Number(m[1]));
        }
      });
      counter = maxSeq;
    }

    counter += 1;
    const tz       = Session.getScriptTimeZone() || 'Asia/Kolkata';
    const datePart = Utilities.formatDate(new Date(), tz, 'yyMMdd');
    const seqPart  = Utilities.formatString('%05d', counter);
    const orderId  = 'ORD-' + datePart + '-' + seqPart;

    target.getRange(newRowNum, idCol).setValue(orderId);
    props.setProperty('order_counter', String(counter));
    Logger.log('handleFormSubmit: Order ID set → ' + orderId);

    // ── STEP 3: Copy to BPMS VRV ─────────────────────────────
    copyOrderToBPMS(newRowNum);
    Logger.log('handleFormSubmit: complete for ' + orderId);

  } catch (err) {
    Logger.log('handleFormSubmit ERROR: ' + err.toString());
  }
}


// ============================================================
// Install the form submit trigger — run this ONCE
// Removes duplicates before installing fresh
// ============================================================
function installFormTrigger() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Remove any existing handleFormSubmit triggers to avoid duplicates
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'handleFormSubmit') {
      ScriptApp.deleteTrigger(t);
      Logger.log('Removed old handleFormSubmit trigger');
    }
  });

  // Install as spreadsheet-level onFormSubmit
  ScriptApp.newTrigger('handleFormSubmit')
    .forSpreadsheet(ss)
    .onFormSubmit()
    .create();

  SpreadsheetApp.getActiveSpreadsheet()
    .toast('Form trigger installed! Form Responses 9 → Orders is now active.', 'Setup', 5);
  Logger.log('installFormTrigger: done');
}


// ============================================================
// Test helper — simulates a form submit using the last row
// of Form Responses 9 so you can test without submitting a form
// ============================================================
function testHandleFormSubmit() {
  const ss     = SpreadsheetApp.getActiveSpreadsheet();
  const source = ss.getSheetByName('Form Responses 9');

  if (!source) {
    Logger.log('testHandleFormSubmit: Form Responses 9 not found');
    return;
  }

  // Build a fake event object pointing at the last row of Form Responses 9
  const lastRow  = source.getLastRow();
  const lastCol  = source.getLastColumn();
  const fakeRange = source.getRange(lastRow, 1, 1, lastCol);

  const fakeEvent = {
    range: fakeRange,
    source: ss,
  };

  Logger.log('testHandleFormSubmit: simulating submit for row ' + lastRow);
  handleFormSubmit(fakeEvent);
}


// ============================================================
// Backfill Order IDs for any Orders rows that are missing them
// ============================================================
function backfillOrderIds() {
  const ss      = SpreadsheetApp.getActive();
  const sh      = ss.getSheetByName(RESPONSES_SHEET_NAME);
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const idCol   = headers.indexOf(ID_HEADER) + 1;

  if (idCol < 1) throw new Error('Order ID column not found in Orders sheet');

  const lastRow = sh.getLastRow();
  const props   = PropertiesService.getScriptProperties();
  let counter   = Number(props.getProperty('order_counter') || '0');

  if (counter === 0) {
    const idValues = sh.getRange(2, idCol, lastRow - 1).getValues().flat();
    const re = /^ORD-\d{6}-(\d{5})$/;
    let maxSeq = 0;
    idValues.forEach(v => {
      if (typeof v === 'string') {
        const m = v.match(re);
        if (m) maxSeq = Math.max(maxSeq, Number(m[1]));
      }
    });
    counter = maxSeq;
  }

  let filled = 0;
  for (let r = 2; r <= lastRow; r++) {
    const existing     = sh.getRange(r, idCol).getValue();
    const hasTimestamp = sh.getRange(r, 1).getValue() !== '';
    if (!existing && hasTimestamp) {
      counter += 1;
      const tz       = Session.getScriptTimeZone() || 'Asia/Kolkata';
      const datePart = Utilities.formatDate(new Date(), tz, 'yyMMdd');
      const seqPart  = Utilities.formatString('%05d', counter);
      sh.getRange(r, idCol).setValue('ORD-' + datePart + '-' + seqPart);
      filled++;
    }
  }

  props.setProperty('order_counter', String(counter));
  SpreadsheetApp.getUi().alert('Backfilled ' + filled + ' Order IDs.');
}





