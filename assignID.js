// const RESPONSES_SHEET_NAME = 'Orders';
// const ID_HEADER = 'Order ID';
// const ID_PREFIX = 'ORD';

// function assignOrderId(e) {
//   const lock = LockService.getScriptLock();
//   lock.waitLock(30000);
//   try {
//     const ss = SpreadsheetApp.getActive();
//     const sh = ss.getSheetByName('Orders');
//     if (!sh) throw new Error('Orders sheet not found');

//     const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
//     let idCol = headers.indexOf('Order ID') + 1;
//     if (idCol < 1) return;

//     const row = sh.getLastRow();
//     if (sh.getRange(row, idCol).getValue()) return;

//     const props = PropertiesService.getScriptProperties();
//     let counter = Number(props.getProperty('complaint_counter') || '0');

//     if (counter === 0) {
//       const idValues = sh.getRange(2, idCol, Math.max(1, sh.getLastRow() - 1)).getValues().flat();
//       const re = /^ORD-\d{6}-(\d{5})$/;
//       let maxSeq = 0;
//       idValues.forEach(v => {
//         if (typeof v === 'string') {
//           const m = v.match(re);
//           if (m) maxSeq = Math.max(maxSeq, Number(m[1]));
//         }
//       });
//       counter = maxSeq;
//     }

//     counter += 1;
//     const tz = Session.getScriptTimeZone() || 'Asia/Kolkata';
//     const datePart = Utilities.formatDate(new Date(), tz, 'yyMMdd');
//     const seqPart = Utilities.formatString('%05d', counter);
//     const id = `ORD-${datePart}-${seqPart}`;

//     sh.getRange(row, idCol).setValue(id);
//     props.setProperty('complaint_counter', String(counter));

//   } finally {
//     lock.releaseLock();
//   }
// }

// function backfillOrderIds() {
//   const ss = SpreadsheetApp.getActive();
//   const sh = ss.getSheetByName(RESPONSES_SHEET_NAME);
//   const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
//   const idCol = headers.indexOf(ID_HEADER) + 1;
//   if (idCol < 1) throw new Error(`Column "${ID_HEADER}" not found`);

//   const lastRow = sh.getLastRow();
//   const idValues = sh.getRange(2, idCol, lastRow - 1).getValues().flat();

//   for (let r = 2; r <= lastRow; r++) {
//     if (!idValues[r - 2]) {
//       assignOrderId({ range: sh.getRange(r, 1) });
//     }
//   }
//   SpreadsheetApp.getUi().alert('Backfilled Order IDs for existing rows.');
// }