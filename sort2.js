// // ============================================================
// // sort.gs  —  BPMS VRV  |  Clean working version
// // Functions: fixSortNow, sortBPMSByTimestamp, colorAllRowsFast,
// //            copyOrderToBPMS, fixRowFormulas
// // ============================================================


// // ============================================================
// // Manual trigger — run from Apps Script toolbar to re-sort & recolor
// // ============================================================
// function fixSortNow() {
//   const ss = SpreadsheetApp.getActiveSpreadsheet();
//   const bpmsSheet = ss.getSheetByName('BPMS VRV');
//   ss.toast('Sorting...', 'BPMS', 3);
//   sortBPMSByTimestamp(bpmsSheet);
//   ss.toast('Coloring...', 'BPMS', 3);
//   colorAllRowsFast(bpmsSheet);
//   ss.toast('Done!', 'BPMS', 3);
// }


// // ============================================================
// // Sort rows 15+ by Col A timestamp — newest first
// // Sorts ENTIRE row so all step data moves with its order
// // ============================================================
// function sortBPMSByTimestamp(bpmsSheet) {
//   const sheet = bpmsSheet ||
//     SpreadsheetApp.getActiveSpreadsheet().getSheetByName('BPMS VRV');

//   const lastRow = sheet.getLastRow();
//   if (lastRow < 15) return;

//   // Find actual last row that has a timestamp in col A
//   const colAData = sheet.getRange(15, 1, lastRow - 14, 1).getValues();
//   let actualLastRow = 14;
//   for (let i = 0; i < colAData.length; i++) {
//     if (colAData[i][0] !== '') actualLastRow = 15 + i;
//   }
//   if (actualLastRow < 15) return;

//   // Sort entire row (all columns) so step data moves with its order
//   sheet.getRange(15, 1, actualLastRow - 14, sheet.getLastColumn())
//        .sort({ column: 1, ascending: false });

//   Logger.log('sortBPMSByTimestamp: sorted to row ' + actualLastRow);
// }


// // ============================================================
// // Color rows 15+ based on step status
// // Batch reads all status columns — fast, no cell-by-cell calls
// // ============================================================
// function colorAllRowsFast(bpmsSheet) {
//   const sheet = bpmsSheet ||
//     SpreadsheetApp.getActiveSpreadsheet().getSheetByName('BPMS VRV');

//   const lastRow = sheet.getLastRow();
//   if (lastRow < 15) return;

//   const totalRows = lastRow - 14;
//   const lastCol   = sheet.getLastColumn();

//   // Status column numbers (1-based) — verified from your sheet
//   const cols = {
//     s2:  13,   // M
//     s3:  16,   // P
//     s6:  29,   // AC
//     s7:  36,   // AJ
//     s8:  40,   // AN
//     s9:  44,   // AR
//     s13: 61,   // BI
//     s14: 66,   // BN
//     s15: 70,   // BR
//     s16: 76,   // BX
//     s27: 120,  // DT
//     s28: 130,  // DZ
//     s29: 135,  // EE
//   };

//   function getCol(colNum) {
//     return sheet.getRange(15, colNum, totalRows, 1).getValues()
//       .map(r => (r[0] || '').toString().toLowerCase().trim());
//   }

//   const s2  = getCol(cols.s2);
//   const s3  = getCol(cols.s3);
//   const s6  = getCol(cols.s6);
//   const s7  = getCol(cols.s7);
//   const s8  = getCol(cols.s8);
//   const s9  = getCol(cols.s9);
//   const s13 = getCol(cols.s13);
//   const s14 = getCol(cols.s14);
//   const s15 = getCol(cols.s15);
//   const s16 = getCol(cols.s16);
//   const s27 = getCol(cols.s27);
//   const s28 = getCol(cols.s28);
//   const s29 = getCol(cols.s29);
//   const tsCol = sheet.getRange(15, 1, totalRows, 1).getValues().map(r => r[0]);

//   const C = {
//     GREEN:  '#C6EFCE',
//     RED:    '#FF0000',
//     PINK:   '#E6B8C2',
//     BLUE:   '#BDD7EE',
//     CREAM:  '#FFF2CC',
//     BROWN:  '#C4A882',
//     ORANGE: '#F4B942',
//     WHITE:  '#FFFFFF',
//   };

//   const done    = s => s === 'done';
//   const inProc  = s => s === 'in progress';
//   const pending = s => s === '' || s === 'in progress';

//   for (let i = 0; i < totalRows; i++) {
//     if (!tsCol[i]) continue; // skip empty rows

//     let color = C.WHITE;

//     if      (done(s28[i]) && done(s29[i]))                                              color = C.GREEN;
//     else if ((done(s27[i]) && done(s28[i]) && inProc(s29[i])) ||
//              (done(s27[i]) && pending(s28[i]) && inProc(s29[i])))                       color = C.BLUE;
//     else if (done(s27[i]) && pending(s28[i]) && done(s29[i]))                           color = C.RED;
//     else if (done(s27[i]) && s29[i] === '')                                             color = C.PINK;
//     else if (done(s13[i]) && pending(s14[i]))                                           color = C.ORANGE;
//     else if (done(s6[i])  && (inProc(s7[i]) || inProc(s8[i]) || inProc(s9[i])))        color = C.CREAM;
//     else if (done(s2[i])  && inProc(s3[i]))                                             color = C.BROWN;

//     sheet.getRange(15 + i, 1, 1, lastCol).setBackground(color);
//   }

//   Logger.log('colorAllRowsFast: colored ' + totalRows + ' rows');
// }


// // ============================================================
// // Fix corrupted TAT formula refs in a single row
// //
// // Root cause: Google Sheets copyTo doubles the column letter
// // and shifts the row when it can't resolve the reference.
// //   P$13  →  PP<row>    (single-letter cols)
// //   AI$13 →  AIAI<row>  (two-letter cols)
// //
// // This reads all formulas in one batch, fixes them, writes back
// // in one batch — called immediately after copyTo.
// // ============================================================
// function fixRowFormulas(sheet, row) {
//   const lastCol = sheet.getLastColumn();
//   const formulas = sheet.getRange(row, 1, 1, lastCol).getFormulas()[0];

//   // Build list of [corrupted_pattern, correct_col] for every TAT column
//   // Single-letter cols that get doubled: H PP TT XX SS VV etc
//   // Two-letter cols that get doubled:    AIAI AMAM AQAQ etc
//   // We match: <doubled_col><any_digits>  and replace with <correct_col>$13

//   // All two-letter TAT columns in your sheet
//   const twoLetterCols = [
//     'AB','AI','AM','AQ','AU','AY',
//     'BC','BG','BL','BP','BV','BZ',
//     'CF','CJ','CN','CR','CV','CZ',
//     'DD','DH','DL','DP','DT','DX','EE'
//   ];

//   let changed = false;

//   const fixed = formulas.map(formula => {
//     if (!formula) return formula;
//     let f = formula;

//     // Fix doubled two-letter cols first (longer match wins, do these first)
//     // e.g. AIAI17 → AI$13
//     for (const col of twoLetterCols) {
//       const doubled = col + col;           // e.g. 'AIAI'
//       // Match doubled col immediately followed by digits (not preceded by letter)
//       const re = new RegExp('(?<![A-Z])' + doubled + '(\\d+)', 'g');
//       if (f.includes(doubled)) {
//         f = f.replace(re, col + '\\$13');
//         // Note: replace uses literal $ in replacement string
//         f = f.split(col + '\\$13').join(col + '$13');
//       }
//     }

//     // Fix doubled single-letter cols
//     // e.g. PP17 → P$13, HH16 → H$13, TT15 → T$13
//     // Match: two identical uppercase letters followed by digits,
//     // NOT preceded by another letter (to avoid matching inside longer refs)
//     f = f.replace(/(?<![A-Z])([A-Z])\1(\d+)/g, (match, letter, digits) => {
//       return letter + '$13';
//     });

//     if (f !== formula) changed = true;
//     return f;
//   });

//   if (changed) {
//     sheet.getRange(row, 1, 1, lastCol).setFormulas([fixed]);
//     Logger.log('fixRowFormulas: fixed row ' + row);
//   } else {
//     Logger.log('fixRowFormulas: no corruption found in row ' + row);
//   }
// }


// // ============================================================
// // Copy a new order from Orders sheet → BPMS VRV
// // Called from onFormSubmit trigger
// // ============================================================
// function copyOrderToBPMS(ordersRow) {
//   const ss            = SpreadsheetApp.getActiveSpreadsheet();
//   const ordersSheet   = ss.getSheetByName('Orders');
//   const bpmsSheet     = ss.getSheetByName('BPMS VRV');
//   const templateSheet = ss.getSheetByName('BPMS-Template');

//   if (!ordersSheet || !bpmsSheet) {
//     Logger.log('ERROR: Orders or BPMS VRV sheet not found');
//     return;
//   }
//   if (!templateSheet) {
//     ss.toast('BPMS-Template sheet missing!', 'Error', 5);
//     Logger.log('ERROR: BPMS-Template sheet not found');
//     return;
//   }

//   if (!ordersRow) ordersRow = ordersSheet.getLastRow();

//   const orderData  = ordersSheet.getRange(ordersRow, 1, 1, 6).getValues()[0];
//   const newOrderId = orderData[1].toString().trim();
//   Logger.log('copyOrderToBPMS: processing order ' + newOrderId);

//   // ── Duplicate check ──────────────────────────────────────
//   const bpmsLastRow = bpmsSheet.getLastRow();
//   if (bpmsLastRow >= 15) {
//     const existingIds = bpmsSheet
//       .getRange(15, 2, bpmsLastRow - 14, 1)
//       .getValues()
//       .map(r => r[0].toString().trim());
//     if (existingIds.includes(newOrderId)) {
//       ss.toast('Order already exists in BPMS VRV!', 'BPMS', 3);
//       Logger.log('Duplicate skipped: ' + newOrderId);
//       return;
//     }
//   }

//   const newRow = Math.max(bpmsLastRow + 1, 15);
//   const lastCol = bpmsSheet.getLastColumn();

//   // ── STEP 1: Copy template row → new row ──────────────────
//   // Brings formulas, dropdowns, validations
//   templateSheet.getRange(2, 1, 1, lastCol)
//     .copyTo(
//       bpmsSheet.getRange(newRow, 1, 1, lastCol),
//       SpreadsheetApp.CopyPasteType.PASTE_NORMAL,
//       false
//     );
//   Logger.log('Template copied to row ' + newRow);

//   // ── STEP 2: Fix corrupted TAT refs immediately after copy ─
//   // copyTo doubles column letters: P$13 → PP<row>
//   // fixRowFormulas corrects them back to P$13
//   fixRowFormulas(bpmsSheet, newRow);

//   // ── STEP 3: Write order data into A:F ────────────────────
//   bpmsSheet.getRange(newRow, 1, 1, 6).setValues([orderData]);
//   bpmsSheet.getRange(newRow, 1).setNumberFormat('dd/MM/yyyy HH:mm:ss');
//   Logger.log('Order data written: ' + newOrderId + ' → row ' + newRow);

//   // ── STEP 4: Restore dropdown validations from template ───
//   const validations = templateSheet
//     .getRange(2, 1, 1, lastCol)
//     .getDataValidations()[0];
//   for (let col = 1; col <= lastCol; col++) {
//     if (validations[col - 1]) {
//       bpmsSheet.getRange(newRow, col).setDataValidation(validations[col - 1]);
//     }
//   }

//   // ── STEP 5: Sort all rows + apply row colors ─────────────
//   sortBPMSByTimestamp(bpmsSheet);
//   colorAllRowsFast(bpmsSheet);

//   ss.toast('Order added & sorted!', 'BPMS', 4);
//   Logger.log('copyOrderToBPMS complete: ' + newOrderId);
// }


// // ============================================================
// // Debug helper — run manually to check what fixRowFormulas
// // would do on any specific row without writing changes
// // ============================================================
// function debugCheckRow() {
//   const ROW_TO_CHECK = 15; // ← change this to any row number
//   const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('BPMS VRV');
//   const lastCol = sheet.getLastColumn();
//   const formulas = sheet.getRange(ROW_TO_CHECK, 1, 1, lastCol).getFormulas()[0];

//   for (let col = 0; col < lastCol; col++) {
//     const f = formulas[col];
//     if (!f) continue;
//     // Flag any formula that looks corrupted (doubled col letters before digits)
//     if (/(?<![A-Z])([A-Z])\1\d+/.test(f) || /(?<![A-Z])([A-Z]{2})\2\d+/.test(f)) {
//       const colLetter = columnToLetter(col + 1);
//       Logger.log('CORRUPTED col ' + colLetter + '(' + (col+1) + '): ' + f.substring(0, 80));
//     }
//   }
//   Logger.log('debugCheckRow done for row ' + ROW_TO_CHECK);
// }


// // ============================================================
// // Utility: convert column number to letter (1=A, 28=AB etc)
// // ============================================================
// function columnToLetter(col) {
//   let letter = '';
//   while (col > 0) {
//     const mod = (col - 1) % 26;
//     letter = String.fromCharCode(65 + mod) + letter;
//     col = Math.floor((col - 1) / 26);
//   }
//   return letter;
// }