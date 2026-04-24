// // // ============================================================
// // // OPTIMIZED Sort + Color - Much Faster!
// // // ============================================================

// // function fixSortNow() {
// //   const ss = SpreadsheetApp.getActiveSpreadsheet();
// //   const bpmsSheet = ss.getSheetByName('BPMS VRV');
// //   ss.toast('Sorting...', 'BPMS', 3);
// //   sortBPMSByTimestamp(bpmsSheet);
// //   ss.toast('Coloring...', 'BPMS', 3);
// //   colorAllRowsFast(bpmsSheet);
// //   ss.toast('Done!', 'BPMS', 3);
// // }

// // function sortBPMSByTimestamp(bpmsSheet) {
// //   const sheet = bpmsSheet ||
// //     SpreadsheetApp.getActiveSpreadsheet().getSheetByName('BPMS VRV');

// //   const lastRow = sheet.getLastRow();
// //   if (lastRow < 15) return;

// //   // Find actual last row with timestamp
// //   const colAData = sheet.getRange(15, 1, lastRow - 14, 1).getValues();
// //   let actualLastRow = 14;
// //   for (let i = 0; i < colAData.length; i++) {
// //     if (colAData[i][0] !== '') actualLastRow = 15 + i;
// //   }

// //   if (actualLastRow < 15) return;

// //   // Sort only A-F newest first
// //   sheet.getRange(15, 1, actualLastRow - 14, 6)
// //        .sort({column: 1, ascending: false});

// //   Logger.log('Sorted to row: ' + actualLastRow);
// // }




// // // ============================================================
// // // FAST color - reads ALL status columns in ONE batch call
// // // instead of cell by cell
// // // ============================================================
// // function colorAllRowsFast(bpmsSheet) {
// //   const sheet = bpmsSheet ||
// //     SpreadsheetApp.getActiveSpreadsheet().getSheetByName('BPMS VRV');

// //   const lastRow = sheet.getLastRow();
// //   if (lastRow < 15) return;

// //   const totalRows = lastRow - 14;

// //   // Read ALL needed status columns in ONE call each
// //   // Column numbers (1-based)
// //   const cols = {
// //     s2:  13,  // M  - Step 2
// //     s3:  16,  // P  - Step 3
// //     s6:  29,  // AC - Step 6
// //     s7:  36,  // AJ - Step 7
// //     s8:  40,  // AN - Step 8
// //     s9:  44,  // AR - Step 9
// //     s13: 61,  // BI - Step 13
// //     s14: 66,  // BN - Step 14
// //     s15: 70,  // BR - Step 15
// //     s16: 76,  // BX - Step 16
// //     s27: 120, // DT - Step 27
// //     s28: 130, // DZ - Step 28
// //     s29: 135, // EE - Step 29
// //   };

// //   // Batch read each status column for ALL rows at once
// //   function getColValues(colNum) {
// //     return sheet.getRange(15, colNum, totalRows, 1).getValues()
// //       .map(r => (r[0] || '').toString().toLowerCase().trim());
// //   }

// //   const s2  = getColValues(cols.s2);
// //   const s3  = getColValues(cols.s3);
// //   const s6  = getColValues(cols.s6);
// //   const s7  = getColValues(cols.s7);
// //   const s8  = getColValues(cols.s8);
// //   const s9  = getColValues(cols.s9);
// //   const s13 = getColValues(cols.s13);
// //   const s14 = getColValues(cols.s14);
// //   const s15 = getColValues(cols.s15);
// //   const s16 = getColValues(cols.s16);
// //   const s27 = getColValues(cols.s27);
// //   const s28 = getColValues(cols.s28);
// //   const s29 = getColValues(cols.s29);

// //   // Also read col A to skip empty rows
// //   const tsCol = sheet.getRange(15, 1, totalRows, 1).getValues()
// //     .map(r => r[0]);

// //   const COLORS = {
// //     GREEN:      '#C6EFCE',
// //     BRIGHT_RED: '#FF0000',
// //     LIGHT_RED:  '#EA9999',
// //     PINK:       '#E6B8C2',
// //     BLUE:       '#BDD7EE',
// //     CREAM:      '#FFF2CC',
// //     BROWN:      '#C4A882',
// //     ORANGE:     '#F4B942',
// //     WHITE:      '#FFFFFF',
// //   };

// //   const done    = s => s === 'done';
// //   const inProc  = s => s === 'in progress';
// //   const pending = s => s === '' || s === 'in progress';

// //   // Build color array for all rows
// //   const lastCol = sheet.getLastColumn();
// //   const backgrounds = [];

// //   for (let i = 0; i < totalRows; i++) {
// //     // Skip empty rows
// //     if (!tsCol[i]) {
// //       backgrounds.push(null); // no change
// //       continue;
// //     }

// //     let color = COLORS.WHITE;

// //     if (done(s28[i]) && done(s29[i])) {
// //       color = COLORS.GREEN;
// //     } else if ((done(s27[i]) && done(s28[i]) && inProc(s29[i])) ||
// //                (done(s27[i]) && pending(s28[i]) && inProc(s29[i]))) {
// //       color = COLORS.BLUE;
// //     } else if (done(s27[i]) && pending(s28[i]) && done(s29[i])) {
// //       color = COLORS.BRIGHT_RED;
// //     } else if (done(s27[i]) && s29[i] === '') {
// //       color = COLORS.PINK;
// //     } else if (done(s13[i]) && pending(s14[i])) {
// //       color = COLORS.ORANGE;
// //     } else if (done(s6[i]) && (inProc(s7[i]) || inProc(s8[i]) || inProc(s9[i]))) {
// //       color = COLORS.CREAM;
// //     } else if (done(s2[i]) && inProc(s3[i])) {
// //       color = COLORS.BROWN;
// //     }

// //     backgrounds.push(color);
// //   }

// //   // Apply colors in BATCH - one setBackgrounds call per row
// //   // Group consecutive same-color rows for efficiency
// //   for (let i = 0; i < totalRows; i++) {
// //     if (backgrounds[i] !== null) {
// //       sheet.getRange(15 + i, 1, 1, lastCol)
// //            .setBackground(backgrounds[i]);
// //     }
// //   }

// //   Logger.log('Coloring done for ' + totalRows + ' rows');
// // }

// // // ============================================================
// // // Update copyOrderToBPMS to use fast color
// // // ============================================================
// // // function copyOrderToBPMS(ordersRow) {
// // //   const ss = SpreadsheetApp.getActiveSpreadsheet();
// // //   const ordersSheet = ss.getSheetByName('Orders');
// // //   const bpmsSheet = ss.getSheetByName('BPMS VRV');

// // //   if (!ordersSheet || !bpmsSheet) {
// // //     Logger.log('Sheet not found!');
// // //     return;
// // //   }

// // //   if (!ordersRow || ordersRow === null) {
// // //     ordersRow = ordersSheet.getLastRow();
// // //   }

// // //   const orderData = ordersSheet.getRange(ordersRow, 1, 1, 6).getValues()[0];

// // //   const bpmsLastRow = bpmsSheet.getLastRow();
// // //   const newRow = Math.max(bpmsLastRow + 1, 15);

// // //   bpmsSheet.getRange(newRow, 1).setValue(orderData[0]); // Timestamp
// // //   bpmsSheet.getRange(newRow, 2).setValue(orderData[1]); // Order ID
// // //   bpmsSheet.getRange(newRow, 3).setValue(orderData[2]); // Site Engineer
// // //   bpmsSheet.getRange(newRow, 4).setValue(orderData[3]); // Project Name
// // //   bpmsSheet.getRange(newRow, 5).setValue(orderData[4]); // Location
// // //   bpmsSheet.getRange(newRow, 6).setValue(orderData[5]); // Total Order Value

// // //   sortBPMSByTimestamp(bpmsSheet);
// // //   colorAllRowsFast(bpmsSheet);

// // //   ss.toast('New order added & sorted!', 'BPMS', 4);
// // // }

// // //new test here 
// // function copyOrderToBPMS(ordersRow) {
// //   const ss = SpreadsheetApp.getActiveSpreadsheet();
// //   const ordersSheet = ss.getSheetByName('Orders');
// //   const bpmsSheet = ss.getSheetByName('BPMS VRV');

// //   if (!ordersSheet || !bpmsSheet) {
// //     Logger.log('Sheet not found!');
// //     return;
// //   }

// //   if (!ordersRow || ordersRow === null) {
// //     ordersRow = ordersSheet.getLastRow();
// //   }

// //   const orderData = ordersSheet.getRange(ordersRow, 1, 1, 6).getValues()[0];
// //   const newOrderId = orderData[1].toString().trim();

// //   Logger.log('New Order ID: ' + newOrderId);

// //   // ✅ Check if Order ID already exists in BPMS VRV (avoid duplicate)
// //   const bpmsLastRow = bpmsSheet.getLastRow();
// //   if (bpmsLastRow >= 15) {
// //     const existingIds = bpmsSheet.getRange(15, 2, bpmsLastRow - 14, 1)
// //                                  .getValues()
// //                                  .map(r => r[0].toString().trim());
// //     if (existingIds.includes(newOrderId)) {
// //       Logger.log('Order ID already exists in BPMS VRV, skipping: ' + newOrderId);
// //       ss.toast('Order already exists in BPMS VRV!', 'BPMS', 3);
// //       return;
// //     }
// //   }

// //   // Write to new row
// //   const newRow = Math.max(bpmsLastRow + 1, 15);

// //   // Write all 6 values at once
// //   bpmsSheet.getRange(newRow, 1, 1, 6).setValues([orderData]);

// //   // ✅ Fix timestamp format - show full date + time
// //   bpmsSheet.getRange(newRow, 1)
// //            .setNumberFormat('dd/MM/yyyy HH:mm:ss');

// //   Logger.log('Written to BPMS VRV row: ' + newRow);

// //   // Sort + Color
// //   sortBPMSByTimestamp(bpmsSheet);
// //   colorAllRowsFast(bpmsSheet);

// //   ss.toast('New order added & sorted!', 'BPMS', 4);
// // }


// // 5:57 25-3-26

// // ============================================================
// // BPMS VRV - Complete Sort + Color Script
// // ============================================================

// function fixSortNow() {
//   const ss = SpreadsheetApp.getActiveSpreadsheet();
//   const bpmsSheet = ss.getSheetByName('BPMS VRV');
//   ss.toast('Sorting all rows...', 'BPMS', 3);
//   sortBPMSByTimestamp(bpmsSheet);
//   ss.toast('Coloring all rows...', 'BPMS', 3);
//   colorAllRowsFast(bpmsSheet);
//   ss.toast('Done! Sorted & colored.', 'BPMS', 3);
// }

// function sortBPMSByTimestamp(bpmsSheet) {
//   const sheet = bpmsSheet ||
//     SpreadsheetApp.getActiveSpreadsheet().getSheetByName('BPMS VRV');

//   const lastRow = sheet.getLastRow();
//   if (lastRow < 15) return;

//   // Find actual last row with data
//   const colAData = sheet.getRange(15, 1, lastRow - 14, 1).getValues();
//   let actualLastRow = 14;
//   for (let i = 0; i < colAData.length; i++) {
//     if (colAData[i][0] !== '') actualLastRow = 15 + i;
//   }

//   if (actualLastRow < 15) return;

//   // ✅ Sort ENTIRE ROW including all steps 2-29
//   sheet.getRange(15, 1, actualLastRow - 14, sheet.getLastColumn())
//        .sort({column: 1, ascending: false});

//   Logger.log('Full row sort done to row: ' + actualLastRow);
// }

// function colorAllRowsFast(bpmsSheet) {
//   const sheet = bpmsSheet ||
//     SpreadsheetApp.getActiveSpreadsheet().getSheetByName('BPMS VRV');

//   const lastRow = sheet.getLastRow();
//   if (lastRow < 15) return;

//   const totalRows = lastRow - 14;
//   const lastCol = sheet.getLastColumn();

//   // Status column numbers (1-based) - verified from your sheet
//   const cols = {
//     s2:  13,  // M
//     s3:  16,  // P
//     s6:  29,  // AC
//     s7:  36,  // AJ
//     s8:  40,  // AN
//     s9:  44,  // AR
//     s13: 61,  // BI
//     s14: 66,  // BN
//     s15: 70,  // BR
//     s16: 76,  // BX
//     s27: 120, // DT
//     s28: 130, // DZ
//     s29: 135, // EE
//   };

//   // Batch read each status column
//   function getColValues(colNum) {
//     return sheet.getRange(15, colNum, totalRows, 1).getValues()
//       .map(r => (r[0] || '').toString().toLowerCase().trim());
//   }

//   const s2  = getColValues(cols.s2);
//   const s3  = getColValues(cols.s3);
//   const s6  = getColValues(cols.s6);
//   const s7  = getColValues(cols.s7);
//   const s8  = getColValues(cols.s8);
//   const s9  = getColValues(cols.s9);
//   const s13 = getColValues(cols.s13);
//   const s14 = getColValues(cols.s14);
//   const s15 = getColValues(cols.s15);
//   const s16 = getColValues(cols.s16);
//   const s27 = getColValues(cols.s27);
//   const s28 = getColValues(cols.s28);
//   const s29 = getColValues(cols.s29);
//   const tsCol = sheet.getRange(15, 1, totalRows, 1).getValues().map(r => r[0]);

//   const COLORS = {
//     GREEN:      '#C6EFCE',
//     BRIGHT_RED: '#FF0000',
//     PINK:       '#E6B8C2',
//     BLUE:       '#BDD7EE',
//     CREAM:      '#FFF2CC',
//     BROWN:      '#C4A882',
//     ORANGE:     '#F4B942',
//     WHITE:      '#FFFFFF',
//   };

//   const done    = s => s === 'done';
//   const inProc  = s => s === 'in progress';
//   const pending = s => s === '' || s === 'in progress';

//   // Apply color row by row
//   for (let i = 0; i < totalRows; i++) {
//     if (!tsCol[i]) continue; // skip empty rows

//     let color = COLORS.WHITE;

//     if (done(s28[i]) && done(s29[i])) {
//       color = COLORS.GREEN;
//     } else if (
//       (done(s27[i]) && done(s28[i]) && inProc(s29[i])) ||
//       (done(s27[i]) && pending(s28[i]) && inProc(s29[i]))
//     ) {
//       color = COLORS.BLUE;
//     } else if (done(s27[i]) && pending(s28[i]) && done(s29[i])) {
//       color = COLORS.BRIGHT_RED;
//     } else if (done(s27[i]) && s29[i] === '') {
//       color = COLORS.PINK;
//     } else if (done(s13[i]) && pending(s14[i])) {
//       color = COLORS.ORANGE;
//     } else if (done(s6[i]) && (inProc(s7[i]) || inProc(s8[i]) || inProc(s9[i]))) {
//       color = COLORS.CREAM;
//     } else if (done(s2[i]) && inProc(s3[i])) {
//       color = COLORS.BROWN;
//     }

//     sheet.getRange(15 + i, 1, 1, lastCol).setBackground(color);
//   }

//   Logger.log('Coloring done for ' + totalRows + ' rows');
// }

// // function copyOrderToBPMS(ordersRow) {
// //   const ss = SpreadsheetApp.getActiveSpreadsheet();
// //   const ordersSheet = ss.getSheetByName('Orders');
// //   const bpmsSheet = ss.getSheetByName('BPMS VRV');

// //   if (!ordersSheet || !bpmsSheet) {
// //     Logger.log('Sheet not found!');
// //     return;
// //   }

// //   if (!ordersRow || ordersRow === null) {
// //     ordersRow = ordersSheet.getLastRow();
// //   }

// //   const orderData = ordersSheet.getRange(ordersRow, 1, 1, 6).getValues()[0];
// //   const newOrderId = orderData[1].toString().trim();

// //   // Check duplicate
// //   const bpmsLastRow = bpmsSheet.getLastRow();
// //   if (bpmsLastRow >= 15) {
// //     const existingIds = bpmsSheet.getRange(15, 2, bpmsLastRow - 14, 1)
// //                                  .getValues()
// //                                  .map(r => r[0].toString().trim());
// //     if (existingIds.includes(newOrderId)) {
// //       Logger.log('Duplicate order, skipping: ' + newOrderId);
// //       return;
// //     }
// //   }

// //   // Write to next empty row
// //   const newRow = Math.max(bpmsLastRow + 1, 15);
// //   bpmsSheet.getRange(newRow, 1, 1, 6).setValues([orderData]);
// //   bpmsSheet.getRange(newRow, 1).setNumberFormat('dd/MM/yyyy HH:mm:ss');

// //   Logger.log('Added to BPMS VRV row: ' + newRow);

// //   // Sort entire rows + color
// //   sortBPMSByTimestamp(bpmsSheet);
// //   colorAllRowsFast(bpmsSheet);

// //   ss.toast('New order added, sorted & colored!', 'BPMS', 4);
// // }


// // 26-3

// // function copyOrderToBPMS(ordersRow) {
// //   const ss = SpreadsheetApp.getActiveSpreadsheet();
// //   const ordersSheet = ss.getSheetByName('Orders');
// //   const bpmsSheet = ss.getSheetByName('BPMS VRV');

// //   if (!ordersSheet || !bpmsSheet) {
// //     Logger.log('Sheet not found!');
// //     return;
// //   }

// //   if (!ordersRow || ordersRow === null) {
// //     ordersRow = ordersSheet.getLastRow();
// //   }

// //   const orderData = ordersSheet.getRange(ordersRow, 1, 1, 6).getValues()[0];
// //   const newOrderId = orderData[1].toString().trim();

// //   // Check duplicate
// //   const bpmsLastRow = bpmsSheet.getLastRow();
// //   if (bpmsLastRow >= 15) {
// //     const existingIds = bpmsSheet.getRange(15, 2, bpmsLastRow - 14, 1)
// //                                  .getValues()
// //                                  .map(r => r[0].toString().trim());
// //     if (existingIds.includes(newOrderId)) {
// //       Logger.log('Duplicate order, skipping: ' + newOrderId);
// //       return;
// //     }
// //   }

// //   const newRow = Math.max(bpmsLastRow + 1, 15);

// //   // ✅ STEP 1 - Copy entire previous row's formulas first
// //   // This copies all Planned/Actual/TimeDelay formulas from last data row
// //   const templateRow = bpmsLastRow; // last existing row has all formulas
// //   if (templateRow >= 15) {
// //     // Copy ALL columns from template row to new row
// //     bpmsSheet.getRange(templateRow, 1, 1, bpmsSheet.getLastColumn())
// //              .copyTo(
// //                bpmsSheet.getRange(newRow, 1, 1, bpmsSheet.getLastColumn()),
// //                SpreadsheetApp.CopyPasteType.PASTE_FORMULA,
// //                false
// //              );
// //   }

// //   // ✅ STEP 2 - Now overwrite A-F with actual order data
// //   bpmsSheet.getRange(newRow, 1, 1, 6).setValues([orderData]);
// //   bpmsSheet.getRange(newRow, 1).setNumberFormat('dd/MM/yyyy HH:mm:ss');

// //   // ✅ STEP 3 - Clear status columns G onwards (except formulas)
// //   // Clear Status dropdowns so new row starts fresh
// //   // Status cols: J, M, P, S, V, AC, AJ, AN, AR, AV, BA, BE, BI, BN, BR, BX, CB, DT, DZ, EE
// //   const statusCols = [10,13,16,19,22,29,36,40,44,48,53,57,61,66,70,76,90,120,130,135];
// //   statusCols.forEach(col => {
// //     bpmsSheet.getRange(newRow, col).clearContent();
// //   });

// //   Logger.log('Added to BPMS VRV row: ' + newRow);

// //   // ✅ STEP 4 - Sort entire row + color
// //   sortBPMSByTimestamp(bpmsSheet);
// //   colorAllRowsFast(bpmsSheet);

// //   ss.toast('New order added, sorted & colored!', 'BPMS', 4);
// // }

// // function copyOrderToBPMS(ordersRow) {
// //   const ss = SpreadsheetApp.getActiveSpreadsheet();
// //   const ordersSheet = ss.getSheetByName('Orders');
// //   const bpmsSheet = ss.getSheetByName('BPMS VRV');
// //   const templateSheet = ss.getSheetByName('BPMS-Template');

// //   if (!ordersSheet || !bpmsSheet) {
// //     Logger.log('Sheet not found!');
// //     return;
// //   }

// //   if (!templateSheet) {
// //     ss.toast('BPMS-Template sheet not found! Please create it first.', 'Error', 5);
// //     Logger.log('BPMS-Template sheet missing!');
// //     return;
// //   }

// //   if (!ordersRow || ordersRow === null) {
// //     ordersRow = ordersSheet.getLastRow();
// //   }

// //   const orderData = ordersSheet.getRange(ordersRow, 1, 1, 6).getValues()[0];
// //   const newOrderId = orderData[1].toString().trim();

// //   // Check duplicate
// //   const bpmsLastRow = bpmsSheet.getLastRow();
// //   if (bpmsLastRow >= 15) {
// //     const existingIds = bpmsSheet.getRange(15, 2, bpmsLastRow - 14, 1)
// //                                  .getValues()
// //                                  .map(r => r[0].toString().trim());
// //     if (existingIds.includes(newOrderId)) {
// //       Logger.log('Duplicate order, skipping: ' + newOrderId);
// //       return;
// //     }
// //   }

// //   const newRow = Math.max(bpmsLastRow + 1, 15);
// //   const lastCol = bpmsSheet.getLastColumn();

// //   // ✅ STEP 1 - Copy ENTIRE template row to new row
// //   // This copies formulas, validations, dropdowns, upload buttons
// //   templateSheet.getRange(2, 1, 1, lastCol)
// //     .copyTo(
// //       bpmsSheet.getRange(newRow, 1, 1, lastCol),
// //       SpreadsheetApp.CopyPasteType.PASTE_NORMAL,
// //       false
// //     );

// //   Logger.log('Template copied to row: ' + newRow);

// //   // ✅ STEP 2 - Clear all data columns so new row starts fresh
// //   // Clear: Actual timestamps, Status dropdowns, Time Delay, Remarks
// //   // Keep: Planned formulas (they auto-calculate from Col A timestamp)

// //   // All Actual columns (every 4th column pattern after each step)
// //   // Based on your sheet: Planned=H, Actual=I, Status=J, TimeDelay=K pattern
// //   // Clear entire row data first except formulas
// //   const clearRanges = [];

// //   // Clear Actual columns (col I=9 and every step's actual col)
// //   // These are the timestamp columns that get set when status = Done
// //   const actualCols = [9,  // Step 2 Actual
// //     calculateActualCols(lastCol) // we'll get all actual cols
// //   ];

// //   // Simpler approach - clear all values but keep formulas
// //   const allFormulas = bpmsSheet.getRange(newRow, 1, 1, lastCol).getFormulas()[0];
// //   const clearValues = new Array(lastCol).fill('');

// //   // Write blanks to entire row first
// //   bpmsSheet.getRange(newRow, 1, 1, lastCol).setValues([clearValues]);

// //   // Then restore formulas
// //   for (let col = 1; col <= lastCol; col++) {
// //     if (allFormulas[col - 1] !== '') {
// //       bpmsSheet.getRange(newRow, col).setFormula(allFormulas[col - 1]);
// //     }
// //   }

// //   Logger.log('Row cleared, formulas restored');

// //   // ✅ STEP 3 - Write actual order data to A-F
// //   bpmsSheet.getRange(newRow, 1, 1, 6).setValues([orderData]);
// //   bpmsSheet.getRange(newRow, 1).setNumberFormat('dd/MM/yyyy HH:mm:ss');

// //   Logger.log('Order data written: ' + newOrderId);

// //   // ✅ STEP 4 - Restore data validations (dropdowns) from template
// //   const templateValidations = templateSheet.getRange(2, 1, 1, lastCol).getDataValidations()[0];
// //   for (let col = 1; col <= lastCol; col++) {
// //     if (templateValidations[col - 1] !== null) {
// //       bpmsSheet.getRange(newRow, col).setDataValidation(templateValidations[col - 1]);
// //     }
// //   }

// //   Logger.log('Validations restored');

// //   // ✅ STEP 5 - Sort + Color
// //   sortBPMSByTimestamp(bpmsSheet);
// //   colorAllRowsFast(bpmsSheet);

// //   ss.toast('New order added with all steps & sorted!', 'BPMS', 4);
// // }

// function copyOrderToBPMS(ordersRow) {
//   const ss = SpreadsheetApp.getActiveSpreadsheet();
//   const ordersSheet = ss.getSheetByName('Orders');
//   const bpmsSheet = ss.getSheetByName('BPMS VRV');
//   const templateSheet = ss.getSheetByName('BPMS-Template');

//   if (!ordersSheet || !bpmsSheet || !templateSheet) {
//     ss.toast('Sheet not found!', 'Error', 5);
//     return;
//   }

//   if (!ordersRow) ordersRow = ordersSheet.getLastRow();

//   const orderData = ordersSheet.getRange(ordersRow, 1, 1, 6).getValues()[0];
//   const newOrderId = orderData[1].toString().trim();

//   // Duplicate check
//   const bpmsLastRow = bpmsSheet.getLastRow();
//   if (bpmsLastRow >= 15) {
//     const existingIds = bpmsSheet.getRange(15, 2, bpmsLastRow - 14, 1)
//                                  .getValues()
//                                  .map(r => r[0].toString().trim());
//     if (existingIds.includes(newOrderId)) {
//       ss.toast('Order already exists!', 'BPMS', 3);
//       return;
//     }
//   }

//   const newRow = Math.max(bpmsLastRow + 1, 15);
//   const lastCol = bpmsSheet.getLastColumn();

//   // STEP 1 — Copy template
//   templateSheet.getRange(2, 1, 1, lastCol)
//     .copyTo(
//       bpmsSheet.getRange(newRow, 1, 1, lastCol),
//       SpreadsheetApp.CopyPasteType.PASTE_NORMAL,
//       false
//     );

//   // ✅ STEP 2 — Fix corrupted TAT formula refs IMMEDIATELY after copy
//   fixRowFormulas(bpmsSheet, newRow);

//   // STEP 3 — Write order data into A-F
//   bpmsSheet.getRange(newRow, 1, 1, 6).setValues([orderData]);
//   bpmsSheet.getRange(newRow, 1).setNumberFormat('dd/MM/yyyy HH:mm:ss');

//   // STEP 4 — Restore dropdowns from template
//   const validations = templateSheet.getRange(2, 1, 1, lastCol).getDataValidations()[0];
//   for (let col = 1; col <= lastCol; col++) {
//     if (validations[col - 1]) {
//       bpmsSheet.getRange(newRow, col).setDataValidation(validations[col - 1]);
//     }
//   }

//   // STEP 5 — Sort + Color
//   sortBPMSByTimestamp(bpmsSheet);
//   colorAllRowsFast(bpmsSheet);

//   ss.toast('Order added!', 'BPMS', 4);
//   Logger.log('copyOrderToBPMS done: ' + newOrderId + ' at row ' + newRow);
// }




// // Helper placeholder - not needed with new approach
// function calculateActualCols(lastCol) {
//   return [];
// }

// function fixAllTemplateFormulas() {
//   const ss = SpreadsheetApp.getActiveSpreadsheet();
//   const templateSheet = ss.getSheetByName('BPMS-Template');
//   const bpmsSheet = ss.getSheetByName('BPMS VRV');
//   const lastCol = templateSheet.getLastColumn();

//   Logger.log('Starting formula fix...');
//   let fixedCount = 0;
//   let clearedCount = 0;

//   for (let col = 1; col <= lastCol; col++) {
//     const formula = templateSheet.getRange(2, col).getFormula();
//     if (formula === '') continue;

//     // Skip #REF! formulas - clear them
//     if (formula.includes('#REF!')) {
//       templateSheet.getRange(2, col).clearContent();
//       clearedCount++;
//       Logger.log('Cleared #REF! at col ' + col);
//       continue;
//     }

//     // Fix wrong row references - all off by 1
//     let fixed = formula
//       .replace(/\$C\$7/g, '$C$8')   // Opening time
//       .replace(/\$D\$7/g, '$D$8')   // Closing time
//       .replace(/\$E\$7/g, '$E$8')   // Working days
//       .replace(/\$A\$7/g, '$A$8')   // NOW() cell
//       .replace(/H\$12/g,  'H$13')   // Step 2 TAT
//       .replace(/P\$12/g,  'P$13')   // Step 3 TAT
//       .replace(/T\$12/g,  'T$13')   // Step 4 TAT
//       .replace(/X\$12/g,  'X$13')   // Step 5 TAT
//       .replace(/AB\$12/g, 'AB$13')  // Step 6 TAT
//       .replace(/AI\$12/g, 'AI$13')  // Step 7 TAT
//       .replace(/AM\$12/g, 'AM$13')  // Step 8 TAT
//       .replace(/AQ\$12/g, 'AQ$13')  // Step 9 TAT
//       .replace(/AU\$12/g, 'AU$13')  // Step 10 TAT
//       .replace(/AY\$12/g, 'AY$13')  // Step 11 TAT
//       .replace(/BC\$12/g, 'BC$13')  // Step 12 TAT
//       .replace(/BG\$12/g, 'BG$13')  // Step 13 TAT
//       .replace(/BL\$12/g, 'BL$13')  // Step 14 TAT
//       .replace(/BP\$12/g, 'BP$13')  // Step 15 TAT
//       .replace(/BV\$12/g, 'BV$13')  // Step 16 TAT
//       .replace(/BZ\$12/g, 'BZ$13')  // Step 17 TAT
//       .replace(/CF\$12/g, 'CF$13')  // Step 18 TAT
//       .replace(/CJ\$12/g, 'CJ$13')  // Step 19 TAT
//       .replace(/CN\$12/g, 'CN$13')  // Step 20 TAT
//       .replace(/CR\$12/g, 'CR$13')  // Step 21 TAT
//       .replace(/CV\$12/g, 'CV$13')  // Step 22 TAT
//       .replace(/CZ\$12/g, 'CZ$13')  // Step 23 TAT
//       .replace(/DD\$12/g, 'DD$13')  // Step 24 TAT
//       .replace(/DH\$12/g, 'DH$13')  // Step 25 TAT
//       .replace(/DL\$12/g, 'DL$13')  // Step 26 TAT
//       .replace(/DP\$12/g, 'DP$13')  // Step 27 TAT
//       .replace(/DT\$12/g, 'DT$13')  // Step 28 TAT (if exists)
//       .replace(/DX\$12/g, 'DX$13')  // Step 29 TAT (if exists)
//       // Generic catch-all for any remaining $12 TAT references
//       .replace(/([A-Z]{1,2})\$12/g, '$1$13');

//     if (fixed !== formula) {
//       templateSheet.getRange(2, col).setFormula(fixed);
//       fixedCount++;
//       Logger.log('Fixed col ' + col + ': ' + fixed.substring(0, 60));
//     }
//   }

//   Logger.log('Done! Fixed: ' + fixedCount + ' | Cleared: ' + clearedCount);

//   // Now also fix row 15 in BPMS VRV (current broken row)
//   fixRow(bpmsSheet, 15, lastCol);

//   ss.toast('Template fixed! Fixed: ' + fixedCount + ' formulas', 'BPMS', 5);
// }

// // Fix a specific row in BPMS VRV with same corrections
// function fixRow(sheet, row, lastCol) {
//   Logger.log('Fixing BPMS VRV row ' + row);
//   let fixedCount = 0;

//   for (let col = 1; col <= lastCol; col++) {
//     const formula = sheet.getRange(row, col).getFormula();
//     if (formula === '') continue;

//     if (formula.includes('#REF!')) {
//       sheet.getRange(row, col).clearContent();
//       continue;
//     }

//     const fixed = formula
//       .replace(/\$C\$7/g, '$C$8')
//       .replace(/\$D\$7/g, '$D$8')
//       .replace(/\$E\$7/g, '$E$8')
//       .replace(/\$A\$7/g, '$A$8')
//       .replace(/([A-Z]{1,2})\$12/g, '$1$13');

//     if (fixed !== formula) {
//       sheet.getRange(row, col).setFormula(fixed);
//       fixedCount++;
//     }
//   }

//   Logger.log('Row ' + row + ' fixed: ' + fixedCount + ' formulas');
// }

// // Run this after fixAllTemplateFormulas to fix ALL existing BPMS VRV rows
// function fixAllBPMSRows() {
//   const ss = SpreadsheetApp.getActiveSpreadsheet();
//   const bpmsSheet = ss.getSheetByName('BPMS VRV');
//   const lastRow = bpmsSheet.getLastRow();
//   const lastCol = bpmsSheet.getLastColumn();

//   ss.toast('Fixing all rows...', 'BPMS', 5);

//   for (let row = 15; row <= lastRow; row++) {
//     fixRow(bpmsSheet, row, lastCol);
//   }

//   ss.toast('All rows fixed!', 'BPMS', 3);
//   Logger.log('All BPMS VRV rows fixed');
// }


// function nuclearFixFast() {
//   const ss = SpreadsheetApp.getActiveSpreadsheet();
//   const bpmsSheet = ss.getSheetByName('BPMS VRV');
//   const templateSheet = ss.getSheetByName('BPMS-Template');
//   const lastCol = bpmsSheet.getLastColumn();
//   const lastRow = bpmsSheet.getLastRow();

//   function fixCorruptedFormula(formula) {
//     if (formula === '' || !formula.includes('16')) return formula;
//     return formula
//       .replace(/\bAA16\b/g, 'A$13').replace(/\bBB16\b/g, 'B$13')
//       .replace(/\bCC16\b/g, 'C$13').replace(/\bDD16\b/g, 'D$13')
//       .replace(/\bEE16\b/g, 'E$13').replace(/\bFF16\b/g, 'F$13')
//       .replace(/\bGG16\b/g, 'G$13').replace(/\bHH16\b/g, 'H$13')
//       .replace(/\bII16\b/g, 'I$13').replace(/\bJJ16\b/g, 'J$13')
//       .replace(/\bKK16\b/g, 'K$13').replace(/\bLL16\b/g, 'L$13')
//       .replace(/\bMM16\b/g, 'M$13').replace(/\bNN16\b/g, 'N$13')
//       .replace(/\bOO16\b/g, 'O$13').replace(/\bPP16\b/g, 'P$13')
//       .replace(/\bQQ16\b/g, 'Q$13').replace(/\bRR16\b/g, 'R$13')
//       .replace(/\bSS16\b/g, 'S$13').replace(/\bTT16\b/g, 'T$13')
//       .replace(/\bUU16\b/g, 'U$13').replace(/\bVV16\b/g, 'V$13')
//       .replace(/\bWW16\b/g, 'W$13').replace(/\bXX16\b/g, 'X$13')
//       .replace(/\bYY16\b/g, 'Y$13').replace(/\bZZ16\b/g, 'Z$13');
//   }

//   // ✅ Fix template - batch read/write
//   ss.toast('Fixing template...', 'BPMS', 3);
//   const tFormulas = templateSheet.getRange(2, 1, 1, lastCol).getFormulas()[0];
//   const tFixed = tFormulas.map(f => fixCorruptedFormula(f));
//   // Write all at once
//   for (let col = 0; col < lastCol; col++) {
//     if (tFixed[col] !== tFormulas[col] && tFixed[col] !== '') {
//       templateSheet.getRange(2, col + 1).setFormula(tFixed[col]);
//     }
//   }
//   Logger.log('Template fixed');

//   // ✅ Fix BPMS VRV - batch read entire sheet at once
//   ss.toast('Fixing BPMS VRV rows...', 'BPMS', 5);
//   const allFormulas = bpmsSheet.getRange(15, 1, lastRow - 14, lastCol).getFormulas();
  
//   // Process all rows
//   for (let r = 0; r < allFormulas.length; r++) {
//     let rowChanged = false;
//     const rowFormulas = allFormulas[r];
    
//     for (let c = 0; c < lastCol; c++) {
//       if (rowFormulas[c] === '') continue;
//       const fixed = fixCorruptedFormula(rowFormulas[c]);
//       if (fixed !== rowFormulas[c]) {
//         rowFormulas[c] = fixed;
//         rowChanged = true;
//       }
//     }
    
//     // Only write back rows that changed
//     if (rowChanged) {
//       // Write only changed cells in this row
//       for (let c = 0; c < lastCol; c++) {
//         const orig = allFormulas[r][c];
//         const fixed = rowFormulas[c];
//         if (fixed !== orig && fixed !== '') {
//           bpmsSheet.getRange(15 + r, c + 1).setFormula(fixed);
//         }
//       }
//       Logger.log('Fixed row: ' + (15 + r));
//     }
//   }

//   Logger.log('Nuclear fix complete!');
//   ss.toast('Done! All formulas fixed.', 'BPMS', 3);
// }

// function rebuildAllFormulas() {
//   const ss = SpreadsheetApp.getActiveSpreadsheet();
//   const bpmsSheet = ss.getSheetByName('BPMS VRV');
//   const templateSheet = ss.getSheetByName('BPMS-Template');
//   const lastRow = bpmsSheet.getLastRow();
//   const lastCol = bpmsSheet.getLastColumn();

//   ss.toast('Reading template formulas...', 'BPMS', 3);

//   // ✅ STEP 1 - Get correct formulas from a KNOWN GOOD ROW
//   // Row 20 or any old complete row that has correct formulas
//   // Find best template row (most formulas, no corruption)
//   let bestRow = -1;
//   let bestCount = 0;
  
//   for (let row = 15; row <= Math.min(lastRow, 30); row++) {
//     const formulas = bpmsSheet.getRange(row, 1, 1, lastCol).getFormulas()[0];
//     const goodCount = formulas.filter(f => 
//       f !== '' && !f.includes('#REF!') && !f.includes('HH') && 
//       !f.includes('PP') && f.includes('$13')
//     ).length;
//     Logger.log('Row ' + row + ' good formulas: ' + goodCount);
//     if (goodCount > bestCount) {
//       bestCount = goodCount;
//       bestRow = row;
//     }
//   }
  
//   Logger.log('Best template row: ' + bestRow + ' with ' + bestCount + ' good formulas');
  
//   if (bestRow === -1) {
//     ss.toast('No good template row found!', 'Error', 5);
//     return;
//   }
  
//   // ✅ STEP 2 - Get formulas from best row
//   const templateFormulas = bpmsSheet.getRange(bestRow, 1, 1, lastCol).getFormulas()[0];
  
//   // ✅ STEP 3 - Fix Holidays reference to use $A$3:$A (static)
//   const cleanFormulas = templateFormulas.map(f => {
//     if (f === '') return f;
//     // Fix dynamic Holidays row reference to static
//     return f.replace(/Holidays!\$A\d+:\$A/g, 'Holidays!$A$3:$A');
//   });
  
//   Logger.log('Template formulas cleaned');

//   // ✅ STEP 4 - Update template sheet with clean formulas
//   for (let col = 0; col < lastCol; col++) {
//     if (cleanFormulas[col] !== '') {
//       // Adjust formula for template row 2
//       const adjustedFormula = adjustFormulaToRow(cleanFormulas[col], bestRow, 2);
//       if (adjustedFormula) {
//         templateSheet.getRange(2, col + 1).setFormula(adjustedFormula);
//       }
//     }
//   }
//   Logger.log('Template updated');

//   // ✅ STEP 5 - Apply correct formulas to ALL data rows
//   ss.toast('Fixing all rows...', 'BPMS', 10);
  
//   for (let row = 15; row <= lastRow; row++) {
//     // Skip rows without timestamp (empty rows)
//     const tsVal = bpmsSheet.getRange(row, 1).getValue();
//     if (tsVal === '') continue;
    
//     // Apply formulas adjusted for this row
//     for (let col = 0; col < lastCol; col++) {
//       if (cleanFormulas[col] === '') continue;
      
//       const adjustedFormula = adjustFormulaToRow(cleanFormulas[col], bestRow, row);
//       if (adjustedFormula) {
//         bpmsSheet.getRange(row, col + 1).setFormula(adjustedFormula);
//       }
//     }
//     Logger.log('Row ' + row + ' done');
//   }

//   ss.toast('All rows rebuilt! Running sort & color...', 'BPMS', 3);
//   sortBPMSByTimestamp(bpmsSheet);
//   colorAllRowsFast(bpmsSheet);
//   ss.toast('Complete!', 'BPMS', 3);
// }

// // Adjust formula row references from sourceRow to targetRow
// function adjustFormulaToRow(formula, sourceRow, targetRow) {
//   if (!formula || formula === '') return '';
//   if (formula.includes('#REF!')) return '';
  
//   // Replace row numbers in formula
//   // Match cell references like A15, H15, I15 etc (not fixed refs like $C$8)
//   // Only replace non-fixed row references
//   const result = formula.replace(
//     /([A-Z]{1,2})(\d+)/g,
//     (match, col, row) => {
//       const rowNum = parseInt(row);
//       // Skip fixed references (preceded by $) and header rows (1-14)
//       if (rowNum <= 14) return match; // keep header refs unchanged
//       if (rowNum === sourceRow) {
//         return col + targetRow; // adjust to target row
//       }
//       return match;
//     }
//   );
  
//   return result;
// }

// // ============================================================
// // Call this inside copyOrderToBPMS after the copyTo call
// // ============================================================
// function fixRowFormulas(sheet, row) {
//   const lastCol = sheet.getLastColumn();

//   // Read all formulas for this row in one batch call
//   const formulas = sheet.getRange(row, 1, 1, lastCol).getFormulas()[0];

//   // Map of corrupted TAT ref → correct ref
//   // Pattern: doubled col letter + wrong row  e.g. PP17 → P$13
//   // Covers all steps in your sheet
//   const tatFixes = [
//     ['PP', 'P'], ['TT', 'T'], ['XX', 'X'], ['HH', 'H'],
//     ['AB', 'AB'], // AB is already 2 chars, won't double — skip
//     ['AIAI', 'AI'], ['AMAM', 'AM'], ['AQAQ', 'AQ'], ['AUAU', 'AU'],
//     ['AYAY', 'AY'], ['BCBC', 'BC'], ['BGBG', 'BG'], ['BLBL', 'BL'],
//     ['BPBP', 'BP'], ['BVBV', 'BV'], ['BZBZ', 'BZ'], ['CFCF', 'CF'],
//     ['CJCJ', 'CJ'], ['CNCN', 'CN'], ['CRCR', 'CR'], ['CVCV', 'CV'],
//     ['CZCZ', 'CZ'], ['DDDD', 'DD'], ['DHDH', 'DH'], ['DLDL', 'DL'],
//     ['DPDP', 'DP'], ['DTDT', 'DT'], ['DXDX', 'DX'], ['EDED', 'ED'],
//   ];

//   let changed = false;
//   const fixed = formulas.map(f => {
//     if (!f) return f;
//     let result = f;

//     // Fix doubled column refs with wrong row number
//     // e.g. PP17 → P$13, AIAI17 → AI$13
//     for (const [bad, good] of tatFixes) {
//       // Match the doubled col + any row number (not already fixed with $)
//       const re = new RegExp(bad + '(\\d+)', 'g');
//       result = result.replace(re, good + '$13');
//     }

//     // Generic catch-all: any remaining doubled single-letter col + row number
//     // e.g. HH17, SS17, VV17 etc
//     result = result.replace(/\b([A-Z])\1(\d+)\b/g, '$1$$$13');

//     if (result !== f) changed = true;
//     return result;
//   });

//   if (changed) {
//     // Write all fixed formulas back in one batch call
//     sheet.getRange(row, 1, 1, lastCol).setFormulas([fixed]);
//     Logger.log('fixRowFormulas: fixed row ' + row);
//   }
// }