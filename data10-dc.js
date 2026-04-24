// function handleDCFormSubmit(e) {
//   const ss = SpreadsheetApp.getActiveSpreadsheet();
  
//   const source = ss.getSheetByName("Form Responses 10");
//   const target = ss.getSheetByName("BPMS-Upload DC");

//   if (!source || !target) {
//     Logger.log("Sheet not found! Check names.");
//     return;
//   }

//   // Copy response to target sheet with column mapping
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
  
//   Logger.log("DC form response routed successfully!");
// }

function handleDCFormSubmit(e) {

    const submittedSheet = e.range.getSheet().getName();
  if (submittedSheet !== "Form Responses 10") return; // ignore other forms
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  const source = ss.getSheetByName("Form Responses 10");
  const target = ss.getSheetByName("BPMS-Upload DC");

  if (!source || !target) {
    Logger.log("Sheet not found!");
    return;
  }

  const sourceHeaders = source.getRange(1, 1, 1, source.getLastColumn()).getValues()[0]
    .map(h => h.toString().trim()); // ← trims all spaces
  const targetHeaders = target.getRange(1, 1, 1, target.getLastColumn()).getValues()[0]
    .map(h => h.toString().trim()); // ← trims all spaces

  const lastSourceRow = source.getLastRow();
  const sourceData = source.getRange(lastSourceRow, 1, 1, source.getLastColumn()).getValues()[0];

  const newRow = [];
  for (var i = 0; i < targetHeaders.length; i++) {
    const sourceIndex = sourceHeaders.indexOf(targetHeaders[i]);
    newRow.push(sourceIndex !== -1 ? sourceData[sourceIndex] : "");
  }
  target.appendRow(newRow);
  Logger.log("Row appended successfully!");
}
