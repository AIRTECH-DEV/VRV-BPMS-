function showChangeNew() {
  // Just go to Apps Script and find: function onChange_new
  // Paste it here

  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var sheet=ss.getActiveSheet();
  var active=ss.getActiveCell();
  var row=active.getRow();
  var col=active.getColumn();
  var val=active.getValue();
  if(val=="Done" ){var ts=sheet.getRange(row, col-1).getValue(); if(ts==""){sheet.getRange(row, col-1).setValue(new Date())}}
  if(val=="Yes"){var ts2=sheet.getRange(row, col-1).getValue(); if(ts2==""){sheet.getRange(row, col-1).setValue(new Date())}}
  if(val=="No" ){var ts3=sheet.getRange(row, col-1).getValue(); if(ts3==""){sheet.getRange(row, col-1).setValue(new Date())}}
  if(val===true ){var ts4=sheet.getRange(row, col-1).getValue(); if(ts4==""){sheet.getRange(row, col-1).setValue(new Date())}}
}