function serviceAriel() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setActiveRange(ss.getRange("A4"));
}
function serviceMatt() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var row = ss.getActiveCell().getRow();
  if(row<188){ss.setActiveRange(ss.getRange("A218"));ss.getActiveRange().getValue();}
  ss.setActiveRange(ss.getRange("A188"));
}
function serviceDave() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var row = ss.getActiveCell().getRow();
  if(row<344){ss.setActiveRange(ss.getRange("A374"));ss.getActiveRange().getValue();}
  ss.setActiveRange(ss.getRange("A344"));
}