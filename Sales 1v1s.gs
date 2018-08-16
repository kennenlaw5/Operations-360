function viewJeff() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var loc=7;
  ss.setActiveRange(ss.getRange("A"+loc));
}
function viewAnna() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var loc=166;
  var row = ss.getActiveCell().getRow();
  if(row<loc){ss.setActiveRange(ss.getRange("B"+(loc+25)));ss.getActiveRange().getValue();}
  ss.setActiveRange(ss.getRange("A"+loc));
}
function viewBen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var loc=356;
  var row = ss.getActiveCell().getRow();
  if(row<loc){ss.setActiveRange(ss.getRange("B"+(loc+25)));ss.getActiveRange().getValue();}
  ss.setActiveRange(ss.getRange("A"+loc));
}
function viewMark() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var loc=546;
  var row = ss.getActiveCell().getRow();
  if(row<loc){ss.setActiveRange(ss.getRange("B"+(loc+25)));ss.getActiveRange().getValue();}
  ss.setActiveRange(ss.getRange("A"+loc));
}
function viewRobb() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var loc=705;
  var row = ss.getActiveCell().getRow();
  if(row<loc){ss.setActiveRange(ss.getRange("B"+(loc+25)));ss.getActiveRange().getValue();}
  ss.setActiveRange(ss.getRange("A"+loc));
}
function viewSeth() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var loc=895;
  var row = ss.getActiveCell().getRow();
  if(row<loc){ss.setActiveRange(ss.getRange("B"+(loc+24)));ss.getActiveRange().getValue();}
  ss.setActiveRange(ss.getRange("A"+loc));
}
function viewDean() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var loc=1085;
  var row = ss.getActiveCell().getRow();
  if(row<loc){ss.setActiveRange(ss.getRange("B"+(loc+25)));ss.getActiveRange().getValue();}
  ss.setActiveRange(ss.getRange("A"+loc));
}