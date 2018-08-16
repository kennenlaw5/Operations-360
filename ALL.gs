function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Utilities').addSubMenu(ui.createMenu('Help').addItem('By Phone','menuItem1').addItem('By Email','menuItem2')).addItem('GSM Report Import','fullImport').addToUi();
  //.addItem('Reset Statistics','reset').addItem('Refresh CA Ranking','rank').addToUi();
  var message = 'The spreadsheet has loaded successfully! Have a great day!';
  var title = 'Complete!';
  SpreadsheetApp.getActiveSpreadsheet().toast(message, title);
}
function menuItem1() {
  SpreadsheetApp.getUi().alert('Call or text (720) 317-5427');
}
function menuItem2() {
  var ui = SpreadsheetApp.getUi();
  var input = ui.prompt('Send Email:','Describe the issue you\'re having in the box below, then press "Ok" to submit your issue via email:',ui.ButtonSet.OK_CANCEL);
  if (input.getSelectedButton() == ui.Button.OK) {
    MailApp.sendEmail('kennen.lawrence@schomp.com','HELP Sales Daily_January',input.getResponseText(),{name:getName()});
    SpreadsheetApp.getActiveSpreadsheet().toast('Email sent successfully! Your issue should be addressed within a few hours! For more immediate assistance, contact Kennen by phone.','Email Sent');
  } else if (input.getSelectedButton() == ui.Button.CANCEL) {
    Logger.log('User cancelled');
  }
}
function getName(){
  //Created By Kennen Lawrence
  //Version 1.0
  var email = Session.getActiveUser().getEmail();
  var name;var first;var last;
  name = email.split("@schomp.com");
  name=name[0];
  name=name.split(".");
  first=name[0];
  last=name[1];
  first= first[0].toUpperCase() + first.substring(1);
  last=last[0].toUpperCase() + last.substring(1);
  name=first+" "+last;
  Logger.log(name);
  return name;
}