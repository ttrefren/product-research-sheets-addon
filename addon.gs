var sheet = SpreadsheetApp.getActiveSheet();
var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
 
function getEmails(label_name) {
  var label = GmailApp.getUserLabelByName(label_name);
  var threads = label.getThreads();
  
  sheet.getRange(1, 1).setValue("Label: " + label_name);
  sheet.getRange(2, 1, 1, 4).setValues([["Date", "Email address", "Message", "Feedback bucket"]]);
  for (var i = 0; i < threads.length; i++) {
    var messages=threads[i].getMessages();
    // use first message for user email address, last message to get entire body
    
    var first_message = messages[0];
    var last_message = messages[messages.length - 1];

    sheet.getRange(i + 3, 1, 1, 3).setValues([[first_message.getDate(), first_message.getFrom(), last_message.getPlainBody()]]);
  }
}

function loadEmailsFromLabel() {
  var ui = SpreadsheetApp.getUi();
  var label = sheet.getRange(1, 1).getValue().split('Label: ')[1];
  if (!label) {
    label = ui.prompt("label for feedback emails to load?").getResponseText();
  }
  getEmails(label);
}  
 
function onOpen() {
  var menuEntries = [ {name: "Load feedback emails from label", functionName: "loadEmailsFromLabel"} ];
  spreadsheet.addMenu("Load emails", menuEntries);
}
