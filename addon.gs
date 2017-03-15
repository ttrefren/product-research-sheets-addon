
var sheet = SpreadsheetApp.getActiveSheet();
var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

function getEmails(label_name) {
  var label = GmailApp.getUserLabelByName(label_name);
  if (!label) {
    throw new Error("No label found with that name. (If your label name has spaces, make sure they are spaces and not dashes)");
  }

  var threads = label.getThreads();

  sheet.getRange(1, 1).setValue("Label: " + label_name);
  sheet.getRange(2, 1, 1, 5).setValues([["Date", "Thread ID", "Email address", "Message", "Feedback bucket"]]);

  var last_row = sheet.getLastRow();
  var previous_email_ids = [];
  if (last_row > 2) {
    Logger.log("Existing data in sheet - grabbing previous thread IDs");
    var previous_email_ids = sheet.getRange(3, 2, last_row - 2).getValues().map(function(item) { return item[0] });
  }

  var insert_row = last_row + 1;
  for (var i = 0; i < threads.length; i++) {
    var messages=threads[i].getMessages();
    // use first message for user email address, last message to get entire body

    var first_message = messages[0];
    var last_message = messages[messages.length - 1];

    var date = first_message.getDate(),
        thread_id = first_message.getId(),
        from_addr = first_message.getFrom(),
        body = last_message.getPlainBody();

    if (previous_email_ids.indexOf(thread_id) == -1) {
      sheet.getRange(insert_row, 1, 1, 4).setValues([[date, thread_id, from_addr, body]]);
      insert_row++;
      Logger.log("Adding row " + thread_id);
    } else {
      Logger.log("Skipping " + thread_id);
    }
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
  SpreadsheetApp.getUi().createAddonMenu().addItem('Load feedback emails from label', 'loadEmailsFromLabel').addToUi();
}

function onInstall() {
  onOpen();
}
