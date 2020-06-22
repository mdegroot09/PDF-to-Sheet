var importReport = () => {

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sheet1');
  var email = sheet.getRange('B2').getValue();
  var threads = GmailApp.search('in:inbox from:"' + email + '"');
  var messages = threads[0].getMessages();
  messages.reverse();
  sheet.getRange('D1').setValue(messages[0].getDate())
//  var attachment = messages[0].getAttachments()[0].setContentType('text/csv').getDataAsString();
  var attachment = messages[0].getAttachments()[0].getDataAsString()
//  var csvData = Utilities.parseCsv(attachment.getDataAsString(), ",")
//  sheet.getRange('A4').setValue(csvData)
//  sheet.getRange('A4').setValue(ContentService.createTextOutput(attachment).getContent())
  sheet.getRange('A4').setValue(attachment);
  
//  attachment.setContentType('text/csv');
//
//  // Is the attachment a CSV file
//  if (attachment.getContentType() === "text/csv") {
//
//    var csvData = Utilities.parseCsv(attachment.getDataAsString(), ",");
//  
//    // Remember to clear the content of the sheet before importing new data
//    // sheet.clearContents().clearFormats();
//    sheet.getRange(4, 1, csvData.length, csvData[0].length).setValues(csvData);
//  
//    // GmailApp.moveMessageToTrash(messages);
//    // GmailApp.moveThreadsToArchive(messages[0])
//  }
}