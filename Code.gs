function importReport(){

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('sheet1');
  var threads = GmailApp.search('in:"Pricing for Casey"');
  var messages = threads[0].getMessages();
  messages.reverse();
  sheet.getRange('D1').setValue(messages[0].getDate())
  sheet.getRange('D2').setValue(new Date())
//  var attachment = messages[0].getAttachments()[0].setContentType('text/csv').getDataAsString();
  var attachment = messages[0].getAttachments()[0].setContentType(MimeType.FOLDER).getDataAsString()
  return sheet.getRange('A4').setValue(attachment);
  
  var csvData = Utilities.parseCsv(attachment.getDataAsString(), " ")
  sheet.clearContents().clearFormats();
  sheet.getRange(1, 1, csvData.length, csvData[0].length).setValues(csvData)
}

function extractTextFromPDF() {

  // PDF File URL
  // You can also pull PDFs from Google Drive
  // var blob = UrlFetchApp.fetch('url').getBlob();
  
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('sheet1');
  
  // var threads = GmailApp.search('in:inbox from:"mdegroot09@gmail.com"');
  var threads = GmailApp.search('in:"Pricing for Casey"');
  var messages = threads[1].getMessages();
  messages.reverse();
  
  var blob = messages[0].getAttachments()[0].copyBlob()
  var resource = {
    title: blob.getName(),
    mimeType: blob.getContentType()
  };
  
  // Enable the Advanced Drive API Service
  var file = Drive.Files.insert(resource, blob, {ocr: true, ocrLanguage: "en"});
  
  var doc = DocumentApp.openById(file.id);
  var text = doc.getBody().getText();
  text = String(text).split('\n')
  
  text.forEach(function(a,i){
    sheet.getRange('A' + (4 + i)).setValue(a);
  })
}