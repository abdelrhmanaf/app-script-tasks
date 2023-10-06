// @ts-nocheck
function CopyToLive(e) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var source = ss.getSheetByName("Form Responses 1");
    var destinationsheet = ss.getSheetByName("live");
    var scriptProperties = PropertiesService.getScriptProperties();
    var lastProcessedRow = parseInt(scriptProperties.getProperty("lastProcessedRow") || 1);
    var numRowsToCopy = source.getLastRow() - lastProcessedRow;
  
    if (numRowsToCopy > 0) {
      var sourceRange = source.getRange(lastProcessedRow + 1, 1, numRowsToCopy, source.getLastColumn());
      var sourceData = sourceRange.getValues();
      var numDestinationColumns = source.getLastColumn() - 1; 
      var destinationRange = destinationsheet.getRange(destinationsheet.getLastRow() + 1, 1, numRowsToCopy, numDestinationColumns);
      var destinationValues = [];
  
  for (var i = 0; i < sourceData.length; i++) {
    var row = sourceData[i];
    var fullName = row[1] + ' ' + row[2];
    var photoLink = row[3];
    var photoHyperlink = '=HYPERLINK("' + photoLink + '","Picture")';
  
    row.splice(1, 3, fullName, photoHyperlink);
    destinationValues.push(row);
  }
  
  destinationRange.setValues(destinationValues);
  
  var newLastProcessedRow = lastProcessedRow + numRowsToCopy;
  scriptProperties.setProperty("lastProcessedRow", newLastProcessedRow.toString());
  }
  }
  
  function onFormSubmit(e) {
    var values =   e.values;
    console.log(values);
    var userEmail = 'mail';
    var subject = 'Form Submission Alert';
    var message = 'A form has been submitted.';
   
   
    var form = FormApp.openById('FormApp');
  
    
    MailApp.sendEmail({
  to:userEmail ,
  subject: subject,
  htmlBody: message +  "\n  Timestamp = " + values[0]+ "\n Full name:" +values[1]+ ' '+values[2]+ "\n Picture: "+values[3]+ "\n  Email: " +values[4]
  });
  }
  
  
  function clearall() {
  
    var app = SpreadsheetApp();
    var spreedsheat =app.getActiveSpreadsheet();
    var  sheet = spreedsheat.getSheetByName("live")
    sheet.getRange("A2:D19").clearContent();
  
  
  }