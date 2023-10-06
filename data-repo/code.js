function onFormSubmit(e) {
    try{
      var formResponseSheet = e.range.getSheet();
      var formResponses = formResponseSheet.getRange(e.range.getRow(), 2, 1, formResponseSheet.getLastColumn() - 1).getValues();
  
      var liveSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Live');
  
      var employeesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Employees');
      var dateValue = formResponses[0][0]; 
      var emailAddress = formResponses[0][1]; 
      Logger.log(emailAddress);
      var employeeName = getEmployeeName(employeesSheet, emailAddress);
  
      var rowData = [employeeName,dateValue].concat(formResponses[0].slice(2));
      liveSheet.appendRow(rowData);
    } catch (e){
      notifyError(e)
    }
  
  }
  
  function notifyError(e){
    MailApp.sendEmail("email", "Error report: Aftersales - Status Sheet", 
                      "\r\nMessage: " + e.message
                      + "\r\nFile: " + e.fileName
                      + "\r\nLine: " + e.lineNumber
                      + "\r\nLine: " + e.stack
                      + "\r\nLink: https://docs.google.com/spreadsheets/d/");
  }
  
  function getEmployeeName(employeesSheet, emailAddress) {
    var emailColumnValues = employeesSheet.getRange(2, 1, employeesSheet.getLastRow() - 1, 1).getValues();
    var nameColumnValues = employeesSheet.getRange(2, 2, employeesSheet.getLastRow() - 1, 1).getValues();
        Logger.log(emailColumnValues);
    for (var i = 0; i < emailColumnValues.length; i++) {
      if (emailColumnValues[i][0] === emailAddress) {
  
        return nameColumnValues[i][0];
      }
  }
    return 'Name not found';
  }
  
  
  
  function sendReminderEmails() {
    try {
      var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      var employeesSheet = spreadsheet.getSheetByName('employees');
      var emailColumn = employeesSheet.getRange('A3:C' + employeesSheet.getLastRow()).getValues();
      var nameColumn = employeesSheet.getRange('B3:B' + employeesSheet.getLastRow()).getValues();
      var formUrl = 'https://forms';
  
      for (var i = 0; i < emailColumn.length; i++) {
        var emailAddress = emailColumn[i][0];
        Logger.log(emailAddress);
        var employeeName = nameColumn[i][0];
        Logger.log(employeeName);
  
        var emailTemplate = HtmlService.createTemplateFromFile("EmployeeWeeklyReminder");
        emailTemplate.employee = employeeName;
        emailTemplate.url = formUrl;
        
        var emailContent = emailTemplate.evaluate().getContent();
  
        var subject = 'Aftersales Weekly Status Form Reminder';
  
        MailApp.sendEmail({
          to: emailAddress,
          subject: subject,
          htmlBody: emailContent
        });
      }
    } catch (e) {
      notifyError(e)
    }
  }
  
  
  
  