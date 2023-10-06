function extractSpreadsheetIdFromUrl(url) {
  var idPattern = /\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/;
  var matches = url.match(idPattern);
  if (matches && matches.length > 1) {
    return matches[1];
  }
  return null;
}



function archiveYearlySheets(spreadsheetId) {
  //var spreadsheetId = extractSpreadsheetIdFromUrl('https://docs.google.com/spreadsheets/'); 
  var archiveYear = '22'; 
  //var archiveYear = new Date().getFullYear().toString().slice(-2);

  var parentFolder = DriveApp.getFileById(spreadsheetId).getParents().next();
  var archiveFolderName = 'Archive';
  var archiveFolder;
  var folders = parentFolder.getFoldersByName(archiveFolderName);

  if (folders.hasNext()) {
    archiveFolder = folders.next();
  } else {
    archiveFolder = parentFolder.createFolder(archiveFolderName);
  }
  var sourceSpreadsheet = SpreadsheetApp.openById(spreadsheetId);
  
  if (!hasDataForYear(sourceSpreadsheet, archiveYear)) {
    console.log('No data for the current year in the original spreadsheet. Skipping.');
    return;
  }

  var archiveSpreadsheets = archiveFolder.getFilesByName(archiveYear + '-' + sourceSpreadsheet.getName());
  var newArchiveSpreadsheetId;

  if (archiveSpreadsheets.hasNext()) {
    var existingArchiveSpreadsheet = archiveSpreadsheets.next();
    newArchiveSpreadsheetId = existingArchiveSpreadsheet.getId();
    console.log('Archive spreadsheet already created.');
  } else {
    var sheets = sourceSpreadsheet.getSheets();
  
    var newArchiveSpreadsheet = SpreadsheetApp.create(archiveYear + '-' + sourceSpreadsheet.getName());
    newArchiveSpreadsheetId = newArchiveSpreadsheet.getId();

    for (var i = 0; i < sheets.length; i++) {
      var sheet = sheets[i];
      var sheetName = sheet.getName();

      if (sheetName.endsWith('-' + archiveYear)) {
        var copiedSheet = sheet.copyTo(newArchiveSpreadsheet);
        copiedSheet.setName(sheetName.replace('-' + archiveYear, '')); 
        }
    }
    var defaultSheet = newArchiveSpreadsheet.getSheetByName('Sheet1');
    if (defaultSheet) {
      newArchiveSpreadsheet.deleteSheet(defaultSheet);
    }

    DriveApp.getFileById(newArchiveSpreadsheetId).moveTo(archiveFolder);
    var archiveSpreadsheetUrl = newArchiveSpreadsheet.getUrl();
  
    var hyperlinkText = archiveYear + '-' + sourceSpreadsheet.getName();
  
    var archiveSheet = sourceSpreadsheet.getSheetByName('Archive');
  if (!archiveSheet) {
    archiveSheet = sourceSpreadsheet.insertSheet('Archive');
  }
    var sheets = sourceSpreadsheet.getSheets();
    sourceSpreadsheet.setActiveSheet(archiveSheet);
    sourceSpreadsheet.moveActiveSheet(sheets.length);

    var protection = archiveSheet.protect();
    protection.setDescription('Archive Sheet Protection');
    archiveSheet.getRange(1, 1, archiveSheet.getMaxRows(), archiveSheet.getMaxColumns()).setValue(null);
    protection.addEditor(Session.getActiveUser());
  
    var archiveSpreadsheet = SpreadsheetApp.openById(newArchiveSpreadsheetId);
    var archiveSheetExists = archiveSpreadsheet.getSheetByName('Archive');

    
    if (archiveSheetExists) {
      console.log('Archive sheet already created.');
    } else {
      var lastRow = archiveSheet.getLastRow() + 1;
      var archiveHyperlink = '=HYPERLINK("' + archiveSpreadsheetUrl + '", "' + hyperlinkText + '")';
      archiveSheet.getRange('A' + lastRow).setFormula(archiveHyperlink);
      console.log('Archive sheet created.');
    }
    sheets_length = sheets.length
    for (var i = 0; i < sheets_length; i++) {
      var sheet = sheets[i];
      var sheetName = sheet.getName();
      if (sheetName.endsWith('-' + archiveYear)) {
        sourceSpreadsheet.deleteSheet(sheet);
        console.log('Deleted sheet: ' + sheetName);
      }
    }
  }
  
}
  
function hasDataForYear(sourceSpreadsheet, archiveYear) {
  var sheets = sourceSpreadsheet.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    var sheet = sheets[i];
    var sheetName = sheet.getName();
    if (sheetName.endsWith('-' + archiveYear)) {
      var dataRange = sheet.getDataRange();
      var values = dataRange.getValues();
      for (var row = 0; row < values.length; row++) {
        for (var col = 0; col < values[row].length; col++) {
          if (values[row][col] !== "") {
            return true; 
          }
        }
      }
    }
  }

  return false; 
}

  
// this fun is for looping throuth the sheet with the all spreadsheets
function archiveMultipleSpreadsheets() {
  var otherSpreadsheetId = 'otherSpreadsheetId';
  var otherSpreadsheet = SpreadsheetApp.openById(otherSpreadsheetId);
  var sheet = otherSpreadsheet.getSheetByName('list'); 

  var spreadsheetIds = sheet.getRange('B2:B' + sheet.getLastRow()).getValues();

  for (var i = 0; i < spreadsheetIds.length; i++) {
    var spreadsheetId = spreadsheetIds[i][0];
    console.log(spreadsheetId);
    if (spreadsheetId) { 
      archiveYearlySheets(extractSpreadsheetIdFromUrl(spreadsheetId));
    }
  }
}
  
  
  