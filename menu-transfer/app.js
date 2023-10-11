function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Custom Menu')
    .addItem('Copy Data', 'showDialog')
    .addToUi();
}

function showDialog() {
  var ui = SpreadsheetApp.getUi();
  var urlPrompt = ui.prompt('Enter the URL of the Emailed sheet');
  
  if (urlPrompt.getSelectedButton() == ui.Button.OK) {
    var sheetUrl = urlPrompt.getResponseText();
    
    var sheetNamePrompt = ui.prompt('Enter the sheet name of the Emailed sheet');
    
    if (sheetNamePrompt.getSelectedButton() == ui.Button.OK) {
      var sheetName = sheetNamePrompt.getResponseText();
      var destFolder = DriveApp.getFileById(SpreadsheetApp.getActiveSpreadsheet().getId()).getParents().next();
      var folder = destFolder.createFolder("TempFolder");
      var fileId = extractFileId(sheetUrl);
      var file = DriveApp.getFileById(fileId);
      var mimeType = file.getMimeType();

      if (mimeType === MimeType.GOOGLE_SHEETS) {
        var googleSheetsId = file.makeCopy("TempGoogleSheets", folder).getId();
        copyDataFromGoogleSheets(googleSheetsId, sheetName);
      } else if (mimeType === MimeType.MICROSOFT_EXCEL) {
        var xlsxId = file.makeCopy("TempXLSX", folder).getId();
        var googleSheetsId = convertXLSXtoGoogleSheets(xlsxId);
        copyDataFromGoogleSheets(googleSheetsId, sheetName);
        DriveApp.getFileById(xlsxId).setTrashed(true);
        DriveApp.getFileById(googleSheetsId).setTrashed(true);
        
      } else {
        ui.alert('Unsupported file format. Please provide a valid XLSX or Google Sheets file.');
      }
      folder.setTrashed(true);
    } else {
      ui.alert('Sheet name not provided. Please try again.');
    }
  } else {
    ui.alert('URL not provided. Please try again.');
  }
}
function extractFileId(url) {
  var fileId = /\/d\/([^/]+)/.exec(url);
  return fileId ? fileId[1] : null;
}

function convertXLSXtoGoogleSheets(fileId) {
  var file = DriveApp.getFileById(fileId);
  var blob = file.getBlob();
  var resource = {
    title: file.getName(),
    mimeType: MimeType.GOOGLE_SHEETS
  };
  var createdFile = Drive.Files.insert(resource, blob);
  return createdFile.id;
}

function copyDataFromGoogleSheets(sourceSheetId, sourceSheetName) {
  var destSheet = SpreadsheetApp.getActiveSheet();
  
  var sourceSpreadsheet = SpreadsheetApp.openById(sourceSheetId);
  var sourceSheet = sourceSpreadsheet.getSheetByName(sourceSheetName);
  
  if (!sourceSheet) {
    Logger.log("Source sheet not found. Make sure the sheet name is correct.");
    return;
  }

  var sourceData = sourceSheet.getDataRange().getValues();
  var destData = destSheet.getDataRange().getValues();
  var columnIndexMapping = { 
    1: 1, 
    3: 5,
    4: 11,
    8: 2,  
    9: 3, 
    10: 4,
    11: 7,
    13: 8,
    14: 9,
    15: 10,
    16: 12, 
    17: 13, 
    18: 14, 
    19: 17,
    23: 18,
  };

  var mappedData = [];
  Logger.log("Source Column Headers: " + sourceData[0]);

  for (var i = 1; i < sourceData.length; i++) {
    var mappedRow = [];
    for (var j = 0; j < sourceData[i].length; j++) {
      var sourceColumnIndex = j;
      var destColumnIndex = columnIndexMapping[sourceColumnIndex];
      
      if (destColumnIndex !== undefined) {
        mappedRow[destColumnIndex] = sourceData[i][j];
      }
    }
    mappedData.push(mappedRow);
  }
  Logger.log("Mapped Destination Columns: " + Object.values(columnIndexMapping));

  if (mappedData.length > 0) {
    destSheet.getRange(destData.length + 1, 1, mappedData.length, mappedData[0].length).setValues(mappedData);
    Logger.log("Data appended successfully with mapping!");
  } else {
    Logger.log("No data found in the source sheet.");
  }
}