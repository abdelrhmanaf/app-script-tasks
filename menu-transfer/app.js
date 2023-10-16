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
      var fileId = extractFileId(sheetUrl);
      var file = DriveApp.getFileById(fileId);
      var mimeType = file.getMimeType();

      if (mimeType === MimeType.MICROSOFT_EXCEL) {
        var googleSheetsId = convertXLSXtoGoogleSheets(fileId);
        copyDataFromGoogleSheets(googleSheetsId, sheetName);
        addNamesToSheet();
        DriveApp.getFileById(googleSheetsId).setTrashed(true);
      } else if (mimeType === MimeType.GOOGLE_SHEETS) {
        copyDataFromGoogleSheets(fileId, sheetName);
        addNamesToSheet();
      } else {
        ui.alert('Unsupported file format. Please provide a valid XLSX or Google Sheets file.');
      }
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
    0: 19,
    1: 1, 
    3: 5,
    4: 11,
    8: 2,  
    9: 3, 
    10: 4,
    11: 7,
    11: 17,
    13: 8,
    14: 9,
    15: 10,
    16: 12, 
    17: 13, 
    18: 14, 
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
function addNamesToSheet_test() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSheet = spreadsheet.getSheetByName('ROMI_acronym_maping');
  var destSheet = spreadsheet.getSheetByName('Trader_Consolidation');
  //var ssexported = SpreadsheetApp.openById('');
  var ssexported = SpreadsheetApp.openById(sheetid);
  //var exportedSheet = ssexported.getSheetByName('Export');
  var exportedSheet = ssexported.getSheetByName(exsheetname);
  var sourceData = sourceSheet.getDataRange().getValues();
  var exporttedData = exportedSheet.getDataRange().getValues();
  var destData = destSheet.getDataRange().getValues();
  var nameToAcronymMapping = createNameToAcronymMapping(sourceData);
  
  var nameColumnIndex = 19; 
  var acronymColumnIndex = 19;

  for (var i = 1; i < destData.length; i++) {
    var name = destData[i][nameColumnIndex];
    if (name && !destData[i][acronymColumnIndex] && nameToAcronymMapping[name]) {
      destData[i][acronymColumnIndex+1] = nameToAcronymMapping[name];
    }
  }
  //destSheet.deleteColumn(nameColumnIndex);
  
 destSheet.getRange(1, 1, destData.length, destData[0].length).setValues(destData);
}

function createNameToAcronymMapping(data) {
  var mapping = {};
  for (var i = 1; i < data.length; i++) {
    var name = data[i][0]; 
    var acronym = data[i][1]; 
    mapping[name] = acronym;
  }
  return mapping;
}