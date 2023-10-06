function moveRowsToBackup() {
    try{
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var liveSheet = ss.getSheetByName('Live');
      var backupSheet = ss.getSheetByName('Backup');
      var statusColumn = 8;
      var targetValue = 'Done';
  
      var data = liveSheet.getDataRange().getValues();
      var headers = data[0];
      var newData = [headers]; 
  
      var rowsToDelete = [];
      var rowsToMove = [];
  
      for (var i = 1; i < data.length; i++) {
        var status = data[i][statusColumn - 1];
        if (status !== "" && status === targetValue) {
          rowsToMove.push(data[i]);
          rowsToDelete.push(i + 1); 
        }
      }
  
      if (rowsToMove.length > 0) {
        var rangeToMove = backupSheet.getRange(backupSheet.getLastRow() + 1, 1, rowsToMove.length, headers.length);
        rangeToMove.setValues(rowsToMove);
  
        for (var j = rowsToDelete.length - 1; j >= 0; j--) {
          liveSheet.deleteRow(rowsToDelete[j]);
        }
  
        Logger.log("Rows moved to backup: " + rowsToMove.length);
      }
    }catch(e) {
      notifyError(e)
    }
  }
  