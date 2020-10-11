function lastRow(sheet,column) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var lastRow = sheet.getLastRow();
  var array = sheet.getRange(column + 1 + ':' + column + lastRow).getValues();
  for (i=0;i<array.length;i++) {
    if (array[i] != '') {       
      var final = i + 1;
    }
  }
  if (final != null) {
    return final;
  } else {
    return 0;
  }
}