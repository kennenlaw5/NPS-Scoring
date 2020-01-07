function driver (input) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  switch (input) {
    case 'main_sheet':
      var mainSheet;
      
      ss.getSheets().forEach(function (sheet) {
        if (sheet.getSheetName().indexOf('YTD') !== -1) mainSheet = sheet.getSheetName();
      });
      
      return mainSheet;
    case 'dealers':
      return ['BMW', 'MINI'];
  }
}

function getNames (dealer /*REQUIRED*/, sheetName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  if (!sheetName) sheetName = driver('main_sheet');
  
  var sheet = ss.getSheetByName(sheetName);
  var values = sheet.getRange(1, 1, sheet.getLastRow()).getDisplayValues();
  var compile = false;
  var found = false;
  var names = [];
  var startRow, endRow;
  dealer = driver('dealers')[dealer];
  
  values.forEach(function (value, index) {
    if (endRow) return;
    if (compile) {
      names.push(value[0]);
      
      if (names.length === 1) startRow = index + 1;
      
      if (value[0].toUpperCase() === 'OTHER') {
        endRow = index + 1;
        compile = false;
        found = false;
      }
    }
    else if (!compile && !found && value[0].toUpperCase().indexOf(dealer) !== -1) found = true;
    else if (found && !compile && value[0].toUpperCase().indexOf('ADVISER') !== -1) compile = true;
  });
  
  return [names, startRow, endRow];
}
