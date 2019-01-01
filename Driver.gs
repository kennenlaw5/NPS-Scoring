function driver (input) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  switch (input) {
    case 'main_sheet':
      var sheet;
      var sheets = ss.getSheets();
      for (var i = 0; i < sheets.length; i++) {
        if (sheets[i].getSheetName().indexOf('YTD') != -1) { sheet = sheets[i].getSheetName(); }
      }
      return sheet;
      break;
    case 'dealers':
      var dealers = ['BMW', 'MINI'];
      return dealers;
      break;
  }
}

function getNames (dealer /*REQUIRED*/, sheet) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (sheet == undefined) { sheet = ss.getSheetByName(driver('main_sheet')); }
  else { sheet = ss.getSheetByName(sheet); }
  var values = sheet.getRange(1, 1, sheet.getLastRow()).getDisplayValues();
  var compile = false;
  var found = false;
  var names = [];
  var first, last;
  dealer = driver('dealers')[dealer];
  
  for (var i = 0; i < values.length; i++) {
    if (compile && values[i][0].toUpperCase().indexOf('ADVISER') == -1) {
      names.push(values[i][0]);
      if (names.length == 1) { first = i + 1; }
      if (values[i][0].toUpperCase() == 'OTHER') { compile = false; found = false; last = i + 1; break; }
    } else if (!compile && values[i][0].toUpperCase() == dealer) {
      found = true;
    } else if (found && !compile && values[i][0].toUpperCase().indexOf('ADVISER') != -1) { compile = true; }
  }
  return [names, first, last];
}
