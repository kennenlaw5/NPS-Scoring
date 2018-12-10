function driver (input) {
  switch (input) {
    case 'main_sheet':
      return 'YTD 2018';
      break;
    case 'dealers':
      var dealers = ['BMW', 'MINI'];
      return dealers;
      break;
  }
}

function getNames (dealer /*REQUIRED*/, sheet) {
  dealer = 0;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (sheet == undefined) { sheet = ss.getSheetByName(driver('main_sheet')); }
  else { sheet = ss.getSheetByName(sheet); }
  var values = sheet.getRange(1, 1, sheet.getLastRow()).getDisplayValues();
  var compile = false;
  var names = [];
  var row;
  dealer = driver('dealers')[dealer];
  
  for (var i = 0; i < values.length; i++) {
    if (compile && values[i][0] != '' && values[i][0].toUpperCase().indexOf('ADVISER') == -1) {
      names.push(values[i][0]);
      if (names.length == 1) { row = i + 1; }
      if (values[i][0].toUpperCase() == 'OTHER') { compile = false; break; }
    } else if (!compile && values[i][0].toUpperCase() == dealer) {
      compile = true;
    }
  }
  return [names, row];
}
