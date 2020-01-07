function newYear() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var password = 'brownies';
  var input, check, year, newYear, row, numRows, names, sheetName;
  while (!check) {
    input = ui.prompt('Please Enter Password', 'Please enter the password, then press "Ok" to wipe all sheets and set them all up for the new year.', ui.ButtonSet.OK_CANCEL);
    
    if (input.getSelectedButton() != ui.Button.OK) return;
    
    if (input.getResponseText() == '') {
      ui.alert('Blank', 'Password cannot be blank! Please try again', ui.ButtonSet.OK);
      continue;
    }
    
    if (input.getResponseText().toLowerCase() != password) {
      ui.alert('Incorrect', 'That password was incorrect. Please try again.', ui.ButtonSet.OK);
      continue;
    }
    
    check = true;
  }
  
  if (ui.alert('Are You Sure?', 'Running this function will be deleting all the previous year\'s information. Are you sure you want to continue?', ui.ButtonSet.YES_NO) != ui.Button.YES) return;
  var sheets = ss.getSheets();
  var dealers = driver('dealers');
  
  if (sheets[0].getSheetName().indexOf('YTD') == -1) {
    ui.alert('Incorrect Order', 'The YTD sheet MUST be at the front of the sheet list. Please find the YTD sheet and click and drag it to the far left, then try again.', ui.ButtonSet.OK);
    return;
  }
  
  year = sheets[0].getSheetName().split('YTD ')[1];
  newYear = parseInt(year, 10) + 1;
  sheets[0].getRange(1, 5).setValue(newYear);
  sheets[0].setName(sheets[0].getName().replace(year, newYear));
  
  for (var i = 1; i < sheets.length; i++) {
    for (var j = 0; j < dealers.length; j++) {
      names = getNames(j, sheets[i].getSheetName());
      row = names[1];
      numRows = parseInt(names[2], 10) - parseInt(row, 10) + 1;
      sheets[i].getRange(row, 3, numRows, 3).setValue('');
    }
    
    sheetName = sheets[i].getSheetName().split(' ');
    sheetName[1] = newYear;
    sheets[i].setName(sheetName.join(' '));
  }
}
