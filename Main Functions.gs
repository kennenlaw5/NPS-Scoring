function onOpen() {
  var ui = SpreadsheetApp.getUi();
  var menu = ui.createMenu('Utilities').addSubMenu(ui.createMenu('BMW').addItem('Detractor','BD').addItem('Passive','BPa').addItem('Promoter','BPr'))
  .addSubMenu(ui.createMenu('Mini').addItem('Detractor','MD').addItem('Passive','MPa').addItem('Promoter','MPr')).addItem('Add Adviser', 'addAdviser').addItem('Remove Adviser', 'removeAdviser');
  var month = new Date().getMonth();
  if (month == 0) { menu.addItem('New Year', 'newYear'); }
  menu.addToUi();
  //.addItem('Reset Statistics','reset').addItem('Refresh CA Ranking','rank').addToUi();
  var message = 'The spreadsheet has loaded successfully! Have a great day!';
  var title = 'Complete!';
  SpreadsheetApp.getActiveSpreadsheet().toast(message, title);
}

function BD() { scraper(0, 2); }

function BPa() { scraper(0, 1); }

function BPr() { scraper(0, 0); }

function MD() { scraper(1, 2); }

function MPa() { scraper(1, 1); }

function MPr() { scraper(1, 0); }

function scraper(dealer,type) {
  //BMW=0, Mini=1;
  //Detractor=2; Passive=1;Promoter=0
  //dealer=0;type=1;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var sheetName = ss.getSheetByName(driver('main_sheet')).getRange('F1').getDisplayValue();
  var sheet = ss.getSheetByName(sheetName);
  if (sheet == null) { 
    ui.alert('Sheet Not Found', 'The sheet named "' + sheetName 
             + '" could not be found. Please check the sheet names for any spelling errors and try again.', ui.ButtonSet.OK);
    return;
  }
  var source = ss.getActiveSheet();
  var column, row;
  var names = getNames(dealer);
  row = names[1];
  names = names[0];
  var found = false;
  var num = 0;
  var range = source.getRange(1, 1, source.getLastRow(), source.getLastColumn()).getValues();
  
  for (var i = 0; i < range[0].length; i++) {
    if(range[0][i] == 'Advisor Name') { column = parseInt(i); }
  }
  
  for (i = 0; i < names.length; i++) {
    names[i] = [names[i], 0];
  }
  
  for (var i = 1; i < range.length; i++) {
    if(range[i][0] != ''){
      found = false;
      for (var j = 0; j < names.length && found == false; j++) {
        if (range[i][column] == names[j][0]) {
          found = true;
          names[j][1] += 1;
        }
      }
      if (found == false) {
        names[names.length-1][1] += 1;
      }
    } else { i = range.length; }
  }
  
  var final = [];
  for (i = 0; i < names.length; i++) {
    final[i] = [names[i][1]];
  }
  
  sheet.getRange(row, parseInt(type) + 3 ,names.length).setValues(final);
  ss.setActiveSheet(sheet);
  ss.deleteSheet(source);
}

function addAdviser() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var input, sheets, row, range, formulas, ytd;
  var valid = false;
  while (!valid) {
    input = ui.prompt('Name', 'Please enter the adviser\'s name exactly as it appears on the NPS site then press "Ok".', ui.ButtonSet.OK_CANCEL);
    if (input.getSelectedButton() != ui.Button.OK) { return; }
    var adviser = input.getResponseText();
    if (adviser == '' || adviser.split(' ').join('') == '') { ui.alert('Field Empty', 'Nothing was entered. Please try again.', ui.ButtonSet.OK); }
    else { valid = true; }
  }
  valid = false;
  while (!valid) {
    input = ui.prompt('Dealership', 'Please enter the dealership the adviser is asigned to. (The options are BMW and Mini)', ui.ButtonSet.OK_CANCEL);
    if (input.getSelectedButton() != ui.Button.OK) { return; }
    var dealer = input.getResponseText();
    if (dealer == '' || dealer.split(' ').join('') == '') { ui.alert('Field Empty', 'Nothing was entered. Please try again.', ui.ButtonSet.OK); continue; }
    dealer = driver('dealers').indexOf(input.getResponseText().toUpperCase());
    if (dealer == -1) { ui.alert('Invalid Input', 'Please enter a valid response. The options are: BMW and Mini', ui.ButtonSet.OK); }
    else { valid = true; }
  }
  
  sheets = ss.getSheets();
  
  for (var i = 0; i < 13; i++) {
    sheets[i].activate();
    row = getNames(dealer, sheets[i].getSheetName())[2];
    range = sheets[i].getRange(row, 2, 1, sheets[i].getLastColumn() - 1);
    formulas = range.getFormulas();
    sheets[i].insertRowBefore(row);
    if (i == 0 && sheets[i].getName().split(' ')[0].toUpperCase() == 'YTD') { ytd = [formulas, row]; }
    else { range.setFormulas(formulas); }
    sheets[i].getRange(row, 1).setValue(adviser);
    if (i == 12 && ytd[0].length > 0) { 
      sheets[0].activate();
      sheets[0].getRange(ytd[1], 2, 1, sheets[0].getLastColumn() - 1).setFormulas(ytd[0]);
    }
  }
  ss.toast(adviser + ' was added successfully to the sheet!', 'Complete!');
}

function removeAdviser() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var input, sheets, row, names;
  var valid = false;
  while (!valid) {
    input = ui.prompt('Name', 'Please enter the adviser\'s name exactly as it appears on the spreadsheet then press "Ok".', ui.ButtonSet.OK_CANCEL);
    if (input.getSelectedButton() != ui.Button.OK) { return; }
    var adviser = input.getResponseText();
    if (adviser == '' || adviser.split(' ').join('') == '') { ui.alert('Field Empty', 'Nothing was entered. Please try again.', ui.ButtonSet.OK); }
    else { valid = true; }
  }
  valid = false;
  while (!valid) {
    input = ui.prompt('Dealership', 'Please enter the dealership the adviser is asigned to. (The options are BMW and Mini)', ui.ButtonSet.OK_CANCEL);
    if (input.getSelectedButton() != ui.Button.OK) { return; }
    var dealer = input.getResponseText();
    if (dealer == '' || dealer.split(' ').join('') == '') { ui.alert('Field Empty', 'Nothing was entered. Please try again.', ui.ButtonSet.OK); continue; }
    dealer = driver('dealers').indexOf(input.getResponseText().toUpperCase());
    if (dealer == -1) { ui.alert('Invalid Input', 'Please enter a valid response. The options are: BMW and Mini', ui.ButtonSet.OK); }
    else { valid = true; }
  }
  
  sheets = ss.getSheets();
  
  for (var i = 0; i < 13; i++) {
    valid = false;
    sheets[i].activate();
    names = getNames(dealer, sheets[i].getSheetName());
    row = names[1];
    names = names[0];
    Logger.log([sheets[i].getSheetName(), row]);
    for (var j = 0; j < names.length; j++) {
      if (names[j].toUpperCase() == adviser.toUpperCase()) { row += j; valid = true; break; }
    }
    if (valid) { sheets[i].deleteRow(row); }
  }
  sheets[0].activate();
  ss.toast(adviser + ' was deleted from the sheet successfully!', 'Complete!');
}

function test(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) { sheets[i] = sheets[i].getName(); }
  var columns = ['C', 'D', 'E'];
  var row = 27;
  var formulas = [];
  for (i = 0; i < 5; i++) {
    formulas[i] = [];
    for (var j = 0; j < 3; j++){
      formulas[i][j] = "=SUM(";
      for (var k = 1; k < sheets.length; k++) {
        formulas[i][j] += "'" + sheets[k] + "'!" + columns[j] + "" + (i + row);
        if ( k + 1 != sheets.length) { formulas[i][j] += ','; }
      }
      formulas[i][j] += ')';
    }
  }
  ss.getSheetByName(sheets[0]).getRange("C27:E31").setValues(formulas);
}
