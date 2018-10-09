function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Utilities').addSubMenu(ui.createMenu('BMW').addItem('Detractor','BD').addItem('Passive','BPa').addItem('Promoter','BPr'))
  .addSubMenu(ui.createMenu('Mini').addItem('Detractor','MD').addItem('Passive','MPa').addItem('Promoter','MPr')).addToUi();
  //.addItem('Reset Statistics','reset').addItem('Refresh CA Ranking','rank').addToUi();
  var message = 'The spreadsheet has loaded successfully! Have a great day!';
  var title = 'Complete!';
  SpreadsheetApp.getActiveSpreadsheet().toast(message, title);
}
function BD(){scraper(0,2);}
function BPa(){scraper(0,1);}
function BPr(){scraper(0,0);}
function MD(){scraper(1,2);}
function MPa(){scraper(1,1);}
function MPr(){scraper(1,0);}

function scraper(dealer,type) {
  //BMW=0, Mini=1;
  //Detractor=2; Passive=1;Promoter=0
  //dealer=0;type=1;
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var ui=SpreadsheetApp.getUi();
  var sheetName=ss.getSheetByName("YTD 2018").getRange("F1").getDisplayValue();
  var sheet=ss.getSheetByName(sheetName);
  var source=ss.getActiveSheet();var column;var row;
  if (dealer == 0) { var names=ss.getSheetByName("YTD 2018").getRange("A5:A19").getValues(); row=5; }
  else if (dealer == 1) { var names=ss.getSheetByName("YTD 2018").getRange("A28:A33").getValues(); row=28; }
  var found=false; var num=0;
  var range=source.getRange(1,1,source.getLastRow(),source.getLastColumn()).getValues();
  Logger.log(range.length);
  for(var i=0;i<range[0].length;i++){
    if(range[0][i]=="Advisor Name"){column=parseInt(i);}
  }
  Logger.log(column);
  for(i=0;i<names.length;i++){
    names[i]=[names[i][0],0];
  }
  Logger.log(names);
  for(var i=1;i<range.length;i++){
    if(range[i][0]!=""){
      found=false;
      for (var j = 0; j < names.length && found == false; j++) {
        if (range[i][column] == names[j][0]) {
          found=true;
          names[j][1]+=1;
          Logger.log("FOUND");
        }
      }
      if(found==false){
        names[names.length-1][1]+=1;
        Logger.log(range[i][column]);
      }
    }else{i=range.length;}
  }
  Logger.log(names);
  var final=[];
  for(i=0;i<names.length;i++){
    final[i]=[names[i][1]];
  }
  Logger.log(sheet);
  sheet.getRange(row,parseInt(type)+3,names.length).setValues(final);
  ss.setActiveSheet(sheet);
  ss.deleteSheet(source);
}
function test(){
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var sheets=ss.getSheets();
  for(var i=0;i<sheets.length;i++){sheets[i]=sheets[i].getName();}
  var columns=["C","D","E"];var row=27;var formulas=[];
  for(i=0;i<5;i++){
    formulas[i]=[];
    for(var j=0;j<3;j++){
      formulas[i][j]="=SUM(";
      for(var k=1;k<sheets.length;k++){
        formulas[i][j]+="'"+sheets[k]+"'!"+columns[j]+""+(i+row);
        if(k+1!=sheets.length){formulas[i][j]+=",";}
      }
      formulas[i][j]+=")"
    }
  }
  ss.getSheetByName(sheets[0]).getRange("C27:E31").setValues(formulas);
  Logger.log(formulas);
}