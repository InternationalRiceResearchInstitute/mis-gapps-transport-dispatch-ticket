function onClickHandler(e){
var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getActiveSheet().getName()
  Logger.log(sheet);  
 
}
