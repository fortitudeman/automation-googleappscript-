// Add Custom Item In Menu
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Automation')
    .addItem('Es Vs Actual Daily Report', 'dailyReport')
    .addItem('Es Vs Actual Monthly Report', 'monthlyReport')
    .addItem('Daily Hours Logged Report', 'logReport')
    .addItem('Daily Hours New Month Creation', 'newMonthCreate')
    .addToUi();
}
function modifySheet(sheet) {
   //delete columns except 'Project','Name','Estimated time'
   var headings = sheet.getDataRange().offset(0, 0, 1).getValues()[0];
                    
   sheet.deleteColumns(1, headings.indexOf('Project'))
   sheet.deleteColumns(2, headings.indexOf('Name')-headings.indexOf('Project')-1);
   sheet.deleteColumns(3, headings.indexOf('Estimated Time')-headings.indexOf('Name')-1);
   sheet.deleteColumns(4, headings.length-headings.indexOf('Estimated Time')-1);
   
   //Set the style 
   sheet.setColumnWidth(1, 300);
   sheet.setColumnWidth(2, 600);
   sheet.setColumnWidth(3, 100);
   
   //Rearrange Columns Order
    var columnSpec = sheet.getRange("A1:A");
    sheet.moveColumns(columnSpec, 4);
    var columnSpec = sheet.getRange("B1:B");
    sheet.moveColumns(columnSpec, 4); 
}
