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
function deleteRows(sheet){
   
    var RANGE = sheet.getDataRange();
    var DELETE_VAL = ['Project Management','PM Activities','CH Internal-Angela Harper'];
    
     // The column to search for the DELETE_VAL (Zero is first)
    var COL_TO_SEARCH = 0;
    var rangeVals = RANGE.getValues();
    var newRangeVals = [];
    
    Logger.log(rangeVals[1][0]);
    for(var i = 0; i < DELETE_VAL.length; i++){
      for(var n = rangeVals.length-1 ; n >=0  ; n--){
          if(rangeVals[n][COL_TO_SEARCH].toLowerCase() === DELETE_VAL[i].toLowerCase()){   
            sheet.deleteRow(n+1);
          };
      } 
    };
 
}
function duplicateProcess(sheet){
    // Sort the sheet without sorting header
    sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).sort(1);
    var column = 1;
    var lastRow = sheet.getLastRow();
    var columnRange = sheet.getRange(1,column,lastRow);
    var rangeArray = columnRange.getValues();
    //Convert to one dimensional array
    rangeArray = [].concat.apply([],rangeArray);
   
    //sort the data and find duplicates
   
    
    var duplicates = [];
    var indexes = [];
    for(var i=0;i<rangeArray.length-1;i++){
        if(rangeArray[i].toLowerCase()==rangeArray[i+1].toLowerCase()){
            duplicates.push(rangeArray[i]);
            indexes.push(i);
        }
    }
    
    var  count = {};
    duplicates.forEach(function(i) { count[i] = (count[i]||0) + 1;});
        
    //Highlight all the duplicates
    for(var i=0;i< indexes.length;i++){
      sheet.getRange(indexes[i]+1, column).setBackground("yellow");
    }
    // Get the duplicated str and its number of repeating
    var dup = [];
    for (var key in count){
       if(count.hasOwnProperty(key)){
           dup.push(key)
       }
    }
    //Add the hypen and number to duplicated cell
    for(var i=0;i<dup.length;i++){
      var firstIndex = rangeArray.indexOf(dup[i]);
      for(var j=0;j<=count[dup[i]];j++){
          var val = sheet.getRange(firstIndex+1+j, column).getValue()+"-"+j;
          sheet.getRange(firstIndex+1+j, column).setValue(val);
      }
    }
    
    
    
}
function dailyReport() {
  var sheet = SpreadsheetApp.getActiveSheet();
  //modifySheet(sheet);
  //deleteRows(sheet)
  duplicateProcess(sheet);
}