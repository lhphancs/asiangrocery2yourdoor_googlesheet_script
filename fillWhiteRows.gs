function fillWhiteRowsBtn(){
  var activeSheet = SpreadsheetApp.getActiveSheet();
  var activeSheetInfo = new SheetInfo(activeSheet);
  
  var productNameCoordinate = getRowColCoordinateOfStr(activeSheetInfo, REPLENISH_HEADER_PRODUCT_NAME);
  var unitSoldAmtLast30DaysCoordinate = getRowColCoordinateOfStr(activeSheetInfo, REPLENISH_HEADER_UNIT_SOLD_LAST_30_DAYS);
  var myCommentCoordinate = getRowColCoordinateOfStr(activeSheetInfo, REPLENISH_HEADER_MY_COMMENT);
  
  var activeSheetValues = activeSheetInfo.sheetValues;
  for(var i = productNameCoordinate.rowIndex+1; i < activeSheetInfo.amtRow; ++i){
    var range = activeSheet.getRange(i+1, productNameCoordinate.colIndex+1);
    var cellBackgroundColor = range.getBackground();
    
    if(cellBackgroundColor == '#ffffff'){
      var unitSoldAmtLast30DaysCellVal = activeSheetValues[i][unitSoldAmtLast30DaysCoordinate.colIndex];
      var sendAmt = 1.5*unitSoldAmtLast30DaysCellVal;
      var formattedDate = Utilities.formatDate(new Date(), "PST", "MM/dd");
      var writeCellValue = formattedDate + " Send " + sendAmt;
      activeSheet.getRange(i+1, myCommentCoordinate.colIndex+1).setValue(writeCellValue);
    }
  }
}
