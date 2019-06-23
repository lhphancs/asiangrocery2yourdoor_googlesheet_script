function fillRedRows(wholesaleMap) {
  var replenishSheet = SpreadsheetApp.getActiveSheet();
  var replenishSheetInfo = new SheetInfo(replenishSheet);
  
  var replenishAsinRowColCoordinate = getRowColCoordinateOfStr(replenishSheetInfo, REPLENISH_HEADER_ASIN);
  var replenishCommentRowColCoordinate = getRowColCoordinateOfStr(replenishSheetInfo, REPLENISH_HEADER_MY_COMMENT);
  
  if(replenishAsinRowColCoordinate && replenishCommentRowColCoordinate){
    colorRowAndCommentIfNeeded(replenishSheet, replenishSheetInfo, wholesaleMap, replenishAsinRowColCoordinate, replenishCommentRowColCoordinate);
  }
  else
    throw( "Undetected headers in sheet. No edits were made.\n\n'" + REPLENISH_HEADER_ASIN + "' or '" + REPLENISH_HEADER_MY_COMMENT + "' was not found in replenish sheet." );
}

function colorRowAndCommentIfNeeded(replenishSheet, replenishSheetInfo, wholesaleMap, replenishAsinRowColCoordinate, replenishCommentRowColCoordinate){
  var replenishSheetValues = replenishSheetInfo.sheetValues;
  
  // rowIndex has + 1 because we want to skip the header
  for(var rowIndex = replenishAsinRowColCoordinate.rowIndex + 1; rowIndex < replenishSheetInfo.amtRow; ++rowIndex){
    var asin = replenishSheetValues[rowIndex][replenishAsinRowColCoordinate.colIndex];
    if(asin in wholesaleMap){
      var wholesaleColor = wholesaleMap[asin].color;
      
      if(wholesaleColor == COLOR_BKGD_RED)
      {
        replenishSheet.getRange(rowIndex+1, replenishCommentRowColCoordinate.colIndex + 1).setValue('Check profit before send');
        replenishSheet.getRange(rowIndex+1, 1, 1, replenishSheetValues[0].length).setBackground(COLOR_BKGD_RED);
      }
      
    }
  }
}