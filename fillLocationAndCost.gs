function fillLocationAndCost(wholesaleMap) {
  var replenishSheet = SpreadsheetApp.getActiveSheet();
  var replenishSheetInfo = new SheetInfo(replenishSheet);
  
  var replenishAsinRowColCoordinate = getRowColCoordinateOfStr(replenishSheetInfo, REPLENISH_HEADER_ASIN);
  var replenishLocationRowColCoordinate = getRowColCoordinateOfStr(replenishSheetInfo, REPLENISH_HEADER_SHELF_LOCATION);
  var replenishCostRowColCoordinate = getRowColCoordinateOfStr(replenishSheetInfo, REPLENISH_HEADER_COST);
  
  if(replenishAsinRowColCoordinate && replenishLocationRowColCoordinate && replenishCostRowColCoordinate){
    writeLocation(replenishSheet, replenishSheetInfo, wholesaleMap, replenishAsinRowColCoordinate, replenishLocationRowColCoordinate, replenishCostRowColCoordinate);
  }
  else
    throw( "Undetected headers in sheet. No edits were made.\n\n'" + REPLENISH_HEADER_ASIN + "' or '" + REPLENISH_HEADER_SHELF_LOCATION + "' or '" + REPLENISH_HEADER_COST + "' was not found in replenish sheet." );
}

function writeLocation(replenishSheet, replenishSheetInfo, wholesaleMap, replenishAsinRowColCoordinate, replenishLocationRowColCoordinate, replenishCostRowColCoordinate){
  var replenishAsinColIndex = replenishAsinRowColCoordinate.colIndex;
  var replenishLocationCol = replenishLocationRowColCoordinate.colIndex + 1;
  var replenishCostCol = replenishCostRowColCoordinate.colIndex + 1;
  var replenishSheetValues = replenishSheetInfo.sheetValues;
  
  // rowIndex has + 1 because we want to skip the header
  for(var rowIndex = replenishAsinRowColCoordinate.rowIndex + 1; rowIndex < replenishSheetInfo.amtRow; ++rowIndex){
    var asin = replenishSheetValues[rowIndex][replenishAsinColIndex];
    if(asin in wholesaleMap){
      var shelfLocation = wholesaleMap[asin].shelfLocation;
      var cost = wholesaleMap[asin].cost;
      replenishSheet.getRange(rowIndex+1, replenishLocationCol).setValue(shelfLocation);
      replenishSheet.getRange(rowIndex+1, replenishCostCol).setValue(cost);
    }
  }
}