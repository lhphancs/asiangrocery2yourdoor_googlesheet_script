function fillUnitsSoldLast30DaysIs0AndDaysSupplyIsMoreThanZero() {
  var replenishSheet = SpreadsheetApp.getActiveSheet();
  var replenishSheetInfo = new SheetInfo(replenishSheet);
  
  var replenishUnitsSoldLast30DaysRowColCoordinate = getRowColCoordinateOfStr(replenishSheetInfo, REPLENISH_HEADER_UNIT_SOLD_LAST_30_DAYS);
  var replenishTotalUnitsRowColCoordinate = getRowColCoordinateOfStr(replenishSheetInfo, REPLENISH_HEADER_TOTAL_UNITS);
  var replenishDaysOfSupplyRowColCoordinate = getRowColCoordinateOfStr(replenishSheetInfo, REPLENISH_HEADER_DAYS_OF_SUPPLY);
  var replenishCommentRowColCoordinate = getRowColCoordinateOfStr(replenishSheetInfo, REPLENISH_HEADER_MY_COMMENT);
   
  if(replenishUnitsSoldLast30DaysRowColCoordinate && replenishDaysOfSupplyRowColCoordinate){
    fillQualifiedRows(replenishSheet, replenishSheetInfo, replenishUnitsSoldLast30DaysRowColCoordinate, replenishTotalUnitsRowColCoordinate, replenishDaysOfSupplyRowColCoordinate, replenishCommentRowColCoordinate);
  }
  else
    throw( "Undetected headers in sheet. No edits were made.\n\n'" + REPLENISH_HEADER_ASIN + "' or '" + REPLENISH_HEADER_MY_COMMENT + "' was not found in replenish sheet." );
}

function fillQualifiedRows(replenishSheet, replenishSheetInfo, replenishUnitsSoldLast30DaysRowColCoordinate, replenishTotalUnitsRowColCoordinate, replenishDaysOfSupplyRowColCoordinate, replenishCommentRowColCoordinate){
  var replenishSheetValues = replenishSheetInfo.sheetValues;
  
  // rowIndex has + 1 because we want to skip the header
  for(var rowIndex = replenishUnitsSoldLast30DaysRowColCoordinate.rowIndex + 1; rowIndex < replenishSheetInfo.amtRow; ++rowIndex){
    var unitsSoldLast30Days = replenishSheetValues[rowIndex][replenishUnitsSoldLast30DaysRowColCoordinate.colIndex];
    var totalUnits = replenishSheetValues[rowIndex][replenishTotalUnitsRowColCoordinate.colIndex];
    var daysOfSupply = replenishSheetValues[rowIndex][replenishDaysOfSupplyRowColCoordinate.colIndex];
    
    if(unitsSoldLast30Days === 0 && totalUnits > 0){
      replenishSheet.getRange(rowIndex+1, replenishCommentRowColCoordinate.colIndex + 1).setValue('Check Repricer');
      replenishSheet.getRange(rowIndex+1, 1, 1, replenishSheetValues[0].length).setBackground(COLOR_PURPLE_BKGD);
    }
    else if(unitsSoldLast30Days > 0 && daysOfSupply === 0){
      var sendAmt = unitsSoldLast30Days < 5 && daysOfSupply === 0 ? 10 : unitsSoldLast30Days * 2;
      replenishSheet.getRange(rowIndex+1, replenishCommentRowColCoordinate.colIndex + 1).setValue('"date" send ' + sendAmt);
    }
  }
}