function handleAsinListEditedCellVal(replenishSheet
        , replenishSheetValues, replenishHeaderCoordinatesObj, editCoordinate, editedCellVal) {
  var isAsinAddMode = isSameWord(editedCellVal, REPLENISH_SENT_ASIN_OPTION_ADD_STR);
  var isAsinDeleteMode = isSameWord(editedCellVal, REPLENISH_SENT_ASIN_OPTION_SUBTRACT_STR);
  var asin = replenishSheetValues[editCoordinate.rowIndex][replenishHeaderCoordinatesObj.asin.colIndex];
  
  if( isBlankVal(asin) )
    throw("Blank asin for current row");
    
  if(isAsinAddMode || isAsinDeleteMode){
    displayMsgScriptRunning();
    
    var color = isAsinAddMode ? COLOR_GREEN_STR : COLOR_GREY_STR;
    var asinListSheet = SpreadsheetApp.openById(ASIN_SENT_LIST_SPREADSHEET_ID).getSheets()[0];
    var asinSheetInfo = new SheetInfo(asinListSheet);
    var asinSheetValues = asinSheetInfo.sheetValues;
    var asinCoordinate = getRowColCoordinateOfStr(asinSheetInfo, ASIN_SENT_HEADER_ASIN);
    if(isAsinAddMode)
      insertAsinInAsinListSheet(asinListSheet, asinSheetInfo, asinSheetValues, asinCoordinate, asin);
      
    else
      deleteAsinInAsinListSheet(asinListSheet, asinSheetInfo, asinSheetValues, asinCoordinate, asin);
      
    replenishSheet.getRange(editCoordinate.rowIndex+1, 1, 1, editCoordinate.colIndex+1).setBackground(color);
  }
}

function insertAsinInAsinListSheet(asinListSheet, asinSheetInfo, asinSheetValues, asinCoordinate, asin){
  //Check if asin is already in there
  for(var i = asinCoordinate.rowIndex + 1; i<asinSheetInfo.amtRow; ++i){
    if( asinSheetValues[i][asinCoordinate.colIndex] == asin )
      throw(asin + " already exists in the asin list. No changes made.");
  }
  asinListSheet.getRange(asinSheetInfo.amtRow+1, asinCoordinate.colIndex+1).setValue(asin);
  displayMsg("Asin list added: " + asin + " to ", "Update successful");
}

function deleteAsinInAsinListSheet(asinListSheet, asinSheetInfo, asinSheetValues, asinCoordinate, asin){
  var asinFound = false;
  var amtAsinFound = 0;
  for(var i = asinCoordinate.rowIndex + 1; i<asinSheetInfo.amtRow; ++i){
    if( asinSheetValues[i][asinCoordinate.colIndex] == asin ){
      asinListSheet.deleteRow(i + 1 - amtAsinFound);
      ++amtAsinFound;
      asinFound = true;
    }
  }
  if(asinFound)
    displayMsg("Asin list deleted: " + asin, "Update successful");
  else
    throw(asin + " was not found in Asin list");
}