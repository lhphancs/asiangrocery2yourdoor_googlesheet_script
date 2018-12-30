function onEditCheckHeadersAndRespond(e){
  var replenishSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var replenishSheetInfo = new SheetInfo(replenishSheet);
  var replenishHeaderCoordinatesObj = getReplenishHeaderCoordinatesObj(replenishSheetInfo);
  if(replenishHeaderCoordinatesObj.hasAllCoordinates)
    onEditCheckEditedCellAndRespond(e, replenishSheet, replenishSheetInfo, replenishHeaderCoordinatesObj);
  else
    throw("Error: All replenish headers were not found. Check script's global vars");
}

function onEditCheckEditedCellAndRespond(e, replenishSheet, replenishSheetInfo, replenishHeaderCoordinatesObj){
  var range = e.range;
  var editCoordinate = new RowColCoordinate(range.getRow()-1, range.getColumn()-1);

  if(editCoordinate.colIndex == replenishHeaderCoordinatesObj.oos.colIndex
  || editCoordinate.colIndex == replenishHeaderCoordinatesObj.asinListAddOrDelete.colIndex){
    var wholesaleSpreadSheet = SpreadsheetApp.openById(READ_WHOLESALE_SPREADSHEET_ID);
    var wholesaleHeadersObj = new WholesaleHeaders(WHOLESALE_HEADER_ASIN, WHOLESALE_HEADER_PACK
                          , WHOLESALE_HEADER_BOX_AMT, WHOLESALE_HEADER_STOCK_NO, WHOLESALE_HEADER_PRODUCT_NAME);
    if(editCoordinate.rowIndex > replenishHeaderCoordinatesObj.oos.rowIndex){
      var replenishSheetValues = replenishSheetInfo.sheetValues;
      var editedCellVal = replenishSheetValues[editCoordinate.rowIndex][editCoordinate.colIndex];
      if(editCoordinate.colIndex == replenishHeaderCoordinatesObj.oos.colIndex){
        handleOosEditedCellVal(wholesaleSpreadSheet, wholesaleHeadersObj, replenishSheet
          , replenishSheetValues, replenishHeaderCoordinatesObj, editCoordinate, editedCellVal);
      }
      else{
        handleAsinListEditedCellVal(replenishSheet
          , replenishSheetValues, replenishHeaderCoordinatesObj, editCoordinate, editedCellVal);
      }
    }
  }
}

function getReplenishHeaderCoordinatesObj(activeSheetInfo){
  var asinCoord = getRowColCoordinateOfStr(activeSheetInfo, REPLENISH_HEADER_ASIN);
  var unitSoldAmtLast30DaysCoord = getRowColCoordinateOfStr(activeSheetInfo, REPLENISH_HEADER_UNIT_SOLD_LAST_30_DAYS);
  var oosCoord = getRowColCoordinateOfStr(activeSheetInfo, REPLENISH_HEADER_OOS);
  var asinListAddOrDeleteCoord = getRowColCoordinateOfStr(activeSheetInfo, REPLENISH_HEADER_ASIN_LIST_ADD_OR_DELETE);
  return new ReplenishHeaderCoordinates(asinCoord, unitSoldAmtLast30DaysCoord, oosCoord, asinListAddOrDeleteCoord);
}