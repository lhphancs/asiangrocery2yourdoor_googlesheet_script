function fillRedRows(wholesaleSpreadSheet){
  var redAsinSet = getRedAsinSet(wholesaleSpreadSheet);
  
  var replenishSheet = SpreadsheetApp.getActiveSheet();
  var replenishSheetInfo = new SheetInfo(replenishSheet);
  var replenishHeaderAsinCoordinate = getRowColCoordinateOfStr(replenishSheetInfo, REPLENISH_HEADER_ASIN);
  var replenishHeaderOssCoordinate = getRowColCoordinateOfStr(replenishSheetInfo, REPLENISH_HEADER_OOS);
  var replenishHeaderMyCommentCoordinate = getRowColCoordinateOfStr(replenishSheetInfo, REPLENISH_HEADER_MY_COMMENT);
  colorRedAsinRows(redAsinSet, replenishSheet, replenishSheetInfo, replenishHeaderAsinCoordinate, replenishHeaderOssCoordinate, replenishHeaderMyCommentCoordinate);
}

function colorRedAsinRows(redAsinSet, replenishSheet, replenishSheetInfo, replenishHeaderAsinCoordinate, replenishHeaderOssCoordinate, replenishHeaderMyCommentCoordinate){
  var replenishSheetValues = replenishSheetInfo.sheetValues;
  for(i = replenishHeaderAsinCoordinate.rowIndex+1; i < replenishSheetInfo.amtRow; ++i){
    var asinCellVal = replenishSheetValues[i][replenishHeaderAsinCoordinate.colIndex];
    if(asinCellVal in redAsinSet){
      replenishSheet.getRange(i+1, 1, 1, replenishHeaderOssCoordinate.colIndex+1).setBackground(COLOR_RED_BKGD);
      replenishSheet.getRange(i+1, replenishHeaderMyCommentCoordinate.colIndex+1).setValue("Check profit before send");
    }
  }
}

function getRedAsinSet(wholesaleSpreadSheet){
  var redAsinSet = {};
  
  var sheets = wholesaleSpreadSheet.getSheets();
  for(var i = 0; i<sheets.length; ++i){
    var sheetInfo = new SheetInfo( sheets[i] );
    var wholesaleHeaderAsinCoordinate = getRowColCoordinateOfStr(sheetInfo, WHOLESALE_HEADER_ASIN);
    
    if(wholesaleHeaderAsinCoordinate != undefined)
      readSheetValuesToAddAsins(sheets[i], sheetInfo, redAsinSet, wholesaleHeaderAsinCoordinate);
  }
  return redAsinSet;
}

function readSheetValuesToAddAsins(sheet, sheetInfo, redAsinSet, wholesaleHeaderAsinCoordinate){
  var sheetValues = sheetInfo.sheetValues;
  var wholesaleHeaderAsinCol = wholesaleHeaderAsinCoordinate.colIndex + 1;
  
  for(i = wholesaleHeaderAsinCoordinate.rowIndex+1; i < sheetInfo.amtRow; ++i){
    var keyCellVal = sheetValues[i][wholesaleHeaderAsinCoordinate.colIndex];
    if(keyCellVal != "" && !(keyCellVal in redAsinSet) ){
      var color = sheet.getRange(i+1, wholesaleHeaderAsinCol).getBackground();
      if(color == COLOR_RED_BKGD)
        redAsinSet[keyCellVal] = true;
    }
  }
}
