function fillUnsent(wholesaleSpreadSheet){
  var colorsToIgnore = [COLOR_RED_BKGD, COLOR_CYAN_BKGD];
  var asinsSent = getAsinsSent();
  
  var replenishSheet = SpreadsheetApp.getActiveSheet();
  var replenishSheetInfo = new SheetInfo(replenishSheet);

  var replenishHeaderAsinCoordinate = getRowColCoordinateOfStr(replenishSheetInfo, REPLENISH_HEADER_ASIN);
  var asinsCurrentlyOnReplenishSheet = getAsinsCurrentlyOnReplenishSheet(replenishSheetInfo, replenishHeaderAsinCoordinate);
  var itemsToWrite = getItemsToWrite(colorsToIgnore, wholesaleSpreadSheet, asinsSent, asinsCurrentlyOnReplenishSheet);

  writeItems(replenishSheet, replenishSheetInfo, itemsToWrite, replenishHeaderAsinCoordinate);
}

function getAsinsCurrentlyOnReplenishSheet(replenishSheetInfo, replenishHeaderAsinCoordinate){
  var asinsCurrentlyOnReplenishSheet = {};

  var sheetValues = replenishSheetInfo.sheetValues;
  for(var rowIndex = replenishHeaderAsinCoordinate.rowIndex + 1; rowIndex < replenishSheetInfo.amtRow; ++rowIndex){
      var asin = sheetValues[rowIndex][replenishHeaderAsinCoordinate.colIndex];
      if( !isBlankVal(asin) )
        asinsCurrentlyOnReplenishSheet[asin] = true;
  }
  return asinsCurrentlyOnReplenishSheet;
}

function writeItems(replenishSheet, replenishSheetInfo, itemsToWrite, replenishHeaderAsinCoordinate){
  var replenishHeaderProductNameCoordinate = getRowColCoordinateOfStr(replenishSheetInfo, REPLENISH_HEADER_PRODUCT_NAME);
  var replenishHeaderShelfLocationCoordinate = getRowColCoordinateOfStr(replenishSheetInfo, REPLENISH_HEADER_SHELF_LOCATION);
  var replenishHeaderMyCommentCoordinate = getRowColCoordinateOfStr(replenishSheetInfo, REPLENISH_HEADER_MY_COMMENT);
  var replenishHeaderOssCoordinate = getRowColCoordinateOfStr(replenishSheetInfo, REPLENISH_HEADER_OOS);
  
  var replenishAsinCol = replenishHeaderAsinCoordinate.colIndex + 1;
  var replenishProductNameCol = replenishHeaderProductNameCoordinate.colIndex + 1;
  var replenishShelfLocationCol = replenishHeaderShelfLocationCoordinate.colIndex + 1;
  var replenishMyCommentCol = replenishHeaderMyCommentCoordinate.colIndex + 1;
  var replenishOosCol = replenishHeaderOssCoordinate.colIndex + 1;
  
  var startWriteRow = replenishSheetInfo.amtRow + 1;
  var writeRow = startWriteRow;
  for(var asinKey in itemsToWrite){
    replenishSheet.getRange(writeRow, replenishAsinCol).setValue(asinKey);
    replenishSheet.getRange(writeRow, replenishProductNameCol).setValue(itemsToWrite[asinKey].productName);
    replenishSheet.getRange(writeRow, replenishShelfLocationCol).setValue(itemsToWrite[asinKey].shelfLocation);
    replenishSheet.getRange(writeRow, replenishMyCommentCol).setValue("Send amtFromHelium / (amtFbaSeller+1)");
    ++writeRow;
  }
  //Now color all written rows
  var amtRowColored = writeRow - startWriteRow;
  if(amtRowColored > 0)
    replenishSheet.getRange(startWriteRow, 1, amtRowColored, replenishOosCol).setBackground(COLOR_GREY_BKGD)
}

function getItemsToWrite(colorsToIgnore, wholesaleSpreadSheet, asinsSent, asinsCurrentlyOnReplenishSheet){
  var itemsToWrite = {};
  var sheets = wholesaleSpreadSheet.getSheets()
  for(var i = 0; i<sheets.length; ++i)
    readSheetValuesToCompleteItemsNotSent(colorsToIgnore, sheets[i], asinsSent, asinsCurrentlyOnReplenishSheet, itemsToWrite);
    
  return itemsToWrite;
}

function readSheetValuesToCompleteItemsNotSent(colorsToIgnore, sheet, asinsSent, asinsCurrentlyOnReplenishSheet, itemsToWrite){
  var sheetInfo = new SheetInfo(sheet);
  var sheetValues = sheetInfo.sheetValues;
  var wholesaleHeaderAsinCoordinate = getRowColCoordinateOfStr(sheetInfo, WHOLESALE_HEADER_ASIN);
  var wholesaleHeaderProductNameCoordinate = getRowColCoordinateOfStr(sheetInfo, WHOLESALE_HEADER_PRODUCT_NAME);
  var wholesaleHeaderLocationCoordinate = getRowColCoordinateOfStr(sheetInfo, WHOLESALE_HEADER_SHELF_LOCATION);
  
  if(wholesaleHeaderAsinCoordinate != undefined && wholesaleHeaderProductNameCoordinate != undefined && wholesaleHeaderLocationCoordinate != undefined){
    // rowIndex has + 1 because we want to skip the header
    for(var rowIndex = wholesaleHeaderAsinCoordinate.rowIndex + 1; rowIndex < sheetInfo.amtRow; ++rowIndex){
      var asin = sheetValues[rowIndex][wholesaleHeaderAsinCoordinate.colIndex];
      if( !isBlankVal(asin) && !(asin in asinsSent) && !(asin in asinsCurrentlyOnReplenishSheet)
      && !cellColorBkgdHasMatch(colorsToIgnore, sheet, rowIndex, wholesaleHeaderAsinCoordinate.colIndex) ){
        itemsToWrite[asin] = {
          productName: sheetValues[rowIndex][wholesaleHeaderProductNameCoordinate.colIndex]
          , shelfLocation: sheetValues[rowIndex][wholesaleHeaderLocationCoordinate.colIndex]};
      }
    }
  }
}

function getAsinsSent(){
  var asinsSent = {};
  
  var asinsSentSpreadSheet = SpreadsheetApp.openById(ASIN_SENT_LIST_SPREADSHEET_ID);
  var firstSheet = asinsSentSpreadSheet.getSheets()[0];
  var sheetInfo = new SheetInfo(firstSheet);
  var asinCoordinate = getRowColCoordinateOfStr(sheetInfo, ASIN_SENT_HEADER_ASIN);
  
  var sheetValues = sheetInfo.sheetValues;
  for(i = asinCoordinate.rowIndex+1; i < sheetInfo.amtRow; ++i){
    var keyCellVal = sheetValues[i][asinCoordinate.colIndex];
    if(keyCellVal != "")
      asinsSent[keyCellVal] = true;
  }
  return asinsSent;
}