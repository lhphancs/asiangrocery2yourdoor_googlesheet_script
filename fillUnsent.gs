function fillUnsent(wholesaleMap){
  var colorsToIgnore = [COLOR_BKGD_RED, COLOR_BKGD_CYAN];
  var replenishSheet = SpreadsheetApp.getActiveSheet();
  var replenishSheetInfo = new SheetInfo(replenishSheet);
  var replenishHeaderAsinCoordinate = getRowColCoordinateOfStr(replenishSheetInfo, REPLENISH_HEADER_ASIN);
  
  var asinsAlreadySent = getAsinsSent();
  var replenishSheetAsins = getAsinsInReplenishSheet(replenishSheetInfo, replenishHeaderAsinCoordinate);
  var itemsToWrite = getItemsToWrite(colorsToIgnore, wholesaleMap, asinsAlreadySent, replenishSheetAsins);

  writeItems(replenishSheet, replenishSheetInfo, itemsToWrite, replenishHeaderAsinCoordinate);
}

function getAsinsInReplenishSheet(replenishSheetInfo, replenishHeaderAsinCoordinate){
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
  
  var startWriteRow = replenishSheetInfo.amtRow + 1;
  var writeRow = startWriteRow;
  for(var asinKey in itemsToWrite){
    replenishSheet.getRange(writeRow, replenishHeaderAsinCoordinate.colIndex + 1).setValue(asinKey);
    replenishSheet.getRange(writeRow, replenishHeaderProductNameCoordinate.colIndex + 1).setValue(itemsToWrite[asinKey].productName);
    replenishSheet.getRange(writeRow, replenishHeaderShelfLocationCoordinate.colIndex + 1).setValue(itemsToWrite[asinKey].shelfLocation);
    replenishSheet.getRange(writeRow, replenishHeaderMyCommentCoordinate.colIndex + 1).setValue("Send amtFromHelium / (amtFbaSeller+1)");
    ++writeRow;
  }
  //Now color all written rows
  var amtRowColored = writeRow - startWriteRow;
  if(amtRowColored > 0){
    replenishSheet.getRange(startWriteRow, 1, amtRowColored, replenishSheetInfo.sheetValues[0].length).setBackground(COLOR_BKGD_GREY);
  }
}

function getItemsToWrite(colorsToIgnore, wholesaleMap, asinsAlreadySent, replenishSheetAsins){
  var itemsToWrite = {};

  for(var wholesaleAsin in wholesaleMap){
    var wholesaleItem = wholesaleMap[wholesaleAsin];
    if( !(wholesaleAsin in replenishSheetAsins) && !cellColorBkgdHasMatch(wholesaleItem.color, colorsToIgnore) ){
      itemsToWrite[wholesaleAsin] = {productName: wholesaleItem.productName, shelfLocation: wholesaleItem.shelfLocation}
    }
  }
  return itemsToWrite;
}

function getAsinsSent(){
  var asinsSent = {};
  
  var asinsSentSpreadSheet = SpreadsheetApp.openById(ASIN_SENT_LIST_SPREADSHEET_ID);
  var firstSheet = asinsSentSpreadSheet.getSheets()[0];
  var sheetInfo = new SheetInfo(firstSheet);
  var asinCoordinate = getRowColCoordinateOfStr(sheetInfo, ASIN_SENT_HEADER_ASIN);
  
  var sheetValues = sheetInfo.sheetValues;
  for(var i = asinCoordinate.rowIndex+1; i < sheetInfo.amtRow; ++i){
    var asin = sheetValues[i][asinCoordinate.colIndex];
    if(asin != "")
      asinsSent[asin] = true;
  }
  return asinsSent;
}