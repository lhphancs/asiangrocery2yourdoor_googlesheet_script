function handleOosEditedCellVal(wholesaleSpreadSheet, wholesaleHeadersObj, replenishSheet
, replenishSheetValues, replenishHeaderCoordinatesObj, editCoordinate, editedCellVal){
  var isAddGreenMode = isSameWord(editedCellVal, REPLENISH_OOS_OPTION_ADD_GREEN_STR);
  var isAddYellowMode = isSameWord(editedCellVal, REPLENISH_OOS_OPTION_ADD_YELLOW_STR);
  var isSubtractMode = isSameWord(editedCellVal, REPLENISH_OOS_OPTION_SUBTRACT_STR);

  if(isAddGreenMode || isAddYellowMode || isSubtractMode){
    displayMsgScriptRunning();
    var color = isAddGreenMode ? COLOR_GREEN_BKGD : isAddYellowMode ? COLOR_YELLOW_BKGD : null;
    if(isAddGreenMode || isAddYellowMode)
      extractRowDataAndExecuteRepurchaseWrite(wholesaleSpreadSheet, wholesaleHeadersObj
      , replenishSheetValues, replenishHeaderCoordinatesObj, editCoordinate, true);
      
    else
      extractRowDataAndExecuteRepurchaseWrite(wholesaleSpreadSheet, wholesaleHeadersObj
      , replenishSheetValues, replenishHeaderCoordinatesObj, editCoordinate, false);
    replenishSheet.getRange(editCoordinate.rowIndex+1, 1
    , 1, replenishHeaderCoordinatesObj.asinListAddOrDelete.colIndex+1).setBackground(color);
  }
}

function extractRowDataAndExecuteRepurchaseWrite(wholesaleSpreadSheet, wholesaleHeadersObj, replenishSheetValues
, replenishHeaderCoordinatesObj, editCoordinate, isAddMode){
  var asin = replenishSheetValues[editCoordinate.rowIndex][replenishHeaderCoordinatesObj.asin.colIndex];
  var unitSoldAmtLast30Days = replenishSheetValues[editCoordinate.rowIndex][replenishHeaderCoordinatesObj.unitSoldAmtLast30Days.colIndex];
  unitSoldAmtLast30Days = unitSoldAmtLast30Days == 0 ? 1 : unitSoldAmtLast30Days; //If 0, pretend it's 1. Makes calculation not equal 0.

  if( isBlankVal(asin) )
    throw("Blank asin for current row");
  else{
    var repurchaseSpreadSheet = SpreadsheetApp.openById(REPURCHASE_SPREADSHEET_ID);
    var wholesaleData = getWholesaleDataForAsin(wholesaleSpreadSheet, wholesaleHeadersObj, asin);
    var errorMsgsContainer = new ErrorMsgsContainer();
    checkValidCalculation(wholesaleData, unitSoldAmtLast30Days, errorMsgsContainer);

    if( errorMsgsContainer.errorMgs.length == 0 ){
      var repurchaseBoxAmt = calculateBoxAmt(wholesaleData, unitSoldAmtLast30Days);
      var repurchaseAddAmt = isAddMode ? repurchaseBoxAmt : -repurchaseBoxAmt;
      
      executeRepurchaseWrite(wholesaleData, repurchaseSpreadSheet, repurchaseAddAmt);
    }
    else{
      var errorMsg = errorMsgsContainer.getErrorMsgs();
      writeErrorToRepurchaseSheet(repurchaseSpreadSheet, asin, errorMsg);
      throw(errorMsg);
    }
  }
}

function checkValidCalculation(wholesaleData, unitSoldAmtLast30Days, errorMsgsContainer){
  if(wholesaleData.wholesalers.length == 0){
    errorMsgsContainer.addError("No wholesaler found.");
    return;
  }
    
  if(typeof wholesaleData.pack != 'number')
    errorMsgsContainer.addError("Pack is not a number on wholesale sheet");
    
  if(typeof wholesaleData.boxAmt != 'number')
    errorMsgsContainer.addError("Box Amount is not a number on wholesale sheet");
    
  if(typeof unitSoldAmtLast30Days != 'number')
    errorMsgsContainer.addError("unitSoldAmtLast30Days is not a number on replenish sheet");
}

function getWholesaleDataForAsin(wholesaleSpreadSheet, wholesaleHeadersObj, asin){
  var wholesaleData = {};
  wholesaleData.wholesalers = [];
  var sheets = wholesaleSpreadSheet.getSheets();

  for(var i = 0; i<sheets.length; ++i){
    var sheetName = sheets[i].getName();
    var sheetInfo = new SheetInfo( sheets[i] );
    var headerAsinCoordinate = getRowColCoordinateOfStr(sheetInfo, wholesaleHeadersObj.asin);
    var headerPackCoordinate = getRowColCoordinateOfStr(sheetInfo, wholesaleHeadersObj.pack);
    var headerBoxAmtCoordinate = getRowColCoordinateOfStr(sheetInfo, wholesaleHeadersObj.boxAmt);
    var headerStockNoCoordinate = getRowColCoordinateOfStr(sheetInfo, wholesaleHeadersObj.stockNo);
    var headerProductNameCoordinate = getRowColCoordinateOfStr(sheetInfo, wholesaleHeadersObj.productName);
    
    var listOfHeadersToCheck = [headerAsinCoordinate, headerPackCoordinate
                                , headerBoxAmtCoordinate, headerStockNoCoordinate, headerProductNameCoordinate];
    var allHeadersFound = true;
    for(var j=0; j<listOfHeadersToCheck.length; ++j){
      Logger.log(i);
      if(listOfHeadersToCheck[j] == undefined){
        allHeadersFound = false;
        break;
      }
    }
    if(allHeadersFound)
      addWholesaleDataIfAsinFound(wholesaleData, sheetName, sheetInfo, headerAsinCoordinate
      , headerPackCoordinate, headerBoxAmtCoordinate, headerStockNoCoordinate, headerProductNameCoordinate, asin);
  }
  return wholesaleData;
}

function addWholesaleDataIfAsinFound(wholesaleData, sheetName, sheetInfo, headerAsinCoordinate
, headerPackCoordinate, headerBoxAmtCoordinate, headerStockNoCoordinate, headerProductNameCoordinate, asin){
Logger.log("AA");
  Logger.log(sheetName);
  Logger.log(headerAsinCoordinate)
  Logger.log("BB");
  
  var sheetValues = sheetInfo.sheetValues;

  for(var i = headerAsinCoordinate.rowIndex+1; i<sheetInfo.amtRow; ++i){
    var asinCellVal = sheetValues[i][headerAsinCoordinate.colIndex];
    if(asinCellVal == asin){
      wholesaleData.wholesalers.push(sheetName);
      
      wholesaleData.pack = sheetValues[i][headerPackCoordinate.colIndex];
      wholesaleData.boxAmt = sheetValues[i][headerBoxAmtCoordinate.colIndex];
      wholesaleData.stockNo = sheetValues[i][headerStockNoCoordinate.colIndex];
      wholesaleData.productName = sheetValues[i][headerProductNameCoordinate.colIndex];
      return;
    }
  }
}

function calculateBoxAmt(wholesaleData, unitSoldAmtLast30Days){
   return 2 * (unitSoldAmtLast30Days * wholesaleData.pack / wholesaleData.boxAmt);
}

function executeRepurchaseWrite(wholesaleData, repurchaseSpreadSheet, repurchaseAddAmt){
  var lastUpdatedRepurchaseAmt = undefined;
  for(var i=0; i<wholesaleData.wholesalers.length; ++i){
    var wholesaler = wholesaleData.wholesalers[i];
    var repurchaseWholesalerSheet = getOrCreateWholesalerSheet(repurchaseSpreadSheet, wholesaler);
    var repurchaseSheetInfo = new SheetInfo(repurchaseWholesalerSheet);

    var repurchaseHeaderCoordinates = new RepurchaseHeaderCoordinates(
      getRowColCoordinateOfStr(repurchaseSheetInfo, REPURCHASE_HEADER_STOCK_NO),
      getRowColCoordinateOfStr(repurchaseSheetInfo, REPURCHASE_HEADER_ROUNDED_REPURCHASE_AMT),
      getRowColCoordinateOfStr(repurchaseSheetInfo, REPURCHASE_HEADER_REPURCHASE_AMT),
      getRowColCoordinateOfStr(repurchaseSheetInfo, REPURCHASE_HEADER_PRODUCT_NAME)
    );
    
    if(!repurchaseHeaderCoordinates.hasAllCoordinates){
      var errorMsg = "Repurchase's (" + wholesaler + "): Has headers that are not found."
      writeErrorToRepurchaseSheet(repurchaseSpreadSheet, "N/A", errorMsg)
      throw(errorMsg);
    }
    
    var dictStockNoToRepurchaseAmt = getExistingWholesalerDictStockNoToRepurchaseAmt(repurchaseSheetInfo, repurchaseHeaderCoordinates);
    if(wholesaleData.stockNo in dictStockNoToRepurchaseAmt)
      dictStockNoToRepurchaseAmt[wholesaleData.stockNo].repurchaseAmt += repurchaseAddAmt;
    else
      dictStockNoToRepurchaseAmt[wholesaleData.stockNo] = {
        productName: wholesaleData.productName,
        repurchaseAmt: repurchaseAddAmt
      };
    
    
    //Don't allow number to be negative
    if(dictStockNoToRepurchaseAmt[wholesaleData.stockNo].repurchaseAmt < 0)
      dictStockNoToRepurchaseAmt[wholesaleData.stockNo].repurchaseAmt = 0;
      
    clearSheetAndRewrite(repurchaseWholesalerSheet, repurchaseSheetInfo, repurchaseHeaderCoordinates, dictStockNoToRepurchaseAmt);
    lastUpdatedRepurchaseAmt = dictStockNoToRepurchaseAmt[wholesaleData.stockNo].repurchaseAmt;
  }
  if(lastUpdatedRepurchaseAmt != undefined)
    displayMsg(wholesaleData.stockNo + " (" + lastUpdatedRepurchaseAmt + ")", "Update successful");
}

function getOrCreateWholesalerSheet(repurchaseSpreadSheet, wholesaler){
  var wholesalerSheet = repurchaseSpreadSheet.getSheetByName(wholesaler);
  if(wholesalerSheet == null){
    wholesalerSheet = repurchaseSpreadSheet.insertSheet(wholesaler);
    wholesalerSheet.getRange(1, REPURCHASE_DEFAULT_COL_STOCK_NO).setValue(REPURCHASE_HEADER_STOCK_NO);
    wholesalerSheet.getRange(1, REPURCHASE_DEFAULT_COL_ROUNDED_REPURCHASE_AMT).setValue(REPURCHASE_HEADER_ROUNDED_REPURCHASE_AMT);
    wholesalerSheet.getRange(1, REPURCHASE_DEFAULT_COL_REPURCHASE_AMT).setValue(REPURCHASE_HEADER_REPURCHASE_AMT);
    wholesalerSheet.getRange(1, REPURCHASE_DEFAULT_COL_PRODUCT_NAME).setValue(REPURCHASE_HEADER_PRODUCT_NAME);
  }
  return wholesalerSheet;
}

function getExistingWholesalerDictStockNoToRepurchaseAmt(repurchaseSheetInfo, repurchaseHeaderCoordinates){
  var sheetValues = repurchaseSheetInfo.sheetValues;
  var dict = {};

  for(var i=repurchaseHeaderCoordinates.stockNo.rowIndex+1; i<repurchaseSheetInfo.amtRow; ++i){    
    var stockNoCellVal = sheetValues[i][repurchaseHeaderCoordinates.stockNo.colIndex];
    var repurchaseAmtCellVal = sheetValues[i][repurchaseHeaderCoordinates.repurchaseAmt.colIndex];
    var productNameCellVal = sheetValues[i][repurchaseHeaderCoordinates.productName.colIndex];
    if(stockNoCellVal == '' || stockNoCellVal == undefined || repurchaseAmtCellVal == '' || repurchaseAmtCellVal == undefined)
      continue;
    dict[stockNoCellVal] = {
      repurchaseAmt: repurchaseAmtCellVal,
      productName: productNameCellVal
    };
  }
  return dict;
}

function clearSheetAndRewrite(repurchaseWholesalerSheet, repurchaseSheetInfo, repurchaseHeaderCoordinates, dictStockNoToRepurchaseAmt){
  var rowNumberBelowHeader = repurchaseHeaderCoordinates.stockNo.rowIndex + 2;
  var stockNoCol = repurchaseHeaderCoordinates.stockNo.colIndex + 1;
  var roundedRepurchaseAmtCol = repurchaseHeaderCoordinates.roundedRepurchaseAmt.colIndex + 1;
  var repurchaseAmtCol = repurchaseHeaderCoordinates.repurchaseAmt.colIndex + 1;
  var productNameCol = repurchaseHeaderCoordinates.productName.colIndex + 1;
  repurchaseWholesalerSheet.deleteRows( rowNumberBelowHeader, repurchaseWholesalerSheet.getLastRow() ); // +1 to convert to row number, and +1 skip headeRow
  
  var i = rowNumberBelowHeader;
  var formattedDate = Utilities.formatDate(new Date(), "PST", "MM/dd/yy HH:mm:ss");
  for(var key in dictStockNoToRepurchaseAmt){
    var repurchaseAmt = dictStockNoToRepurchaseAmt[key].repurchaseAmt;
    if(repurchaseAmt <= 0)
      continue;
    var roundedRepurchaseAmt = repurchaseAmt > 0 && repurchaseAmt <= 1 ? 1 : roundPositiveDecimalToOnes(repurchaseAmt, 4);
    var productName = dictStockNoToRepurchaseAmt[key].productName;
    repurchaseWholesalerSheet.getRange(i, stockNoCol).setValue(key);
    repurchaseWholesalerSheet.getRange(i, roundedRepurchaseAmtCol).setValue(roundedRepurchaseAmt);
    repurchaseWholesalerSheet.getRange(i, repurchaseAmtCol).setValue(repurchaseAmt);
    repurchaseWholesalerSheet.getRange(i, productNameCol).setValue(productName);
    ++i;
  }
}

function writeErrorToRepurchaseSheet(repurchaseSpreadSheet, asin, errorMsg){
  var replenishSheetName = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName();
  var errorSheet = getOrCreateRepurchaseErrorSheet(repurchaseSpreadSheet);
  var rowAfterLast = errorSheet.getLastRow() + 1;
  var formattedDate = Utilities.formatDate(new Date(), "PST", "MM/dd/yy HH:mm:ss");
  
  var errorSheetInfo = new SheetInfo(errorSheet);
  var errorReplenishSheetNameCoordinate = getRowColCoordinateOfStr(errorSheetInfo, ERROR_HEADER_REPLENISH_SHEETNAME);
  var errorAsinCoordinate = getRowColCoordinateOfStr(errorSheetInfo, ERROR_HEADER_ASIN);
  var errorErrorMsgCoordinate = getRowColCoordinateOfStr(errorSheetInfo, ERROR_HEADER_ERROR_MSG);
  var errorTimestampCoordinate = getRowColCoordinateOfStr(errorSheetInfo, ERROR_HEADER_TIMESTAMP);
  
  errorSheet.getRange(rowAfterLast, errorReplenishSheetNameCoordinate.colIndex+1).setValue(replenishSheetName);
  errorSheet.getRange(rowAfterLast, errorAsinCoordinate.colIndex+1).setValue(asin);
  errorSheet.getRange(rowAfterLast, errorErrorMsgCoordinate.colIndex+1).setValue(errorMsg);
  errorSheet.getRange(rowAfterLast, errorTimestampCoordinate.colIndex+1).setValue(formattedDate);
}

function getOrCreateRepurchaseErrorSheet(repurchaseSpreadSheet){
  var errorSheet = repurchaseSpreadSheet.getSheetByName(ERROR_SHEETNAME);
  if(errorSheet == null){
    errorSheet = repurchaseSpreadSheet.insertSheet(ERROR_SHEETNAME);
    errorSheet.getRange(1, ERROR_DEFAULT_COL_REPLENISH_SHEETNAME).setValue(ERROR_HEADER_REPLENISH_SHEETNAME);
    errorSheet.getRange(1, ERROR_DEFAULT_COL_ASIN).setValue(ERROR_HEADER_ASIN);
    errorSheet.getRange(1, ERROR_DEFAULT_COL_ERROR_MSG).setValue(ERROR_HEADER_ERROR_MSG);
    errorSheet.getRange(1, ERROR_DEFAULT_COL_TIMESTAMP).setValue(ERROR_HEADER_TIMESTAMP);
  }
  return errorSheet;
}