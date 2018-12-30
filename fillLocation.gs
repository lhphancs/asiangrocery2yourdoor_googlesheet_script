function fillLocation(wholesaleSpreadSheet) {
  var writeHeaders = ['shelf location']; //These headers will have column written. Can add more to this array like 'product name'.
  
  var writeSheet = SpreadsheetApp.getActiveSheet();
  var writeSheetInfo = new SheetInfo(writeSheet);
  
  var writeKeyRowColCoordinate = getRowColCoordinateOfStr(writeSheetInfo, WHOLESALE_HEADER_ASIN);
  var dictOfWriteHeadersRowColCoordinate = getDictOfCoordinates(writeSheetInfo, writeHeaders);
  
  var errorMsgsContainer = new ErrorMsgsContainer();
  if( allHeadersAreFound(writeSheetInfo, dictOfWriteHeadersRowColCoordinate, errorMsgsContainer) ){
    var wholesaleDataDict = getDataDict(wholesaleSpreadSheet, WHOLESALE_HEADER_ASIN, writeHeaders, errorMsgsContainer);
    writeLocation(writeSheet, writeSheetInfo, wholesaleDataDict, writeKeyRowColCoordinate, dictOfWriteHeadersRowColCoordinate);
    var successMsg = "All location were written.";
    if(errorMsgsContainer.errorMgs.length > 0){
      successMsg += "\n\nWarnings:\n"
      successMsg += errorMsgsContainer.getErrorMsgs();
    }
  }
  else
    throw( "Undetected headers in sheet. No edits were made.\n\n" + errorMsgsContainer.getErrorMsgs() );
}

function addUndefinedHeaderErrors(sheetInfo, errorMsgsContainer, arrayOfHeadersNotFound){
  var headersNotFoundStr = "";
  for each(var header in arrayOfHeadersNotFound)
    headersNotFoundStr += header + ', ';
  headersNotFoundStr = headersNotFoundStr.substring(0, headersNotFoundStr.length - 2); // Remove ending space and comma
  
  var errorMsg = getSheetErrorString(sheetInfo.title, "These headers inside bracket were not found\n[  " + headersNotFoundStr + "  ]"); 
  errorMsgsContainer.addError(errorMsg);
}

function getDataFromRowIndex(sheetValues, rowIndex, dictOfValidHeadersRowColCoordinate){
  var data = {};
  for(var header in dictOfValidHeadersRowColCoordinate){
    var headerCoordinate = dictOfValidHeadersRowColCoordinate[header];
    data[header] = sheetValues[rowIndex][headerCoordinate.colIndex];
  }
  return data;
}

function readSheetValuesToCompleteDataDict(sheetInfo, dataDict, wholesaleHeaderAsinCoordinate, dictOfValidHeadersRowColCoordinate){
  var sheetValues = sheetInfo.sheetValues;
  
  for(i = wholesaleHeaderAsinCoordinate.rowIndex+1; i < sheetInfo.amtRow; ++i){
    var keyCellVal = sheetValues[i][wholesaleHeaderAsinCoordinate.colIndex];
    if(keyCellVal != "")
      dataDict[keyCellVal] = getDataFromRowIndex(sheetValues, i, dictOfValidHeadersRowColCoordinate);
  }
}

function getDictWithValidValuesOnly(dict, errorMsgsContainer, sheetTitle){
  var retDict = {};
  for(var key in dict)
    if(dict[key] == undefined)
      errorMsgsContainer.addError( getSheetErrorString(sheetTitle, "'" + key + "' was not found") );
    else
      retDict[key] = dict[key];
  return retDict;
}

function getDataDict(wholesaleSpreadSheet, WHOLESALE_HEADER_ASIN, writeHeaders, errorMsgsContainer){
  var dataDict = {};
  var sheets = wholesaleSpreadSheet.getSheets();
  
  for(var i = 0; i<sheets.length; ++i){
    var sheetInfo = new SheetInfo( sheets[i] );
    var wholesaleHeaderAsinCoordinate = getRowColCoordinateOfStr(sheetInfo, WHOLESALE_HEADER_ASIN);
    
    if(wholesaleHeaderAsinCoordinate == undefined)
      errorMsgsContainer.addError( getSheetErrorString(sheetInfo.title, WHOLESALE_HEADER_ASIN + "' was not found in sheet.") );
    else{
      var dictOfHeadersRowColCoordinate = getDictOfCoordinates(sheetInfo, writeHeaders);
      var dictOfValidHeadersRowColCoordinate = getDictWithValidValuesOnly(dictOfHeadersRowColCoordinate, errorMsgsContainer, sheetInfo.title);
      readSheetValuesToCompleteDataDict(sheetInfo, dataDict, wholesaleHeaderAsinCoordinate, dictOfValidHeadersRowColCoordinate);
    }
  }
  return dataDict;
}

function getValueFromDictWithKeyAndHeader(wholesaleDataDict, keyCellVal, header){
  if(header in wholesaleDataDict[keyCellVal]){
    return wholesaleDataDict[keyCellVal][header];
  }
  return undefined;
}

function writeLocation(writeSheet, writeSheetInfo, wholesaleDataDict, writeKeyRowColCoordinate, dictOfWriteHeadersRowColCoordinate){
  var writeKeyColIndex = writeKeyRowColCoordinate.colIndex;
  var sheetValues = writeSheetInfo.sheetValues;
  
  // rowIndex has + 1 because we want to skip the header
  for(var rowIndex = writeKeyRowColCoordinate.rowIndex + 1; rowIndex < writeSheetInfo.amtRow; ++rowIndex){
    var keyCellVal = sheetValues[rowIndex][writeKeyColIndex];
    if(keyCellVal in wholesaleDataDict){
      for(var header in dictOfWriteHeadersRowColCoordinate){
        var writeLocationCol = dictOfWriteHeadersRowColCoordinate[header].colIndex + 1;
        var writeVal = getValueFromDictWithKeyAndHeader(wholesaleDataDict, keyCellVal, header);
        writeSheet.getRange(rowIndex+1, writeLocationCol).setValue(writeVal);
      }
    }
  }
}

function getDictOfCoordinates(sheetInfo, strs){
  var dictOfCoordinates = {};
  for each(var str in strs)
    dictOfCoordinates[str] = getRowColCoordinateOfStr(sheetInfo, str);
  
  return dictOfCoordinates;
}

function getArrayOfUndefinedHeaders(dictOfHeadersRowColCoordinate){
  var arrayOfUndefinedHeaders = [];
  for(var header in dictOfHeadersRowColCoordinate){
    if(dictOfHeadersRowColCoordinate[header] == undefined)
      arrayOfUndefinedHeaders.push(header);
  }
  return arrayOfUndefinedHeaders;
}

function allHeadersAreFound(sheetInfo, dictOfWriteHeadersRowColCoordinate, errorMsgsContainer){
  var allHeadersAreFound = true;
  var arrayOfUndefinedHeaders = getArrayOfUndefinedHeaders(dictOfWriteHeadersRowColCoordinate);
      
  if(arrayOfUndefinedHeaders.length != 0){
    addUndefinedHeaderErrors(sheetInfo, errorMsgsContainer, arrayOfUndefinedHeaders);
    allHeadersAreFound = false;
  }
    
  return allHeadersAreFound;
}

function getSheetErrorString(sheetTitle, msg){
  return sheetTitle + ": " + msg;
}