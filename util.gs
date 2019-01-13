function displayMsg(msg, title){
  SpreadsheetApp.getActiveSpreadsheet().toast(msg, title, -1);
}

function displayMsgScriptRunning(){
  displayMsg('Script is currently running...', 'Processing');
}

function getRowColCoordinateOfStr(sheetInfo, str){
  for(i = 0; i < sheetInfo.amtRow; ++i){
    for(j = 0; j < sheetInfo.amtCol; ++j){
      var cellVal = sheetInfo.sheetValues[i][j];
      if( typeof(cellVal) != 'string' )
        continue;
      
      if(cellVal.toUpperCase() == str.toUpperCase() )
        return new RowColCoordinate(i, j);
    }
  }
  return undefined;
}


function roundPositiveDecimalToOnes(val, lowestIntToRoundUp){
  var tenthPlaceDigit = (val*10)%10;
  var floorVal = Math.floor(val);
  return tenthPlaceDigit >= lowestIntToRoundUp ? floorVal+1 : floorVal;
}

function isBlankVal(val){
  return val == '' || val == undefined;
}

function isSameWord(word1, word2){
  return word1.toUpperCase() == word2.toUpperCase();
}

function cellColorBkgdHasMatch(colorStrArray, sheet, rowIndex, colIndex){
  var range = sheet.getRange(rowIndex+1, colIndex+1);
  var cellBackgroundColor = range.getBackground();

  for(var i=0; i<colorStrArray.length; ++i)
    if(colorStrArray[i] == cellBackgroundColor)
      return true;
      
  return false;
}