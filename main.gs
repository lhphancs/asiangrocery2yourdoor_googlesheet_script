/*           
Coder Note:
  Anything that has "index" in variable name means counting begins at 0.
  This is because googlesheet reads and write cells beginning with value of 1...
  But when their function returns 2d array, reading that array must have counting begin at 0
*/

function executeScript(scriptFunction, e){
  try{
    scriptFunction(e);
  }
  catch(error){
    // Need this check because script runs even if they type in non OOS col, 
    if(error != undefined)
      displayMsg(error, "Error");
  }
}

function executeFillLocation(e){
  var wholesaleSpreadSheet = SpreadsheetApp.openById(READ_WHOLESALE_SPREADSHEET_ID);
  fillLocation(wholesaleSpreadSheet, WHOLESALE_HEADER_ASIN);
}

function executeFillRepurchase(e){
  var wholesaleSpreadSheet = SpreadsheetApp.openById(READ_WHOLESALE_SPREADSHEET_ID);
  var wholesaleHeadersObj = new WholesaleHeaders(WHOLESALE_HEADER_ASIN, WHOLESALE_HEADER_PACK
                          , WHOLESALE_HEADER_BOX_AMT, WHOLESALE_HEADER_STOCK_NO, WHOLESALE_HEADER_PRODUCT_NAME);
  fillRepurchase(e, wholesaleSpreadSheet, wholesaleHeadersObj);
}

function executeFillWhiteRowsBtn(e){
  displayMsgScriptRunning();
  fillWhiteRowsBtn();
}


// Note: 'onEdit' is a reserve function for googlesheet script. Can't use onEdit directly due to permission error.
function onCellEdit(e){
  executeScript(executeFillRepurchase, e);
}

function onClickFillLocationBtn(e){
  displayMsgScriptRunning();
  executeScript(executeFillLocation, e);
  displayMsg("Script successfully wrote locations", "Update Complete");
}

function onClickFillWhiteRowsBtn(e){
  executeScript(executeFillWhiteRowsBtn, e);
  displayMsg("Script successfully wrote white rows.", "Update Complete");
}