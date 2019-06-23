/*           
Coder Note:
  Anything that has "index" in variable name means counting begins at 0.
  This is because googlesheet reads and write cells beginning with value of 1...
  But when their function returns 2d array, reading that array must have counting begin at 0
*/

function executeScript(scriptFunction, e){
  var lock = LockService.getDocumentLock();
  try{
    while( !lock.tryLock(3000) ){
      continue;
    }
    try{
      scriptFunction(e);
    }
    catch(error){
       displayMsg(error, "Error");
    }
      
  }
  catch(error){
    displayMsg(error, "Error");
  }
  finally{
    lock.releaseLock();
  }
}

function executeFillRepurchase(e){
  var wholesaleSpreadSheet = SpreadsheetApp.openById(READ_WHOLESALE_SPREADSHEET_ID);
  fillRepurchase(wholesaleSpreadSheet, e);
}

function fillAll(e){
  var wholesaleSpreadSheet = SpreadsheetApp.openById(READ_WHOLESALE_SPREADSHEET_ID);
  var wholesaleMap = getWholesaleMap(wholesaleSpreadSheet);
  
  fillLocation(wholesaleMap);
  fillUnitsSoldLast30DaysIs0AndDaysSupplyIsMoreThanZero();
  fillRedRows(wholesaleMap);
  fillUnsent(wholesaleMap);
  
  displayMsg("Scripts ran successfully!", "Update Complete");
}

function onBtnFillWhiteRows(e){
  displayMsgScriptRunning();
  executeScript(fillWhiteRows, e);
}

function onRunAllScriptsBtn(e){
  displayMsgScriptRunning();
  executeScript(fillAll, e);
}


// Note: 'onEdit' is a reserve function for googlesheet script. Can't use onEdit directly due to permission error.
function onCellEdit(e){
  executeScript(onEditCheckHeadersAndRespond, e);
}