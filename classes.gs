var ErrorMsgsContainer = function(){
  this.errorMgs = [];
  
  this.addError = function(msg){
    this.errorMgs.push(msg);
  }
  
  this.getErrorMsgs = function(){
    var msg = "";
    for(var i = 0; i < this.errorMgs.length; ++i)
      msg += this.errorMgs[i] + '\n';
    msg = msg.substring(0, msg.length - 1);
    return msg;
  }
}

var RowColCoordinate = function(row, col){
  this.rowIndex = row;
  this.colIndex = col;
};

var SheetInfo = function(sheet){
  this.title = sheet.getName();
  
  var rangeData = sheet.getDataRange();
  this.amtRow = rangeData.getLastRow();
  this.amtCol = rangeData.getLastColumn();
  this.sheetValues = sheet.getRange(1, 1, this.amtRow, this.amtCol).getValues(); //Retrives values as 2d array
}

var ReplenishHeaderCoordinates = function(asinCoord, unitSoldLast30DaysCoord, oosCoord){
  this.asin = asinCoord;
  this.unitSoldAmtLast30Days = unitSoldLast30DaysCoord;
  this.oos = oosCoord;
  
  this.hasAllCoordinates = asinCoord != undefined && unitSoldLast30DaysCoord != undefined && oosCoord != undefined;
}

var RepurchaseHeaderCoordinates = function(stockNoCoord, roundedRepurchaseAmtCoord, repurchaseAmtCoord, productNameCoord){
  this.stockNo = stockNoCoord;
  this.roundedRepurchaseAmt = roundedRepurchaseAmtCoord;
  this.repurchaseAmt = repurchaseAmtCoord;
  this.productName = productNameCoord;
  
  this.hasAllCoordinates = stockNoCoord != undefined && roundedRepurchaseAmtCoord != undefined
    && repurchaseAmtCoord != undefined && productNameCoord != undefined;
}

var WholesaleHeaders = function(asin, pack, boxAmt, stockNo, productName){
  this.asin = asin;
  this.pack = pack;
  this.boxAmt = boxAmt;
  this.stockNo = stockNo;
  this.productName = productName;
}