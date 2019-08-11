function populateMapForSheet(map, sheet){
  var sheetInfo = new SheetInfo(sheet);
  var wholesaleAsinCoordinate = getRowColCoordinateOfStr(sheetInfo, WHOLESALE_HEADER_ASIN);
  var wholesaleProductNameCoordinate = getRowColCoordinateOfStr(sheetInfo, WHOLESALE_HEADER_PRODUCT_NAME);
  var wholesaleShelfLocationCoordinate = getRowColCoordinateOfStr(sheetInfo, WHOLESALE_HEADER_SHELF_LOCATION);
  var wholesaleCostCoordinate = getRowColCoordinateOfStr(sheetInfo, WHOLESALE_HEADER_COST);
  var wholesaleCommentCoordinate = getRowColCoordinateOfStr(sheetInfo, WHOLESALE_HEADER_IMPORTANT_COMMENT);
  
  if(wholesaleAsinCoordinate && wholesaleProductNameCoordinate && wholesaleShelfLocationCoordinate){
    var sheetValues = sheetInfo.sheetValues;
    for(var i = wholesaleAsinCoordinate.rowIndex+1; i < sheetInfo.amtRow; ++i){
      var row = sheetValues[i];
      var asin = sheetValues[i][wholesaleAsinCoordinate.colIndex];
      if( isBlankVal(asin) ){
        continue;
      }
        
      var productName = sheetValues[i][wholesaleProductNameCoordinate.colIndex];
      var shelfLocation = sheetValues[i][wholesaleShelfLocationCoordinate.colIndex];
      var cost = sheetValues[i][wholesaleCostCoordinate.colIndex];
      var comment = sheetValues[i][wholesaleCommentCoordinate.colIndex];
      var color = sheet.getRange(i+1, wholesaleAsinCoordinate.colIndex).getBackground(); //+1 because counting starts at 1 for getRange
      map[asin] = {productName: productName, shelfLocation: shelfLocation, comment: comment, cost: cost, color: color};
    }
  }
}

function getWholesaleMap(wholesaleSheet){
  var map = {};

  var sheets = wholesaleSheet.getSheets();
  for(var i = 0; i<sheets.length; ++i){
    populateMapForSheet(map, sheets[i]);
  }
  return map;
}