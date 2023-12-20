// Author: Dai Hung PHAM
// Date: May 24, 2023
//
// Important Notice:
// This code is the intellectual property of Dai Hung PHAM. (daihung.pham@yahoo.fr)
// Do not reuse or share this code without explicit permission.
// If you wish to use or modify this code, please contact Dai Hung PHAM for authorization.
//getColumnLetter_////////////////////////////////////////////////////////////////////////////////////
/**

*/
function getColumnLetter_(columnNumber) {
  var columnLetter = "";
  while (columnNumber > 0) {
    var remainder = (columnNumber - 1) % 26;
    columnLetter = String.fromCharCode(65 + remainder) + columnLetter;
    columnNumber = Math.floor((columnNumber - 1) / 26);
  }
  return columnLetter;
}


//getColumnNumber_////////////////////////////////////////////////////////////////////////////////////
/**

*/
function getColumnNumber_(columnLetter) {
  var column = 0;
  for (var i = 0; i < columnLetter.length; i++) {
    column *= 26;
    column += columnLetter.charCodeAt(i) - 'A'.charCodeAt(0) + 1;
  }
  return column;
}


//getIdProductbycellReference////////////////////////////////////////////////////////////////////////////////////
/**

*/
function getIdProductbycellReference_(cellReference) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PRODUCTS_Sheet.Name);
  var columnLetter = cellReference.charAt(0);
  var row = parseInt(cellReference.substring(1));
  var cell = sheet.getRange(row, 2);
  var idProduct = cell.getValue();
  return idProduct;
}

/*
(idproduct, newdataset)
updateProductReference_
updateProductName_
updateProductDescShort_
updateProductDesc_
updateProductPrice_
updateProductCategory_
updateProductManufacturer_
updateProductCondition_
updateProductActive_
updateProductQty_
updateProductEAN13_
*/

//translateCellReference////////////////////////////////////////////////////////////////////////////////////
/**

*/
function translateCellReference_(cellReference) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PRODUCTS_Sheet.Name);
  var putData = sheet.getRange(cellReference).getValue();
  var columnNames = [
    "Reference", "Quantite", "Prix HT", "Nom", "Categories", "Condition", "Active",
    "EAN13", "Description court", "Description long", "Marque", "id tax rules group"
  ];
  var functionNames = [
    /*"updateProductReference_",
    "updateProductQty_",
    "updateProductPrice_",
    "updateProductName_",
    "updateProductCategory_",
    "updateProductCondition_",
    "updateProductActive_",
    "updateProductEAN13_",
    "updateProductDescShort_",
    "updateProductDesc_",
    "updateProductManufacturer_",
    "updateIdtaxRulesGroup"
    */
    "putReference_",
    "putQuantite_",
    "putPrixHT_",
    "putNom_",
    "putCategories_",
    "putCondition_",
    "putActive_",
    "putEAN13_",
    "putDescriptionCourt_",
    "putDescriptionLong_",
    "putMarque_",
    "putIdtaxRulesGroup_"
  ];
  
  var match = cellReference.match(/([A-Z]+)(\d+)/);
  if (match) {
    var columnLetter = match[1];
    var rowIndex = parseInt(match[2]);
    var columnIndex = columnLetter.charCodeAt(0) - 65 + 1;
    var columnName = columnNames[columnIndex - 3];
    var functionName = functionNames[columnIndex - 3]
    return {
      idProduct: getIdProductbycellReference_(cellReference),
      nameColumn: columnName,
      ligne: rowIndex,
      column: columnIndex,
      function : functionName,
      data : putData
    };
  }
  
  return null;
}



//callFunctionByName_////////////////////////////////////////////////////////////////////////////////////
/**

*/
async function callFunctionByName_(functionName, ...parameters) {
  if (typeof this[functionName] === "function") {
    return await this[functionName](...parameters);
  } else {
    Logger.log("Function not found: " + functionName);
  }
}




/*
//testT1////////////////////////////////////////////////////////////////////////////////////

function testT1(){

  // Example usage:

  var translation = translateCellReference("E4");
  Logger.log(translation);

  putData = translateCellReference("F2");
  callFunctionByName(putData.function,putData.idProduct,putData.data);

  translation = translateCellReference("C2");
  Logger.log(translation);
  // Example usage:
  // Example usage:
  var functionName = "functionName2";
  var parameter1 = "Hello";
  var parameter2 = "World!";
  var parameter3 = 123;
  callFunctionByName(functionName, parameter1, parameter2, parameter3);

}

function functionName1() {
  Logger.log("Function 1 called");
}

function functionName2(a, b, c) {
  Logger.log("Function 2 called with parameters: " + a + ", " + b + ", " + c);
}

//test2////////////////////////////////////////////////////////////////////////////////////

function test2() {
var str = "E2, D3, F3, I3, H5, F5";
var result = removeElement(str, 'E2');
Logger.log(result); // Logs: "D3, F3, I3, H5, F5"

var str = "E2";
var result = removeElement(str, 'E2');
Logger.log(result); // Logs: "D3, F3, I3, H5, F5"


var str = "E2, D3";
var result = removeElement(str, 'E2');
Logger.log(result); // Logs: "D3, F3, I3, H5, F5"


var str = "E2, D3, F3, I3, H5, F5";
var result = removeElement(str, 'D3');
Logger.log(result); // Logs: "D3, F3, I3, H5, F5"


var str = "E2, D3, F3, I3, H5, F5";
var result = removeElement(str, 'F9');
Logger.log(result); // Logs: "D3, F3, I3, H5, F5"
}


*/