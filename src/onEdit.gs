// Author: Dai Hung PHAM
// Date: May 24, 2023
//
// Important Notice:
// This code is the intellectual property of Dai Hung PHAM. (daihung.pham@yahoo.fr)
// Do not reuse or share this code without explicit permission.
// If you wish to use or modify this code, please contact Dai Hung PHAM for authorization.
//onEdit//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
/**

*/
function onEdit(e) {
  var range = e.range;
  var column = range.getColumn();
  var row = range.getRow();
  var sheet = range.getSheet();

  // If the edit was made in column B (2nd column) and it's an empty cell
  if(column == 2 && sheet.getRange(row, column).getValue() == "") {
    sheet.getRange(row, 17).setValue("");  // Set the corresponding Q column cell as ""
    return;
  }
  
  // If the edit was made between column C (3rd) and column N (14th)
  if(column >= 3 && column <= 14) {
    var cellQ = sheet.getRange(row, 17);
    var prevValue = cellQ.getValue();
    var cellRef = range.getA1Notation();
    
    // Check if the cell in column Q is empty or not
    if(prevValue != "") {
      var cellRefs = prevValue.split(",");
      if(cellRefs.indexOf(cellRef) == -1) {
        cellQ.setValue(prevValue + ',' + cellRef); // append the edited cell reference only if it's not already recorded
      }
    } else {
      cellQ.setValue(cellRef); // write the edited cell reference
    }
  }
}
