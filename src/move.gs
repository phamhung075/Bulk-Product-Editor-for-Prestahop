// Author: Dai Hung PHAM
// Date: May 24, 2023
//
// Important Notice:
// This code is the intellectual property of Dai Hung PHAM. (daihung.pham@yahoo.fr)
// Do not reuse or share this code without explicit permission.
// If you wish to use or modify this code, please contact Dai Hung PHAM for authorization.
/**
  * A configuration object for debugging purposes related to products.
  * 
  * @type {Object}
  * @property {boolean} DEBUG_value Enables or disables logging of values
  * @property {boolean} DEBUG_request Enables or disables logging of requests
  * @property {boolean} DEBUG_xml Enables or disables logging of xml
 */
const DEBUG_MOVE = {
  DEBUG_value: true,
  DEBUG_request: true,
  DEBUG_xml:true,
}

//initialiser_Sheet_Product////////////////////////////////////////////////////////////////////////////////////////////////////////////
/**
 *  Initializes the "PRODUCTS" sheet by clearing its content, excluding the first row and first column,
 *  setting the header row, and preserving rows with non-empty cells in column B.
*/


function initialiser_Sheet_Product() {
  if (DEBUG_MOVE.DEBUG_value) {
    Logger.log(`run initialiser_Sheet_Product`);
  }
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(PRODUCTS_Sheet.Name);
  const removeRange = sheet.getRange("A2:Q");
  removeRange.clearContent();
  /*
  const keepRange = sheet.getRange(PRODUCTS_Sheet.Plage_product_id);
  const removeRange1 = sheet.getRange(PRODUCTS_Sheet.Plage_data);
  const removeRange2 = sheet.getRange(PRODUCTS_Sheet.Plage_change_history);
  const keepValues = keepRange.getValues().flat().filter(Boolean).map(row => [row]);
  
  if (DEBUG_MOVE.DEBUG_value) {
    
    Logger.log(`PRODUCTS_Sheet.Plage_product_id of initialiser_Sheet_Product(): ${PRODUCTS_Sheet.Plage_product_id}`);
    Logger.log(`keepRange of initialiser_Sheet_Product(): ${keepValues}`);
    Logger.log(`removeRange of initialiser_Sheet_Product(): ${removeRange1}`);
    Logger.log(`keepValues of initialiser_Sheet_Product(): ${keepValues}`);
    Logger.log(`PRODUCTS_Sheet.HEADERS.length: ${PRODUCTS_Sheet.HEADERS.length}`);
    Logger.log(`PRODUCTS_Sheet.HEADERS: ${PRODUCTS_Sheet.HEADERS}`);
  }

  removeRange1.clearContent();
  removeRange2.clearContent();
  sheet.getRange(1, 3, 1, PRODUCTS_Sheet.HEADERS.length).setValues([PRODUCTS_Sheet.HEADERS]);
  sheet.getRange(2, 2, keepValues.length, 1).setValues(keepValues);

    */
}

function initialiser_Sheet_Product_keepIDs() {
  if (DEBUG_MOVE.DEBUG_value) {
    Logger.log(`run initialiser_Sheet_Product`);
  }
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(PRODUCTS_Sheet.Name);

  const keepRange = sheet.getRange(PRODUCTS_Sheet.Plage_product_id);
  const removeRange1 = sheet.getRange(PRODUCTS_Sheet.Plage_data);
  const removeRange2 = sheet.getRange(PRODUCTS_Sheet.Plage_change_history);
  const keepValues = keepRange.getValues().flat().filter(Boolean).map(row => [row]);
  
  if (DEBUG_MOVE.DEBUG_value) {
    
    Logger.log(`PRODUCTS_Sheet.Plage_product_id of initialiser_Sheet_Product(): ${PRODUCTS_Sheet.Plage_product_id}`);
    Logger.log(`keepRange of initialiser_Sheet_Product(): ${keepValues}`);
    Logger.log(`removeRange of initialiser_Sheet_Product(): ${removeRange1}`);
    Logger.log(`keepValues of initialiser_Sheet_Product(): ${keepValues}`);
    Logger.log(`PRODUCTS_Sheet.HEADERS.length: ${PRODUCTS_Sheet.HEADERS.length}`);
    Logger.log(`PRODUCTS_Sheet.HEADERS: ${PRODUCTS_Sheet.HEADERS}`);
  }

  removeRange1.clearContent();
  removeRange2.clearContent();
  sheet.getRange(1, 3, 1, PRODUCTS_Sheet.HEADERS.length).setValues([PRODUCTS_Sheet.HEADERS]);
  sheet.getRange(2, 2, keepValues.length, 1).setValues(keepValues);
}

//initialiser_Sheet_LEBONCOINCSV_////////////////////////////////////////////////////////////////////////////////////////////////////////////////
/**
 *  Initializes the "LEBONCOINCSV" sheet by clearing its content from row 2 to the end of column B.
*/

function initialiser_Sheet_LEBONCOINCSV_() {
  if (DEBUG_MOVE.DEBUG_value) {
    Logger.log(`run initialiser_Sheet_LEBONCOINCSV_`);
  }
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LEBONCOINCSV.Name).getRange(LEBONCOINCSV.Plage_data).clearContent();
}


//initialiser_Sheet_FNACCSV_////////////////////////////////////////////////////////////////////////////////////////////////////////////////
/**
 *  Initializes the "FNACCSV" sheet by clearing its content from row 2 to the end of column B.
*/


function initialiser_Sheet_FNACCSV_() {
  if (DEBUG_MOVE.DEBUG_value) {
    Logger.log(`run initialiser_Sheet_FNACCSV_`);
  }
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName(FNACCSV.Name).getRange(FNACCSV.Plage_data).clearContent();
}


//initialiser_Sheet_RAKUTENCSV_////////////////////////////////////////////////////////////////////////////////////////////////////////////////
/**
 *  Initializes the "RAKUTENCSV" sheet by clearing its content from row 2 to the end of column B.
*/


function initialiser_Sheet_RAKUTENCSV_() {
  if (DEBUG_MOVE.DEBUG_value) {
    Logger.log(`run initialiser_Sheet_RAKUTENCSV_`);
  }
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName(RAKUTENCSV.Name).getRange(RAKUTENCSV.Plage_data).clearContent();
}


//initializeSheetFilter_////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
/**
 * Initializes the ORDERS_Check_Stock sheet by clearing its content from row 2 to the end of column B.
*/

function initializeSheetFilter_() {
  if (DEBUG_MOVE.DEBUG_value) {
    Logger.log(`run initializeSheetFilter_`);
  }
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ORDERS_Check_Stock.Name).getRange(ORDERS_Check_Stock.Plage_data).clearContent();
}



//myVLookup_////////////////////////////////////////////////////////////////////////////////////////////////////////////////
/**
 *  Searches for a value in a specific column of a range in a given sheet and returns the value in a specified return column.
 * 
 *  @param {string} searchValue The value to search for in the search column.
 *  @param {string} sheetName The name of the sheet to search.
 *  @param {string} range The range to search in the format "A1:B10".
 *  @param {number} searchColumn The column number to search in, starting from 1.
 *  @param {number} returnColumn The column number to return the value from, starting from 1.
 *  @return {string} The value found in the specified return column or "#N/A" if the value is not found.
*/


function myVLookup_(searchValue, sheetName, range, searchColumn, returnColumn) {
  if (DEBUG_MOVE.DEBUG_value) {
    Logger.log(`run myVLookup_`);
  }
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  var data = sheet.getRange(range).getValues();
  for (var i = 0; i < data.length; i++) {
    if (data[i][searchColumn - 1] == searchValue) {
      return data[i][returnColumn - 1];
    }
  }
  return "#N/A"; // Return "#N/A" if value not found
}


//createDropdownList_//////////////////////////////////////////////////////////////////////////////////////////////////////////////
/**
 *  Creates a dropdown list for a given sheet, cell range and list of values.
 * 
 *  @param {string[]} dropDownList The list of values for the dropdown
 *  @param {string} sheetName The name of the sheet where the dropdown will be created
 *  @param {string} plage The cell range where the dropdown will be created
 *  @return {void}
*/
function createDropdownList_(dropDownList, sheetName, plage) {
  if (DEBUG_MOVE.DEBUG_value) {
    Logger.log(`run createDropdownList_`);
  }
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName); // modify sheet name here

  
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(dropDownList) // modify drop-down list here
    .setAllowInvalid(false)
    .build();

  var targetRange = sheet.getRange(plage); // modify range here
  targetRange.setDataValidation(rule);
}




//get_IdsProductbyRef_//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//https://tabletandtv.com/api/products?filter[reference]=%25[test]%25
/**
 *  Retrieves the IDs of products with a given reference from a PrestaShop API endpoint.
 * 
 *  @param {string} url The URL of the API endpoint to retrieve the product IDs from.
 *  @return {Promise<string[]>} A promise that resolves with an array of product IDs or rejects with an error message.
*/


async function get_IdsProductbyRef_(url) {
  if (DEBUG_MOVE.DEBUG_value) {
    Logger.log(`run get_IdsProductbyRef_`);
  }
  return new Promise(async (resolve, reject) => {
    const xml = await getContentText_base64EncodedAuthorizationKey_(url, false);
    if (DEBUG_MOVE.DEBUG_xml) {
      Logger.log(`xml of get_IdsProductbyRef_(): ${xml}`);
    }
    const document = XmlService.parse(xml);
    const root = document.getRootElement();
    const productElements = root
      .getChildren("products")[0]
      .getChildren("product");
    if (productElements.length === 0) {
      reject("No product elements found in API response");
    } else {
      const ids = productElements.map(productElement => productElement.getAttribute("id").getValue());
      resolve(ids);
      if (DEBUG_MOVE.DEBUG_value) {
        Logger.log(`ids of get_IdsProductbyRef_(): ${ids}`);
      }
    }
  });
}



//filter_button////////////////////////////////////////////////////////////////////////////////////////////////////////////////
/**
 * Filters products on a sheet based on their reference using a PrestaShop API endpoint.
*/
async function filter_button() {
  if (DEBUG_MOVE.DEBUG_value) {
    Logger.log(`run filter_button`);
  }
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(FILTER_Sheet.Name);
  SpreadsheetApp.flush();
  const data = await get_IdsProductbyRef_(getlinkStockAPIbyRef_());
  const mappedData = data.map(row => [row]);
  const range = sheet.getRange(2, 1, mappedData.length, mappedData[0].length);
  range.clearContent();
  sheet.getRange(FILTER_Sheet.Plage_product_id).clearContent();
  range.setValues(mappedData);
  SpreadsheetApp.flush();
  if (DEBUG_MOVE.DEBUG_value) {
    Logger.log(`finds data of filter_button(): ${mappedData}`);
  }
  await updateList_();
}

//updateList_////////////////////////////////////////////////////////////////////////////////////////////////////////////////
/**
 *  Updates the list of products in the sheet by adding a new value if it does not already exist.
 * 
 * @return {void}
*/

function updateList_() {
  if (DEBUG_MOVE.DEBUG_value) {
    Logger.log(`run updateList_`);
  }
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(FILTER_Sheet.Name);
  const listRange = sheet.getRange(FILTER_Sheet.Plage_history);
  const listValues = listRange.getValues().flat().filter(Boolean).map(row => [row]);
  const sourceValue = sheet.getRange(FILTER_Sheet.Option_Filter1_Cell_Value).getValue();
  
  if (!listValues.includes(sourceValue)) { // Check if value already exists in the list
    if (DEBUG_MOVE.DEBUG_value) {
      Logger.log(`new value in history getProductStockQty_(): ${sourceValue}`);
    }
    const cell = sheet.getRange(`C${listValues.length+12}`);
    cell.setValue(sourceValue);
  }
}


//moveAtoAS_FILTER_to_PRODUCTS////////////////////////////////////////////////////////////////////////////////////////////////////////////////
/**
 *  Moves data from the A column of the "FILTER" sheet to the B column of the "PRODUCTS" sheet.
 * 
 * @return {void}
*/
function moveAtoAS_FILTER_to_PRODUCTS() {
  if (DEBUG_MOVE.DEBUG_value) {
    Logger.log(`run moveAtoAS_FILTER_to_PRODUCTS`);
  }
  moveRangeS1ToS2_(FILTER_Sheet.Plage_product_id, PRODUCTS_Sheet.Plage_product_id, FILTER_Sheet.Name, PRODUCTS_Sheet.Name);
}

//moveAtoAS_ORDERS_to_PRODUCTS//////////////////////////////////////////////////////////////////////////////////////////////////////////////
/**
 *  Moves data from the A column of the "ORDERS" sheet to the B column of the "PRODUCTS" sheet, without clearing the source data.
 * 
 *  @return {void}
*/

function moveAtoAS_ORDERS_to_PRODUCTS() {
  if (DEBUG_MOVE.DEBUG_value) {
    Logger.log(`run moveAtoAS_ORDERS_to_PRODUCTS`);
  }
  moveRangeS1ToS2notClear_(ORDERS_Sheet.Plage_product_id, PRODUCTS_Sheet.Plage_product_id, ORDERS_Sheet.Name, PRODUCTS_Sheet.Name);
}

//moveRangeS1ToS2notClear_//////////////////////////////////////////////////////////////////////////////////////////////////////////////
/**
 *  Moves data from the A column of the "ORDERS" sheet to the B column of the "PRODUCTS" sheet, without clearing the source data.
 * 
 * @return {void}
*/

function moveRangeS1ToS2notClear_(nameCopyrangelikeA1A, namePastrangelikeA1A, sheet1Name, sheet2Name) {
  if (DEBUG_MOVE.DEBUG_value) {
    Logger.log(`run moveRangeS1ToS2notClear_`);
  }
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet1 = ss.getSheetByName(sheet1Name);
  const sheet2 = ss.getSheetByName(sheet2Name);

  const range1 = sheet1.getRange(nameCopyrangelikeA1A);
  const range2 = sheet2.getRange(namePastrangelikeA1A);

  range2.clearContent();

  const values1 = range1.getValues().flat().filter(Boolean).map(row => [row[0]]);
  
  if (DEBUG_MOVE.DEBUG_value) {
    Logger.log(`values filter moveRangeS1ToS2_(): ${values1}`);
  }

  const destinationRange = sheet2.getRange(2, 2, values1.length, 1);
  
  if (DEBUG_MOVE.DEBUG_value) {
    Logger.log(`values moveRangeS1ToS2_(): ${values1}`);
  }

  range1.copyTo(destinationRange, { contentsOnly: true });
}

//moveRangeS1ToS2_//////////////////////////////////////////////////////////////////////////////////////////////////////////////
/**
 *  Moves data from the A column of the "ORDERS" sheet to the B column of the "PRODUCTS" sheet, clearing the source data.
 * 
 * @return {void}
*/
function moveRangeS1ToS2_(nameCopyrangelikeA1A, namePastrangelikeA1A, sheet1Name, sheet2Name) {
  if (DEBUG_MOVE.DEBUG_value) {
    Logger.log(`run moveRangeS1ToS2_`);
  }
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet1 = ss.getSheetByName(sheet1Name);
  const sheet2 = ss.getSheetByName(sheet2Name);

  const range1 = sheet1.getRange(nameCopyrangelikeA1A);
  const range2 = sheet2.getRange(namePastrangelikeA1A);

  range2.clearContent();

  const values1 = range1.getValues().flat().filter(Boolean).map(row => [row[0]]);
  
  if (DEBUG_MOVE.DEBUG_value) {
    Logger.log(`values filter moveRangeS1ToS2_(): ${values1}`);
  }

  const destinationRange = sheet2.getRange(2, 2, values1.length, 1);
  
  if (DEBUG_MOVE.DEBUG_value) {
    Logger.log(`values moveRangeS1ToS2_(): ${values1}`);
  }

  range1.copyTo(destinationRange, { contentsOnly: true });
  range1.clearContent();
}
