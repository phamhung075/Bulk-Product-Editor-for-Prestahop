// Author: Dai Hung PHAM
// Date: May 24, 2023
//
// Important Notice:
// This code is the intellectual property of Dai Hung PHAM. (daihung.pham@yahoo.fr)
// Do not reuse or share this code without explicit permission.
// If you wish to use or modify this code, please contact Dai Hung PHAM for authorization.
const DEBUG_GETLINK = {
  DEBUG_value: false,
  DEBUG_request: false,
}
//getlinkProductAPIbyID_//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
/**
 *  Retrieves the API link for a product with the given ID.
 * 
 *  @param {string} idProduct The ID of the product for which to retrieve the API link
 *  @return {string} The API link for the product with the given ID
*/
function showAlert(alert) {
  var ui = SpreadsheetApp.getUi();
  ui.alert('Alert', alert, ui.ButtonSet.OK);
}
function getDataFromApi_(method) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG_Sheet.Name);
  const idPOST = sheet.getRange(CONFIG_Sheet.IdPOST).getValue();
  const domain = sheet.getRange(CONFIG_Sheet.Prestashop_domain_Cell).getValue();
  const passPOST = sheet.getRange(CONFIG_Sheet.PassPOST).getValue();
  const url = 'https://apisyncproduct-v230518.oa.r.appspot.com/ecom';
  const data = {
    'pseudo': idPOST,
    'password': passPOST,
    'domain': domain,
    'methodGS': method,
  };

  const options = {
    'method' : 'POST',
    'payload' : JSON.stringify(data),
    'contentType': 'application/json',
    'muteHttpExceptions': true
  };

  const response = UrlFetchApp.fetch(url, options);
  if (response.getResponseCode() == 401) {
    showAlert("Bad login or password");
    return "Bad login or password";
  }
  if (response.getResponseCode() == 402) {
    sheet.getRange(CONFIG_Sheet.Credit).setValue(0);
    showAlert("Payment Required: Please add more credit");
    return "Payment Required: Please add more credit";
  }
    if (response.getResponseCode() == 403) {
    sheet.getRange(CONFIG_Sheet.Credit).setValue(0);
    showAlert("Invalid domain configuration");
    return "Invalid domain configuration";
  }
  if (response.getResponseCode() == 200) {
    const json = response.getContentText();
    try {
      const responseParse = JSON.parse(json);
      Logger.log(responseParse.data);
      //return data;
      sheet.getRange(CONFIG_Sheet.Credit).setValue(responseParse.credit);
      return responseParse.data;
    } catch(e) {
      Logger.log("Error: " + e);
      Logger.log("Response: " + json);
    }
  } else {
    Logger.log("Response code: " + response.getResponseCode());
    Logger.log("Response body: " + response.getContentText());
    showAlert(response.getContentText());
  }
}


//take1Credit_//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
/**

*/
function take1Credit_() {
  Logger.log(getDataFromApi_('takeCredit'));
}



//getlinkAPIbyIDProduct_//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
/**

*/
function getlinkAPIbyIDProduct_(idProduct) {
  const apiMethod = getDataFromApi_('getlinkProductAPIbyID');
  if (DEBUG_GETLINK.DEBUG_value){
    Logger.log(`apiMethod getlinkProductAPIbyID_(): ${apiMethod}`);
  }
  const url = `${apiMethod}${idProduct}`;
  if (DEBUG_PRODUCT_GET.DEBUG_request || DEBUG_PRODUCT_PUT.DEBUG_request || DEBUG_GETLINK.DEBUG_request){
    Logger.log(`getlinkProductAPIbyID_(): ${url}`);
  }
  return url;
}

//getlinkProductAPIbyID_//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
/**

*/
function getlinkProductAPIbyID_(idProduct) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG_Sheet.Name);
  const vSite = sheet.getRange(CONFIG_Sheet.Prestashop_domain_Cell).getValue();
  //return `${vSite}/api/products/${idProduct}`;
  const url = `${vSite}/api/products/${idProduct}`;
  if (DEBUG_PRODUCT_GET.DEBUG_request || DEBUG_PRODUCT_PUT.DEBUG_request){
    Logger.log(`getlinkProductAPIbyID_(): ${url}`);
  }
  return url;
}



//getlinkAPIbyIDStockProduct_//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
/**
 *  Retrieves the API link for a product stock with the given ID.
 * 
 *  @param {string} productStockID The ID of the product stock for which to retrieve the API link
 *  @return {string} The API link for the product stock with the given ID
*/

function getlinkAPIbyIDStockProduct_(productStockID) {
  const apiMethod = getDataFromApi_('getlinkStockAPIbyID');
  if (DEBUG_GETLINK.DEBUG_value){
    Logger.log(`apiMethod getlinkStockAPIbyID_(): ${apiMethod}`);
  }
  const url = `${apiMethod}${productStockID}`;
  if (DEBUG_PRODUCT_GET.DEBUG_request || DEBUG_PRODUCT_PUT.DEBUG_request || DEBUG_GETLINK.DEBUG_request){
    Logger.log(`getlinkStockAPIbyID_(): ${url}`);
  }
  return url;
}


//getlinkStockAPIbyID_//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
/**
 *  Retrieves the API link for a product stock with the given ID.
 * 
 *  @param {string} productStockID The ID of the product stock for which to retrieve the API link
 *  @return {string} The API link for the product stock with the given ID
*/

function getlinkStockAPIbyID_(productStockID) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG_Sheet.Name);
  const domain = sheet.getRange(CONFIG_Sheet.Prestashop_domain_Cell).getValue();
  //return `${domain}/api/stock_availables/${productStockID}`;
  const url = `${domain}/api/stock_availables/${productStockID}`;
  if (DEBUG_PRODUCT_GET.DEBUG_request || DEBUG_PRODUCT_PUT.DEBUG_request){
    Logger.log(`getlinkStockAPIbyID_(): ${url}`);
  }
  return url;
}


//getlinkStockAPIbyRef_//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
/**
 *  Retrieves the API link for a product stock with a given reference.
 * 
 *  @return {string} The API link for the product stock with the given reference
*/

function getlinkStockAPIbyRef_() {
  take1Credit_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG_Sheet.Name);
  const domain = sheet.getRange(CONFIG_Sheet.Prestashop_domain_Cell).getValue();
  const sheet2 = ss.getSheetByName(FILTER_Sheet.Name);

  // Get the values of the cells
  const typeFil = sheet2.getRange(FILTER_Sheet.TypefilRange).getValues();
  const dataFil = sheet2.getRange(FILTER_Sheet.DatafilRange).getValues();

  let url = domain + "/api/products?";
  let firstFilter = true;

  for (let i = 0; i < dataFil.length; i++) {
    if (dataFil[i][0] !== "#WaitingForNextVersion" && dataFil[i][0] !== null && dataFil[i][0] !== "") {
      if (firstFilter) {
        firstFilter = false;
      } else {
        url += "&";
      }
      url += "filter[" + typeFil[i][0] + "]=" + encodeURIComponent('%') + "[" + encodeURIComponent(dataFil[i][0]) + "]" + encodeURIComponent('%');
    }
  }

  if (DEBUG_MOVE.DEBUG_request) {
    Logger.log(`getlinkStockAPIbyRef_(): ${url}`);
  }

  return url;
}



//getlinkCountriesIDsAPI_////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
/**
 *  Retrieves the API link for all countries.
 * 
 *  @return {string} The API link for all countries
*/

function getlinkCountriesIDsAPI_() {
  const apiMethod = getDataFromApi_('getlinkCountriesIDsAPI');
  if (DEBUG_GETLINK.DEBUG_value){
    Logger.log(`apiMethod getlinkCountriesIDsAPI_(): ${apiMethod}`);
  }
  const url = `${apiMethod}`;
  if (DEBUG_PRODUCT_GET.DEBUG_request || DEBUG_PRODUCT_PUT.DEBUG_request || DEBUG_GETLINK.DEBUG_request){
    Logger.log(`getlinkCountriesIDsAPI_(): ${url}`);
  }
  return url;
}


//getlinkCategoriesIDsAPI_////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
/**
 *  Retrieves the API link for all categories.
 * 
 *  @return {string} The API link for all categories
*/

function getlinkCategoriesIDsAPI_(){
  const apiMethod = getDataFromApi_('getlinkCategoriesIDsAPI');
  if (DEBUG_GETLINK.DEBUG_value){
    Logger.log(`apiMethod getlinkCategoriesIDsAPI_(): ${apiMethod}`);
  }
  const url = `${apiMethod}`;
  if (DEBUG_PRODUCT_GET.DEBUG_request || DEBUG_PRODUCT_PUT.DEBUG_request || DEBUG_GETLINK.DEBUG_request){
    Logger.log(`getlinkCategoriesIDsAPI_(): ${url}`);
  }
  return url;
}



//getObjetsAllManufactureLinks_////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
/**
 *  Gets the URL for retrieving the IDs of all manufacturers in the PrestaShop API.
 *  This function retrieves the domain name from the config sheet, and then appends the necessary path
 *  to the PrestaShop API to retrieve the IDs of all manufacturers.
 * 
 *  @return {string} The URL for retrieving the IDs of all manufacturers in the PrestaShop API.
*/

function getlinkManufacturersIDsAPI_(){
  const apiMethod = getDataFromApi_('getlinkManufacturersIDsAPI');
  if (DEBUG_GETLINK.DEBUG_value){
    Logger.log(`apiMethod getlinkManufacturersIDsAPI_(): ${apiMethod}`);
  }
  const url = `${apiMethod}`;
  if (DEBUG_PRODUCT_GET.DEBUG_request || DEBUG_PRODUCT_PUT.DEBUG_request || DEBUG_GETLINK.DEBUG_request){
    Logger.log(`getlinkManufacturersIDsAPI_(): ${url}`);
  }
  return url;
}


//getlinkTaxsIDsAPI_////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
/**
 * Retrieves the URL for the Tax Rule Groups API endpoint from the active Google Spreadsheet configuration sheet.
 * 
 * @return {string} The URL for the Tax Rule Groups API endpoint.
*/

function getlinkTaxsIDsAPI_(){
  const apiMethod = getDataFromApi_('getlinkTaxsIDsAPI');
  if (DEBUG_GETLINK.DEBUG_value){
    Logger.log(`apiMethod getlinkTaxsIDsAPI_(): ${apiMethod}`);
  }
  const url = `${apiMethod}`;
  if (DEBUG_PRODUCT_GET.DEBUG_request || DEBUG_PRODUCT_PUT.DEBUG_request || DEBUG_GETLINK.DEBUG_request){
    Logger.log(`getlinkTaxsIDsAPI_(): ${url}`);
  }
  return url;
}



