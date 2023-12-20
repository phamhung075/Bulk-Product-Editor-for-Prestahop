// Author: Dai Hung PHAM
// Date: May 24, 2023
//
// Important Notice:
// This code is the intellectual property of Dai Hung PHAM. (daihung.pham@yahoo.fr)
// Do not reuse or share this code without explicit permission.
// If you wish to use or modify this code, please contact Dai Hung PHAM for authorization.
/**
  * A configuration object for debugging purposes related to manufacturers.
  * 
  * @type {Object}
  * @property {boolean} DEBUG_value Enables or disables logging of manufacturers values
  * @property {boolean} DEBUG_request Enables or disables logging of manufacturers-related API requests
 */
const DEBUG_MANUFACTURES = {
  DEBUG_value: false,
  DEBUG_request: false,
  DEBUG_xml:false,
}

//manufacturers manufacturer

//getObjetsAllManufactureLinks_////////////////////////////////////////////////////////////////////////////////////
/**
 *  Retrieves all manufacturer links from the API and returns them as an object with manufacturer IDs as keys and links as values.
 * 
 *  @return {Promise<Object<string, string>>} A promise that resolves with an object containing manufacturer IDs as keys and links as values
*/
async function getObjetsAllManufactureLinks_() {
  const url = getlinkManufacturersIDsAPI_();
  const manufacturersContent = await getContentText_base64EncodedAuthorizationKey_(url, false);
  const document = XmlService.parse(manufacturersContent);
  if(DEBUG_MANUFACTURES.DEBUG_xml){
      Logger.log("document: " + document);
  }
  var root = document.getRootElement();
  var manufacturersElement = root.getChild('manufacturers');
  var manufacturerElements = manufacturersElement.getChildren('manufacturer');
  var manufacturerLinks = {};

  for (var i = 0; i < manufacturerElements.length; i++) {
    var manufacturer = manufacturerElements[i];
    var id = manufacturer.getAttribute('id').getValue();
    var xlinkNamespace = XmlService.getNamespace('xlink', 'http://www.w3.org/1999/xlink');
    var href = manufacturer.getAttribute('href', xlinkNamespace).getValue();
    manufacturerLinks[id] = href;
  }

  return manufacturerLinks;
}


//getManufactureIds_////////////////////////////////////////////////////////////////////////////////////
/**
 *  Retrieves an array of manufacturer IDs from an object containing manufacturer links.
 * 
 *  @param {Object<string, string>} manufacturerLinks An object containing manufacturer IDs as keys and links as values
 *  @return {Array<string>} An array of manufacturer IDs
*/
function getManufactureIds_(manufacturerLinks) {
  var ids = [];
  for (var id in manufacturerLinks) {
    ids.push(id);
  }
  return ids;
}



//getManufactureXlinks_////////////////////////////////////////////////////////////////////////////////////
/**
 *  Retrieves an array of manufacturer links from an object containing manufacturer links.
 *  @param {Object<string, string>} manufacturerLinks An object containing manufacturer IDs as keys and links as values
 *  @return {Array<string>} An array of manufacturer links
*/
function getManufactureXlinks_(manufacturerLinks) {
  var xlinks = [];
  for (var id in manufacturerLinks) {
    xlinks.push(manufacturerLinks[id]);
  }
  return xlinks;
}



//getLanguageNameFromManufactureLink_////////////////////////////////////////////////////////////////////////////////////
/**

*/
async function getLanguageNameFromManufactureLink_(manufacturerLink) {
  const manufacturerContent = await getContentText_base64EncodedAuthorizationKey_(manufacturerLink, true);
  const document = XmlService.parse(manufacturerContent);
  var root = document.getRootElement();
  var manufacturerElement = root.getChild('manufacturer');
  return nameElement = manufacturerElement.getChild('name').getText();
}



//writeManufactureInfoToSheet////////////////////////////////////////////////////////////////////////////////////
/**

*/
async function writeManufactureInfoToSheet() {
  const manufacturerLinks = await getObjetsAllManufactureLinks_();
  const manufacturerIds = getManufactureIds_(manufacturerLinks);
  const manufacturerXlinks = getManufactureXlinks_(manufacturerLinks);
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG_Sheet.Name);
  const column = sheet.getRange(CONFIG_Sheet.Manufactures_DropDown_Column_Cell).getValue();
  let begin = sheet.getRange(CONFIG_Sheet.Manufacturers_Lenght_Cell).getValue() || 0;
  sheet.getRange(CONFIG_Sheet.Manufacturers_Name_Plage).setValue("MANUFACTURES");
  for (let i = begin ; i < manufacturerIds.length; i++) {
    const languageName = await getLanguageNameFromManufactureLink_(manufacturerXlinks[i]);
    if(DEBUG_MANUFACTURES.DEBUG_value){
      Logger.log("id: " + manufacturerIds[i] + " languageName:" + languageName);
    }
    sheet.getRange(CONFIG_Sheet.Manufacturers_Lenght_Cell).setValue(i+1);
    sheet.getRange(i + 2, column).setValue(manufacturerIds[i]);     // Column D: Manufacturer IDs
    sheet.getRange(i + 2, column +1 ).setValue(languageName+ " ("+manufacturerIds[i]+")");      // Column E: Language Manufacturers Names
  }
  createDropdownListManufactures_();
}


//createDropdownListManufactures_////////////////////////////////////////////////////////////////////////////////////
/**

*/
function createDropdownListManufactures_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet1 = ss.getSheetByName(CONFIG_Sheet.Name);
  var columnGetDropDown = sheet1.getRange(CONFIG_Sheet.Manufactures_DropDown_Column_Cell).getValue()+1;
  if(DEBUG_MANUFACTURES.DEBUG_value){
    Logger.log("columnGetDropDown:"+ columnGetDropDown);
  }
  var sheet2 = ss.getSheetByName(PRODUCTS_Sheet.Name);
  var dataRange = sheet1.getRange(2, columnGetDropDown, sheet1.getLastRow() - 1);
  var values = dataRange.getValues();
  var flatValues = values.flat().filter(String); // Supprime les cellules vides
  
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(flatValues, true)
    .setAllowInvalid(false)
    .build();

  var targetRange = sheet2.getRange(PRODUCTS_Sheet.Plage_product_manufacturer);
  targetRange.setDataValidation(rule);
}
