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
  let root = document.getRootElement();
  let manufacturersElement = root.getChild('manufacturers');
  let manufacturerElements = manufacturersElement.getChildren('manufacturer');
  let manufacturerLinks = {};

  for (let i = 0; i < manufacturerElements.length; i++) {
    let manufacturer = manufacturerElements[i];
    let id = manufacturer.getAttribute('id').getValue();
    let xlinkNamespace = XmlService.getNamespace('xlink', 'http://www.w3.org/1999/xlink');
    let href = manufacturer.getAttribute('href', xlinkNamespace).getValue();
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
  let ids = [];
  for (let id in manufacturerLinks) {
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
  let xlinks = [];
  for (let id in manufacturerLinks) {
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
  let root = document.getRootElement();
  let manufacturerElement = root.getChild('manufacturer');
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
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet1 = ss.getSheetByName(CONFIG_Sheet.Name);
  let columnGetDropDown = sheet1.getRange(CONFIG_Sheet.Manufactures_DropDown_Column_Cell).getValue()+1;
  if(DEBUG_MANUFACTURES.DEBUG_value){
    Logger.log("columnGetDropDown:"+ columnGetDropDown);
  }
  let sheet2 = ss.getSheetByName(PRODUCTS_Sheet.Name);
  let dataRange = sheet1.getRange(2, columnGetDropDown, sheet1.getLastRow() - 1);
  let values = dataRange.getValues();
  let flatValues = values.flat().filter(String); // Supprime les cellules vides
  
  let rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(flatValues, true)
    .setAllowInvalid(false)
    .build();

  let targetRange = sheet2.getRange(PRODUCTS_Sheet.Plage_product_manufacturer);
  targetRange.setDataValidation(rule);
}
