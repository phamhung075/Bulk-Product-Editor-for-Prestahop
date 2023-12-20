// Author: Dai Hung PHAM
// Date: May 24, 2023
//
// Important Notice:
// This code is the intellectual property of Dai Hung PHAM. (daihung.pham@yahoo.fr)
// Do not reuse or share this code without explicit permission.
// If you wish to use or modify this code, please contact Dai Hung PHAM for authorization.
/**
  * A configuration object for debugging purposes related to taxs.
  * 
  * @type {Object}
  * @property {boolean} DEBUG_value Enables or disables logging of taxs values
  * @property {boolean} DEBUG_request Enables or disables logging of taxs-related API requests
 */
const DEBUG_TAXS = {
  DEBUG_value: false,
  DEBUG_request: false,
  DEBUG_xml:false,
}

//taxs tax

//getObjetsAllTaxLinks_////////////////////////////////////////////////////////////////////////////////////
/**
 *  Retrieves all tax links from the API and returns them as an object with tax IDs as keys and links as values.
 * 
 *  @return {Promise<Object<string, string>>} A promise that resolves with an object containing tax IDs as keys and links as values
*/
async function getObjetsAllTaxLinks_() {
  const url = getlinkTaxsIDsAPI_();
  const taxsContent = await getContentText_base64EncodedAuthorizationKey_(url, false);
  const document = XmlService.parse(taxsContent);
  if(DEBUG_TAXS.DEBUG_xml){
      Logger.log("document: " + document);
  }
  var root = document.getRootElement();
  var taxsElement = root.getChild('tax_rule_groups');
  var taxElements = taxsElement.getChildren('tax_rule_group');
  var taxLinks = {};

  for (var i = 0; i < taxElements.length; i++) {
    var tax = taxElements[i];
    var id = tax.getAttribute('id').getValue();
    var xlinkNamespace = XmlService.getNamespace('xlink', 'http://www.w3.org/1999/xlink');
    var href = tax.getAttribute('href', xlinkNamespace).getValue();
    taxLinks[id] = href;
  }

  return taxLinks;
}


//getTaxIds_////////////////////////////////////////////////////////////////////////////////////
/**
 *  Retrieves an array of tax IDs from an object containing tax links.
 * 
 *  @param {Object<string, string>} taxLinks An object containing tax IDs as keys and links as values
 *  @return {Array<string>} An array of tax IDs
*/
function getTaxIds_(taxLinks) {
  var ids = [];
  for (var id in taxLinks) {
    ids.push(id);
  }
  return ids;
}



//getTaxXlinks_////////////////////////////////////////////////////////////////////////////////////
/**
 *  Retrieves an array of tax links from an object containing tax links.
 *  @param {Object<string, string>} taxLinks An object containing tax IDs as keys and links as values
 *  @return {Array<string>} An array of tax links
*/
function getTaxXlinks_(taxLinks) {
  var xlinks = [];
  for (var id in taxLinks) {
    xlinks.push(taxLinks[id]);
  }
  return xlinks;
}



//getLanguageNameFromTaxLink_////////////////////////////////////////////////////////////////////////////////////
/**

*/
async function getLanguageNameFromTaxLink_(taxLink) {
  const taxContent = await getContentText_base64EncodedAuthorizationKey_(taxLink, true);
  const document = XmlService.parse(taxContent);
  var root = document.getRootElement();
  var taxElement = root.getChild('tax_rule_group');
  return nameElement = taxElement.getChild('name').getText();
}



//writeTaxInfoToSheet////////////////////////////////////////////////////////////////////////////////////
/**

*/
async function writeTaxInfoToSheet() {
  const taxLinks = await getObjetsAllTaxLinks_();
  const taxIds = getTaxIds_(taxLinks);
  const taxXlinks = getTaxXlinks_(taxLinks);
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG_Sheet.Name);
  const column = sheet.getRange(CONFIG_Sheet.Taxs_DropDown_Column_Cell).getValue();
  let begin = sheet.getRange(CONFIG_Sheet.Taxs_Lenght_Cell).getValue() || 0;
  sheet.getRange(CONFIG_Sheet.Taxs_Name_Plage).setValue("TAXES");
  for (let i = begin ; i < taxIds.length; i++) {
    const languageName = await getLanguageNameFromTaxLink_(taxXlinks[i]);
    if(DEBUG_TAXS.DEBUG_value){
      Logger.log("id: " + taxIds[i] + " languageName:" + languageName);
    }
    sheet.getRange(CONFIG_Sheet.Taxs_Lenght_Cell).setValue(i+1);
    sheet.getRange(i + 2, column).setValue(taxIds[i]);     // Column D: Tax IDs
    sheet.getRange(i + 2, column +1 ).setValue(languageName+ " ("+taxIds[i]+")");      // Column E: Language Taxs Names
  }
  createDropdownListTaxs_();
}



//createDropdownListTaxs_////////////////////////////////////////////////////////////////////////////////////
/**

*/
function createDropdownListTaxs_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet1 = ss.getSheetByName(CONFIG_Sheet.Name);
  var columnGetDropDown = sheet1.getRange(CONFIG_Sheet.Taxs_DropDown_Column_Cell).getValue()+1;
  if(DEBUG_TAXS.DEBUG_value){
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

  var targetRange = sheet2.getRange(PRODUCTS_Sheet.Plage_product_id_tax_rules_group);
  targetRange.setDataValidation(rule);
}
