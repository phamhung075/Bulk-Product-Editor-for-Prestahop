// Author: Dai Hung PHAM
// Date: May 24, 2023
//
// Important Notice:
// This code is the intellectual property of Dai Hung PHAM. (daihung.pham@yahoo.fr)
// Do not reuse or share this code without explicit permission.
// If you wish to use or modify this code, please contact Dai Hung PHAM for authorization.
/**
  * A configuration object for debugging purposes related to countries.
  * 
  * @type {Object}
  * @property {boolean} DEBUG_value Enables or disables logging of countries values
  * @property {boolean} DEBUG_request Enables or disables logging of countries-related API requests
 */
const DEBUG_COUNTRIES = {
  DEBUG_value: false,
  DEBUG_request: false,
}


//getObjetsAllCountryLinks_////////////////////////////////////////////////////////////////////////////////////
/**
 *  Retrieves all country links from the API and returns them as an object with country IDs as keys and links as values.
 * 
 *  @return {Promise<Object<string, string>>} A promise that resolves with an object containing country IDs as keys and links as values
*/
async function getObjetsAllCountryLinks_() {
  const url = getlinkCountriesIDsAPI_();
  const countriesContent = await getContentText_base64EncodedAuthorizationKey_(url, true);
  const document = XmlService.parse(countriesContent);
  let root = document.getRootElement();
  let countriesElement = root.getChild('countries');
  let countryElements = countriesElement.getChildren('country');
  let countryLinks = {};

  for (let i = 0; i < countryElements.length; i++) {
    let country = countryElements[i];
    let id = country.getAttribute('id').getValue();
    let xlinkNamespace = XmlService.getNamespace('xlink', 'http://www.w3.org/1999/xlink');
    let href = country.getAttribute('href', xlinkNamespace).getValue();
    countryLinks[id] = href;
  }

  return countryLinks;
}


//getCountryIds_////////////////////////////////////////////////////////////////////////////////////
/**
 *  Retrieves an array of country IDs from an object containing country links.
 * 
 *  @param {Object<string, string>} countryLinks An object containing country IDs as keys and links as values
 *  @return {Array<string>} An array of country IDs
*/
function getCountryIds_(countryLinks) {
  let ids = [];
  for (let id in countryLinks) {
    ids.push(id);
  }
  return ids;
}



//getCountryXlinks_////////////////////////////////////////////////////////////////////////////////////
/**
 *  Retrieves an array of country links from an object containing country links.
 *  @param {Object<string, string>} countryLinks An object containing country IDs as keys and links as values
 *  @return {Array<string>} An array of country links
*/
function getCountryXlinks_(countryLinks) {
  let xlinks = [];
  for (let id in countryLinks) {
    xlinks.push(countryLinks[id]);
  }
  return xlinks;
}



//getLanguageNameFromCountryLink_////////////////////////////////////////////////////////////////////////////////////
/**

*/
async function getLanguageNameFromCountryLink_(countryLink) {
  const countryContent = await getContentText_base64EncodedAuthorizationKey_(countryLink, true);
  const document = XmlService.parse(countryContent);
  let root = document.getRootElement();
  let countryElement = root.getChild('country');
  let nameElement = countryElement.getChild('name');
  let languageElements = nameElement.getChildren('language');
  
  // Assuming there's only one language element, return the language name
  if (languageElements.length > 0) {
    return languageElements[0].getText();
  }
  return "";
}



//writeCountryInfoToSheet//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
/**

*/
async function writeCountryInfoToSheet() {
  const countryLinks = await getObjetsAllCountryLinks_();
  const countryIds = getCountryIds_(countryLinks);
  const countryXlinks = getCountryXlinks_(countryLinks);
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG_Sheet.Name);
  let begin = sheet.getRange(CONFIG_Sheet.Countries_Lenght_Cell).getValue() || 0;
  sheet.getRange(CONFIG_Sheet.Countries_Name_Plage).setValue("COUNTRIES");
  for (let i = begin ; i < countryIds.length; i++) {
    const languageName = await getLanguageNameFromCountryLink_(countryXlinks[i]);
    sheet.getRange(CONFIG_Sheet.Countries_Lenght_Cell).setValue(i+1);
    sheet.getRange(i + 2, 4).setValue(countryIds[i]);     // Column D: Country IDs
    sheet.getRange(i + 2, 5).setValue(languageName);      // Column E: Language Countries Names
  }
}

