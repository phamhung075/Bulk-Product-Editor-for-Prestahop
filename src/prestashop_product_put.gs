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
  * @property {boolean} DEBUG_value Enables or disables logging of product values (e.g., stock quantity)
  * @property {boolean} DEBUG_request Enables or disables logging of product-related API requests
 */
const DEBUG_PRODUCT_PUT = {
  DEBUG_value: false,
  DEBUG_request: false,
}

//getDataLigne_//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
/**
  * Retrieves and processes product data from the specified sheet in the Google Spreadsheet.
  * 
  * This function reads product data from the PRODUCTS_Sheet.Name in the active Google Spreadsheet.
  * It filters out rows where the id is empty and selects only the id, ref, qty, prixht, and name columns.
  * The function returns the processed data as a 2D array.
  * 
  * @return {Array[]} An array of rows, each containing selected product data in the following format [id, ref, qty, prixht, name, col6, col7, col8, col9, col10, col11, col12, col13]
 */


const getDataLigne_ = () => {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(PRODUCTS_Sheet.Name);
  const data = sheet.getDataRange().getValues();

  // Filter rows where id is not empty
  const filteredData = data.slice(1).filter(row => {
    return row[1] !== ''
  });

  // Get the id, ref, qty, prixht, name and other column
  const selectedData = filteredData.map(row => [row[1], row[2], row[3], row[4], row[5], row[6], row[7], row[8], row[9], row[10], row[11], row[12], row[13]]);
  if (DEBUG_PRODUCT_PUT.DEBUG_value) {
    Logger.log(`getDataLigne_() return: ${selectedData}`);
  }
  return selectedData;
};


//extractNumber_//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
/**
  * Extracts the number within parentheses in a given string.
  * 
  * This function takes a string containing a number inside parentheses as input,
  * extracts the number using a regular expression, and returns the extracted number
  * as an integer. If no number is found, the function returns null.
  * 
  * @param {string} data The string containing a number inside parentheses.
  * @return {number|null} The extracted number as an integer, or null if no number is found.
 */

function extractNumber_(data) {
  const result = getMatch_(/\((\d+)\)/, data);
  if (result) {
    var number = parseInt(result[1]);
    if (DEBUG_PRODUCT_PUT.DEBUG_value) {
      Logger.log(`extractNumber_() return number(int): ${number}`);
    }
    return number;
  } else {
    if (DEBUG_PRODUCT_PUT.DEBUG_value) {
      Logger.log(`error extractNumber_() return: null or number is not find`);
    }
    return null;
  }
}



//getMatch_//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
/**
  * Finds the first match of a given regular expression in a given string.
  * 
  * This function takes a regular expression and a string as input, executes the regular
  * expression on the string, and returns the first match found. If no match is found,
  * the function returns null.
  * 
  * @param {RegExp} regex - The regular expression to search for in the string.
  * @param {string} str - The string to search for the regular expression in.
  * @return {Array|null} The first match found as an array, or null if no match is found.
 */

const getMatch_ = (regex, str) => {
  const match = regex.exec(str);
  return match ? match : null;
};



//ifExist_//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
/**
  * Checks if a specified ID exists within a given range on a Google Sheet.
  * 
  * @param {number|string} id The ID to search for within the specified range
  * @param {string} plage The range within which to search for the ID
  * @param {string} sheetName The name of the Google Sheet containing the range
  * @return {boolean} true if the ID exists within the specified range, false otherwise
 */


function ifExist_(id, plage, sheetName) {
  // Ouvrir la feuille de calcul active
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  // Sélectionner la feuille avec le nom fourni
  const sheet = spreadsheet.getSheetByName(sheetName);

  // Obtenir la plage spécifiée
  const range = sheet.getRange(plage);

  // Obtenir les valeurs de la plage sous forme de tableau 2D
  const values = range.getValues();

  // Vérifier si l'ID existe en utilisant 'some' et 'flatMap'
  return values.flatMap(row => row).some(value => value === id);
}





////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////PUT API ZONE////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////





//updateProductReference_///////////////////////////////////////////////////////////////////
/**
  * Updates the product reference for a given product ID with a new reference value.
  * 
  * @param {string} productID The product ID for which the reference needs to be updated
  * @param {string} newReference The new reference value to update the product with
  * @return {Promise<string>} A promise that resolves with a success message if the update was successful or an error message if there was an issue
 */

async function updateProductReference_(productID, newReference) {
  const url = getlinkProductAPIbyID_(productID);
  if (DEBUG_PRODUCT_PUT.DEBUG_value) {
    Logger.log("----- updateProductReference_ url: " + url);
  }
  const xml = await getContentText_base64EncodedAuthorizationKey_(url, true);
  const document = XmlService.parse(xml);
  const root = document.getRootElement();
  const product = root.getChild("product");

  const reference = product.getChild("reference").getText();
  if (reference !== newReference) {
    product.getChild("manufacturer_name").detach();
    product.getChild("position_in_category").detach();
    product.getChild("quantity").detach();


    product.getChild("reference").setText(newReference);
    const updateResponse = await putXML_base64EncodedAuthorizationKey_(url, document, true);

    if (updateResponse.getResponseCode() === 200) {
      const message = `Product ${productID} reference updated successfully! (${reference} -> ${newReference})`;
      if (DEBUG_PRODUCT_PUT.DEBUG_value) {
        Logger.log("message output: " + message);
      }
      return message;
    } else {
      const error = `Error updating product reference: ${updateResponse.getResponseCode()} ${updateResponse.getContentText()}`;
      Logger.log(error);
      return error;
    }
  }
  else {
    const message = `Product ${productID} reference not change!`;
    if (DEBUG_PRODUCT_PUT.DEBUG_value) {
      Logger.log("message output: " + message);
    }
    return message;
  }
}



//putReference_//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
/**
 * This async function updates the reference for a product based on its product ID.
 * It retrieves the API URL for the product ID using the getlinkAPIbyIDProduct_ function.
 * If DEBUG_PRODUCT_PUT.DEBUG_value is true, it logs the URL to the Google Apps Script logger.
 * It then fetches the XML content from the API URL using the getContentText_base64EncodedAuthorizationKey_ function,
 * parses the XML document, and retrieves the root element and the product element. 
 * t obtains the current reference from the product element and compares it with the new reference.
 * If they are different, it updates the reference by setting the new reference.
 * It removes certain child elements from the product XML before updating the reference.
 * It then sends the updated XML document to the API using the putXML_base64EncodedAuthorizationKey_ function,
 * with the true flag indicating base64-encoded authorization key.
 * If the update is successful (response code 200), it returns a success message.
 * If there is an error during the update, it logs the error and returns an error message.
 * If the current reference is the same as the new reference, it returns a message indicating no change.
 * @param {string} productID - The product ID for which the reference will be updated.
 * @param {string} newReference - The new reference.
 * @return {Promise<string>} A promise that resolves with a message indicating the result of the reference update.
 */
async function putReference_(productID, newReference) {
  const url = getlinkAPIbyIDProduct_(productID);
  if (DEBUG_PRODUCT_PUT.DEBUG_value) {
    Logger.log("----- putReference_ url: " + url);
  }
  const xml = await getContentText_base64EncodedAuthorizationKey_(url, true);
  const document = XmlService.parse(xml);
  const root = document.getRootElement();
  const product = root.getChild("product");

  const reference = product.getChild("reference").getText();
  if (reference !== newReference) {
    product.getChild("manufacturer_name").detach();
    product.getChild("position_in_category").detach();
    product.getChild("quantity").detach();


    product.getChild("reference").setText(newReference);
    const updateResponse = await putXML_base64EncodedAuthorizationKey_(url, document, true);

    if (updateResponse.getResponseCode() === 200) {
      const message = `Product ${productID} reference updated successfully! (${reference} -> ${newReference})`;
      if (DEBUG_PRODUCT_PUT.DEBUG_value) {
        Logger.log("message output: " + message);
      }
      return message;
    } else {
      const error = `Error updating product reference: ${updateResponse.getResponseCode()} ${updateResponse.getContentText()}`;
      Logger.log(error);
      return error;
    }
  }
  else {
    const message = `Product ${productID} reference not change!`;
    if (DEBUG_PRODUCT_PUT.DEBUG_value) {
      Logger.log("message output: " + message);
    }
    return message;
  }
}




//updateProductName_//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
/**
  * Updates the product name for a given product ID with a new name value.
  * 
  * @param {string} productID The product ID for which the name needs to be updated
  * @param {string} newName The new name value to update the product with
  * @return {Promise<string>} A promise that resolves with a success message if the update was successful or an error message if there was an issue
 */

async function updateProductName_(productID, newName) {
  const url = getlinkProductAPIbyID_(productID);
  if (DEBUG_PRODUCT_PUT.DEBUG_value) {
    Logger.log("----- updateProductName_ url: " + url);
  }
  const xml = await getContentText_base64EncodedAuthorizationKey_(url, true);
  const document = XmlService.parse(xml);
  const root = document.getRootElement();
  const product = root.getChild("product");

  const nameElements = product.getChildren("name");
  for (let i = 0; i < nameElements.length; i++) {
    let nameElement = nameElements[i];
    let languageElements = nameElement.getChildren("language");
    for (let j = 0; j < languageElements.length; j++) {
      let languageElement = languageElements[j];
      if (languageElement.getAttribute("id").getValue() == "1") {
        oldName = languageElement.getText();
        //Logger.log("Old name: " + oldName);
        languageElement.setText(newName);
        //Logger.log("New name: " + languageElement.getText());
        break;
      }
    }
  }
  if (oldName !== newName) {
    product.getChild("manufacturer_name").detach();
    product.getChild("position_in_category").detach();
    product.getChild("quantity").detach();

    const updateResponse = await putXML_base64EncodedAuthorizationKey_(url, document, true);

    if (updateResponse.getResponseCode() === 200) {
      const message = `Product ${productID} name updated successfully! (${oldName} -> ${newName})`;
      if (DEBUG_PRODUCT_PUT.DEBUG_value) {
        Logger.log("message output: " + message);
      }
      return message;
    } else {
      const error = `Error updating product name: ${updateResponse.getResponseCode()} ${updateResponse.getContentText()}`;
      Logger.log(error);
      return error;
    }
  }
  else {
    const message = `Product ${productID} name not change!`;
    if (DEBUG_PRODUCT_PUT.DEBUG_value) {
      Logger.log("message output: " + message);
    }
    return message;
  }
}



//putNom_//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
/**
 * This async function updates the name for a product based on its product ID.
 * It retrieves the API URL for the product ID using the getlinkAPIbyIDProduct_ function.
 * If DEBUG_PRODUCT_PUT.DEBUG_value is true, it logs the URL to the Google Apps Script logger.
 * It then fetches the XML content from the API URL using the getContentText_base64EncodedAuthorizationKey_ function,
 * parses the XML document, and retrieves the root element and the product element.
 * It iterates through the "name" elements and finds the language element with ID "1" (English).
 * It obtains the current name from the language element and compares it with the new name.
 * If they are different, it updates the language element by setting the new name.
 * It removes certain child elements from the product XML before updating the name.
 * It then sends the updated XML document to the API using the putXML_base64EncodedAuthorizationKey_ function,
 * with the true flag indicating base64-encoded authorization key.
 * If the update is successful (response code 200), it returns a success message.
 * If there is an error during the update, it logs the error and returns an error message.
 * If the current name is the same as the new name, it returns a message indicating no change.
 * @param {string} productID - The product ID for which the name will be updated.
 * @param {string} newName - The new name.
 * @return {Promise<string>} A promise that resolves with a message indicating the result of the name update.
 */
async function putNom_(productID, newName) {
  const url = getlinkAPIbyIDProduct_(productID);
  if (DEBUG_PRODUCT_PUT.DEBUG_value) {
    Logger.log("----- putNom_ url: " + url);
  }
  const xml = await getContentText_base64EncodedAuthorizationKey_(url, true);
  const document = XmlService.parse(xml);
  const root = document.getRootElement();
  const product = root.getChild("product");

  const nameElements = product.getChildren("name");
  for (let i = 0; i < nameElements.length; i++) {
    let nameElement = nameElements[i];
    let languageElements = nameElement.getChildren("language");
    for (let j = 0; j < languageElements.length; j++) {
      let languageElement = languageElements[j];
      if (languageElement.getAttribute("id").getValue() == "1") {
        oldName = languageElement.getText();
        //Logger.log("Old name: " + oldName);
        languageElement.setText(newName);
        //Logger.log("New name: " + languageElement.getText());
        break;
      }
    }
  }
  if (oldName !== newName) {
    product.getChild("manufacturer_name").detach();
    product.getChild("position_in_category").detach();
    product.getChild("quantity").detach();

    const updateResponse = await putXML_base64EncodedAuthorizationKey_(url, document, true);

    if (updateResponse.getResponseCode() === 200) {
      const message = `Product ${productID} name updated successfully! (${oldName} -> ${newName})`;
      if (DEBUG_PRODUCT_PUT.DEBUG_value) {
        Logger.log("message output: " + message);
      }
      return message;
    } else {
      const error = `Error updating product name: ${updateResponse.getResponseCode()} ${updateResponse.getContentText()}`;
      Logger.log(error);
      return error;
    }
  }
  else {
    const message = `Product ${productID} name not change!`;
    if (DEBUG_PRODUCT_PUT.DEBUG_value) {
      Logger.log("message output: " + message);
    }
    return message;
  }
}


//updateProductDescShort_//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
/**
 * Updates the short description for a single product in the PrestaShop API.
 * This asynchronous function retrieves the XML data for a single product using its ID, updates the short description,
 * and then sends the updated XML data back to the API to update the product. If successful, it returns a success message
 * with the old and new short descriptions. If unsuccessful, it returns an error message with the response code and content text.
 * 
 * @param {number} productID - The ID of the product to update.
 * @param {string} newDescShort - The new short description for the product.
 * @return {Promise<string>} A promise that resolves with a message indicating whether the update was successful or not.
*/


async function updateProductDescShort_(productID, newDescShort) {
  const url = getlinkProductAPIbyID_(productID);
  if (DEBUG_PRODUCT_PUT.DEBUG_value) {
    Logger.log("----- updateProductDescShort_ url: " + url);
  }
  const xml = await getContentText_base64EncodedAuthorizationKey_(url, true);
  const document = XmlService.parse(xml);
  const root = document.getRootElement();
  const product = root.getChild("product");

  const nameElements = product.getChildren("description_short");
  for (let i = 0; i < nameElements.length; i++) {
    let nameElement = nameElements[i];
    let languageElements = nameElement.getChildren("language");
    for (let j = 0; j < languageElements.length; j++) {
      let languageElement = languageElements[j];
      if (languageElement.getAttribute("id").getValue() == "1") {
        oldDescShort = languageElement.getText();
        //Logger.log("Old DescShort: " + oldDescShort);
        languageElement.setText(newDescShort);
        //Logger.log("New DescShort: " + languageElement.getText());
        break;
      }
    }
  }
  if (oldDescShort !== newDescShort) {
    product.getChild("manufacturer_name").detach();
    product.getChild("position_in_category").detach();
    product.getChild("quantity").detach();

    const updateResponse = await putXML_base64EncodedAuthorizationKey_(url, document, true);

    if (updateResponse.getResponseCode() === 200) {
      const message = `Product ${productID} description short updated successfully! (${oldDescShort} -> ${newDescShort})`;
      if (DEBUG_PRODUCT_PUT.DEBUG_value) {
        Logger.log("message output: " + message);
      }
      return message;
    } else {
      const error = `Error updating product description short: ${updateResponse.getResponseCode()} ${updateResponse.getContentText()}`;
      Logger.log(error);
      return error;
    }
  }
  else {
    const message = `Product ${productID} description short not change!`;
    if (DEBUG_PRODUCT_PUT.DEBUG_value) {
      Logger.log("message output: " + message);
    }
    return message;
  }
}



//putDescriptionCourt_//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
/**
 * This async function updates the short description for a product based on its product ID.
 * It retrieves the API URL for the product ID using the getlinkAPIbyIDProduct_ function.
 * If DEBUG_PRODUCT_PUT.DEBUG_value is true, it logs the URL to the Google Apps Script logger.
 * It then fetches the XML content from the API URL using the getContentText_base64EncodedAuthorizationKey_ function,
 * parses the XML document, and retrieves the root element and the product element.
 * It iterates through the "description_short" elements and finds the language element with ID "1" (English).
 * It obtains the current short description from the language element and compares it with the new short description.
 * If they are different, it updates the language element by setting the new short description.
 * It removes certain child elements from the product XML before updating the short description.
 * It then sends the updated XML document to the API using the putXML_base64EncodedAuthorizationKey_ function,
 * with the true flag indicating base64-encoded authorization key.
 * If the update is successful (response code 200), it returns a success message.
 * If there is an error during the update, it logs the error and returns an error message.
 * If the current short description is the same as the new short description, it returns a message indicating no change.
 * @param {string} productID - The product ID for which the short description will be updated.
 * @param {string} newDescShort - The new short description.
 * @return {Promise<string>} A promise that resolves with a message indicating the result of the short description update.
 */
async function putDescriptionCourt_(productID, newDescShort) {
  const url = getlinkAPIbyIDProduct_(productID);
  if (DEBUG_PRODUCT_PUT.DEBUG_value) {
    Logger.log("----- putDescriptionCourt_ url: " + url);
  }
  const xml = await getContentText_base64EncodedAuthorizationKey_(url, true);
  const document = XmlService.parse(xml);
  const root = document.getRootElement();
  const product = root.getChild("product");

  const nameElements = product.getChildren("description_short");
  for (let i = 0; i < nameElements.length; i++) {
    let nameElement = nameElements[i];
    let languageElements = nameElement.getChildren("language");
    for (let j = 0; j < languageElements.length; j++) {
      let languageElement = languageElements[j];
      if (languageElement.getAttribute("id").getValue() == "1") {
        oldDescShort = languageElement.getText();
        //Logger.log("Old DescShort: " + oldDescShort);
        languageElement.setText(newDescShort);
        //Logger.log("New DescShort: " + languageElement.getText());
        break;
      }
    }
  }
  if (oldDescShort !== newDescShort) {
    product.getChild("manufacturer_name").detach();
    product.getChild("position_in_category").detach();
    product.getChild("quantity").detach();

    const updateResponse = await putXML_base64EncodedAuthorizationKey_(url, document, true);

    if (updateResponse.getResponseCode() === 200) {
      const message = `Product ${productID} description short updated successfully! (${oldDescShort} -> ${newDescShort})`;
      if (DEBUG_PRODUCT_PUT.DEBUG_value) {
        Logger.log("message output: " + message);
      }
      return message;
    } else {
      const error = `Error updating product description short: ${updateResponse.getResponseCode()} ${updateResponse.getContentText()}`;
      Logger.log(error);
      return error;
    }
  }
  else {
    const message = `Product ${productID} description short not change!`;
    if (DEBUG_PRODUCT_PUT.DEBUG_value) {
      Logger.log("message output: " + message);
    }
    return message;
  }
}

//updateProductDesc_//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
/**
 * Updates the full description for a single product in the PrestaShop API.
 * This asynchronous function retrieves the XML data for a single product using its ID, updates the full description,
 * and then sends the updated XML data back to the API to update the product. If successful, it returns a success message
 * with the old and new full descriptions. If unsuccessful, it returns an error message with the response code and content text.
 * 
 * @param {number} productID - The ID of the product to update.
 * @param {string} newDesc - The new full description for the product.
 * @return {Promise<string>} A promise that resolves with a message indicating whether the update was successful or not.
*/


async function updateProductDesc_(productID, newDesc) {
  const url = getlinkProductAPIbyID_(productID);
  if (DEBUG_PRODUCT_PUT.DEBUG_value) {
    Logger.log("----- updateProductDesc_ url: " + url);
  }
  const xml = await getContentText_base64EncodedAuthorizationKey_(url, true);
  const document = XmlService.parse(xml);
  const root = document.getRootElement();
  const product = root.getChild("product");

  const nameElements = product.getChildren("description");
  for (let i = 0; i < nameElements.length; i++) {
    let nameElement = nameElements[i];
    let languageElements = nameElement.getChildren("language");
    for (let j = 0; j < languageElements.length; j++) {
      let languageElement = languageElements[j];
      if (languageElement.getAttribute("id").getValue() == "1") {
        oldDesc = languageElement.getText();
        //Logger.log("Old Description: " + oldDesc);
        languageElement.setText(newDesc);
        //Logger.log("New Description: " + languageElement.getText());
        break;
      }
    }
  }
  if (oldDesc !== newDesc) {
    product.getChild("manufacturer_name").detach();
    product.getChild("position_in_category").detach();
    product.getChild("quantity").detach();

    const updateResponse = await putXML_base64EncodedAuthorizationKey_(url, document, true);

    if (updateResponse.getResponseCode() === 200) {
      const message = `Product ${productID} description updated successfully! (${oldDesc} -> ${newDesc})`;
      if (DEBUG_PRODUCT_PUT.DEBUG_value) {
        Logger.log("message output: " + message);
      }
      return message;
    } else {
      const error = `Error updating product description : ${updateResponse.getResponseCode()} ${updateResponse.getContentText()}`;
      Logger.log(error);
      return error;
    }
  }
  else {
    const message = `Product ${productID} description not change!`;
    if (DEBUG_PRODUCT_PUT.DEBUG_value) {
      Logger.log("message output: " + message);
    }
    return message;
  }
}



//putDescriptionLong_//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
/**
 * This async function updates the long description for a product based on its product ID.
 * It retrieves the API URL for the product ID using the getlinkAPIbyIDProduct_ function.
 * If DEBUG_PRODUCT_PUT.DEBUG_value is true, it logs the URL to the Google Apps Script logger.
 * It then fetches the XML content from the API URL using the getContentText_base64EncodedAuthorizationKey_ function,
 * parses the XML document, and retrieves the root element and the product element.
 * It iterates through the "description" elements and finds the language element with ID "1" (English).
 * It obtains the current long description from the language element and compares it with the new long description.
 * If they are different, it updates the language element by setting the new long description.
 * It removes certain child elements from the product XML before updating the long description.
 * It then sends the updated XML document to the API using the putXML_base64EncodedAuthorizationKey_ function,
 * with the true flag indicating base64-encoded authorization key.
 * If the update is successful (response code 200), it returns a success message.
 * If there is an error during the update, it logs the error and returns an error message.
 * If the current long description is the same as the new long description, it returns a message indicating no change.
 * @param {string} productID - The product ID for which the long description will be updated.
 * @param {string} newDesc - The new long description.
 * @return {Promise<string>} A promise that resolves with a message indicating the result of the long description update.
 */
async function putDescriptionLong_(productID, newDesc) {
  const url = getlinkAPIbyIDProduct_(productID);
  if (DEBUG_PRODUCT_PUT.DEBUG_value) {
    Logger.log("----- putDescriptionLong_ url: " + url);
  }
  const xml = await getContentText_base64EncodedAuthorizationKey_(url, true);
  const document = XmlService.parse(xml);
  const root = document.getRootElement();
  const product = root.getChild("product");

  const nameElements = product.getChildren("description");
  for (let i = 0; i < nameElements.length; i++) {
    let nameElement = nameElements[i];
    let languageElements = nameElement.getChildren("language");
    for (let j = 0; j < languageElements.length; j++) {
      let languageElement = languageElements[j];
      if (languageElement.getAttribute("id").getValue() == "1") {
        oldDesc = languageElement.getText();
        //Logger.log("Old Description: " + oldDesc);
        languageElement.setText(newDesc);
        //Logger.log("New Description: " + languageElement.getText());
        break;
      }
    }
  }
  if (oldDesc !== newDesc) {
    product.getChild("manufacturer_name").detach();
    product.getChild("position_in_category").detach();
    product.getChild("quantity").detach();

    const updateResponse = await putXML_base64EncodedAuthorizationKey_(url, document, true);

    if (updateResponse.getResponseCode() === 200) {
      const message = `Product ${productID} description updated successfully! (${oldDesc} -> ${newDesc})`;
      if (DEBUG_PRODUCT_PUT.DEBUG_value) {
        Logger.log("message output: " + message);
      }
      return message;
    } else {
      const error = `Error updating product description : ${updateResponse.getResponseCode()} ${updateResponse.getContentText()}`;
      Logger.log(error);
      return error;
    }
  }
  else {
    const message = `Product ${productID} description not change!`;
    if (DEBUG_PRODUCT_PUT.DEBUG_value) {
      Logger.log("message output: " + message);
    }
    return message;
  }
}

//updateProductPrice_//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
/**
  * Updates the product price for a given product ID with a new price value.
  * 
  * @param {string} productID The product ID for which the price needs to be updated
  * @param {number} newPrice The new price value to update the product with
  * @return {Promise<string>} A promise that resolves with a success message if the update was successful or an error message if there was an issue
 */

async function updateProductPrice_(productID, newPrice) {
  const url = getlinkProductAPIbyID_(productID);
  if (DEBUG_PRODUCT_PUT.DEBUG_value) {
    Logger.log("----- updateProductPrice_ url: " + url);
  }

  const xml = await getContentText_base64EncodedAuthorizationKey_(url, true);
  const document = XmlService.parse(xml);
  const root = document.getRootElement();
  const product = root.getChild("product");

  const oldPrice = product.getChild("price").getText();

  if (parseInt(oldPrice) !== newPrice) {
    product.getChild("manufacturer_name").detach();
    product.getChild("position_in_category").detach();
    product.getChild("quantity").detach();

    product.getChild("price").setText(newPrice);
    const updateResponse = await putXML_base64EncodedAuthorizationKey_(url, document, true);

    if (updateResponse.getResponseCode() === 200) {
      const message = `Product ${productID} price updated successfully! (${oldPrice} -> ${newPrice})`;
      if (DEBUG_PRODUCT_PUT.DEBUG_value) {
        Logger.log("message output: " + message);
      }
      return message;
    } else {
      const error = `Error updating product price: ${updateResponse.getResponseCode()} ${updateResponse.getContentText()}`;
      Logger.log(error);
      return error;
    }
  }
  else {
    const message = `Product ${productID} price not change!`;
    if (DEBUG_PRODUCT_PUT.DEBUG_value) {
      Logger.log("message output: " + message);
    }
    return message;
  }
}


//putCategories_//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
/**
 * This async function updates the HT price for a product based on its product ID.
 * It retrieves the API URL for the product ID using the getlinkAPIbyIDProduct_ function.
 * If DEBUG_PRODUCT_PUT.DEBUG_value is true, it logs the URL to the Google Apps Script logger.
 * It then fetches the XML content from the API URL using the getContentText_base64EncodedAuthorizationKey_ function,
 * parses the XML document, and retrieves the root element and the product element.
 * It obtains the current HT price from the product XML and compares it with the new HT price.
 * If they are different, it updates the XML document by setting the new HT price.
 * It removes certain child elements from the product XML before updating the HT price.
 * It then sends the updated XML document to the API using the putXML_base64EncodedAuthorizationKey_ function,
 * with the true flag indicating base64-encoded authorization key.
 * If the update is successful (response code 200), it returns a success message.
 * If there is an error during the update, it logs the error and returns an error message.
 * If the current HT price is the same as the new HT price, it returns a message indicating no change.
 * @param {string} productID - The product ID for which the HT price will be updated.
 * @param {number} newPrice - The new HT price.
 * @return {Promise<string>} A promise that resolves with a message indicating the result of the HT price update.
 */

async function putPrixHT_(productID, newPrice) {
  const url = getlinkAPIbyIDProduct_(productID);
  if (DEBUG_PRODUCT_PUT.DEBUG_value) {
    Logger.log("----- putPrixHT_ url: " + url);
  }

  const xml = await getContentText_base64EncodedAuthorizationKey_(url, true);
  const document = XmlService.parse(xml);
  const root = document.getRootElement();
  const product = root.getChild("product");

  const oldPrice = product.getChild("price").getText();

  if (parseInt(oldPrice) !== newPrice) {
    product.getChild("manufacturer_name").detach();
    product.getChild("position_in_category").detach();
    product.getChild("quantity").detach();

    product.getChild("price").setText(newPrice);
    const updateResponse = await putXML_base64EncodedAuthorizationKey_(url, document, true);

    if (updateResponse.getResponseCode() === 200) {
      const message = `Product ${productID} price updated successfully! (${oldPrice} -> ${newPrice})`;
      if (DEBUG_PRODUCT_PUT.DEBUG_value) {
        Logger.log("message output: " + message);
      }
      return message;
    } else {
      const error = `Error updating product price: ${updateResponse.getResponseCode()} ${updateResponse.getContentText()}`;
      Logger.log(error);
      return error;
    }
  }
  else {
    const message = `Product ${productID} price not change!`;
    if (DEBUG_PRODUCT_PUT.DEBUG_value) {
      Logger.log("message output: " + message);
    }
    return message;
  }
}



//updateProductCategory_//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
/**
  * Updates the product category for a given product ID with a new category value.
  * 
  * @param {string} productID The product ID for which the category needs to be updated
  * @param {number} newCategory The new category value to update the product with
  * @return {Promise<string>} A promise that resolves with a success message if the update was successful or an error message if there was an issue
 */

async function updateProductCategory_(productID, nCategory) {
  const newCategory = extractNumber_(nCategory);
  const url = getlinkProductAPIbyID_(productID);
  if (DEBUG_PRODUCT_PUT.DEBUG_value) {
    Logger.log("----- updateProductCategory_ url: " + url);
  }

  const xml = await getContentText_base64EncodedAuthorizationKey_(url, true);
  const document = XmlService.parse(xml);
  const root = document.getRootElement();
  const product = root.getChild("product");

  const oldCategory = product.getChild("id_category_default").getText();

  if (parseInt(oldCategory) !== newCategory) {
    product.getChild("manufacturer_name").detach();
    product.getChild("position_in_category").detach();
    product.getChild("quantity").detach();
    product.getChild("associations").getChild("categories").detach();

    product.getChild("id_category_default").setText(newCategory);
    const updateResponse = await putXML_base64EncodedAuthorizationKey_(url, document, true);

    if (updateResponse.getResponseCode() === 200) {
      const message = `Product ${productID} category updated successfully! (${oldCategory} -> ${newCategory})`;
      if (DEBUG_PRODUCT_PUT.DEBUG_value) {
        Logger.log("message output: " + message);
      }
      return message;
    } else {
      const error = `Error updating product category: ${updateResponse.getResponseCode()} ${updateResponse.getContentText()}`;
      Logger.log(error);
      return error;
    }
  }
  else {
    const message = `Product ${productID} category not change!`;
    if (DEBUG_PRODUCT_PUT.DEBUG_value) {
      Logger.log("message output: " + message);
    }
    return message;
  }
}


//putCategories_//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
/**
 * This async function updates the category for a product based on its product ID.
 * It extracts the numeric value from nCategory using the extractNumber_ function.
 * It retrieves the API URL for the product ID using the getlinkAPIbyIDProduct_ function.
 * If DEBUG_PRODUCT_PUT.DEBUG_value is true, it logs the URL to the Google Apps Script logger.
 * It then fetches the XML content from the API URL using the getContentText_base64EncodedAuthorizationKey_ function,
 * parses the XML document, and retrieves the root element and the product element.
 * It obtains the current category value from the product XML and compares it with the new category.
 * If they are different, it updates the XML document by setting the new category value.
 * It removes certain child elements from the product XML before updating the category.
 * It then sends the updated XML document to the API using the putXML_base64EncodedAuthorizationKey_ function,
 * with the true flag indicating base64-encoded authorization key.
 * If the update is successful (response code 200), it returns a success message.
 * If there is an error during the update, it logs the error and returns an error message.
 * If the current category is the same as the new category, it returns a message indicating no change.
 * @param {string} productID - The product ID for which the category will be updated.
 * @param {string} nCategory - The new category.
 * @return {Promise<string>} A promise that resolves with a message indicating the result of the category update.
 */

async function putCategories_(productID, nCategory) {
  const newCategory = extractNumber_(nCategory);
  const url = getlinkAPIbyIDProduct_(productID);
  if (DEBUG_PRODUCT_PUT.DEBUG_value) {
    Logger.log("----- putCategories_ url: " + url);
  }

  const xml = await getContentText_base64EncodedAuthorizationKey_(url, true);
  const document = XmlService.parse(xml);
  const root = document.getRootElement();
  const product = root.getChild("product");

  const oldCategory = product.getChild("id_category_default").getText();

  if (parseInt(oldCategory) !== newCategory) {
    product.getChild("manufacturer_name").detach();
    product.getChild("position_in_category").detach();
    product.getChild("quantity").detach();
    product.getChild("associations").getChild("categories").detach();

    product.getChild("id_category_default").setText(newCategory);
    const updateResponse = await putXML_base64EncodedAuthorizationKey_(url, document, true);

    if (updateResponse.getResponseCode() === 200) {
      const message = `Product ${productID} category updated successfully! (${oldCategory} -> ${newCategory})`;
      if (DEBUG_PRODUCT_PUT.DEBUG_value) {
        Logger.log("message output: " + message);
      }
      return message;
    } else {
      const error = `Error updating product category: ${updateResponse.getResponseCode()} ${updateResponse.getContentText()}`;
      Logger.log(error);
      return error;
    }
  }
  else {
    const message = `Product ${productID} category not change!`;
    if (DEBUG_PRODUCT_PUT.DEBUG_value) {
      Logger.log("message output: " + message);
    }
    return message;
  }
}




//updateProductManufacturer_//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
/**
 * Updates the manufacturer name for a single product in the PrestaShop API.
 * This asynchronous function retrieves the XML data for a single product using its ID, updates the manufacturer name,
 * and then sends the updated XML data back to the API to update the product. If successful, it returns a success message
 * with the old and new manufacturer IDs. If unsuccessful, it returns an error message with the response code and content text.
 * 
 * @param {number} productID - The ID of the product to update.
 * @param {number} newManufacturer - The new manufacturer ID for the product.
 * @return {Promise<string>} A promise that resolves with a message indicating whether the update was successful or not.
*/


async function updateProductManufacturer_(productID, nManufacturer) {
  const newManufacturer = extractNumber_(nManufacturer);
  const url = getlinkProductAPIbyID_(productID);
  if (DEBUG_PRODUCT_PUT.DEBUG_value) {
    Logger.log("----- updateProductManufacturer_ url: " + url);
  }

  const xml = await getContentText_base64EncodedAuthorizationKey_(url, true);
  const document = XmlService.parse(xml);
  const root = document.getRootElement();
  const product = root.getChild("product");

  const oldManufacturer = product.getChild("id_manufacturer").getText();

  if (parseInt(oldManufacturer) !== newManufacturer) {
    product.getChild("manufacturer_name").detach();
    product.getChild("position_in_category").detach();
    product.getChild("quantity").detach();

    product.getChild("id_manufacturer").setText(newManufacturer);
    const updateResponse = await putXML_base64EncodedAuthorizationKey_(url, document, true);

    if (updateResponse.getResponseCode() === 200) {
      const message = `Product ${productID} manufacturer updated successfully! (${oldManufacturer} -> ${newManufacturer})`;
      if (DEBUG_PRODUCT_PUT.DEBUG_value) {
        Logger.log("message output: " + message);
      }
      return message;
    } else {
      const error = `Error updating product manufacturer: ${updateResponse.getResponseCode()} ${updateResponse.getContentText()}`;
      Logger.log(error);
      return error;
    }
  }
  else {
    const message = `Product ${productID} manufacturer not change!`;
    if (DEBUG_PRODUCT_PUT.DEBUG_value) {
      Logger.log("message output: " + message);
    }
    return message;
  }
}

//putMarque_//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
/**
 * This async function updates the manufacturer for a product based on its product ID.
 * It extracts the numeric value from nManufacturer using the extractNumber_ function.
 * It retrieves the API URL for the product ID using the getlinkAPIbyIDProduct_ function.
 * If DEBUG_PRODUCT_PUT.DEBUG_value is true, it logs the URL to the Google Apps Script logger.
 * It then fetches the XML content from the API URL using the getContentText_base64EncodedAuthorizationKey_ function,
 * parses the XML document, and retrieves the root element and the product element.
 * It obtains the current manufacturer value from the product XML and compares it with the new manufacturer.
 * If they are different, it updates the XML document by setting the new manufacturer value.
 * It removes certain child elements from the product XML before updating the manufacturer.
 * It then sends the updated XML document to the API using the putXML_base64EncodedAuthorizationKey_ function,
 * with the true flag indicating base64-encoded authorization key.
 * If the update is successful (response code 200), it returns a success message.
 * If there is an error during the update, it logs the error and returns an error message.
 * If the current manufacturer is the same as the new manufacturer, it returns a message indicating no change.
 * @param {string} productID - The product ID for which the manufacturer will be updated.
 * @param {string} nManufacturer - The new manufacturer.
 * @return {Promise<string>} A promise that resolves with a message indicating the result of the manufacturer update.
 */
async function putMarque_(productID, nManufacturer) {
  const newManufacturer = extractNumber_(nManufacturer);
  const url = getlinkAPIbyIDProduct_(productID);
  if (DEBUG_PRODUCT_PUT.DEBUG_value) {
    Logger.log("----- putMarque_ url: " + url);
  }

  const xml = await getContentText_base64EncodedAuthorizationKey_(url, true);
  const document = XmlService.parse(xml);
  const root = document.getRootElement();
  const product = root.getChild("product");

  const oldManufacturer = product.getChild("id_manufacturer").getText();

  if (parseInt(oldManufacturer) !== newManufacturer) {
    product.getChild("manufacturer_name").detach();
    product.getChild("position_in_category").detach();
    product.getChild("quantity").detach();

    product.getChild("id_manufacturer").setText(newManufacturer);
    const updateResponse = await putXML_base64EncodedAuthorizationKey_(url, document, true);

    if (updateResponse.getResponseCode() === 200) {
      const message = `Product ${productID} manufacturer updated successfully! (${oldManufacturer} -> ${newManufacturer})`;
      if (DEBUG_PRODUCT_PUT.DEBUG_value) {
        Logger.log("message output: " + message);
      }
      return message;
    } else {
      const error = `Error updating product manufacturer: ${updateResponse.getResponseCode()} ${updateResponse.getContentText()}`;
      Logger.log(error);
      return error;
    }
  }
  else {
    const message = `Product ${productID} manufacturer not change!`;
    if (DEBUG_PRODUCT_PUT.DEBUG_value) {
      Logger.log("message output: " + message);
    }
    return message;
  }
}


//updateProductCondition_//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
/**
  * Updates the product condition for a given product ID with a new condition value.
  * 
  * @param {string} productID The product ID for which the condition needs to be updated
  * @param {string} newCondition The new condition value to update the product with (e.g., "new", "used", "refurbished")
  * @return {Promise<string>} A promise that resolves with a success message if the update was successful or an error message if there was an issue
 */


async function updateProductCondition_(productID, newCondition) {
  const url = getlinkProductAPIbyID_(productID);
  if (DEBUG_PRODUCT_PUT.DEBUG_value) {
    Logger.log("----- updateProductCondition_ url: " + url);
  }

  const xml = await getContentText_base64EncodedAuthorizationKey_(url, true);
  const document = XmlService.parse(xml);
  const root = document.getRootElement();
  const product = root.getChild("product");

  const oldCondition = product.getChild("condition").getText();

  if (oldCondition !== newCondition) {
    product.getChild("manufacturer_name").detach();
    product.getChild("position_in_category").detach();
    product.getChild("quantity").detach();

    product.getChild("condition").setText(newCondition);
    const updateResponse = await putXML_base64EncodedAuthorizationKey_(url, document, true);

    if (updateResponse.getResponseCode() === 200) {
      const message = `Product ${productID} condition updated successfully! (${oldCondition} -> ${newCondition})`;
      if (DEBUG_PRODUCT_PUT.DEBUG_value) {
        Logger.log("message output: " + message);
      }
      return message;
    } else {
      const error = `Error updating product condition: ${updateResponse.getResponseCode()} ${updateResponse.getContentText()}`;
      Logger.log(error);
      return error;
    }
  }
  else {
    const message = `Product ${productID} condition not change!`;
    if (DEBUG_PRODUCT_PUT.DEBUG_value) {
      Logger.log("message output: " + message);
    }
    return message;
  }
}


//putCondition_//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
/**
 * This async function updates the condition for a product based on its product ID.
 * It retrieves the API URL for the product ID using the getlinkAPIbyIDProduct_ function.
 * If DEBUG_PRODUCT_PUT.DEBUG_value is true, it logs the URL to the Google Apps Script logger.
 * It then fetches the XML content from the API URL using the getContentText_base64EncodedAuthorizationKey_ function,
 * parses the XML document, and retrieves the root element and the product element.
 * It obtains the current condition value from the product XML and compares it with the new condition.
 * If they are different, it updates the XML document by setting the new condition value.
 * It removes certain child elements from the product XML before updating the condition.
 * It then sends the updated XML document to the API using the putXML_base64EncodedAuthorizationKey_ function,
 * with the true flag indicating base64-encoded authorization key.
 * If the update is successful (response code 200), it returns a success message.
 * If there is an error during the update, it logs the error and returns an error message.
 * If the current condition is the same as the new condition, it returns a message indicating no change.
 * @param {string} productID - The product ID for which the condition will be updated.
 * @param {string} newCondition - The new condition.
 * @return {Promise<string>} A promise that resolves with a message indicating the result of the condition update.
 */
async function putCondition_(productID, newCondition) {
  const url = getlinkAPIbyIDProduct_(productID);
  if (DEBUG_PRODUCT_PUT.DEBUG_value) {
    Logger.log("----- putCondition_ url: " + url);
  }

  const xml = await getContentText_base64EncodedAuthorizationKey_(url, true);
  const document = XmlService.parse(xml);
  const root = document.getRootElement();
  const product = root.getChild("product");

  const oldCondition = product.getChild("condition").getText();

  if (oldCondition !== newCondition) {
    product.getChild("manufacturer_name").detach();
    product.getChild("position_in_category").detach();
    product.getChild("quantity").detach();

    product.getChild("condition").setText(newCondition);
    const updateResponse = await putXML_base64EncodedAuthorizationKey_(url, document, true);

    if (updateResponse.getResponseCode() === 200) {
      const message = `Product ${productID} condition updated successfully! (${oldCondition} -> ${newCondition})`;
      if (DEBUG_PRODUCT_PUT.DEBUG_value) {
        Logger.log("message output: " + message);
      }
      return message;
    } else {
      const error = `Error updating product condition: ${updateResponse.getResponseCode()} ${updateResponse.getContentText()}`;
      Logger.log(error);
      return error;
    }
  }
  else {
    const message = `Product ${productID} condition not change!`;
    if (DEBUG_PRODUCT_PUT.DEBUG_value) {
      Logger.log("message output: " + message);
    }
    return message;
  }
}

//updateProductActive_//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
/**
 *  Updates the active condition of a product in the PrestaShop API using the product ID and new active value provided.
 * 
 *  @param {number} productID - The ID of the product to update.
 *  @param {number} newActive - The new active value to set for the product.
 *  @return {Promise<string>} - A promise that resolves with a message indicating the success or failure of the update.
*/


async function updateProductActive_(productID, newActive) {
  const url = getlinkProductAPIbyID_(productID);
  if (DEBUG_PRODUCT_PUT.DEBUG_value) {
    Logger.log("----- updateProductActive_ url: " + url);
  }

  const xml = await getContentText_base64EncodedAuthorizationKey_(url, true);
  const document = XmlService.parse(xml);
  const root = document.getRootElement();
  const product = root.getChild("product");

  const oldActive = product.getChild("active").getText();

  if (parseInt(oldActive) !== newActive) {
    product.getChild("manufacturer_name").detach();
    product.getChild("position_in_category").detach();
    product.getChild("quantity").detach();

    product.getChild("active").setText(newActive);
    const updateResponse = await putXML_base64EncodedAuthorizationKey_(url, document, true);

    if (updateResponse.getResponseCode() === 200) {
      const message = `Product ${productID} active ON/OFF updated successfully! (${oldActive} -> ${newActive})`;
      if (DEBUG_PRODUCT_PUT.DEBUG_value) {
        Logger.log("message output: " + message);
      }
      return message;
    } else {
      const error = `Error updating product active ON/OFF: ${updateResponse.getResponseCode()} ${updateResponse.getContentText()}`;
      Logger.log(error);
      return error;
    }
  }
  else {
    const message = `Product ${productID} active ON/OFF not change!`;
    if (DEBUG_PRODUCT_PUT.DEBUG_value) {
      Logger.log("message output: " + message);
    }
    return message;
  }
}

//putActive_//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
/**
 * This async function updates the active status (ON/OFF) for a product based on its product ID.
 * It retrieves the API URL for the product ID using the getlinkAPIbyIDProduct_ function.
 * If DEBUG_PRODUCT_PUT.DEBUG_value is true, it logs the URL to the Google Apps Script logger.
 * It then fetches the XML content from the API URL using the getContentText_base64EncodedAuthorizationKey_ function,
 * parses the XML document, and retrieves the root element and the product element.
 * It obtains the current active status from the product XML and compares it with the new active status.
 * If they are different, it updates the XML document by setting the new active status.
 * It removes certain child elements from the product XML before updating the active status.
 * It then sends the updated XML document to the API using the putXML_base64EncodedAuthorizationKey_ function,
 * with the true flag indicating base64-encoded authorization key.
 * If the update is successful (response code 200), it returns a success message.
 * If there is an error during the update, it logs the error and returns an error message.
 * If the current active status is the same as the new active status, it returns a message indicating no change.
 * @param {string} productID - The product ID for which the active status will be updated.
 * @param {number} newActive - The new active status (0 or 1).
 * @return {Promise<string>} A promise that resolves with a message indicating the result of the active status update.
 */
async function putActive_(productID, newActive) {
  const url = getlinkAPIbyIDProduct_(productID);
  if (DEBUG_PRODUCT_PUT.DEBUG_value) {
    Logger.log("----- putActive_ url: " + url);
  }

  const xml = await getContentText_base64EncodedAuthorizationKey_(url, true);
  const document = XmlService.parse(xml);
  const root = document.getRootElement();
  const product = root.getChild("product");

  const oldActive = product.getChild("active").getText();

  if (parseInt(oldActive) !== newActive) {
    product.getChild("manufacturer_name").detach();
    product.getChild("position_in_category").detach();
    product.getChild("quantity").detach();

    product.getChild("active").setText(newActive);
    const updateResponse = await putXML_base64EncodedAuthorizationKey_(url, document, true);

    if (updateResponse.getResponseCode() === 200) {
      const message = `Product ${productID} active ON/OFF updated successfully! (${oldActive} -> ${newActive})`;
      if (DEBUG_PRODUCT_PUT.DEBUG_value) {
        Logger.log("message output: " + message);
      }
      return message;
    } else {
      const error = `Error updating product active ON/OFF: ${updateResponse.getResponseCode()} ${updateResponse.getContentText()}`;
      Logger.log(error);
      return error;
    }
  }
  else {
    const message = `Product ${productID} active ON/OFF not change!`;
    if (DEBUG_PRODUCT_PUT.DEBUG_value) {
      Logger.log("message output: " + message);
    }
    return message;
  }
}





//updateProductEAN13_//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
/**
 *  Updates the EAN13 code for a single product in the PrestaShop API.
 *  This asynchronous function retrieves the XML data for a single product using its ID, updates the EAN13 code,
 *  and then sends the updated XML data back to the API to update the product. If successful, it returns a success message
 *  with the old and new EAN13 codes. If unsuccessful, it returns an error message with the response code and content text.
 * 
 *  @param {number} productID - The ID of the product to update.
 *  @param {number} newEAN - The new EAN13 code for the product.
 *  @return {Promise<string>} A promise that resolves with a message indicating whether the update was successful or not.
*/


async function updateProductEAN13_(productID, newEAN) {
  const url = getlinkProductAPIbyID_(productID);
  if (DEBUG_PRODUCT_PUT.DEBUG_value) {
    Logger.log("----- updateProductEAN13_ url: " + url);
  }

  const xml = await getContentText_base64EncodedAuthorizationKey_(url, true);
  const document = XmlService.parse(xml);
  const root = document.getRootElement();
  const product = root.getChild("product");

  const oldEAN = product.getChild("ean13").getText();

  if (oldEAN !== newEAN.toString()) {
    product.getChild("manufacturer_name").detach();
    product.getChild("position_in_category").detach();
    product.getChild("quantity").detach();

    product.getChild("ean13").setText(newEAN);
    const updateResponse = await putXML_base64EncodedAuthorizationKey_(url, document, true);

    if (updateResponse.getResponseCode() === 200) {
      const message = `Product ${productID} EAN13 updated successfully! (${oldEAN} -> ${newEAN})`;
      if (DEBUG_PRODUCT_PUT.DEBUG_value) {
        Logger.log("message output: " + message);
      }
      return message;
    } else {
      const error = `Error updating product EAN13: ${updateResponse.getResponseCode()} ${updateResponse.getContentText()}`;
      Logger.log(error);
      return error;
    }
  }
  else {
    const message = `Product ${productID} EAN13 not change!`;
    if (DEBUG_PRODUCT_PUT.DEBUG_value) {
      Logger.log("message output: " + message);
    }
    return message;
  }
}



//putEAN13_//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
/**
 * This async function updates the EAN13 for a product based on its product ID.
 * It retrieves the API URL for the product ID using the getlinkAPIbyIDProduct_ function.
 * If DEBUG_PRODUCT_PUT.DEBUG_value is true, it logs the URL to the Google Apps Script logger.
 * It then fetches the XML content from the API URL using the getContentText_base64EncodedAuthorizationKey_ function,
 * parses the XML document, and retrieves the root element and the product element.
 * It obtains the current EAN13 value from the product XML and compares it with the new EAN13.
 * If they are different, it updates the XML document by setting the new EAN13 value.
 * It removes certain child elements from the product XML before updating the EAN13.
 * It then sends the updated XML document to the API using the putXML_base64EncodedAuthorizationKey_ function,
 * with the true flag indicating base64-encoded authorization key.
 * If the update is successful (response code 200), it returns a success message.
 * If there is an error during the update, it logs the error and returns an error message.
 * If the current EAN13 is the same as the new EAN13, it returns a message indicating no change.
 * @param {string} productID - The product ID for which the EAN13 will be updated.
 * @param {string|number} newEAN - The new EAN13.
 * @return {Promise<string>} A promise that resolves with a message indicating the result of the EAN13 update.
 */
async function putEAN13_(productID, newEAN) {
  const url = getlinkAPIbyIDProduct_(productID);
  if (DEBUG_PRODUCT_PUT.DEBUG_value) {
    Logger.log("----- putEAN13_ url: " + url);
  }

  const xml = await getContentText_base64EncodedAuthorizationKey_(url, true);
  const document = XmlService.parse(xml);
  const root = document.getRootElement();
  const product = root.getChild("product");

  const oldEAN = product.getChild("ean13").getText();

  if (oldEAN !== newEAN.toString()) {
    product.getChild("manufacturer_name").detach();
    product.getChild("position_in_category").detach();
    product.getChild("quantity").detach();

    product.getChild("ean13").setText(newEAN);
    const updateResponse = await putXML_base64EncodedAuthorizationKey_(url, document, true);

    if (updateResponse.getResponseCode() === 200) {
      const message = `Product ${productID} EAN13 updated successfully! (${oldEAN} -> ${newEAN})`;
      if (DEBUG_PRODUCT_PUT.DEBUG_value) {
        Logger.log("message output: " + message);
      }
      return message;
    } else {
      const error = `Error updating product EAN13: ${updateResponse.getResponseCode()} ${updateResponse.getContentText()}`;
      Logger.log(error);
      return error;
    }
  }
  else {
    const message = `Product ${productID} EAN13 not change!`;
    if (DEBUG_PRODUCT_PUT.DEBUG_value) {
      Logger.log("message output: " + message);
    }
    return message;
  }
}


//updateProductQty_//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
/**
  * Updates the product quantity for a given product ID with a new quantity value.
  * 
  * @param {string} productID The product ID for which the quantity needs to be updated
  * @param {number} newQty The new quantity value to update the product with
  * @return {Promise<string>} A promise that resolves with a success message if the update was successful or an error message if there was an issue
 */


async function updateProductQty_(productID, newQty) {
  const url = getlinkProductAPIbyID_(productID);
  if (DEBUG_PRODUCT_PUT.DEBUG_value) {
    Logger.log("----- updateProductQty_ url: " + url);
  }

  const productContent = await getContentText_base64EncodedAuthorizationKey_(url, true);
  const document = XmlService.parse(productContent);
  const root = document.getRootElement();
  const product = root.getChild("product");
  const stockID = product
    .getChild("associations")
    .getChild("stock_availables")
    .getChild("stock_available")
    .getChild("id").getValue();
  //Logger.log("stockID: "+ stockID)
  const urlstockAPI = getlinkStockAPIbyID_(stockID);
  const stockContent = await getContentText_base64EncodedAuthorizationKey_(urlstockAPI);
  const stockDocument = XmlService.parse(stockContent);
  const stockRoot = stockDocument.getRootElement();
  const stock = stockRoot.getChildren("stock_available")[0];
  const quantity = stock.getChild("quantity").getText();

  if (parseInt(quantity) !== newQty) {
    stock.getChild("quantity").setText(newQty);
    const updateResponse = await putXML_base64EncodedAuthorizationKey_(urlstockAPI, stockDocument, false);

    if (updateResponse.getResponseCode() === 200) {
      const message = `Product ${productID} quantity updated successfully! (${quantity} -> ${newQty})`;
      if (DEBUG_PRODUCT_PUT.DEBUG_value) {
        Logger.log("message output: " + message);
      }
      return message;
    } else {
      const error = `Error updating product quantity: ${updateResponse.getResponseCode()} ${updateResponse.getContentText()}`;
      Logger.log(error);
      return error;
    }
  }
  else {
    const message = `Product ${productID} quantity not change!`;
    if (DEBUG_PRODUCT_PUT.DEBUG_value) {
      Logger.log("message output: " + message);
    }
    return message;
  }
}


//putQuantite_//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
/**
 * This async function updates the quantity for a product based on its product ID.
 * It retrieves the API URL for the product ID using the getlinkAPIbyIDProduct_ function.
 * If DEBUG_PRODUCT_PUT.DEBUG_value is true, it logs the URL to the Google Apps Script logger.
 * It then fetches the product XML content from the API URL using the getContentText_base64EncodedAuthorizationKey_ function,
 * parses the XML document, and retrieves the root element and the product element.
 * It extracts the stock ID from the XML and retrieves the API URL for the stock ID using the getlinkAPIbyIDStockProduct_ function.
 * It fetches the stock XML content from the stock API URL and parses the XML document.
 * It retrieves the quantity from the stock XML.
 * It compares the current quantity with the new quantity. If they are different, it updates the stock XML document by setting the new quantity.
 * It then sends the updated stock XML document to the stock API using the putXML_base64EncodedAuthorizationKey_ function,
 * with the false flag indicating no base64-encoded authorization key.
 * If the update is successful (response code 200), it returns a success message.
 * If there is an error during the update, it logs the error and returns an error message.
 * If the current quantity is the same as the new quantity, it returns a message indicating no change.
 * @param {string} productID - The product ID for which the quantity will be updated.
 * @param {number} newQty - The new quantity.
 * @return {Promise<string>} A promise that resolves with a message indicating the result of the quantity update.
 */
async function putQuantite_(productID, newQty) {
  const url = getlinkAPIbyIDProduct_(productID);
  if (DEBUG_PRODUCT_PUT.DEBUG_value) {
    Logger.log("----- putQuantite_ url: " + url);
  }

  const productContent = await getContentText_base64EncodedAuthorizationKey_(url, true);
  const document = XmlService.parse(productContent);
  const root = document.getRootElement();
  const product = root.getChild("product");
  const stockID = product
    .getChild("associations")
    .getChild("stock_availables")
    .getChild("stock_available")
    .getChild("id").getValue();
  //Logger.log("stockID: "+ stockID)
  const urlstockAPI = getlinkAPIbyIDStockProduct_(stockID);
  const stockContent = await getContentText_base64EncodedAuthorizationKey_(urlstockAPI);
  const stockDocument = XmlService.parse(stockContent);
  const stockRoot = stockDocument.getRootElement();
  const stock = stockRoot.getChildren("stock_available")[0];
  const quantity = stock.getChild("quantity").getText();

  if (parseInt(quantity) !== newQty) {
    stock.getChild("quantity").setText(newQty);
    const updateResponse = await putXML_base64EncodedAuthorizationKey_(urlstockAPI, stockDocument, false);

    if (updateResponse.getResponseCode() === 200) {
      const message = `Product ${productID} quantity updated successfully! (${quantity} -> ${newQty})`;
      if (DEBUG_PRODUCT_PUT.DEBUG_value) {
        Logger.log("message output: " + message);
      }
      return message;
    } else {
      const error = `Error updating product quantity: ${updateResponse.getResponseCode()} ${updateResponse.getContentText()}`;
      Logger.log(error);
      return error;
    }
  }
  else {
    const message = `Product ${productID} quantity not change!`;
    if (DEBUG_PRODUCT_PUT.DEBUG_value) {
      Logger.log("message output: " + message);
    }
    return message;
  }
}




//putIdtaxRulesGroup_//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
/**
 * This async function updates the tax rules group for a product based on its product ID.
 * It extracts the numeric value from nTaxe using the extractNumber_ function.
 * It retrieves the API URL for the product ID using the getlinkAPIbyIDProduct_ function.
 * If DEBUG_PRODUCT_PUT.DEBUG_value is true, it logs the URL to the Google Apps Script logger.
 * It then fetches the XML content from the API URL using the getContentText_base64EncodedAuthorizationKey_ function,
 * parses the XML document, and retrieves the root element and the product element.
 * It obtains the current tax rules group value from the product XML and compares it with the new tax rules group.
 * If they are different, it updates the XML document by setting the new tax rules group value.
 * It removes certain child elements from the product XML before updating the tax rules group.
 * It then sends the updated XML document to the API using the putXML_base64EncodedAuthorizationKey_ function,
 * with the true flag indicating base64-encoded authorization key.
 * If the update is successful (response code 200), it returns a success message.
 * If there is an error during the update, it logs the error and returns an error message.
 * If the current tax rules group is the same as the new tax rules group, it returns a message indicating no change.
 * @param {string} productID - The product ID for which the tax rules group will be updated.
 * @param {string} nTaxe - The new tax rules group.
 * @return {Promise<string>} A promise that resolves with a message indicating the result of the tax update.
 */
async function putIdtaxRulesGroup_(productID, nTaxe) {
  const newTaxe = extractNumber_(nTaxe);
  const url = getlinkAPIbyIDProduct_(productID);
  if (DEBUG_PRODUCT_PUT.DEBUG_value) {
    Logger.log("----- putIdtaxRulesGroup_ url: " + url);
  }

  const xml = await getContentText_base64EncodedAuthorizationKey_(url, true);
  const document = XmlService.parse(xml);
  const root = document.getRootElement();
  const product = root.getChild("product");

  const oldTaxe = product.getChild("id_tax_rules_group").getText();

  if (parseInt(oldTaxe) !== newTaxe) {
    product.getChild("manufacturer_name").detach();
    product.getChild("position_in_category").detach();
    product.getChild("quantity").detach();

    product.getChild("id_tax_rules_group").setText(newTaxe);
    const updateResponse = await putXML_base64EncodedAuthorizationKey_(url, document, true);

    if (updateResponse.getResponseCode() === 200) {
      const message = `Product ${productID} taxe updated successfully! (${oldTaxe} -> ${newTaxe})`;
      if (DEBUG_PRODUCT_PUT.DEBUG_value) {
        Logger.log("message output: " + message);
      }
      return message;
    } else {
      const error = `Error updating product taxe: ${updateResponse.getResponseCode()} ${updateResponse.getContentText()}`;
      Logger.log(error);
      return error;
    }
  }
  else {
    const message = `Product ${productID} taxe not change!`;
    if (DEBUG_PRODUCT_PUT.DEBUG_value) {
      Logger.log("message output: " + message);
    }
    return message;
  }
}

//updateProductALL_//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
/**
  * Updates multiple attributes (reference, quantity, price, name) of a product for a given product ID.
  * 
  * @param {string} productID The product ID for which the attributes need to be updated
  * @param {string} newReference The new reference value to update the product with
  * @param {number} newQty The new quantity value to update the product with
  * @param {number} newPrice The new price value to update the product with
  * @param {string} newName The new name value to update the product with
  * @return {Promise<string>} A promise that resolves with a success message if the update was successful or an error message if there was an issue
 */


async function updateProductALL_(productID, newReference, newQty, newPrice, newName) {
  const url = getlinkProductAPIbyID_(productID);
  if (DEBUG_PRODUCT_PUT.DEBUG_value) {
    Logger.log("----- updateProductALL_ url: " + url);
  }
  let message = "product " + productID + ":";
  const productContent = await getContentText_base64EncodedAuthorizationKey_(url, true);
  const document = XmlService.parse(productContent);
  const root = document.getRootElement();
  const product = root.getChild("product");
  const reference = product.getChild("reference").getText();
  const oldPrice = product.getChild("price").getText();
  const nameElements = product.getChildren("name");
  for (let i = 0; i < nameElements.length; i++) {
    let nameElement = nameElements[i];
    let languageElements = nameElement.getChildren("language");
    for (let j = 0; j < languageElements.length; j++) {
      let languageElement = languageElements[j];
      if (languageElement.getAttribute("id").getValue() == "1") {
        oldName = languageElement.getText();
        //Logger.log("Old name: " + oldName);
        //Logger.log("New name: " + languageElement.getText());
        break;
      }
    }
  }
  if ((oldName !== newName) || (parseInt(oldPrice) !== newPrice) || (reference !== newReference)) {
    product.getChild("manufacturer_name").detach();
    product.getChild("position_in_category").detach();
    product.getChild("quantity").detach();

    if (reference !== newReference) {
      product.getChild("reference").setText(newReference);
      message += " reference(" + reference + " -> " + newReference + ")";
    }


    if (oldName !== newName) {
      for (let i = 0; i < nameElements.length; i++) {
        let nameElement = nameElements[i];
        let languageElements = nameElement.getChildren("language");
        for (let j = 0; j < languageElements.length; j++) {
          let languageElement = languageElements[j];
          if (languageElement.getAttribute("id").getValue() == "1") {
            //Logger.log("Old name: " + oldName);
            languageElement.setText(newName);
            //Logger.log("New name: " + languageElement.getText());
            break;
          }
        }
      }
      message += " name(" + oldName + " -> " + newName + ")";
    }

    if (parseInt(oldPrice) !== newPrice) {
      product.getChild("price").setText(newPrice);
      message += " price(" + oldPrice + " -> " + newPrice + ")";
    }
    const updateResponse = await putXML_base64EncodedAuthorizationKey_(url, document, true);
    if (updateResponse.getResponseCode() === 200) {
      message += " update successful !";
    }
    else {
      const error = `Error updating product ${productID}: ${updateResponse.getResponseCode()} ${updateResponse.getContentText()}`;
      Logger.log(error);
      return error;
    }
  }
  else {
    message = `Product ${productID} reference, name, price not change !`;
    //Logger.log("message output: " + message);
  }

  //update quantity
  const stockID = product
    .getChild("associations")
    .getChild("stock_availables")
    .getChild("stock_available")
    .getChild("id")
    .getValue();
  //Logger.log("stockID: "+ stockID)
  const urlstockAPI = getlinkStockAPIbyID_(stockID);
  const stockContent = await getContentText_base64EncodedAuthorizationKey_(urlstockAPI);
  const stockDocument = XmlService.parse(stockContent);
  const stockRoot = stockDocument.getRootElement();
  const stock = stockRoot.getChildren("stock_available")[0];
  const quantity = stock.getChild("quantity").getText();
  if (parseInt(quantity) !== newQty) {
    stock.getChild("quantity").setText(newQty);
    const updateStockResponse = await putXML_base64EncodedAuthorizationKey_(urlstockAPI, stockDocument, false);

    if (updateStockResponse.getResponseCode() === 200) {
      message += `, update quantity(${quantity} -> ${newQty}) successful !`;
      // Logger.log("message output: " + message);

    } else {
      const error = `Error updating product quantity: ${updateStockResponse.getResponseCode()} ${updateStockResponse.getContentText()}`;
      return error;
    }
  }
  else {
    message += `, quantity not change`;
    if (DEBUG_PRODUCT_PUT.DEBUG_value) {
      Logger.log("message output: " + message);
    }
  }
  return message;
}




////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////Menu EXport API ZONE////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////



//putProductID_API_reference/////////////////////////////////////////////////////////////////////////////////////////////
/**
  * Updates the reference attribute of multiple products and writes the update messages to the spreadsheet.
  * 
  * This function retrieves the data for each product from the 'PRODUCTS_Sheet.Name' sheet, updates the product reference
  * using the 'updateProductReference_' function, and writes the update message to the 'PRODUCTS_Sheet.COLUMN_MESSAGE_PUT' column
  * in the same sheet.
 */


async function putProductID_API_reference() {
  let array = getDataLigne_();
  let n_ligne = array.length;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(PRODUCTS_Sheet.Name);

  const updatePromises = array.map(async (item, i) => {
    const message = await updateProductReference_(item[0], item[1]);
    sheet.getRange(i + 2, PRODUCTS_Sheet.COLUMN_MESSAGE_PUT).setValue(message);
  });

  await Promise.all(updatePromises);
}



//putProductID_API_Qty//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
/**
  * Updates the quantity of multiple products and writes the update messages to the spreadsheet.
  * 
  * This function retrieves the data for each product from the 'PRODUCTS_Sheet.Name' sheet, updates the product quantity
  * using the 'updateProductQty_' function, and writes the update message to the 'PRODUCTS_Sheet.COLUMN_MESSAGE_PUT' column
  * in the same sheet.
 */


async function putProductID_API_Qty() {
  let array = getDataLigne_();
  let n_ligne = array.length;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(PRODUCTS_Sheet.Name);

  const updatePromises = array.map(async (item, i) => {
    const message = await updateProductQty_(item[0], item[2]);
    sheet.getRange(i + 2, PRODUCTS_Sheet.COLUMN_MESSAGE_PUT).setValue(message);
  });

  await Promise.all(updatePromises);
}


//putProductID_API_Price//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
/**
  * Updates the price of multiple products and writes the update messages to the spreadsheet.
  * 
  * This function retrieves the data for each product from the 'PRODUCTS_Sheet.Name' sheet, updates the product price
  * using the 'updateProductPrice_' function, and writes the update message to the 'PRODUCTS_Sheet.COLUMN_MESSAGE_PUT' column
  * in the same sheet.
 */


async function putProductID_API_Price() {
  let array = getDataLigne_();
  let n_ligne = array.length;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(PRODUCTS_Sheet.Name);

  const updatePromises = array.map(async (item, i) => {
    const message = await updateProductPrice_(item[0], item[3]);
    sheet.getRange(i + 2, PRODUCTS_Sheet.COLUMN_MESSAGE_PUT).setValue(message);
  });

  await Promise.all(updatePromises);
}



//putProductID_API_Name//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
/**
  * Updates the name of multiple products and writes the update messages to the spreadsheet.
  * 
  * This function retrieves the data for each product from the 'PRODUCTS_Sheet.Name' sheet, updates the product name
  * using the 'updateProductName_' function, and writes the update message to the 'PRODUCTS_Sheet.COLUMN_MESSAGE_PUT' column
  * in the same sheet.
 */


async function putProductID_API_Name() {
  let array = getDataLigne_();
  let n_ligne = array.length;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(PRODUCTS_Sheet.Name);

  const updatePromises = array.map(async (item, i) => {
    const message = await updateProductName_(item[0], item[4]);
    sheet.getRange(i + 2, PRODUCTS_Sheet.COLUMN_MESSAGE_PUT).setValue(message);
  });

  await Promise.all(updatePromises);
}



//putProductID_API_Category//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
/**
  * Updates the category for each product listed in the Google Spreadsheet using the PrestaShop API.
  * 
  * This function retrieves product data from the active Google Spreadsheet, processes it, and updates
  * the category for each product using the PrestaShop API. It then writes the update status message
  * for each product in the specified message column on the spreadsheet.
  * 
  * @return {Promise<void>} A promise that resolves when all product category updates have been attempted.
 */


async function putProductID_API_Category() {
  let array = getDataLigne_();
  let n_ligne = array.length;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(PRODUCTS_Sheet.Name);

  const updatePromises = array.map(async (item, i) => {
    const categoryId = item[5];
    const message = categoryId !== null ? await updateProductCategory_(item[0], categoryId) : 'No category found';
    sheet.getRange(i + 2, PRODUCTS_Sheet.COLUMN_MESSAGE_PUT).setValue(message);
  });

  await Promise.all(updatePromises);
}



//putProductID_API_Condition//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
/**
  * Updates the condition for each product listed in the Google Spreadsheet using the PrestaShop API.
 *
  * This function retrieves product data from the active Google Spreadsheet, processes it, and updates
  * the condition for each product using the PrestaShop API. It then writes the update status message
  * for each product in the specified message column on the spreadsheet.
  * 
  * @return {Promise<void>} A promise that resolves when all product condition updates have been attempted.
 */

async function putProductID_API_Condition() {
  let array = getDataLigne_();
  let n_ligne = array.length;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(PRODUCTS_Sheet.Name);

  const updatePromises = array.map(async (item, i) => {
    const message = await updateProductCondition_(item[0], item[6]);
    sheet.getRange(i + 2, PRODUCTS_Sheet.COLUMN_MESSAGE_PUT).setValue(message);
  });

  await Promise.all(updatePromises);
}



//putProductID_API_Active//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
/**
 *  Updates the condition for each product listed in the Google Spreadsheet using the PrestaShop API.
 *  This asynchronous function retrieves product data from the active Google Spreadsheet, processes it, and updates
 *  the condition for each product using the PrestaShop API. It then writes the update status message for each product
 *  in the specified message column on the spreadsheet. This function utilizes the getDataLigne_() and updateProductActive_()
 *  functions.
 * 
 *  @return {Promise<void>} A promise that resolves when all product condition updates have been attempted.
*/


async function putProductID_API_Active() {
  let array = getDataLigne_();
  let n_ligne = array.length;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(PRODUCTS_Sheet.Name);

  const updatePromises = array.map(async (item, i) => {
    const message = await updateProductActive_(item[0], item[7]);
    sheet.getRange(i + 2, PRODUCTS_Sheet.COLUMN_MESSAGE_PUT).setValue(message);
  });

  await Promise.all(updatePromises);
}




//putProductID_API_EAN13//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
/**
 *  Updates the EAN13 code for each product listed in the Google Spreadsheet using the PrestaShop API.
 *  This asynchronous function retrieves product data from the active Google Spreadsheet, processes it, and updates
 *  the EAN13 code for each product using the PrestaShop API. It then writes the update status message for each product
 *  in the specified message column on the spreadsheet. This function utilizes the getDataLigne_() and updateProductEAN13_()
 *  functions.
 * 
 *  @return {Promise<void>} A promise that resolves when all product EAN13 updates have been attempted.
*/


async function putProductID_API_EAN13() {
  let array = getDataLigne_();
  let n_ligne = array.length;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(PRODUCTS_Sheet.Name);

  const updatePromises = array.map(async (item, i) => {
    const message = await updateProductEAN13_(item[0], item[8]);
    sheet.getRange(i + 2, PRODUCTS_Sheet.COLUMN_MESSAGE_PUT).setValue(message);
  });

  await Promise.all(updatePromises);
}



//putProductID_API_DescShort//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
/**
 *  Updates the short description for each product listed in the Google Spreadsheet using the PrestaShop API.
 *  This asynchronous function retrieves product data from the active Google Spreadsheet, processes it, and updates
 *  the short description for each product using the PrestaShop API. It then writes the update status message for each product
 *  in the specified message column on the spreadsheet. This function utilizes the getDataLigne_() and updateProductDescShort_()
 *  functions.
 * 
 * @return {Promise<void>} A promise that resolves when all product short description updates have been attempted.
*/


async function putProductID_API_DescShort() {
  let array = getDataLigne_();
  let n_ligne = array.length;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(PRODUCTS_Sheet.Name);

  const updatePromises = array.map(async (item, i) => {
    const message = await updateProductDescShort_(item[0], item[9]);
    sheet.getRange(i + 2, PRODUCTS_Sheet.COLUMN_MESSAGE_PUT).setValue(message);
  });

  await Promise.all(updatePromises);
}


//putProductID_API_Description//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
/**
 * Updates the full description for each product listed in the Google Spreadsheet using the PrestaShop API.
 * This asynchronous function retrieves product data from the active Google Spreadsheet, processes it, and updates
 * the full description for each product using the PrestaShop API. It then writes the update status message for each product
 * in the specified message column on the spreadsheet. This function utilizes the getDataLigne_() and updateProductDesc_()
 * functions.
 * 
 * @return {Promise<void>} A promise that resolves when all product full description updates have been attempted.
*/


async function putProductID_API_Description() {
  function loadDescription (){
    return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG_Sheet.Name).getRange(CONFIG_Sheet.DescriptionLoad).getValue();
  }

  let array = getDataLigne_();
  let n_ligne = array.length;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(PRODUCTS_Sheet.Name);

  const updatePromises = array.map(async (item, i) => {
    let message = "#N/A"
    if(loadDescription()){
      message = await updateProductDesc_(item[0], item[10]);
    } else{
      message = "Need activate option -Load description- on -CONFIG- sheet for load";
    }

    sheet.getRange(i + 2, PRODUCTS_Sheet.COLUMN_MESSAGE_PUT).setValue(message);
  });

  await Promise.all(updatePromises);
}




//putProductID_API_Manufacturers//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
/**
 * Updates the manufacturer name for each product listed in the Google Spreadsheet using the PrestaShop API.
 * This asynchronous function retrieves product data from the active Google Spreadsheet, processes it, and updates
 * the manufacturer name for each product using the PrestaShop API. It then writes the update status message for each product
 * in the specified message column on the spreadsheet. This function utilizes the getDataLigne_() and updateProductManufacturer_()
 * functions.
 * 
 * @return {Promise<void>} A promise that resolves when all product manufacturer updates have been attempted.
*/


async function putProductID_API_Manufacturers() {
  let array = getDataLigne_();
  let n_ligne = array.length;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(PRODUCTS_Sheet.Name);

  const updatePromises = array.map(async (item, i) => {
    const manufacturerId = item[11];
    const message = await updateProductManufacturer_(item[0], manufacturerId);
    sheet.getRange(i + 2, PRODUCTS_Sheet.COLUMN_MESSAGE_PUT).setValue(message);
  });

  await Promise.all(updatePromises);
}

//putProductID_API_ALL//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
/**
 * Updates all attributes (reference, quantity, price, name) of multiple products based on their product ID using the updateProductALL_ function.
 * The function retrieves data for multiple products from a spreadsheet and calls the updateProductALL_ function for each product,
 * updating all its attributes. It then sets the result message for each product in the spreadsheet.
 * 
 * @return {Promise<void>} A promise that resolves when all products have been updated with their respective new attributes.
*/


async function putProductID_API_ALL() {
  let array = getDataLigne_();
  let n_ligne = array.length;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(PRODUCTS_Sheet.Name);

  const updatePromises = array.map(async (item, i) => {
    const message = await updateProductALL_(item[0], item[1], item[2], item[3], item[4]);
    sheet.getRange(i + 2, PRODUCTS_Sheet.COLUMN_MESSAGE_PUT).setValue(message);
  });

  await Promise.all(updatePromises);
}


//compileData_//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
/**
 * This function compiles data from a spreadsheet by retrieving values from a specific range.
 * It filters out empty values and non-string values, then splits the string values by comma and flattens the resulting array.
 * The final array is logged to the Google Apps Script logger.
 * 
 * @return {Array} The compiled array of values from the spreadsheet.
*/
function compileData_() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PRODUCTS_Sheet.Name);
  var range = sheet.getRange('Q2:Q');
  var values = range.getValues();

  var resultArray = values.filter(value => value[0] !== '' && typeof value[0] === 'string') 
                         .map(value => value[0].split(','))
                         .flat();

  Logger.log(resultArray); // this will log the final array to the Google Apps Script logger
  return resultArray;
}

//removeElement_//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
/**
 * This function removes a specified element from a comma-separated string.
 * It splits the string into an array based on commas, trims any leading or trailing spaces,
 * filters out the specified element, and joins the remaining elements back into a comma-separated string.
 * 
 * @param {string} str - The comma-separated string from which the element will be removed.
 * @param {string} elem - The element to be removed from the string.
 * @return {string} The modified string with the specified element removed.
*/


function removeElement_(str, elem) {
  return str.split(',')
            .map(s => s.trim()) // trim spaces
            .filter(el => el !== elem)
            .join(', ');
}


//handleDataChanges//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
/**
 * This async function handles data changes by performing a series of operations on each data change.
 * It retrieves the data changes using the compileData_ function and obtains the target sheet from the active spreadsheet.
 * For each data change, it translates the cell reference, retrieves the target cell, and removes the specified element from its value.
 * The updated cell value is then set.
 * The function also calls a specified function using the callFunctionByName_ function, passing the necessary arguments.
 * It retrieves the message cell, combines the previous message with the function result, and sets the updated message.
 * 
 * @return {Promise<void>} A promise that resolves when all data changes have been handled.
*/
async function handleDataChanges() {
  const dataChanges = compileData_();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PRODUCTS_Sheet.Name);

  for(const dataChange of dataChanges){
    const putData = translateCellReference_(dataChange);
    const cell = sheet.getRange(putData.ligne, getColumnNumber_("Q"));
    const updatedCellValue = removeElement_(cell.getValue(), dataChange);
    
    // Update the cell value after removing the element
    cell.setValue(updatedCellValue);

    // Call the function with necessary arguments
    const messageCell = sheet.getRange(putData.ligne, getColumnNumber_("O"));
    const functionResult = await callFunctionByName_(putData.function, putData.idProduct, putData.data);
    const message = [messageCell.getValue(), functionResult].join('');
    messageCell.setValue(message);
  }
}







