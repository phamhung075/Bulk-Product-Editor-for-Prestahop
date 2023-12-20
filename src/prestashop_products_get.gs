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
const DEBUG_PRODUCT_GET = {
  DEBUG_value: true,
  DEBUG_request: true,
}


//getProductStockQty_//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
/**
  * Retrieves the stock quantity for a product from the provided API endpoint URL.
  * 
  * @param {string} urlStock The API endpoint URL to fetch the stock quantity
  * @return {Promise<number>} A promise that resolves with the stock quantity as a number
 */


async function getProductStockQty_(urlStock) {
  //async function getProductStockQty(urlStock="/api/stock_availables/63180"){
  const content = await getContentText_base64EncodedAuthorizationKey_(urlStock, true);
  const root = XmlService.parse(content).getRootElement();
  const stock = root.getChildren("stock_available")[0];
  const qty = stock.getChild("quantity").getText();
  if (DEBUG_PRODUCT_GET.DEBUG_value) {
    Logger.log(`getProductStockQty_() return: ${qty}`);
  }
  return qty;
}


//getProductStockID_//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
/**
  * Extracts the stock ID from the provided product content in XML format.
  * 
  * @param {string} productContent The product content in XML format
  * @return {string} The extracted stock ID value
 */


function getProductStockID_(productContent) {
  const productDocument = XmlService.parse(productContent);
  const productRoot = productDocument.getRootElement();
  const product = productRoot.getChildren("product")[0];
  return product
    .getChild("associations")
    .getChild("stock_availables")
    .getChild("stock_available")
    .getChild("id").getValue();
}


//gethrefQuantityProduct_//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
/**
  * Retrieves the xlink:href value from the associations element of a provided product.
  * 
  * @param {Object} product The product object containing the associations element
  * @return {string} The xlink:href value from the associations element
 */


const gethrefQuantityProduct_ = (product) => {
  const { getValue } = product.getChildren("associations")//;.getChild("stock_availables")//;.getChild("stock_available");//.getAttribute("xlink:href");
  if (DEBUG_PRODUCT_GET.DEBUG_value) {
    Logger.log(`getProductStockQty_() return: ${getValue()}`);
  }
  return getValue();
};




//getProductDataFromContent_//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
/**
  * Extracts various product data from the provided product and stock content in XML format.
  * 
  * @param {string} productContent The product content in XML format
  * @param {string} stockContent The stock content in XML format
  * @return {Array} An array containing the extracted product data: reference, quantity, price, name, category, condition, active status, EAN13, short description, long description, brand, and tax ID
 */

function getProductDataFromContent_(productContent, stockContent) {
  function loadDescription (){
    return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG_Sheet.Name).getRange(CONFIG_Sheet.DescriptionLoad).getValue();
  }

  const productDocument = XmlService.parse(productContent);
  const productRoot = productDocument.getRootElement();
  const product = productRoot.getChildren("product")[0];
  const productId = product.getChild("id").getText();
  const reference = product.getChild("reference").getText();
  const nameElements = product.getChildren("name");
  let ppname = "#N/A";
  for (let i = 0; i < nameElements.length; i++) {
    let nameElement = nameElements[i];
    let languageElements = nameElement.getChildren("language");
    for (let j = 0; j < languageElements.length; j++) {
      let languageElement = languageElements[j];
      if (languageElement.getAttribute("id").getValue() == "1") {
        ppname = languageElement.getText();
        break;
      }
    }
  }

  const descShortElements = product.getChildren("description_short");
  let descShort = "#N/A";
  for (let i = 0; i < descShortElements.length; i++) {
    let descShortElement = descShortElements[i];
    let languageElements = descShortElement.getChildren("language");
    for (let j = 0; j < languageElements.length; j++) {
      let languageElement = languageElements[j];
      if (languageElement.getAttribute("id").getValue() == "1") {
        descShort = languageElement.getText();
        break;
      }
    }
  }
  
  let descLong = "#N/A";

  const descLongElements = product.getChildren("description");

  for (let i = 0; i < descLongElements.length; i++) {
    let descLongElement = descLongElements[i];
    let languageElements = descLongElement.getChildren("language");
    for (let j = 0; j < languageElements.length; j++) {
      let languageElement = languageElements[j];
      if (languageElement.getAttribute("id").getValue() == "1") {
        if(loadDescription()){
          descLong = languageElement.getText();
          break;
        }
        
        else{
          descLong = "Need activate option -Load description- on -CONFIG- sheet for load"
        }
      }
    }
  }  
  const columnGetDropDown = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG_Sheet.Name).getRange(CONFIG_Sheet.Categories_DropDown_Column_Cell).getValue();
  const plage_Categories_vlookup = getColumnLetter_(columnGetDropDown-1)+"2:"+getColumnLetter_(columnGetDropDown);
  Logger.log("plage_Categories_vlookup:"+plage_Categories_vlookup);
  const stockDocument = XmlService.parse(stockContent);
  const stockRoot = stockDocument.getRootElement();
  const stock = stockRoot.getChildren("stock_available")[0];
  const quantity = stock.getChild("quantity").getText();
  const price = product.getChild("price").getText();
  const categoryid = product.getChild("id_category_default").getText();
  const category = myVLookup_(categoryid, CONFIG_Sheet.Name, plage_Categories_vlookup, 1, 2);
  const condition = product.getChild("condition").getText();
  const active = product.getChild("active").getText();
  const ean13 = product.getChild("ean13").getText();
  const idmarque = product.getChild("id_manufacturer").getText();
  const marque = myVLookup_(parseInt(idmarque), CONFIG_Sheet.Name, CONFIG_Sheet.Plage_Manufacturer_vlookup, 1, 2);
  const idTax = product.getChild("id_tax_rules_group").getText();
  const tax = myVLookup_(parseInt(idTax), CONFIG_Sheet.Name, CONFIG_Sheet.Plage_Tax_vlookup, 1, 2);
  const image = product.getChild("id_default_image").getText();
  Logger.log("image:"+image);

  return [reference, quantity, price, ppname, category, condition, parseInt(active), ean13, descShort, descLong, marque, tax, image, productId]     // < xx +^
}



//importProductData_//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
/**
  * Imports product data from the provided API endpoint URL by fetching both product and stock content.
  * 
  * @param {string} product_urlAPI The API endpoint URL to fetch the product content
  * @return {Promise<Array>} A promise that resolves with an array containing the extracted product data: 
  * reference, quantity, price, name, category, condition, active status, EAN13, short description, long description, brand, and tax ID
 */

async function importProductData_(product_urlAPI) {
  const productContent = await getContentText_base64EncodedAuthorizationKey_(product_urlAPI, true);
  const stockID = getProductStockID_(productContent);
  Logger.log("stockID: " + stockID)
  const stockAPILink = getlinkStockAPIbyID_(stockID);
  const stockContent = await getContentText_base64EncodedAuthorizationKey_(stockAPILink, true);
  const [reference, quantity, price, ppname, category, condition, active, ean13, descShort, descLong, marque, idTaxe, image, productId] = getProductDataFromContent_(productContent, stockContent);  //<< xx
  if (DEBUG_PRODUCT_GET.DEBUG_value) {
    Logger.log(`importProductData_() return: ${[reference, quantity, price, ppname, category, condition, active, ean13, descShort, descLong, marque, idTaxe, image, productId]}`); //< xx
  }
  return [reference, quantity, price, ppname, category, condition, active, ean13, descShort, descLong, marque, idTaxe, image, productId];                                                           //< xx
}





//get_RefPriceQty_byID//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
/**
  * Retrieves the reference, price, and quantity for a list of product IDs and populates the associated Google Sheet.
  * This function processes the product IDs in batches to avoid exceeding API rate limits.
  * 
  * @return {Promise<Array>} A promise that resolves with an array containing arrays of the extracted product data: reference, quantity, price, and name
 */
function removeTrailingSlash(url) {
  return url.replace(/\/$/, '');
}
async function get_RefPriceQty_byID(batchSize=10) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(PRODUCTS_Sheet.Name);
  const csheet = ss.getSheetByName(CONFIG_Sheet.Name);
  const webUrl = removeTrailingSlash(csheet.getRange(CONFIG_Sheet.Prestashop_domain_Cell).getValue())+"/"+csheet.getRange(CONFIG_Sheet.AdminAccess).getValue();
  initialiser_Sheet_Product();
  moveAtoAS_FILTER_to_PRODUCTS();
  createDropdownList_(PRESTASHOP.PRESTASHOP_Products_conditions, PRODUCTS_Sheet.Name, PRODUCTS_Sheet.Plage_product_condition);
  createDropdownList_(PRESTASHOP.PRESTASHOP_Products_active, PRODUCTS_Sheet.Name, PRODUCTS_Sheet.Plage_product_active);
  const productIds = sheet.getRange(PRODUCTS_Sheet.Plage_product_id).getValues().filter(row => row[0] !== "");

  if (DEBUG_PRODUCT_GET.DEBUG_value) {
    Logger.log(`quantity Ids in List: ${productIds.length}`);
  }
  batchSize = GET_SIZE; // number of products to fetch at a time
  const results = []; // Declare an array to store the final data

  for (let startIndex = 0; startIndex < productIds.length; startIndex += batchSize) {
    take1Credit_();
    const batchIds = productIds.slice(startIndex, startIndex + batchSize).map(row => row[0]);
    const importedData = (await Promise.allSettled(batchIds.map(id => {
      if (id !== "") {
        const apiLink = getlinkProductAPIbyID_(id);
        return importProductData_(apiLink);
      }
      return null;
    }))).filter(result => result.status === 'fulfilled' && result.value !== null); // Filter out null values and rejected promises


    importedData.forEach((data, i) => {
      if(Array.isArray(data.value)) {
        const productId = data.value.pop();
        const imageId = data.value.pop();
        const row = startIndex + i + 2; // add 2 because of header row and 0-indexing
        sheet.getRange(row, 3, 1, data.value.length).setValues([data.value]);
        if (csheet.getRange(CONFIG_Sheet.AdminAccess).getValue()!==null){
          sheet.getRange(row, 1).setValue('=HYPERLINK("'+webUrl+'/index.php/sell/catalog/products/' + productId + '"; IMAGE("https://okpasneuf.com/img/tmp/product_mini_' + imageId + '.jpg"; 1))');
        }else {
          sheet.getRange(row, 1).setValue('=IMAGE("https://okpasneuf.com/img/tmp/product_mini_' + imageId + '.jpg"; 1)');
        }
        results.push(data.value); // Push the required data to the results array
      }
    });
  }
  if (DEBUG_PRODUCT_GET.DEBUG_value) {
    Logger.log(`get_RefPriceQty_byID return: ${results}`);
  }
  return results; // Return the results array containing [reference, quantity, price, ppname...]
}


async function get_RefPriceQty_byID_Menu(batchSize=10) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(PRODUCTS_Sheet.Name);
  const csheet = ss.getSheetByName(CONFIG_Sheet.Name);
  const webUrl = removeTrailingSlash(csheet.getRange(CONFIG_Sheet.Prestashop_domain_Cell).getValue())+"/"+csheet.getRange(CONFIG_Sheet.AdminAccess).getValue();
  initialiser_Sheet_Product_keepIDs();
  createDropdownList_(PRESTASHOP.PRESTASHOP_Products_conditions, PRODUCTS_Sheet.Name, PRODUCTS_Sheet.Plage_product_condition);
  createDropdownList_(PRESTASHOP.PRESTASHOP_Products_active, PRODUCTS_Sheet.Name, PRODUCTS_Sheet.Plage_product_active);
  const productIds = sheet.getRange(PRODUCTS_Sheet.Plage_product_id).getValues().filter(row => row[0] !== "");

  if (DEBUG_PRODUCT_GET.DEBUG_value) {
    Logger.log(`quantity Ids in List: ${productIds.length}`);
  }
  batchSize = GET_SIZE; // number of products to fetch at a time
  const results = []; // Declare an array to store the final data

  for (let startIndex = 0; startIndex < productIds.length; startIndex += batchSize) {
    take1Credit_();
    const batchIds = productIds.slice(startIndex, startIndex + batchSize).map(row => row[0]);
    const importedData = (await Promise.allSettled(batchIds.map(id => {
      if (id !== "") {
        const apiLink = getlinkProductAPIbyID_(id);
        return importProductData_(apiLink);
      }
      return null;
    }))).filter(result => result.status === 'fulfilled' && result.value !== null); // Filter out null values and rejected promises


    importedData.forEach((data, i) => {
      if(Array.isArray(data.value)) {
        const productId = data.value.pop();
        const imageId = data.value.pop();
        const row = startIndex + i + 2; // add 2 because of header row and 0-indexing
        sheet.getRange(row, 3, 1, data.value.length).setValues([data.value]);
        if (csheet.getRange(CONFIG_Sheet.AdminAccess).getValue()!==null){
          sheet.getRange(row, 1).setValue('=HYPERLINK("'+webUrl+'/index.php/sell/catalog/products/' + productId + '"; IMAGE("https://okpasneuf.com/img/tmp/product_mini_' + imageId + '.jpg"; 1))');
        }else {
          sheet.getRange(row, 1).setValue('=IMAGE("https://okpasneuf.com/img/tmp/product_mini_' + imageId + '.jpg"; 1)');
        }



        results.push(data.value); // Push the required data to the results array
      }
    });
  }
  if (DEBUG_PRODUCT_GET.DEBUG_value) {
    Logger.log(`get_RefPriceQty_byID return: ${results}`);
  }
  return results; // Return the results array containing [reference, quantity, price, ppname...]
}




