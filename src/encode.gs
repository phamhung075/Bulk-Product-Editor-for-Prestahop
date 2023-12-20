// Author: Dai Hung PHAM
// Date: May 24, 2023
//
// Important Notice:
// This code is the intellectual property of Dai Hung PHAM. (daihung.pham@yahoo.fr)
// Do not reuse or share this code without explicit permission.
// If you wish to use or modify this code, please contact Dai Hung PHAM for authorization.
//getContentText_base64EncodedAuthorizationKey_//////////////////////////////////////////////////////////////////////////////////////
/**
 *  Retrieves the content text of a URL using a base64-encoded authorization key.
 * 
 *  @param {string} url The URL to retrieve the content of
 *  @param {boolean} withTimestamp Whether to include a timestamp in the request
 *  @return {Promise<string>} A promise that resolves with the content text of the URL
*/
async function getContentText_base64EncodedAuthorizationKey_(url, withTimestamp = true) {
  const encodedUrl = Utilities.base64Encode(url);
  const content = await fetchContent_(encodedUrl, withTimestamp);
  return content;
}


//fetchContent_////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
/**
 *  Fetches the content of a URL using a base64-encoded authorization key and optional timestamp.
 * 
 *  @param {string} encodedUrl The base64-encoded URL to fetch the content of
 *  @param {boolean} withTimestamp Whether to include a timestamp in the request
 *  @return {Promise<string>} A promise that resolves with the content of the URL
 *  @throws {Error} If the API endpoint returns an error response
*/
async function fetchContent_(encodedUrl, withTimestamp) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG_Sheet.Name);
  const apiKey = sheet.getRange('B2').getValue();
  const base64EncodedAuthorizationKey = Utilities.base64EncodeWebSafe(apiKey + ':');
  const headers = {
    'Authorization': 'Basic ' + base64EncodedAuthorizationKey
  };
  const options = {
    "method": "GET",
    "headers": headers,
    "muteHttpExceptions": true
  };
  let url = Utilities.newBlob(Utilities.base64Decode(encodedUrl)).getDataAsString();
  if (withTimestamp) {
    url += `?t=${new Date().getTime()}`;
  }
  const response = await UrlFetchApp.fetch(url, options);
  if (response.getResponseCode() !== 200) {
  throw new Error(`The API endpoint returned an error response: ${response.getContentText()}`);
  }
  return response.getContentText();
}


//putXML_base64EncodedAuthorizationKey_/////////////////////////////////////////////////////////////////////////////////////
/**
 *  Sends a PUT request with an XML document to a URL using a base64-encoded authorization key and optional timestamp.
 * 
 *  @param {string} url The URL to send the PUT request to
 *  @param {XmlService.Document} document The XML document to send in the payload
 *  @param {boolean} withTimestamp Whether to include a timestamp in the request
 *  @return {Promise<HTTPResponse>} A promise that resolves with the HTTPResponse object of the request
 *  @throws {Error} If the URL parameter is not provided
*/
async function putXML_base64EncodedAuthorizationKey_(url, document, withTimestamp = false) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG_Sheet.Name);
  const vSite = sheet.getRange('B1').getValue();
  const apiKey = sheet.getRange('B2').getValue();
  const timestamp = new Date().getTime();

  if (withTimestamp) {
    url += `?t=${timestamp}`;
  }

  const base64EncodedAuthorizationKey = Utilities.base64EncodeWebSafe(apiKey + ':');
  const headers = {
    'Authorization': 'Basic ' + base64EncodedAuthorizationKey
  };

  if (!url) {
    throw new Error("The URL parameter is required.");
  }
  const modifiedXml = XmlService.getPrettyFormat().format(document);

  const options = {
    headers: headers,
    method: "PUT",
    contentType: "application/xml",
    payload: modifiedXml,
    muteHttpExceptions: true,
  };
  return UrlFetchApp.fetch(url, options);
}

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
/*
//"/api/products/140567"
//"/api/products?filter[reference]=%25[test]%25"
// /api/stock_availables/93556
*/