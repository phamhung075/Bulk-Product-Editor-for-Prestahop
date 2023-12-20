// Author: Dai Hung PHAM
// Date: May 24, 2023
//
// Important Notice:
// This code is the intellectual property of Dai Hung PHAM. (daihung.pham@yahoo.fr)
// Do not reuse or share this code without explicit permission.
// If you wish to use or modify this code, please contact Dai Hung PHAM for authorization.
//get_IdsProductbyRefList//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
/**
 *  This function retrieves the quantity of a product from a sheet called "GetStockList" by using the reference of the product located in column A. 
 *  It first retrieves the values in column A and filters out any null values. 
 *  It then loops through the non-null values, calls an external API with the reference of the product to get the product data in XML format, 
 * parses the XML to find the product's ID, and calls another function getlinkStockAPIbyRefList_ to import the product data. 
 *  If the imported product data matches the reference of the product in the loop, the quantity of the product is retrieved 
 * and entered into the corresponding row in column B of the "GetStockList" sheet.
*/

function get_IdsProductbyRefList(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("GetStockList");
  const valuesList = sheet.getRange('A1:A').getValues();
  let qtyP=[];

  // use the filter() method to get only the non-null values
  let nonNullValues = valuesList.filter(function(value) {
      return value != "";
  });

  Logger.log (nonNullValues);

  for (let i=0;i<nonNullValues.length;i++){
    let valuePass = nonNullValues[i];
    let apiLink = getlinkStockAPIbyRefList_(valuePass);
    Logger.log(apiLink) ;


    let timestamp = new Date().getTime();
    let headers = {'Cache-Control': 'no-cache'};
    let response = UrlFetchApp.fetch(apiLink);
    // Check the response code to see if the request was successful
    if (response.getResponseCode() != 200) {
      let re="Error getting product data: " + response.getResponseCode() + " " + response.getContentText();
      Logger.log(re);
      return re;
    }
    let xml = response.getContentText();
    Logger.log(xml);

    // Parse the XML data using the XmlService
    let document = XmlService.parse(xml);

    // Get the root element of the XML document
    let root = document.getRootElement();

    // Find the element that contains the reference you want to change
    let products = root.getChildren("products");
    let ids = [];
    for (let j = 0; j < products.length; j++) {
      let product = products[j];
      let productElements = product.getChildren("product");
      for (let k = 0; k < productElements.length; k++) {
        let productElement = productElements[k];        
        let idproduct=productElement.getAttribute("id").getValue();
        ids.push(idproduct);

        //Logger.log(idproduct);
        let thisProduct = importProductData_(idproduct);
        if(thisProduct[0]==nonNullValues[i]){
          Logger.log("qty:"+thisProduct[1]);
          //qtyP.push(thisProduct[1]);
          r="B"+(i+1);
          Logger.log(r);

          sheet.getRange(r).setValue(thisProduct[1]);

        }  
      }
    }    
  }
} 


/*
This function takes in a reference of a product as a parameter and retrieves the API link to fetch the product's data by reference. It first retrieves the site and key information from a sheet called SHEET_CONFIG and then retrieves the filter type and value from another sheet called "FILTER" which is the reference of the product. It then combines the site, filter, and key to form the API link and returns it. The returned link is then used in the main function to fetch the product data from the API.
*/






/*
 * script to export data of the named sheet as an individual csv files
 * sheet downloaded to Google Drive and then downloaded as a CSV file
 * file named according to the name of the sheet
 * original author: Michael Derazon (https://gist.github.com/mderazon/9655893)
*/

/*
This function exports the data from a sheet called "rakutenExport" to a CSV file. It first gets the active spreadsheet and the sheet to be exported. Then, it creates a new folder in Google Drive with the name of the spreadsheet and a timestamp. The function then appends ".csv" to the sheet name to create the file name. The data from the sheet is then converted to CSV format and saved in a new file in the newly created folder. Finally, it gets the download URL of the file and calls another function called "showurl" and pass the download url to it.
*/


//showurl_Rakuten_//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
/**
 *  Creates an HTML output that displays a link to download the file, and shows it in a modal dialog box using the Google Sheets UI service.
 * 
 *  @param {string} downloadURL The download link of the file to be displayed
*/
function showurl_Rakuten_(downloadURL) {
  //Change what the download button says here
  let link = HtmlService.createHtmlOutput('<a href="' + downloadURL + '">Click here to download</a>');
  SpreadsheetApp.getUi().showModalDialog(link, 'Your CSV file: rakutenExport is ready!');
}



//convertRangeToCsvFile_//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
/**
 *  Converts the data in a given sheet to a CSV file format, with semicolons as delimiters between columns.
 * 
 *  @param {string} csvFileName The name to give the CSV file
 *  @param {Sheet} sheet The sheet containing the data to be converted
 *  @return {string|undefined} The CSV file contents, or undefined if there is only one row of data in the sheet
*/

function convertRangeToCsvFile_(csvFileName, sheet) {
  // Get available data range in the spreadsheet
  const activeRange = sheet.getDataRange();

  try {
    const data = activeRange.getValues();

    // Check if there is more than one row of data
    if (data.length > 1) {
      const escapeCommas = cell => {
        if (cell.toString().indexOf(',') !== -1) {
          return `"${cell}"`;
        }
        return cell;
      };

      const isNotEmptyRow = row => row.some(cell => cell !== '');

      const processRow = row => row.map(escapeCommas).join(';');

      const csv = data
        .filter(isNotEmptyRow)
        .map(processRow)
        .join('\r\n');

      return csv;
    }

    return undefined;
  } catch (err) {
    Logger.log(err);
    Browser.msgBox(err);
  }
}



//showurl_directUpload_//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
/**
 *  Creates an HTML output that displays a link to download a file that was uploaded to Rakuten, 
 * and shows it in a modal dialog box using the Google Sheets UI service.
 * 
 *  @param {string} downloadURL The download link of the file to be displayed
*/
function showurl_directUpload_(downloadURL) {
  //Change what the download button says here
  let link = HtmlService.createHtmlOutput('<a href="' + downloadURL + '">Download the file that was uploaded on Rakuten./a>');
  SpreadsheetApp.getUi().showModalDialog(link, 'Your CSV file upload: rakutenExport is done!');
}


//saveAsCSV_Fnac//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
/**
 * Saves the contents of the "Fnac" sheet in CSV format to Google Drive.
 * If there have been no changes in the stock, displays a message stating that no update is necessary.
 * Shows a modal dialog box with a link to download the created CSV file.
*/
function saveAsCSV_Fnac() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(FNACCSV.Name);
  
  //get_stocks_Rakuten();
  Utilities.sleep(2000);

  if (sheet.getRange("A1").getValue() === "") {
    SpreadsheetApp.getUi().showModalDialog(
      HtmlService.createHtmlOutput(
        'As there has been no change in current stock, there is no need to create a new file.'),
        'No Update');
    return;
  }

  // create a folder from the name of the spreadsheet
  let folderName = ss.getName().toLowerCase().replace(/ /g,'_') + '_csv_';
  let folder = DriveApp.getFoldersByName(folderName).next();
  if (!folder) {
    folder = DriveApp.createFolder(folderName);
  }
  // append ".csv" extension to the sheet name
  let date = Utilities.formatDate(new Date(), "GMT", "ddMMyy");
  let time = Utilities.formatDate(new Date(), "GMT", "HHmm");
  fileName = sheet.getName() + "_" + date + "-" + time + ".csv";
  // convert all available sheet data to csv format
  let csvFile = convertRangeToCsvFile_(fileName, sheet);
  // create a file in the Docs List with the given name and the csv data
  let file = folder.createFile(fileName, csvFile);
  //File download
  let fileID = file.getId();
  Logger.log(fileID);
  showurl_Fnac_(file.getDownloadUrl().slice(0, -8));
}



//showurl_Fnac_//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
/**
 *  Creates an HTML output that displays a link to download the file, and shows it in a modal dialog box using the Google Sheets UI service.
 * 
 *  @param {string} downloadURL The download link of the file to be displayed
*/


function showurl_Fnac_(downloadURL) {
  //Change what the download button says here
  let link = HtmlService.createHtmlOutput('<a href="' + downloadURL + '">Click here to download</a>');
  SpreadsheetApp.getUi().showModalDialog(link, 'Your CSV file: fnacExport is ready!');
}

//saveAsCSV_Leboncoin//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
/**
 *  Saves the contents of the "Leboncoin" sheet in CSV format to Google Drive.
 *  If there have been no changes in the stock, displays a message stating that no update is necessary.
 *  Shows a modal dialog box with a link to download the created CSV file.
*/

function saveAsCSV_Leboncoin() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(LEBONCOINCSV);
  
  //get_stocks_Rakuten();
  Utilities.sleep(2000);

  if (sheet.getRange("A1").getValue() === "") {
    SpreadsheetApp.getUi().showModalDialog(
      HtmlService.createHtmlOutput(
        'As there has been no change in current stock, there is no need to create a new file.'),
        'No Update');
    return;
  }

  // create a folder from the name of the spreadsheet
  let folderName = ss.getName().toLowerCase().replace(/ /g,'_') + '_csv_';
  let folder = DriveApp.getFoldersByName(folderName).next();
  if (!folder) {
    folder = DriveApp.createFolder(folderName);
  }
  // append ".csv" extension to the sheet name
  let date = Utilities.formatDate(new Date(), "GMT", "ddMMyy");
  let time = Utilities.formatDate(new Date(), "GMT", "HHmm");
  fileName = sheet.getName() + "_" + date + "-" + time + ".csv";
  // convert all available sheet data to csv format
  let csvFile = convertRangeToCsvFile_(fileName, sheet);
  // create a file in the Docs List with the given name and the csv data
  let file = folder.createFile(fileName, csvFile);
  //File download
  let fileID = file.getId();
  Logger.log(fileID);
  showurl_Fnac_(file.getDownloadUrl().slice(0, -8));
 
}


//showurl_Leboncoin_//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
/**
 *  Creates an HTML output that displays a link to download the file, and shows it in a modal dialog box using the Google Sheets UI service.
 * 
 *  @param {string} downloadURL The download link of the file to be displayed
*/

function showurl_Leboncoin_(downloadURL) {
  //Change what the download button says here
  let link = HtmlService.createHtmlOutput('<a href="' + downloadURL + '">Click here to download</a>');
  SpreadsheetApp.getUi().showModalDialog(link, 'Your CSV file: fnacExport is ready!');
}

