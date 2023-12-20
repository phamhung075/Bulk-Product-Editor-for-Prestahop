// Author: Dai Hung PHAM
// Date: May 24, 2023
//
// Important Notice:
// This code is the intellectual property of Dai Hung PHAM. (daihung.pham@yahoo.fr)
// Do not reuse or share this code without explicit permission.
// If you wish to use or modify this code, please contact Dai Hung PHAM for authorization.
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// function make menu on Google Sheet when open
/**
 * Runs when the Google Sheet is opened and creates a custom menu on the UI with various options for the ProElectro tools library.
*/
function onOpen() {
  MENU.makeMenu();
}

/**
 * A collection of functions that create custom menus on the Google Sheet UI with various options for the ProElectro tools library.
 * Creates a custom menu on the Google Sheet UI with various options for the ProElectro tools library.
*/

const MENU = {

// Import //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////  

  makeMenu: function(){
    SpreadsheetApp.getUi()
      .createMenu('_PRODUCTS_')
      .addItem('Delete all data products', 'initialiser_Sheet_Product_keepIDs')
      .addItem('Get data products', 'get_RefPriceQty_byID_Menu')
      .addToUi();
  },


/*
//Custom Menu//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

    SpreadsheetApp.getUi()
        .createMenu('Custom Menu')
        .addItem('Show sidebar', 'showSidebar')
        .addItem('addMesSidebar', 'addMesSidebar')
      .addToUi();
*/


}




