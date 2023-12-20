// Author: Dai Hung PHAM
// Date: May 24, 2023
//
// Important Notice:
// This code is the intellectual property of Dai Hung PHAM. (daihung.pham@yahoo.fr)
// Do not reuse or share this code without explicit permission.
// If you wish to use or modify this code, please contact Dai Hung PHAM for authorization.
/**
  * A configuration object for debugging purposes related to categories.
  * 
  * @type {Object}
  * @property {boolean} DEBUG_value Enables or disables logging of categories values
  * @property {boolean} DEBUG_request Enables or disables logging of categories-related API requests
 */
const DEBUG_CATEGORIES = {
  DEBUG_value: false,
  DEBUG_request: false,
}

//getObjetsAllcategoryLinks_////////////////////////////////////////////////////////////////////////////////////
/**

*/
async function getObjetsAllcategoryLinks_() {
  const url = getlinkCategoriesIDsAPI_();
  const categoriesContent = await getContentText_base64EncodedAuthorizationKey_(url, true);
  const document = XmlService.parse(categoriesContent);
  let root = document.getRootElement();
  let categoriesElement = root.getChild('categories');
  let categoryElements = categoriesElement.getChildren('category');
  let categoryLinks = {};

  for (let i = 0; i < categoryElements.length; i++) {
    let category = categoryElements[i];
    let id = category.getAttribute('id').getValue();
    let xlinkNamespace = XmlService.getNamespace('xlink', 'http://www.w3.org/1999/xlink');
    let href = category.getAttribute('href', xlinkNamespace).getValue();
    categoryLinks[id] = href;
  }

  return categoryLinks;
}


//getcategoryIds_////////////////////////////////////////////////////////////////////////////////////
/**

*/
function getcategoryIds_(categoryLinks) {
  let ids = [];
  for (let id in categoryLinks) {
    ids.push(id);
  }
  return ids;
}


//getcategoryXlinks_////////////////////////////////////////////////////////////////////////////////////
/**

*/
function getcategoryXlinks_(categoryLinks) {
  let xlinks = [];
  for (let id in categoryLinks) {
    xlinks.push(categoryLinks[id]);
  }
  return xlinks;
}


//getDataFromCategoryLink_////////////////////////////////////////////////////////////////////////////////////
/**

*/
async function getDataFromCategoryLink_(categoryLink) {
  const categoryContent = await getContentText_base64EncodedAuthorizationKey_(categoryLink, true);
  const document = XmlService.parse(categoryContent);
  const root = document.getRootElement();
  const categoryElement = root.getChild('category');
  const nameElement = categoryElement.getChild('name');
  const languageElements = nameElement.getChildren('language');
  let ctgname = "#N/A";
  // Assuming there's only one language element, return the language name
  if (languageElements.length > 0) {
    ctgname = languageElements[0].getText();
  }
  const level_depthElement = categoryElement.getChild('level_depth').getValue();
  const id_parent = categoryElement.getChild('id_parent').getValue();
  return [ctgname,level_depthElement,id_parent];
}


//writecategoryInfoToSheet////////////////////////////////////////////////////////////////////////////////////
/**

*/
async function writecategoryInfoToSheet() {
  const categoryLinks = await getObjetsAllcategoryLinks_();
  const categoryIds = getcategoryIds_(categoryLinks);
  const categoryXlinks = getcategoryXlinks_(categoryLinks);
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG_Sheet.Name);
  let begin = sheet.getRange(CONFIG_Sheet.Categories_Lenght_Cell).getValue() || 0;
  sheet.getRange(CONFIG_Sheet.Categories_Name_Plage).setValue("CATEGORIES");
  for (let i = begin ; i < categoryIds.length; i++) {
    const data = await getDataFromCategoryLink_(categoryXlinks[i]);
    Logger.log(categoryIds[i], data);
    const row = i + 2;
    sheet.getRange(CONFIG_Sheet.Categories_Lenght_Cell).setValue(i + 1);
    sheet.getRange(row, 7).setValue(categoryIds[i]); // Column D: category IDs
    //sheet.getRange(row, 9).setValue(categoryXlinks[i]);  // Column F: Xlinks
    sheet.getRange(row, 8, 1, data.length).setValues([data]);
  }
  reorganizeCategories_();
  reorganizeCategories2_();
}


//reorganizeCategories_////////////////////////////////////////////////////////////////////////////////////
/**

*/
function reorganizeCategories_() {
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  let data = sheet.getRange(CONFIG_Sheet.Plage_Categories_Origin).getValues();

  let categories = [];

  data.forEach(function (row) {
    let id = row[0];
    let name = row[1];
    let level = row[2];
    let parentId = row[3];

    if (id === "" && name === "" && level === "" && parentId === "") {
      return;
    }

    categories.push({
      id: id,
      name: name,
      level: level,
      parentId: parentId
    });
  });

  categories.sort(function (a, b) {
    if (a.parentId !== b.parentId) {
      return a.parentId - b.parentId;
    }
    return a.id - b.id;
  });

  let idToNameMap = {};
  categories.forEach(function (category) {
    idToNameMap[category.id] = category.name;
  });

  let parentIdToChildrenMap = {};
  categories.forEach(function (category) {
    if (category.parentId in parentIdToChildrenMap) {
      parentIdToChildrenMap[category.parentId].push(category);
    } else {
      parentIdToChildrenMap[category.parentId] = [category];
    }
  });

  function buildOutput(parentId, output) {
    if (parentId in parentIdToChildrenMap) {
      parentIdToChildrenMap[parentId].forEach(function (category) {
        let parentName = idToNameMap[category.parentId] || "";
        let row = [
          category.id,
          category.name + "  (" + category.id + ")" + "  ---  " + parentName + "  (" + category.parentId + ")",
          category.level,
          category.parentId
        ];
        output.push(row);
        buildOutput(category.id, output);
      });
    }
  }

  let output = [];
  buildOutput(0, output);

  output.forEach(function (row, index) {
    let rowNumber = index + 2;
    let level = parseInt(row[2], 10);
    let columnOffset = 12 + level;
    let outputRange = sheet.getRange(rowNumber, columnOffset, 1, row.length);
    outputRange.setValues([row]);
  });
}

//reorganizeCategories2_////////////////////////////////////////////////////////////////////////////////////
/**

*/
function reorganizeCategories2_() {
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  let data = sheet.getRange(CONFIG_Sheet.Plage_Categories_Origin).getValues();
  let categories = [];
  const columnWrite = sheet.getRange(CONFIG_Sheet.Plage_Categories_Origin).getColumn()+14;
  const listWrite = columnWrite+1;

  
  data.forEach(function (row) {
    let id = row[0];
    let name = row[1];
    let level = row[2];
    let parentId = row[3];

    if (id === "" && name === "" && level === "" && parentId === "") {
      return;
    }

    categories.push({
      id: id,
      name: name,
      level: level,
      parentId: parentId
    });
  });

  categories.sort(function (a, b) {
    if (a.parentId !== b.parentId) {
      return a.parentId - b.parentId;
    }
    return a.id - b.id;
  });

  let idToNameMap = {};
  categories.forEach(function (category) {
    idToNameMap[category.id] = category.name;
  });

  let parentIdToChildrenMap = {};
  categories.forEach(function (category) {
    if (category.parentId in parentIdToChildrenMap) {
      parentIdToChildrenMap[category.parentId].push(category);
    } else {
      parentIdToChildrenMap[category.parentId] = [category];
    }
  });

  function buildOutput(parentId, output) {
    if (parentId in parentIdToChildrenMap) {
      parentIdToChildrenMap[parentId].forEach(function (category) {
        let parentName = idToNameMap[category.parentId] || "";
        let row = [
          category.id,
          category.name + "  (" + category.id + ")" + "  ---  " + parentName + "  (" + category.parentId + ")",
          category.level,
          category.parentId
        ];
        output.push(row);
        buildOutput(category.id, output);
      });
    }
  }

  let output = [];
  buildOutput(0, output);

  output.forEach(function (row, index) {
    let rowNumber = index + 2;
    let level = parseInt(row[2], 10);
    let columnOffset = columnWrite;
    let outputRange = sheet.getRange(rowNumber, columnOffset, 1, row.length);
    outputRange.setValues([row]);
  });
  sheet.getRange(CONFIG_Sheet.Categories_DropDown_Column_Cell).setValue(listWrite);
  createDropdownListCategories_();
}

//createDropdownListCategories_////////////////////////////////////////////////////////////////////////////////////
/**

*/
function createDropdownListCategories_() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet1 = ss.getSheetByName(CONFIG_Sheet.Name);
  let columnGetDropDown = sheet1.getRange(CONFIG_Sheet.Categories_DropDown_Column_Cell).getValue();
  let dataRange = sheet1.getRange(2, columnGetDropDown, sheet1.getLastRow() - 1);
  
  let sheet2 = ss.getSheetByName(PRODUCTS_Sheet.Name);

  let values = dataRange.getValues();
  let flatValues = values.flat().filter(String); // Supprime les cellules vides
  
  let rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(flatValues, true)
    .setAllowInvalid(false)
    .build();

  let targetRange = sheet2.getRange(PRODUCTS_Sheet.Plage_product_category);
  targetRange.setDataValidation(rule);
}
