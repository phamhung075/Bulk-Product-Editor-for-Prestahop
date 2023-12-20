// Author: Dai Hung PHAM
// Date: May 24, 2023
//
// Important Notice:
// This code is the intellectual property of Dai Hung PHAM. (daihung.pham@yahoo.fr)
// Do not reuse or share this code without explicit permission.
// If you wish to use or modify this code, please contact Dai Hung PHAM for authorization.
const PREFIXES_REF = ["IPAD-", "TAB-","MACBOOK-Pro","MACBOOK-Air", "WATCH-"];
const FNAC_REF_SUBFIXES = "fff15";
const RAKUTEN_REF_SUBFIXES = "rrr11";
const GET_SIZE = 5; //use on: get_RefPriceQty_byID() - prestashop_products_get.gs


const RAKUTENCSV = {
  Name: "_RakutenSyncCommande_",
  Cell_first_data: "A2",
  Plage_product_ref: "A2:A",
  Plage_data: "A2:B",
};

const LEBONCOINCSV = {
  Name: "_leboncoinSyncCommande_",
  Cell_first_data: "A2",
  Plage_product_id: "A2:A",
  Plage_data: "A2:B",
};

const FNACCSV = {
  Name: "_FnacSyncCommande_",
  Cell_first_data: "A2",
  Plage_product_ref: "A2:A",
  Plage_data: "A2:B",
};

const PRESTASHOP = {
  PRESTASHOP_Products_conditions: ["new","used", "refurbished"],
  PRESTASHOP_Products_active: [0,1],
};



/**
 *  Configuration object for a sheet related to the PrestaShop API.
 *  Contains the sheet name, the cell position for the API site value, and the cell ranges for categories data.
*/
const CONFIG_Sheet = {
  Name: "_CONFIG_",
  Plage_Categories_Origin: "G2:J",
  Plage_Manufacturer_vlookup:"AA2:AB",
  Plage_Tax_vlookup:"AD2:AE",
  Categories_DropDown_Column_Cell: "B9",
  Manufactures_DropDown_Column_Cell: "B10",
  Taxs_DropDown_Column_Cell: "B11",
  Prestashop_domain_Cell: "B1",
  Prestashop_API_key_Cell: "B2",
  IdPOST: "B3",
  PassPOST: "B4",
  Credit: "B5",
  AdminAccess: "B17",
  RakutenVersion_listing_Cell: "B6",
  DescriptionLoad: "B15",
  Categories_Lenght_Cell: "H1",
  Categories_Name_Plage: "G1",
  Countries_Lenght_Cell: "E1",
  Countries_Name_Plage: "D1",
  Manufacturers_Lenght_Cell: "AB1",
  Manufacturers_Name_Plage: "AA1",
  Taxs_Lenght_Cell: "AE1",
  Taxs_Name_Plage: "AD1",
};

const PRODUCTS_Sheet = {
  Name: "_PRODUCTS_",
  Plage_data: "B2:O",
  Plage_change_history: "Q2:Q",
  Plage_product_id: "B2:B",
  Plage_product_reference: "C2:C",
  Plage_product_quantity: "D2:D",
  Plage_product_price: "E2:E",
  Plage_product_name: "F2:F",
  Plage_product_category: "G2:G",
  Plage_product_condition: "H2:H",
  Plage_product_active: "I2:I",
  Plage_product_EAN13: "J2:J",
  Plage_product_description_short: "K2:K",
  Plage_product_description_long: "L2:L",
  Plage_product_manufacturer: "M2:M",
  Plage_product_id_tax_rules_group: "N2:N",
  Plage_product_update_result: "O2:O",
  COLUMN_MESSAGE_PUT: 15,
  HEADERS: ["Reference",
            "Quantity",
            "Price excluding tax",
            "Name",
            "Categories",
            "Condition",
            "Active",
            "EAN13",
            "Description short",
            "Description long",
            "Manufactures",
            "Tax rules group",
            "Update Result",
            Date(),
          ],
  };

const FILTER_Sheet = {
  Name: "_FILTER_",
  Plage_product_id: "A2:A",
  Plage_history: "C12:C",
  Option_Filter1_Cell: "B2",
  Option_Filter2_Cell: "B3",
  Option_Filter3_Cell: "B4",
  Option_Filter4_Cell: "B5",
  Option_Filter5_Cell: "B6",
  Option_Filter6_Cell: "B7",
  Option_Filter1_Cell_Value: "C2",
  Option_Filter2_Cell_Value: "C3",
  Option_Filter3_Cell_Value: "C4",
  Option_Filter4_Cell_Value: "C5",
  Option_Filter5_Cell_Value: "C6",
  Option_Filter6_Cell_Value: "C7",
  TypefilRange: "B2:B7",
  DatafilRange: "C2:C7",
};

const ORDERS_Sheet = {
  Name: "_ORDERS_",
  Plage_product_id: "B2:B",
  Plage_product_data: "B2:J",
  Last_Order_import: "O8",
  Min_Id_Oders_import: "O2",
  Sort_by: "O4",
  Limit: "O5",
  TypeFilter: "O6",
  Number_Order_fund: "O7",
  MaxminOrders_isActive: "P2",
  MaxminOrders: "N2",
};

const LOGS_Sheet = {
  Name: "_LOGs_change",
  Header_range: "A1:J1",
};
const ORDERS_InTraitement = {
  Name: "_IDs_LastOrders_",
  Plage_order_id: "A1:A",
};

const ORDERS_Check_Stock = {
  HEADERS: ["order_id", "product_id", "product_reference", "product_name", "product_quantity_left"],
  Name: "_STOCK_Check_",
  Plage_data: "A3:E",
  Plage_order_id: "A1:A",
  IDs_colume_COMMANDE_CHECK_STOCK_range: "B3:B",
  Date_min: "H1",
  Date_max: "H2",
  Quantity_left_max: "H3",
  Quantity_left_min: "H4",
  Id_Order_min: "H5",
};
