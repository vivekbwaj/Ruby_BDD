Feature:As a GBSS user I want to create products and poll for 20 minutes to check for their registration and the create PO for these products

Scenario Outline:Create products 
  Given the user is logged in
  And needs "<location_supplier>" "<prodType>" product of "<Department>" of color "<Colors>" where "<ProdCaseType>" template selection "<templateName>"
  Then the user gets the result in "C:\PO_Registration_Process_Flow.xls" at "<PO_Num>" and entity "<prodType>"

Examples:
|location_supplier|prodType|Department|Colors|ProdCaseType|templateName|PO_Num|
|Indent|GMProducts|520-BED COVERINGS|Gold|OuterGreaterInner||1,2,3,4,5,6|
#|Local|GMProducts|520-BED COVERINGS|Gold,Blue,Black|OuterGreaterInner|Generic - Various 10|1, 10, 19, 20, 23, 24, 25, 26, 27, 79, 80, 81, 82, 92, 146,  147,  148,  150|
#|Local|APPProducts|265-CITY DRESSING|Gold,Green,Grey|makeSizeInactive|Lds - Regular 10|29,  30, 57, 58, 59, 76, 77, 78, 83, 84, 85, 86, 87, 88, 143,  144|
#|Indent|GMProducts|520-BED COVERINGS|Gold,Blue,Black|OuterGreaterInner|Generic - Various 10|2,  3,  4,  5,  6,  7,  8,  9,  11, 12, 13, 14, 15, 16, 17, 18, 21, 22, 28, 31, 32, 33, 34, 35, 36, 37, 38, 39, 40, 89, 90, 91, 117,  118,  119,  120,  121,  122,  123,  124,  125,  126,  127,  128,  129,  130,  131,  132,  133,  134,  135,  136,  137,  138,  139,  140,  145,  149|
#|Indent|APPProducts|265-CITY DRESSING|Gold,Green,Grey|makeSizeInactive|Lds - Regular 10|41, 42, 43, 44, 45, 46, 47, 48, 49, 50, 51, 52, 53, 54, 55, 56, 60, 61, 62, 63, 64, 65, 66, 67, 68, 69, 70, 71, 72, 73, 74, 75, 93, 94, 95, 96, 97, 98, 99, 100,  101,  102,  103,  104,  105,  106,  107,  108,  109,  110,  111,  112,  113,  114,  115,  116,  141,  142|


#|Local|APPProducts|265-CITY DRESSING|Gold,Green,Grey|makeSizeInactive|Lds - Regular 10|20,  21, 38, 39, 40, 41, 42, 43, 50, 51, 52, 53, 54, 55|
#|Indent|GMProducts|520-BED COVERINGS|Gold,Blue,Black|OuterGreaterInner|Generic - Various 10|2,  3,  4,  5,  6,  7,  8,  9,  10, 11, 12, 13, 14, 15, 16, 19, 44, 45, 46, 47, 48, 49, 78, 79, 80, 81, 82, 83, 84, 85, 86, 117,  118,  119,  120,  121,  122,  123,  124,  125,  126,  127,  128|
#|Indent|APPProducts|265-CITY DRESSING|Gold,Green,Grey|makeSizeInactive|Lds - Regular 10|22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 37, 62, 63, 64, 65, 66, 67, 68, 69, 70, 71, 72, 73, 74, 75, 76, 77, 93, 94, 95, 96, 97, 98, 99, 100,  101,  102,  103,  104,  105,  106,  107,  108,  109,  110,  111,  112,  113,  114,  115,  116|
#|Local|GMProducts|520-BED COVERINGS|Gold,Blue,Black|OuterGreaterInner|Generic - Various 10|17,  18, 56, 57, 58, 59, 60, 61, 87, 88, 89, 90, 91, 92|


Scenario:Polling Products
   Given the user is logged in
   And fetches all "ProductsForPO" ids from "C:\PO_Registration_Process_Flow.xls"
   Then validates "ProductsForPO" registration to "C:\PO_Registration_Process_Flow.xls"
   Then user logs out

Scenario:Create Purchase Order
   Given the user is logged in
   And wants to create POs from "C:\PO_Registration_Process_Flow.xls"
   Then user logs out

Scenario:Clear products id dictionary
   Given the user clears the ProdIdDictionary

Scenario:Polling PO
   Given the user is logged in
   And fetches all "PO" ids from "C:\PO_Registration_Process_Flow.xls"
   Then validates "PO" registration to "C:\Shipment.txt"