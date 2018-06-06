Feature:As a GBSS user I want to create products and poll for 20 minutes to check for their registration

Scenario Outline:Create products 
   Given the user is logged in
   And needs "<location_supplier>" "<prodType>" product of "<Department>" of color "<Colors>" where "<ProdCaseType>" template selection "<templateName>"
   Then the user gets the result in "C:\Products.xls" at "CreateProductAndPoll" and entity "<prodType>"

Examples:
|location_supplier|prodType|Department|Colors|ProdCaseType|templateName|
#|Local|GMProducts|520-BED COVERINGS|Black|makeSizeInactive|Generic - Various 10|
#|Local|APPProducts|265-CITY DRESSING|Gold,Black|makeSizeInactive|Lds - Regular 10|
#|Local|APPProducts|265-CITY DRESSING|Gold|OuterGreaterInner|Lds - Regular 10|
#|Local|APPProducts|160-LADIES FASHION ACCESSORIES|Gold|OuterGreaterInner|No_Variations|
#|Local|APPProducts|265-CITY DRESSING|Gold|makeSizeInactive|Lds - Regular 10|
|Local|GMProducts|980-AUDIO & PHOTO TELCO|Gold|OuterGreaterInner|Generic - Various 5|
#|Local|GMProducts|520-BED COVERINGS|Gold|OuterGreaterInner|Generic - Various 10|
#|Local|GMProducts|700-SEASONAL GIFTING|Gold|innerEqualOuter||
#|Local|GMProducts|520-BED COVERINGS|Gold|innerEqualOuter|Generic - Various 10|
#|Local|APPProducts|160-LADIES FASHION ACCESSORIES|Gold|innerEqualOuter|No_Variations|
#|Local|GMProducts|700-SEASONAL GIFTING|Gold,Blue,Black|OuterGreaterInner||
|Indent|APPProducts|245-RUNWAY TO RACK|Gold,Blue,Bronze|innerEqualOuter|Lds - Regular 3|
#|Indent|APPProducts|265-CITY DRESSING|Gold,Blue,Bronze|OuterGreaterInner|Lds - Regular 10|
#|Indent|GMProducts|520-BED COVERINGS|Gold,Blue,Black|OuterGreaterInner|Generic - Various 10|


Scenario:Polling Products
   Given the user is logged in
   And fetches all "Products" ids from "C:\Products.xls"
   Then validates "Products" registration to "C:\Products.xls"

