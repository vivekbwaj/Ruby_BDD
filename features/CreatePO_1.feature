Feature:As a GBSS user I want to create a Purchase Order

Scenario:Create Purchase Order
   Given the user is logged in
   And wants to create POs from "C:\PO_Registration_Process_Flow_1.xls"
   Then user logs out

Scenario:Polling PO
   Given the user is logged in
   And fetches all "PO" ids from "C:\PO_Registration_Process_Flow_1.xls"
   Then validates "PO" registration to "C:\Shipment_1.txt"


