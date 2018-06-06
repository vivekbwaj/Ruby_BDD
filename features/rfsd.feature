Feature:As a GBSS user I want to do RFSD

Scenario:Create rfsd
   Given the user is logged in
   And wants to rfsd from "C:\PO_Registration_Process_Flow.xls"
   Then user logs out

Scenario:Polling PO
   Given the user is logged in
   And fetches all "PO" ids from "C:\PO_Registration_Process_Flow.xls"
   Then validates "PO" registration to "C:\Shipment.txt"
