require 'chronic'
require 'time'
require 'colorize'
require 'date'
require 'win32ole'

$col_dict=WIN32OLE.new('Scripting.Dictionary')
$po_field_dict=WIN32OLE.new('Scripting.Dictionary')

def createPO

 begin
  
         case $type
      
          when "GM"
              loadGMproductBrief("#{$prod_ref_id}")

          when "APP"
              loadAPPproductBrief("#{$prod_ref_id}")
          end   

      ###### Go to new commitment
      theElement("NewCommitment","Portal").flash  
      theElement("NewCommitment","Portal").click  
      sleep 8
      theElement("Shipments","POrder").click 
      sleep 5  

          case $po_Case
      
          when "case_1"
               po_case_1
               #one shipment appears in the shipment section in shipments tab

           when "case_2"
               po_case_1
               #creating PO by first creating Product
               #calling  po_case_1 because everything remains same with only extra step after opening allocation window

           when "case_3"
               po_case_1    

           when "case_4"
               po_case_1       

           when "case_5"
               po_case_1 

           when "case_1_size"
               po_case_1 

           when "case_2_size"
               po_case_1   

           when "case_3_size"
               po_case_1   

           when "case_4_size"
               po_case_1   


          else     
            puts "Invalid PO case number"
          end

       if theBrowser.div(:class=>"errorMessage-body").present?
                buttons = theBrowser.elements(:class=>"x-btn-text")
                    buttons.each do |close|
                      if close.text=="Close" then
                        close.click
                        sleep 2
                      end
                    end
           emailtext_po="Warning message pop-up appeared: Some error while entering data"
           $purchaseOrder_id=emailtext_po
           subj = "Warning message pop-up appeared: Some error while entering data"
           $status="Failed"
            puts "Warning message pop-up appeared: Some error while entering data".red       

       else
          theElement("HeaderTab","Header").click
           sleep 2
          puts theBrowser.div(:id=>"display_orderNo$0").text
          puts ""
           emailtext_po=theBrowser.div(:id=>"display_orderNo$0").text
           $purchaseOrder_id=emailtext_po
           subj = "Purchase order created"
           $status="Done"
       end

       email(emailtext_po,subj,"")
       theElement("HomeTab","Portal").click
  
   rescue Exception => e  
           emailtext_po=e.message + e.backtrace.inspect[0..300]  
           puts "#{$po_Num}".yellow
           puts "#{emailtext_po}".red
           $purchaseOrder_id=emailtext_po[0..100]
           subj = "Failed:po_Num #{$po_Num}: Data entry"
           $status="Failed data entry"

           email(emailtext_po,subj,"") 
           if isElementPresent("ShipWinClose","POrder",2)
              theElement("ShipWinClose","POrder").click
              sleep 1 
              if theBrowser.div(:class=>"x-window x-window-dlg").button(:text=>"Yes").present?
                theBrowser.div(:class=>"x-window x-window-dlg").button(:text=>"Yes").click
                sleep 1
              end
          end         
       theElement("HomeTab","Portal").click
    end

end


def getAllFieldValues
    $col_dict.each.with_index do |val,index|
          value=$po_field_dict[val]
           if value.class==Float
              cell_content=Integer(value)
          else
               cell_content=value
           end
        # puts "#{$col_dict[val]} ------> #{$po_field_dict[val]}" 
 
       case $col_dict[val]
  
      when "po_Num"
           $po_Num=cell_content
           # puts "#{$po_Num}"

      when "po_Case"
        $po_Case=cell_content
        # puts "#{$po_Case}"

      when "prod_ref_id"
        $prod_ref_id=cell_content
        # puts "#{$prod_ref_id}"

      when "type"
        $type=cell_content
        # puts "#{$type}"

      when "entity_type"
        $entity_type=cell_content
        # puts "#{$entity_type}"        

      when "distribution_Method"
        $distribution_Method=cell_content
        # puts "#{$distribution_Method}"    

      when "po_Type"
        $po_Type=cell_content
        # puts "#{$po_Type}"             

      when "supply_chain_Event"
        $supply_chain_Event=cell_content
        # puts "#{$supply_chain_Event}" 

      when "order_Type"
        $order_Type=cell_content
        # puts "#{$order_Type}"         

      when "fcl_lcl"
        $fcl_lcl=cell_content
        # puts "#{$fcl_lcl}" 

      when "dc_proc_type"
        $dc_proc_type=cell_content
        # puts "#{$dc_proc_type}" 

      when "allocation_VIC"
        $allocation_VIC=cell_content
        # puts "#{$allocation_VIC}"
 
      when "allocation_WA"
        $allocation_WA=cell_content
        # puts "#{$allocation_WA}"

      when "allocation_NSW"
        $allocation_NSW=cell_content
        # puts "#{$allocation_NSW}"
 
      when "allocation_QLD"
        $allocation_QLD=cell_content
        # puts "#{$allocation_QLD}"

      when "allocation_Total"
        $allocation_Total=cell_content
        # puts "#{$allocation_Total}"
   
      when "dcNo_VIC"
        $dcNo_VIC=cell_content
        # puts "#{$dcNo_VIC}"

      when "dcNo_NSW"
        $dcNo_NSW=cell_content
        # puts "#{$dcNo_NSW}"
 
      when "dcNo_WA"
        $dcNo_WA=cell_content
        # puts "#{$dcNo_WA}"

      when "dcNo_QLD"
        $dcNo_QLD=cell_content
        # puts "#{$dcNo_QLD}"

      when "dcNo_Total"
        $dcNo_Total=cell_content
        # puts "#{$dcNo_Total}"

      when "test_case_Description"   
        $test_case_Description=cell_content
        # puts "#{$test_case_Description}"
       
      else     
        # puts "Invalid column name or not required"
      end
    end   
    createPO
end

Given(/^wants to create POs from "([^"]*)"$/) do |sheetName|
 
  excel = WIN32OLE.new('Excel.Application')
  excel.visible = true
  workbook = excel.Workbooks.open(sheetName)

  worksheet = workbook.Worksheets('Order')
  worksheet.Activate
  num_Of_rows=worksheet.UsedRange.Rows.Count-1
  num_Of_cols=worksheet.UsedRange.Columns.Count

  for col in 2..worksheet.UsedRange.Columns.Count
    $col_dict[col]=worksheet.Cells(1,col).Value 
  end 

   $col_dict.each.with_index do |val,index|
    # puts "index: #{index} for column number: #{val} has column_name: #{$col_dict[val]}" 
   end

   for r in 2..worksheet.UsedRange.Rows.Count
       if worksheet.Cells(r,1).Value=="Y"
          $excel_RowNum=r #this is used to focus on row to enter the pass fail or output any value in a particular column
          for c in 2..worksheet.UsedRange.Columns.Count
            $po_field_dict[c]=worksheet.Cells(r,c).Value
            # puts "#{$col_dict[c]} ------> #{$po_field_dict[c]}" 
          end
          getAllFieldValues
          worksheet.Cells(r,1).Value=$status
          worksheet.Cells(r,23).Value=$purchaseOrder_id
          workbook.saved = true #Setting it to True is a simple way of preventing the "Do you wish to save..." dialog appearing when Excel is closed
          workbook.Save 
        else
         # puts "PO #{worksheet.Cells(r,2).Value} not required" 
        end 
         # puts ""
   end

 # workbook.saved = true #Setting it to True is a simple way of preventing the "Do you wish to save..." dialog appearing when Excel is closed
 # workbook.Save
 excel.ActiveWorkbook.Close(0)
 excel.Quit()
    

end



