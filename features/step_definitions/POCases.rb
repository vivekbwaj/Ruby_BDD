require 'chronic'
require 'time'
require 'colorize'
require 'date'
require 'win32ole'



  def po_case_1

     dc_date=DateTime.now.next_month.next_month.strftime "%d/%m/%Y"
     show_date=DateTime.now.next_month.next_month.next_month.strftime "%d/%m/%Y"

     theBrowser.div(:class=>"x-grid3-cell-inner x-grid3-col-onShowDate").click
     sleep 1
     theBrowser.send_keys [:control, 'a'], :backspace
     theBrowser.input(:id=>"onShowDate/cDateField").send_keys show_date
     sleep 1
     theBrowser.div(:class=>"x-grid3-cell-inner x-grid3-col-dcDueDate").click
     sleep 1
     theBrowser.send_keys [:control, 'a'], :backspace
     theBrowser.input(:id=>"dcDueDate/cDateField").send_keys dc_date
     sleep 1

     theBrowser.div(:class=>"x-grid3-cell-inner x-grid3-col-advertisingDate").click
     sleep 1
     theBrowser.input(:id=>"advertisingDate/cDateField").send_keys show_date
     sleep 1     

    #####enter notes and comments######
      theBrowser.div(:id=>"cbx-orderShipments-order__orderShipments__notes-0-33").click
      sleep 1
      theBrowser.send_keys [:control, 'a'], :backspace
      theBrowser.send_keys "Test PO"

      theBrowser.div(:id=>"cbx-orderShipments-order__orderShipments__comments-0-40").click
      sleep 1
      theBrowser.send_keys "Test PO"

    if $dc_proc_type!="Conveyable"
      theBrowser.execute_script("document.getElementsByClassName('x-grid3-cell-inner x-grid3-col-dcProcTypeId')[0].click()")
      theBrowser.send_keys :down
      sleep 1
      theBrowser.send_keys $dc_proc_type
      sleep 1
      theBrowser.div(:class=>"x-combo-list-item",:text=>"#{$dc_proc_type}").click
    end
sleep 1
    #### shipments page fields
    if $supply_chain_Event!="NA"
      theBrowser.execute_script("document.getElementsByClassName('x-grid3-cell-inner x-grid3-col-supplyChainEventId')[0].click()")
      sleep 1
      theBrowser.send_keys [:control, 'a'], :backspace
      theBrowser.send_keys $supply_chain_Event
      sleep 1
      theBrowser.div(:class=>"x-combo-list-item",:text=>"#{$supply_chain_Event}").click
    end

      theBrowser.execute_script("document.getElementsByClassName('x-grid3-cell-inner x-grid3-col-distributionMethodId')[0].click()")
      sleep 2
      theBrowser.send_keys $distribution_Method

    if $po_Type!="NA"
      theBrowser.execute_script("document.getElementsByClassName('x-grid3-cell-inner x-grid3-col-poTypeId')[0].click()")
      sleep 2
      theBrowser.send_keys $po_Type
    end
      
    if $order_Type!="NA"
      theBrowser.execute_script("document.getElementsByClassName('x-grid3-cell-inner x-grid3-col-orderTypeId')[0].click()")
      sleep 2
      theBrowser.send_keys $order_Type
    end

    if $fcl_lcl!="NA"
      theBrowser.execute_script("document.getElementsByClassName('x-grid3-cell-inner x-grid3-col-lclFclId')[0].click()")
      sleep 2
      theBrowser.send_keys $fcl_lcl    
    end

      theBrowser.execute_script("document.getElementsByClassName('x-grid3-cell-inner x-grid3-col-vasCodeId')[0].click()")
      sleep 2
      theBrowser.send_keys "No"  

    ########################
    theElement("Save","Portal").click

     sleep 8
    buttons = theBrowser.elements(:class=>"inline_button")  
      buttons.each do |shipBtn|
          if shipBtn.attribute_value("value")=="DC Allocation" then
            shipBtn.click
            sleep 4
          end
        end
     $size_PO=false
          case $po_Case
      
           when "case_2"
                delete_colours_allocation
           when "case_3"
               change_casePK    

           when "case_4"
               deleteColor_Addagain   

           when "case_1_size"
                  $size_PO=true    
                  theBrowser.input(:id=>"colorSizeChoiceS").click
                  theBrowser.div(:class=>"x-window x-window-dlg").button(:text=>"Yes").click
                  sleep 5  
            when "case_2_size"  
                  $size_PO=true             
                  theBrowser.input(:id=>"colorSizeChoiceS").click
                  theBrowser.div(:class=>"x-window x-window-dlg").button(:text=>"Yes").click    
                  sleep 2
                  delete_sizes
            when "case_3_size"  
                  $size_PO=true             
                  theBrowser.input(:id=>"colorSizeChoiceS").click
                  theBrowser.div(:class=>"x-window x-window-dlg").button(:text=>"Yes").click    
                  sleep 2
                  change_casePK 

            when "case_4_size"  
                  $size_PO=true             
                  theBrowser.input(:id=>"colorSizeChoiceS").click
                  theBrowser.div(:class=>"x-window x-window-dlg").button(:text=>"Yes").click    
                  sleep 2
                  deleteColor_Addagain                                 
          end


  ########Fill Allocation percentages
    if $allocation_QLD!="NA"
      if $size_PO==true
       theElement("QLDpercentageSize","POrder").send_keys $allocation_QLD        
      else
       theElement("QLDpercentage","POrder").send_keys $allocation_QLD
      end
    end

    if $allocation_VIC!="NA"
      if $size_PO==true
       theElement("VICpercentageSize","POrder").send_keys $allocation_VIC
      else
       theElement("VICpercentage","POrder").send_keys $allocation_VIC           
      end      
    end

    if $allocation_WA!="NA"
      if $size_PO==true
       theElement("WApercentageSize","POrder").send_keys $allocation_WA
      else
       theElement("WApercentage","POrder").send_keys $allocation_WA         
      end      
    end

    if $allocation_NSW!="NA"
      if $size_PO==true
       theElement("NSWpercentageSize","POrder").send_keys $allocation_NSW
      else
       theElement("NSWpercentage","POrder").send_keys $allocation_NSW        
      end       
    end

    if $allocation_Total!="NA"
      
        if $size_PO==true
         theBrowser.input(:id=>"national_percentage_size$0").click
         theBrowser.send_keys [:control, 'a'], :backspace 
         theElement("NationalpercentageSize","POrder").send_keys $allocation_Total
        else
         theBrowser.input(:id=>"national_percentage_color$0").click
         theBrowser.send_keys [:control, 'a'], :backspace 
         theElement("Nationalpercentage","POrder").send_keys $allocation_Total        
        end      

    else 
      if $size_PO==true
         theBrowser.input(:id=>"national_percentage_size$0").click
         theBrowser.send_keys [:control, 'a'], :backspace 
      else
       theBrowser.input(:id=>"national_percentage_color$0").click
       theBrowser.send_keys [:control, 'a'], :backspace 
      end 
    end


  ###############################   

  ######## Allocate  DC Number

    if $dcNo_QLD!="NA"
       theElement("QLDcode","POrder").send_keys $dcNo_QLD
       theBrowser.div(:class=>"x-combo-list-item",:text=>"#{$dcNo_QLD}").click
    end

    if $dcNo_VIC!="NA"
       theElement("VICcode","POrder").send_keys $dcNo_VIC
       theBrowser.div(:class=>"x-combo-list-item",:text=>"#{$dcNo_VIC}").click  
    end

    if $dcNo_WA!="NA"
       theElement("WAcode","POrder").send_keys $dcNo_WA
       theBrowser.div(:class=>"x-combo-list-item",:text=>"#{$dcNo_WA}").click   
    end

    if $dcNo_NSW!="NA"
       theElement("NSWcode","POrder").send_keys $dcNo_NSW
       theBrowser.div(:class=>"x-combo-list-item",:text=>"#{$dcNo_NSW}").click
    end

    if $dcNo_Total!="NA"
       if $size_PO==true
         theBrowser.input(:id=>"nationalDcNoId$0").click           
       end 
       theBrowser.input(:id=>"nationalDcNoId$0").click  
       sleep 1
       theBrowser.send_keys [:control, 'a'], :backspace
       theBrowser.input(:id=>"nationalDcNoId$0").send_keys $dcNo_Total
       sleep 1
       theBrowser.div(:class=>"x-combo-list-item",:text=>"#{$dcNo_Total}").click  
    end   

  ################################

  if $size_PO==true
      quantities=theBrowser.execute_script("return document.getElementById('shipmentSizes').getElementsByClassName('x-grid3-cell-inner x-grid3-col-minQty')")
      quantity=theBrowser.execute_script("return document.getElementsByClassName('x-grid3-cell-inner x-grid3-col-sizeCodeQty')")
        ########Fill Colours Quantity
           q=0
              quantity.each do |qty|
                qty.click
                temp_text=Integer(quantities[q].text)
                qty_value=temp_text*20
                q+=1
                sleep 1
                theBrowser.send_keys qty_value
              end
        #############################    
        theBrowser.button(:id=>"refreshSizeQtyByPercentage").click
        sleep 7   
   else
     quantities=theBrowser.execute_script("return document.getElementsByClassName('x-grid3-cell-inner x-grid3-col-minQty')")
     quantity=theBrowser.execute_script("return document.getElementsByClassName('x-grid3-cell-inner x-grid3-col-sellColorQty')")
        ########Fill Colours Quantity
           q=0
              quantity.each do |qty|
                qty.click
                temp_text=Integer(quantities[q].text)
                qty_value=temp_text*20
                q+=1
                sleep 1
                theBrowser.send_keys qty_value
              end
        #############################      

     theBrowser.button(:id=>"refreshColorQtyByPercentage").click  
      sleep 7   
  end
    

     theElement("ShipmentSave","POrder").click
     sleep 10
     theElement("ShipWinClose","POrder").click
     sleep 1
     theElement("Save","Portal").click
     sleep 10
     theElement("Register","Portal").click   
     sleep 8
     
  end
##*********************************************************************** end of case-1 **************************************************************************************
def delete_colours_allocation
viv=theBrowser.div(:id=>"shipmentSizes").elements(:class=>"x-grid3-row-checker")

   tick_shipments_toDelete=theBrowser.div(:id=>"shipmentColors").elements(:class=>"x-grid3-row-checker")
     for ds in 1..tick_shipments_toDelete.length-1
         tick_shipments_toDelete[ds].click
     end
   theBrowser.div(:id=>"shipmentColorsGridButtons").button(:id=>"delete").click
     theElement("ShipmentSave","POrder").click
     sleep 8
 
end

def change_casePK

  if $size_PO== true
     casePK=theBrowser.execute_script("return document.getElementsByClassName('x-grid3-col x-grid3-cell x-grid3-td-keycodeCasePackId  x-grid3-dirty-cell')")
  else
     casePK=theBrowser.execute_script("return document.getElementsByClassName('x-grid3-col x-grid3-cell x-grid3-td-casePackId')")
  end
     default=casePK[0].text
     casePK[0].click
     sleep 1
     theBrowser.send_keys :down
     sleep 1     

    for cl in 0..theBrowser.elements(:class=>"x-combo-list-item").length-1
      if theBrowser.elements(:class=>"x-combo-list-item")[cl].present?
           if theBrowser.elements(:class=>"x-combo-list-item")[cl].text==default
             if theBrowser.elements(:class=>"x-combo-list-item")[cl+1].present?
                theBrowser.elements(:class=>"x-combo-list-item")[cl+1].click
             end
           else
              theBrowser.elements(:class=>"x-combo-list-item")[cl].click
           end
       end
     end
    
end

def deleteColor_Addagain

  ##### delete one the last one
     tick_shipments_toDelete=theBrowser.div(:id=>"shipmentColors").elements(:class=>"x-grid3-row-checker")
     tick_shipments_toDelete[tick_shipments_toDelete.length-1].click
     theBrowser.div(:id=>"shipmentColorsGridButtons").button(:id=>"delete").click
       theElement("ShipmentSave","POrder").click
       sleep 8

  crs_array=Array.new
  color_row_ship=theBrowser.div(:id=>"shipmentColors").elements(:class=>"x-grid3-row-table")
   color_row_ship.each do |crs|
     crs_array << crs.div(:class=>"x-grid3-cell-inner x-grid3-col-sellColor").text
  end

  # ##click select button
   theBrowser.div(:id=>"shipmentColorsGridButtons").button(:id=>"addItems").click   
   sleep 4
   theBrowser.input(:id=>"btn-orderShipmentColorsPopupedshipmentColorsSearch").click
   sleep 3

  color_row=theBrowser.div(:id=>"orderShipmentColorsPopuped").elements(:class=>"x-grid3-row")
   color_row.each do |cr|
     already_present=false 
      crs_array.each do |crsa|
        if crsa==cr.div(:class=>"x-grid3-cell-inner x-grid3-col-sellColor").text
           already_present=true
        end
      end
     
     if already_present==false
        cr.div(:class=>"x-grid3-cell-inner x-grid3-col-sellColor").double_click
        break
     end
  end

end

def delete_sizes
   tick_sizes_toDelete=theBrowser.div(:id=>"shipmentSizes").elements(:class=>"x-grid3-row-checker")
     for ds in 1..tick_sizes_toDelete.length-1
         tick_sizes_toDelete[ds].click
     end
   theBrowser.div(:id=>"sizeGroup").button(:id=>"delete").click
   sleep 2
     theElement("ShipmentSave","POrder").click
     sleep 8
 
end

##  FOR multiple briefs in a single PO
# theBrowser.div(:id=>"orderShipmentItemsGridButtons").button(:id=>"addItems").click