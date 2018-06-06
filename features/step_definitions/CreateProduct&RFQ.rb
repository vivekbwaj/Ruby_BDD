require 'chronic'
require 'time'
require 'colorize'
require 'date'
require 'win32ole'

$fillerText=""
$department=""
$colors=""
$loggedIn = false
$refId=""
$startRow_excel=2
$rows=""

Given(/^needs "([^"]*)" "([^"]*)" product of "([^"]*)" of color "([^"]*)" where "([^"]*)" template selection "([^"]*)"$/) do |sup_loc,brief,arg1, arg2, arg3,template| 

  $prodType=brief
  $department="#{arg1}"
  $colors=arg2.split(",")
  $case_type_prod=arg3
  $templateName=template
  $sup_location=sup_loc

  if $sup_location=="Indent"
    $venNameSAPNum="4105213"
    $factory="" # not being used currently
    $sA="2001 - TGT Sourcing Asia"
    $cOrigin="China"
    $aO="Shanghai"
    # $portOfLoading="Shenzhen"
    $portOfLoading="Shanghai"
    $currency="USD"
  else
    $venNameSAPNum="3866"
    $factory="" # not being used currently
    $sA="4000 - No Agent"
    $cOrigin="Australia"
    $aO="Australia"
    $portOfLoading="Melbourne"

    # $portOfLoading="Shenzhen"
    $currency="AUD"
  end

  createNow

end
	
def createNow

		################# loading product brief page ##########
		##  NA means don' search for any particular  Refid 
    
    if $prodType=="GMProducts"
		  loadGMproductBrief("NA")    	
    else
      loadAPPproductBrief("NA")        
    end
 		####################### New and Fill Header page details #########################################
			 theElement("New","Portal").flash	
			 theElement("New","Portal").click
			 style_num=DateTime.now
			 waitForObjectToDisappear("Copy","Portal",10)
     		 $fillerText=style_num.strftime "%H%M%d%m%y"
             # puts "#{$fillerText}"
             selectDropdownValue("Department",$department,"Header")
                    selectDropdownValue("SourcingAgent",$sA,"Header")
        selectDropdownValue("AgentOffice",$aO,"Header")  
			  # selectJSdropdown("Department",$department,"Header")
        onElement("Description","set",$fillerText,"Header")
        onElement("EstFOB","set","5","Header") 
        onElement("EstELC","set","7","Header") 
        onElement("EstSellPrice","set","29.95","Header") 
        onElement("InitialOrder","set","100","Header") 
        onElement("MonthlyOrder","set","150","Header")         
			  # selectJSdropdown("SourcingAgent",$sA,"Header")
        # selectJSdropdown("AgentOffice",$aO,"Header")	
        # selectJSdropdown("TechDesigner","Anubha Bhatnagar","Header")
        # selectJSdropdown("DesignStylist","Designer Designer","Header")
        # selectJSdropdown("Buyer","BUYER BUYER","Header")
        # selectJSdropdown("MerchAssistant","Brett Rees","Header")
        # selectJSdropdown("Planner","Erik Chan","Header")
        # selectJSdropdown("Merchandiser","Ian Rabie","Header")
        # selectJSdropdown("ProdMerchandiser","Chirag Pandya","Header")
        # selectJSdropdown("QATester1","Mark Richmund","Header")
        # selectJSdropdown("Track","Normal","Header")
			  theElement("Range","Header").click
			  sleep 2

			  ## to select rangeName
			  theBrowser.input(:id=>"prffsef_mc.ref_4").click
			  sleep 2
			  theBrowser.send_keys :tab
			  theBrowser.send_keys :enter
			  if isElementPresent("RangeFirstResult","Header",10) then
			  	 theElement("RangeFirstResult","Header").double_click
			  	else
			  		raise "Range Results table didn't load"
			  end

			  ## to select vendor
        theElement("vendorName","Header").click
        sleep 2
        onElement("vendorSAPNum","set",$venNameSAPNum,"Header")
        theBrowser.send_keys :enter
        if isElementPresent("vendorFirstResult","Header",10) then
           theElement("vendorFirstResult","Header").double_click
          else
            raise "Vendor Results table didn't load"
        end            
			  
        #to select factory
        theElement("factoryName","Header").click
        sleep 2
        theBrowser.input(:id=>"mf.factoryFullNameEn").click
        theBrowser.send_keys :enter
        if isElementPresent("factoryFirstResult","Header",10) then
           theElement("factoryFirstResult","Header").double_click
          else
            raise "Factory Results table didn't load"
        end
    
         dc_date=DateTime.now.next_month.strftime "%d/%m/%Y"
     show_date=DateTime.now.next_month.next_month.strftime "%d/%m/%Y"

  # when on Show date is entered rest of the dates are calculated on it's own
			 			 			 theBrowser.input(:id=>"inStoreDate$0").send_keys DateTime.now.next_month.next_month.strftime "%d/%m/%Y" 
			  # hit save

			  theElement("Save","Portal").click
			  sleep 8
        $refId=theElement("RefNo","Portal").text
        puts theElement("RefNo","Portal").text
        # onElement("StyleNo","set",$refId,"Header")
       
     ####################### Fill Specification page details #########################################                  

          if isElementPresent("SpecificationTab","Specification",5) then
             theElement("SpecificationTab","Specification").click
              if isElementPresent("ItemDesc","Specification",3) then
                onElement("ItemDesc","set",$fillerText,"Specification")
 
				  # hit save
			       theElement("Save","Portal").click
			       sleep 8
              end
          else
             raise "Specification tab not found"   
          end

          ################################################################
           if $prodType=="APP" 
            if isElementPresent("MeasurementTab","Measurement",5)
                  theElement("MeasurementTab","Measurement").click
                  sleep 2
                  theBrowser.input(:id=>"templateTypeId$0").click
                  theBrowser.send_keys :down
                  theBrowser.send_keys :enter
                  # theBrowser.input(:id=>"templateNamePopupBtn").click
                  # sleep 2
                  # theBrowser.div(:id=>"styleMeasPopupPopupedDIV").div(:id=>"btn-styleMeasPopupPopupedmeasSearch").click
                 theElement("Save","Portal").click
                 sleep 8
            else
                 raise "Measurement tab not found"   
            end
          end  
          ###############################################################
     ####################### Fill costing page details ######################################### 

          if isElementPresent("CostingTab","Costing",5) then
             theElement("CostingTab","Costing").click

                if $prodType=="APPProducts" 
                   fillCostingValues("set1")

                else
                   fillCostingValues("set2")
                end
				  # hit save
				  theElement("Save","Portal").click
				  sleep 8
          else
             raise "Costing tab not found"   
          end
     ####################### Fill color size page details ######################################### 
 
          if isElementPresent("ColorSizeTab","ColorSize",5) then
               theElement("ColorSizeTab","ColorSize").click    
               sleep 2

               		###### more than one color
               			numOfColors=$colors.length
               			for c in 2..numOfColors
			    			       buttons = theBrowser.elements(:class=>"x-btn-text")
    						       buttons.each do |add|
     							      if add.text=="Add" then
     							      	add.click
     								      sleep 2
     							      end
     						       end
               			end

               			contrastColorEnabled=theBrowser.elements(:class=>"x-grid3-cell-inner x-grid3-col-contractColourId",:tag_name=>"div")
               			nC=contrastColorEnabled.size
               			colorIndex=0
               			contrastColorEnabled.each do |clrs|
               				clrs.click
               				sleep 2
                            theBrowser.send_keys [:control, 'a'], :backspace
                            colorSelect=$colors[colorIndex]
		       				       theBrowser.div(:class=>"x-combo-list-item",:text=>"#{colorSelect}").click 
		       				       sleep 2                            
                            colorIndex=colorIndex+1
                              if colorIndex==nC-1 then
                              	break
                              end
               			end

               			sellColorEnabled=theBrowser.execute_script("return document.getElementsByClassName('x-grid3-cell-inner x-grid3-col-sellColor')")
               			nsC=theBrowser.execute_script("return document.getElementsByClassName('x-grid3-cell-inner x-grid3-col-sellColor').length")
               			for sC in 0..nsC-1
               			  colorSelect=$colors[sC]
               			  theBrowser.execute_script("document.getElementsByClassName('x-grid3-cell-inner x-grid3-col-sellColor')[#{sC}].click()")
               			  sleep 1
                      theBrowser.send_keys [:control, 'a'], :backspace
               			  theBrowser.send_keys "#{colorSelect}"
               			end
              			
               		
                  ######
                  theElement("AutoAssignAPN","ColorSize").click 

##############################fill casepack details

                  case $case_type_prod

                    when "innerEqualOuter"
                           if $templateName!=""
                             theElement("Template","ColorSize").click
                             sleep 3
                               if theBrowser.div(:id=>"styleSizesPopuped").div(:text=>"#{$templateName}").present?
                                 theBrowser.div(:id=>"styleSizesPopuped").div(:text=>"#{$templateName}").double_click
                                 sleep 2
                                 if theBrowser.span(:class=>"x-window-header-text",:text=>"Please confirm").present?
                                    theBrowser.div(:class=>"x-window x-window-dlg").button(:text=>"Yes").click
                                    sleep 2
                                    fillCasepacks
                                    enterMinQty 1                            
                                 end
                               else
                                 raise "#{templateName} template doesn't exist"
                               end
                           end
                    when "OuterGreaterInner"
                           if $templateName!=""
                             theElement("Template","ColorSize").click
                             sleep 3
                               if theBrowser.div(:id=>"styleSizesPopuped").div(:text=>"#{$templateName}").present?
                                 theBrowser.div(:id=>"styleSizesPopuped").div(:text=>"#{$templateName}").double_click
                                 sleep 2
                                 if theBrowser.span(:class=>"x-window-header-text",:text=>"Please confirm").present?
                                    theBrowser.div(:class=>"x-window x-window-dlg").button(:text=>"Yes").click
                                    sleep 2
                                   fillCasepacks
                                   enterMinQty 2   
                                   sleep 1                                 
                                 end
                               else
                                 raise "#{templateName} template doesn't exist"
                               end
                           else
                                   fillCasepacks
                                   enterMinQty 2                              
                           end

                    when "makeSizeInactive"
                          $makeInactive=true
                          # here template selection is always there, we need to select template and make some sizes inactive
                             theElement("Template","ColorSize").click
                             sleep 3
                               if theBrowser.div(:id=>"styleSizesPopuped").div(:text=>"#{$templateName}").present?
                                 theBrowser.div(:id=>"styleSizesPopuped").div(:text=>"#{$templateName}").double_click
                                 sleep 2
                                 if theBrowser.span(:class=>"x-window-header-text",:text=>"Please confirm").present?
                                    theBrowser.div(:class=>"x-window x-window-dlg").button(:text=>"Yes").click
                                    sleep 2
                                   fillCasepacks
                                   enterMinQty 2   
                                   sleep 1                                 
                                 end
                               else
                                 raise "#{templateName} template doesn't exist"
                               end
                    else
                    end

##################################################################################################

              theElement("Save","Portal").click
    			   sleep 10
          else
             raise "Color-size tab not found"   
          end

     ####################### Fill Keycode page details #########################################

            theElement("KeyCodeTab","Keycode").click
            sleep 5
            onElement("ProdName","set","Vivek_product","Keycode")
            onElement("ArthurSDesc","set","Vivek_product","Keycode")
            onElement("WebProdName","set","Vivek_product","Keycode")
            onElement("ProdFeatures","set","Vivek_product","Keycode")            
            onElement("POSDesc","set","Sample","Keycode")
            onElement("SPLLine1","set","Vivek_product","Keycode")
            onElement("MinPresDepth","set","5","Keycode")
            selectDropdownValue("Barcode","TABAL","Keycode") 

			    buttons = theBrowser.elements(:class=>"x-btn-text")
    				buttons.each do |btn|
     					if btn.text=="Add..." then
     						btn.click
     						sleep 3
							    buttons = theBrowser.elements(:class=>"x-btn-text")
				    				buttons.each do |btn|
				     					if btn.text=="Search" then
				     						btn.click
				     						break
				     					end		
				    				end     						
     						break
     					end		
    				end
			 sleep 5
			 theBrowser.div(:text=>"10480").double_click
			# theElement("AssortmentFirstResult","Keycode").double_click 
			sleep 1
			theElement("Save","Portal").click
			  sleep 8
			theElement("Applydefault","Keycode").click
			sleep 1
		    theElement("YesButton","Portal").flash
		    theElement("YesButton","Portal").click
		    sleep 4

      for c in 0..$colors.length-1
          theBrowser.input(:id=>"colourKeycodeId$0").click
          theBrowser.send_keys :down  
          keycode_options=theBrowser.execute_script("return document.getElementsByClassName('x-layer x-combo-list dropdownList dropdownList-colourKeycodeId$0')[0].getElementsByTagName('div')[0].getElementsByTagName('div')")      
 
            keycode_options.each do |keycol|
              # binding.pry
              if keycol.text.include? $colors[c] then
                 keycol.click
                 sleep 2
                numoffields=theBrowser.execute_script("return document.getElementsByClassName('x-grid3-cell-inner x-grid3-col-subTypeCode').length")
                # binding.pry
                  for row in 0..numoffields-1
                     check_text=theBrowser.execute_script("return document.getElementsByClassName('x-grid3-cell-inner x-grid3-col-subTypeCode')[#{row}].innerHTML")
                     if check_text=="Primary Colour"
                        indexIs=row
                        break
                    end
                  end
                theBrowser.execute_script("document.getElementsByClassName('x-grid3-cell-inner x-grid3-col-selectCode')[#{indexIs}].getElementsByTagName('input')[0].click()")
                sleep 2
                theBrowser.input(:id=>"mcl.name").send_keys $colors[c]
                theBrowser.send_keys :enter
                theBrowser.div(:id=>"productAttributeCodePopupedGrid").div(:class=>"x-grid3-cell-inner x-grid3-col-name").double_click
                theElement("Save","Portal").click
                sleep 8
                break
            end
           end 

      end

 
		  theElement("Register","Portal").click
			sleep 8
			theElement("HeaderTab","Header").click
			sleep 3
			regDone="No"
			reg=theBrowser.elements(:class=>"header_value")
			  reg.each do |regStatus|
			  	if regStatus.text.include? "LOCKED" then
			  		puts "LOCKED".yellow
			  		regStatus.flash
					regDone="Yes"			  		
			  		break
			  	end	
			  end

			  if regDone=="Yes" then
			  		theElement("RefNo","Portal").flash
			        emailtext="ReferenceId: #{$refId}"
              subj = "#{$sup_location} #{$prodType} for department #{$department} #{$colors} where #{$case_type_prod} using size template #{$templateName}"              
			        email(emailtext,subj,"")

			  else
			  	raise "Product registration not done:Some fields might be incorrect/missing"
			  end
end			


Then(/^the user gets the result in "([^"]*)" at "([^"]*)" and entity "([^"]*)"$/) do |file,po_Num_OR_flowType,poll_entity|

    sheetName="Order"
  
        excel = WIN32OLE.new('Excel.Application')
        excel.visible = true #true means that the excel will be visible , false= excel runs in background
        workbook = excel.workbooks.open("#{file}")
        worksheet = workbook.worksheets("#{sheetName}")
        worksheet.Activate
      if po_Num_OR_flowType=="CreateProductAndPoll"
         worksheet.Cells($startRow_excel,2).Value=$refId
         worksheet.Cells($startRow_excel,3).Value=poll_entity  
         $startRow_excel+=1    

      else
        $rows=po_Num_OR_flowType.split(",")
        if $rows.length >0
          $rows.each do |r|
            r=Integer(r)
            # puts "row number is #{r}"
              worksheet.Cells(r+1,4).Value=$refId
              worksheet.Cells(r+1,5).Value=poll_entity
          end   
         else
          puts "PO not required for this brief"
         end 
      end  
         # worksheet.Range("A#{$startRow_excel}").value=$refId
         # worksheet.Range("B#{$startRow_excel}").value=$fillerText
         # $startRow_excel+=1

        workbook.saved = true #Setting it to True is a simple way of preventing the "Do you wish to save..." dialog appearing when Excel is closed
        workbook.Save
        excel.ActiveWorkbook.Close(0)
        excel.Quit()
end

def fillCasepacks

  num_of_sizes=theBrowser.execute_script("return document.getElementsByClassName('x-grid3-cell-inner x-grid3-col-pack2A').length")

    if $makeInactive==true
      upperLimit=num_of_sizes-2
      theBrowser.execute_script("document.getElementsByClassName('x-grid3-cell-inner x-grid3-col-pack2A')[#{num_of_sizes-1}].click()")
      sleep 1
      theBrowser.send_keys [:control, 'a'], :backspace                        
      theBrowser.send_keys "0"      
      theBrowser.div(:id=>"sizesGrid").elements(:class=>"x-grid3-row-checker")[num_of_sizes-1].click
      theBrowser.button(:id=>"inactive").click
      sleep 2
      $makeInactive=false
    else
      upperLimit=num_of_sizes-1    
    end

   for i in 0..upperLimit
      theBrowser.execute_script("document.getElementsByClassName('x-grid3-cell-inner x-grid3-col-pack2A')[#{i}].click()")
      sleep 1
      theBrowser.send_keys [:control, 'a'], :backspace                        
      theBrowser.send_keys "1"

      theBrowser.execute_script("document.getElementsByClassName('x-grid3-cell-inner x-grid3-col-pack2B')[#{i}].click()")
      sleep 1
      theBrowser.send_keys [:control, 'a'], :backspace  
      theBrowser.send_keys "2"

      theBrowser.execute_script("document.getElementsByClassName('x-grid3-cell-inner x-grid3-col-pack2C')[#{i}].click()")
      sleep 1
      theBrowser.send_keys [:control, 'a'], :backspace  
      theBrowser.send_keys "3"

      theBrowser.execute_script("return document.getElementById('sizesGrid').getElementsByClassName('x-grid3-col x-grid3-cell x-grid3-td-variationaPack')[#{i}].click()")
      sleep 1
      theBrowser.send_keys [:control, 'a'], :backspace  
      theBrowser.send_keys "5"

      theBrowser.execute_script("return document.getElementById('sizesGrid').getElementsByClassName('x-grid3-col x-grid3-cell x-grid3-td-variationaMinQty')[#{i}].click()")
      sleep 1
      theBrowser.send_keys [:control, 'a'], :backspace  
      theBrowser.send_keys "5"

      theBrowser.execute_script("return document.getElementById('sizesGrid').getElementsByClassName('x-grid3-col x-grid3-cell x-grid3-td-variationbPack')[#{i}].click()")
      sleep 1
      theBrowser.send_keys [:control, 'a'], :backspace  
      theBrowser.send_keys "10"

      theBrowser.execute_script("return document.getElementById('sizesGrid').getElementsByClassName('x-grid3-col x-grid3-cell x-grid3-td-variationbMinQty')[#{i}].click()")
      sleep 1
      theBrowser.send_keys [:control, 'a'], :backspace  
      theBrowser.send_keys "10"
    end    

   theElement("Save","Portal").click
   sleep 10

 end

def enterMinQty(outerInnerRatio)                         

   packATotal=theBrowser.div(:class=>"x-grid3-cell-inner x-grid3-col-packA").text
   aMin=Integer(packATotal)*outerInnerRatio
   packBTotal=theBrowser.div(:class=>"x-grid3-cell-inner x-grid3-col-packB").text
   bMin=Integer(packBTotal)*outerInnerRatio
   packCTotal=theBrowser.div(:class=>"x-grid3-cell-inner x-grid3-col-packC").text
   cMin=Integer(packCTotal)*outerInnerRatio

    ############Fill min. quantity for pack A,B,C
    minQty=theBrowser.execute_script("return document.getElementsByClassName('x-grid3-cell-inner x-grid3-col-packaMin').length")
      for minA in 0..minQty-1
        theBrowser.execute_script("document.getElementsByClassName('x-grid3-cell-inner x-grid3-col-packaMin')[#{minA}].click()")
        sleep 1
        theBrowser.send_keys [:control, 'a'], :backspace
         theBrowser.send_keys aMin
      end                    

    minQty=theBrowser.execute_script("return document.getElementsByClassName('x-grid3-cell-inner x-grid3-col-packbMin').length")
      for minB in 0..minQty-1
        theBrowser.execute_script("document.getElementsByClassName('x-grid3-cell-inner x-grid3-col-packbMin')[#{minB}].click()")
        sleep 1
        theBrowser.send_keys [:control, 'a'], :backspace
         theBrowser.send_keys bMin
      end  

    minQty=theBrowser.execute_script("return document.getElementsByClassName('x-grid3-cell-inner x-grid3-col-packcMin').length")
      for minC in 0..minQty-1
        theBrowser.execute_script("document.getElementsByClassName('x-grid3-cell-inner x-grid3-col-packcMin')[#{minC}].click()")
        sleep 1
        theBrowser.send_keys [:control, 'a'], :backspace
        theBrowser.send_keys cMin
      end  
end

