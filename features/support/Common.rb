require 'chronic'
require 'time'
require 'colorize'
require 'date'
require 'win32ole'

# $loggedIn = false
$found=""
$ids_dict=WIN32OLE.new('Scripting.Dictionary')
$poNum_dict=WIN32OLE.new('Scripting.Dictionary')
#$ids_dict.Items() will give you the content
$shipmentFile = ""
$secondsStart="NA"

Given(/^the user is logged in$/) do

   if $loggedIn==false then
		  if isElementPresent("Login","Portal",10) then
			  onElement("Username", "set", ENV['uname'], "Portal")
		      onElement("Password", "set", ENV['pwd'], "Portal")
		      theElement("Login", "Portal").flash
		      theElement("Login", "Portal").click
		      $loggedIn = true

		      if isElementPresent("Logout","Portal",10) then
		        puts "Logged in successfully"
		      else
		        raise "Clicked on login button but didn't land on expected page"
		      end
		  else
		    raise "Couldn't log in"
		  end
   else		  
			theBrowser.span(:text=>"Home").click
			sleep 10
	end		

end

Then(/^user logs out$/) do
	if isElementPresent("Logout","Portal",10) then
		theElement("Logout","Portal").click
		$loggedIn=false
	end
end 

Given(/^the user clears the ProdIdDictionary$/) do
	$ids_dict.RemoveAll()
end


def email(text,subj,attachment_file)
# puts "#{text} : #{subj} "
# puts ""
	outlook = WIN32OLE.new('Outlook.Application')
	message = outlook.CreateItem(0)
    message.Subject = "#{subj}"
	message.Body = "#{text}"
	message.Recipients.Add 'TEST.TRADING@target.com.au'
	# message.Recipients.Add 'vivek.bhardwaj@target.com.au'
	# message.Recipients.Add 'mohammad.khan@target.com.au'	
    #message.Recipients.Add 'Chirag.Pandya@target.com.au'
    #message.Recipients.Add 'Anubha.Bhatnagar@target.com.au'
	# message.Recipients.Add 'David.Manison@target.com.au'
	#message.Recipients.Add 'Darren.Hauenstein@target.com.au'
	# message.Recipients.Add 'Simon.Shand@target.com.au'
	# message.Recipients.Add 'Mario.Miksic@target.com.au'
	# message.Recipients.Add 'Steven.Kerovic@target.com.au'
	# message.Recipients.Add 'Sajith.Purushothaman@target.com.au'

	if attachment_file!=""
	 message.Attachments.Add("#{attachment_file}")
	end
	#Want to save as a draft?
	message.Save
	#Want to send instead?
	message.Send

end

Given(/^fetches all "([^"]*)" ids from "([^"]*)"$/) do |entity,file|

  $excel = WIN32OLE.new('Excel.Application')
  $excel.visible = true
  $workbook = $excel.Workbooks.open("#{file}")
  $worksheet = $workbook.Worksheets('Order')
  $worksheet.Activate

case entity

	when "PO"
  	  $column_number=23
	   for r in 2..$worksheet.UsedRange.Rows.Count
	       if $worksheet.Cells(r,1).Value=="Done"
			  $ids_dict[r]=$worksheet.Cells(r,$column_number).Value 
			  $poNum_dict[r]=$worksheet.Cells(r,2).Value 
	        end 
	   end

	when "ProductsForPO"
 	  $column_number=4
	   for r in 2..$worksheet.UsedRange.Rows.Count
	       if $worksheet.Cells(r,5).Value=="APPProducts" || $worksheet.Cells(r,5).Value=="GMProducts"
			  $ids_dict[r]=$worksheet.Cells(r,$column_number).Value 
			  $poNum_dict[r]=$worksheet.Cells(r,2).Value 
	        end 
	   end 	  

	when "Products"
	   for r in 2..$worksheet.UsedRange.Rows.Count
			  $ids_dict[r]=$worksheet.Cells(r,2).Value 
			  $poNum_dict[r]=$worksheet.Cells(r,1).Value       
	   end 

 	else
 		raise "Unidentifiable entity for registration validation"
	end


 # $workbook.saved = true #Setting it to True is a simple way of preventing the "Do you wish to save..." dialog appearing when Excel is closed
 # $workbook.Save
 # $excel.ActiveWorkbook.Close(0)
 # $excel.Quit()
end

   def selectDropdownValue(elementName,value,pageKey)
		theElement(elementName,pageKey).click
		theBrowser.send_keys [:control, 'a'], :backspace
		theBrowser.div(:class=>"x-combo-list-item",:text=>"#{value}").click
    end

   def selectJSdropdown(elementName,value,pageKey)
        locator=theElement(elementName,pageKey).id
		theElement(elementName,pageKey).click
		theBrowser.send_keys [:control, 'a'], :backspace
		class_variable="x-layer x-combo-list dropdownList dropdownList-"+locator
		options=theBrowser.execute_script("return document.getElementsByClassName('#{class_variable}')[0].getElementsByTagName('div')[0].getElementsByTagName('div')")
		options.each do |opt|
			if opt.text==value
				opt.click
				sleep 1
				break
			end
		end
    end



def poll(id,poll_entity)

case poll_entity

	when "PO"
		identifier_number="v_orderNo"
		identifier_status="dropDown_orderRegStatusId"
		result_Type	="PORegnStatus"

	when "APPProducts"
		identifier_number="v_refNo"
		identifier_status="dropDown_productRegStatusId"
		result_Type	="ProdRegnStatus"
	when "GMProducts"
		identifier_number="v_refNo"
		identifier_status="dropDown_productRegStatusId"
		result_Type	="ProdRegnStatus"		
	end

		if isElementPresent(poll_entity,"Portal",10) then
			theElement(poll_entity,"Portal").flash
			theElement(poll_entity,"Portal").click
		       if isElementPresent("AdvSearch","Portal",10) then
		       	  	if poll_entity=="APPProducts" || poll_entity=="GMProducts"
		       	  		if id.class==Float
		       	  			id=Integer(id)
		       	  		end
		       	  	end
		       	  	theBrowser.input(:id=>"#{identifier_number}").send_keys "#{id}"		       	  	
                    theBrowser.input(:id=>"#{identifier_status}").send_keys "REGISTERED"
                    sleep 2
                    theBrowser.div(:class=>"x-combo-list-item",:text=>"REGISTERED").click
		       	    theElement("AdvSearch","Portal").flash
		       	    theElement("AdvSearch","Portal").click
		       	     if isElementPresent(result_Type,"Portal",2) then
		       	     	theElement(result_Type,"Portal").flash  
		       	        $found=true		
		       	        if poll_entity=="PO"
		       	        	$shipment_num=theElement("ShipmentNum","Portal").text
		       	        	# theElement("ShipmentNum","Portal").click
		       	        	# sleep 4
		       	        	# theElement("TransmissionHistory","POrder").click
		       	        	# sleep 3
		       	        	# if theElement("ODRAck","POrder").present?
		       	        	#   if theElement("ODRAck","POrder").text=="Acknowledged"
		       	        	# 	theElement("ODRAck","POrder").flash
		       	        	#   else
		       	        	#   	$found=false
		       	        	#   end      	        	 
		       	        	# else
		       	        	# 	$found=false
		       	        	# end
		       	        end
		       	        theElement("HomeTab","Portal").click
		     	    else	
		       	     	$found=false
		       	     	theElement("HomeTab","Portal").click
		       	    end	
		        else
		        	raise "Advanced Search button not found"
		        end
		else
		    raise "#{poll_entity} link not found"
		end 	
 return $found		
end

Then(/^validates "([^"]*)" registration to "([^"]*)"$/) do |arg1,attachment|
    		
    		if arg1=="Products" || arg1=="ProductsForPO"
    		 
    		else
    	      $shipmentFile=File.new("#{attachment}", "r+")
    		end

	$ids_dict.each.with_index do |val,index|
		# binding.pry
    		time2 = Time.now
    		if arg1=="Products"
    		  poll_entity=$worksheet.Cells(val,3).Value
    		else
    	      poll_entity=$worksheet.Cells(val,5).Value	
    		end

			while poll($ids_dict[val],poll_entity)==false  do
				sleep 60
				# puts "time is #{Time.now}"
				if Time.now > time2 + 1200
			    	puts "#{$ids_dict[val]} didn't get registered in 20 minutes"

			    	case arg1

			    	when "PO"
			    	$worksheet.Cells(val,$column_number+1).Value="Failed registration"
			    	break				    		
			    	
			    	when "ProductsForPO"
			    	$worksheet.Cells(val,28).Value="Failed registration"
			    	break				   
			    	
			    	when "Products"
			    	$worksheet.Cells(val,4).Value="Failed registration"
			    	break	

			    	end

			    end		
			end     

		    if $found==true
		    	case arg1

		    	when "PO" 
			    	begin
			    	  	$shipmentFile.syswrite("#{val-1}:#{$shipment_num} ")
			    	  	$shipmentFile.syswrite("\n")
			    	  	email($ids_dict[val],"PO registered with Shipment id: #{$shipment_num}","")
	         	    rescue Exception => e
	         	    	body=e.message + e.backtrace.inspect[0..300]  
	         	    	subj="#{$shipment_num} not written to textfile"
			    	end
    	    	  $worksheet.Cells(val,$column_number+1).Value=$shipment_num

		    	when "ProductsForPO"

    	    	  $worksheet.Cells(val,$column_number+1).Value="PO"
    	    	  $worksheet.Cells(val,28).Value="Product registered"

		    	when "Products"

    	    	  $worksheet.Cells(val,4).Value="Product registered"

		    	else
		    	  puts "Incorrect argument in validates registration step" 
		    	end
		    	
  
		    end	  	    
   end     
	 $workbook.saved = true #Setting it to True is a simple way of preventing the "Do you wish to save..." dialog appearing when Excel is closed
	 $workbook.Save
	 $excel.ActiveWorkbook.Close(0)
	 $excel.Quit()
	 email("#{arg1} file","PFA ids > #{arg1}: >",attachment)

end

def loadGMproductBrief(referenceID)

		if isElementPresent("GMProducts","Portal",10) then
			theElement("GMProducts","Portal").flash
			theElement("GMProducts","Portal").click
		       if isElementPresent("AdvSearch","Portal",10) then
		       	  if referenceID !="NA"
		       	  	theBrowser.input(:id=>"v_refNo").flash
		       	  	theBrowser.input(:id=>"v_refNo").send_keys "#{referenceID}"
		       	  end
		       	   theElement("AdvSearch","Portal").flash
		       	   theElement("AdvSearch","Portal").click
		       	     if isElementPresent("Results","Portal",10) then
		       	     	theElement("Results","Portal").flash
		       	     	theElement("Results","Portal").click
		       	     	if isElementPresent("PleaseConfirmWindow","Portal",3)
		       	     		theElement("YesButton","Portal").flash
		       	     	    theElement("YesButton","Portal").click
		       	     	end
		   	     	    if isElementPresent("HeaderPage","Header",2) then
		      	     	   theElement("HeaderPage","Header").flash
		      	     	   # puts "On Header tab"
		       	     	else
		       	     	   raise "GM Product brief page didn't load"
		       	     	end    
		     	    else	
		       	     	raise "GM Products not found"
		       	    end	

		        else
		        	raise "Advanced Search button not found on Brief(GM) page"
		        end
		else
		    raise "Open Briefs (GM) link not found"
		end 	

end

def loadAPPproductBrief(referenceID)

		if isElementPresent("APPProducts","Portal",10) then
			theElement("APPProducts","Portal").flash
			theElement("APPProducts","Portal").click
		       if isElementPresent("AdvSearch","Portal",10) then
		       	  if referenceID !="NA"
		       	  	theBrowser.input(:id=>"v_refNo").flash
		       	  	theBrowser.input(:id=>"v_refNo").send_keys "#{referenceID}"
		       	  end
		       	   theElement("AdvSearch","Portal").flash
		       	   theElement("AdvSearch","Portal").click
		       	     if isElementPresent("Results","Portal",10) then
		       	     	theElement("Results","Portal").flash
		       	     	theElement("Results","Portal").click
		       	     	if isElementPresent("PleaseConfirmWindow","Portal",3)
		       	     		theElement("YesButton","Portal").flash
		       	     	    theElement("YesButton","Portal").click
		       	     	end
		   	     	    if isElementPresent("HeaderPage","Header",2) then
		      	     	   theElement("HeaderPage","Header").flash
		      	     	   # puts "On Header tab"
		       	     	else
		       	     	   raise "APP Product brief page didn't load"
		       	     	end    
		     	    else	
		       	     	raise "APP Products not found"
		       	    end	

		        else
		        	raise "Advanced Search button not found on Brief(APP) page"
		        end
		else
		    raise "Open Briefs (APP) link not found"
		end 	

end

def loadQM(referenceID)

		if isElementPresent("QM","Portal",37) then
			theElement("QM","Portal").flash
			theElement("QM","Portal").click
	
		       if isElementPresent("AdvSearch","Portal",10) then
		       	  if referenceID !="NA"
		       	  	theBrowser.input(:id=>"v_styleNo").flash
		       	  	theBrowser.input(:id=>"v_styleNo").send_keys "#{referenceID}"
		       	  end
		       	   theElement("AdvSearch","Portal").flash
		       	   theElement("AdvSearch","Portal").click
		       end
		       	     
		else
		    raise "QM record not found"
		end 	

end


def fillCostingValues(set)

  case set

  when "set1"
  	# binding.pry
    theBrowser.input(:id=>"countryOfOriginId$0").parent.img.click
	sleep 1
    theBrowser.div(:class=>"x-layer x-combo-list dropdownList dropdownList-countryOfOriginId$0").div.div(:class=>"x-combo-list-item",:text=>"#{$cOrigin}").click

	theBrowser.input(:id=>"portOfLoadingId$0").parent.img.click
	sleep 1
    theBrowser.div(:class=>"x-layer x-combo-list dropdownList dropdownList-portOfLoadingId$0").div.div(:class=>"x-combo-list-item",:text=>"#{$portOfLoading}").click


  	   selectDropdownValue("Currency",$currency,"Costing")
  	   # selectJSdropdown("Currency",$currency,"Costing")
       onElement("VendorUnitCost","set","6.56","Costing")
       onElement("LCLVendorUnitCost","set","6.56","Costing")
       onElement("FCLVendorUnitCost","set","6.56","Costing")
       onElement("ReOrderDays","set","90","Costing")
       onElement("ProdLeadTime","set","60","Costing")
       theBrowser.input(:id=>"moq$0").send_keys "5000"
       theBrowser.input(:id=>"mcq$0").send_keys "300"
       # theBrowser.execute_script("return document.getElementById('moq$0').value='5000'")
       # theBrowser.execute_script("return document.getElementById('mcq$0').value='300'")
       onElement("OuterHeight","set","57","Costing")
       onElement("OuterWidth","set","50","Costing")
       onElement("OuterLength","set","16","Costing")
       onElement("OuterWeight","set","6","Costing")
       onElement("InnerHeight","set","57","Costing")
       onElement("InnerWidth","set","50","Costing")
       onElement("InnerLength","set","16","Costing")
       onElement("InnerWeight","set","16","Costing") 
       onElement("InstoreHeight","set","2.5","Costing")    
       onElement("InstoreWidth","set","46","Costing")
       onElement("InstoreLength","set","76","Costing")
       onElement("InstoreWeight","set","0.4","Costing")
       onElement("UnitsPerOuter","set","15","Costing")
       onElement("InnersPerOuter","set","1","Costing")                             
       onElement("UnitsPerInner","set","15","Costing")
       selectDropdownValue("PackageType","BAG","Costing") 
       selectDropdownValue("SupplyChainEvent","BULK FCL","Costing") 
       #  selectDropdownValue("CountryOfOrigin",$cOrigin,"Costing")
       # selectDropdownValue("PortOfLoading",$portOfLoading,"Costing")
       # selectJSdropdown("CountryOfOrigin",$cOrigin,"Costing")
       # selectJSdropdown("PortOfLoading",$portOfLoading,"Costing")
       selectDropdownValue("FCLLCL","LCL","Costing") 
       selectDropdownValue("FactoryPacked","No","Costing")
  

  when "set2"
  	# binding.pry
    theBrowser.input(:id=>"countryOfOriginId$0").parent.img.click
	sleep 1
    theBrowser.div(:class=>"x-layer x-combo-list dropdownList dropdownList-countryOfOriginId$0").div.div(:class=>"x-combo-list-item",:text=>"#{$cOrigin}").click

	theBrowser.input(:id=>"portOfLoadingId$0").parent.img.click
	sleep 1
    theBrowser.div(:class=>"x-layer x-combo-list dropdownList dropdownList-portOfLoadingId$0").div.div(:class=>"x-combo-list-item",:text=>"#{$portOfLoading}").click

  	selectDropdownValue("Currency",$currency,"Costing")
  	   # selectJSdropdown("Currency",$currency,"Costing")
       onElement("VendorUnitCost","set","12.5","Costing")
       onElement("LCLVendorUnitCost","set","12.5","Costing")
       onElement("FCLVendorUnitCost","set","12.5","Costing")
       onElement("ReOrderDays","set","75","Costing")
       onElement("ProdLeadTime","set","60","Costing")
       theBrowser.input(:id=>"moq$0").send_keys "5000"
       theBrowser.input(:id=>"mcq$0").send_keys "300"
       # theBrowser.execute_script("return document.getElementById('moq$0').value='5000'")
       # theBrowser.execute_script("return document.getElementById('mcq$0').value='300'")
       onElement("OuterHeight","set","84","Costing")
       onElement("OuterWidth","set","44.5","Costing")
       onElement("OuterLength","set","44.5","Costing")
       onElement("OuterWeight","set","10","Costing")
       onElement("InnerHeight","set","84","Costing")
       onElement("InnerWidth","set","44.5","Costing")
       onElement("InnerLength","set","44.5","Costing")
       onElement("InnerWeight","set","10","Costing")
       onElement("InstoreHeight","set","55","Costing")    
       onElement("InstoreWidth","set","44","Costing")
       onElement("InstoreLength","set","44","Costing")
       onElement("InstoreWeight","set","2.5","Costing")
       onElement("UnitsPerOuter","set","3","Costing")
       onElement("InnersPerOuter","set","1","Costing")                             
       onElement("UnitsPerInner","set","3","Costing")
       selectDropdownValue("PackageType","BAG","Costing") 
       selectDropdownValue("SupplyChainEvent","BULK FCL","Costing") 
       #  selectDropdownValue("CountryOfOrigin",$cOrigin,"Costing")
       # selectDropdownValue("PortOfLoading",$portOfLoading,"Costing")
       # selectJSdropdown("CountryOfOrigin",$cOrigin,"Costing")
       # selectJSdropdown("PortOfLoading",$portOfLoading,"Costing")
       selectDropdownValue("FCLLCL","LCL","Costing") 
       selectDropdownValue("FactoryPacked","No","Costing")

   end             
end