require 'chronic'
require 'time'
require 'colorize'
require 'date'
require 'win32ole'

$col_dict=WIN32OLE.new('Scripting.Dictionary')
$po_field_dict=WIN32OLE.new('Scripting.Dictionary')

def doRFSD

 begin
  
        # case $type
      
        # when "GM"
         #     loadGMproductBrief("#{$prod_ref_id}")

         #when "APP"
          #    loadAPPproductBrief("#{$prod_ref_id}")

          #when "QM"
               loadQM("#{$v_styleNo}")
  end   


        ###### Go to RFSD
      theBrowser.a(:xpath,"//a[contains(text(),'Sample & Document Submission')]").flash
      theBrowser.a(:xpath,"//a[contains(text(),'Sample & Document Submission')]").click

      #sleep 2
     
     # theBrowser.div(:id=>"btn_search").flash

      sleep 5

    
 
     theBrowser.div(:id=>"btn_searchAdv").flash
     theBrowser.div(:id=>"btn_searchAdv").click

     sleep 2

     #theBrowser.div(:class=>"x-grid3-cell-inner x-grid3-col-styleNo").flash
     
     print "zzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzz"



     theBrowser.div(:xpath,"//div[text()='New']//parent::td//parent::tr//div[@class='x-grid3-cell-inner x-grid3-col-styleNo']").click
     # theBrowser.div(:id=>"btn_search").click

     # sleep 1

   # theElement("Search","Portal").flash    
   #  theElement("Search","Portal").click

  #theBrowser.div(:xpath,"//div[text()='Red Tag Sample']//parent::td//parent::tr//div[@class='x-grid3-row-checker']").click
  #theBrowser.div(:xpath,"//div[text()='Test report product safety (toxicology liquid filled articles)']//parent::td//parent::tr//div[@class='x-grid3-row-checker']").click
  #theBrowser.div(:xpath,"//div[text()='Development Sample']//parent::td//parent::tr//div[@class='x-grid3-row-checker']").click
 #theBrowser.div(:xpath,'//*[@id="ext-gen8376"]/div[4]/table/tbody/tr[1]/td[4]/div', :class=>"x-grid3-row-checker").click
   
  #  sleep 1

    #theBrowser.div(:id=>"request").click
   
   #   sleep 3

 #  theBrowser.div(:class=>"x-grid3-cell-inner x-grid3-col-sellColor").click
 #  theBrowser.div(:id=>"sellColor_ccg_core_select_").send_keys "CMLFGE"

 #   sleep 3

 #  theBrowser.div(:id=>"selectedEmail").flash
 #  theBrowser.div(:id=>"selectedEmail").to_subtype.clear

    #  sleep 1

 # theBrowser.div(:id=>"selectedEmail").send_keys "chirag.pandya@target.com.au"

 # sleep 4




         #theBrowser.div(:xpath,"//div[text()='Development Sample']//parent::td//parent::tr//div[@class='inline_button']").click
     
     
    # sleep 1 

     #theBrowser.div(:id=>"btn_closeSampleDocDetail").flash
     #theBrowser.div(:id=>"btn_closeSampleDocDetail").click

    
    #  dc_date=DateTime.now.next_month.next_month.strftime "%d/%m/%Y"
     #submit_date=DateTime.now.next_month.next_month.next_month.strftime "%d/%m/%Y"

    # theBrowser.div(:class=>"x-grid3-cell-inner x-grid3-col-checker").click
     
     #theBrowser.send_keys [:control, 'a'], :backspace
     
    sleep 1
   #print theBrowser.div(:css,".ytb-text").text




    # theBrowser.div(:xpath,"//div[@class="x-grid3-cell-inner x-grid3-col-submitDate"]").click
    # theBrowser.div(:class=>"x-form-field-wrap  x-trigger-wrap-focus").click
#    print "!!!!!!!!!!!!!!!!!!!!!!!!!!!!Sample Submission is the next step!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!"
     #theBrowser.element(:xpath,"//div[text()='Development Sample']//parent::td//parent::tr//div[@class='x-grid3-col x-grid3-cell x-grid3-td-submitDate ']").flash
     #theBrowser.element(:xpath,"//div[text()='Development Sample']//parent::td//parent::tr//div[@class='x-grid3-col x-grid3-cell x-grid3-td-submitDate ']").click
    
   
     #theBrowser.element(:xpath,"//div[text()='Development Sample']//parent::td//parent::tr//div[@class='x-grid3-cell-inner x-grid3-col-submitDate']").click
    #driver.findElement(:xpath,"//div[text()='Online Sample']//parent::td//parent::tr//div[@class='x-form-trigger x-form-date-trigger ']").flash()
    
###############***********************************************

  print "CPCPCPCPCPCPCPCPCPCPCPCPCPCPCPCPCP"
  #theBrowser.element(:css,"#samplesGrid .x-grid3-scroller .x-grid3-body .x-grid3-row.x-grid3-row-first .x-grid3-row-table .x-grid3-col.x-grid3-cell.x-grid3-td-submitDate ").flash
  #theBrowser.element(:css,"#samplesGrid .x-grid3-scroller .x-grid3-body .x-grid3-row.x-grid3-row-first .x-grid3-row-table .x-grid3-col.x-grid3-cell.x-grid3-td-submitDate ").click
  #theBrowser.element(:css,".x-layer .x-form-trigger.x-form-date-trigger").flash 
  #theBrowser.element(:css,".x-layer .x-form-trigger.x-form-date-trigger").click  
  #sleep 2
  #theBrowser.element(:css,".x-date-picker.x-unselectable .x-date-bottom .x-btn-text").flash
  #sleep 1
  #theBrowser.element(:css,".x-date-picker.x-unselectable .x-date-bottom .x-btn-text").click
  #sleep 1
  #theBrowser.element(:css,"#samplesGrid .x-grid3-scroller .x-grid3-row.x-grid3-row-first .x-grid3-row-checker").flash
  #sleep 1
  #theBrowser.element(:css,"#samplesGrid .x-grid3-scroller .x-grid3-row.x-grid3-row-first .x-grid3-row-checker").click
 # sleep 1
  #theBrowser.div(:id=>"submitSamples").flash
  #theBrowser.div(:id=>"submitSamples").click
  #sleep 3
  theBrowser.div(:id=>"btn_search").flash
  theBrowser.div(:id=>"btn_search").click

  print "TTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTT"

  sleep 3
  theBrowser.a(:xpath,"//a[contains(text(),'Request for Sample/Document')]").flash
  theBrowser.a(:xpath,"//a[contains(text(),'Request for Sample/Document')]").click
  sleep 3

  theBrowser.input(:id=>"v_description").flash
  theBrowser.input(:id=>"v_description").send_keys "CP"
  sleep 2
  
  theBrowser.div(:id=>"btn_searchAdv").flash
  theBrowser.div(:id=>"btn_searchAdv").click
  sleep 1
  theBrowser.element(:css,".x-grid3-cell-inner.x-grid3-col-styleNo").click
  sleep 3
  theBrowser.div(:xpath,"//div[text()='Online Sample']//parent::td//parent::tr//div[@class='x-grid3-row-checker']").click
 # theBrowser.div(:xpath,"//div[text()='Test report product safety (toxicology liquid filled articles)']//parent::td//parent::tr//div[@class='x-grid3-row-checker']").click
  #theBrowser.div(:xpath,"//div[text()='Development Sample']//parent::td//parent::tr//div[@class='x-grid3-row-checker']").click
  sleep 1
  #theBrowser.div(:xpath,"//div[text()='New']//parent::td//parent::tr//div[@class='x-grid3-row-checker']").flash
  #theBrowser.div(:xpath,"//div[text()='New']//parent::td//parent::tr//div[@class='x-grid3-row-checker']").click
  sleep 1
  theBrowser.div(:id=>"request").flash
  theBrowser.div(:id=>"request").click
  sleep 3
  theBrowser.div(:id=>"selectedEmail").flash
  theBrowser.div(:id=>"selectedEmail").to_subtype.clear
  sleep 1
  theBrowser.div(:id=>"selectedEmail").send_keys "chirag.pandya@target.com.au"
  sleep 2

  print "PPPPPPPPPP123123123PPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPP"
  print "Pwwwwwwwwwwwwwwwwwww"
  theBrowser.element(:css,".x-grid3-cell-inner.x-grid3-col-sellColor").flash
  theBrowser.element(:css,".x-grid3-cell-inner.x-grid3-col-sellColor").click

  print "Pwwwwwwwwwwwwwwwwwww"
  theBrowser.element(:css,".x-trigger-wrap-focus .x-form-trigger.x-form-arrow-trigger").flash
  print "JJJJJKKKKKKKKKKKKKKKKKKK****@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@"
  theBrowser.element(:css,".x-trigger-wrap-focus .x-form-trigger.x-form-arrow-trigger").click
  print "CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC"
  sleep 2
  theBrowser.element(:css,".x-layer.x-combo-list.dropdownList .x-combo-list-inner .x-combo-list-item.x-combo-selected").flash
  theBrowser.element(:css,".x-layer.x-combo-list.dropdownList .x-combo-list-inner .x-combo-list-item.x-combo-selected").click

  #print theBrowser.element(:css,"#sellColor_ccg_core_select_")
  #theBrowser.element(:css,".x-form-field-wrap.x-trigger-wrap-focus .x-form-trigger.x-form-arrow-trigger ").options[1].click

  
  sleep 2

  print "$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$"

  #theBrowser.element(:css,"#selectedSampleDoc .x-panel-bwrap .x-panel-body .x-grid3-scroller .x-grid3-row-table .x-grid3-col.x-grid3-cell.x-grid3-td-checker.x-grid3-cell-first .x-grid3-cell-inner.x-grid3-col-checker .x-grid3-row-checker").flash
  theBrowser.element(:css,"#selectedSampleDoc .x-grid3-scroller .x-grid3-body .x-grid3-row.x-grid3-row-first .x-grid3-col.x-grid3-cell.x-grid3-td-checker.x-grid3-cell-first .x-grid3-cell-inner.x-grid3-col-checker").flash
  theBrowser.element(:css,"#selectedSampleDoc .x-grid3-scroller .x-grid3-body .x-grid3-row.x-grid3-row-first .x-grid3-col.x-grid3-cell.x-grid3-td-checker.x-grid3-cell-first .x-grid3-cell-inner.x-grid3-col-checker").click
  #theBrowser.element(:xpath,"//div[text()='L']//parent::td//parent::tr//div[@class='x-grid3-row-checker']").flash
  #theBrowser.element(:xpath,"//div[text()='L']//parent::td//parent::tr//div[@class='x-grid3-row-checker']").click
  sleep 1
  #theBrowser.element(:css,".#selectedSampleDoc .x-panel-bwrap .x-panel-body .x-grid3-scroller .x-grid3-row-table .x-grid3-col.x-grid3-cell.x-grid3-td-checker.x-grid3-cell-first .x-grid3-cell-inner.x-grid3-col-checker .x-grid3-row-checker").click
  sleep 2
  theBrowser.div(:id=>"copy").flash
  theBrowser.div(:id=>"copy").click
  print "PPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPP"
  theBrowser.element(:css,".x-grid3-row.x-grid3-row-last .x-grid3-row-table .x-grid3-cell-inner.x-grid3-col-sellColor").flash
  theBrowser.element(:css,".x-grid3-row.x-grid3-row-last .x-grid3-row-table .x-grid3-cell-inner.x-grid3-col-sellColor").click
  #theBrowser.element(:css,".x-grid3-row-last .x-grid3-col-sellColor").flash
  #theBrowser.element(:css,".x-grid3-row-last .x-grid3-col-sellColor").click
  sleep 2
  theBrowser.element(:css,".x-form-field-wrap.x-trigger-wrap-focus .x-form-trigger.x-form-arrow-trigger").flash
  theBrowser.element(:css,".x-form-field-wrap.x-trigger-wrap-focus .x-form-trigger.x-form-arrow-trigger").click
  sleep 3

  theBrowser.element(:xpath,"//div[text()='Grey']").flash
  theBrowser.element(:xpath,"//div[text()='Grey']").click


  #theBrowser.element(:css,".x-form-field-wrap.x-trigger-wrap-focus .x-form-text.x-form-field.x-form-focus").flash
  #theBrowser.element(:css,".x-form-field-wrap.x-trigger-wrap-focus .x-form-text.x-form-field.x-form-focus").click
  #theBrowser.element(:css,"#selectedSampleDoc .x-panel.x-grid-panel .x-panel-bwrap .x-grid3 .x-grid3-viewport .x-grid3-scroller .x-layer.x-editor.x-small-editor.x-grid-editor .x-form-field-wrap.x-trigger-wrap-focus .x-form-trigger.x-form-arrow-trigger").flash
  #theBrowser.element(:css,"#selectedSampleDoc .x-panel.x-grid-panel .x-panel-bwrap .x-grid3 .x-grid3-viewport .x-grid3-scroller .x-layer.x-editor.x-small-editor.x-grid-editor .x-form-field-wrap.x-trigger-wrap-focus .x-form-trigger.x-form-arrow-trigger").click
  #sleep 1
  print "^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^"

  #theBrowser.div(:id=>"sellColor_ccg_core_select_").flash
  #theBrowser.input(:id=>"sellColor_ccg_core_select_").click
  #theBrowser.send_keys :down

  
  sleep 3
  theBrowser.div(:id=>"add").flash
  theBrowser.div(:id=>"add").click

  print "It has been working till this point.... continue from here....04th May 2018"
  #sleep 1
  #theBrowser.element(:css,".x-form-field-wrap.x-trigger-wrap-focus .x-form-text.x-form-field.x-form-focus").flash
  #theBrowser.element(:css,".x-form-field-wrap.x-trigger-wrap-focus .x-form-text.x-form-field.x-form-focus").click
  #sleep 1
  #theBrowser.element(:css,".x-form-field-wrap.x-trigger-wrap-focus .x-form-trigger.x-form-arrow-trigger").click
  #sleep 1
  theBrowser.element(:css,"#selectedSampleDoc .x-grid3-row-last .x-grid3-cell-inner.x-grid3-col-typeId").flash
  theBrowser.element(:css,"#selectedSampleDoc .x-grid3-row-last .x-grid3-cell-inner.x-grid3-col-typeId").click
  sleep 1
  #theBrowser.element(:css,".x-form-field-wrap.x-trigger-wrap-focus .x-form-trigger.x-form-arrow-trigger").flash
  #theBrowser.element(:css,".x-form-field-wrap.x-trigger-wrap-focus .x-form-trigger.x-form-arrow-trigger").click
  #theBrowser.options(:text,"Document").click
  sleep 1
  #theBrowser.div(:id=>"typeId_ccg_core_select_").flash
  #theBrowser.div(:id=>"typeId_ccg_core_select_").send_keys "Document"
  #theBrowser.div(:id=>"typeId_ccg_core_select_").click

  print "CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC"

  sleep 2
  theBrowser.element(:css,"#selectedSampleDoc .x-grid3-row-last .x-grid3-cell-inner x-grid3-col-sampleReqId").flash
  theBrowser.element(:css,"#selectedSampleDoc .x-grid3-row-last .x-grid3-cell-inner x-grid3-col-sampleReqId").click
  sleep 1
  theBrowser.element(:css,".x-form-field-wrap.x-trigger-wrap-focus .x-form-trigger.x-form-arrow-trigger").flash
  theBrowser.element(:css,".x-form-field-wrap.x-trigger-wrap-focus .x-form-trigger.x-form-arrow-trigger").click 
  sleep 1
  theBrowser.div(:id=>"sampleReqId_ccg_core_select_").flash
  theBrowser.div(:id=>"sampleReqId_ccg_core_select_").send_keys "Bulk Test Report"
  sleep 1
  theBrowser.element(:xpath,"//div[text()='Bronze']").flash
  theBrowser.element(:xpath,"//div[text()='Bronze']").click
  sleep 1

  print "^&^&^&^&^&^&^&^&^&^&^&^&^&^&^&^^&^&^&^&^&^&^&^&^&&^&^&^"

  sleep 1
  theBrowser.div(:id=>"btn_confirmRequest").flash
  #theBrowser.div(:id=>"btn_confirmRequest").click
  sleep 3
  theBrowser.div(:id=>"btn_search").click



  #theBrowser.element(:css,"#selectedSampleDoc .x-grid3-header .x-grid3-header-offset .x-grid3-hd.x-grid3-cell.x-grid3-td-checker.x-grid3-cell-first .x-grid3-hd-inner.x-grid3-hd-checker .x-grid3-hd-checker").flash
  #theBrowser.element(:css,"#selectedSampleDoc .x-grid3-header .x-grid3-header-offset .x-grid3-hd.x-grid3-cell.x-grid3-td-checker.x-grid3-cell-first .x-grid3-hd-inner.x-grid3-hd-checker .x-grid3-hd-checker").click
  #theBrowser.element(:css,"#selectedSampleDoc .x-grid3-header .x-grid3-header-offset .x-grid3-hd.x-grid3-cell.x-grid3-td-checker.x-grid3-cell-first .x-grid3-hd-inner.x-grid3-hd-checker .x-grid3-hd-checker").click
 # theBrowser.element(:css,"#selectedSampleDoc .x-grid3-header .x-grid3-header-offset .x-grid3-hd.x-grid3-cell.x-grid3-td-checker.x-grid3-cell-first .x-grid3-hd-inner.x-grid3-hd-checker .x-grid3-hd-checker").click
  
  #theBrowser.element(:css,"#selectedSampleDoc .x-panel-ml .x-grid3 .x-grid3-scroller .x-grid3-row.x-grid3-dirty-row.x-grid3-row-last .x-grid3-row-table .x-grid3-col.x-grid3-cell.x-grid3-td-sellColor.x-grid3-dirty-cell .x-grid3-cell-inner.x-grid3-col-sellColor").flash
  #theBrowser.element(:css,"#selectedSampleDoc .x-panel-ml .x-grid3 .x-grid3-scroller .x-grid3-row.x-grid3-dirty-row.x-grid3-row-last .x-grid3-row-table .x-grid3-col.x-grid3-cell.x-grid3-td-sellColor.x-grid3-dirty-cell .x-grid3-cell-inner.x-grid3-col-sellColor").click

  #theBrowser.element(:css,".x-form-text.x-form-field.x-form-focus").flash
  
  #theBrowser.element(:css,".x-form-text.x-form-field.x-form-focus").to_subtype.clear
  

  
  #theBrowser.element(:css,".x-layer.x-editor.x-small-editor.x-grid-editor .x-form-field-wrap.x-trigger-wrap-focus .x-form-text.x-form-field.x-form-focus").click

  print "************************************************************************************************************************"
  #theBrowser.element(:css,".x-grid3-body .x-grid3-row.x-grid3-dirty-row.x-grid3-row-last .x-grid3-col.x-grid3-cell.x-grid3-td-sellColor.x-grid3-dirty-cell .x-grid3-cell-inner.x-grid3-col-sellColor").flash
  #theBrowser.element(:css,".x-grid3-body .x-grid3-row.x-grid3-dirty-row.x-grid3-row-last .x-grid3-col.x-grid3-cell.x-grid3-td-sellColor.x-grid3-dirty-cell .x-grid3-cell-inner.x-grid3-col-sellColor").click
  sleep 1
  #theBrowser.element(:css,"#selectedSampleDoc .x-panel-bwrap .x-panel-body .x-grid3-scroller .x-grid3-body .x-grid3-row.x-grid3-dirty-row.x-grid3-row-last .x-grid3-col.x-grid3-cell.x-grid3-td-sellColor.x-grid3-dirty-cell .x-grid3-cell-inner.x-grid3-col-sellColor").flash
  #theBrowser.element(:css,"#selectedSampleDoc .x-panel-bwrap .x-panel-body .x-grid3-scroller .x-grid3-body .x-grid3-row.x-grid3-dirty-row.x-grid3-row-last .x-grid3-col.x-grid3-cell.x-grid3-td-sellColor.x-grid3-dirty-cell .x-grid3-cell-inner.x-grid3-col-sellColor").click
  #theBrowser.element(:css,"#selectedSampleDoc .x-grid3-row.x-grid3-dirty-row.x-grid3-row-last .x-grid3-row-table .x-grid3-col.x-grid3-cell.x-grid3-td-sellColor .x-grid3-cell-inner.x-grid3-col-sellColor").flash
  #theBrowser.element(:css,".x-layer.x-editor.x-small-editor.x-grid-editor .x-form-field-wrap.x-trigger-wrap-focus .x-form-text.x-form-field.x-form-focus").flash
  #theBrowser.element(:css,".x-grid3-row.x-grid3-dirty-row.x-grid3-row-last .x-grid3-row-table .x-grid3-col.x-grid3-cell.x-grid3-td-sellColor.x-grid3-dirty-cell .x-grid3-cell-inner.x-grid3-col-sellColor").flash
  #theBrowser.element(:css,".x-grid3-row.x-grid3-dirty-row.x-grid3-row-last .x-grid3-row-table .x-grid3-col.x-grid3-cell.x-grid3-td-sellColor.x-grid3-dirty-cell .x-grid3-cell-inner.x-grid3-col-sellColor").click
  #theBrowser.div(:xpath,"//div[@class='x-grid3-row   x-grid3-dirty-row  x-grid3-row-last ']//div[@class='x-grid3-cell-inner x-grid3-col-sellColor']//parent::td//parent::tr//div[@class='x-grid3-cell-inner x-grid3-col-sellColor']").flash

  #theBrowser.element(:css,"#selectedSampleDoc .x-panel-body .x-grid3-scroller .x-layer.x-editor.x-small-editor.x-grid-editor .x-form-field-wrap.x-trigger-wrap-focus .x-form-text.x-form-field.x-form-focus").flash
  #theBrowser.element(:css,"#selectedSampleDoc .x-panel-body .x-grid3-scroller .x-layer.x-editor.x-small-editor.x-grid-editor .x-form-field-wrap.x-trigger-wrap-focus .x-form-text.x-form-field.x-form-focus").click
  #print theBrowser.element(:css,".x-grid3-row.x-grid3-dirty-row.x-grid3-row-last tr:nth-child(1) .x-grid3-col.x-grid3-cell.x-grid3-td-sellColor .x-grid3-cell-inner.x-grid3-col-sellColor")
  #theBrowser.element(:css,".x-grid3-row.x-grid3-dirty-row.x-grid3-row-last tr:nth-child(1) .x-grid3-col.x-grid3-cell.x-grid3-td-sellColor .x-grid3-cell-inner.x-grid3-col-sellColor").flash
  #theBrowser.element(:css,".x-grid3-row.x-grid3-dirty-row.x-grid3-row-last tr:nth-child(1) .x-grid3-col.x-grid3-cell.x-grid3-td-sellColor .x-grid3-cell-inner.x-grid3-col-sellColor").click
  #theBrowser.element(:css,".x-grid3-row.x-grid3-dirty-row.x-grid3-row-last .x-grid3-col.x-grid3-cell.x-grid3-td-sellColor .x-grid3-cell-inner.x-grid3-col-sellColor").click
  #print"asdsajdsadsadsadsadsad"
  #theBrowser.element(:css,".x-grid3-row.x-grid3-dirty-row.x-grid3-row-last .x-grid3-cell-inner.x-grid3-col-sellColor").send_keys("Testing")
  #theBrowser.element(:css,".x-grid3-row.x-grid3-dirty-row.x-grid3-row-last .x-grid3-col.x-grid3-cell.x-grid3-td-sellColor.x-grid3-dirty-cell .x-grid3-cell-inner.x-grid3-col-sellColor").flash
  #theBrowser.element(:css,".x-grid3-row.x-grid3-dirty-row.x-grid3-row-last .x-grid3-col.x-grid3-cell.x-grid3-td-sellColor.x-grid3-dirty-cell .x-grid3-cell-inner.x-grid3-col-sellColor").click
  #print theBrowser.div(:css,".x-grid3-row.x-grid3-dirty-row.x-grid3-row-last  .x-grid3-cell-inner.x-grid3-col-sellColor")
  #theBrowser.div(:css,".x-grid3-row.x-grid3-dirty-row.x-grid3-row-last  .x-grid3-cell-inner.x-grid3-col-sellColor").flash
  #theBrowser.div(:css,".x-grid3-row.x-grid3-dirty-row.x-grid3-row-last  .x-grid3-cell-inner.x-grid3-col-sellColor").click 
  #print theBrowser.element(:css,".x-grid3-row.x-grid3-dirty-row.x-grid3-row-last .x-grid3-cell-inner.x-grid3-col-sellColor")
  #theBrowser.element(:css,".x-grid3-row.x-grid3-dirty-row.x-grid3-row-last .x-grid3-cell-inner.x-grid3-col-sellColor").flash
  #theBrowser.element(:css,".x-grid3-row.x-grid3-dirty-row.x-grid3-row-last .x-grid3-cell-inner.x-grid3-col-sellColor").click
  #print theBrowser.div(:id=>"sellColor_ccg_core_select_").div(:class=>"x-layer x-editor x-small-editor x-grid-editor")
  #theBrowser.div(:id=>"sellColor_ccg_core_select_").div(:class=>"x-layer x-editor x-small-editor x-grid-editor").flash
  #theBrowser.div(:id=>"sellColor_ccg_core_select_").div(:class=>"x-layer x-editor x-small-editor x-grid-editor").click
 

  #theBrowser.element(:xpath,"//div[@class='x-layer x-editor x-small-editor x-grid-editor']//div[@class='x-form-field-wrap  x-trigger-wrap-focus']").flash
  #theBrowser.element(:xpath,"//div[@class='x-grid3-row   x-grid3-dirty-row  x-grid3-row-last ']//div[@class='x-grid3-cell-inner x-grid3-col-sellColor']//parent::td//parent::tr//div[@class='x-grid3-cell-inner x-grid3-col-sellColor']").click
  #theBrowser.element(:css,".x-layer.x-editor.x-small-editor.x-grid-editor .x-form-field-wrap.x-trigger-wrap-focus .x-form-text.x-form-field.x-form-focus").click
  #theBrowser.element(:css,".x-grid3-row.x-grid3-dirty-row.x-grid3-row-last .x-grid3-cell-inner.x-grid3-col-sellColor").flash
  #theBrowser.element(:css,".x-grid3-row.x-grid3-dirty-row.x-grid3-row-last .x-grid3-cell-inner.x-grid3-col-sellColor").click
  #heBrowser.element(:css,"#selectedSampleDoc .x-grid3-row.x-grid3-dirty-row.x-grid3-row-last .x-grid3-row-table .x-grid3-col.x-grid3-cell.x-grid3-td-sellColor .x-grid3-cell-inner.x-grid3-col-sellColor ").click
  #theBrowser.div(:id=>"sampleReqId_ccg_core_select_").click
  #theBrowser.element(:css,".x-form-field-wrap.x-trigger-wrap-focus .x-form-text.x-form-field.x-form-focus").click
  #theBrowser.element(:css,".x-layer.x-editor.x-small-editor.x-grid-editor .x-form-field-wrap.x-trigger-wrap-focus #sellColor_ccg_core_select_").click
  #theBrowser.element(:css,".x-layer.x-editor.x-small-editor.x-grid-editor .x-form-text.x-form-field.x-form-focus").click
  #theBrowser.element(:css,"#selectedSampleDoc .x-panel.x-grid-panel .x-panel-bwrap .x-grid3 .x-grid3-viewport .x-grid3-scroller .x-layer.x-editor.x-small-editor.x-grid-editor .x-form-field-wrap.x-trigger-wrap-focus .x-form-trigger.x-form-arrow-trigger").flash
  print "iiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiii"
  #theBrowser.element(:css,"#selectedSampleDoc .x-panel.x-grid-panel .x-panel-bwrap .x-grid3 .x-grid3-viewport .x-grid3-scroller .x-layer.x-editor.x-small-editor.x-grid-editor .x-form-field-wrap.x-trigger-wrap-focus .x-form-trigger.x-form-arrow-trigger").click



 # theBrowser.div(:class=>"x-grid3-cell-inner x-grid3-col-sellColor").click
 #  theBrowser.div(:id=>"sellColor_ccg_core_select_").send_keys "CMLFGE"
  sleep 1
  theBrowser.div(:id=>"btn_confirmRequest").flash
  #theBrowser.div(:id=>"btn_confirmRequest").click
  sleep 3
  theBrowser.div(:id=>"btn_search").click







 # if(1st table.exists)
  # {
   #   String[] check = driver.findElements(:xpath, "//form[@id='tab_headerForm']").child()
    #  For ( int i= 0; i <check.count() ; i ++ )
     # {
      #  if(tr.exits)
       #   {
        #    click
         #   write
          #}
      #}
   #}

#}
###################################************************************

# driver.findElement(:xpath,"//form[@id='tab_headerForm']//parent:://div[text()='Online Sample']//parent::td//parent::tr//div[@class='x-grid3-cell-inner x-grid3-col-submitDate']").click
# theBrowser.send_keys [:control, 'a'], :backspace
#driver.findElement(:xpath,"//form[@id='tab_headerForm']//parent:://div[text()='Online Sample']//parent::td//parent::tr//div[@class='x-grid3-cell-inner x-grid3-col-submitDate']").send_keys DateTime.now.next_month.next_month.strftime "%d/%m/%Y"
#driver.findElement(:xpath,"//form[@id='tab_headerForm']//parent:://div[@id='samplesGrid']//parent::/div[@id='x-grid3-scroller']//parent:://div[text()='Online Sample']//parent::td//parent::tr//div[@class='x-grid3-cell-inner x-grid3-col-submitDate']").flash()
#driver.findElement(:xpath,"//form[@id='tab_headerForm']//parent:://div[@id='samplesGrid']//parent::/div[@id='x-grid3-scroller']//parent:://div[text()='Online Sample']//parent::td//parent::tr//div[@class='x-grid3-cell-inner x-grid3-col-submitDate']").send_keys DateTime.now.next_month.next_month.strftime "%d/%m/%Y"
   # sleep 2
    #driver.findElement(:xpath,"//div[text()='Advertising Sample']//parent::td//parent::tr//div[@class='x-form-trigger x-form-date-trigger ']").click()
    # theBrowser.element(:xpath,"//div[text()='Development Sample']//parent::td//parent::tr//div[@class='x-form-trigger x-form-textarea-trigger']").click
     #theBrowser.element(:xpath,"//div[text()='Development Sample']//parent::td//parent::tr//div[@class='x-grid3-cell-inner x-grid3-col-tplTypeId']").send_keys "Hi CP"

     #theBrowser.element(:xpath,"//div[id='submitDate/cDateField']//parent::td//parent::tr//div[@class='x-grid3-col x-grid3-cell x-grid3-td-submitDate ']").flash
     #theBrowser.element(:xpath,"//div[text()='Development Sample']//parent::td//parent::tr//div[@class='x-grid3-col x-grid3-cell x-grid3-td-submitDate ']").click
     
    # sleep 2
     #theBrowser.div(:xpath,"//div[text()='Development Sample']//parent::td//parent::tr//div[@x-grid3-col x-grid3-cell x-grid3-td-submitDate ']").send_keys [:control, 'a'], :backspace
     #theBrowser.input(:id=>"submitDate/cDateField").send_keys "10/06/2018"

     #CPCPCPCPCPCP
    # theBrowser.element(:xpath,"//div[text()='Colour Sample']//parent::td//parent::tr//div[@class=' x-form-text x-form-field datefield_default_style x-form-focus']").flash
   
   # theBrowser.element(:xpath,"//div[text()='Colour Sample']//parent::td//parent::tr//div[@id='samples']").click
   # sleep 1
  # theBrowser.element(:xpath,"//div[text()='Colour Sample']//parent::td//parent::tr//div[@id='samples']").send_keys DateTime.now.next_month.next_month.strftime "%d/%m/%Y"

     #theBrowser.element(:xpath,"//div[id='Sample Type']//parent:://div[text()='Online Sample']//parent::td//parent::tr//div[@class='x-grid3-cell-inner x-grid3-col-submitDate']").send_keys DateTime.now.next_month.next_month.strftime "%d/%m/%Y"
     #theBrowser.div(:xpath,"//div[text()='Development Sample']//parent::td//parent::tr//div[@class='x-form-trigger x-form-date-trigger ']").click
     #theBrowser.div(:class=>"x-btn-text").click

    # theBrowser.div(:class=>"x-form-trigger x-form-date-trigger").click

     #theBrowser.div(:id=>"ext-gen1106").flash
     #theBrowser.div(:id=>"ext-gen1106").click
    # theBrowser.div(:id=>"submitDate/cDateField").click
     #theBrowser.div(:id=>"submitDate/cDateField").send_keys DateTime.now.next_month.next_month.strftime "%d/%m/%Y" 
   # theBrowser.div(:id=>"submitDate").send_keys "09/05/2018"
     #theBrowser.findElement(:id=>"submitDate/cDateField$0").send_keys "10/06/2018"
     #theBrowser.div(:class=>"x-grid3-cell-inner x-grid3-col-submitDate").click
     #theBrowser.div(:class=>"x-grid3-cell-inner x-grid3-col-submitDate").send_keys "10/06/2018"

   #theBrowser.input(:class=>"x-grid3-cell-inner x-grid3-col-submitDate").send_keys DateTime.now.next_month.next_month.strftime "%d/%m/%Y" 
     #theBrowser.input(:id=>"submitDate/cDateField").send_keys DateTime.now.next_month.next_month.strftime "%d/%m/%Y" 
   #  theBrowser.input(:class=>"x-grid3-cell-inner x-grid3-col-submitDate").send_keys DateTime.now.next_month.next_month.strftime "%d/%m/%Y" 
       
     #theBrowser.div(:class=>"x-btn-text").button(:text=>"Yes").flash
     #theBrowser.div(:class=>"x-btn-text").button(:text=>"Yes").click
     #theBrowser.input(:id=>"submitDate/cDateField").send_keys DateTime.now.next_month.next_month.strftime "%d/%m/%Y" 

     sleep 2

    # theBrowser.div(:class=>"x-grid3-cell-inner x-grid3-col-submitQty").send_keys "9"

     sleep 1

     #theBrowser.div(:class=>"x-btn-wrap x-btn").flash

     sleep 1

     #theBrowser.div(:class=>"x-btn-wrap x-btn").click


      #theBrowser.div(:xpath,"//div[text()='Development Sample']//parent::td//parent::tr//div[@class='x-grid3-cell-inner x-grid3-col-details']").flash
    # theBrowser.div(:xpath,"//div[text()='Red Tag Sample']//parent::td//parent::tr//div[@class='x-grid3-cell-inner x-grid3-col-details']").flash
      #theBrowser.div(:xpath,"//div[text()='Pre Production Sample']//parent::td//parent::tr//div[@class='x-grid3-cell-inner x-grid3-col-details']").flash
     # sleep 1
 
    #  theBrowser.div(:xpath,"//div[text()='Development Sample']//parent::td//parent::tr//div[@class='x-grid3-cell-inner x-grid3-col-details']").click
    #theBrowser.div(:xpath,"//div[text()='Red Tag Sample']//parent::td//parent::tr//div[@class='x-grid3-cell-inner x-grid3-col-details']").click
     # theBrowser.div(:xpath,"//div[text()='Pre Production Sample']//parent::td//parent::tr//div[@class='x-grid3-cell-inner x-grid3-col-details']").click
     # sleep 2

     # theBrowser.div(:id=>"btn_submitDoc").click

      sleep 1

      theElement("YesButton","Portal").flash
      theElement("YesButton","Portal").click

    # theBrowser.div(:class=>"x-form-trigger x-form-arrow-trigger").click

      sleep 2

      #driver.findElement(By.xpath
      #selectByVisibleText.selectByVisibleText("Approved")
      #theBrowser.div(:id=>"ext-gen2844",:div=>"Approved").flash 
      #theBrowser.div(:class=>"x-combo-list-item ", :index=> Approved)
      #theBrowser.div(:xpath,"//div[text()='Approved']").flash
      #selectDropdownValue("QMAssessments","Approved","QualityManagement") 
      # Select dropdown=new Select(driver.findElement(By.name(“partyStatus”)));
      #dropdown.selectByValue(“Approved”);
      #theBrowser.div(:xpath,"//div[text()='Colour Sample']//parent::td//parent::tr//div[@class='x-grid3-row-checker']").click
      #theBrowser.div(:class=>" x-combo-list-item").click
      #theBrowser.div(:class=>" x-form-text x-form-field x-form-focus").send_keys "Approved"
      #theBrowser.select_list(:id=>'ext-gen2844').option(:index=> 0).select
      #selectCPDropdownValue("QMAssessments","Approved","QualityManagement")
      theBrowser.div(:name=>"partyStatus", :index=> Approved)


      sleep 2

      theBrowser.div(:id=>"btn_saveSampleDocDetail").click









    # theBrowser.div(:id=>"x-grid3-cell-inner x-grid3-col-styleNo").click



     # webDriver.findElement(:xpath,"//a[contains(text(),'QM(GM)')]").click

      #theElement("HomeTab","Portal").click
      #theBrowser.div(:CSS,"//a[onclick*='quality-qa-gm']").click

     #theBrowser.div(:xpath,"//a[onclick*='quality-qa-gm']").click

     # theElement("Applydefault","Keycode").click
     # sleep 1
     #   theElement("YesButton","Portal").flash
     #  theElement("YesButton","Portal").click
      #  sleep 4









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

      when "v_styleNo"
        $v_styleNo=cell_content
        # puts "#{$v_styleNo}}"
       
      else     
        # puts "Invalid column name or not required"
      end
    end   
    doRFSD
end

Given(/^wants to rfsd from "([^"]*)"$/) do |sheetName|
 
  excel = WIN32OLE.new('Excel.Application')
  excel.visible = false
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



