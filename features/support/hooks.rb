require 'watir-webdriver'
require 'pry'

#require 'data_magic'

    # puts __FILE__

  if ENV['DEBUG'] != true
    (ENV['RESULTSDIR'].nil?) ? results_dir =  results_dir = "results/#{Time.now.strftime("%d%m%Y_%H%M%S")}" : results_dir = ENV['RESULTSDIR']
  end

######## Select browser ###########
  case ENV['br']
    
  when "chrome"
    puts "Running on Chrome"
    chromedriver_path=File.expand_path('../../driverDependencies/chromedriver.exe', File.dirname(__FILE__))
    profile=Selenium::WebDriver::Chrome.driver_path = chromedriver_path
    browser = Watir::Browser.new :chrome, :switches => %w[--disable-popup-blocking]

     # profile = Selenium::WebDriver::Chrome::Profile.new
     # # profile.native_events = false
     # browser = Watir::Browser.new :chrome, :profile => profile

  when "firefox"
    puts "Running on Firefox"    
     profile = Selenium::WebDriver::Firefox::Profile.new
     profile.native_events = false
     browser = Watir::Browser.new :firefox, :profile => profile
     # for latest firefox download geckodriver ,and append its path in environment variables

  when "grid"
    puts "Running on Grid"
    caps = Selenium::WebDriver::Remote::Capabilities.firefox
    caps.platform = :WINDOWS
    caps[:name] = "Watir WebDriver"

    browser = Watir::Browser.new(
      :remote,
      :url => "http://localhost:4444/wd/hub",
      :desired_capabilities => caps)

  else
    puts "Running Headless"
    headless_path=File.expand_path('../../driverDependencies/phantomjs.exe', File.dirname(__FILE__))
    profile=Selenium::WebDriver::PhantomJS.path = headless_path
    browser = Watir::Browser.new :phantomjs, :switches => %w[--disable-popup-blocking]    
  end

   browser.driver.manage.timeouts.implicit_wait = 5 # 5 seconds
   browser.driver.manage.window.maximize

#############################################

  loggedIn = false
  baseurl = ENV['URL']
  envPrefix = ENV['ENV_PREFIX']
  testData = YAML.load_file("features/support/data/#{envPrefix}_testdata.yml") if(!envPrefix.nil?)
  log = Logger.new STDOUT
  log.level = Logger::WARN
  log.level = Logger::DEBUG if(!ENV['DEBUG'].nil?)
  log.formatter = proc { |severity, time, progname, msg|"#{severity} #{time} #{caller[4].split('/').last}> #{msg}\n"}
  
  timestamp = ENV['RESULTSDIR'].split("_")[1]

#############################################################################################
    Before do 
      
      $baseurl = baseurl
      $browser = browser
      $loggedIn = loggedIn
      $envPrefix = envPrefix
      $testData = testData
      $log = log
      $timestamp = timestamp
      # puts "TimeStampPrefix is " + timestamp

      TestRun.browser.goto TestRun.baseurl if(!TestRun.loggedIn)

      

    end
###############################################################################################
    After do |scenario|
      begin
        if scenario.failed?
          Dir::mkdir('results/screenshots') if not File.directory?('results/screenshots')
          screenshot = "./results/screenshots/FAILED_#{scenario.name.gsub(' ','_').gsub(/[^0-9A-Za-z_]/, '')}.png"
          #Browser::BROWSER.driver.save_screenshot(screenshot)
          TestRun.browser.driver.save_screenshot(screenshot)
          embed screenshot, 'image/png'
        end
      #browser.windows.each {|w| w.close rescue nil} if(browser.windows > 1)
      ensure
        #@browser.quit
        loggedIn = TestRun.loggedIn
        # (TestRun.browser.windows.count..2).each { |windowNo|      
        #   TestRun.log.info "Closing Window: #{TestRun.browser.windows[(windowNo-1)].title}"
        #   TestRun.browser.windows[(windowNo-1)].close 
        # }  
        TestRun.browser.windows.count.downto(2) do |windowNo|
          TestRun.log.info "Current Window count = #{windowNo}"
          TestRun.log.info "Closing Window: #{TestRun.browser.windows[(windowNo-1)].title}"
          TestRun.browser.windows[(windowNo-1)].close
        end
      end
    end

at_exit do
  TestRun.browser.quit
end

