require 'chronic'
require 'time'
require 'colorize'


Given(/^test1 that the framework starts$/) do
	
	LoadURL("http://rubygems.org")
	sleep 2

end