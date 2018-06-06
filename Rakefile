require 'rubygems'
require 'cucumber'
require 'cucumber/rake/task'

# Cucumber::Rake::Task.new(:features) do |t|
#   t.profile = 'default'
# end


# task :default => :features

task :ProductToPO => [:Product,:PO]do
	puts "Kick-off Product ->> PO"
end

task :Product do
	sh 'cucumber features\CreateProduct.feature -p debug -p Portal -p browser'
end

task :PO do
	sh 'cucumber features\CreatePO.feature -p debug -p Portal -p browser'
	# sh 'gem list'
end


task :default => :ProductToPO