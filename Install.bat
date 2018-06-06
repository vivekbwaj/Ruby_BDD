Set StartInDirectory=%CD%

rem %StartInDirectory%

rem +----------------------------------+
rem 1. Installing Java
rem +----------------------------------+

rem Check current version

java -version

rem  Assumed already installed, if not then kill this script and run:
rem  jre-8u121-windows-x64.exe

rem  Enter to continue...

pause


rem +----------------------------------+
rem 2. Installing Ruby
rem +----------------------------------+

rubyinstaller-2.3.1-x64

rem 2.1 Verifying Ruby...

rem  %PATH%

set PATH=C:\Ruby23-x64\bin;%PATH%

rem  %PATH%

ruby -v

rem  Enter to continue...

pause


rem +----------------------------------+
rem 3. Installing Ruby DevKit
rem +----------------------------------+

rem 3.1 Creating directory...

md "C:\RubyDevKit"

copy DevKit-mingw64-64-4.7.2-20130224-1432-sfx.exe "C:\RubyDevKit"

C:

cd "C:\RubyDevKit"

dir

rem 3.2 Executing extractor...

DevKit-mingw64-64-4.7.2-20130224-1432-sfx

rem 3.3 Initialising installer...

ruby dk.rb init

dir config.yml

type config.yml

rem 3.4 Adding Ruby path

echo - C:\Ruby23-x64 >> config.yml

type config.yml

rem 3.4 Begin install...

ruby dk.rb install

rem  Enter to continue...

pause



rem +----------------------------------+
rem 4. Installing Ruby Gems
rem +----------------------------------+

rem 4.1 Installing specific gems

T:

cd %StartInDirectory%\Gems

gem install --force --local *.gem

rem  Enter to continue...

pause
