@echo off

title CodeCracker By NotepadCodes!

color 5E

echo CodeCracher V1

echo              By NotepadCodes.webs.com

echo ******************************************************************

echo.

net user

echo type in a username in the option above:

Set /p username=

net user %username%

echo Changing Password...