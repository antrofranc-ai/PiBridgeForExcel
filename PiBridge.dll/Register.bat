@echo off
::set global variables
setlocal EnableDelayedExpansion
set APPNAME=PiBridge
set LOGFILE=%USERPROFILE%\Documents\%APPNAME%_Install.log
set DLLPATH=%~dp0
set WORDIR=%~dp0
set BATFILE=%~0
set EXCELBIT=1
set OSBIT=1
set ARCHITECTURE=1
set INSTALLDIR=C:\Users\Public\Documents
set EXCELBUILD=11.0
set EXCELVERSION=2003
set NETFRAME=%WINDIR%\Microsoft.NET\Framework\v4.0*
set PIDIR=C:\Zerodha\Pi
set PIPATH=C:\Zerodha\Pi\Pi.exe

set MSG=#START#
echo %MSG% >>"%LOGFILE%"

set MSG=Batch file started at %DATE% %TIME%
echo %MSG% && echo %MSG% >>"%LOGFILE%"

set MSG=Registering PiBridge DLL using regasm
echo %MSG% && echo %MSG% >>"%LOGFILE%"

set MSG=APP NAME : %APPNAME%
echo %MSG% && echo %MSG% >>"%LOGFILE%"

set MSG=LOG FILE : %LOGFILE%
echo %MSG% && echo %MSG% >>"%LOGFILE%"

set MSG=BAT FILE : %BATFILE%
echo %MSG% >>"%LOGFILE%"

set MSG=WORKING DIRECTORY : %WORDIR%
echo %MSG% >>"%LOGFILE%"

set MSG=DLL PATH : %DLLPATH%
echo %MSG% && echo %MSG% >>"%LOGFILE%"

set MSG=PI PATH : %PIPATH%
echo %MSG% && echo %MSG% >>"%LOGFILE%"

set MSG=PI DIR : %PIDIR%
echo %MSG% && echo %MSG% >>"%LOGFILE%"

echo. && echo.
set MSG=*** User Consent to run this batch file ***
echo %MSG% && echo %MSG% >>"%LOGFILE%"

set MSG=This batch file will register PiBridge DLL for Excel. Official installer only supports for AmiBroker. Since the PiBridge is COM, any windows application can consume it. Note PiBridge is 32 bit, so your Excel should be 32 bit. **PiBridge will not work in 64bit MS Office**.
echo %MSG% && echo %MSG% >>"%LOGFILE%"

echo. && echo.
set MSG=*** Actually what this batch file does? ***
echo %MSG% && echo %MSG% >>"%LOGFILE%"

set MSG=Since this file can be viewed in Notepad, You can open it yourself and see what it does.
echo %MSG% && echo %MSG% >>"%LOGFILE%"

set MSG=For others, here is what it does...
echo %MSG% && echo %MSG% >>"%LOGFILE%"

set MSG=* Checks for Administrator rights
echo %MSG% && echo %MSG% >>"%LOGFILE%"

set MSG=* Checks for OS Architecture and Registering DLL accordingly
echo %MSG% && echo %MSG% >>"%LOGFILE%"


echo. && echo.
echo Do you want to proceed [Type Y/N and Press Enter]? >>"%LOGFILE%"
set /P INPUT=Do you want to proceed [Type Y/N and Press Enter]?
echo %INPUT% && echo %INPUT% >>"%LOGFILE%"

if /I "%INPUT%" equ "N" (
set ERR=%APPNAME% DLL is not registered. Exiting Command Prompt.
echo !ERR! && echo !ERR! >>"%LOGFILE%"
echo x=msgbox^("!ERR!" ,48, "%APPNAME% Error"^) > %TEMP%\msgbox.vbs && start /w %TEMP%\msgbox.vbs
exit
) 

if NOT "%DLLPATH%"=="%DLLPATH:(=%" (
set ERR=Your folder contains special characters '^('.Please rename the folder without special character and try again.
echo !ERR! && echo !ERR! >>"%LOGFILE%"
echo x=msgbox^("!ERR!" ,48, "%APPNAME% Error"^) > %TEMP%\msgbox.vbs && start /w %TEMP%\msgbox.vbs
exit
)

if NOT "%DLLPATH%"=="%DLLPATH:)=%" (
set ERR=Your folder contains special characters '^)'.Please rename the folder without special character and try again.
echo !ERR! && echo !ERR! >>"%LOGFILE%"
echo x=msgbox^("!ERR!" ,48, "%APPNAME% Error"^) > %TEMP%\msgbox.vbs && start /w %TEMP%\msgbox.vbs
exit
)

::::::::ADMIN
echo. && echo.
set MSG=Checking for Administrative rights
echo %MSG% && echo %MSG% >>"%LOGFILE%"

net session >nul 2>&1
if NOT %errorLevel% == 0 (
set ERR=You do not have Administrative rights. I assume you are running this file without extracting it. You can try as follows. Extract the downloaded Zip file contents to a folder. Right click on the Register.bat file and select Run As Administrator
echo !ERR! && echo !ERR! >>"%LOGFILE%"
echo x=msgbox^("!ERR!" ,16, "%APPNAME% Error"^) > %TEMP%\msgbox.vbs && start /w %TEMP%\msgbox.vbs
exit
) 

set MSG=ADMIN RIGHTS : YES
echo %MSG% && echo %MSG% >>"%LOGFILE%"

echo. && echo.
set MSG=Checking Pi Installation
echo %MSG% && echo %MSG% >>"%LOGFILE%"

if NOT Exist %PIPATH% (
set ERR=Pi is not installed.
echo !ERR! && echo !ERR! >>"%LOGFILE%"
echo x=msgbox^("!ERR!" ,48, "%APPNAME% Error"^) > %TEMP%\msgbox.vbs && start /w %TEMP%\msgbox.vbs
exit
)

set MSG=PI INSTALLED : YES
echo %MSG% && echo %MSG% >>"%LOGFILE%"

echo. && echo.
set MSG=Checking whether Excel is running or not
echo %MSG% && echo %MSG% >>"%LOGFILE%"

tasklist /FI "IMAGENAME eq EXCEL.exe" 2>NUL | find /I /N "EXCEL.exe">NUL
if "%ERRORLEVEL%"=="0" (
set ERR=Excel is running. Please close Excel and try again.
echo !ERR! && echo !ERR! >>"%LOGFILE%"
echo x=msgbox^("!ERR!" ,48, "%APPNAME% Error"^) > %TEMP%\msgbox.vbs && start /w %TEMP%\msgbox.vbs
exit
)

set MSG=EXCEL RUNNING : NO
echo %MSG% && echo %MSG% >>"%LOGFILE%"

echo. && echo.
set MSG=Checking whether Pi is running or not
echo %MSG% && echo %MSG% >>"%LOGFILE%"

tasklist /FI "IMAGENAME eq Pi.exe" 2>NUL | find /I /N "Pi.exe">NUL
if "%ERRORLEVEL%"=="0" (
set ERR=Pi is running. Please close Pi and try again.
echo !ERR! && echo !ERR! >>"%LOGFILE%"
echo x=msgbox^("!ERR!" ,48, "%APPNAME% Error"^) > %TEMP%\msgbox.vbs && start /w %TEMP%\msgbox.vbs
exit
)

set MSG=PI RUNNING : NO
echo %MSG% && echo %MSG% >>"%LOGFILE%"

echo. && echo.
set MSG=Checking whether AmiBroker is running or not
echo %MSG% && echo %MSG% >>"%LOGFILE%"

tasklist /FI "IMAGENAME eq Broker.exe" 2>NUL | find /I /N "Broker.exe">NUL
if "%ERRORLEVEL%"=="0" (
set ERR=AmiBroker is running. Please close AmiBroker and try again.
echo !ERR! && echo !ERR! >>"%LOGFILE%"
echo x=msgbox^("!ERR!" ,48, "%APPNAME% Error"^) > %TEMP%\msgbox.vbs && start /w %TEMP%\msgbox.vbs
exit
)

set MSG=AMIBROKER RUNNING : NO
echo %MSG% && echo %MSG% >>"%LOGFILE%"

::::::::OS
echo. && echo.
set MSG=Checking for OS Architecture
echo %MSG% && echo %MSG% >>"%LOGFILE%"

set MSG=PROCESSOR_ARCHITECTURE : %PROCESSOR_ARCHITECTURE%
echo %MSG% && echo %MSG% >>"%LOGFILE%"

reg query "HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Environment" /v PROCESSOR_ARCHITECTURE  | find "AMD64" >nul 2>&1
if %errorlevel% equ 0 (
set OSBIT=64
set ARCHITECTURE=AMD64
) else (
	reg query "HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Environment" /v PROCESSOR_ARCHITECTURE  | find "x86" >nul 2>&1
	if !errorlevel! equ 0 (
	set OSBIT=32
	set ARCHITECTURE=X86
	) 
)

if %OSBIT% == 1 (
set ERR=Unable to find Operating System Architecture. Please try again. If problem persits contact Administrator.
echo !ERR! && echo !ERR! >>"%LOGFILE%"
echo x=msgbox^("!ERR!" ,16, "%APPNAME% Error"^) > %TEMP%\msgbox.vbs && start /w %TEMP%\msgbox.vbs
exit
) 

set MSG=CPU ARCHITECTURE : %ARCHITECTURE%
echo %MSG% && echo %MSG% >>"%LOGFILE%"
set MSG=OS BIT : %OSBIT%
echo %MSG% && echo %MSG% >>"%LOGFILE%"
echo. && echo.


set MSG=Setting System Directory as per OS Bit
echo %MSG% && echo %MSG% >>"%LOGFILE%"

if %OSBIT% == 32 (
set INSTALLDIR=C:\Windows\System32
) else (
set INSTALLDIR=C:\Windows\SysWOW64
)

set MSG=Checking .Net Framework Installation
echo %MSG% && echo %MSG% >>"%LOGFILE%"

if not exist %NETFRAME% (
set ERR=Unable to find .Net Framework installation. Please try again. If problem persits contact Administrator.
echo !ERR! && echo !ERR! >>"%LOGFILE%"
echo x=msgbox^("!ERR!" ,16, "%APPNAME% Error"^) > %TEMP%\msgbox.vbs && start /w %TEMP%\msgbox.vbs
exit
)

set MSG=DOTNET FRAME : %NETFRAME%
echo %MSG% && echo %MSG% >>"%LOGFILE%"

set MSG=Copying Files to %INSTALLDIR%
echo %MSG% && echo %MSG% >>"%LOGFILE%"

xcopy /s /h /y /r "%DLLPATH%%APPNAME%.dll" "%INSTALLDIR%" >nul 2>&1
if NOT %errorlevel% equ 0 (
set ERR=Unable to copy file %DLLPATH%%APPNAME%.dll : You are running the batch file without extracting or File may be missing or You do not have rights to copy
echo !ERR! && echo !ERR! >>"%LOGFILE%"
echo x=msgbox^("!ERR!" ,16, "%APPNAME% Error"^) > %TEMP%\msgbox.vbs && start /w %TEMP%\msgbox.vbs
exit
)

xcopy /s /h /y /r "%DLLPATH%client.ini" "%PIDIR%" >nul 2>&1
if NOT %errorlevel% equ 0 (
set ERR=Unable to copy file %DLLPATH%client.ini : You are running the batch file without extracting or File may be missing or You do not have rights to copy
echo !ERR! && echo !ERR! >>"%LOGFILE%"
echo x=msgbox^("!ERR!" ,16, "%APPNAME% Error"^) > %TEMP%\msgbox.vbs && start /w %TEMP%\msgbox.vbs
exit
)

set MSG=Successfully copied files
echo %MSG% && echo %MSG% >>"%LOGFILE%"

set MSG=Registering DLL files
echo %MSG% && echo %MSG% >>"%LOGFILE%"

CD %NETFRAME%
::Unregister first if already installed
Regasm "%INSTALLDIR%\%APPNAME%.dll" /codebase /u >nul 2>&1
if NOT %errorlevel% equ 0 (
set ERR=Unable to unregister. Don't worry, May be this is your first installation
echo !ERR! && echo !ERR! >>"%LOGFILE%"
)

Regasm "%INSTALLDIR%\%APPNAME%.dll" /codebase  >nul 2>&1
if NOT %errorlevel% equ 0 (
set ERR=Unable to Register DLL. Please try again. If problem persits contact Administrator.
echo !ERR! && echo !ERR! >>"%LOGFILE%"
echo x=msgbox^("!ERR!" ,16, "%APPNAME% Error"^) > %TEMP%\msgbox.vbs && start /w %TEMP%\msgbox.vbs
exit
)

::::::::END
echo. && echo.
set MSG=Successfully registered %APPNAME% DLL in the system.Now You can use PiBridge with Excel.
echo %MSG% && echo %MSG% >>"%LOGFILE%"
echo x=msgbox^("Successfully registered %APPNAME% DLL in the system." ^& vbCrLf ^& "Now You can use PiBridge with Excel" ,64, "%APPNAME% Success"^) > %TEMP%\msgbox.vbs && start /w %TEMP%\msgbox.vbs
::Pause >nul
