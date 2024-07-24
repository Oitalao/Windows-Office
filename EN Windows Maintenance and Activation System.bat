@echo off
:: Checks if the script is being run with administrator permissions
openfiles >nul 2>&1
if %errorlevel% neq 0 (
    echo Please run this script as administrator.
    echo Right-click the file and select "Run as administrator"
    pause
    exit /b
)

:menu
cls
echo ===================================================
echo INTERACTIVE SYSTEM MAINTENANCE AND ACTIVATION MENU
echo MADE BY: OITALAO
echo ===================================================
echo 0. Hacker Mode
echo 1. Cleaning and Maintenance
echo 2. Activate Windows
echo 3. Activate Office
echo 4. Debloat Windows 11 
echo 5. Exit
echo =======================
set /p op="Choose an option: "

if %op%==0 goto hacker_mode
if %op%==1 goto cleaning_and_maintenance
if %op%==2 goto activate_windows
if %op%==3 goto activate_office
if %op%==4 goto debloat_windows11
if %op%==5 goto exit

:hacker_mode
color 0a
goto menu

:cleaning_and_maintenance
echo Deleting temporary files...
del /q /f /s "%TEMP%\*.*"
del /q /f /s "C:\Windows\Temp\*.*"
del /q /f /s "C:\Windows\Prefetch\*.*"
echo Temporary files cleanup completed!

echo Running Disk Cleanup...
cleanmgr /sageset:1
cleanmgr /sagerun:1

echo Running CHKDSK...
chkdsk /f /r

echo Resetting IP configuration...
netsh int ip reset

echo Clearing DNS cache...
ipconfig /flushdns

echo Running SFC...
sfc /scannow

echo Running DISM - ScanHealth...
dism /online /cleanup-image /scanhealth

echo Running DISM - RestoreHealth (if needed)...
dism /online /cleanup-image /restorehealth

echo Updating all Winget packages...
winget upgrade --all

echo Registering AppX packages for all users...
powershell -Command "Get-AppXPackage -AllUsers | Where-Object { $_.InstallLocation -like '*SystemApps*' } | Foreach { Add-AppxPackage -DisableDevelopmentMode -Register '$($_.InstallLocation)\AppXManifest.xml' }"

echo All tasks have been executed!
pause
goto menu

:activate_windows
cls
echo =======================
echo ACTIVATE WINDOWS
echo =======================
echo 1. Windows 10
echo 2. Windows 11
echo 3. Return to menu
echo =======================
set /p op2="Choose an option: "

if %op2%==1 goto activate_windows10
if %op2%==2 goto activate_windows11
if %op2%==3 goto menu

:activate_windows10
echo Activating Windows 10...
slmgr /ipk TX9XD-98N7V-6WMQ6-BX7FG-H8Q99
slmgr /skms kms8.msguides.com
slmgr /ato
echo Windows 10 activated!
pause
goto activate_windows

:activate_windows11
echo Activating Windows 11...
slmgr /ipk W269N-WFGWX-YVC9B-4J6C9-T83GX
slmgr /skms kms8.msguides.com
slmgr /ato
echo Windows 11 activated!
pause
goto activate_windows

:activate_office
cls
echo =======================
echo ACTIVATE OFFICE
echo =======================
echo PLEASE READ THE WARNINGS CAREFULLY!!!!!!!!!!!!!!!!!!!!
echo The Office package is an external application to Windows, so it must be installed from the link provided in option 3.
echo YOU MUST COMPLETE THE INSTALLATION PROCEDURE BEFORE ACTIVATION!
echo IF ONE DOESN'T WORK, TRY THE OTHER! Office tends to have activation issues!
echo 1. Method 1
echo 2. Method 2
echo 3. I don't have Office
echo 4. Return to menu
echo =======================
set /p op3="Choose an option: "

if %op3%==1 goto activate_office_method1
if %op3%==2 goto activate_office_method2
if %op3%==3 goto dont_have_office
if %op3%==4 goto menu

:activate_office_method1
echo Navigating to Microsoft Office directory...
& (if exist "%ProgramFiles%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles%\Microsoft Office\Office16") ^
&(if exist "%ProgramFiles(x86)%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles(x86)%\Microsoft Office\Office16") ^
&(for /f %%x in ('dir /b ..\root\Licenses16\ProPlus2019VL*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul) ^
&(for /f %%x in ('dir /b ..\root\Licenses16\ProPlus2019VL*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul) ^
& echo. ^
& echo ============================================================================ ^
& echo Activating your Office... ^
& cscript //nologo slmgr.vbs /ckms >nul ^
& cscript //nologo ospp.vbs /setprt:1688 >nul ^
& cscript //nologo ospp.vbs /unpkey:6MWKP >nul ^
& cscript //nologo ospp.vbs /inpkey:NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP >nul ^
& set i=1
:server
if %i%==1 set KMS_Sev=kms7.MSGuides.com
if %i%==2 set KMS_Sev=kms8.MSGuides.com
if %i%==3 set KMS_Sev=kms9.MSGuides.com
if %i%==4 goto notsupported
cscript //nologo ospp.vbs /sethst:%KMS_Sev% >nul ^
& echo ============================================================================ ^
& echo. ^
& echo.
cscript //nologo ospp.vbs /act | find /i "successful" && (echo. ^& echo ============================================================================ ^& if errorlevel 2 exit) || (echo Failed to connect to my KMS server! Trying to connect to another one... & echo Please wait... & echo. & echo. & set /a i+=1 & goto server)
explorer "http://MSGuides.com"
goto halt
:notsupported
echo. ^
& echo ============================================================================ ^
& echo Sorry! Your Office version is not supported. ^
& echo Please try installing the latest available: bit.ly/aiomsp
:halt
pause >nul
goto menu

:activate_office_method2
echo Navigating to Microsoft Office directory (Program Files x86)...
cd /d "%ProgramFiles(x86)%\Microsoft Office\Office16"

echo Navigating to Microsoft Office directory (Program Files)...
cd /d "%ProgramFiles%\Microsoft Office\Office16"

echo Installing KMS licenses...
for /f %%x in ('dir /b ..\root\Licenses16\ProPlus2021VL_KMS*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x"

echo Setting KMS port...
cscript ospp.vbs /setprt:1688

echo Removing product key...
cscript ospp.vbs /unpkey:6F7TH >nul

echo Entering new product key...
cscript ospp.vbs /inpkey:FXYTK-NJJ8C-GB6DW-3DYQT-6F7TH

echo Setting KMS host...
cscript ospp.vbs /sethst:e8.us.to

echo Activating Office...
cscript ospp.vbs /act

echo Process completed.
pause
goto menu

:dont_have_office
echo Opening Office website...
start https://msguides.com/download-microsoft-office-windows-os
goto menu

:debloat_windows11
@echo off
:: Set the execution policy to Unrestricted
powershell -Command "Set-ExecutionPolicy Unrestricted -Force"

:: Run the Windows 11 debloat script
powershell -Command "& {Invoke-Expression (Invoke-RestMethod 'https://raw.githubusercontent.com/Raphire/Win11Debloat/master/Get.ps1')}"

pause
goto menu

:exit
echo Exiting in 1 second...
timeout /t 1 /nobreak >nul
exit
