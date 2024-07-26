@echo off
:: Checks if the script is being run with administrator permissions
openfiles >nul 2>&1
if %errorlevel% neq 0 (
    echo Please run this script as an administrator.
    echo Right-click on the file and select "Run as administrator"
    pause
    exit /b
)

:menu
cls
echo ===================================================
echo INTERACTIVE SYSTEM ACTIVATION AND MAINTENANCE MENU
echo MADE BY: OITALAO
echo ===================================================
echo 0. Hacker Mode
echo 1. Cleaning and Maintenance
echo 2. Activate Windows
echo 3. Activate Office
echo 4. Debloat Windows 11 
echo 5. Basic Programs
echo 6. System Performance Optimization
echo 7. Exit
echo =======================
set /p op="Choose an option: "

if %op%==0 goto hacker_mode
if %op%==1 goto cleaning_maintenance
if %op%==2 goto activate_windows
if %op%==3 goto activate_office
if %op%==4 goto debloat_windows11
if %op%==5 goto basic_programs
if %op%==6 goto performance_optimization
if %op%==7 goto exit

:hacker_mode
color 0a
goto menu

:cleaning_maintenance
echo Deleting temporary files...
del /q /f /s "%TEMP%\*.*"
del /q /f /s "C:\Windows\Temp\*.*"
del /q /f /s "C:\Windows\Prefetch\*.*"
echo Temporary file cleaning completed!

echo Running Disk Cleanup...
cleanmgr /sageset:1
cleanmgr /sagerun:1

echo Running CHKDSK...
chkdsk /f /r

echo Resetting IP configuration...
netsh int ip reset

echo Clearing DNS...
ipconfig /flushdns

echo Running SFC...
sfc /scannow

echo Running DISM - ScanHealth...
dism /online /cleanup-image /scanhealth

echo Running DISM - RestoreHealth (if necessary)...
dism /online /cleanup-image /restorehealth

echo Running Windows Malicious Software Removal Tool...
mrt /F /Q

echo Checking and installing Windows updates...
powershell -Command "Install-Module -Name PSWindowsUpdate -Force -SkipPublisherCheck"
powershell -Command "Import-Module PSWindowsUpdate; Get-WindowsUpdate -Install -AcceptAll -AutoReboot"

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
echo READ THE WARNINGS CAREFULLY!!!!!!!!!!!!!!!!!!!!
echo The Office package is an external application to Windows, so it must be installed through the link provided in option 3.
echo YOU MUST COMPLETE THE INSTALLATION PROCEDURE BEFORE ACTIVATION!
echo IF ONE DOES NOT WORK, TRY THE OTHER! Office tends to have issues with activation!
echo 1. Method 1
echo 2. Method 2
echo 3. I don't have Office
echo 4. Return to menu
echo =======================
set /p op3="Choose an option: "

if %op3%==1 goto activate_office_method1
if %op3%==2 goto activate_office_method2
if %op3%==3 goto no_office
if %op3%==4 goto menu

:activate_office_method1
echo Navigating to the Microsoft Office directory...
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
cscript //nologo ospp.vbs /act | find /i "successful" && (echo. ^& echo ============================================================================ ^& if errorlevel 2 exit) || (echo Connection to my KMS server failed! Trying to connect to another... & echo Please wait... & echo. & echo. & set /a i+=1 & goto server)
explorer "http://MSGuides.com"
goto halt
:notsupported
echo. ^
& echo ============================================================================ ^
& echo Sorry! There is no support for your Office version. ^
& echo Please try installing the latest available: bit.ly/aiomsp
:halt
pause >nul
goto menu

:activate_office_method2
echo Navigating to the Microsoft Office directory (Program Files x86)...
cd /d "%ProgramFiles(x86)%\Microsoft Office\Office16"

echo Navigating to the Microsoft Office directory (Program Files)...
cd /d "%ProgramFiles%\Microsoft Office\Office16"

echo Installing KMS licenses...
for /f %%x in ('dir /b ..\root\Licenses16\ProPlus2021VL_KMS*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x"

echo Configuring KMS port...
cscript ospp.vbs /setprt:1688

echo Removing product key...
cscript ospp.vbs /unpkey:6F7TH >nul

echo Inserting new product key...
cscript ospp.vbs /inpkey:FXYTK-NJJ8C-GB6DW-3DYQT-6F7TH

echo Configuring KMS host...
cscript ospp.vbs /sethst:e8.us.to

echo Activating Office...
cscript ospp.vbs /act

echo Process completed.
pause
goto menu

:no_office
echo Opening Office site...
start https://msguides.com/download-microsoft-office-windows-os
goto menu

:debloat_windows11
@echo off
:: Set execution policy to Unrestricted
powershell -Command "Set-ExecutionPolicy Unrestricted -Force"

:: Run the Windows 11 debloat script
powershell -Command "& {Invoke-Expression (Invoke-RestMethod 'https://raw.githubusercontent.com/Raphire/Win11Debloat/master/Get.ps1')}"

pause
goto menu

:basic_programs
echo Installing Google Chrome...
winget install -e --id Google.Chrome

echo Installing VLC Media Player...
winget install -e --id VideoLAN.VLC

echo Installing Java Runtime Environment...
winget install -e --id Oracle.JavaRuntimeEnvironment

echo Installing aTube Catcher...
winget install -e --id DsNET.atube-catcher

echo Installing 7-Zip...
winget install -e --id 7zip.7zip

echo Installing qBittorrent...
winget install -e --id qBittorrent.qBittorrent

echo Installing Adobe Acrobat Reader DC...
winget install -e --id Adobe.Acrobat.Reader.64-bit

echo All basic programs have been installed!
pause
goto menu

:performance_optimization
cls
echo =======================
echo SYSTEM PERFORMANCE OPTIMIZATION
echo =======================
echo 1. Disable unnecessary services
echo 2. Adjust power settings
echo 3. Enable maximum performance mode
echo 4. Return to menu
echo =======================
set /p op4="Choose an option: "

if %op4%==1 goto disable_services
if %op4%==2 goto adjust_power
if %op4%==3 goto max_performance
if %op4%==4 goto menu

:disable_services
echo Disabling unnecessary services...
sc config "DiagTrack" start= disabled
sc config "dmwappushservice" start= disabled
sc config "TrkWks" start= disabled
sc config "WMPNetworkSvc" start= disabled
sc stop "DiagTrack"
sc stop "dmwappushservice"
sc stop "TrkWks"
sc stop "WMPNetworkSvc"
echo Services disabled!
pause
goto performance_optimization

:adjust_power
echo Adjusting power settings for high performance...
powercfg /change standby-timeout-ac 0
powercfg /change hibernate-timeout-ac 0
powercfg /change monitor-timeout-ac 0
powercfg /change disk-timeout-ac 0
powercfg /setactive SCHEME_MIN
echo Power settings adjusted!
pause
goto performance_optimization

:max_performance
echo Enabling maximum performance mode...
powercfg /duplicatescheme e9a42b02-d5df-448d-aa00-03f14749eb61
powercfg /setactive e9a42b02-d5df-448d-aa00-03f14749eb61
echo Maximum performance mode enabled!
pause
goto performance_optimization

:exit
echo Exiting in 1 second...
timeout /t 1 /nobreak >nul
exit
