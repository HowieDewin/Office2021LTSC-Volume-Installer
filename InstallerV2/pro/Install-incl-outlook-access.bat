@echo off
:menu
echo ===================================
echo  Office 2021LTSC PRO Installer
echo ===================================
echo This installer does not include; 
echo Publisher, Teams, OneNote, OneDrive.
echo
echo It does include; 
echo Access, Outlook.
echo.
echo 1. Confirm and Begin Script
echo 2. Exit
echo.
set /p choice="Enter your choice (1 or 2): "

if "%choice%"=="1" goto begin
if "%choice%"=="2" goto exit

echo Invalid choice. Please select 1 or 2.
echo.
goto menu

:begin
echo Beginning install of Office-2021LTSC-Volume, this install takes upwards of 5-10 minutes, please be patient.

REM Get the directory of the current script
set "current_dir=%~dp0"

REM Define the setup executable and configuration XML file
set "setup_exe=setup.exe"
set "config_xml=pro-incl-outlook-access.xml"

REM Run the setup executable with the specified parameters
"%current_dir%%setup_exe%" /configure "%current_dir%%config_xml%"

REM Echo error level indicating success or failure also state working directory upon exit then pause
echo Error level: %errorlevel%
if %errorlevel% neq 0 (
    echo Setup failed with error code %errorlevel%.
) else (
    echo Setup completed successfully.
)

echo Current directory: %current_dir%
pause

:exit
echo Exiting the script. Have a great day!
pause
exit
