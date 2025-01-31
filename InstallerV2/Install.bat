@echo off

setlocal enabledelayedexpansion

echo Checking for local Office2021-LTSC install repositories...
REM Check if the pro\Office directory exists
if not exist "%~dp0pro\Office" (
  echo "The directory 'pro\Office' is missing."
  set "missing=pro\Office"
)

REM Check if the standard\Office directory exists
if not exist "%~dp0standard\Office" (
  echo "The directory 'standard\Office' is missing."
  if defined missing (
    set "missing=!missing!, standard\Office"
  ) else (
    set "missing=standard\Office"
  )
)

REM If both directories are missing, prompt the user for download
if defined missing (
  echo "The following directories are missing: %missing%"
  echo "Do you want to download the missing directories? (Y/N)"
  set /p choice=
  if /i "%choice%"=="Y" (
    call "%~dp0downloader\downloader.bat"
  ) else (
    echo "Download canceled."
    exit /b
  )
)


echo "Both directories exist. Proceeding to menu..."
timeout /t 3 /nobreak

:menu
cls
echo ====================================
echo   Office 2021LTSC Master Installer
echo ====================================
echo Double-check that you made sure to 
echo remove existing installs of office
echo before continuing or this will fail!
echo ====================================
echo First, select which version you need:
echo.
echo 1. Standard
echo 2. Pro
echo 3. Exit Program
echo.
set /p choice="Enter your choice (1, 2, or 3): "

if "%choice%"=="1" goto v-standard
if "%choice%"=="2" goto v-pro
if "%choice%"=="3" goto exit

echo Invalid choice. Please select 1, 2, or 3.
echo.
goto menu

:v-standard

	:stn-menu
		cls
		echo ====================================
		echo  Office 2021LTSC STANDARD Installer
		echo ====================================
		echo This is a STANDARD installer, NOT PRO. 
		echo It includes only Word, Excel, Powerpoint. 
		echo (And outlook but thats not included here)
		echo.
		echo 1. Confirm and Begin Script
		echo 2. Go Back
		echo 3. Exit Program
		echo.
		set /p choice="Enter your choice (1, 2, or 3): "

		if "%choice%"=="1" goto begin-stn
		if "%choice%"=="2" goto menu
		if "%choice%"=="3" goto exit

		echo Invalid choice. Please select 1, 2, or 3.
		echo.
		goto stn-menu

	:begin-stn
		cls
		echo Beginning install of Office-2021LTSC-Standard-Volume, this install takes upwards of 5-10 minutes, please be patient.
		REM Get the directory of the current script
		set "current_dir=%~dp0\standard\"
		REM Define the setup executable and configuration XML file
		set "setup_exe=setup.exe"
		set "config_xml=stn-minimal.xml"
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

:v-pro

	:promenu
		cls
		echo =====================================
		echo  Office 2021LTSC PRO-MASTER Installer
		echo =====================================
		echo Double-check that you made sure to 
		echo remove any pre-existing installs of office
		echo before continuing or this will fail!
		echo
		echo First, select which version you need:
		echo.
		echo 1. pro-minimal
		echo 2. pro-incl-access
		echo 3. pro-incl-outlook
		echo 4. pro-incl-outlook-access
		echo 5. Go Back
		echo 6. Exit Program
		echo.
		set /p choice="Enter your choice (1, 2, 3, 4, 5, or 6): "

		if "%choice%"=="1" goto pro-min
		if "%choice%"=="2" goto pro-access
		if "%choice%"=="3" goto pro-outlook
		if "%choice%"=="4" goto pro-outlook-access
		if "%choice%"=="5" goto menu
		if "%choice%"=="6" goto exit
        	echo Invalid choice. Please select 1, 2, 3, 4, 5, or 6.
		echo.
		goto promenu

	:pro-min
		REM Get the directory of the current script
		set "current_dir=%~dp0\pro\"
		REM Define the setup executable and configuration XML file
		set "setup_exe=setup.exe"
		set "config_xml=pro-minimal.xml"
		REM Run the setup executable with the specified parameters
		"%current_dir%%setup_exe%" /configure "%current_dir%%config_xml%"
		goto errorlvl

	:pro-access
		REM Get the directory of the current script
		set "current_dir=%~dp0\pro\"
		REM Define the setup executable and configuration XML file
		set "setup_exe=setup.exe"
		set "config_xml=pro-incl-access.xml"
		REM Run the setup executable with the specified parameters
		"%current_dir%%setup_exe%" /configure "%current_dir%%config_xml%"
		goto errorlvl

	:pro-outlook
		REM Get the directory of the current script
		set "current_dir=%~dp0\pro\"
		REM Define the setup executable and configuration XML file
		set "setup_exe=setup.exe"
		set "config_xml=pro-incl-outlook.xml"
		REM Run the setup executable with the specified parameters
		"%current_dir%%setup_exe%" /configure "%current_dir%%config_xml%"
		goto errorlvl

	:pro-outlook-access
		REM Get the directory of the current script
		set "current_dir=%~dp0\pro\"
		REM Define the setup executable and configuration XML file
		set "setup_exe=setup.exe"
		set "config_xml=pro-incl-outlook-access.xml"
		REM Run the setup executable with the specified parameters
		"%current_dir%%setup_exe%" /configure "%current_dir%%config_xml%"
		goto errorlvl

	:errorlvl
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
