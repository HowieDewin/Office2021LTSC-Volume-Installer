@echo off
:menu
	echo ====================================
	echo  	 Office 2021LTSC Downloader
	echo ====================================
	echo This is a downloader, for retreiving 
	echo the full office2021-LTSC files and 
	echo placing them in the proper directory
	echo to be used as a local master repo for
	echo performing multiple local deployments.
	echo ====================================
	echo Choose your version below:
	echo.
	echo 1. Standard
	echo 2. Pro
	echo 3. Exit
	echo.
	set /p choice="Enter your choice (1, 2, or 3): "

	if "%choice%"=="1" or "2" goto dl
	REM if "%choice%"=="2" goto dl
	if "%choice%"=="3" goto exit

	echo Invalid choice. Please select 1, 2, or 3.
	echo.
	goto menu

:dl
	echo Beginning download of Office-2021LTSC-Volume-STANDARD, this install takes upwards of 5-10 minutes, please be patient.

	REM Get the directory of the current script
	set "current_dir=%~dp0"

	REM Define the setup executable and configuration XML file
	set "setup_exe=setup.exe"
	if "%choice%"=="1" (
		set "config_xml=configuration-stn.xml"
		) else (
			set "config_xml=configuration-pro.xml"
			)
	REM Run the setup executable with the specified parameters
	"%current_dir%%setup_exe%" /download "%current_dir%%config_xml%"

	REM Echo error level indicating success or failure also state working directory upon exit then pause
	echo Error level: %errorlevel%
	if %errorlevel% neq 0 (
		echo Setup failed with error code %errorlevel%.
	) else (
		echo Setup completed successfully, copying master repository to proper directory.
		if "%choice%"=="1" (
			xcopy %~dp0downloader\Office %~dp0standard\
		) else(
			xcopy %~dp0downloader\Office %~dp0pro\		
			)
		)
		

	echo Current directory: %current_dir%
	pause

:exit
	echo Exiting the script. Have a great day!
	pause
	exit
