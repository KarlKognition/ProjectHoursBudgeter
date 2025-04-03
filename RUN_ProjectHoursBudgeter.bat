@echo off
rem Navigate to the folder Script.
cd /d "%~dp0"
cd phb_app

rem Virtual environment
set "VENV_NAME=.venv"
if not exist "%VENV_NAME%" (
    echo Virtual environment not found. Creating...
    python -m venv "%VENV_NAME%"
    echo Virtual environment created.
) else (
    echo Virtual environment already exists. Using...
)
rem Activate the virtual environment
call %VENV_NAME%\Scripts\activate.bat
rem Check whether requirements.txt exists.
if not exist requirements.txt (
    echo.
    echo requirements.txt not found. Please ask the developer for their Python installation requirements. Exiting....
    deactivate
    echo.
    pause
    exit /b
)

echo Installing required packages...
pip install --upgrade -r requirements.txt

rem Check whether the packages are installed successfully
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo Could not install required packages.
    echo Please check your connection or installation rights. Exiting...
    echo.
    deactivate
    pause
    exit /b
)

rem return to the main folder to be able to start the module
cd ..
echo Start the majestic budgeting of hours!
python -m phb_app

rem Congratulate the user on the hopefully successful run of the wizard
echo.
echo ***********************************
echo ******** GLORIOUS SUCCESS! ********
echo ***********************************
echo ..... Or you've just cancelled the operation or worse yet crashed the program.
echo You may close this window at any time. Or not.... But then don't press any key here!
pause