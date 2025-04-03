@echo off
rem Navigate to the folder Script.
cd /d "%~dp0"

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
echo Virtual environment activated.
pause
