@echo off
REM Navigate to the folder Script.
cd Script

echo "*************NOTE*************"
echo "TESTS CANNOT RUN IN BACKGROUND."
echo "Continuing means waiting for the tests to finish running before continuing other tasks."

choice /c yn /n /m "Press y to continue or n to exit."
if not %errorlevel% == 1 (
    exit /b
) else (
    rem Check whether requirements.txt exists.
    if not exist requirements.txt (
        echo "requirements.txt not found. Please ask the developer for their Python installation requirements. Exiting."
        exit /b
    )

    echo "Installing required packages..."
    pip install --user -r requirements.txt

    if exist __pycache__ (
        rmdir /s /q __pycache__
        echo "__pycache__ deleted."
    ) else (
        echo "No __pycache__ found."
    )

    echo "Looking for all TEST files."

    for %%f in (*unit_tests*.py) do (
        echo Running %%f
        python "%%f"
        echo.
    )

    echo "All found unit test files run."

    pause
)