@echo off
echo Setting up Email Automation Desktop...

REM Check if venv exists
if not exist "venv" (
    echo Creating virtual environment...
    python -m venv venv
)

REM Activate virtual environment
echo Activating virtual environment...
call venv\Scripts\activate.bat

REM Install dependencies
echo Installing dependencies...
pip install -r requirements.txt

echo.
echo Setup complete! Now you can run:
echo.
echo 1. Test the app: python test_app.py
echo 2. Run the app: python src/main.py
echo.
pause