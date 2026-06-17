@echo off
echo ============================================
echo   Build EXE - Drug Concentration Tool
echo ============================================
echo.

echo [1/4] Creating virtual environment...
python -m venv venv
if errorlevel 1 (
    echo.
    echo ERROR: Python not found.
    echo Please install Python and check "Add Python to PATH".
    pause
    exit /b
)

call venv\Scripts\activate.bat

echo [2/4] Installing packages...
python -m pip install --upgrade pip
pip install -r requirements.txt

echo [3/4] Building exe...
pyinstaller --onefile --windowed --name DrugConcCalculator main.py

echo.
echo [4/4] Done!
echo Your exe is here:  dist\DrugConcCalculator.exe
echo ============================================
pause
