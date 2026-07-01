@echo off
echo ============================================
echo  Junior Supervisor Bill Generator - Setup
echo ============================================
echo.

echo [1/3] Installing Node.js dependencies...
cd backend
call npm install
if %errorlevel% neq 0 (
    echo ERROR: npm install failed. Make sure Node.js is installed.
    pause
    exit /b 1
)
echo Done.
echo.

echo [2/3] Installing Python PDF merge library...
pip install pypdf
if %errorlevel% neq 0 (
    echo WARNING: pip install pypdf failed.
    echo Make sure Python is installed and added to PATH.
    echo Download from: https://www.python.org/downloads/
)
echo Done.
echo.

echo [3/3] Checking template file...
if not exist "templates\Junior_Supervisor.xlsx" (
    echo ERROR: Template file missing!
    echo Please copy Junior_Supervisor.xlsx into the backend\templates\ folder.
    pause
    exit /b 1
)
echo Template found.
echo.

echo ============================================
echo  Setup complete! Starting server...
echo  Open http://localhost:4000 in your browser
echo ============================================
echo.
node server.js
pause
