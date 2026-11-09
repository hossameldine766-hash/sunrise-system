@echo off
chcp 65001 >nul
title Build Sunrise Sales System
color 0A

cls
echo ============================================================
echo              Build Sunrise Sales System
echo ============================================================
echo.

REM ═══════════════════════════════════════════════════════════
REM [0] التأكد أن السكربت يعمل من مجلد المشروع
REM ═══════════════════════════════════════════════════════════
cd /d "%~dp0"
echo [0/5] Current folder: %cd%
echo.

if not exist "bazar1.py" (
    echo [ERROR] bazar1.py not found in this folder!
    echo.
    echo Current folder: %cd%
    echo.
    pause
    exit /b 1
)
echo   [OK] bazar1.py found
echo.

REM ═══════════════════════════════════════════════════════════
REM [1] التحقق من وجود .venv
REM ═══════════════════════════════════════════════════════════
echo [1/5] Checking .venv folder...

if not exist ".venv" (
    echo [ERROR] .venv folder NOT found!
    echo.
    echo Please create it first with:
    echo   py -m venv .venv
    echo.
    pause
    exit /b 1
)

if not exist ".venv\Scripts\activate.bat" (
    echo [ERROR] .venv\Scripts\activate.bat NOT found!
    echo.
    pause
    exit /b 1
)

echo   [OK] .venv found
echo.

REM ═══════════════════════════════════════════════════════════
REM [2] تنظيف الملفات القديمة
REM ═══════════════════════════════════════════════════════════
echo [2/5] Cleaning old build files...

if exist "dist" (
    echo   - Removing dist...
    rmdir /s /q "dist" >nul 2>&1
)
if exist "build" (
    echo   - Removing build...
    rmdir /s /q "build" >nul 2>&1
)
if exist "SunriseSystem.spec" (
    echo   - Removing SunriseSystem.spec...
    del /q "SunriseSystem.spec" >nul 2>&1
)

echo   [OK] Cleaned
echo.

REM ═══════════════════════════════════════════════════════════
REM [3] تفعيل البيئة الافتراضية
REM ═══════════════════════════════════════════════════════════
echo [3/5] Activating virtual environment...

call ".venv\Scripts\activate.bat"

if errorlevel 1 (
    echo [ERROR] Failed to activate .venv!
    pause
    exit /b 1
)

echo   [OK] Activated
echo.

REM ═══════════════════════════════════════════════════════════
REM [4] التحقق من PyInstaller
REM ═══════════════════════════════════════════════════════════
echo [4/5] Checking PyInstaller installation...

python -c "import PyInstaller" >nul 2>&1

if errorlevel 1 (
    echo [ERROR] PyInstaller is NOT installed in .venv!
    echo.
    echo Please install it first with:
    echo   .venv\Scripts\activate.bat
    echo   pip install pyinstaller
    echo.
    pause
    exit /b 1
)

echo   [OK] PyInstaller found
echo.

REM ═══════════════════════════════════════════════════════════
REM [5] تشغيل البناء
REM ═══════════════════════════════════════════════════════════
echo [5/5] Building EXE...
echo.
echo Please wait 3-7 minutes. DO NOT close this window.
echo.

python -m PyInstaller --onefile --windowed ^
    --icon=app.ico ^
    --name="SunriseSystem" ^
     --splash splash.png ^
    --add-data "m.png;." ^
    --add-data "app.ico;." ^
    --add-data "splash.png;." ^
    --hidden-import=tkcalendar ^
    --hidden-import=PIL._tkinter_finder ^
    --hidden-import=bidi ^
    --hidden-import=arabic_reshaper ^
    --hidden-import=openpyxl ^
    --hidden-import=openpyxl.drawing ^
    --hidden-import=openpyxl.drawing.image ^
    --hidden-import=openpyxl.styles ^
    --clean ^
    bazar1.py

if errorlevel 1 (
    echo.
    echo ============================================================
    echo  [ERROR] Build FAILED!
    echo ============================================================
    echo.
    echo Check the error messages above.
    pause
    exit /b 1
)

echo.
echo ============================================================
echo  [SUCCESS] Build completed!
echo.
echo  File: dist\SunriseSystem.exe
echo ============================================================
echo.

if exist "dist" (
    echo Opening dist folder...
    start "" "dist"
)

echo.
echo Done! Press any key to close...
pause >nul