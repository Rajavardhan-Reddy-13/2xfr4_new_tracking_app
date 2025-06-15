@echo off
setlocal enabledelayedexpansion
cd /d "%~dp0"

REM ===== CONFIG =====
set APP_NAME=AllTestsGUI
set MAIN_PY=Summary.py
set VENV_DIR=venv

echo ==========================================
echo Building ONE-FILE EXE: %APP_NAME%
echo Main: %MAIN_PY%
echo Folder: %cd%
echo ==========================================

REM --- sanity check ---
if not exist "%MAIN_PY%" (
  echo [ERROR] Cannot find "%MAIN_PY%" in this folder.
  dir /b *.py
  pause
  exit /b 1
)

REM --- ensure Slot efficiency.py exists (your code loads this exact name) ---
if not exist "Slot efficiency.py" (
  echo [ERROR] Missing file: "Slot efficiency.py"
  echo Your Summary.py loads this file dynamically for Slot Efficiency tab.
  pause
  exit /b 1
)

REM --- create efficiency.py alias for compatibility (some builds look for it) ---
if not exist "efficiency.py" (
  echo [INFO] Creating efficiency.py alias from "Slot efficiency.py" ...
  copy /Y "Slot efficiency.py" "efficiency.py" >nul
)

REM --- create venv if missing ---
if not exist "%VENV_DIR%\Scripts\python.exe" (
  echo [INFO] Creating venv in "%VENV_DIR%"...
  python -m venv "%VENV_DIR%"
)

call "%VENV_DIR%\Scripts\activate"

echo [INFO] Python:
python -c "import sys,struct,platform; print(sys.executable); print('bits=',struct.calcsize('P')*8,'machine=',platform.machine())"
echo.

REM --- deps ---
python -m pip install --upgrade pip
python -m pip install -U pyinstaller pyqt5 pandas numpy openpyxl pyodbc matplotlib

REM --- clean old builds ---
echo [INFO] Cleaning build/dist...
if exist "build" rmdir /s /q "build"
if exist "dist" rmdir /s /q "dist"
del /q *.spec 2>nul

REM --- build onefile (bundle dynamic-loaded .py files) ---
echo [INFO] Building --onefile...
python -m PyInstaller --noconfirm --clean --onefile --windowed ^
  --name "%APP_NAME%" ^
  --add-data "Slot efficiency.py;." ^
  --add-data "efficiency.py;." ^
  --add-data "trial0.py;." ^
  "%MAIN_PY%"

if errorlevel 1 (
  echo.
  echo [ERROR] Build failed. Scroll up for details.
  pause
  exit /b 1
)

echo.
echo [OK] Build finished:
echo %cd%\dist\%APP_NAME%.exe
echo.
echo IMPORTANT (your choice): keep these next to EXE externally:
echo   spec_limits.txt
echo   yield_spec_limits.txt
echo   yield_saved_lists.xlsx (if used)
echo.
pause
endlocal
