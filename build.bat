@echo off
echo.
echo ========================================
echo   WeComAutoApp Packaging Script
echo   Mode: GUI only (no console window)
echo ========================================
echo.

echo [Step 0] Killing old WeComAutoApp process...
taskkill /F /IM WeComAutoApp.exe >nul 2>&1
if %errorlevel% equ 0 (
    echo   Killed old process. Waiting 2s...
    timeout /t 2 /nobreak >nul
) else (
    echo   No running process found. OK.
)
echo.

set "PATH=C:\Windows\System32;C:\Windows;C:\Windows\System32\Wbem;C:\Windows\System32\WindowsPowerShell\v1.0;%PATH%"

echo [Step 1] Finding Python...
set PYTHON_CMD=

python --version >nul 2>&1
if %errorlevel% equ 0 (
    set PYTHON_CMD=python
    goto :found
)
python3 --version >nul 2>&1
if %errorlevel% equ 0 (
    set PYTHON_CMD=python3
    goto :found
)
py --version >nul 2>&1
if %errorlevel% equ 0 (
    set PYTHON_CMD=py
    goto :found
)

echo ERROR: Python not found!
pause
exit /b 1

:found
echo   Found: %PYTHON_CMD%
%PYTHON_CMD% --version
echo.

echo [Step 2] Installing dependencies...
%PYTHON_CMD% -m pip install pyinstaller pywin32 Pillow -q --no-warn-script-location
echo   Done.
echo.

echo [Step 3] Cleaning old files...
if exist build rd /s /q build 2>nul
if exist dist rd /s /q dist 2>nul
if exist WeComAutoApp.spec del /f /q WeComAutoApp.spec 2>nul
if exist "%~dp0WeComAutoApp.exe" del /f /q "%~dp0WeComAutoApp.exe" 2>nul
echo   Done.
echo.

echo [Step 4] Packaging with PyInstaller (30-90 seconds)...
echo   Entry: app.py
echo   Mode: --windowed (GUI only, no console)
echo.

%PYTHON_CMD% -m PyInstaller --onefile --windowed --name "WeComAutoApp" --icon NONE --clean --noconfirm --hidden-import win32gui --hidden-import win32api --hidden-import win32con --hidden-import win32ui --hidden-import win32process --hidden-import pywintypes --hidden-import pythoncom --hidden-import PIL --hidden-import PIL.Image --hidden-import tkinter --hidden-import tkinter.ttk --hidden-import tkinter.scrolledtext "%~dp0app.py"

if %errorlevel% neq 0 (
    echo.
    echo ERROR: Packaging failed!
    pause
    exit /b 1
)

echo.
echo [Step 5] Copying output...
if exist "dist\WeComAutoApp.exe" (
    copy /Y "dist\WeComAutoApp.exe" "%~dp0WeComAutoApp.exe" >nul
    echo.
    echo ========================================
    echo   SUCCESS! WeComAutoApp.exe created!
    echo   Location: %~dp0WeComAutoApp.exe
    echo.
    echo   No console window - GUI only!
    echo ========================================
) else (
    echo ERROR: WeComAutoApp.exe not found in dist folder
)

echo.
echo Cleaning temp files...
rd /s /q build 2>nul
rd /s /q dist 2>nul
del /f /q WeComAutoApp.spec 2>nul
rd /s /q __pycache__ 2>nul
echo Done.
echo.
pause
