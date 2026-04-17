@echo off
setlocal enabledelayedexpansion

set "MYDIR=%~dp0"
set "PY_DIR=%MYDIR%python_embed"
set "PY_EXE=%PY_DIR%\python.exe"
set "PY_URL=https://www.python.org/ftp/python/3.11.9/python-3.11.9-embed-amd64.zip"
set "PIP_URL=https://bootstrap.pypa.io/get-pip.py"

echo.
echo ========================================
echo   WeCom Auto Group Creator v1.0
echo ========================================
echo.

REM --- Check existing Python ---
if exist "%PY_EXE%" goto :CHECK_PIP

echo [1/4] Downloading Python 3.11...
powershell -ExecutionPolicy Bypass -Command "Invoke-WebRequest -Uri '%PY_URL%' -OutFile '%MYDIR%python.zip'"
if not exist "%MYDIR%python.zip" (
    echo ERROR: Download failed.
    pause
    exit /b 1
)

echo [2/4] Extracting...
powershell -ExecutionPolicy Bypass -Command "Expand-Archive -Path '%MYDIR%python.zip' -DestinationPath '%PY_DIR%' -Force"
del "%MYDIR%python.zip"

REM --- Enable site-packages in _pth file ---
set "PTH_FILE="
for %%f in ("%PY_DIR%\python*._pth") do set "PTH_FILE=%%f"
if defined PTH_FILE (
    echo import site>> "!PTH_FILE!"
)

:CHECK_PIP
REM --- Check pip ---
"%PY_EXE%" -m pip --version >nul 2>&1
if %errorlevel%==0 goto :HAS_PIP

echo [3/4] Installing pip...
powershell -ExecutionPolicy Bypass -Command "Invoke-WebRequest -Uri '%PIP_URL%' -OutFile '%MYDIR%get-pip.py'"
"%PY_EXE%" "%MYDIR%get-pip.py" --no-warn-script-location
del "%MYDIR%get-pip.py" 2>nul

:HAS_PIP
echo [4/4] Checking dependencies...
"%PY_EXE%" -m pip install pywin32 Pillow -q --no-warn-script-location 2>nul

echo.
echo Python ready: %PY_EXE%
echo.

REM --- Run ---
"%PY_EXE%" "%MYDIR%wecom_auto.py"

echo.
pause
