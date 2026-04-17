@echo off
setlocal enabledelayedexpansion
chcp 65001 >nul 2>&1
title WeCom Inspector v2.1
color 0A

echo.
echo ========================================================
echo    WeCom UIAutomation Inspector v2.1
echo    Auto-setup: No Python installation needed
echo ========================================================
echo.

set "MYDIR=%~dp0"
set "PYDIR=%MYDIR%python_embed"
set "PYTHON=%PYDIR%\python.exe"

:: --- Check if exe exists ---
if exist "%MYDIR%wecom_inspector.exe" (
    echo [OK] Found exe, launching...
    "%MYDIR%wecom_inspector.exe"
    goto :end
)

:: --- Check if embedded Python ready ---
if exist "%PYTHON%" (
    if exist "%PYDIR%\Lib\site-packages\uiautomation" (
        echo [OK] Environment ready, launching...
        echo.
        "%PYTHON%" "%MYDIR%wecom_inspector.py"
        goto :end
    )
)

:: --- First run: auto setup ---
echo ========================================================
echo   First run - auto downloading Python environment
echo   This takes about 30-60 seconds, please wait...
echo ========================================================
echo.

:: Step 1: Download embedded Python
echo [1/4] Downloading Python 3.11 (embedded)...
if not exist "%PYDIR%" mkdir "%PYDIR%"

powershell -ExecutionPolicy Bypass -Command ^
    "[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12;" ^
    "$urls = @(" ^
    "  'https://www.python.org/ftp/python/3.11.9/python-3.11.9-embed-amd64.zip'," ^
    "  'https://registry.npmmirror.com/-/binary/python/3.11.9/python-3.11.9-embed-amd64.zip'" ^
    ");" ^
    "$ok = $false;" ^
    "foreach ($u in $urls) {" ^
    "  try {" ^
    "    Write-Host '  Trying:' $u;" ^
    "    (New-Object Net.WebClient).DownloadFile($u, '%PYDIR%\python.zip');" ^
    "    Write-Host '  [OK] Downloaded';" ^
    "    $ok = $true; break" ^
    "  } catch { Write-Host '  [FAIL] Try next...' }" ^
    "};" ^
    "if (-not $ok) { Write-Host '  [ERROR] Download failed'; exit 1 }"

if not exist "%PYDIR%\python.zip" (
    echo [ERROR] Python download failed. Check network.
    pause
    exit /b 1
)

:: Step 2: Extract
echo [2/4] Extracting Python...
powershell -ExecutionPolicy Bypass -Command ^
    "Expand-Archive -Force '%PYDIR%\python.zip' '%PYDIR%'"
del "%PYDIR%\python.zip" 2>nul

:: Enable pip: uncomment "import site" in ._pth file
for %%F in ("%PYDIR%\python*._pth") do (
    powershell -ExecutionPolicy Bypass -Command ^
        "(Get-Content '%%F') -replace '#import site','import site' | Set-Content '%%F'"
)
echo   [OK] Extracted

:: Step 3: Install pip
echo [3/4] Installing pip...
powershell -ExecutionPolicy Bypass -Command ^
    "[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12;" ^
    "(New-Object Net.WebClient).DownloadFile('https://bootstrap.pypa.io/get-pip.py', '%PYDIR%\get-pip.py')"

"%PYTHON%" "%PYDIR%\get-pip.py" --no-warn-script-location -q 2>nul
del "%PYDIR%\get-pip.py" 2>nul
echo   [OK] pip ready

:: Step 4: Install dependencies
echo [4/4] Installing uiautomation + pywin32...
"%PYTHON%" -m pip install uiautomation pywin32 -q --no-warn-script-location 2>nul
echo   [OK] Dependencies ready

echo.
echo ========================================================
echo   Setup complete! Launching inspector...
echo ========================================================
echo.

"%PYTHON%" "%MYDIR%wecom_inspector.py"

:end
echo.
pause
endlocal
