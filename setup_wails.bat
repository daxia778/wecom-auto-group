@echo off
set "PATH=C:\Windows\System32;C:\Windows;C:\Program Files\Go\bin;C:\Program Files\nodejs;%USERPROFILE%\go\bin;%PATH%"

echo.
echo ========================================
echo   Installing Wails CLI
echo ========================================
echo.

echo [1/3] Installing Wails CLI...
go install github.com/wailsapp/wails/v2/cmd/wails@latest
if %errorlevel% neq 0 (
    echo FAILED to install Wails CLI
    echo Result: FAILED > "%~dp0setup_result.txt"
    pause
    exit /b 1
)
echo   OK

echo.
echo [2/3] Checking wails...
wails version
if %errorlevel% neq 0 (
    echo   Trying full path...
    "%USERPROFILE%\go\bin\wails.exe" version
)

echo.
echo [3/3] Wails doctor...
wails doctor 2>nul || "%USERPROFILE%\go\bin\wails.exe" doctor 2>nul

echo.
echo ========================================
echo   DONE
echo ========================================
echo Result: SUCCESS > "%~dp0setup_result.txt"
echo Go: >> "%~dp0setup_result.txt"
go version >> "%~dp0setup_result.txt" 2>&1
echo Node: >> "%~dp0setup_result.txt"
node --version >> "%~dp0setup_result.txt" 2>&1
echo Wails: >> "%~dp0setup_result.txt"
wails version >> "%~dp0setup_result.txt" 2>&1
"%USERPROFILE%\go\bin\wails.exe" version >> "%~dp0setup_result.txt" 2>&1
echo. >> "%~dp0setup_result.txt"

pause
