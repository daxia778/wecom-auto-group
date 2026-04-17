@echo off
set "PATH=C:\Windows\System32;C:\Windows;C:\Program Files\Go\bin;C:\Program Files\nodejs;%USERPROFILE%\go\bin;%PATH%"

echo.
echo ============================================
echo   WeComAutoGroup - Go Build
echo ============================================
echo.

REM Kill old process
taskkill /F /IM WeComAutoGroup.exe >nul 2>&1

echo [1/4] Checking tools...
go version
wails version 2>nul || (
    echo   Installing Wails CLI...
    go install github.com/wailsapp/wails/v2/cmd/wails@latest
)
echo.

echo [2/4] Initializing Go modules...
if not exist go.mod (
    go mod init wecom-auto-group
)
go mod tidy
echo.

echo [3/4] Building EXE (wails build)...
wails build -clean -platform windows/amd64
echo.

if exist "build\bin\WeComAutoGroup.exe" (
    echo ============================================
    echo   SUCCESS! EXE created:
    echo   build\bin\WeComAutoGroup.exe
    echo ============================================
    copy "build\bin\WeComAutoGroup.exe" "WeComAutoGroup.exe" >nul 2>&1
    echo   Also copied to: WeComAutoGroup.exe
) else (
    echo ============================================
    echo   BUILD FAILED - check errors above
    echo ============================================
)
echo.
pause
