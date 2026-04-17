@echo off
set "PATH=C:\Windows\System32;C:\Windows;%PATH%"
cd /d D:\wecom-auto-group

echo Cleaning old Python files...
del /q app.py 2>nul
del /q wecom_auto.py 2>nul
del /q wecom_inspector.py 2>nul

echo Cleaning old batch files...
del /q build.bat 2>nul
del /q test.bat 2>nul
del /q check_env.bat 2>nul
del /q setup_wails.bat 2>nul
del /q env_result.txt 2>nul
del /q setup_result.txt 2>nul

del /q "一键打包.bat" 2>nul
del /q "一键运行.bat" 2>nul
del /q "修复PowerShell路径.bat" 2>nul
del /q "启动建群.bat" 2>nul
del /q "打包主程序.bat" 2>nul
del /q "打包成exe.bat" 2>nul

echo Cleaning old EXE...
del /q WeComAutoApp.exe 2>nul

echo Cleaning old config (Go version has built-in keys)...
del /q config.json 2>nul

echo.
echo Done! Remaining files:
dir /b
echo.

REM Self-delete
del /q "%~f0"
