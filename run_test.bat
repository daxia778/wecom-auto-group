@echo off
chcp 65001 >nul
echo ================================
echo   WeCom API 测试 (服务器中转)
echo ================================
echo.
cd /d "%~dp0"
go run . --test
echo.
echo ================================
pause
