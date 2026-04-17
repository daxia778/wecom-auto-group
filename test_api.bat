@echo off
chcp 65001 >nul 2>&1
echo ====================================
echo   WeCom API 集成测试
echo ====================================
echo.
echo 正在编译并运行 API 测试...
echo.
cd /d "%~dp0"
go run . --test
echo.
echo ====================================
echo   测试完成
echo ====================================
pause
