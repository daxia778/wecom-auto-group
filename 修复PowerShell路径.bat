@echo off
chcp 65001 >nul 2>&1
title 修复 PowerShell 系统 PATH

echo.
echo ╔══════════════════════════════════════════════════╗
echo ║    修复 PowerShell 到系统 PATH                  ║
echo ║                                                  ║
echo ║    将 PowerShell 目录添加到用户环境变量          ║
echo ║    修复后需要重启终端/IDE 才生效                 ║
echo ╚══════════════════════════════════════════════════╝
echo.

:: 检查 PowerShell 是否已在 PATH 中
where powershell >nul 2>&1
if %errorlevel% equ 0 (
    echo ✅ PowerShell 已经在 PATH 中，无需修复！
    powershell --version
    pause
    exit /b 0
)

:: 检查 PowerShell 是否存在
set PS_DIR=C:\Windows\System32\WindowsPowerShell\v1.0
if not exist "%PS_DIR%\powershell.exe" (
    echo ❌ 未找到 PowerShell: %PS_DIR%\powershell.exe
    echo    请确认 Windows 系统完整性
    pause
    exit /b 1
)

echo [1/2] 找到 PowerShell: %PS_DIR%\powershell.exe
echo.

:: 将 PowerShell 目录添加到用户 PATH（永久生效）
echo [2/2] 正在添加到用户 PATH 环境变量...

:: 读取当前用户 PATH
for /f "tokens=2*" %%A in ('reg query "HKCU\Environment" /v Path 2^>nul') do set "CURRENT_PATH=%%B"

:: 检查是否已包含
echo %CURRENT_PATH% | findstr /I /C:"%PS_DIR%" >nul 2>&1
if %errorlevel% equ 0 (
    echo ✅ PATH 中已包含 PowerShell 目录，但可能未刷新
    echo    请重启你正在使用的 IDE / 终端
    pause
    exit /b 0
)

:: 添加到用户 PATH
if defined CURRENT_PATH (
    setx PATH "%CURRENT_PATH%;%PS_DIR%" >nul 2>&1
) else (
    setx PATH "%PS_DIR%" >nul 2>&1
)

if %errorlevel% equ 0 (
    echo.
    echo ══════════════════════════════════════════════════
    echo   ✅ 修复成功！
    echo.
    echo   已将以下路径添加到用户 PATH:
    echo   %PS_DIR%
    echo.
    echo   ⚠️  重要: 请执行以下操作使其生效:
    echo.
    echo   1. 完全关闭并重新打开你的 IDE (如 VS Code / Cursor)
    echo   2. 或者注销并重新登录 Windows
    echo.
    echo   验证方法: 打开新的 CMD 窗口，输入 powershell --version
    echo ══════════════════════════════════════════════════
) else (
    echo ❌ 添加失败，请尝试手动添加:
    echo    1. Win+R 输入 sysdm.cpl 回车
    echo    2. 高级 → 环境变量
    echo    3. 用户变量 → Path → 编辑 → 新建
    echo    4. 输入: %PS_DIR%
    echo    5. 确定保存，重启 IDE
)

echo.
pause
