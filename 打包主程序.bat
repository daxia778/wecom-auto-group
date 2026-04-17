@echo off
chcp 65001 >nul 2>&1
title 打包 WeComAutoApp 主程序为 EXE
color 0A

echo.
echo ╔══════════════════════════════════════════════════════╗
echo ║    企微自动建群 - 打包成独立 EXE                    ║
echo ║                                                      ║
echo ║    打包后生成: WeComAutoApp.exe                      ║
echo ║    包含: app.py + wecom_auto.py (OCR引擎)           ║
echo ║    只需 EXE + config.json 即可运行                   ║
echo ╚══════════════════════════════════════════════════════╝
echo.

:: ─── 检查 Python ───
set PYTHON_CMD=
where python >nul 2>&1
if %errorlevel% equ 0 ( set PYTHON_CMD=python & goto :has_python )
where python3 >nul 2>&1
if %errorlevel% equ 0 ( set PYTHON_CMD=python3 & goto :has_python )
where py >nul 2>&1
if %errorlevel% equ 0 ( set PYTHON_CMD=py & goto :has_python )

echo ❌ 未找到 Python！
echo    请先安装 Python: Microsoft Store 搜索 "Python 3.11"
echo    或从 https://www.python.org/downloads/ 下载安装
pause & exit /b 1

:has_python
echo [0/5] 检测 Python 版本...
%PYTHON_CMD% --version
echo.

:: ─── 安装依赖 ───
echo [1/5] 安装 PyInstaller 打包工具...
%PYTHON_CMD% -m pip install pyinstaller -q --no-warn-script-location 2>nul
if %errorlevel% neq 0 (
    echo   ⚠️ pip install 失败, 尝试 --user 模式...
    %PYTHON_CMD% -m pip install pyinstaller -q --user --no-warn-script-location 2>nul
)
echo   ✅ PyInstaller 就绪
echo.

echo [2/5] 安装运行时依赖...
%PYTHON_CMD% -m pip install pywin32 Pillow -q --no-warn-script-location 2>nul
echo   ✅ pywin32 + Pillow 就绪
echo.

:: ─── 清理旧的打包产物 ───
echo [3/5] 清理旧文件...
if exist build rd /s /q build 2>nul
if exist dist rd /s /q dist 2>nul
if exist WeComAutoApp.spec del /f /q WeComAutoApp.spec 2>nul
echo   ✅ 已清理
echo.

:: ─── 开始打包 ───
echo [4/5] 开始打包 (约 30~90 秒, 请耐心等待)...
echo   入口文件: app.py
echo   自动包含: wecom_auto.py (通过 import 自动检测)
echo.

%PYTHON_CMD% -m PyInstaller ^
    --onefile ^
    --console ^
    --name "WeComAutoApp" ^
    --icon NONE ^
    --clean ^
    --noconfirm ^
    --hidden-import win32gui ^
    --hidden-import win32api ^
    --hidden-import win32con ^
    --hidden-import win32ui ^
    --hidden-import win32process ^
    --hidden-import pywintypes ^
    --hidden-import pythoncom ^
    --hidden-import PIL ^
    --hidden-import PIL.Image ^
    --hidden-import tkinter ^
    --hidden-import tkinter.ttk ^
    --hidden-import tkinter.scrolledtext ^
    "%~dp0app.py"

if %errorlevel% neq 0 (
    echo.
    echo ════════════════════════════════════════════════════
    echo   ❌ 打包失败！
    echo.
    echo   常见原因:
    echo   1. PyInstaller 版本不兼容, 试试: pip install pyinstaller==6.3.0
    echo   2. pywin32 未正确安装, 试试: pip install pywin32==306
    echo   3. 杀毒软件拦截, 请暂时关闭后重试
    echo ════════════════════════════════════════════════════
    pause & exit /b 1
)

:: ─── 整理输出 ───
echo.
echo [5/5] 整理输出文件...

if exist "dist\WeComAutoApp.exe" (
    copy /Y "dist\WeComAutoApp.exe" "%~dp0WeComAutoApp.exe" >nul

    :: 获取文件大小
    for %%A in ("%~dp0WeComAutoApp.exe") do set EXE_SIZE=%%~zA
    set /a EXE_MB=%EXE_SIZE% / 1048576

    echo.
    echo ══════════════════════════════════════════════════════
    echo   ✅ 打包成功！
    echo.
    echo   📦 文件: %~dp0WeComAutoApp.exe
    echo   📏 大小: %EXE_MB% MB ^(%EXE_SIZE% bytes^)
    echo.
    echo   📋 使用方法:
    echo     1. 在 WeComAutoApp.exe 同目录放置 config.json
    echo     2. 打开企业微信并登录
    echo     3. 双击 WeComAutoApp.exe 启动
    echo.
    echo   💡 提示: 首次启动较慢 (解压临时文件), 后续会快
    echo ══════════════════════════════════════════════════════
) else (
    echo ❌ 未找到打包产物 dist\WeComAutoApp.exe
    echo    请检查上方的错误日志
)

:: 清理临时文件
echo.
echo 清理打包临时文件...
rd /s /q build 2>nul
rd /s /q dist 2>nul
del /f /q WeComAutoApp.spec 2>nul
rd /s /q __pycache__ 2>nul
echo ✅ 临时文件已清理

echo.
pause
