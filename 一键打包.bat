@echo off
chcp 65001 >nul 2>&1
title 一键打包 WeComAutoApp

:: 强制确保系统关键目录在 PATH 中
set "PATH=C:\Windows\System32;C:\Windows;C:\Windows\System32\Wbem;C:\Windows\System32\WindowsPowerShell\v1.0;%PATH%"

echo.
echo ══════════════════════════════════════════════════════
echo   企微自动建群 - 一键打包成 EXE
echo ══════════════════════════════════════════════════════
echo.

:: ─── 第1步: 查找 Python ───
echo [1/5] 查找 Python...
set PYTHON_CMD=

where python >nul 2>&1
if %errorlevel% equ 0 (
    set PYTHON_CMD=python
    goto :found_python
)
where python3 >nul 2>&1
if %errorlevel% equ 0 (
    set PYTHON_CMD=python3
    goto :found_python
)
where py >nul 2>&1
if %errorlevel% equ 0 (
    set PYTHON_CMD=py
    goto :found_python
)

:: 尝试常见安装路径
if exist "C:\Python311\python.exe" (
    set "PYTHON_CMD=C:\Python311\python.exe"
    goto :found_python
)
if exist "C:\Python310\python.exe" (
    set "PYTHON_CMD=C:\Python310\python.exe"
    goto :found_python
)
for /f "delims=" %%i in ('dir /b /s "C:\Users\%USERNAME%\AppData\Local\Programs\Python\python.exe" 2^>nul') do (
    set "PYTHON_CMD=%%i"
    goto :found_python
)
for /f "delims=" %%i in ('dir /b /s "C:\Users\%USERNAME%\AppData\Local\Microsoft\WindowsApps\python*.exe" 2^>nul') do (
    set "PYTHON_CMD=%%i"
    goto :found_python
)

echo.
echo   ❌ 未找到 Python！
echo   请安装 Python: https://www.python.org/downloads/
echo.
pause
exit /b 1

:found_python
echo   ✅ 找到 Python: %PYTHON_CMD%
%PYTHON_CMD% --version
echo.

:: ─── 第2步: 安装依赖 ───
echo [2/5] 安装 PyInstaller 和运行依赖...
%PYTHON_CMD% -m pip install pyinstaller pywin32 Pillow -q --no-warn-script-location 2>nul
if %errorlevel% neq 0 (
    echo   ⚠️ 常规安装失败, 尝试 --user 模式...
    %PYTHON_CMD% -m pip install pyinstaller pywin32 Pillow -q --user --no-warn-script-location 2>nul
)
echo   ✅ 依赖安装完成
echo.

:: ─── 第3步: 清理旧文件 ───
echo [3/5] 清理旧的打包产物...
if exist build rd /s /q build 2>nul
if exist dist rd /s /q dist 2>nul
if exist WeComAutoApp.spec del /f /q WeComAutoApp.spec 2>nul
echo   ✅ 已清理
echo.

:: ─── 第4步: 打包 ───
echo [4/5] 开始 PyInstaller 打包 (约 30~90 秒)...
echo   入口: app.py
echo   模式: 单文件 (--onefile)
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
    echo   ❌ 打包失败！错误信息在上方。
    echo.
    echo   常见原因:
    echo     1. 杀毒软件拦截 → 暂时关闭杀毒软件
    echo     2. pip 版本过旧 → python -m pip install --upgrade pip
    echo     3. pywin32 问题 → pip install pywin32==306
    echo.
    pause
    exit /b 1
)

:: ─── 第5步: 整理输出 ───
echo.
echo [5/5] 整理输出文件...

if exist "dist\WeComAutoApp.exe" (
    copy /Y "dist\WeComAutoApp.exe" "%~dp0WeComAutoApp.exe" >nul

    echo.
    echo ══════════════════════════════════════════════════════
    echo   ✅ 打包成功！
    echo.
    echo   📦 文件: %~dp0WeComAutoApp.exe
    echo.
    echo   📋 使用方法:
    echo     1. 在 WeComAutoApp.exe 同目录放置 config.json
    echo     2. 打开企业微信并登录
    echo     3. 双击 WeComAutoApp.exe 启动
    echo ══════════════════════════════════════════════════════
) else (
    echo   ❌ 未找到 dist\WeComAutoApp.exe
    echo   请检查上方错误日志
)

:: 清理临时文件
echo.
echo 清理临时文件...
rd /s /q build 2>nul
rd /s /q dist 2>nul
del /f /q WeComAutoApp.spec 2>nul
rd /s /q __pycache__ 2>nul
echo ✅ 已清理

echo.
pause
