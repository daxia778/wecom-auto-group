@echo off
chcp 65001 >nul 2>&1
title 打包成 EXE
color 0E

echo.
echo ╔══════════════════════════════════════════════════╗
echo ║  将探查器打包成独立 EXE (无需 Python 即可运行)  ║
echo ║                                                  ║
echo ║  打包后生成: wecom_inspector.exe                ║
echo ║  只需要这一个文件即可，不需要 Python             ║
echo ╚══════════════════════════════════════════════════╝
echo.

:: 检查 Python
where python >nul 2>&1
if %errorlevel% equ 0 ( set PYTHON_CMD=python & goto :has_python )
where python3 >nul 2>&1
if %errorlevel% equ 0 ( set PYTHON_CMD=python3 & goto :has_python )
where py >nul 2>&1
if %errorlevel% equ 0 ( set PYTHON_CMD=py & goto :has_python )

echo ❌ 需要 Python 来打包 (打包完成后就不需要了)
echo    请先安装 Python: Microsoft Store 搜索 "Python 3.11"
pause & exit /b 1

:has_python
echo [1/4] 安装打包工具 PyInstaller...
%PYTHON_CMD% -m pip install pyinstaller -q 2>nul
echo   ✅ PyInstaller 就绪

echo [2/4] 安装运行依赖...
%PYTHON_CMD% -m pip install uiautomation pywin32 -q 2>nul
echo   ✅ 依赖就绪

echo [3/4] 开始打包 (约30~60秒)...
%PYTHON_CMD% -m PyInstaller ^
    --onefile ^
    --console ^
    --name "wecom_inspector" ^
    --icon NONE ^
    --clean ^
    --noconfirm ^
    --hidden-import uiautomation ^
    --hidden-import win32gui ^
    --hidden-import win32process ^
    --hidden-import win32api ^
    --hidden-import win32con ^
    --hidden-import pywintypes ^
    --hidden-import pythoncom ^
    --collect-all uiautomation ^
    "%~dp0wecom_inspector.py"

if %errorlevel% neq 0 (
    echo.
    echo ❌ 打包失败！请截图错误信息发回来
    pause & exit /b 1
)

echo [4/4] 整理文件...

:: 复制到当前目录
if exist "dist\wecom_inspector.exe" (
    copy /Y "dist\wecom_inspector.exe" "%~dp0wecom_inspector.exe" >nul
    echo.
    echo ══════════════════════════════════════════════════
    echo   ✅ 打包成功！
    echo.
    echo   生成文件: %~dp0wecom_inspector.exe
    echo   大小: 
    for %%A in ("%~dp0wecom_inspector.exe") do echo     %%~zA bytes
    echo.
    echo   这个 exe 可以直接双击运行，不需要 Python！
    echo   把它发给任何 Windows 电脑都能用
    echo ══════════════════════════════════════════════════
) else (
    echo ❌ 未找到打包产物
)

:: 清理临时文件
echo.
echo 清理临时文件...
rd /s /q build 2>nul
rd /s /q dist 2>nul
del /f /q wecom_inspector.spec 2>nul
rd /s /q __pycache__ 2>nul

echo.
pause
