@echo off
set "OUTFILE=%~dp0env_result.txt"
set "PATH=C:\Windows\System32;C:\Windows;C:\Windows\System32\Wbem;C:\Windows\System32\WindowsPowerShell\v1.0;%PATH%"

echo Checking environment... > "%OUTFILE%"
echo. >> "%OUTFILE%"

echo === Go === >> "%OUTFILE%"
where go >> "%OUTFILE%" 2>&1
go version >> "%OUTFILE%" 2>&1
echo. >> "%OUTFILE%"

echo === Node === >> "%OUTFILE%"
where node >> "%OUTFILE%" 2>&1
node --version >> "%OUTFILE%" 2>&1
echo. >> "%OUTFILE%"

echo === npm === >> "%OUTFILE%"
where npm >> "%OUTFILE%" 2>&1
npm --version >> "%OUTFILE%" 2>&1
echo. >> "%OUTFILE%"

echo === Git === >> "%OUTFILE%"
where git >> "%OUTFILE%" 2>&1
git --version >> "%OUTFILE%" 2>&1
echo. >> "%OUTFILE%"

echo === Search Git on D === >> "%OUTFILE%"
if exist "D:\Program Files\Git\cmd\git.exe" echo FOUND: D:\Program Files\Git\cmd\git.exe >> "%OUTFILE%"
if exist "D:\Git\cmd\git.exe" echo FOUND: D:\Git\cmd\git.exe >> "%OUTFILE%"
if exist "D:\Git\bin\git.exe" echo FOUND: D:\Git\bin\git.exe >> "%OUTFILE%"
if exist "C:\Program Files\Git\cmd\git.exe" echo FOUND: C:\Program Files\Git\cmd\git.exe >> "%OUTFILE%"
echo. >> "%OUTFILE%"

echo === Search Go === >> "%OUTFILE%"
if exist "C:\Program Files\Go\bin\go.exe" echo FOUND: C:\Program Files\Go\bin\go.exe >> "%OUTFILE%"
if exist "D:\Go\bin\go.exe" echo FOUND: D:\Go\bin\go.exe >> "%OUTFILE%"
if exist "C:\Go\bin\go.exe" echo FOUND: C:\Go\bin\go.exe >> "%OUTFILE%"
echo. >> "%OUTFILE%"

echo Done. Results in: %OUTFILE%
