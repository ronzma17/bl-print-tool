@echo off
chcp 65001 >nul
setlocal

set "SCRIPT=%~dp0prep_print.py"
set "TARGET=%~1"

if not exist "%SCRIPT%" goto noscript
if "%TARGET%"=="" goto nodrop

echo.
echo   Folder: %TARGET%
echo.
echo   1 = check only
echo   2 = make PDF, 50 per file
echo   3 = make PDF, all in one
echo.
set /p MODE=   Type a number then Enter: 
echo.

if "%MODE%"=="1" goto mode1
if "%MODE%"=="2" goto mode2
if "%MODE%"=="3" goto mode3
echo   Invalid choice.
goto theend

:mode1
python "%SCRIPT%" "%TARGET%" --dry-run
goto theend

:mode2
python "%SCRIPT%" "%TARGET%" --chunk 50
goto theend

:mode3
python "%SCRIPT%" "%TARGET%"
goto theend

:noscript
echo   ERROR: prep_print.py is missing.
goto theend

:nodrop
echo   Drag your folder onto this bat file.
goto theend

:theend
echo.
pause
