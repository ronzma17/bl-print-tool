@echo off
chcp 65001 >nul
setlocal
title Bill of Lading Print Tool

set "BASE=%~dp0"
set "SCRIPT=%BASE%prep_print.py"
set "INBOX=%BASE%1_DROP_ZIPS_HERE"
set "OUTBOX=%BASE%2_READY_TO_PRINT"
set "ARCHIVE=%BASE%3_ARCHIVE"

if not exist "%INBOX%"   mkdir "%INBOX%"
if not exist "%OUTBOX%"  mkdir "%OUTBOX%"
if not exist "%ARCHIVE%" mkdir "%ARCHIVE%"
if not exist "%SCRIPT%"  goto noscript

echo.
echo ==========================================
echo   BILL OF LADING  ^>  PRINT  TOOL
echo ==========================================
echo.
echo   Step 1  put the zip files into
echo           1_DROP_ZIPS_HERE
echo.
echo      1 = CHECK ONLY
echo      2 = MAKE PDF   50 per file
echo      3 = MAKE PDF   all in one
echo      0 = open the folders
echo.
set /p MODE=   Type a number then Enter: 
echo.

if "%MODE%"=="0" goto openall
if "%MODE%"=="1" goto mode1
if "%MODE%"=="2" goto mode2
if "%MODE%"=="3" goto mode3
echo   Invalid choice.
goto theend

:mode1
python "%SCRIPT%" "%INBOX%" --out "%OUTBOX%" --dry-run
goto finished

:mode2
python "%SCRIPT%" "%INBOX%" --out "%OUTBOX%" --archive "%ARCHIVE%" --chunk 50
goto finished

:mode3
python "%SCRIPT%" "%INBOX%" --out "%OUTBOX%" --archive "%ARCHIVE%"
goto finished

:finished
echo.
echo   Done.  Open 2_READY_TO_PRINT and print.
echo   Duplex + Flip on LONG edge + Letter
echo.
start "" "%OUTBOX%"
goto theend

:openall
start "" "%INBOX%"
start "" "%OUTBOX%"
goto theend

:noscript
echo   ERROR: prep_print.py is missing.
goto theend

:theend
echo.
pause
