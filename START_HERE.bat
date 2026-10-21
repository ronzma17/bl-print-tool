@echo off
chcp 65001 >nul
setlocal
title Bill of Lading Print Tool

set "BASE=%~dp0"
set "S1=%BASE%prep_print.py"
set "S2=%BASE%prep_prearrival.py"
set "INBOX=%BASE%1_DROP_ZIPS_HERE"
set "OUTBOX=%BASE%2_READY_TO_PRINT"
set "ARCHIVE=%BASE%3_ARCHIVE"
set "PABOX=%BASE%4_DROP_PREARRIVAL_HERE"

if not exist "%INBOX%"   mkdir "%INBOX%"
if not exist "%OUTBOX%"  mkdir "%OUTBOX%"
if not exist "%ARCHIVE%" mkdir "%ARCHIVE%"
if not exist "%PABOX%"   mkdir "%PABOX%"
if not exist "%S1%"      goto noscript

echo.
echo ==========================================
echo   BILL OF LADING  ^>  PRINT  TOOL
echo ==========================================
echo.
echo   SHEET 1   B/L + Waybill   (DUPLEX)
echo     put zip into  1_DROP_ZIPS_HERE
echo.
echo      1 = CHECK ONLY
echo      2 = MAKE PDF   50 per file
echo      3 = MAKE PDF   all in one
echo.
echo   SHEET 2   Pre-Arrival Notice   (SINGLE SIDE)
echo     put zip into  4_DROP_PREARRIVAL_HERE
echo     run this AFTER sheet 1
echo.
echo      4 = CHECK ONLY
echo      5 = MAKE PDF
echo.
echo   CATCH-UP   Waybill arrived later   (SINGLE SIDE)
echo     put the waybill zip into  1_DROP_ZIPS_HERE
echo     prints the BACK of sheet 1, same order
echo.
echo      6 = MAKE WAYBILL PDF
echo.
echo      0 = open the folders
echo.
set /p MODE=   Type a number then Enter: 
echo.

if "%MODE%"=="0" goto openall
if "%MODE%"=="1" goto mode1
if "%MODE%"=="2" goto mode2
if "%MODE%"=="3" goto mode3
if "%MODE%"=="4" goto mode4
if "%MODE%"=="5" goto mode5
if "%MODE%"=="6" goto mode6
echo   Invalid choice.
goto theend

:mode1
python "%S1%" "%INBOX%" --out "%OUTBOX%" --dry-run
goto finished

:mode2
python "%S1%" "%INBOX%" --out "%OUTBOX%" --archive "%ARCHIVE%" --chunk 50
goto finished

:mode3
python "%S1%" "%INBOX%" --out "%OUTBOX%" --archive "%ARCHIVE%"
goto finished

:mode4
if not exist "%S2%" goto noscript2
python "%S2%" "%PABOX%" --out "%OUTBOX%" --dry-run
goto finished

:mode5
if not exist "%S2%" goto noscript2
python "%S2%" "%PABOX%" --out "%OUTBOX%" --archive "%ARCHIVE%"
goto finished

:mode6
python "%S1%" "%INBOX%" --out "%OUTBOX%" --archive "%ARCHIVE%" --waybill-only
goto finished

:finished
echo.
echo ==========================================
echo   Done.  Open 2_READY_TO_PRINT and print.
echo   PRINT.pdf           duplex, flip LONG edge
echo   PRINT_SHEET2.pdf    SINGLE side
echo   PRINT_WAYBILL.pdf   SINGLE side, paper flipped
echo ==========================================
echo.
start "" "%OUTBOX%"
goto theend

:openall
start "" "%INBOX%"
start "" "%PABOX%"
start "" "%OUTBOX%"
goto theend

:noscript
echo   ERROR: prep_print.py is missing.
goto theend

:noscript2
echo   ERROR: prep_prearrival.py is missing.
goto theend

:theend
echo.
pause
