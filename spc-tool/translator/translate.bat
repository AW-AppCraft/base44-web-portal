@echo off
REM ===================================================================
REM  SPC Translator - double-click launcher
REM
REM  Edit the three paths below, then double-click this file.
REM  Needs Python 3 on the PC (python.org, tick "Add to PATH").
REM ===================================================================

set INBOX=C:\SPC\inbox
set OUTBOX=C:\SPC\spc-ready
set REJECTS=C:\SPC\rejects

REM --- Optional: force the feature order regardless of source column order.
REM     Leave blank to take the reading columns in the order they appear.
set FEATURES=

REM --- Rescan interval in seconds. Set to 0 for a single pass.
set INTERVAL=30

REM ===================================================================
cd /d "%~dp0"

if not exist "%INBOX%" (
    echo Input folder does not exist: %INBOX%
    echo Edit the INBOX line in this file.
    pause
    exit /b 1
)

set ARGS=--in "%INBOX%" --out "%OUTBOX%" --rejects "%REJECTS%"
if not "%FEATURES%"=="" set ARGS=%ARGS% --features "%FEATURES%"
if not "%INTERVAL%"=="0" set ARGS=%ARGS% --watch %INTERVAL%

python spc_translate.py %ARGS%

if errorlevel 1 (
    echo.
    echo Finished with rejected files - see translation_log.csv in %INBOX%
)
pause
