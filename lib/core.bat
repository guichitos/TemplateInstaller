rem ======================================================
rem === CORE LIBRARY - Common utilities and shared tasks ===
rem ------------------------------------------------------
rem Provides general-purpose functions for the installer,
rem including closing Office applications (Word, PowerPoint,
rem Excel), logging helpers, and any reusable core logic
rem required by other modules.
rem ======================================================

@echo off
if "%~1"=="" exit /b
set "FUNC=%~1"
shift
goto %FUNC%

:Log
rem Args: LOG_FILE, MESSAGE
set "LOG_FILE=%~1"
shift
echo [%DATE% %TIME%] %* >> "%LOG_FILE%"
exit /b

:CloseOfficeApps
rem Args: LOG_FILE
set "LOG_FILE=%~1"
call :Log "%LOG_FILE%" Closing Office apps...
taskkill /IM WINWORD.EXE /F >nul 2>&1
taskkill /IM POWERPNT.EXE /F >nul 2>&1
taskkill /IM EXCEL.EXE /F >nul 2>&1
exit /b
