@echo off
setlocal enabledelayedexpansion

rem ======================================================
rem === INSTALLER LIBRARY FOR BASE TEMPLATES (v2) ===
rem ------------------------------------------------------
rem Handles installation of default Office templates:
rem   - Word → Normal.dotx / Normal.dotm
rem   - PowerPoint → Blank.potx / Blank.potm
rem   - Excel → Book.xltx / Book.xltm / Sheet.xltx / Sheet.xltm
rem Performs backup, overwrites safely, logs activity,
rem and optionally creates debug evidence files when
rem IsDesignModeEnabled=true.
rem ======================================================

if "%~1"=="" exit /b
set "FUNC=%~1"
shift
goto %FUNC%

:InstallApp
rem Args: APP SRC_NAME DST_DIR DST_NAME LOG_FILE BASE_DIR IsDesignModeEnabled
set "AppName=%~1"
set "SourceFileName=%~2"
set "DestinationDirectory=%~3"
set "DestinationFileName=%~4"
set "LogFilePath=%~5"
set "SourceDirectory=%~6"
set "IsDesignModeEnabled=%~7"

set "SourceFilePath=%SourceDirectory%%SourceFileName%"
set "DestinationFilePath=%DestinationDirectory%\%DestinationFileName%"
set "BackupFilePath=%DestinationDirectory%\%~n4_backup%~x4"

rem === Check if source exists ===
if not exist "%SourceFilePath%" (
    if /I "%IsDesignModeEnabled%"=="true" echo [ERROR] Source file not found: "%SourceFilePath%"
    if /I "%IsDesignModeEnabled%"=="true" echo [ERROR] Source file not found "%SourceFilePath%". >> "%LogFilePath%"
    exit /b
)

rem === Ensure destination folder exists ===
if not exist "%DestinationDirectory%" mkdir "%DestinationDirectory%" 2>nul

rem === Create backup if existing file found ===
if exist "%DestinationFilePath%" (
    copy /Y "%DestinationFilePath%" "%BackupFilePath%" >nul 2>&1
    if /I "%IsDesignModeEnabled%"=="true" (
        echo [BACKUP] Created for %AppName% template at "%BackupFilePath%"
        echo [BACKUP] Created for %AppName% template at "%BackupFilePath%" >> "%LogFilePath%"
    )
)

rem === Copy new template ===
copy /Y "%SourceFilePath%" "%DestinationFilePath%" >nul 2>&1
if exist "%DestinationFilePath%" (
    if /I "%IsDesignModeEnabled%"=="true" (
        echo [OK] Installed %AppName% template at "%DestinationFilePath%"
        echo [OK] Installed %AppName% template at "%DestinationFilePath%" >> "%LogFilePath%"
    )
) else (
    if /I "%IsDesignModeEnabled%"=="true" (
        echo [ERROR] Copy failed for "%SourceFilePath%"
        echo [ERROR] Copy failed for "%SourceFilePath%" >> "%LogFilePath%"
    )
)

endlocal
exit /b
