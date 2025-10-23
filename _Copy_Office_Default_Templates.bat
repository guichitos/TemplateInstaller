rem _Main_installer.bat
@echo off
setlocal enabledelayedexpansion
chcp 65001 >nul

rem ======================================================
rem === UNIVERSAL OFFICE TEMPLATES INSTALLER (MAIN) ======
rem ------------------------------------------------------
rem Entry point for the modular installer system.
rem Coordinates environment checks, closes Office apps,
rem installs base templates for Word, PowerPoint, and Excel,
rem copies user custom templates, and logs all operations.
rem ======================================================

rem === DESIGN / DEBUG MODE CONTROL ======================
rem If IsDesignModeEnabled=true  → shows console messages and generates log.
rem If IsDesignModeEnabled=false → runs silently (no output, no log file).
set "IsDesignModeEnabled=false"

rem === Base paths and library references ================
set "BaseDirectoryPath=%~dp0"
set "LibraryDirectoryPath=%BaseDirectoryPath%lib"
set "LogsDirectoryPath=%BaseDirectoryPath%logs"
set "LogFilePath=%LogsDirectoryPath%\install_log_all.txt"

echo Executing. Please wait...
rem === Initialize log only if design mode is enabled ====
if /I "%IsDesignModeEnabled%"=="true" (
    if not exist "%LogsDirectoryPath%" mkdir "%LogsDirectoryPath%"
    echo. > "%LogFilePath%"
    echo [%DATE% %TIME%] --- START TEMPLATES INSTALLATION --- >> "%LogFilePath%"
)

rem === Library references ===============================
set "InstallerLibraryPath=%LibraryDirectoryPath%\installer_apps.bat"
set "CopyLibraryPath=%LibraryDirectoryPath%\copy_templates.bat"
set "RegistryLibraryPath=%LibraryDirectoryPath%\registry_tools.bat"

rem === Header message =============
if /I "%IsDesignModeEnabled%"=="true" (
    title TEMPLATE INSTALLER - DEBUG MODE
    echo [DEBUG] Design mode is enabled.
    echo [INFO] Script is running from: %BaseDirectoryPath%
)

rem === Environment verification and Office shutdown =====
if /I "%IsDesignModeEnabled%"=="true" (
    echo.
    echo [INFO] Verifying environment and closing Office applications...
    call :CheckEnvironment "%LogFilePath%"
    call :CloseOfficeApps "%LogFilePath%"
    echo [OK] Environment verification and Office app closure completed.
    echo [OK] Environment verification and Office app closure completed. >> "%LogFilePath%"
) else (
    call :CheckEnvironment "" >nul 2>&1
    call :CloseOfficeApps "" >nul 2>&1
)



rem === Install base templates for Word, PowerPoint, Excel ===
if /I "%IsDesignModeEnabled%"=="true" (
    echo.
    echo [INFO] Starting base template installation phase...
    rem --- Word templates (Normal.dotx / Normal.dotm) ---
    call "%InstallerLibraryPath%" :InstallApp "WORD" "Normal.dotx" "%APPDATA%\Microsoft\Templates" "Normal.dotx" "%LogFilePath%" "%BaseDirectoryPath%" "%IsDesignModeEnabled%"
    call "%InstallerLibraryPath%" :InstallApp "WORD" "Normal.dotm" "%APPDATA%\Microsoft\Templates" "Normal.dotm" "%LogFilePath%" "%BaseDirectoryPath%" "%IsDesignModeEnabled%"
    rem --- PowerPoint templates (Blank.potx / Blank.potm) ---
    call "%InstallerLibraryPath%" :InstallApp "POWERPOINT" "Blank.potx" "%APPDATA%\Microsoft\Templates" "Blank.potx" "%LogFilePath%" "%BaseDirectoryPath%" "%IsDesignModeEnabled%"
    call "%InstallerLibraryPath%" :InstallApp "POWERPOINT" "Blank.potm" "%APPDATA%\Microsoft\Templates" "Blank.potm" "%LogFilePath%" "%BaseDirectoryPath%" "%IsDesignModeEnabled%"
    rem --- Excel templates (Book / Sheet in xltx & xltm) ---
    call "%InstallerLibraryPath%" :InstallApp "EXCEL" "Book.xltx" "%APPDATA%\Microsoft\Excel\XLSTART" "Book.xltx" "%LogFilePath%" "%BaseDirectoryPath%" "%IsDesignModeEnabled%"
    call "%InstallerLibraryPath%" :InstallApp "EXCEL" "Book.xltm" "%APPDATA%\Microsoft\Excel\XLSTART" "Book.xltm" "%LogFilePath%" "%BaseDirectoryPath%" "%IsDesignModeEnabled%"
    call "%InstallerLibraryPath%" :InstallApp "EXCEL" "Sheet.xltx" "%APPDATA%\Microsoft\Excel\XLSTART" "Sheet.xltx" "%LogFilePath%" "%BaseDirectoryPath%" "%IsDesignModeEnabled%"
    call "%InstallerLibraryPath%" :InstallApp "EXCEL" "Sheet.xltm" "%APPDATA%\Microsoft\Excel\XLSTART" "Sheet.xltm" "%LogFilePath%" "%BaseDirectoryPath%" "%IsDesignModeEnabled%"
) else (
    rem --- Word templates (Normal.dotx / Normal.dotm) ---
    call "%InstallerLibraryPath%" :InstallApp "WORD" "Normal.dotx" "%APPDATA%\Microsoft\Templates" "Normal.dotx" "" "%BaseDirectoryPath%" "%IsDesignModeEnabled%" >nul 2>&1
    call "%InstallerLibraryPath%" :InstallApp "WORD" "Normal.dotm" "%APPDATA%\Microsoft\Templates" "Normal.dotm" "" "%BaseDirectoryPath%" "%IsDesignModeEnabled%" >nul 2>&1
    rem --- PowerPoint templates (Blank.potx / Blank.potm) ---
    call "%InstallerLibraryPath%" :InstallApp "POWERPOINT" "Blank.potx" "%APPDATA%\Microsoft\Templates" "Blank.potx" "" "%BaseDirectoryPath%" "%IsDesignModeEnabled%" >nul 2>&1
    call "%InstallerLibraryPath%" :InstallApp "POWERPOINT" "Blank.potm" "%APPDATA%\Microsoft\Templates" "Blank.potm" "" "%BaseDirectoryPath%" "%IsDesignModeEnabled%" >nul 2>&1
    rem --- Excel templates (Book / Sheet in xltx & xltm) ---
    call "%InstallerLibraryPath%" :InstallApp "EXCEL" "Book.xltx" "%APPDATA%\Microsoft\Excel\XLSTART" "Book.xltx" "" "%BaseDirectoryPath%" "%IsDesignModeEnabled%" >nul 2>&1
    call "%InstallerLibraryPath%" :InstallApp "EXCEL" "Book.xltm" "%APPDATA%\Microsoft\Excel\XLSTART" "Book.xltm" "" "%BaseDirectoryPath%" "%IsDesignModeEnabled%" >nul 2>&1
    call "%InstallerLibraryPath%" :InstallApp "EXCEL" "Sheet.xltx" "%APPDATA%\Microsoft\Excel\XLSTART" "Sheet.xltx" "" "%BaseDirectoryPath%" "%IsDesignModeEnabled%" >nul 2>&1
    call "%InstallerLibraryPath%" :InstallApp "EXCEL" "Sheet.xltm" "%APPDATA%\Microsoft\Excel\XLSTART" "Sheet.xltm" "" "%BaseDirectoryPath%" "%IsDesignModeEnabled%" >nul 2>&1
)


rem === Copy custom templates and update registry MRUs ===
if /I "%IsDesignModeEnabled%"=="true" (
    echo.
    echo [INFO] Starting custom template copy phase...
)

if not exist "%CopyLibraryPath%" (
    if /I "%IsDesignModeEnabled%"=="true" (
        echo [ERROR] Copy templates library not found: "%CopyLibraryPath%"
        echo [ERROR] Copy templates library not found: "%CopyLibraryPath%" >> "%LogFilePath%"
    )
) else (
    if /I "%IsDesignModeEnabled%"=="true" (
        echo [DEBUG] Calling copy_templates from: "%CopyLibraryPath%" >> "%LogFilePath%"
        call "%CopyLibraryPath%" :CopyAll "%LogFilePath%" "%BaseDirectoryPath%" "%RegistryLibraryPath%" "%IsDesignModeEnabled%"

    ) else (
        call "%CopyLibraryPath%" :CopyAll "%LogFilePath%" "%BaseDirectoryPath%" "%RegistryLibraryPath%" "%IsDesignModeEnabled%">nul 2>&1
    )
)

rem === Finalization and optional pause ==================
if /I "%IsDesignModeEnabled%"=="true" (
    echo [%DATE% %TIME%] --- UNIVERSAL INSTALLATION COMPLETED --- >> "%LogFilePath%"
    echo.
    echo [FINAL] Universal Office Template installation completed successfully.
    echo Log file saved at: "%LogFilePath%"
    echo ----------------------------------------------------
    pause
)
Echo Successfully executed.
pause
goto :EndOfScript

:CheckEnvironment
rem Args: LOG_FILE
set "LOG_FILE=%~1"

if defined LOG_FILE (
    echo [%DATE% %TIME%] Checking environment... >> "%LOG_FILE%"
)

openfiles >nul 2>&1
if %errorlevel% NEQ 0 (
    if defined LOG_FILE (
        echo [%DATE% %TIME%] Elevation required. Attempting to relaunch as admin... >> "%LOG_FILE%"
    )
    powershell -NoProfile -ExecutionPolicy Bypass -Command ^
        "Start-Process 'cmd.exe' -ArgumentList '/c','\"%~f0\"' -Verb RunAs"
    exit /b
)

if defined LOG_FILE (
    echo [%DATE% %TIME%] Environment check passed (already running as admin). >> "%LOG_FILE%"
)

exit /b

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

:EndOfScript
endlocal
exit /b
