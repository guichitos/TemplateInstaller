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
set "IsDesignModeEnabled=true"

rem === Base paths and library references ================
set "BaseDirectoryPath=%~dp0"
set "LibraryDirectoryPath=%BaseDirectoryPath%lib"
set "LogsDirectoryPath=%BaseDirectoryPath%logs"
set "LogFilePath=%LogsDirectoryPath%\install_log_all.txt"

rem === Initialize log only if design mode is enabled ====
if /I "%IsDesignModeEnabled%"=="true" (
    if not exist "%LogsDirectoryPath%" mkdir "%LogsDirectoryPath%"
    echo. > "%LogFilePath%"
    echo [%DATE% %TIME%] --- START TEMPLATES INSTALLATION --- >> "%LogFilePath%"
)

rem === Library references ===============================
set "CoreLibraryPath=%LibraryDirectoryPath%\core.bat"
set "InstallerLibraryPath=%LibraryDirectoryPath%\installer_apps.bat"
set "CopyLibraryPath=%LibraryDirectoryPath%\copy_templates.bat"
set "RegistryLibraryPath=%LibraryDirectoryPath%\registry_tools.bat"
set "EnvironmentLibraryPath=%LibraryDirectoryPath%\env_check.bat"

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
    call "%EnvironmentLibraryPath%" :CheckEnvironment "%LogFilePath%"
    call "%CoreLibraryPath%" :CloseOfficeApps "%LogFilePath%"
    echo [OK] Environment verification and Office app closure completed.
    echo [OK] Environment verification and Office app closure completed. >> "%LogFilePath%"
) else (
    call "%EnvironmentLibraryPath%" :CheckEnvironment >nul 2>&1
    call "%CoreLibraryPath%" :CloseOfficeApps >nul 2>&1
)



rem === Install base templates for Word, PowerPoint, Excel ===
if /I "%IsDesignModeEnabled%"=="true" (
    echo.
    echo [INFO] Starting base template installation phase...
    rem --- Word templates (Normal.dotx / Normal.dotm) ---
    for %%T in ("Normal.dotx" "Normal.dotm") do (
        call "%InstallerLibraryPath%" :InstallApp "WORD" %%~T "%APPDATA%\Microsoft\Templates" %%~T "%LogFilePath%" "%BaseDirectoryPath%" "%IsDesignModeEnabled%"
    )
    rem --- PowerPoint templates (Blank.potx / Blank.potm) ---
    for %%T in ("Blank.potx" "Blank.potm") do (
        call "%InstallerLibraryPath%" :InstallApp "POWERPOINT" %%~T "%APPDATA%\Microsoft\Templates" %%~T "%LogFilePath%" "%BaseDirectoryPath%" "%IsDesignModeEnabled%"
    )
    rem --- Excel templates (Book / Sheet in xltx & xltm) ---
    for %%T in ("Book.xltx" "Book.xltm" "Sheet.xltx" "Sheet.xltm") do (
        call "%InstallerLibraryPath%" :InstallApp "EXCEL" %%~T "%APPDATA%\Microsoft\Excel\XLSTART" %%~T "%LogFilePath%" "%BaseDirectoryPath%" "%IsDesignModeEnabled%"
    )
) else (
    rem --- Word templates (Normal.dotx / Normal.dotm) ---
    for %%T in ("Normal.dotx" "Normal.dotm") do (
        call "%InstallerLibraryPath%" :InstallApp "WORD" %%~T "%APPDATA%\Microsoft\Templates" %%~T "" "%BaseDirectoryPath%" "%IsDesignModeEnabled%" >nul 2>&1
    )
    rem --- PowerPoint templates (Blank.potx / Blank.potm) ---
    for %%T in ("Blank.potx" "Blank.potm") do (
        call "%InstallerLibraryPath%" :InstallApp "POWERPOINT" %%~T "%APPDATA%\Microsoft\Templates" %%~T "" "%BaseDirectoryPath%" "%IsDesignModeEnabled%" >nul 2>&1
    )
    rem --- Excel templates (Book / Sheet in xltx & xltm) ---
    for %%T in ("Book.xltx" "Book.xltm" "Sheet.xltx" "Sheet.xltm") do (
        call "%InstallerLibraryPath%" :InstallApp "EXCEL" %%~T "%APPDATA%\Microsoft\Excel\XLSTART" %%~T "" "%BaseDirectoryPath%" "%IsDesignModeEnabled%" >nul 2>&1
    )
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

endlocal
exit /b
