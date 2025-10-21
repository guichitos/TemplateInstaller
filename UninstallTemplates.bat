@echo off
setlocal enabledelayedexpansion
chcp 65001 >nul

rem ===========================================================
rem === UNIVERSAL OFFICE TEMPLATE UNINSTALLER (v1.2) ==========
rem -----------------------------------------------------------
rem Uses the same hardcoded base paths as main_installer.bat
rem to remove Normal.dotm, Blank.potx, and Book.xltx,
rem restoring backups if available.
rem ===========================================================

rem === Mode and logging configuration ========================
if "%IsDesignModeEnabled%"=="" set "IsDesignModeEnabled=true"

set "BaseDirectoryPath=%~dp0"
set "LibraryDirectoryPath=%BaseDirectoryPath%lib"
set "LogsDirectoryPath=%BaseDirectoryPath%logs"
set "LogFilePath=%LogsDirectoryPath%\uninstall_log.txt"

if /I "%IsDesignModeEnabled%"=="true" (
    if not exist "%LogsDirectoryPath%" mkdir "%LogsDirectoryPath%"
    echo [%DATE% %TIME%] --- START UNINSTALL --- > "%LogFilePath%"
    title OFFICE TEMPLATE UNINSTALLER - DEBUG MODE
    echo [DEBUG] Running from: %BaseDirectoryPath%
)

rem === Define base template paths (same as main_installer.bat) ===
set "WORD_PATH=%APPDATA%\Microsoft\Templates"
set "PPT_PATH=%APPDATA%\Microsoft\Templates"
set "EXCEL_PATH=%APPDATA%\Microsoft\Excel\XLSTART"

echo.
echo [TARGET CLEANUP PATHS]
echo ----------------------------
echo WORD PATH:       %WORD_PATH%
echo POWERPOINT PATH: %PPT_PATH%
echo EXCEL PATH:      %EXCEL_PATH%
echo ----------------------------

if /I "%IsDesignModeEnabled%"=="true" (
    echo [INFO] --- TARGET CLEANUP PATHS --- >> "%LogFilePath%"
    echo Word path: %WORD_PATH% >> "%LogFilePath%"
    echo PowerPoint path: %PPT_PATH% >> "%LogFilePath%"
    echo Excel path: %EXCEL_PATH% >> "%LogFilePath%"
    echo ---------------------------- >> "%LogFilePath%"
)

rem === Define files ==========================================
set "WordFile=%WORD_PATH%\Normal.dotm"
set "WordBackup=%WORD_PATH%\Normal_backup.dotm"

set "PptFile=%PPT_PATH%\Blank.potx"
set "PptBackup=%PPT_PATH%\Blank_backup.potx"

set "ExcelFile=%EXCEL_PATH%\Book.xltx"
set "ExcelBackup=%EXCEL_PATH%\Book_backup.xltx"

rem === Folder existence check ================================
for %%D in ("%WORD_PATH%" "%PPT_PATH%" "%EXCEL_PATH%") do (
    if not exist "%%~D" (
        echo [WARN] Missing folder: %%~D
        if /I "%IsDesignModeEnabled%"=="true" (
            echo [WARN] Missing folder: %%~D >> "%LogFilePath%"
        )
    )
)

rem === Helper routine: delete & restore =======================
call :ProcessFile "Word" "%WordFile%" "%WordBackup%" "%LogFilePath%"
call :ProcessFile "PowerPoint" "%PptFile%" "%PptBackup%" "%LogFilePath%"
call :ProcessFile "Excel" "%ExcelFile%" "%ExcelBackup%" "%LogFilePath%"

if /I "%IsDesignModeEnabled%"=="true" (
    echo [%DATE% %TIME%] --- UNINSTALL COMPLETED --- >> "%LogFilePath%"
    echo.
    echo [FINAL] Uninstallation process finished successfully.
    echo Log saved at: "%LogFilePath%"
    echo --------------------------------------------------------
    pause
)

endlocal
exit /b


:ProcessFile
rem ===========================================================
rem Args: AppName, TargetFile, BackupFile, LogFile
rem ===========================================================
setlocal enabledelayedexpansion
set "AppName=%~1"
set "TargetFile=%~2"
set "BackupFile=%~3"
set "LogFile=%~4"

if /I "%IsDesignModeEnabled%"=="true" (
    echo.>>"%LogFile%"
    echo [INFO] Processing %AppName%...>>"%LogFile%"
)

rem === Step 1: Always delete current template (factory reset) ===
if exist "%TargetFile%" (
    del /F /Q "%TargetFile%" >nul 2>&1
    if exist "%TargetFile%" (
        echo [ERROR] Could not delete "%TargetFile%". File may be locked. >> "%LogFile%"
    ) else (
        echo [OK] Deleted "%TargetFile%" >> "%LogFile%"
    )
) else (
    echo [INFO] "%TargetFile%" not found. >> "%LogFile%"
)

rem === Step 2: Restore from backup if available ===
if exist "%BackupFile%" (
    copy /Y "%BackupFile%" "%TargetFile%" >nul 2>&1
    if exist "%TargetFile%" (
        del /F /Q "%BackupFile%" >nul 2>&1
        if exist "%BackupFile%" (
            echo [WARN] Restored "%TargetFile%" but could not delete backup. >> "%LogFile%"
        ) else (
            echo [OK] Restored "%TargetFile%" and deleted backup. >> "%LogFile%"
        )
    ) else (
        echo [ERROR] Backup copy failed for "%AppName%". >> "%LogFile%"
    )
) else (
    rem === No backup found, ensure no template remains ===
    if exist "%TargetFile%" del /F /Q "%TargetFile%" >nul 2>&1
    if not exist "%TargetFile%" (
        echo [OK] No backup found; folder left clean for "%AppName%". >> "%LogFile%"
    ) else (
        echo [ERROR] Could not clean template for "%AppName%". >> "%LogFile%"
    )
)

endlocal
exit /b 0
