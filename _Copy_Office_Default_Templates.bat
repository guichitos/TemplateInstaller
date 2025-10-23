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
    call :InstallApp "WORD" "Normal.dotx" "%APPDATA%\Microsoft\Templates" "Normal.dotx" "%LogFilePath%" "%BaseDirectoryPath%" "%IsDesignModeEnabled%"
    call :InstallApp "WORD" "Normal.dotm" "%APPDATA%\Microsoft\Templates" "Normal.dotm" "%LogFilePath%" "%BaseDirectoryPath%" "%IsDesignModeEnabled%"
    rem --- PowerPoint templates (Blank.potx / Blank.potm) ---
    call :InstallApp "POWERPOINT" "Blank.potx" "%APPDATA%\Microsoft\Templates" "Blank.potx" "%LogFilePath%" "%BaseDirectoryPath%" "%IsDesignModeEnabled%"
    call :InstallApp "POWERPOINT" "Blank.potm" "%APPDATA%\Microsoft\Templates" "Blank.potm" "%LogFilePath%" "%BaseDirectoryPath%" "%IsDesignModeEnabled%"
    rem --- Excel templates (Book / Sheet in xltx & xltm) ---
    call :InstallApp "EXCEL" "Book.xltx" "%APPDATA%\Microsoft\Excel\XLSTART" "Book.xltx" "%LogFilePath%" "%BaseDirectoryPath%" "%IsDesignModeEnabled%"
    call :InstallApp "EXCEL" "Book.xltm" "%APPDATA%\Microsoft\Excel\XLSTART" "Book.xltm" "%LogFilePath%" "%BaseDirectoryPath%" "%IsDesignModeEnabled%"
    call :InstallApp "EXCEL" "Sheet.xltx" "%APPDATA%\Microsoft\Excel\XLSTART" "Sheet.xltx" "%LogFilePath%" "%BaseDirectoryPath%" "%IsDesignModeEnabled%"
    call :InstallApp "EXCEL" "Sheet.xltm" "%APPDATA%\Microsoft\Excel\XLSTART" "Sheet.xltm" "%LogFilePath%" "%BaseDirectoryPath%" "%IsDesignModeEnabled%"
) else (
    rem --- Word templates (Normal.dotx / Normal.dotm) ---
    call :InstallApp "WORD" "Normal.dotx" "%APPDATA%\Microsoft\Templates" "Normal.dotx" "" "%BaseDirectoryPath%" "%IsDesignModeEnabled%" >nul 2>&1
    call :InstallApp "WORD" "Normal.dotm" "%APPDATA%\Microsoft\Templates" "Normal.dotm" "" "%BaseDirectoryPath%" "%IsDesignModeEnabled%" >nul 2>&1
    rem --- PowerPoint templates (Blank.potx / Blank.potm) ---
    call :InstallApp "POWERPOINT" "Blank.potx" "%APPDATA%\Microsoft\Templates" "Blank.potx" "" "%BaseDirectoryPath%" "%IsDesignModeEnabled%" >nul 2>&1
    call :InstallApp "POWERPOINT" "Blank.potm" "%APPDATA%\Microsoft\Templates" "Blank.potm" "" "%BaseDirectoryPath%" "%IsDesignModeEnabled%" >nul 2>&1
    rem --- Excel templates (Book / Sheet in xltx & xltm) ---
    call :InstallApp "EXCEL" "Book.xltx" "%APPDATA%\Microsoft\Excel\XLSTART" "Book.xltx" "" "%BaseDirectoryPath%" "%IsDesignModeEnabled%" >nul 2>&1
    call :InstallApp "EXCEL" "Book.xltm" "%APPDATA%\Microsoft\Excel\XLSTART" "Book.xltm" "" "%BaseDirectoryPath%" "%IsDesignModeEnabled%" >nul 2>&1
    call :InstallApp "EXCEL" "Sheet.xltx" "%APPDATA%\Microsoft\Excel\XLSTART" "Sheet.xltx" "" "%BaseDirectoryPath%" "%IsDesignModeEnabled%" >nul 2>&1
    call :InstallApp "EXCEL" "Sheet.xltm" "%APPDATA%\Microsoft\Excel\XLSTART" "Sheet.xltm" "" "%BaseDirectoryPath%" "%IsDesignModeEnabled%" >nul 2>&1
)


rem === Detect Office personal template directories ==========
if /I "%IsDesignModeEnabled%"=="true" (
    echo.
    echo [INFO] Detecting Office personal template folders...
)

if /I "%IsDesignModeEnabled%"=="true" (
    call :DetectOfficePaths "%LogFilePath%" "%IsDesignModeEnabled%"
) else (
    call :DetectOfficePaths "" "%IsDesignModeEnabled%" >nul 2>&1
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

:InstallApp
rem Args: APP SRC_NAME DST_DIR DST_NAME LOG_FILE BASE_DIR DESIGN_MODE
setlocal
set "AppName=%~1"
set "SourceFileName=%~2"
set "DestinationDirectory=%~3"
set "DestinationFileName=%~4"
set "LogFilePath=%~5"
set "SourceDirectory=%~6"
set "DesignMode=%~7"

set "SourceFilePath=%SourceDirectory%%SourceFileName%"
set "DestinationFilePath=%DestinationDirectory%\%DestinationFileName%"
set "BackupFilePath=%DestinationDirectory%\%~n4_backup%~x4"

if not exist "%SourceFilePath%" (
    if /I "%DesignMode%"=="true" echo [ERROR] Source file not found: "%SourceFilePath%"
    if /I "%DesignMode%"=="true" if defined LogFilePath echo [ERROR] Source file not found "%SourceFilePath%". >> "%LogFilePath%"
    endlocal
    exit /b
)

if not exist "%DestinationDirectory%" mkdir "%DestinationDirectory%" 2>nul

if exist "%DestinationFilePath%" (
    copy /Y "%DestinationFilePath%" "%BackupFilePath%" >nul 2>&1
    if /I "%DesignMode%"=="true" (
        echo [BACKUP] Created for %AppName% template at "%BackupFilePath%"
        if defined LogFilePath echo [BACKUP] Created for %AppName% template at "%BackupFilePath%" >> "%LogFilePath%"
    )
)

copy /Y "%SourceFilePath%" "%DestinationFilePath%" >nul 2>&1
if exist "%DestinationFilePath%" (
    if /I "%DesignMode%"=="true" (
        echo [OK] Installed %AppName% template at "%DestinationFilePath%"
        if defined LogFilePath echo [OK] Installed %AppName% template at "%DestinationFilePath%" >> "%LogFilePath%"
    )
) else (
    if /I "%DesignMode%"=="true" (
        echo [ERROR] Copy failed for "%SourceFilePath%"
        if defined LogFilePath echo [ERROR] Copy failed for "%SourceFilePath%" >> "%LogFilePath%"
    )
)

endlocal
exit /b

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

:DetectOfficePaths
rem Args: LOG_FILE, DESIGN_MODE
setlocal enabledelayedexpansion
set "LOG_FILE=%~1"
set "DESIGN_MODE=%~2"

set "WORD_PATH="
set "PPT_PATH="
set "EXCEL_PATH="

for /f "tokens=1,2,*" %%A in (
  'reg query "HKCU\Software\Microsoft\Office\16.0\Word\Options" /v "PersonalTemplates" 2^>nul ^| find /I "PersonalTemplates"'
) do set "WORD_PATH=%%C"

for /f "tokens=1,2,*" %%A in (
  'reg query "HKCU\Software\Microsoft\Office\16.0\PowerPoint\Options" /v "PersonalTemplates" 2^>nul ^| find /I "PersonalTemplates"'
) do set "PPT_PATH=%%C"

for /f "tokens=1,2,*" %%A in (
  'reg query "HKCU\Software\Microsoft\Office\16.0\Excel\Options" /v "PersonalTemplates" 2^>nul ^| find /I "PersonalTemplates"'
) do set "EXCEL_PATH=%%C"

call :CleanPath WORD_PATH
call :CleanPath PPT_PATH
call :CleanPath EXCEL_PATH

if /I "!DESIGN_MODE!"=="true" (
    echo [DEBUG] Word templates folder detected: !WORD_PATH!
    echo [DEBUG] PowerPoint templates folder detected: !PPT_PATH!
    echo [DEBUG] Excel templates folder detected: !EXCEL_PATH!
)

if defined LOG_FILE (
    if defined WORD_PATH  echo [%DATE% %TIME%] Word templates folder detected: !WORD_PATH! >> "!LOG_FILE!"
    if defined PPT_PATH   echo [%DATE% %TIME%] PowerPoint templates folder detected: !PPT_PATH! >> "!LOG_FILE!"
    if defined EXCEL_PATH echo [%DATE% %TIME%] Excel templates folder detected: !EXCEL_PATH! >> "!LOG_FILE!"
)

endlocal & (
    if defined WORD_PATH  set "WORD_PATH=%WORD_PATH%"
    if defined PPT_PATH   set "PPT_PATH=%PPT_PATH%"
    if defined EXCEL_PATH set "EXCEL_PATH=%EXCEL_PATH%"
)
exit /b

:CleanPath
setlocal enabledelayedexpansion
set "VAR_NAME=%~1"
for /f "tokens=2 delims==" %%A in ('set !VAR_NAME! 2^>nul') do set "VALUE=%%A"

if defined VALUE (
    set "VALUE=!VALUE:"=!"
    for /f "tokens=* delims= " %%Z in ('echo(!VALUE!') do set "VALUE=%%Z"
    set "VALUE=!VALUE:        =!"
    if "!VALUE:~0,1!"=="\" set "VALUE=C:!VALUE!"
    if "!VALUE:~-1!"=="\" set "VALUE=!VALUE:~0,-1!"
)

endlocal & if defined VALUE set "%~1=%VALUE%"
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
