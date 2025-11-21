@echo off
setlocal enabledelayedexpansion
chcp 65001 >nul

rem ===========================================================
rem === UNIVERSAL OFFICE TEMPLATE UNINSTALLER (v1.2) ==========
rem ===========================================================

rem If wrapper passed the launcher directory (payload), use it.
if not "%~1"=="" (
    set "LauncherDirectory=%~1"
) else (
    rem Fallback: assume current directory is the launcher/payload location
    set "LauncherDirectory=%CD%"
)

rem ScriptDirectory = real location of this uninstaller (in AppData)
set "ScriptDirectory=%~dp0"

if /I "%IsDesignModeEnabled%"=="true" (
    call :DebugTrace "[INFO] Script directory (uninstaller) resolved to: %ScriptDirectory%"
    call :DebugTrace "[INFO] Launcher/payload directory resolved to: %LauncherDirectory%"
)

rem === Mode and logging configuration ========================
rem true  = verbose mode with console messages, logging, and final pause.
rem false = silent mode (no console output or pause).
set "IsDesignModeEnabled=true"

call :DebugTrace "[FLAG] Script initialization started."

set "UserLaunchDirectory=%CD%"

rem Usamos la carpeta del launcher para resolver la payload real
call :ResolveBaseDirectory "%LauncherDirectory%" BaseDirectoryPath
call :ResolveBaseDirectory "%UserLaunchDirectory%" LaunchDirectoryPath

set "BaseHasPayload=0"
set "LaunchHasPayload=0"

call :HasTemplatePayload "%BaseDirectoryPath%" BaseHasPayload
if /I not "%LaunchDirectoryPath%"=="%BaseDirectoryPath%" call :HasTemplatePayload "%LaunchDirectoryPath%" LaunchHasPayload

if "!BaseHasPayload!"=="0" if "!LaunchHasPayload!"=="1" (
    set "BaseDirectoryPath=!LaunchDirectoryPath!"
    if /I "%IsDesignModeEnabled%"=="true" call :DebugTrace "[INFO] No payload found at primary path; using launch directory payload location instead."
)

rem OJO: aquí ya volvemos a usar ScriptDirectory (AppData) para libs y logs
set "LibraryDirectoryPath=%ScriptDirectory%lib"
set "LogsDirectoryPath=%ScriptDirectory%logs"
set "LogFilePath=%LogsDirectoryPath%\uninstall_log.txt"
set "OfficeTemplateLib=%ScriptDirectory%1-2. ResolveAppProperties.bat"

if not exist "%OfficeTemplateLib%" (
    echo [ERROR] Shared library not found: "%OfficeTemplateLib%"
    exit /b 1
)

if /I "%IsDesignModeEnabled%"=="true" (
    if not exist "%LogsDirectoryPath%" mkdir "%LogsDirectoryPath%"
    echo [%DATE% %TIME%] --- START UNINSTALL --- > "%LogFilePath%"
    title OFFICE TEMPLATE UNINSTALLER - DEBUG MODE
    echo [DEBUG] Running from payload base: %BaseDirectoryPath%
)

call :DebugTrace "[FLAG] Target paths and logging configured."


rem === Define base template paths (same as main_installer.bat) ===
set "WORD_PATH=%APPDATA%\Microsoft\Templates"
set "PPT_PATH=%APPDATA%\Microsoft\Templates"
set "EXCEL_PATH=%APPDATA%\Microsoft\Excel\XLSTART"

rem Detect the Document Themes folder using the same logic as the installer
set "APPDATA_EXPANDED="
set "THEME_PATH="
for /f "delims=" %%T in ('powershell -NoLogo -Command "$app=(Get-ItemProperty -Path \"HKCU:\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\User Shell Folders\" -Name AppData -ErrorAction SilentlyContinue).AppData; if ($app) {[Environment]::ExpandEnvironmentVariables($app)}"') do set "APPDATA_EXPANDED=%%T"
if not defined APPDATA_EXPANDED set "APPDATA_EXPANDED=%APPDATA%"
if defined APPDATA_EXPANDED set "THEME_PATH=!APPDATA_EXPANDED!\Microsoft\Templates\Document Themes"

if /I "%IsDesignModeEnabled%"=="true" (
    echo.
    echo [TARGET CLEANUP PATHS]
    echo ----------------------------
    echo WORD PATH:       %WORD_PATH%
    echo POWERPOINT PATH: %PPT_PATH%
    echo EXCEL PATH:      %EXCEL_PATH%
    echo THEMES PATH:     !THEME_PATH!
    echo ----------------------------
)

if /I "%IsDesignModeEnabled%"=="true" (
    echo [INFO] --- TARGET CLEANUP PATHS --- >> "%LogFilePath%"
    echo Word path: %WORD_PATH% >> "%LogFilePath%"
    echo PowerPoint path: %PPT_PATH% >> "%LogFilePath%"
    echo Excel path: %EXCEL_PATH% >> "%LogFilePath%"
    echo Themes path: !THEME_PATH! >> "%LogFilePath%"
    echo ---------------------------- >> "%LogFilePath%"
)

rem === Detect custom template folders for optional cleanup ===
set "WORD_CUSTOM_TEMPLATE_PATH="
set "PPT_CUSTOM_TEMPLATE_PATH="
set "EXCEL_CUSTOM_TEMPLATE_PATH="
set "DEFAULT_CUSTOM_TEMPLATE_DIR="
call :DetectCustomTemplatePaths "%LogFilePath%" "%IsDesignModeEnabled%"

if /I "%IsDesignModeEnabled%"=="true" (
    call :DebugTrace "[DEBUG] Custom template cleanup targets:"
    if defined WORD_CUSTOM_TEMPLATE_PATH (
        call :DebugTrace "        Word: !WORD_CUSTOM_TEMPLATE_PATH!"
    ) else (
        call :DebugTrace "        Word: <not detected>"
    )
    if defined PPT_CUSTOM_TEMPLATE_PATH (
        call :DebugTrace "        PowerPoint: !PPT_CUSTOM_TEMPLATE_PATH!"
    ) else (
        call :DebugTrace "        PowerPoint: <not detected>"
    )
    if defined EXCEL_CUSTOM_TEMPLATE_PATH (
        call :DebugTrace "        Excel: !EXCEL_CUSTOM_TEMPLATE_PATH!"
    ) else (
        call :DebugTrace "        Excel: <not detected>"
    )
)

call :DebugTrace "[FLAG] Built-in template definitions resolved."

rem === Define files ==========================================
set "WordFile=%WORD_PATH%\Normal.dotx"
set "WordBackup=%WORD_PATH%\Normal_backup.dotx"
set "WordMacroFile=%WORD_PATH%\Normal.dotm"
set "WordMacroBackup=%WORD_PATH%\Normal_backup.dotm"
set "WordEmailFile=%WORD_PATH%\NormalEmail.dotx"
set "WordEmailBackup=%WORD_PATH%\NormalEmail_backup.dotx"
set "WordEmailMacroFile=%WORD_PATH%\NormalEmail.dotm"
set "WordEmailMacroBackup=%WORD_PATH%\NormalEmail_backup.dotm"

set "PptFile=%PPT_PATH%\Blank.potx"
set "PptBackup=%PPT_PATH%\Blank_backup.potx"
set "PptMacroFile=%PPT_PATH%\Blank.potm"
set "PptMacroBackup=%PPT_PATH%\Blank_backup.potm"

set "ExcelBookFile=%EXCEL_PATH%\Book.xltx"
set "ExcelBookBackup=%EXCEL_PATH%\Book_backup.xltx"
set "ExcelBookMacroFile=%EXCEL_PATH%\Book.xltm"
set "ExcelBookMacroBackup=%EXCEL_PATH%\Book_backup.xltm"

set "ExcelSheetFile=%EXCEL_PATH%\Sheet.xltx"
set "ExcelSheetBackup=%EXCEL_PATH%\Sheet_backup.xltx"
set "ExcelSheetMacroFile=%EXCEL_PATH%\Sheet.xltm"
set "ExcelSheetMacroBackup=%EXCEL_PATH%\Sheet_backup.xltm"

rem === Helper routine: delete & restore =======================
call :ProcessFile "Word (.dotx)" "%WordFile%" "%WordBackup%" "%LogFilePath%"
call :ProcessFile "Word (.dotm)" "%WordMacroFile%" "%WordMacroBackup%" "%LogFilePath%"
call :ProcessFile "Word Email (.dotx)" "%WordEmailFile%" "%WordEmailBackup%" "%LogFilePath%"
call :ProcessFile "Word Email (.dotm)" "%WordEmailMacroFile%" "%WordEmailMacroBackup%" "%LogFilePath%"
call :ProcessFile "PowerPoint (.potx)" "%PptFile%" "%PptBackup%" "%LogFilePath%"
call :ProcessFile "PowerPoint (.potm)" "%PptMacroFile%" "%PptMacroBackup%" "%LogFilePath%"
call :ProcessFile "Excel Book (.xltx)" "%ExcelBookFile%" "%ExcelBookBackup%" "%LogFilePath%"
call :ProcessFile "Excel Book (.xltm)" "%ExcelBookMacroFile%" "%ExcelBookMacroBackup%" "%LogFilePath%"
call :ProcessFile "Excel Sheet (.xltx)" "%ExcelSheetFile%" "%ExcelSheetBackup%" "%LogFilePath%"
call :ProcessFile "Excel Sheet (.xltm)" "%ExcelSheetMacroFile%" "%ExcelSheetMacroBackup%" "%LogFilePath%"

set "THEME_PAYLOAD_TRACK="
if defined THEME_PATH (
    for %%F in ("%BaseDirectoryPath%*.thmx") do (
        if exist "%%~fF" set "THEME_PAYLOAD_TRACK=!THEME_PAYLOAD_TRACK!;%%~nxF;"
    )
)

rem Clean Document Themes by comparing against installer payloads and only delete matches
if defined THEME_PATH if exist "!THEME_PATH!" (
    for /f "delims=" %%T in ('dir /A-D /B "!THEME_PATH!\*.thmx" 2^>nul') do (
        set "THEME_HAS_PAYLOAD=0"
        if defined THEME_PAYLOAD_TRACK (
            echo !THEME_PAYLOAD_TRACK! | find /I ";%%~nT%%~xT;" >nul && set "THEME_HAS_PAYLOAD=1"
        )

        if "!THEME_HAS_PAYLOAD!"=="1" (
            set "CurrentThemeFile=!THEME_PATH!\%%~nxT"
            set "CurrentThemeBackup=!THEME_PATH!\%%~nT_backup%%~xT"
            call :ProcessFile "Office Theme (%%~nxT)" "!CurrentThemeFile!" "!CurrentThemeBackup!" "%LogFilePath%"
        ) else (
            if /I "%IsDesignModeEnabled%"=="true" call :DebugTrace "        [SKIP] Preserved Office Theme (%%~nxT) with no installer match."
        )
    )
)

call :DebugTrace "[FLAG] Starting custom template cleanup."

call :RemoveCustomTemplates "%BaseDirectoryPath%" "%LogFilePath%" "%IsDesignModeEnabled%" "!WORD_CUSTOM_TEMPLATE_PATH!" "!PPT_CUSTOM_TEMPLATE_PATH!" "!EXCEL_CUSTOM_TEMPLATE_PATH!"

echo.
call :DebugTrace "[FLAG] Clearing template MRU entries."

call :ResetOfficeTemplateMruLists "%LogFilePath%" "%IsDesignModeEnabled%"

call :DebugTrace "[FLAG] Finalizing uninstaller."

call :Finalize "%LogFilePath%"

endlocal
exit /b

rem Base dir resolver keeps template source tied to the unpacked executable location
:ResolveBaseDirectory
setlocal
set "RBD_INPUT=%~1"
set "RBD_OUTPUT_VAR=%~2"

if "%RBD_INPUT:~-1%" NEQ "\\" set "RBD_INPUT=%RBD_INPUT%\\"

set "RBD_FOUND="
for %%D in ("%RBD_INPUT%" "%RBD_INPUT%payload\\" "%RBD_INPUT%templates\\" "%RBD_INPUT%extracted\\") do (
    for %%F in ("%%~D*.dot*" "%%~D*.pot*" "%%~D*.xlt*" "%%~D*.thmx") do (
        if exist "%%~fF" set "RBD_FOUND=%%~D"
    )
    if defined RBD_FOUND goto :RBD_Found
)

:RBD_Found
if not defined RBD_FOUND set "RBD_FOUND=%RBD_INPUT%"

endlocal & set "%RBD_OUTPUT_VAR%=%RBD_FOUND%"
exit /b 0

:HasTemplatePayload
setlocal enabledelayedexpansion
set "HP_PATH=%~1"
set "HP_OUT=%~2"
set "HP_FOUND=0"

if not defined HP_PATH goto :HasTemplatePayloadEnd
if "!HP_PATH:~-1!" NEQ "\\" set "HP_PATH=!HP_PATH!\\"

for %%F in ("!HP_PATH!*.dot*" "!HP_PATH!*.pot*" "!HP_PATH!*.xlt*" "!HP_PATH!*.thmx") do (
    if exist "%%~fF" set "HP_FOUND=1"
)

:HasTemplatePayloadEnd
set "HP_RESULT=!HP_FOUND!"
endlocal & if not "%HP_OUT%"=="" set "%HP_OUT%=%HP_RESULT%"
exit /b 0



:DetectCustomTemplatePaths
set "DCTP_LOG_FILE=%~1"
set "DCTP_DESIGN_MODE=%~2"
if not defined DCTP_DESIGN_MODE set "DCTP_DESIGN_MODE=%IsDesignModeEnabled%"
set "WORD_CUSTOM_TEMPLATE_PATH="
set "PPT_CUSTOM_TEMPLATE_PATH="
set "EXCEL_CUSTOM_TEMPLATE_PATH="
set "DEFAULT_CUSTOM_TEMPLATE_DIR="
set "DEFAULT_CUSTOM_DIR_STATUS=unknown"
set "DCTP_DOCUMENTS_PATH="
set "DCTP_OFFICE_VERSIONS=16.0 15.0 14.0 12.0"

for /f "delims=" %%D in ('powershell -NoLogo -Command "$path=(Get-ItemProperty -Path \"HKCU:\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\User Shell Folders\" -Name Personal -ErrorAction SilentlyContinue).Personal; if ($path) {[Environment]::ExpandEnvironmentVariables($path)}"') do set "DCTP_DOCUMENTS_PATH=%%D"

if defined DCTP_DOCUMENTS_PATH (
    if "!DCTP_DOCUMENTS_PATH:~-1!"=="\" set "DCTP_DOCUMENTS_PATH=!DCTP_DOCUMENTS_PATH:~0,-1!"
    set "DEFAULT_CUSTOM_TEMPLATE_DIR=!DCTP_DOCUMENTS_PATH!\Custom Templates"
) else (
    set "DEFAULT_CUSTOM_TEMPLATE_DIR=%USERPROFILE%\Documents\Custom Templates"
)

if not defined DEFAULT_CUSTOM_TEMPLATE_DIR set "DEFAULT_CUSTOM_TEMPLATE_DIR=%USERPROFILE%\Documents\Custom Templates"

if defined DEFAULT_CUSTOM_TEMPLATE_DIR (
    if exist "!DEFAULT_CUSTOM_TEMPLATE_DIR!" (
        set "DEFAULT_CUSTOM_DIR_STATUS=exists"
    ) else (
        mkdir "!DEFAULT_CUSTOM_TEMPLATE_DIR!" >nul 2>&1
        if not errorlevel 1 (
            set "DEFAULT_CUSTOM_DIR_STATUS=created"
        ) else (
            set "DEFAULT_CUSTOM_DIR_STATUS=create_failed"
        )
    )
)

for %%V in (!DCTP_OFFICE_VERSIONS!) do (
    if not defined WORD_CUSTOM_TEMPLATE_PATH (
        for /f "tokens=1,2,*" %%A in (
          'reg query "HKCU\Software\Microsoft\Office\%%V\Word\Options" /v "PersonalTemplates" 2^>nul ^| find /I "PersonalTemplates"'
        ) do set "WORD_CUSTOM_TEMPLATE_PATH=%%C"
    )
    if not defined PPT_CUSTOM_TEMPLATE_PATH (
        for /f "tokens=1,2,*" %%A in (
          'reg query "HKCU\Software\Microsoft\Office\%%V\PowerPoint\Options" /v "PersonalTemplates" 2^>nul ^| find /I "PersonalTemplates"'
        ) do set "PPT_CUSTOM_TEMPLATE_PATH=%%C"
    )
    if not defined EXCEL_CUSTOM_TEMPLATE_PATH (
        for /f "tokens=1,2,*" %%A in (
          'reg query "HKCU\Software\Microsoft\Office\%%V\Excel\Options" /v "PersonalTemplates" 2^>nul ^| find /I "PersonalTemplates"'
        ) do set "EXCEL_CUSTOM_TEMPLATE_PATH=%%C"
    )
)

for %%V in (!DCTP_OFFICE_VERSIONS!) do (
    if not defined WORD_CUSTOM_TEMPLATE_PATH (
        for /f "tokens=1,2,*" %%A in (
          'reg query "HKCU\Software\Microsoft\Office\%%V\Common\General" /v "UserTemplates" 2^>nul ^| find /I "UserTemplates"'
        ) do set "WORD_CUSTOM_TEMPLATE_PATH=%%C"
    )
    if not defined PPT_CUSTOM_TEMPLATE_PATH (
        for /f "tokens=1,2,*" %%A in (
          'reg query "HKCU\Software\Microsoft\Office\%%V\Common\General" /v "UserTemplates" 2^>nul ^| find /I "UserTemplates"'
        ) do set "PPT_CUSTOM_TEMPLATE_PATH=%%C"
    )
    if not defined EXCEL_CUSTOM_TEMPLATE_PATH (
        for /f "tokens=1,2,*" %%A in (
          'reg query "HKCU\Software\Microsoft\Office\%%V\Common\General" /v "UserTemplates" 2^>nul ^| find /I "UserTemplates"'
        ) do set "EXCEL_CUSTOM_TEMPLATE_PATH=%%C"
    )
)

if not defined WORD_CUSTOM_TEMPLATE_PATH set "WORD_CUSTOM_TEMPLATE_PATH=!DEFAULT_CUSTOM_TEMPLATE_DIR!"
if not defined PPT_CUSTOM_TEMPLATE_PATH set "PPT_CUSTOM_TEMPLATE_PATH=!DEFAULT_CUSTOM_TEMPLATE_DIR!"
if not defined EXCEL_CUSTOM_TEMPLATE_PATH set "EXCEL_CUSTOM_TEMPLATE_PATH=!DEFAULT_CUSTOM_TEMPLATE_DIR!"

call :CleanPath WORD_CUSTOM_TEMPLATE_PATH
call :CleanPath PPT_CUSTOM_TEMPLATE_PATH
call :CleanPath EXCEL_CUSTOM_TEMPLATE_PATH

exit /b 0

:CleanPath
call "%OfficeTemplateLib%" :CleanPath %*
exit /b %errorlevel%

:RemoveCustomTemplates
setlocal enabledelayedexpansion
set "BASE_DIR=%~1"
set "LOG_FILE=%~2"
set "DESIGN_MODE=%~3"
set "WORD_DIR=%~4"
set "PPT_DIR=%~5"
set "EXCEL_DIR=%~6"

if not defined BASE_DIR exit /b 0
if "!BASE_DIR:~-1!" NEQ "\" set "BASE_DIR=!BASE_DIR!\"

if /I "!DESIGN_MODE!"=="true" (
    call :DebugTrace "        [DEBUG] RemoveCustomTemplates invoked with:"
    call :DebugTrace "        Base dir: !BASE_DIR!"
    call :DebugTrace "        Word dir: !WORD_DIR!"
    call :DebugTrace "        PPT dir: !PPT_DIR!"
    call :DebugTrace "        Excel dir: !EXCEL_DIR!"
)

set /a CUSTOM_REMOVED_COUNT=0
set /a CUSTOM_SKIP_COUNT=0
set /a CUSTOM_ERROR_COUNT=0
set /a CUSTOM_TOTAL_CANDIDATES=0
set "CUSTOM_GENERIC_SKIP_LIST=Normal.dotx NormalEmail.dotx Blank.potx Book.xltx Normal.dotm NormalEmail.dotm Blank.potm Book.xltm Sheet.xltx Sheet.xltm"

call :CleanCustomTemplateFiles "!WORD_DIR!" ".dotx .dotm" "!BASE_DIR!" "%LOG_FILE%" "!DESIGN_MODE!" "Word custom templates"
call :CleanCustomTemplateFiles "!PPT_DIR!" ".potx .potm" "!BASE_DIR!" "%LOG_FILE%" "!DESIGN_MODE!" "PowerPoint custom templates"
call :CleanCustomTemplateFiles "!EXCEL_DIR!" ".xltx .xltm" "!BASE_DIR!" "%LOG_FILE%" "!DESIGN_MODE!" "Excel custom templates"

if /I "!DESIGN_MODE!"=="true" (
    call :DebugTrace "[INFO] Custom template cleanup summary: Removed !CUSTOM_REMOVED_COUNT!, skipped !CUSTOM_SKIP_COUNT!, errors !CUSTOM_ERROR_COUNT!."
)

endlocal
exit /b 0

:CleanCustomTemplateFiles
echo [FLAG] CleanCustomTemplateFiles invoked with parameters %*
set "CCF_TARGET_DIR=%~1"
set "CCF_EXT_LIST=%~2"
set "CCF_BASE_DIR=%~3"
call :NormalizePath CCF_BASE_DIR
set "CCF_LOG_FILE=%~4"
set "CCF_DESIGN_MODE=%~5"
set "CCF_LABEL=%~6"

if not defined CCF_TARGET_DIR exit /b 0
if "!CCF_TARGET_DIR!"=="" exit /b 0
if not exist "!CCF_TARGET_DIR!" (
    if /I "!CCF_DESIGN_MODE!"=="true" call :DebugTrace "[INFO] !CCF_LABEL! not found at '!CCF_TARGET_DIR!' - skipping."
    exit /b 0
)

set "CCF_TOP_LEVEL_COUNT=0"
set "CCF_RECURSIVE_COUNT=0"

for /f %%C in ('dir /A /B "!CCF_TARGET_DIR!" 2^>nul ^| find /C /V ""') do set "CCF_TOP_LEVEL_COUNT=%%C"
for /f %%C in ('dir /A /B /S "!CCF_TARGET_DIR!" 2^>nul ^| find /C /V ""') do set "CCF_RECURSIVE_COUNT=%%C"

    set "CCF_DIR_FILE_COUNT=0"
    set "CCF_DIR_REMOVED=0"
    set "CCF_DIR_SKIPPED=0"
    set "CCF_DIR_ERRORS=0"

    for %%E in (!CCF_EXT_LIST!) do (
        set "CCF_EXT_COUNT=0"
        set "CCF_EXT_REMOVED=0"
        set "CCF_EXT_SKIPPED=0"
        set "CCF_EXT_ERRORS=0"
        for /f %%C in ('dir /A-D /B /S "!CCF_TARGET_DIR!\*%%~E" 2^>nul ^| find /C /V ""') do set "CCF_EXT_COUNT=%%C"

        for /f "delims=" %%F in ('dir /A-D /B /S "!CCF_TARGET_DIR!\*%%~E" 2^>nul') do (
            if exist "%%~fF" (
                set "CCF_FILE=%%~nxF"
                set /a CUSTOM_TOTAL_CANDIDATES+=1
                set /a CCF_DIR_FILE_COUNT+=1
                set "CCF_SKIP_GENERIC=0"
                for %%G in (!CUSTOM_GENERIC_SKIP_LIST!) do (
                    if /I "!CCF_FILE!"=="%%~G" set "CCF_SKIP_GENERIC=1"
                )

                if "!CCF_SKIP_GENERIC!"=="1" (
                    rem === Preserve generic system templates ===
                    set /a CUSTOM_SKIP_COUNT+=1
                    set /a CCF_DIR_SKIPPED+=1
                    set /a CCF_EXT_SKIPPED+=1
                    if /I "!CCF_DESIGN_MODE!"=="true" call :DebugTrace "[SKIP] Preserved generic template !CCF_FILE! in !CCF_LABEL!."
                ) else (
                    set "CCF_INSTALLER_FILE=!CCF_BASE_DIR!!CCF_FILE!"

                    rem === Files that ARE part of installer payload MUST be deleted ===
                    if exist "!CCF_INSTALLER_FILE!" (
                        set "CCF_DELETE_REASON=installer payload match"
                        del /F /Q "%%~fF" >nul 2>&1
                        if exist "%%~fF" (
                            set /a CUSTOM_ERROR_COUNT+=1
                            set /a CCF_DIR_ERRORS+=1
                            set /a CCF_EXT_ERRORS+=1
                            if /I "!CCF_DESIGN_MODE!"=="true" call :DebugTrace "[ERROR] Could not delete !CCF_FILE! from !CCF_LABEL!."
                        ) else (
                            set /a CUSTOM_REMOVED_COUNT+=1
                            set /a CCF_DIR_REMOVED+=1
                            set /a CCF_EXT_REMOVED+=1
                            if /I "!CCF_DESIGN_MODE!"=="true" call :DebugTrace "[OK] Deleted !CCF_FILE! from !CCF_LABEL! (!CCF_DELETE_REASON!)."
                        )
                    ) else (
                        rem === Files NOT in installer payload MUST be PRESERVED ===
                        set /a CUSTOM_SKIP_COUNT+=1
                        set /a CCF_DIR_SKIPPED+=1
                        set /a CCF_EXT_SKIPPED+=1
                        if /I "!CCF_DESIGN_MODE!"=="true" call :DebugTrace "[SKIP] Preserved !CCF_FILE! in !CCF_LABEL! (user/custom file)."
                    )
                )

            ) else (
                if /I "!CCF_DESIGN_MODE!"=="true" call :DebugTrace "[WARN] Candidate vanished before delete: %%~fF"
            )
        )
    )

    set "CCF_REMOVED_DIRS=0"
    for /f "delims=" %%D in ('dir /AD /B /S "!CCF_TARGET_DIR!" ^| sort /R') do (
        rd "%%~fD" 2>nul && set /a CCF_REMOVED_DIRS+=1
    )

    if /I "!CCF_DESIGN_MODE!"=="true" call :DebugTrace "[INFO] !CCF_LABEL!: removed !CCF_REMOVED_DIRS! empty directories."

set "CCF_TARGET_DIR="
set "CCF_EXT_LIST="
set "CCF_BASE_DIR="
set "CCF_LOG_FILE="
set "CCF_DESIGN_MODE="
set "CCF_LABEL="
set "CCF_FILE="
set "CCF_INSTALLER_FILE="
set "CCF_SKIP_GENERIC="
set "CCF_TOP_LEVEL_COUNT="
set "CCF_RECURSIVE_COUNT="
set "CCF_EXT_COUNT="
set "CCF_DIR_REMOVED="
set "CCF_DIR_SKIPPED="
set "CCF_DIR_ERRORS="
set "CCF_EXT_REMOVED="
set "CCF_EXT_SKIPPED="
set "CCF_EXT_ERRORS="
set "CCF_REMOVED_DIRS="
exit /b 0


rem --- ResetOfficeTemplateMruLists ==============================
:ResetOfficeTemplateMruLists
setlocal enabledelayedexpansion
set "ROML_LOG_FILE=%~1"
set "ROML_DESIGN_MODE=%~2"
if not defined ROML_DESIGN_MODE set "ROML_DESIGN_MODE=%IsDesignModeEnabled%"

set "ROML_OFFICE_APPS=Word PowerPoint Excel"

for %%A in (!ROML_OFFICE_APPS!) do (
    set "ROML_TARGETS="
    call :GetAppMruTargets ROML_TARGETS "%%~A"

    if defined ROML_TARGETS (
        for %%T in (!ROML_TARGETS!) do (
            set "ROML_CURRENT=%%~T"
            if not "!ROML_CURRENT!"=="" (
                call :ClearMruKey "!ROML_CURRENT!" "%%~A template MRU" "!ROML_LOG_FILE!" "!ROML_DESIGN_MODE!"
            )
        )
    )
)

endlocal
exit /b 0

:GetAppMruTargets
set "GAM_OUT=%~1"
set "GAM_APP=%~2"
if "%GAM_OUT%"=="" exit /b 0

setlocal EnableDelayedExpansion
set "APP_NAME=%GAM_APP%"
set "MRU_TARGET_PATHS="

call :ResolveAppProperties "!APP_NAME!"
echo [DEBUG] Entering GetAppMruTargets with parameters %*
echo .
echo [DEBUG] Resolved properties for !APP_NAME!: PROP_REG_NAME=!PROP_REG_NAME!
echo.
if defined PROP_REG_NAME (
    set "MRU_VAR=!PROP_MRU_VAR!"
    set "MRU_PATH="
    set "MRU_CONTAINER_PATH="
    set "MRU_CONTAINER_ID="

    call :DetectAdalContainer MRU_CONTAINER_ID MRU_CONTAINER_PATH "!PROP_REG_NAME!"


    if not errorlevel 1 if defined MRU_CONTAINER_PATH set "MRU_PATH=!MRU_CONTAINER_PATH!\File MRU"

    if not defined MRU_PATH (
        call :DetectMRUPath "!APP_NAME!"
        for /f "tokens=2 delims==" %%V in ('set !MRU_VAR! 2^>nul') do set "MRU_PATH=%%V"
    )

    if not defined MRU_PATH (
        set "MRU_PATH=HKCU\Software\Microsoft\Office\16.0\!PROP_REG_NAME!\Recent Templates\File MRU"
    )

    set "!MRU_VAR!=!MRU_PATH!"
    if defined MRU_PATH set "MRU_TARGET_PATHS=""!MRU_PATH!"""
    if defined MRU_TARGET_PATHS set "MRU_TARGET_PATHS=!MRU_TARGET_PATHS:""="!"

    set "AUTH_CONTAINER_TARGETS="
    call :CollectAuthContainerPaths AUTH_CONTAINER_TARGETS "!PROP_REG_NAME!"
    if defined AUTH_CONTAINER_TARGETS (
        set "AUTH_CONTAINER_TARGETS=!AUTH_CONTAINER_TARGETS:""="!"
        for %%C in (!AUTH_CONTAINER_TARGETS!) do (
            set "CURRENT_CONTAINER=%%~C"
            if not "!CURRENT_CONTAINER!"=="" call :AppendUniquePath MRU_TARGET_PATHS "!CURRENT_CONTAINER!\File MRU"
        )
    )

    if not defined MRU_TARGET_PATHS set "MRU_TARGET_PATHS=""!MRU_PATH!"""
    if defined MRU_TARGET_PATHS set "MRU_TARGET_PATHS=!MRU_TARGET_PATHS:""="!"
)

set "RESULT=%MRU_TARGET_PATHS%"

for %%# in (1) do (
    endlocal
    if not "%GAM_OUT%"=="" set "%GAM_OUT%=%RESULT%"
)

exit /b 0

:DetectAdalContainer
call "%OfficeTemplateLib%" :DetectAdalContainer %*
exit /b %errorlevel%

:BuildAuthContainerCache
call "%OfficeTemplateLib%" :BuildAuthContainerCache %*
exit /b %errorlevel%

:ResetAuthContainerCache
call "%OfficeTemplateLib%" :ResetAuthContainerCache %*
exit /b %errorlevel%

:ScanRecentTemplateKey
call "%OfficeTemplateLib%" :ScanRecentTemplateKey %*
exit /b %errorlevel%

:HandleRecentTemplateSubkey
call "%OfficeTemplateLib%" :HandleRecentTemplateSubkey %*
exit /b %errorlevel%

:RegisterAuthContainer
call "%OfficeTemplateLib%" :RegisterAuthContainer %*
exit /b %errorlevel%

:CollectAuthContainerPaths
call "%OfficeTemplateLib%" :CollectAuthContainerPaths %*
exit /b %errorlevel%

:AppendUniquePath
call "%OfficeTemplateLib%" :AppendUniquePath %*
exit /b %errorlevel%

:DetectMRUPath
call "%OfficeTemplateLib%" :DetectMRUPath %*
exit /b %errorlevel%

:ResolveAppProperties
echo [DEBUG] Going ResolveAppProperties en other file with parameters %*
call "%OfficeTemplateLib%" :ResolveAppProperties %*
echo [DEBUG] Returning ResolveAppProperties en other file with PROP_REG_NAME=%PROP_REG_NAME%
exit /b %errorlevel%

:ClearMruKey
setlocal DisableDelayedExpansion
set "CMK_KEY=%~1"
set "CMK_LABEL=%~2"
set "CMK_LOG_FILE=%~3"
set "CMK_DESIGN_MODE=%~4"
if not defined CMK_DESIGN_MODE set "CMK_DESIGN_MODE=%IsDesignModeEnabled%"

set "CMK_KEY=%CMK_KEY:"=%"
for /f "tokens=* delims=" %%A in ('echo %CMK_KEY%') do set "CMK_KEY=%%A"

if "%CMK_KEY%"=="" (
    endlocal
    exit /b 0
)

reg query "%CMK_KEY%" >nul 2>&1
if errorlevel 1 (
    endlocal
    exit /b 0
)

powershell -NoLogo -NoProfile -Command ^
    "$path='Registry::%CMK_KEY%';" ^
    "if (Test-Path -LiteralPath $path) {" ^
    "    $item = Get-Item -LiteralPath $path;" ^
    "    foreach ($name in $item.GetValueNames()) {" ^
    "        if ([string]::IsNullOrEmpty($name)) {" ^
    "            Set-ItemProperty -LiteralPath $path -Name '(default)' -Value $null -ErrorAction SilentlyContinue;" ^
    "        } else {" ^
    "            Remove-ItemProperty -LiteralPath $path -Name $name -ErrorAction SilentlyContinue;" ^
    "        }" ^
    "    }" ^
    "}" >nul 2>&1

if errorlevel 1 (
    rem Failed to clean – no logging
) else (
    rem Cleaned successfully – no logging
)

endlocal
exit /b 0

:ProcessFile
rem ===========================================================
rem Args: AppName, TargetFile, BackupFile, LogFile
rem ===========================================================
setlocal enabledelayedexpansion
set "AppName=%~1"
set "TargetFile=%~2"
set "BackupFile=%~3"
set "LogFile=%~4"

rem === Step 1: Always delete current template (factory reset) ===
if exist "%TargetFile%" (
    del /F /Q "%TargetFile%" >nul 2>&1
    if exist "%TargetFile%" (
        set "Message=[%AppName%] [ERROR] Could not delete %TargetFile%. File may be locked."
    ) else (
        set "Message=[%AppName%] [OK] Deleted %TargetFile%"
    )
) else (
    set "Message=[%AppName%] [INFO] %TargetFile% not found."
)

rem === Step 2: Restore from backup if available ===
if exist "%BackupFile%" (
    copy /Y "%BackupFile%" "%TargetFile%" >nul 2>&1
    if exist "%TargetFile%" (
        del /F /Q "%BackupFile%" >nul 2>&1
        if exist "%BackupFile%" (
            set "Message=[%AppName%] [WARN] Restored %TargetFile% but could not delete backup."
        ) else (
            set "Message=[%AppName%] [OK] Restored %TargetFile% and deleted backup."
        )
    ) else (
        set "Message=[%AppName%] [ERROR] Backup copy failed for %AppName%."
    )
) else (
    rem === No backup found, ensure no template remains ===
    if exist "%TargetFile%" del /F /Q "%TargetFile%" >nul 2>&1
    if not exist "%TargetFile%" (
        set "Message=[%AppName%] [OK] No backup found; folder left clean for %AppName%."
    ) else (
        set "Message=[%AppName%] [ERROR] Could not clean template for %AppName%."
    )
)

rem === Step 3: Emit verbose trace if enabled ===
if /I "%IsDesignModeEnabled%"=="true" (
    call :DebugTrace "        !Message!"
    if defined LogFile (>>"%LogFile%" echo [%DATE% %TIME%] !Message!)
)

endlocal
exit /b 0

:Finalize
setlocal enabledelayedexpansion
if /I not "%IsDesignModeEnabled%"=="true" (
    endlocal
    exit /b 0
)

set "ResolvedLogPath=%~1"

>>"%~1" echo [%DATE% %TIME%] --- UNINSTALL COMPLETED ---

endlocal
exit /b 0

:DebugTrace
if /I not "%IsDesignModeEnabled%"=="true" exit /b 0
setlocal enabledelayedexpansion
set "DebugMessage=%~1"
if defined DebugMessage (
    echo !DebugMessage!
) else (
    echo.
)
endlocal
exit /b 0

:Log
call "%OfficeTemplateLib%" :Log %*
exit /b %errorlevel%

:NormalizePath
setlocal
set "NP_VAR=%~1"
set "NP_VAL=!%NP_VAR%!"

rem Remove trailing backslashes
:NP_LOOP
if "!NP_VAL!"=="\" goto NP_END
if "!NP_VAL:~-1!"=="\" (
    set "NP_VAL=!NP_VAL:~0,-1!"
    goto NP_LOOP
)

:NP_END
rem Add exactly one backslash
set "NP_VAL=!NP_VAL!\"
endlocal & set "%~1=%NP_VAL%"
exit /b 0
