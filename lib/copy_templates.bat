rem copy_templates.bat
@echo off
chcp 65001 >nul
setlocal enabledelayedexpansion

rem ==========================================================
rem === COPY_TEMPLATES.BAT - Copy user templates and update MRU
rem ==========================================================

rem === Entry dispatcher =====================================
if "%~1"=="" exit /b
set "FUNC=%~1"
if "%FUNC:~0,1%"==":" set "FUNC=%FUNC:~1%"
goto %FUNC%


:CopyAll
rem Args: LOG_FILE BASE_DIR REGISTRY_LIB IsDesignModeEnabled
shift
set "LOG_FILE=%~1"
set "BASE_DIR=%~2"
set "REGISTRY_LIB=%~3"
set "IsDesignModeEnabled=%~4"

setlocal enabledelayedexpansion
set /a TOTAL_FILES=0
set /a TOTAL_ERRORS=0

if /I "%IsDesignModeEnabled%"=="true" (
    title COPY_TEMPLATES DEBUG MODE
    echo [DEBUG] copy_templates.bat started
    echo [INFO] Script running from: %~dp0
    echo [INFO] Arguments: %*
    echo.
)

set "REGISTRY_LIB_PPT=%~dp0lib\registry_tools_ppt.bat"
set "REGISTRY_LIB_WORD=%~dp0lib\registry_tools_word.bat"
set "REGISTRY_LIB_EXCEL=%~dp0lib\registry_tools_excel.bat"

if defined WORD_PATH if "!WORD_PATH:~-1!"=="\" set "WORD_PATH=!WORD_PATH:~0,-1!"
if defined PPT_PATH  if "!PPT_PATH:~-1!"=="\"  set "PPT_PATH=!PPT_PATH:~0,-1!"
if defined EXCEL_PATH if "!EXCEL_PATH:~-1!"=="\" set "EXCEL_PATH=!EXCEL_PATH:~0,-1!"

if /I "%IsDesignModeEnabled%"=="true" (
    echo.
    echo [DEBUG] === Template destinations received ===
    echo   WORD_PATH = !WORD_PATH!
    echo   PPT_PATH  = !PPT_PATH!
    echo   EXCEL_PATH= !EXCEL_PATH!
    echo [DEBUG] ==================================
    echo.
)

rem ==========================================================
rem === STAGE 1: DETECT MRU PATHS ============================
rem ==========================================================
if exist "%REGISTRY_LIB_PPT%" (
    call "%REGISTRY_LIB_PPT%" DetectPowerPointMRUPath
    setlocal enabledelayedexpansion
    echo [DEBUG] PowerPoint MRU detected: !PPT_MRU_PATH!
    echo [DEBUG] PowerPoint MRU detected: !PPT_MRU_PATH! >> "%LOG_FILE%"
    endlocal
) else if /I "%IsDesignModeEnabled%"=="true" (
    echo [WARNING] registry_tools_ppt.bat not found. MRU for PowerPoint skipped.
    echo [WARNING] registry_tools_ppt.bat not found >> "%LOG_FILE%"
)


if exist "%REGISTRY_LIB_WORD%" (
    call "%REGISTRY_LIB_WORD%" DetectWordMRUPath
    echo [DEBUG] Word MRU detected: %WORD_MRU_PATH% >> "%LOG_FILE%"
) else if /I "%IsDesignModeEnabled%"=="true" (
    echo [WARNING] registry_tools_word.bat not found. MRU for Word skipped.
    echo [WARNING] registry_tools_word.bat not found >> "%LOG_FILE%"
)

if exist "%REGISTRY_LIB_EXCEL%" (
    call "%REGISTRY_LIB_EXCEL%" DetectExcelMRUPath
    echo [DEBUG] Excel MRU detected: %EXCEL_MRU_PATH% >> "%LOG_FILE%"
) else if /I "%IsDesignModeEnabled%"=="true" (
    echo [WARNING] registry_tools_excel.bat not found. MRU for Excel skipped.
    echo [WARNING] registry_tools_excel.bat not found >> "%LOG_FILE%"
)

if /I "%IsDesignModeEnabled%"=="true" (
    echo.
)

rem ==========================================================
rem === STAGE 2: FILE LISTING AND VALIDATION ================
rem ==========================================================
if /I "%IsDesignModeEnabled%"=="true" (
    echo [DEBUG] --- Scanning BASE_DIR for templates --- >> "%LOG_FILE%"
    echo [INFO] Searching templates in "%BASE_DIR%"...
    echo -----------------------------------------------
    dir /b "%BASE_DIR%\*.dot*" "%BASE_DIR%\*.pot*" "%BASE_DIR%\*.xlt*" 2>nul
    echo -----------------------------------------------
    echo.
)

if errorlevel 1 (
    if /I "%IsDesignModeEnabled%"=="true" (
        echo [WARNING] No template files found in "%BASE_DIR%".
        echo [WARNING] No .dotx / .potx / .xltx files detected. >> "%LOG_FILE%"
    )
)

rem ==========================================================
rem === STAGE 3: DESTINATION PATH VALIDATION ================
rem ==========================================================
if /I "%IsDesignModeEnabled%"=="true" (
    echo [DEBUG] Verifying destination paths...
)
for %%P in ("!WORD_PATH!" "!PPT_PATH!" "!EXCEL_PATH!") do (
    if exist "%%~P" (
        if /I "%IsDesignModeEnabled%"=="true" echo [OK] Valid folder: %%~P
    ) else (
        if /I "%IsDesignModeEnabled%"=="true" (
            echo [ERROR] Missing folder: %%~P
            echo [ERROR] Missing folder: %%~P >> "%LOG_FILE%"
        )
    )
)
if /I "%IsDesignModeEnabled%"=="true" (
    echo.
)

rem ==========================================================
rem === STAGE 4: FILE COPY AND REGISTRATION =================
rem ==========================================================
if /I "%IsDesignModeEnabled%"=="true" (
    echo [DEBUG] Starting file copy stage...
    echo [DEBUG] BASE_DIR = "%BASE_DIR%"
    echo -----------------------------------------------
)

for %%F in ("%BASE_DIR%\*.dotx" "%BASE_DIR%\*.dotm" "%BASE_DIR%\*.potx" "%BASE_DIR%\*.potm" "%BASE_DIR%\*.xltx" "%BASE_DIR%\*.xltm") do (
    if exist "%%~fF" (
        set "FN=%%~nxF"
        set "EXT=%%~xF"

        rem === Skip generic templates ===
        set "SKIP=0"
        if /I "!FN!"=="Normal.dotx" set "SKIP=1"
        if /I "!FN!"=="Blank.potx" set "SKIP=1"
        if /I "!FN!"=="Book.xltx" set "SKIP=1"
        if /I "!FN!"=="Normal.dotm" set "SKIP=1"
        if /I "!FN!"=="Blank.potm" set "SKIP=1"
        if /I "!FN!"=="Book.xltm" set "SKIP=1"
        if /I "!FN!"=="Sheet.xltx" set "SKIP=1"
        if /I "!FN!"=="Sheet.xltm" set "SKIP=1"

        rem === Determine destination ===
        set "DEST="
        if /I "!EXT!"==".dotx" set "DEST=!WORD_PATH!"
        if /I "!EXT!"==".dotm" set "DEST=!WORD_PATH!"
        if /I "!EXT!"==".potx" set "DEST=!PPT_PATH!"
        if /I "!EXT!"==".potm" set "DEST=!PPT_PATH!"
        if /I "!EXT!"==".xltx" set "DEST=!EXCEL_PATH!"
        if /I "!EXT!"==".xltm" set "DEST=!EXCEL_PATH!"

        if /I "%IsDesignModeEnabled%"=="true" (
            echo.
            echo [DEBUG] Processing file: !FN!
            echo [DEBUG] Extension detected: !EXT!
        )

        if "!SKIP!"=="1" (
            if /I "%IsDesignModeEnabled%"=="true" (
                echo [INFO] Skipped generic file: !FN!
                echo [INFO] Skipped generic: !FN! >> "%LOG_FILE%"
            )
        ) else if defined DEST (
            if /I "%IsDesignModeEnabled%"=="true" (
                echo [DEBUG] Destination assigned: !DEST!
                echo [ACTION] Copying !FN! → !DEST!
                echo [DEBUG] Copying !FN! to !DEST! >> "%LOG_FILE%"
            )
            mkdir "!DEST!" 2>nul
            if /I "%IsDesignModeEnabled%"=="true" (
                copy /Y "%%~fF" "!DEST!\" >> "%LOG_FILE%" 2>&1
            ) else (
                copy /Y "%%~fF" "!DEST!\" >nul 2>&1
            )

            if exist "!DEST!\!FN!" (
                if /I "%IsDesignModeEnabled%"=="true" (
                    echo [OK] Successfully copied: !FN!
                    echo [RESULT] Success → !FN! >> "%LOG_FILE%"
                )
                set /a TOTAL_FILES+=1

                rem === ADDED: MRU registration ===
                if /I "!EXT!"==".potx" call "%REGISTRY_LIB_PPT%" SimulateRegEntry "!FN!" "!DEST!\!FN!" "%LOG_FILE%"
                if /I "!EXT!"==".potm" call "%REGISTRY_LIB_PPT%" SimulateRegEntry "!FN!" "!DEST!\!FN!" "%LOG_FILE%"
                if /I "!EXT!"==".dotx" call "%REGISTRY_LIB_WORD%" SimulateRegEntry "!FN!" "!DEST!\!FN!" "%LOG_FILE%"
                if /I "!EXT!"==".dotm" call "%REGISTRY_LIB_WORD%" SimulateRegEntry "!FN!" "!DEST!\!FN!" "%LOG_FILE%"
                if /I "!EXT!"==".xltx" call "%REGISTRY_LIB_EXCEL%" SimulateRegEntry "!FN!" "!DEST!\!FN!" "%LOG_FILE%"
                if /I "!EXT!"==".xltm" call "%REGISTRY_LIB_EXCEL%" SimulateRegEntry "!FN!" "!DEST!\!FN!" "%LOG_FILE%"
                rem === END ADDED ===

            ) else (
                if /I "%IsDesignModeEnabled%"=="true" (
                    echo [ERROR] Failed to copy: !FN!
                    echo [RESULT] Error → !FN! >> "%LOG_FILE%"
                )
                set /a TOTAL_ERRORS+=1
            )
        ) else (
            if /I "%IsDesignModeEnabled%"=="true" (
                echo [WARNING] No destination assigned for !FN!
                echo [WARNING] No destination → !FN! >> "%LOG_FILE%"
            )
        )
        if /I "%IsDesignModeEnabled%"=="true" (
            echo -----------------------------------------------
        )
    )
)

if /I "%IsDesignModeEnabled%"=="true" (
    echo [DEBUG] Copy loop finished
    echo [DEBUG] TOTAL_FILES=!TOTAL_FILES! TOTAL_ERRORS=!TOTAL_ERRORS!
)

rem ==========================================================
rem === STAGE 5: FINAL SUMMARY ==============================
rem ==========================================================
if /I "%IsDesignModeEnabled%"=="true" (
    echo.
    echo [FINAL] Copy phase completed.
    echo   Files copied: !TOTAL_FILES!
    echo   Files with errors: !TOTAL_ERRORS!
    echo ----------------------------------------------------------
    echo [DEBUG] Total copied: !TOTAL_FILES!, errors: !TOTAL_ERRORS! >> "%LOG_FILE%"
    echo.
)

endlocal
exit /b
