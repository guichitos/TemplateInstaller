@echo off
setlocal enabledelayedexpansion

:: Parameters
::  1  - Design mode flag (true/false) for verbose logging
::  2  - Enable file selection (true/false)
::  3  - Open Document Themes folder flag
::  4  - Document Themes folder path
::  5  - Document Themes file name to select (optional)
::  6  - Open Custom Office Templates folder flag
::  7  - Custom Office Templates folder path
::  8  - Custom Office Templates file name to select (optional)
::  9  - Open Roaming Templates folder flag
:: 10  - Roaming Templates folder path
:: 11  - Roaming Templates file name to select (optional)
:: 12  - Open Excel startup folder flag
:: 13  - Excel startup folder path
:: 14  - Excel startup file name to select (optional)
:: 15  - Open Custom Office Templates alternate folder flag
:: 16  - Custom Office Templates alternate folder path
:: 17  - Custom Office Templates alternate file name to select (optional)
set "DESIGN_MODE=%~1"
set "SELECT_FILES=%~2"
set "OPEN_THEME=%~3"
set "THEME_PATH=%~4"
set "THEME_FILE=%~5"
set "OPEN_CUSTOM=%~6"
set "CUSTOM_PATH=%~7"
set "CUSTOM_FILE=%~8"
set "OPEN_ROAMING=%~9"

shift
set "ROAMING_PATH=%~9"
shift
set "ROAMING_FILE=%~9"
shift
set "OPEN_EXCEL=%~9"
shift
set "EXCEL_PATH=%~9"
shift
set "EXCEL_FILE=%~9"
shift
set "OPEN_CUSTOM_ALT=%~9"
shift
set "CUSTOM_ALT_PATH=%~9"
shift
set "CUSTOM_ALT_FILE=%~9"

if /I "!DESIGN_MODE!"=="true" (
    echo [DEBUG] Template folder opener invoked with arguments:
    echo         DESIGN_MODE="!DESIGN_MODE!"
    echo         SELECT_FILES="!SELECT_FILES!"
    echo         OPEN_THEME="!OPEN_THEME!" THEME_PATH="!THEME_PATH!" THEME_FILE="!THEME_FILE!"
    echo         OPEN_CUSTOM="!OPEN_CUSTOM!" CUSTOM_PATH="!CUSTOM_PATH!" CUSTOM_FILE="!CUSTOM_FILE!"
    echo         OPEN_ROAMING="!OPEN_ROAMING!" ROAMING_PATH="!ROAMING_PATH!" ROAMING_FILE="!ROAMING_FILE!"
    echo         OPEN_EXCEL="!OPEN_EXCEL!" EXCEL_PATH="!EXCEL_PATH!" EXCEL_FILE="!EXCEL_FILE!"
    echo         OPEN_CUSTOM_ALT="!OPEN_CUSTOM_ALT!" CUSTOM_ALT_PATH="!CUSTOM_ALT_PATH!" CUSTOM_ALT_FILE="!CUSTOM_ALT_FILE!"
)

set "OPENED_TEMPLATE_FOLDERS=;"

call :OpenIfEnabled "!OPEN_THEME!" "!THEME_PATH!" "!THEME_FILE!" "Document Themes folder"
call :OpenIfEnabled "!OPEN_CUSTOM!" "!CUSTOM_PATH!" "!CUSTOM_FILE!" "Custom Office Templates folder"
call :OpenIfEnabled "!OPEN_CUSTOM_ALT!" "!CUSTOM_ALT_PATH!" "!CUSTOM_ALT_FILE!" "Custom Office Templates alternate folder"
call :OpenIfEnabled "!OPEN_ROAMING!" "!ROAMING_PATH!" "!ROAMING_FILE!" "Roaming Templates folder"
call :OpenIfEnabled "!OPEN_EXCEL!" "!EXCEL_PATH!" "!EXCEL_FILE!" "Excel startup folder"

exit /b 0

:OpenIfEnabled
set "FLAG=%~1"
set "TARGET=%~2"
set "FILENAME=%~3"
set "LABEL=%~4"

if not defined TARGET exit /b
if not defined LABEL set "LABEL=template folder"

set "SHOULD_OPEN=0"
for %%B in (1 true yes on) do if /I "!FLAG!"=="%%B" set "SHOULD_OPEN=1"
if not "!SHOULD_OPEN!"=="1" exit /b

call :NormalizePath "!TARGET!" TARGET_COMPARE
set "TOKEN=;!TARGET_COMPARE!;"
if not "!OPENED_TEMPLATE_FOLDERS:%TOKEN%=!"=="!OPENED_TEMPLATE_FOLDERS!" exit /b

set "IS_ONEDRIVE=0"
if not "!TARGET:\OneDrive\=!"=="!TARGET!" set "IS_ONEDRIVE=1"
if not "!TARGET:\\OneDrive\\=!"=="!TARGET!" set "IS_ONEDRIVE=1"

set "SHOULD_SELECT=0"
for %%B in (1 true yes on) do if /I "!SELECT_FILES!"=="%%B" set "SHOULD_SELECT=1"

if /I "!DESIGN_MODE!"=="true" (
    if defined FILENAME (
        echo [ACTION] Opening !LABEL!; selection requested: "!FILENAME!"
    ) else (
        echo [ACTION] Opening !LABEL!: "!TARGET!"
    )
)

set "FINAL_SELECTION="
if "!SHOULD_SELECT!"=="1" if "!IS_ONEDRIVE!"=="0" (
    call :ResolveSelectionTarget "!TARGET!" "!FILENAME!" FINAL_SELECTION
)

if defined FINAL_SELECTION (
    if exist "!FINAL_SELECTION!" (
        start "" explorer.exe /select,"!FINAL_SELECTION!"
    ) else (
        if /I "!DESIGN_MODE!"=="true" echo [WARN] Selection target not found; opening folder instead: "!FINAL_SELECTION!"
        start "" "!TARGET!"
    )
) else (
    if "!SHOULD_SELECT!"=="1" if "!IS_ONEDRIVE!"=="1" if /I "!DESIGN_MODE!"=="true" echo [DEBUG] Selection skipped because the path is under OneDrive.
    start "" "!TARGET!"
)

set "OPENED_TEMPLATE_FOLDERS=!OPENED_TEMPLATE_FOLDERS!!TOKEN!"
exit /b

:ResolveSelectionTarget
set "RST_BASE=%~1"
set "RST_NAME=%~2"
set "RST_OUT=%~3"
if "%RST_OUT%"=="" exit /b
set "%RST_OUT%="

if not defined RST_NAME exit /b
if "!RST_NAME!"=="" exit /b

set "RST_CANDIDATE=!RST_NAME!"
if "!RST_NAME:~1,1!"==":" goto _SetSelection
if "!RST_NAME:~0,2!"=="\\\\" goto _SetSelection

call :NormalizePath "!RST_BASE!" RST_BASE_TRIMMED
if defined RST_BASE_TRIMMED (
    set "RST_CANDIDATE=!RST_BASE_TRIMMED!\!RST_NAME!"
)

:_SetSelection
set "%RST_OUT%=%RST_CANDIDATE%"
exit /b

:NormalizePath
set "NP_INPUT=%~1"
set "NP_OUTPUT_VAR=%~2"
if "%NP_OUTPUT_VAR%"=="" exit /b
setlocal enabledelayedexpansion
set "NP_WORK=!NP_INPUT!"
:_TrimLoop
if defined NP_WORK if "!NP_WORK:~-1!"==" " set "NP_WORK=!NP_WORK:~0,-1!" & goto _TrimLoop
if defined NP_WORK if "!NP_WORK:~-1!"=="\\" set "NP_WORK=!NP_WORK:~0,-1!" & goto _TrimLoop
endlocal & set "%NP_OUTPUT_VAR%=%NP_WORK%"
exit /b
