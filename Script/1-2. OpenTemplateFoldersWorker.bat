@echo off
setlocal EnableDelayedExpansion

set "OF_DESIGN_MODE=%~1"
set "OF_OPEN_DOC=%~2"
set "OF_DOC_PATH=%~3"
set "OF_DOC_SELECT=%~4"
set "OF_OPEN_CUSTOM=%~5"
set "OF_CUSTOM_PATH=%~6"
set "OF_OPEN_CUSTOM_ALT=%~7"
set "OF_CUSTOM_ALT_PATH=%~8"
set "OF_OPEN_ROAMING=%~9"
set "OF_ROAMING_PATH=%~10"
set "OF_OPEN_EXCEL=%~11"
set "OF_EXCEL_PATH=%~12"
set "OF_EXCEL_SELECT=%~13"

set "OPENED_TEMPLATE_FOLDERS=;"

call :OpenFolderIfRequested "%OF_OPEN_DOC%" "%OF_DOC_PATH%" "%OF_DESIGN_MODE%" "Document Themes folder" "%OF_DOC_SELECT%"
call :OpenFolderIfRequested "%OF_OPEN_CUSTOM%" "%OF_CUSTOM_PATH%" "%OF_DESIGN_MODE%" "Custom Office Templates folder" ""
call :OpenFolderIfRequested "%OF_OPEN_CUSTOM_ALT%" "%OF_CUSTOM_ALT_PATH%" "%OF_DESIGN_MODE%" "Custom Office Templates alternate folder" ""
call :OpenFolderIfRequested "%OF_OPEN_ROAMING%" "%OF_ROAMING_PATH%" "%OF_DESIGN_MODE%" "Roaming Templates folder" ""
call :OpenFolderIfRequested "%OF_OPEN_EXCEL%" "%OF_EXCEL_PATH%" "%OF_DESIGN_MODE%" "Excel startup folder" "%OF_EXCEL_SELECT%"

exit /b 0

:OpenFolderIfRequested
set "REQ_OPEN=%~1"
set "TARGET_PATH=%~2"
set "DESIGN_MODE=%~3"
set "FOLDER_LABEL=%~4"
set "SELECT_PATH=%~5"

if /I not "%REQ_OPEN%"=="true" exit /b
if "%TARGET_PATH%"=="" exit /b
call :NormalizePath "%TARGET_PATH%" TARGET_COMPARE
set "TOKEN=;%TARGET_COMPARE%;"
if "!OPENED_TEMPLATE_FOLDERS:%TOKEN%=!"=="!OPENED_TEMPLATE_FOLDERS!" (
    if /I "%DESIGN_MODE%"=="true" (
        if defined SELECT_PATH (
            echo [ACTION] Opening !FOLDER_LABEL! and selecting: !SELECT_PATH!
        ) else (
            echo [ACTION] Opening !FOLDER_LABEL!: !TARGET_PATH!
        )
    )

    if defined SELECT_PATH (
        if exist "%SELECT_PATH%" (
            start "" explorer /select,"!SELECT_PATH!"
        ) else (
            start "" explorer "!TARGET_PATH!"
        )
    ) else (
        start "" explorer "!TARGET_PATH!"
    )
    set "OPENED_TEMPLATE_FOLDERS=!OPENED_TEMPLATE_FOLDERS!!TOKEN!"
)
exit /b

:NormalizePath
set "NP_INPUT=%~1"
set "NP_OUTPUT_VAR=%~2"
if "%NP_OUTPUT_VAR%"=="" exit /b
setlocal EnableDelayedExpansion
set "NP_WORK=!NP_INPUT!"
:_TrimLoop
if defined NP_WORK if "!NP_WORK:~-1!"==" " set "NP_WORK=!NP_WORK:~0,-1!" & goto _TrimLoop
if defined NP_WORK if "!NP_WORK:~-1!"=="\\" set "NP_WORK=!NP_WORK:~0,-1!" & goto _TrimLoop
endlocal & set "%NP_OUTPUT_VAR%=%NP_WORK%"
exit /b
