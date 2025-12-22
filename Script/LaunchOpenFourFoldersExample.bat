@echo off
setlocal

set "OPEN_THEME=1"
set "OPEN_CUSTOM=1"
set "OPEN_ROAMING=1"
set "OPEN_EXCEL=1"
set "OPEN_CUSTOM_ALT=1"

set "SELECT_FILES=1"
set "THEME_FILE="
set "CUSTOM_FILE="
set "ROAMING_FILE="
set "EXCEL_FILE=aDebe conservarse.xltx"
set "CUSTOM_ALT_FILE=Debe conservarse.xltx"

set "ScriptDirectory=%~dp0"
set "OpenFoldersScript=%ScriptDirectory%OpenFourFolders.bat"

if exist "%OpenFoldersScript%" (
    call "%OpenFoldersScript%" ^
        %OPEN_THEME% ^
        %OPEN_CUSTOM% ^
        %OPEN_ROAMING% ^
        %OPEN_EXCEL% ^
        %OPEN_CUSTOM_ALT% ^
        %SELECT_FILES% ^
        "%THEME_FILE%" ^
        "%CUSTOM_FILE%" ^
        "%ROAMING_FILE%" ^
        "%EXCEL_FILE%" ^
        "%CUSTOM_ALT_FILE%"
) else (
    echo Could not find %OpenFoldersScript%
)

endlocal
