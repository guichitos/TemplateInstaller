@echo off
setlocal

set "OPEN_THEME=1"
set "OPEN_CUSTOM=0"
set "OPEN_ROAMING=1"
set "OPEN_EXCEL=0"
set "OPEN_CUSTOM_ALT=1"

set "ScriptDirectory=%~dp0"
set "OpenFoldersScript=%ScriptDirectory%OpenFourFolders.bat"

if exist "%OpenFoldersScript%" (
    call "%OpenFoldersScript%" %OPEN_THEME% %OPEN_CUSTOM% %OPEN_ROAMING% %OPEN_EXCEL% %OPEN_CUSTOM_ALT%
) else (
    echo Could not find %OpenFoldersScript%
)

endlocal
