@echo off
setlocal enabledelayedexpansion

set "OPEN_THEME=%~1"
set "OPEN_CUSTOM=%~2"
set "OPEN_ROAMING=%~3"
set "OPEN_EXCEL=%~4"
set "OPEN_CUSTOM_ALT=%~5"

if not defined OPEN_THEME set "OPEN_THEME=1"
if not defined OPEN_CUSTOM set "OPEN_CUSTOM=1"
if not defined OPEN_ROAMING set "OPEN_ROAMING=1"
if not defined OPEN_EXCEL set "OPEN_EXCEL=1"
if not defined OPEN_CUSTOM_ALT set "OPEN_CUSTOM_ALT=1"

set "ScriptDirectory=%~dp0"
set "OfficeTemplateLib=%ScriptDirectory%1-2. AuthContainerTools.bat"

set "APPDATA_EXPANDED="
for /f "delims=" %%T in ('powershell -NoLogo -Command "$app=(Get-ItemProperty -Path \"HKCU:\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\User Shell Folders\" -Name AppData -ErrorAction SilentlyContinue).AppData; if ($app) {[Environment]::ExpandEnvironmentVariables($app)}"') do set "APPDATA_EXPANDED=%%T"
if not defined APPDATA_EXPANDED set "APPDATA_EXPANDED=%APPDATA%"

set "THEME_PATH=%APPDATA_EXPANDED%\Microsoft\Templates\Document Themes"
set "ROAMING_TEMPLATE_PATH=%APPDATA_EXPANDED%\Microsoft\Templates"
set "EXCEL_STARTUP_PATH=%APPDATA_EXPANDED%\Microsoft\Excel\XLSTART"

set "DOCUMENTS_PATH="
for /f "delims=" %%D in ('powershell -NoLogo -Command "$path=(Get-ItemProperty -Path \"HKCU:\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\User Shell Folders\" -Name Personal -ErrorAction SilentlyContinue).Personal; if ($path) {[Environment]::ExpandEnvironmentVariables($path)}"') do set "DOCUMENTS_PATH=%%D"
if defined DOCUMENTS_PATH (
    if "!DOCUMENTS_PATH:~-1!"=="\" set "DOCUMENTS_PATH=!DOCUMENTS_PATH:~0,-1!"
    set "DEFAULT_CUSTOM_DIR=!DOCUMENTS_PATH!\Custom Templates"
) else (
    set "DEFAULT_CUSTOM_DIR=%USERPROFILE%\Documents\Custom Templates"
)
if not defined DEFAULT_CUSTOM_DIR set "DEFAULT_CUSTOM_DIR=%USERPROFILE%\Documents\Custom Templates"

set "CUSTOM_OFFICE_TEMPLATE_PATH="
for %%V in (16.0 15.0 14.0 12.0) do (
    if not defined CUSTOM_OFFICE_TEMPLATE_PATH (
        for /f "tokens=1,2,*" %%A in (
          'reg query "HKCU\Software\Microsoft\Office\%%V\Word\Options" /v "PersonalTemplates" 2^>nul ^| find /I "PersonalTemplates"'
        ) do set "CUSTOM_OFFICE_TEMPLATE_PATH=%%C"
    )
)
for %%V in (16.0 15.0 14.0 12.0) do (
    if not defined CUSTOM_OFFICE_TEMPLATE_PATH (
        for /f "tokens=1,2,*" %%A in (
          'reg query "HKCU\Software\Microsoft\Office\%%V\Common\General" /v "UserTemplates" 2^>nul ^| find /I "UserTemplates"'
        ) do set "CUSTOM_OFFICE_TEMPLATE_PATH=%%C"
    )
)
if not defined CUSTOM_OFFICE_TEMPLATE_PATH if defined DEFAULT_CUSTOM_DIR set "CUSTOM_OFFICE_TEMPLATE_PATH=%DEFAULT_CUSTOM_DIR%"
if not defined CUSTOM_OFFICE_TEMPLATE_PATH set "CUSTOM_OFFICE_TEMPLATE_PATH=%USERPROFILE%\Documents\Custom Templates"

set "CUSTOM_OFFICE_TEMPLATE_ALT_PATH="
for %%V in (16.0 15.0 14.0 12.0) do (
    if not defined CUSTOM_OFFICE_TEMPLATE_ALT_PATH (
        for /f "tokens=1,2,*" %%A in (
          'reg query "HKCU\Software\Microsoft\Office\%%V\PowerPoint\Options" /v "PersonalTemplates" 2^>nul ^| find /I "PersonalTemplates"'
        ) do set "CUSTOM_OFFICE_TEMPLATE_ALT_PATH=%%C"
    )
)
for %%V in (16.0 15.0 14.0 12.0) do (
    if not defined CUSTOM_OFFICE_TEMPLATE_ALT_PATH (
        for /f "tokens=1,2,*" %%A in (
          'reg query "HKCU\Software\Microsoft\Office\%%V\Common\General" /v "UserTemplates" 2^>nul ^| find /I "UserTemplates"'
        ) do set "CUSTOM_OFFICE_TEMPLATE_ALT_PATH=%%C"
    )
)
if not defined CUSTOM_OFFICE_TEMPLATE_ALT_PATH if defined CUSTOM_OFFICE_TEMPLATE_PATH set "CUSTOM_OFFICE_TEMPLATE_ALT_PATH=%CUSTOM_OFFICE_TEMPLATE_PATH%"
if not defined CUSTOM_OFFICE_TEMPLATE_ALT_PATH if defined DEFAULT_CUSTOM_DIR set "CUSTOM_OFFICE_TEMPLATE_ALT_PATH=%DEFAULT_CUSTOM_DIR%"
if not defined CUSTOM_OFFICE_TEMPLATE_ALT_PATH set "CUSTOM_OFFICE_TEMPLATE_ALT_PATH=%USERPROFILE%\Documents\Custom Templates"


if exist "%OfficeTemplateLib%" (
    call "%OfficeTemplateLib%" :CleanPath APPDATA_EXPANDED
    call "%OfficeTemplateLib%" :CleanPath THEME_PATH
    call "%OfficeTemplateLib%" :CleanPath ROAMING_TEMPLATE_PATH
    call "%OfficeTemplateLib%" :CleanPath EXCEL_STARTUP_PATH
    call "%OfficeTemplateLib%" :CleanPath DOCUMENTS_PATH
    call "%OfficeTemplateLib%" :CleanPath DEFAULT_CUSTOM_DIR
    call "%OfficeTemplateLib%" :CleanPath CUSTOM_OFFICE_TEMPLATE_PATH
    call "%OfficeTemplateLib%" :CleanPath CUSTOM_OFFICE_TEMPLATE_ALT_PATH
) else (
    if "!APPDATA_EXPANDED:~-1!"=="\" set "APPDATA_EXPANDED=!APPDATA_EXPANDED:~0,-1!"
    if "!THEME_PATH:~-1!"=="\" set "THEME_PATH=!THEME_PATH:~0,-1!"
    if "!ROAMING_TEMPLATE_PATH:~-1!"=="\" set "ROAMING_TEMPLATE_PATH=!ROAMING_TEMPLATE_PATH:~0,-1!"
    if "!EXCEL_STARTUP_PATH:~-1!"=="\" set "EXCEL_STARTUP_PATH=!EXCEL_STARTUP_PATH:~0,-1!"
    if "!DOCUMENTS_PATH:~-1!"=="\" set "DOCUMENTS_PATH=!DOCUMENTS_PATH:~0,-1!"
    if "!DEFAULT_CUSTOM_DIR:~-1!"=="\" set "DEFAULT_CUSTOM_DIR=!DEFAULT_CUSTOM_DIR:~0,-1!"
    if "!CUSTOM_OFFICE_TEMPLATE_PATH:~-1!"=="\" set "CUSTOM_OFFICE_TEMPLATE_PATH=!CUSTOM_OFFICE_TEMPLATE_PATH:~0,-1!"
    if "!CUSTOM_OFFICE_TEMPLATE_ALT_PATH:~-1!"=="\" set "CUSTOM_OFFICE_TEMPLATE_ALT_PATH=!CUSTOM_OFFICE_TEMPLATE_ALT_PATH:~0,-1!"
)

call :OpenIfEnabled "!OPEN_THEME!" "%THEME_PATH%"
call :OpenIfEnabled "!OPEN_CUSTOM!" "%CUSTOM_OFFICE_TEMPLATE_PATH%"
call :OpenIfEnabled "!OPEN_ROAMING!" "%ROAMING_TEMPLATE_PATH%"
call :OpenIfEnabled "!OPEN_EXCEL!" "%EXCEL_STARTUP_PATH%"
call :OpenIfEnabled "!OPEN_CUSTOM_ALT!" "%CUSTOM_OFFICE_TEMPLATE_ALT_PATH%"

goto :EOF

:OpenIfEnabled
set "FLAG=%~1"
set "TARGET=%~2"

set "SHOULD_OPEN=0"
for %%B in (1 true yes on) do if /I "!FLAG!"=="%%B" set "SHOULD_OPEN=1"

if "!SHOULD_OPEN!"=="1" start "" "%TARGET%"
set "FLAG="
set "TARGET="
set "SHOULD_OPEN="
goto :EOF
