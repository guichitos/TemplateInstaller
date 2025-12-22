@echo off
setlocal enabledelayedexpansion

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
if not defined CUSTOM_OFFICE_TEMPLATE_ALT_PATH if defined DOCUMENTS_PATH set "CUSTOM_OFFICE_TEMPLATE_ALT_PATH=!DOCUMENTS_PATH!\Plantillas personalizadas de Office"
if not defined CUSTOM_OFFICE_TEMPLATE_ALT_PATH set "CUSTOM_OFFICE_TEMPLATE_ALT_PATH=%USERPROFILE%\Documents\Plantillas personalizadas de Office"


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

start "" "%THEME_PATH%"
start "" "%CUSTOM_OFFICE_TEMPLATE_PATH%"
start "" "%ROAMING_TEMPLATE_PATH%"
start "" "%EXCEL_STARTUP_PATH%"
start "" "%CUSTOM_OFFICE_TEMPLATE_ALT_PATH%"
