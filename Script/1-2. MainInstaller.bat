@echo off
setlocal enabledelayedexpansion

chcp 65001 >nul

rem Acá puede aditarse la lista de autores permitidos
set "DEFAULT_ALLOWED_TEMPLATE_AUTHORS=www.grada.cc;www.gradaz.com"
rem =========================================================

rem Delay (in seconds) between opening the Document Themes folder and launching apps
set "DOCUMENT_THEME_OPEN_DELAY_SECONDS=15"

set "IsDesignModeEnabled=false"

set "ScriptDirectory=%~dp0"
set "BaseHint=%~1"

if not defined BaseHint if defined PIN_LAUNCHER_DIR set "BaseHint=%PIN_LAUNCHER_DIR%"
if not defined BaseHint set "BaseHint=%ScriptDirectory%"

rem Si el instalador se ejecuta directamente desde %APPDATA% sin pista de carpeta,
rem advierte y sale para evitar usar una ruta sin plantillas.
if /I "%BaseHint%"=="%ScriptDirectory%" if /I "%ScriptDirectory:~0,12%"=="%APPDATA%\\" (
    echo [ERROR] No se recibio la ruta de las plantillas. Ejecute el instalador desde "1. Pin templates..." para que se le pase la carpeta correcta.
    exit /b 1
)

call :ResolveBaseDirectory "%BaseHint%" BaseDirectoryPath
if /I "%IsDesignModeEnabled%"=="true" echo [DEBUG] Base directory resolved to: %BaseDirectoryPath%

set "OfficeTemplateLib=%ScriptDirectory%1-2. AuthContainerTools.bat"
set "MRUTools=%ScriptDirectory%1-2. MRU-PathResolver.bat"
set "RepairTemplateMRUScript=%ScriptDirectory%1-2. Repair Office template MRU.bat"
set "ResolveAppProps=%ScriptDirectory%1-2. ResolveAppProperties.bat"
set "MRUInit=%ScriptDirectory%1-2. InitializeMRUSystem.bat"

if not exist "%OfficeTemplateLib%" (
    echo [ERROR] Shared library not found: "%OfficeTemplateLib%"
    exit /b 1
)

if not exist "%MRUTools%" (
    echo [ERROR] MRU library not found: "%MRUTools%"
    exit /b 1
)

if not exist "%RepairTemplateMRUScript%" (
    echo [ERROR] MRU repair script not found: "%RepairTemplateMRUScript%"
    exit /b 1
)

if not defined AuthorValidationEnabled set "AuthorValidationEnabled=TRUE"
if not defined AllowedTemplateAuthors set "AllowedTemplateAuthors=%DEFAULT_ALLOWED_TEMPLATE_AUTHORS%"

if /I "%~1"=="--check-author" (
    set "CTA_CLI_TARGET=%~2"
    set "CTA_CLI_MODE=%~3"
    if not defined CTA_CLI_MODE set "CTA_CLI_MODE=%IsDesignModeEnabled%"
    call :CheckTemplateAuthorAllowed "%CTA_CLI_TARGET%" CTA_CLI_RESULT "%CTA_CLI_MODE%" ""
    if not defined CTA_CLI_RESULT set "CTA_CLI_RESULT="
    echo %CTA_CLI_RESULT%
    exit /b
)

if /I "%IsDesignModeEnabled%"=="true" (
    title TEMPLATE INSTALLER - DEBUG MODE
    echo [DEBUG] Design mode is enabled.
    echo [INFO] Script is running from: %BaseDirectoryPath%
) else (
    title TEMPLATE INSTALLER
    echo Installing custom templates and applying them as the new Microsoft Office defaults
    echo Executing...
    
)

if /I "%IsDesignModeEnabled%"=="true" (
    echo.
    echo [INFO] Verifying environment and closing Office applications...
    call :CheckEnvironment
    echo [INFO] Closing Office applications...
    call :CloseOfficeApps
    echo [OK] Environment verification and Office app closure completed.
) else (
    call :CheckEnvironment "" >nul 2>&1
    call :CloseOfficeApps "" >nul 2>&1
)

if /I "%IsDesignModeEnabled%"=="true" (
    echo.
    echo [INFO] Cleaning template MRU entries before installation...
    call "%RepairTemplateMRUScript%"
) else (
    call "%RepairTemplateMRUScript%" >nul 2>&1
)

set "FORCE_OPEN_WORD=0"
set "FORCE_OPEN_PPT=0"
set "FORCE_OPEN_EXCEL=0"
set "GLOBAL_ITEM_COUNT_WORD=0"
set "GLOBAL_ITEM_COUNT_PPT=0"
set "GLOBAL_ITEM_COUNT_EXCEL=0"
set "LAST_INSTALL_STATUS=0"
set "LAST_INSTALLED_PATH="
set "OPENED_TEMPLATE_FOLDERS=;"
set "OPEN_DOCUMENT_THEME_FLAG=false"
set "DOCUMENT_THEME_SELECT="
set "WORD_BASE_TEMPLATE_DIR=%APPDATA%\Microsoft\Templates"
set "PPT_BASE_TEMPLATE_DIR=%APPDATA%\Microsoft\Templates"
set "EXCEL_BASE_TEMPLATE_DIR=%APPDATA%\Microsoft\Excel\XLSTART"
if /I "%IsDesignModeEnabled%"=="true" (
    echo.
    echo [INFO] Starting base template installation phase...
)
call :InstallBaseTemplates "%IsDesignModeEnabled%"

if /I "%IsDesignModeEnabled%"=="true" (
    echo.
    echo [INFO] Detecting Office personal template folders...
)

call :DetectOfficePaths "" "%IsDesignModeEnabled%"

if /I "%IsDesignModeEnabled%"=="true" (
    echo.
    echo [INFO] Starting custom template copy phase...
)

set "CopyAllErrorLevel=0"

if /I "%IsDesignModeEnabled%"=="true" (
    echo [DEBUG] Invoking CopyAll with base directory and design mode enabled.
    call :CopyAll "" "%BaseDirectoryPath%" "%IsDesignModeEnabled%"
    set "CopyAllErrorLevel=!errorlevel!"
    echo [DEBUG] :CopyAll returned with errorlevel !CopyAllErrorLevel!
    if not "!CopyAllErrorLevel!"=="0" (
        echo [WARN] Non-zero errorlevel detected after CopyAll execution: !CopyAllErrorLevel!
    )
) else (
    call :CopyAll "" "%BaseDirectoryPath%" "%IsDesignModeEnabled%"
    set "CopyAllErrorLevel=!errorlevel!"
)

if /I "%IsDesignModeEnabled%"=="true" (
    echo [DEBUG] Completed CopyAll invocation block (errorlevel=!CopyAllErrorLevel!)
)

call :HandleDocumentThemeFolderOpen "%OPEN_DOCUMENT_THEME_FLAG%" "%IsDesignModeEnabled%" "%THEME_PATH%" "%DOCUMENT_THEME_SELECT%"

if /I "%OPEN_DOCUMENT_THEME_FLAG%"=="true" (
    if /I "%IsDesignModeEnabled%"=="true" echo [INFO] Waiting %DOCUMENT_THEME_OPEN_DELAY_SECONDS% seconds before launching Office apps.
    timeout /t %DOCUMENT_THEME_OPEN_DELAY_SECONDS% /nobreak >nul 2>&1
)

if /I "%IsDesignModeEnabled%"=="true" (
    echo.
    echo [FINAL] Universal Office Template installation completed successfully.
    echo ----------------------------------------------------
)
call :LaunchOfficeApps "%FORCE_OPEN_WORD%" "%FORCE_OPEN_PPT%" "%FORCE_OPEN_EXCEL%" "%IsDesignModeEnabled%" ""
call :EndOfScript
goto :EOF

:InstallBaseTemplates
set "IBT_DesignMode=%~1"

call :InstallApp "WORD" "Normal.dotx" "%APPDATA%\Microsoft\Templates" "Normal.dotx" "" "%BaseDirectoryPath%" "%IBT_DesignMode%"
if "!LAST_INSTALL_STATUS!"=="1" (
    set "FORCE_OPEN_WORD=1"
    call :OpenTemplateFolder "%WORD_BASE_TEMPLATE_DIR%" "" "%IBT_DesignMode%" "base Word template folder" "!LAST_INSTALLED_PATH!"
)
call :InstallApp "WORD" "Normal.dotm" "%APPDATA%\Microsoft\Templates" "Normal.dotm" "" "%BaseDirectoryPath%" "%IBT_DesignMode%"
if "!LAST_INSTALL_STATUS!"=="1" (
    set "FORCE_OPEN_WORD=1"
    call :OpenTemplateFolder "%WORD_BASE_TEMPLATE_DIR%" "" "%IBT_DesignMode%" "base Word template folder" "!LAST_INSTALLED_PATH!"
)
call :InstallApp "WORD" "NormalEmail.dotx" "%APPDATA%\Microsoft\Templates" "NormalEmail.dotx" "" "%BaseDirectoryPath%" "%IBT_DesignMode%"
if "!LAST_INSTALL_STATUS!"=="1" (
    set "FORCE_OPEN_WORD=1"
    call :OpenTemplateFolder "%WORD_BASE_TEMPLATE_DIR%" "" "%IBT_DesignMode%" "base Word template folder" "!LAST_INSTALLED_PATH!"
)
call :InstallApp "WORD" "NormalEmail.dotm" "%APPDATA%\Microsoft\Templates" "NormalEmail.dotm" "" "%BaseDirectoryPath%" "%IBT_DesignMode%"
if "!LAST_INSTALL_STATUS!"=="1" (
    set "FORCE_OPEN_WORD=1"
    call :OpenTemplateFolder "%WORD_BASE_TEMPLATE_DIR%" "" "%IBT_DesignMode%" "base Word template folder" "!LAST_INSTALLED_PATH!"
)
call :InstallApp "POWERPOINT" "Blank.potx" "%APPDATA%\Microsoft\Templates" "Blank.potx" "" "%BaseDirectoryPath%" "%IBT_DesignMode%"
if "!LAST_INSTALL_STATUS!"=="1" (
    set "FORCE_OPEN_PPT=1"
    call :OpenTemplateFolder "%PPT_BASE_TEMPLATE_DIR%" "" "%IBT_DesignMode%" "base PowerPoint template folder" "!LAST_INSTALLED_PATH!"
)
call :InstallApp "POWERPOINT" "Blank.potm" "%APPDATA%\Microsoft\Templates" "Blank.potm" "" "%BaseDirectoryPath%" "%IBT_DesignMode%"
if "!LAST_INSTALL_STATUS!"=="1" (
    set "FORCE_OPEN_PPT=1"
    call :OpenTemplateFolder "%PPT_BASE_TEMPLATE_DIR%" "" "%IBT_DesignMode%" "base PowerPoint template folder" "!LAST_INSTALLED_PATH!"
)
call :InstallApp "EXCEL" "Book.xltx" "%APPDATA%\Microsoft\Excel\XLSTART" "Book.xltx" "" "%BaseDirectoryPath%" "%IBT_DesignMode%"
if "!LAST_INSTALL_STATUS!"=="1" (
    set "FORCE_OPEN_EXCEL=1"
    call :OpenTemplateFolder "%EXCEL_BASE_TEMPLATE_DIR%" "" "%IBT_DesignMode%" "base Excel template folder" "!LAST_INSTALLED_PATH!"
)
call :InstallApp "EXCEL" "Book.xltm" "%APPDATA%\Microsoft\Excel\XLSTART" "Book.xltm" "" "%BaseDirectoryPath%" "%IBT_DesignMode%"
if "!LAST_INSTALL_STATUS!"=="1" (
    set "FORCE_OPEN_EXCEL=1"
    call :OpenTemplateFolder "%EXCEL_BASE_TEMPLATE_DIR%" "" "%IBT_DesignMode%" "base Excel template folder" "!LAST_INSTALLED_PATH!"
)
call :InstallApp "EXCEL" "Sheet.xltx" "%APPDATA%\Microsoft\Excel\XLSTART" "Sheet.xltx" "" "%BaseDirectoryPath%" "%IBT_DesignMode%"
if "!LAST_INSTALL_STATUS!"=="1" (
    set "FORCE_OPEN_EXCEL=1"
    call :OpenTemplateFolder "%EXCEL_BASE_TEMPLATE_DIR%" "" "%IBT_DesignMode%" "base Excel template folder" "!LAST_INSTALLED_PATH!"
)
call :InstallApp "EXCEL" "Sheet.xltm" "%APPDATA%\Microsoft\Excel\XLSTART" "Sheet.xltm" "" "%BaseDirectoryPath%" "%IBT_DesignMode%"
if "!LAST_INSTALL_STATUS!"=="1" (
    set "FORCE_OPEN_EXCEL=1"
    call :OpenTemplateFolder "%EXCEL_BASE_TEMPLATE_DIR%" "" "%IBT_DesignMode%" "base Excel template folder" "!LAST_INSTALLED_PATH!"
)
exit /b 0

:CheckTemplateAuthorAllowed
set "CTA_Target=%~1"
set "CTA_OutputVar=%~2"
set "CTA_DesignMode=%~3"
set "CTA_LogFile=%~4"

if not defined CTA_OutputVar set "CTA_OutputVar=AUTHOR_RESULT"
if not defined CTA_DesignMode set "CTA_DesignMode=%IsDesignModeEnabled%"

setlocal EnableExtensions EnableDelayedExpansion
set "CTA_Target=%CTA_Target%"
set "CTA_Result=TRUE"
set "CTA_Error=0"
set "CTA_Status="

if "!CTA_Target!"=="" (
    set "CTA_Result=FALSE"
    set "CTA_Error=1"
    set "CTA_Status=[ERROR] Debe especificar una ruta de archivo o carpeta."
    goto CTA_HandleMessage
)

if exist "!CTA_Target!\" (
    set "CTA_IsDir=1"
) else if exist "!CTA_Target!" (
    set "CTA_IsDir=0"
) else (
    set "CTA_Result=FALSE"
    set "CTA_Error=1"
    set "CTA_Status=[ERROR] No se encontró la ruta: "!CTA_Target!""
    goto CTA_HandleMessage
)

if "!CTA_IsDir!"=="1" (
    set "TAL_OUTPUT_FILE=%TEMP%\tal_output_!RANDOM!_!RANDOM!.tmp"
    if exist "!TAL_OUTPUT_FILE!" del "!TAL_OUTPUT_FILE!" >nul 2>&1

    powershell -NoProfile -ExecutionPolicy Bypass -Command ^
      "$target='!CTA_Target!'; $out='!TAL_OUTPUT_FILE!';" ^
      "Add-Type -AssemblyName System.IO.Compression.FileSystem;" ^
      "$exts=@('dotx','dotm','xltx','xltm','potx','potm');" ^
      "$files=Get-ChildItem -LiteralPath $target -File | Where-Object { $exts -contains $_.Extension.TrimStart('.') };" ^
      "foreach($f in $files){" ^
      " try{$zip=[IO.Compression.ZipFile]::OpenRead($f.FullName);" ^
      "  $core=$zip.Entries|Where-Object{$_.FullName -eq 'docProps/core.xml'};" ^
      "  if($core){$r=New-Object IO.StreamReader($core.Open());$xml=[xml]$r.ReadToEnd();$r.Close();$a=$xml.coreProperties.creator;}" ^
      "  if($a){'Archivo: '+$f.Name+' - Autor: '+$a|Out-File $out -Append -Encoding UTF8}else{'Archivo: '+$f.Name+' - Autor: [VACÍO]'|Out-File $out -Append -Encoding UTF8}" ^
      " }catch{'[ERROR] '+$f.Name+' → '+$_.Exception.Message|Out-File $out -Append -Encoding UTF8}finally{if($zip){$zip.Dispose()}}" ^
      "}"

    if exist "!TAL_OUTPUT_FILE!" (
        type "!TAL_OUTPUT_FILE!"
        del "!TAL_OUTPUT_FILE!" >nul 2>&1
    )

    set "CTA_Status=[INFO] Autores listados para la carpeta "!CTA_Target!"."
    goto CTA_HandleMessage
)

set "CTA_Check=!AuthorValidationEnabled!"
if not defined CTA_Check set "CTA_Check=TRUE"

if /I "!CTA_Check!"=="FALSE" (
    set "CTA_Result=TRUE"
    set "CTA_Status=[INFO] Validación de autores deshabilitada; se permite "!CTA_Target!"."
    goto CTA_HandleMessage
)

set "authorList="
for /f "usebackq delims=" %%A in (`powershell -NoProfile -ExecutionPolicy Bypass -Command ^
    "$f='!CTA_Target!';" ^
    "Add-Type -AssemblyName System.IO.Compression.FileSystem;" ^
    "try{$zip=[IO.Compression.ZipFile]::OpenRead($f);" ^
    "  $core=$zip.Entries|Where-Object{$_.FullName -eq 'docProps/core.xml'};" ^
    "  if($core){$r=New-Object IO.StreamReader($core.Open());$xml=[xml]$r.ReadToEnd();$r.Close();$a=$xml.coreProperties.creator;}" ^
    "  if($a){Write-Host $a}else{Write-Host '[VACÍO]'}" ^
    "}catch{Write-Host '[ERROR] '+$_.Exception.Message}finally{if($zip){$zip.Dispose()}}"`) do set "authorList=%%A"
if /I "!CTA_DesignMode!"=="true" (
    echo [DEBUG] Autor encontrado: !authorList!
)

if not defined authorList (
    set "CTA_Result=FALSE"
    set "CTA_Status=[WARN] No se pudo obtener el autor para "!CTA_Target!"."
    goto CTA_HandleMessage
)

if /I "!authorList!"=="[VACÍO]" (
    set "CTA_Result=FALSE"
    set "CTA_Status=[WARN] File without Author: "!CTA_Target!"."
    goto CTA_HandleMessage
)

if /I "!authorList:~0,7!"=="[ERROR]" (
    set "CTA_Result=FALSE"
    set "CTA_Status=!authorList!"
    goto CTA_HandleMessage
)

set "authorList=!authorList:,=;!"
set "CTA_Allowed=!AllowedTemplateAuthors!"

set "CTA_Result=TRUE"
for %%a in (!authorList!) do (
    set "author=%%~a"
    for /f "tokens=* delims= " %%t in ("!author!") do set "author=%%t"
    if "!author:~-1!"==" " set "author=!author:~0,-1!"

    set "found=FALSE"
    for %%E in (!CTA_Allowed!) do (
        if /I "%%~E"=="!author!" set "found=TRUE"
    )

    if /I "!found!"=="FALSE" set "CTA_Result=FALSE"
)

if /I "!CTA_DesignMode!"=="true" (
    if /I "!CTA_Result!"=="TRUE" (
        echo [DEBUG] [OK] Autores aprobados
    ) else (
        echo [DEBUG] [FAIL] Autores rechazados
    )
)

goto CTA_HandleMessage

:CTA_HandleMessage
if not defined CTA_Result set "CTA_Result=FALSE"
if not defined CTA_Error set "CTA_Error=0"

if defined CTA_Status (
    if /I "!CTA_DesignMode!"=="true" echo !CTA_Status!
)

set "CTA_FinalResult=!CTA_Result!"
set "CTA_FinalError=!CTA_Error!"
endlocal & (
    if not "%CTA_OutputVar%"=="" set "%CTA_OutputVar%=%CTA_FinalResult%"
)


exit /b %CTA_FinalError%

:ResolveBaseDirectory
setlocal
set "RBD_INPUT=%~1"
set "RBD_OUTPUT_VAR=%~2"

if "%RBD_INPUT:~-1%" NEQ "\\" set "RBD_INPUT=%RBD_INPUT%\\"

set "RBD_FOUND="
for %%D in ("%RBD_INPUT%" "%RBD_INPUT%payload\\" "%RBD_INPUT%templates\\" "%RBD_INPUT%extracted\\"") do (
    for %%F in ("%%~D*.dot*" "%%~D*.pot*" "%%~D*.xlt*" "%%~D*.thmx") do (
        if exist "%%~fF" set "RBD_FOUND=%%~D"
    )
    if defined RBD_FOUND goto :ResolveBaseDirectoryFound
)

:ResolveBaseDirectoryFound
if not defined RBD_FOUND set "RBD_FOUND=%RBD_INPUT%"

endlocal & set "%RBD_OUTPUT_VAR%=%RBD_FOUND%"
exit /b 0

:InstallApp
setlocal EnableDelayedExpansion
set "AppName=%~1"
set "SourceFileName=%~2"
set "DestinationDirectory=%~3"
set "DestinationFileName=%~4"
set "SourceDirectory=%~6"
set "DesignMode=%~7"

if not "%SourceDirectory:~-1%"=="\\" set "SourceDirectory=%SourceDirectory%\\"

set "SourceFilePath=%SourceDirectory%%SourceFileName%"
set "DestinationFilePath=%DestinationDirectory%\%DestinationFileName%"
set "INSTALL_SUCCESS=0"
set "INSTALLED_PATH="

if /I "%DesignMode%"=="true" (
    echo [DEBUG] Source path resolved: "%SourceFilePath%"
    echo [DEBUG] Destination path resolved: "%DestinationFilePath%"
)

if not exist "%SourceFilePath%" (
    if /I "%DesignMode%"=="true" echo.
    if /I "%DesignMode%"=="true" echo [WARNING] Source file not found: "%SourceFilePath%"
    if /I "%DesignMode%"=="true" echo.
    endlocal & set "LAST_INSTALL_STATUS=0"
    exit /b
)

call :CheckTemplateAuthorAllowed "%SourceFilePath%" AUTHOR_RESULT "%DesignMode%" ""

if /I "!AUTHOR_RESULT!"=="FALSE" (
    if /I "%DesignMode%"=="true" (
        echo [BLOCKED] Author not allowed for "%SourceFilePath%"
    )
    endlocal & set "LAST_INSTALL_STATUS=0" & set "LAST_INSTALLED_PATH="
    exit /b
)

if not exist "%DestinationDirectory%" mkdir "%DestinationDirectory%" 2>nul

call :BackupExistingTemplate "%DestinationDirectory%" "%DestinationFileName%" "%DesignMode%" LAST_BACKUP_CREATED LAST_BACKUP_PATH

copy /Y "%SourceFilePath%" "%DestinationFilePath%" >nul 2>&1

if exist "%DestinationFilePath%" (
    if /I "%DesignMode%"=="true" (
        echo [OK] Installed %AppName% template at "%DestinationFilePath%"
    )
    set "INSTALL_SUCCESS=1"
    set "INSTALLED_PATH=%DestinationFilePath%"
) else (
    if /I "%DesignMode%"=="true" (
        echo [ERROR] Copy failed for "%SourceFilePath%"
    )
)

endlocal & set "LAST_INSTALL_STATUS=%INSTALL_SUCCESS%" & set "LAST_INSTALLED_PATH=%INSTALLED_PATH%"
exit /b


:BackupExistingTemplate
if /I "%IsDesignModeEnabled%"=="true" (
echo .
echo "[DEBUG] BackupExistingTemplate called with args: %*")
rem ===========================================================
rem Args: DestinationDirectory, DestinationFileName, DesignMode, OutputFlagVar, OutputPathVar
rem ===========================================================
setlocal EnableDelayedExpansion
set "BET_DestinationDirectory=%~1"
set "BET_DestinationFileName=%~2"
set "BET_DesignMode=%~3"
set "BET_OutputFlagVar=%~4"
set "BET_OutputPathVar=%~5"

set "BET_BackupCreated=0"
set "BET_BackupPath="
set "BET_TargetFile=%BET_DestinationDirectory%\%BET_DestinationFileName%"

if /I "%IsDesignModeEnabled%"=="true" (
    echo [DEBUG] BET_DestinationDirectory="%BET_DestinationDirectory%"
    echo [DEBUG] BET_DestinationFileName="%BET_DestinationFileName%"
    echo [DEBUG] BET_TargetFile="%BET_TargetFile%"
)

if exist "%BET_TargetFile%" (
    set "BET_BackupDir=%BET_DestinationDirectory%\Backup"
    if /I "%BET_DesignMode%"=="true" echo [DEBUG] Preparing backup directory at "!BET_BackupDir!"

    if not exist "!BET_BackupDir!" (
        mkdir "!BET_BackupDir!" >nul 2>&1
        if /I "%BET_DesignMode%"=="true" echo [DEBUG] mkdir result for "!BET_BackupDir!": !errorlevel!
    )

    if not exist "!BET_BackupDir!" (
        if /I "%BET_DesignMode%"=="true" echo [WARN] Could not create backup directory: "!BET_BackupDir!"
        goto :BET_End
    )

    for /f "delims=" %%T in ('powershell -NoProfile -Command "Get-Date -Format yyyy.MM.dd.HHmm"') do set "BET_Timestamp=%%T"
    if not defined BET_Timestamp set "BET_Timestamp=%DATE%_%TIME%"

    set "BET_BackupPath=!BET_BackupDir!\!BET_Timestamp!_%BET_DestinationFileName%"

    if /I "%BET_DesignMode%"=="true" (
        echo [INFO] Preparing backup.
        echo         Source : "!BET_TargetFile!"
        echo         Backup : "!BET_BackupPath!"
    )

    if not exist "!BET_TargetFile!" (
        if /I "%BET_DesignMode%"=="true" echo [ERROR] Backup source not found: "!BET_TargetFile!"
        goto :BET_End
    )

    copy /Y "!BET_TargetFile!" "!BET_BackupPath!" >nul 2>&1

    if exist "!BET_BackupPath!" (
        set "BET_BackupCreated=1"
        if /I "%BET_DesignMode%"=="true" echo [BACKUP] Created for %BET_DestinationFileName% at "!BET_BackupPath!"
    ) else (
        if /I "%BET_DesignMode%"=="true" echo [WARN] Failed to create backup for %BET_DestinationFileName% at "!BET_BackupPath!"
    )
) else (
    if /I "%BET_DesignMode%"=="true" echo [INFO] No existing file to backup for %BET_DestinationFileName% at "!BET_TargetFile!".
)

:BET_End

if /I "%IsDesignModeEnabled%"=="true" (
    echo [DEBUG] BET_BackupCreated="!BET_BackupCreated!"
    echo [DEBUG] BET_BackupPath="!BET_BackupPath!"
)

:BET_End

if /I "%IsDesignModeEnabled%"=="true" (
    echo [DEBUG] BET_BackupCreated="%BET_BackupCreated%"
    echo [DEBUG] BET_BackupPath="%BET_BackupPath%"
)

:BET_End

endlocal & (
    if not "%BET_OutputFlagVar%"=="" set "%BET_OutputFlagVar%=%BET_BackupCreated%"
    if not "%BET_OutputPathVar%"=="" set "%BET_OutputPathVar%=%BET_BackupPath%"
)
if /I "%IsDesignModeEnabled%"=="true" (
echo [DEBUG] BackupExistingTemplate completed.
echo .
)
exit /b 0


:CheckEnvironment
echo [DEBUG] Environment check starting...
echo [DEBUG] Environment check completed.
exit /b 0

:DetectOfficePaths
setlocal enabledelayedexpansion
set "LOG_FILE=%~1"
set "DESIGN_MODE=%~2"

set "WORD_PATH="
set "PPT_PATH="
set "EXCEL_PATH="
set "THEME_PATH="
set "WORD_PATH_FALLBACK=0"
set "PPT_PATH_FALLBACK=0"
set "EXCEL_PATH_FALLBACK=0"
set "THEME_PATH_STATUS=unknown"
set "DEFAULT_CUSTOM_DIR="
set "DEFAULT_CUSTOM_DIR_CREATED=0"
set "DEFAULT_CUSTOM_DIR_STATUS=unknown"
set "DOCUMENTS_PATH="
set "OFFICE_VERSIONS=16.0 15.0 14.0 12.0"


for /f "delims=" %%T in ('powershell -NoLogo -Command "$app=(Get-ItemProperty -Path \"HKCU:\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\User Shell Folders\" -Name AppData -ErrorAction SilentlyContinue).AppData; if ($app) {[Environment]::ExpandEnvironmentVariables($app)}"') do set "APPDATA_EXPANDED=%%T"
if not defined APPDATA_EXPANDED set "APPDATA_EXPANDED=%APPDATA%"
if defined APPDATA_EXPANDED set "THEME_PATH=!APPDATA_EXPANDED!\Microsoft\Templates\Document Themes"


for /f "delims=" %%D in ('powershell -NoLogo -Command "$path=(Get-ItemProperty -Path \"HKCU:\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\User Shell Folders\" -Name Personal -ErrorAction SilentlyContinue).Personal; if ($path) {[Environment]::ExpandEnvironmentVariables($path)}"') do set "DOCUMENTS_PATH=%%D"

if defined DOCUMENTS_PATH (
    if "!DOCUMENTS_PATH:~-1!"=="\" set "DOCUMENTS_PATH=!DOCUMENTS_PATH:~0,-1!"
    set "DEFAULT_CUSTOM_DIR=!DOCUMENTS_PATH!\Custom Templates"
) else (
    set "DEFAULT_CUSTOM_DIR=%USERPROFILE%\Documents\Custom Templates"
)

if not defined DEFAULT_CUSTOM_DIR set "DEFAULT_CUSTOM_DIR=%USERPROFILE%\Documents\Custom Templates"
if defined DEFAULT_CUSTOM_DIR (
    if exist "!DEFAULT_CUSTOM_DIR!" (
        set "DEFAULT_CUSTOM_DIR_STATUS=exists"
    ) else (
        mkdir "!DEFAULT_CUSTOM_DIR!" >nul 2>&1
        if not errorlevel 1 (
            set "DEFAULT_CUSTOM_DIR_CREATED=1"
            set "DEFAULT_CUSTOM_DIR_STATUS=created"
        ) else (
            set "DEFAULT_CUSTOM_DIR_STATUS=create_failed"
        )
    )
)
if defined THEME_PATH (
    if exist "!THEME_PATH!" (
        set "THEME_PATH_STATUS=exists"
    ) else (
        mkdir "!THEME_PATH!" >nul 2>&1
        if not errorlevel 1 (
            set "THEME_PATH_STATUS=created"
        ) else (
            set "THEME_PATH_STATUS=create_failed"
        )
    )
)

for %%V in (!OFFICE_VERSIONS!) do (
    if not defined WORD_PATH (
        for /f "tokens=1,2,*" %%A in (
          'reg query "HKCU\Software\Microsoft\Office\%%V\Word\Options" /v "PersonalTemplates" 2^>nul ^| find /I "PersonalTemplates"'
        ) do set "WORD_PATH=%%C"
    )
    if not defined PPT_PATH (
        for /f "tokens=1,2,*" %%A in (
          'reg query "HKCU\Software\Microsoft\Office\%%V\PowerPoint\Options" /v "PersonalTemplates" 2^>nul ^| find /I "PersonalTemplates"'
        ) do set "PPT_PATH=%%C"
    )
    if not defined EXCEL_PATH (
        for /f "tokens=1,2,*" %%A in (
          'reg query "HKCU\Software\Microsoft\Office\%%V\Excel\Options" /v "PersonalTemplates" 2^>nul ^| find /I "PersonalTemplates"'
        ) do set "EXCEL_PATH=%%C"
    )
)

for %%V in (!OFFICE_VERSIONS!) do (
    if not defined WORD_PATH (
        for /f "tokens=1,2,*" %%A in (
          'reg query "HKCU\Software\Microsoft\Office\%%V\Common\General" /v "UserTemplates" 2^>nul ^| find /I "UserTemplates"'
        ) do set "WORD_PATH=%%C"
    )
    if not defined PPT_PATH (
        for /f "tokens=1,2,*" %%A in (
          'reg query "HKCU\Software\Microsoft\Office\%%V\Common\General" /v "UserTemplates" 2^>nul ^| find /I "UserTemplates"'
        ) do set "PPT_PATH=%%C"
    )
    if not defined EXCEL_PATH (
        for /f "tokens=1,2,*" %%A in (
          'reg query "HKCU\Software\Microsoft\Office\%%V\Common\General" /v "UserTemplates" 2^>nul ^| find /I "UserTemplates"'
        ) do set "EXCEL_PATH=%%C"
    )
)

if not defined WORD_PATH if defined DEFAULT_CUSTOM_DIR (
    set "WORD_PATH=!DEFAULT_CUSTOM_DIR!"
    set "WORD_PATH_FALLBACK=1"
)
if not defined PPT_PATH if defined DEFAULT_CUSTOM_DIR (
    set "PPT_PATH=!DEFAULT_CUSTOM_DIR!"
    set "PPT_PATH_FALLBACK=1"
)
if not defined EXCEL_PATH if defined DEFAULT_CUSTOM_DIR (
    set "EXCEL_PATH=!DEFAULT_CUSTOM_DIR!"
    set "EXCEL_PATH_FALLBACK=1"
)

call "%OfficeTemplateLib%" :CleanPath WORD_PATH
call "%OfficeTemplateLib%" :CleanPath PPT_PATH
call "%OfficeTemplateLib%" :CleanPath EXCEL_PATH
if defined THEME_PATH call "%OfficeTemplateLib%" :CleanPath THEME_PATH

if /I "!DESIGN_MODE!"=="true" (
    if "!DEFAULT_CUSTOM_DIR_STATUS!"=="created" (
        echo [INFO] Created default "Custom Templates" folder at: !DEFAULT_CUSTOM_DIR!
    ) else if "!DEFAULT_CUSTOM_DIR_STATUS!"=="exists" (
        echo [DEBUG] Default "Custom Templates" folder already exists at: !DEFAULT_CUSTOM_DIR!
    ) else if "!DEFAULT_CUSTOM_DIR_STATUS!"=="create_failed" (
        echo [WARNING] Failed to create default "Custom Templates" folder at: !DEFAULT_CUSTOM_DIR!
    )
    if defined WORD_PATH (
        if "!WORD_PATH_FALLBACK!"=="1" (
            echo [INFO] Word templates folder defaulted to: !WORD_PATH!
        ) else (
            echo [DEBUG] Word templates folder detected: !WORD_PATH!
        )
    ) else (
        echo [WARNING] No Word templates folder detected from registry.
    )
    if defined PPT_PATH (
        if "!PPT_PATH_FALLBACK!"=="1" (
            echo [INFO] PowerPoint templates folder defaulted to: !PPT_PATH!
        ) else (
            echo [DEBUG] PowerPoint templates folder detected: !PPT_PATH!
        )
    ) else (
        echo [WARNING] No PowerPoint templates folder detected from registry.
    )
    if defined EXCEL_PATH ( 
        if "!EXCEL_PATH_FALLBACK!"=="1" (
            echo [INFO] Excel templates folder defaulted to: !EXCEL_PATH!
        ) else (
            echo [DEBUG] Excel templates folder detected: !EXCEL_PATH!
        )
    ) else (
        echo [WARNING] No Excel templates folder detected from registry.
    )
    if defined THEME_PATH (
        if "!THEME_PATH_STATUS!"=="created" (
            echo [INFO] Document Themes folder created at: !THEME_PATH!
        ) else if "!THEME_PATH_STATUS!"=="exists" (
            echo [DEBUG] Document Themes folder detected: !THEME_PATH!
        ) else if "!THEME_PATH_STATUS!"=="create_failed" (
            echo [WARNING] Failed to create Document Themes folder at: !THEME_PATH!
        )
    ) else (
        echo [WARNING] Document Themes folder path could not be determined.
    )
)

endlocal & (
    set "WORD_PATH=%WORD_PATH%"
    set "PPT_PATH=%PPT_PATH%"
    set "EXCEL_PATH=%EXCEL_PATH%"
    set "THEME_PATH=%THEME_PATH%"
)

exit /b

:CloseOfficeApps
echo [DEBUG] Entering Closing Office applications with args: %*
taskkill /IM WINWORD.EXE /F >nul 2>&1
taskkill /IM POWERPNT.EXE /F >nul 2>&1
taskkill /IM EXCEL.EXE /F >nul 2>&1
echo [DEBUG] Exiting Closing Office applications...
exit /b

:HandleDocumentThemeFolderOpen
set "DT_SHOULD_OPEN=%~1"
set "DT_DESIGN_MODE=%~2"
set "DT_TARGET=%~3"
set "DT_SELECT=%~4"

if /I "%DT_SHOULD_OPEN%"=="true" (
    if defined DT_TARGET if exist "%DT_TARGET%" (
        call :OpenTemplateFolder "%DT_TARGET%" "" "%DT_DESIGN_MODE%" "Document Themes folder" "%DT_SELECT%"
    ) else if /I "%DT_DESIGN_MODE%"=="true" (
        echo [DEBUG] Document Themes folder not opened because the path is unavailable.
    )
) else if /I "%DT_DESIGN_MODE%"=="true" (
    echo [DEBUG] Document Themes folder open flag is false; skipping launch.
)
exit /b

:OpenTemplateFolder
set "TARGET_PATH=%~1"
set "DESIGN_MODE=%~2"
set "FOLDER_LABEL=%~3"
set "SELECT_PATH=%~4"

if "%TARGET_PATH%"=="" exit /b
if not exist "%TARGET_PATH%" exit /b
if "%FOLDER_LABEL%"=="" set "FOLDER_LABEL=template folder"
if not defined OPENED_TEMPLATE_FOLDERS set "OPENED_TEMPLATE_FOLDERS=;"
set "TOKEN=;%TARGET_PATH%;"
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

:LaunchOfficeApps
setlocal EnableDelayedExpansion
set "OPEN_WORD_FLAG=%~1"
set "OPEN_PPT_FLAG=%~2"
set "OPEN_EXCEL_FLAG=%~3"
set "LAUNCH_DESIGN_MODE=%~4"
set "ANY_LAUNCH=0"
if not defined OPEN_WORD_FLAG set "OPEN_WORD_FLAG=0"
if not defined OPEN_PPT_FLAG set "OPEN_PPT_FLAG=0"
if not defined OPEN_EXCEL_FLAG set "OPEN_EXCEL_FLAG=0"

if /I "!OPEN_WORD_FLAG!"=="1" (
    set "ANY_LAUNCH=1"
    call :LaunchSingleOfficeApp "winword.exe" "Microsoft Word" "!LAUNCH_DESIGN_MODE!"
) else (
    if /I "!LAUNCH_DESIGN_MODE!"=="true" (
        echo [INFO] Microsoft Word will remain closed no new templates applied.
    )
)

if /I "!OPEN_PPT_FLAG!"=="1" (
    set "ANY_LAUNCH=1"
    call :LaunchSingleOfficeApp "powerpnt.exe" "Microsoft PowerPoint" "!LAUNCH_DESIGN_MODE!"
) else if /I "!LAUNCH_DESIGN_MODE!"=="true" (
    echo [INFO] Microsoft PowerPoint will remain closed no new templates applied.
)

if /I "!OPEN_EXCEL_FLAG!"=="1" (
    set "ANY_LAUNCH=1"
    call :LaunchSingleOfficeApp "excel.exe" "Microsoft Excel" "!LAUNCH_DESIGN_MODE!"
) else if /I "!LAUNCH_DESIGN_MODE!"=="true" (
    echo [INFO] Microsoft Excel will remain closed no new templates applied.
)
if /I "%DESIGN_MODE%"=="true" (
    if "!ANY_LAUNCH!"=="0" if /I "!LAUNCH_DESIGN_MODE!"=="true" echo [INFO] No Office applications need to be relaunched.
)

endlocal
exit /b

:LaunchSingleOfficeApp
setlocal EnableDelayedExpansion
set "APP_EXECUTABLE=%~1"
set "APP_FRIENDLY=%~2"
set "APP_DESIGN_MODE=%~3"
set "APP_PATH="
set "APP_PATH_RESOLVED=0"
set "APP_RESOLUTION_SOURCE=PATH"


for /f "usebackq delims=" %%P in (`where "!APP_EXECUTABLE!" 2^>nul`) do (
    if not defined APP_PATH (
        set "APP_PATH=%%~fP"
        set "APP_PATH_RESOLVED=1"
        set "APP_RESOLUTION_SOURCE=PATH"
    )
)
if not defined APP_PATH (
    for %%R in ("HKLM\Software\Microsoft\Windows\CurrentVersion\App Paths\!APP_EXECUTABLE!" "HKCU\Software\Microsoft\Windows\CurrentVersion\App Paths\!APP_EXECUTABLE!" "HKLM\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\!APP_EXECUTABLE!") do (
        if not defined APP_PATH (
            for /f "tokens=1,*" %%A in ('reg query %%~R /ve 2^>nul ^| findstr /I "REG_SZ"') do (
                if not defined APP_PATH (
                    set "APP_PATH=%%B"
                    set "APP_RESOLUTION_SOURCE=REGISTRY"
                )
            )
        )
    )
)
if defined APP_PATH set "APP_PATH=!APP_PATH:"=!"

if defined APP_PATH if exist "!APP_PATH!" (
    set "APP_PATH_RESOLVED=1"
)

if "!APP_PATH_RESOLVED!"=="0" (
    for %%B in ("!ProgramFiles!" "!ProgramFiles(x86)!" "!ProgramW6432!") do (
        if not "%%~B"=="" (
            if not defined APP_PATH (
                for %%V in (16 15 14 13 12) do (
                    if not defined APP_PATH if exist "%%~B\Microsoft Office\root\Office%%V\!APP_EXECUTABLE!" (
                        set "APP_PATH=%%~B\Microsoft Office\root\Office%%V\!APP_EXECUTABLE!"
                        set "APP_RESOLUTION_SOURCE=PROGRAMFILES"
                    )
                    if not defined APP_PATH if exist "%%~B\Microsoft Office\Office%%V\!APP_EXECUTABLE!" (
                        set "APP_PATH=%%~B\Microsoft Office\Office%%V\!APP_EXECUTABLE!"
                        set "APP_RESOLUTION_SOURCE=PROGRAMFILES"
                    )
                )
            )
        )
    )
)

if defined APP_PATH if exist "!APP_PATH!" (
    set "APP_PATH_RESOLVED=1"
)
if not defined APP_PATH (
    set "APP_PATH=!APP_EXECUTABLE!"
)

if /I "!APP_DESIGN_MODE!"=="true" (
    if "!APP_PATH_RESOLVED!"=="1" (
        echo [ACTION] Launching !APP_FRIENDLY! from "!APP_PATH!" source - !APP_RESOLUTION_SOURCE!
    ) else (
        echo [WARN] Unable to resolve full path for !APP_FRIENDLY! - attempting to launch via PATH lookup !APP_EXECUTABLE!
    )
)

set "APP_LAUNCH_TARGET=!APP_PATH!"
if "!APP_PATH_RESOLVED!"=="0" set "APP_LAUNCH_TARGET=!APP_EXECUTABLE!"
start "" "!APP_LAUNCH_TARGET!" >nul 2>&1
set "APP_START_ERROR=!errorlevel!"
if not "!APP_START_ERROR!"=="0" (
    if /I "!APP_DESIGN_MODE!"=="true" echo [WARN] Unable to launch !APP_FRIENDLY! using "!APP_LAUNCH_TARGET!" errorlevel=!APP_START_ERROR!
)

endlocal
exit /b

:CopyAll
setlocal enabledelayedexpansion
set "LOG_FILE=%~1"
set "BASE_DIR=%~2"
set "IsDesignModeEnabled=%~3"

if not defined BASE_DIR set "BASE_DIR=%~dp0"
if not "%BASE_DIR:~-1%"=="\\" set "BASE_DIR=%BASE_DIR%\\"

set /a TOTAL_FILES=0
set /a TOTAL_ERRORS=0
set /a TOTAL_BLOCKED=0
set "OPEN_WORD=0"
set "OPEN_PPT=0"
set "OPEN_EXCEL=0"
set "OPEN_THEME=0"
set "WORD_SELECT="
set "PPT_SELECT="
set "EXCEL_SELECT="
set "THEME_SELECT="

if defined FORCE_OPEN_WORD if "!FORCE_OPEN_WORD!"=="1" set "OPEN_WORD=1"
if defined FORCE_OPEN_PPT if "!FORCE_OPEN_PPT!"=="1" set "OPEN_PPT=1"
if defined FORCE_OPEN_EXCEL if "!FORCE_OPEN_EXCEL!"=="1" set "OPEN_EXCEL=1"

if defined WORD_PATH if "!WORD_PATH:~-1!"=="\" set "WORD_PATH=!WORD_PATH:~0,-1!"
if defined PPT_PATH  if "!PPT_PATH:~-1!"=="\"  set "PPT_PATH=!PPT_PATH:~0,-1!"
if defined EXCEL_PATH if "!EXCEL_PATH:~-1!"=="\" set "EXCEL_PATH=!EXCEL_PATH:~0,-1!"
if defined THEME_PATH if "!THEME_PATH:~-1!"=="\" set "THEME_PATH=!THEME_PATH:~0,-1!"

call "%MRUTools%" :DetectMRUPath WORD ADAL
call "%MRUTools%" :DetectMRUPath WORD LIVEID
call "%MRUTools%" :DetectMRUPath POWERPOINT ADAL
call "%MRUTools%" :DetectMRUPath POWERPOINT LIVEID
call "%MRUTools%" :DetectMRUPath EXCEL ADAL
call "%MRUTools%" :DetectMRUPath EXCEL LIVEID


if /I "%IsDesignModeEnabled%"=="true" (
    echo [INFO] Scanning BASE_DIR for templates...
    echo -----------------------------------------------
    dir /b "%BASE_DIR%\*.dot*" "%BASE_DIR%\*.pot*" "%BASE_DIR%\*.xlt*" 2>nul
    echo -----------------------------------------------
    echo.
)

for %%P in ("!WORD_PATH!" "!PPT_PATH!" "!EXCEL_PATH!" "!THEME_PATH!") do (
    if not "%%~P"=="" (
        if not exist "%%~P" mkdir "%%~P" >nul 2>&1
    )
)

if /I "%IsDesignModeEnabled%"=="true" (
    echo [DEBUG] Starting file copy stage...
    echo -----------------------------------------------
)

for %%F in ("%BASE_DIR%*.dotx" "%BASE_DIR%*.dotm" "%BASE_DIR%*.potx" "%BASE_DIR%*.potm" "%BASE_DIR%*.xltx" "%BASE_DIR%*.xltm" "%BASE_DIR%*.thmx") do (
    if exist "%%~fF" (
        set "FN=%%~nxF"
        set "EXT=%%~xF"


        if /I "%IsDesignModeEnabled%"=="true" (
            echo [DEBUG] Iteración actual:
            echo     %%F
            echo     FN=%%~nxF
            echo     EXT=%%~xF
            echo ----------------------------
        )

        set "SKIP=0"
        for %%G in (Normal.dotx NormalEmail.dotx Blank.potx Book.xltx Normal.dotm NormalEmail.dotm Blank.potm Book.xltm Sheet.xltx Sheet.xltm) do (
            if /I "!FN!"=="%%G" set "SKIP=1"
        )

        set "DEST="
        if /I "!EXT!"==".dotx" set "DEST=!WORD_PATH!"
        if /I "!EXT!"==".dotm" set "DEST=!WORD_PATH!"
        if /I "!EXT!"==".potx" set "DEST=!PPT_PATH!"
        if /I "!EXT!"==".potm" set "DEST=!PPT_PATH!"
        if /I "!EXT!"==".xltx" set "DEST=!EXCEL_PATH!"
        if /I "!EXT!"==".xltm" set "DEST=!EXCEL_PATH!"
        if /I "!EXT!"==".thmx" set "DEST=!THEME_PATH!"

        if /I "%IsDesignModeEnabled%"=="true" (
            echo.
            echo [DEBUG] Processing file: !FN!
            echo [DEBUG] Destination assigned: !DEST!
        )

        if "!SKIP!"=="1" (
            if /I "%IsDesignModeEnabled%"=="true" (
                echo [INFO] Skipped default file: !FN!
            )
        ) else if defined DEST ( 
            set "AUTHOR_RESULT="
            call :CheckTemplateAuthorAllowed "%%~fF" AUTHOR_RESULT "%IsDesignModeEnabled%" ""

            if /I "!AUTHOR_RESULT!"=="FALSE" (
                if /I "%IsDesignModeEnabled%"=="true" (
                    echo [BLOCKED] Author not allowed for "!FN!"
                )
                set /a TOTAL_BLOCKED+=1
            ) else (
                copy /Y "%%~fF" "!DEST!\" >nul 2>&1
                if exist "!DEST!\!FN!" (
                    if /I "%IsDesignModeEnabled%"=="true" echo [OK] Copied: !FN!
                    set /a TOTAL_FILES+=1
                    if /I "!DEST!"=="!WORD_PATH!" (
                        set "OPEN_WORD=1"
                        if "!WORD_SELECT!"=="" set "WORD_SELECT=!DEST!\!FN!"
                    )
                    if /I "!DEST!"=="!PPT_PATH!" (
                        set "OPEN_PPT=1"
                        if "!PPT_SELECT!"=="" set "PPT_SELECT=!DEST!\!FN!"
                    )
                    if /I "!DEST!"=="!EXCEL_PATH!" (
                        set "OPEN_EXCEL=1"
                        if "!EXCEL_SELECT!"=="" set "EXCEL_SELECT=!DEST!\!FN!"
                    )
                    if /I "!DEST!"=="!THEME_PATH!" (
                        set "OPEN_THEME=1"
                        if "!THEME_SELECT!"=="" set "THEME_SELECT=!DEST!\!FN!"
                    )
                    rem Register MRU
                    if /I "!EXT!"==".dotx" call :SimulateRegEntry WORD "!FN!" "!DEST!\!FN!" ""
                    if /I "!EXT!"==".dotm" call :SimulateRegEntry WORD "!FN!" "!DEST!\!FN!" ""
                    if /I "!EXT!"==".potx" call :SimulateRegEntry POWERPOINT "!FN!" "!DEST!\!FN!" ""
                    if /I "!EXT!"==".potm" call :SimulateRegEntry POWERPOINT "!FN!" "!DEST!\!FN!" ""
                    if /I "!EXT!"==".xltx" call :SimulateRegEntry EXCEL "!FN!" "!DEST!\!FN!" ""
                    if /I "!EXT!"==".xltm" call :SimulateRegEntry EXCEL "!FN!" "!DEST!\!FN!" ""
                ) else (
                    if /I "%IsDesignModeEnabled%"=="true" echo [ERROR] Failed to copy: !FN!
                    set /a TOTAL_ERRORS+=1
                )
            )
        ) else (
            if /I "%IsDesignModeEnabled%"=="true" (
                echo [WARNING] No destination assigned for !FN!
            )
        )

        if /I "%IsDesignModeEnabled%"=="true" echo -----------------------------------------------
    )
)

if /I "%IsDesignModeEnabled%"=="true" (
    echo entra a debug checkpoint
    echo [DEBUG] Copy loop finished
    echo [DEBUG] TOTAL_FILES=!TOTAL_FILES! TOTAL_ERRORS=!TOTAL_ERRORS! TOTAL_BLOCKED=!TOTAL_BLOCKED!
)

if "!OPEN_WORD!"=="1" if exist "!WORD_PATH!" (
    call :OpenTemplateFolder "!WORD_PATH!" "" "%IsDesignModeEnabled%" "Word template folder" "!WORD_SELECT!"
)

if "!OPEN_PPT!"=="1" if exist "!PPT_PATH!" (
    call :OpenTemplateFolder "!PPT_PATH!" "" "%IsDesignModeEnabled%" "PowerPoint template folder" "!PPT_SELECT!"
)

if "!OPEN_EXCEL!"=="1" if exist "!EXCEL_PATH!" (
    call :OpenTemplateFolder "!EXCEL_PATH!" "" "%IsDesignModeEnabled%" "Excel template folder" "!EXCEL_SELECT!"
)

set "OPEN_DOCUMENT_THEME_FLAG=false"
if "!OPEN_THEME!"=="1" set "OPEN_DOCUMENT_THEME_FLAG=true"

if /I "%IsDesignModeEnabled%"=="true" (
    echo [DEBUG] Exiting CopyAll routine - pre-endlocal
)

if /I "%IsDesignModeEnabled%"=="true" (
    echo.
    echo [FINAL] Copy phase completed.
    echo   Files copied: !TOTAL_FILES!
    echo   Files with errors: !TOTAL_ERRORS!
    echo   Files blocked - author: !TOTAL_BLOCKED!
    echo ----------------------------------------------------------
)

endlocal & (
    set "FORCE_OPEN_WORD=%OPEN_WORD%"
    set "FORCE_OPEN_PPT=%OPEN_PPT%"
    set "FORCE_OPEN_EXCEL=%OPEN_EXCEL%"
    set "OPEN_DOCUMENT_THEME_FLAG=%OPEN_DOCUMENT_THEME_FLAG%"
    set "DOCUMENT_THEME_SELECT=%THEME_SELECT%"
)
exit /b

:AddMruTargetFromVar
echo [DEBUG - AddMruTargetFromVar] Entering AddMruTargetFromVar with args: %*
set "AMT_TARGET=%~1"
set "AMT_VAR_NAME=%~2"
set "AMT_APP=%~3"
set "AMT_AUTH=%~4"

if "%AMT_TARGET%"=="" exit /b 0
if "%AMT_VAR_NAME%"=="" exit /b 0

set "AMT_VALUE="
call set "AMT_VALUE=%%%AMT_VAR_NAME%%%"

if not defined AMT_VALUE (
    if not "%AMT_APP%"=="" if not "%AMT_AUTH%"=="" call "%MRUTools%" :DetectMRUPath "%AMT_APP%" "%AMT_AUTH%"
    set "AMT_VALUE="
    call set "AMT_VALUE=%%%AMT_VAR_NAME%%%"
)

if not defined AMT_VALUE exit /b 0

call "%OfficeTemplateLib%" :AppendUniquePath %AMT_TARGET% "!AMT_VALUE!"
exit /b 0

:SimulateRegEntry
setlocal enabledelayedexpansion
set "APP_NAME=%~1"
set "FILE_NAME=%~2"
set "FULL_PATH=%~3"
set "LOG_FILE=%~4"

call "%ResolveAppProps%" "%APP_NAME%"
if not defined PROP_REG_NAME (

    endlocal
    exit /b 1
)
set "SCRIPT_DIR=%~dp0"
set "COUNTER_VAR=!PROP_COUNTER_VAR!"
set "LOCAL_LOGGING=true"

if /I "%IsDesignModeEnabled%"=="false" set "LOCAL_LOGGING=false"
set "MRU_TARGET_PATHS="
set "MRU_FALLBACK="

call :ResolveAppProperties "%APP_NAME%"
if /I "%IsDesignModeEnabled%"=="true" (
echo [DEBUG] PROP_REG_NAME=!PROP_REG_NAME!  
echo [DEBUG] PROP_COUNTER_VAR=!PROP_COUNTER_VAR!
echo [DEBUG] PROP_MRU_VAR_ADAL=!PROP_MRU_VAR_ADAL!  
echo [DEBUG] PROP_MRU_VAR_LIVEID=!PROP_MRU_VAR_LIVEID!  
)

if not defined MRU_TARGET_PATHS (
    if defined MRU_FALLBACK (
        set "MRU_TARGET_PATHS=!MRU_FALLBACK!"
    )
)

if defined MRU_TARGET_PATHS (
    set "MRU_TARGET_PATHS=!MRU_TARGET_PATHS:"="!"
)

if /I "%IsDesignModeEnabled%"=="true" echo [DEBUG] Rutas MRU a actualizar: !MRU_TARGET_PATHS!

set "REG_VALUE=Item 1"
set "REG_DATA=[F00000000][T01ED6D7E58D00000][O00000000]*%FULL_PATH%"

set "META_VALUE=Item Metadata 1"
set "META_DATA=<Metadata><AppSpecific><id>%FULL_PATH%</id><nm>%BASENAME%</nm><du>%FULL_PATH%</du></AppSpecific></Metadata>"

for %%T in ("!PROP_MRU_VAR_ADAL!" "!PROP_MRU_VAR_LIVEID!") do (

    set "CURRENT_TARGET=%%~T"

    if "!CURRENT_TARGET!"=="" (
        if /I "%IsDesignModeEnabled%"=="true" echo [DEBUG] Target vacío, salto.
    ) else (

        for %%N in ("!FILE_NAME!") do set "BASENAME=%%~nN"

        set "META_DATA=<Metadata><AppSpecific><id>!FULL_PATH!</id><nm>!BASENAME!</nm><du>!FULL_PATH!</du></AppSpecific></Metadata>"

        set "CURRENT_REG_VALUE=!REG_VALUE!"
        set "CURRENT_META_VALUE="
        set "NEEDS_SHIFT=1"
        set "EXISTING_ITEM="
        set "EXISTING_META="

        call :FindExistingMRUEntry "!CURRENT_TARGET!" "!FULL_PATH!" EXISTING_ITEM EXISTING_META

        if defined EXISTING_ITEM (
            set "CURRENT_REG_VALUE=!EXISTING_ITEM!"
            set "NEEDS_SHIFT=0"
            if /I "%IsDesignModeEnabled%"=="true" echo [DEBUG] Entrada MRU existente encontrada en "!CURRENT_TARGET!".
        )

        if defined EXISTING_META (
            set "CURRENT_META_VALUE=!EXISTING_META!"
        )

        if "!NEEDS_SHIFT!"=="0" (
            if not defined CURRENT_META_VALUE (
                set "ITEM_INDEX="
                for /f "tokens=2 delims= " %%# in ("!CURRENT_REG_VALUE!") do if not defined ITEM_INDEX set "ITEM_INDEX=%%#"
                if not defined ITEM_INDEX (
                    for /f "tokens=3 delims= " %%# in ("!CURRENT_REG_VALUE!") do if not defined ITEM_INDEX set "ITEM_INDEX=%%#"
                )
                if defined ITEM_INDEX set "CURRENT_META_VALUE=Item Metadata !ITEM_INDEX!"
            )
        ) else (
            call :ShiftMRUEntries "!PROP_REG_NAME!" "!CURRENT_TARGET!" "%IsDesignModeEnabled%" "!LOCAL_LOGGING!" "!LOG_FILE!"
            set "CURRENT_META_VALUE=!META_VALUE!"
        )

        if not defined CURRENT_META_VALUE set "CURRENT_META_VALUE=!META_VALUE!"

        if /I "%IsDesignModeEnabled%"=="true" echo [DEBUG] Escribiendo !CURRENT_REG_VALUE! en "!CURRENT_TARGET!"
        reg add "!CURRENT_TARGET!" /v "!CURRENT_REG_VALUE!" /t REG_SZ /d "!REG_DATA!" /f >nul 2>&1

        if /I "%IsDesignModeEnabled%"=="true" echo [DEBUG] Escribiendo !CURRENT_META_VALUE! en "!CURRENT_TARGET!"
        reg add "!CURRENT_TARGET!" /v "!CURRENT_META_VALUE!" /t REG_SZ /d "!META_DATA!" /f >nul 2>&1

        set "TARGET_APPLIED=1"
    )
)


if "!TARGET_APPLIED!"=="0" (
    if /I "%IsDesignModeEnabled%"=="true" echo [WARNING] No se encontraron rutas MRU donde registrar "!FILE_NAME!".
)
exit /b

:ResolveAppProperties

set "RAP_APP=%~1"
set "PROP_MRU_VAR_ADAL="
set "PROP_MRU_VAR_LIVEID="
set "PROP_COUNTER_VAR="

if /I "%RAP_APP%"=="WORD" (
    set "PROP_MRU_VAR_ADAL=%WORD_MRU_ADAL%"
    set "PROP_MRU_VAR_LIVEID=%WORD_MRU_LIVEID%"
    set "PROP_COUNTER_VAR=%GLOBAL_ITEM_COUNT_WORD%"
) else if /I "%RAP_APP%"=="POWERPOINT" (
    set "PROP_MRU_VAR_ADAL=%PPT_MRU_ADAL%"
    set "PROP_MRU_VAR_LIVEID=%PPT_MRU_LIVEID%"
    set "PROP_COUNTER_VAR=%GLOBAL_ITEM_COUNT_PPT%"
) else if /I "%RAP_APP%"=="EXCEL" (
    set "PROP_MRU_VAR_ADAL=%EXCEL_MRU_ADAL%"
    set "PROP_MRU_VAR_LIVEID=%EXCEL_MRU_LIVEID%"
    set "PROP_COUNTER_VAR=%GLOBAL_ITEM_COUNT_EXCEL%"
)

exit /b 0

:FindExistingMRUEntry
setlocal EnableDelayedExpansion
set "FIND_MRU_PATH=%~1"
set "FIND_TARGET=%~2"
set "OUT_VALUE=%~3"
set "OUT_META=%~4"
set "FOUND_VALUE="
set "FOUND_META="

if not defined FIND_MRU_PATH goto :find_mru_exit
if not defined FIND_TARGET goto :find_mru_exit

for /f "skip=2 tokens=* delims=" %%L in ('reg query "!FIND_MRU_PATH!" 2^>nul') do (
    set "LINE=%%L"
    if defined LINE (
        set "LINE_TO_SEARCH=!LINE!"
        set "TARGET_TO_SEARCH=!FIND_TARGET!"
        call :EscapeForCmd LINE_TO_SEARCH
        call :EscapeForCmd TARGET_TO_SEARCH
        echo(!LINE_TO_SEARCH!| findstr /I /C:"!TARGET_TO_SEARCH!" >nul
        if not errorlevel 1 (
            set "VALUE_NAME_RAW="
            set "WORK_LINE=!LINE:REG_SZ=|!"
            if not "!WORK_LINE!"=="!LINE!" (
                for /f "tokens=1 delims=|" %%P in ("!WORK_LINE!") do set "VALUE_NAME_RAW=%%P"
                call :TrimWhitespaceVar VALUE_NAME_RAW
                if defined VALUE_NAME_RAW (
                    if /I "!VALUE_NAME_RAW:~0,13!"=="Item Metadata" (
                        if not defined FOUND_META set "FOUND_META=!VALUE_NAME_RAW!"
                    ) else if /I "!VALUE_NAME_RAW:~0,4!"=="Item" (
                        if not defined FOUND_VALUE set "FOUND_VALUE=!VALUE_NAME_RAW!"
                    )
                )
            )
        )
        if defined FOUND_VALUE if defined FOUND_META goto :find_mru_exit
    )
)

:find_mru_exit
endlocal & (
    if not "%OUT_VALUE%"=="" set "%OUT_VALUE%=%FOUND_VALUE%"
    if not "%OUT_META%"=="" set "%OUT_META%=%FOUND_META%"
)
exit /b 0

:EscapeForCmd
rem Args: VAR_NAME
setlocal EnableDelayedExpansion
set "ESC_VAR=%~1"
if "%ESC_VAR%"=="" (
    endlocal
    exit /b 0
)
set "ESC_VALUE="
for /f "tokens=1* delims==" %%A in ('set %ESC_VAR% 2^>nul') do (
    if /I "%%A"=="%ESC_VAR%" set "ESC_VALUE=%%B"
)
if not defined ESC_VALUE (
    endlocal
    exit /b 0
)
set "ESC_VALUE=!ESC_VALUE:^=^^!"
set "ESC_VALUE=!ESC_VALUE:|=^|!"
set "ESC_VALUE=!ESC_VALUE:&=^&!"
set "ESC_VALUE=!ESC_VALUE:<=^<!"
set "ESC_VALUE=!ESC_VALUE:>=^>!"
set "ESC_VALUE=!ESC_VALUE:(=^(!"
set "ESC_VALUE=!ESC_VALUE:)=^)!"
endlocal & set "%ESC_VAR%=%ESC_VALUE%"
exit /b 0

:ShiftMRUEntries
setlocal EnableDelayedExpansion
set "APP_KEY=%~1"
set "TARGET_MRU=%~2"
set "DESIGN_MODE=%~3"
set "LOCAL_LOGGING=%~4"
set "LOG_FILE=%~5"
set "OFFSET=1"

if not defined TARGET_MRU (
    endlocal
    exit /b 0
)

if /I "%DESIGN_MODE%"=="true" echo [DEBUG] Ajustando índices MRU para %APP_KEY%...

set "TMP_FILE=%TEMP%\mru_shift_%RANDOM%.txt"
if exist "!TMP_FILE!" del "!TMP_FILE!" >nul 2>&1

set "FOUND_VALUES="

for /f "skip=2 tokens=* delims=" %%L in ('reg query "!TARGET_MRU!" 2^>nul') do (
    set "LINE=%%L"
    if not "!LINE!"=="" (
        set "HASREG=!LINE:REG_SZ=!"
        if not "!HASREG!"=="!LINE!" (
            set "WORK_LINE=!LINE:REG_SZ=|!"
            for /f "tokens=1 delims=|" %%P in ("!WORK_LINE!") do set "VALUE_NAME_RAW=%%P"
            call :TrimWhitespaceVar VALUE_NAME_RAW
            if defined VALUE_NAME_RAW (
                set "FIRST="
                set "SECOND="
                set "THIRD="
                for /f "tokens=1-3" %%a in ("!VALUE_NAME_RAW!") do (
                    if not defined FIRST set "FIRST=%%a"
                    if not defined SECOND set "SECOND=%%b"
                    if not defined THIRD set "THIRD=%%c"
                )
                set "BASE="
                set "INDEX="
                if /I "!FIRST!"=="Item" (
                    if /I "!SECOND!"=="Metadata" (
                        set "BASE=Item Metadata"
                        set "INDEX=!THIRD!"
                    ) else (
                        set "BASE=Item"
                        set "INDEX=!SECOND!"
                    )
                )
                if defined INDEX (
                    echo(!INDEX!| findstr /R "^[0-9][0-9]*$" >nul
                    if not errorlevel 1 (
                        set "FOUND_VALUES=1"
                        set "PAD=0000000000!INDEX!"
                        set "PAD=!PAD:~-10!"
                        >>"!TMP_FILE!" echo(!PAD!^|!VALUE_NAME_RAW!
                    )
                )
            )
        )
    )
)

if not defined FOUND_VALUES (
    if /I "%DESIGN_MODE%"=="true" echo [DEBUG] No se encontraron entradas MRU previas para %APP_KEY%.
    if exist "!TMP_FILE!" del "!TMP_FILE!" >nul 2>&1
    endlocal
    exit /b 0
)

for /f "usebackq tokens=1* delims=|" %%A in (`sort /R "!TMP_FILE!"`) do (
    call :ShiftMRURename "%%B" "%OFFSET%" "!TARGET_MRU!" "%DESIGN_MODE%" "%LOCAL_LOGGING%" "%LOG_FILE%" "%APP_KEY%"
)

if exist "!TMP_FILE!" del "!TMP_FILE!" >nul 2>&1

if /I "%DESIGN_MODE%"=="true" echo [DEBUG] Reindexado MRU completado para %APP_KEY%.

endlocal
exit /b 0


:ShiftMRURename
setlocal EnableDelayedExpansion
set "ORIGINAL_NAME=%~1"
set "OFFSET=%~2"
set "MRU_PATH=%~3"
set "DESIGN_MODE=%~4"
set "LOCAL_LOGGING=%~5"
set "LOG_FILE=%~6"
set "APP_KEY=%~7"

if "%ORIGINAL_NAME%"=="" (
    endlocal
    exit /b 0
)

set "FIRST="
set "SECOND="
set "THIRD="
for /f "tokens=1-3" %%a in ("!ORIGINAL_NAME!") do (
    if not defined FIRST set "FIRST=%%a"
    if not defined SECOND set "SECOND=%%b"
    if not defined THIRD set "THIRD=%%c"
)

set "BASE="
set "INDEX="
if /I "!FIRST!"=="Item" (
    if /I "!SECOND!"=="Metadata" (
        set "BASE=Item Metadata"
        set "INDEX=!THIRD!"
    ) else (
        set "BASE=Item"
        set "INDEX=!SECOND!"
    )
)

if not defined INDEX (
    endlocal
    exit /b 0
)

set /a NEW_INDEX=INDEX+OFFSET
if /I "!BASE!"=="Item Metadata" (
    set "NEW_NAME=Item Metadata !NEW_INDEX!"
) else (
    set "NEW_NAME=Item !NEW_INDEX!"
)

set "DATA_LINE="
for /f "skip=2 tokens=* delims=" %%L in ('reg query "!MRU_PATH!" /v "!ORIGINAL_NAME!" 2^>nul') do set "DATA_LINE=%%L"
if not defined DATA_LINE (
    endlocal
    exit /b 0
)

set "DATA_LINE=!DATA_LINE:*REG_SZ=!"
call :TrimWhitespaceVar DATA_LINE
set "DATA=!DATA_LINE!"

if /I "%DESIGN_MODE%"=="true" echo [DEBUG] Renombrando "!ORIGINAL_NAME!" a "!NEW_NAME!" en "!MRU_PATH!" para %APP_KEY%.

reg add "!MRU_PATH!" /v "!NEW_NAME!" /t REG_SZ /d "!DATA!" /f >nul
reg delete "!MRU_PATH!" /v "!ORIGINAL_NAME!" /f >nul

endlocal
exit /b 0

:TrimWhitespaceVar
setlocal EnableDelayedExpansion
set "VALUE=!%~1!"
:TrimLeadingWS
if defined VALUE if "!VALUE:~0,1!"==" " (
    set "VALUE=!VALUE:~1!"
    goto :TrimLeadingWS
)
:TrimTrailingWS
if defined VALUE if "!VALUE:~-1!"==" " (
    set "VALUE=!VALUE:~0,-1!"
    goto :TrimTrailingWS
)
endlocal & set "%~1=%VALUE%"
exit /b 0

:EndOfScript
if /I "%IsDesignModeEnabled%"=="true" (
    echo [DEBUG] Entering EndOfScript finalizer...
    pause
)
echo Ready
endlocal
exit /b 0

