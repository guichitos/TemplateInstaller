@echo off
setlocal enabledelayedexpansion
chcp 65001 >nul

rem ======================================================
rem === UNIVERSAL OFFICE TEMPLATES INSTALLER (v2025.10.15)
rem ======================================================

set "SRC_DIR=%~dp0"
set "LOG_FILE=%SRC_DIR%install_log_all.txt"
echo. > "%LOG_FILE%"
echo [%DATE% %TIME%] --- START UNIVERSAL INSTALLATION --- >> "%LOG_FILE%"

rem === Close Office apps ===
echo Closing Word, PowerPoint, and Excel if open... >> "%LOG_FILE%"
taskkill /IM WINWORD.EXE /F >nul 2>&1
taskkill /IM POWERPNT.EXE /F >nul 2>&1
taskkill /IM EXCEL.EXE /F >nul 2>&1

rem ======================================================
rem === INSTALL WORD TEMPLATE
rem ======================================================
call :InstallApp "WORD" "GenericTemplate.dotm" "%APPDATA%\Microsoft\Templates" "Normal.dotm"

rem ======================================================
rem === INSTALL POWERPOINT TEMPLATE
rem ======================================================
call :InstallApp "POWERPOINT" "GenericTemplate.potx" "%APPDATA%\Microsoft\Templates" "Blank.potx"

rem ======================================================
rem === INSTALL EXCEL TEMPLATE
rem ======================================================
call :InstallApp "EXCEL" "GenericTemplate.xltx" "%APPDATA%\Microsoft\Excel\XLSTART" "Book.xltx"

rem ======================================================
rem === COPY CUSTOM TEMPLATES FOR ALL APPLICATIONS
rem ======================================================
echo. >> "%LOG_FILE%"
echo [%DATE% %TIME%] --- COPYING CUSTOM TEMPLATES (.dot*, .pot*, .xlt*) --- >> "%LOG_FILE%"

set "FOUND_ANY=0"
set /a TOTAL_FILES=0
set /a TOTAL_FOLDERS=0

for /d %%O in ("%USERPROFILE%\OneDrive*") do (
  for %%S in ("Documents" "Documentos") do (
    set "CANDIDATE=%%~fO\%%~S\Custom Office Templates"
    if exist "!CANDIDATE!" (
      call :CopyCustomTemplates "!CANDIDATE!"
      set "FOUND_ANY=1"
      set /a TOTAL_FOLDERS+=1
    )
  )
)

for %%L in ("%USERPROFILE%\Documents\Custom Office Templates" "%USERPROFILE%\Documentos\Custom Office Templates") do (
  if exist "%%~fL" (
    call :CopyCustomTemplates "%%~fL"
    set "FOUND_ANY=1"
    set /a TOTAL_FOLDERS+=1
  )
)

if "!FOUND_ANY!"=="0" (
  set "DEFAULT_DIR=%USERPROFILE%\Documents\Custom Office Templates"
  mkdir "!DEFAULT_DIR!" 2>nul
  call :CopyCustomTemplates "!DEFAULT_DIR!"
  set /a TOTAL_FOLDERS+=1
)

echo Total folders updated: !TOTAL_FOLDERS! >> "%LOG_FILE%"
echo Total templates copied: !TOTAL_FILES! >> "%LOG_FILE%"
echo. >> "%LOG_FILE%"

echo [%DATE% %TIME%] --- UNIVERSAL INSTALLATION COMPLETED --- >> "%LOG_FILE%"
endlocal
pause
exit /b


:InstallApp
set "APP=%~1"
set "SRC_FILE=%SRC_DIR%%~2"
set "DST_DIR=%~3"
set "DST_FILE=%DST_DIR%\%~4"
set "BACKUP_FILE=%DST_DIR%\%~n4_backup%~x4"

echo. >> "%LOG_FILE%"
echo --- [%APP% INSTALLATION] --- >> "%LOG_FILE%"
echo Source: "%SRC_FILE%" >> "%LOG_FILE%"
echo Destination folder: "%DST_DIR%" >> "%LOG_FILE%"

if not exist "%SRC_FILE%" (
  echo ERROR: Source file not found "%SRC_FILE%". >> "%LOG_FILE%"
  goto :eof
)

if not exist "%DST_DIR%" (
  mkdir "%DST_DIR%" 2>nul
  echo Created destination folder "%DST_DIR%". >> "%LOG_FILE%"
)

if exist "%DST_FILE%" (
  copy /Y "%DST_FILE%" "%BACKUP_FILE%" >nul
  echo Backup created: "%BACKUP_FILE%" >> "%LOG_FILE%"
) else (
  echo No existing template to back up. >> "%LOG_FILE%"
)

if exist "%DST_FILE%" (
  del /F /Q "%DST_FILE%"
  if exist "%DST_FILE%" (
    echo ERROR: Could not delete old file. >> "%LOG_FILE%"
    goto :eof
  )
)

copy /Y "%SRC_FILE%" "%DST_FILE%" >nul
if exist "%DST_FILE%" (
  echo Installed new template "%DST_FILE%". >> "%LOG_FILE%"
) else (
  echo ERROR: Copy failed for "%SRC_FILE%". >> "%LOG_FILE%"
)

echo Origin marker > "%SRC_DIR%prueba_origen_%APP%.txt"
echo Destination marker > "%DST_DIR%\prueba_destino_%APP%.txt"
echo Created prueba_origen_%APP%.txt and prueba_destino_%APP%.txt >> "%LOG_FILE%"
goto :eof


:CopyCustomTemplates
set "TARGET_DIR=%~1"
echo Copying templates to: "!TARGET_DIR!" >> "%LOG_FILE%"
mkdir "!TARGET_DIR!" 2>nul

for %%F in ("%SRC_DIR%*.dotx" "%SRC_DIR%*.dotm" "%SRC_DIR%*.potx" "%SRC_DIR%*.potm" "%SRC_DIR%*.xltx" "%SRC_DIR%*.xltm") do (
  if exist "%%~fF" (
    set "FN=%%~nxF"
    if /I not "!FN!"=="Normal.dotx" if /I not "!FN!"=="Normal.dotm" if /I not "!FN!"=="Blank.potx" if /I not "!FN!"=="Blank.potm" if /I not "!FN!"=="Book.xltx" if /I not "!FN!"=="Book.xltm" if /I not "!FN!"=="Sheet.xltx" if /I not "!FN!"=="Sheet.xltm" (
      copy /Y "%%~fF" "!TARGET_DIR!" >nul
      if exist "!TARGET_DIR!\%%~nxF" (
        echo Copied %%~nxF to "!TARGET_DIR!" >> "%LOG_FILE%"
        set /a TOTAL_FILES+=1
      ) else (
        echo ERROR: Failed to copy %%~nxF to "!TARGET_DIR!" >> "%LOG_FILE%"
      )
    )
  )
)

echo Custom templates marker > "!TARGET_DIR!\prueba_custom.txt"
echo Created prueba_custom.txt in "!TARGET_DIR!" >> "%LOG_FILE%"
goto :eof
