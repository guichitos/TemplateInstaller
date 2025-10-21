@echo off
rem ======================================================
rem === CUSTOM TEMPLATES COPY LIBRARY (v2025.10.17)
rem ------------------------------------------------------
rem Copies additional Office templates (.dotx, .potx, .xltx, etc.)
rem from the installer’s base directory to all valid user
rem "Custom Office Templates" folders (Documents, OneDrive, etc.).
rem Excludes GenericTemplate.* files, logs operations, and
rem optionally integrates with registry_tools.bat to simulate
rem registry entries for PowerPoint templates.
rem ======================================================

if "%~1"=="" exit /b
set "FUNC=%~1"
shift
goto %FUNC%


:CopyAll
rem Args: LOG_FILE BASE_DIR REG_LIB
set "LOG_FILE=%~1"
set "BASE_DIR=%~2"
set "REG_LIB=%~3"

setlocal enabledelayedexpansion
set /a TOTAL_FILES=0
set /a TOTAL_FOLDERS=0
set "FOUND_ANY=0"

rem === Buscar carpetas de destino del usuario ===
for /d %%O in ("%USERPROFILE%\OneDrive*") do (
  for %%S in ("Documents" "Documentos") do (
    set "CANDIDATE=%%~fO\%%~S\Custom Office Templates"
    if exist "!CANDIDATE!" (
      call :CopyCustomTemplates "!CANDIDATE!" "!LOG_FILE!" "!BASE_DIR!" "!REG_LIB!"
      set "FOUND_ANY=1"
      set /a TOTAL_FOLDERS+=1
    )
  )
)

for %%L in ("%USERPROFILE%\Documents\Custom Office Templates" "%USERPROFILE%\Documentos\Custom Office Templates") do (
  if exist "%%~fL" (
    call :CopyCustomTemplates "%%~fL" "!LOG_FILE!" "!BASE_DIR!" "!REG_LIB!"
    set "FOUND_ANY=1"
    set /a TOTAL_FOLDERS+=1
  )
)

if "!FOUND_ANY!"=="0" (
  set "DEFAULT_DIR=%USERPROFILE%\Documents\Custom Office Templates"
  mkdir "!DEFAULT_DIR!" 2>nul
  call :CopyCustomTemplates "!DEFAULT_DIR!" "!LOG_FILE!" "!BASE_DIR!" "!REG_LIB!"
  set /a TOTAL_FOLDERS+=1
)

echo Total folders updated: !TOTAL_FOLDERS! >> "!LOG_FILE!"
echo Total templates copied: !TOTAL_FILES! >> "!LOG_FILE!"
endlocal
exit /b


:CopyCustomTemplates
rem Args: TARGET_DIR LOG_FILE BASE_DIR REG_LIB
set "TARGET_DIR=%~1"
set "LOG_FILE=%~2"
set "BASE_DIR=%~3"
set "REG_LIB=%~4"

setlocal enabledelayedexpansion
echo Copying templates to: "!TARGET_DIR!" >> "!LOG_FILE!"
mkdir "!TARGET_DIR!" 2>nul

for %%F in ("%BASE_DIR%*.dotx" "%BASE_DIR%*.dotm" "%BASE_DIR%*.potx" "%BASE_DIR%*.potm" "%BASE_DIR%*.xltx" "%BASE_DIR%*.xltm") do (
  if exist "%%~fF" (
    rem === Excluir plantillas genéricas ===
    if /I not "%%~nxF"=="GenericTemplate.dotm" if /I not "%%~nxF"=="GenericTemplate.potx" if /I not "%%~nxF"=="GenericTemplate.xltx" (
      copy /Y "%%~fF" "!TARGET_DIR!" >nul
      if exist "!TARGET_DIR!\%%~nxF" (
        echo Copied %%~nxF to "!TARGET_DIR!" >> "!LOG_FILE!"
        set /a TOTAL_FILES+=1

        rem === Si es PowerPoint y hay registro disponible ===
        if exist "!REG_LIB!" (
          if /I "%%~xF"==".potx" (
            call "!REG_LIB!" :SimulateRegEntry "%%~nxF" "!TARGET_DIR!\%%~nxF!" "!LOG_FILE!"
          )
          if /I "%%~xF"==".potm" (
            call "!REG_LIB!" :SimulateRegEntry "%%~nxF" "!TARGET_DIR!\%%~nxF!" "!LOG_FILE!"
          )
        )
      ) else (
        echo ERROR: Failed to copy %%~nxF to "!TARGET_DIR!" >> "!LOG_FILE!"
      )
    )
  )
)

rem === Crear marcador de control ===
if "%DEBUG_MODE%"=="1" (
  echo Custom templates marker > "!TARGET_DIR!\prueba_custom.txt"
  echo Created prueba_custom.txt in "!TARGET_DIR!" >> "!LOG_FILE!"
)

endlocal
exit /b
