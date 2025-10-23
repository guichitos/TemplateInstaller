@echo off
setlocal EnableDelayedExpansion
chcp 65001 >nul

rem ======================================================
rem === SHIFT POWERPOINT MRU INDICES (versión limpia) ===
rem ======================================================
set "OFFSET=1"

echo ======================================================
echo [INFO] Iniciando ajuste de índices MRU de PowerPoint
echo [INFO] Desplazamiento configurado: %OFFSET%
echo ======================================================

echo [INFO] Detectando clave MRU de PowerPoint...
if not defined PPT_MRU_PATH call :DetectPowerPointMRUPath
if not defined PPT_MRU_PATH (
  echo [ERROR] No se pudo detectar la clave MRU de PowerPoint.
  exit /b 1
)
echo [INFO] Clave MRU detectada: "%PPT_MRU_PATH%"

set "TMP_FILE=%TEMP%\ppt_shift_%RANDOM%.txt"
if exist "%TMP_FILE%" del "%TMP_FILE%" >nul 2>&1

set "FOUND_VALUES="
echo [INFO] Recolectando valores existentes...
set /a LINECOUNT=0

rem === Leer el registro sin usar pipes ni findstr (para evitar errores con '<' o '&')
for /f "skip=2 tokens=* delims=" %%L in ('reg query "%PPT_MRU_PATH%" 2^>nul') do (
  set /a LINECOUNT+=1
  set "LINE=%%L"

  if not "!LINE!"=="" (
    set "HASREG=!LINE:REG_SZ=!"
    if not "!HASREG!"=="!LINE!" (
      rem Reemplazar REG_SZ por '|' para aislar nombre del valor
      set "WORK_LINE=!LINE:REG_SZ=|!"
      for /f "tokens=1 delims=|" %%P in ("!WORK_LINE!") do set "VALUE_NAME_RAW=%%P"
      call :Trim VALUE_NAME_RAW

      if defined VALUE_NAME_RAW (
        set "FIRST=" & set "SECOND=" & set "THIRD="
        for /f "tokens=1-3" %%a in ("!VALUE_NAME_RAW!") do (
          if not defined FIRST set "FIRST=%%a"
          if not defined SECOND set "SECOND=%%b"
          if not defined THIRD set "THIRD=%%c"
        )

        set "BASE=" & set "INDEX="
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
            >>"%TMP_FILE%" echo(!PAD!^|!VALUE_NAME_RAW!
          )
        )
      )
    )
  )
)

if not defined FOUND_VALUES (
  echo [INFO] No se encontraron valores "Item" para ajustar.
  exit /b 0
)

echo [INFO] Reetiquetando valores existentes con desplazamiento %OFFSET%...
for /f "usebackq tokens=1* delims=|" %%A in (`sort /R "%TMP_FILE%"`) do (
  call :ShiftValue "%%B"
)

del "%TMP_FILE%" >nul 2>&1
echo [OK] Proceso completado correctamente.
exit /b 0


:ShiftValue
setlocal EnableDelayedExpansion

if "%~1"=="" (
  endlocal
  exit /b 0
)

set "ORIGINAL_NAME=%~1"

rem === Separar en partes ===
setlocal DisableDelayedExpansion
for /f "tokens=1-3" %%a in ("%ORIGINAL_NAME%") do (
  endlocal
  set "FIRST=%%a"
  set "SECOND=%%b"
  set "THIRD=%%c"
)
setlocal EnableDelayedExpansion

rem === Determinar base e índice ===
set "BASE=" & set "INDEX="
if /I "!FIRST!"=="Item" (
  if /I "!SECOND!"=="Metadata" (
    set "BASE=Item Metadata"
    set "INDEX=!THIRD!"
  ) else (
    set "BASE=Item"
    set "INDEX=!SECOND!"
  )
)

rem === Validar que el índice sea numérico ===
set "TEST=!INDEX!"
for /f "delims=0123456789" %%Z in ("!TEST!") do set "NONNUM=%%Z"
if defined NONNUM (
  endlocal
  exit /b 0
)

set /a NEW_INDEX=INDEX+OFFSET
set "NEW_NAME=!BASE! !NEW_INDEX!"

echo [INFO] Procesando: "!ORIGINAL_NAME!" → "!NEW_NAME!"

rem === Borrar archivo asociado si existe ===
set "CURRENT_DIR=%~dp0"
set "FILE_PATH=%CURRENT_DIR%!NEW_NAME!"
if exist "!FILE_PATH!" del /f /q "!FILE_PATH!"

rem === Leer valor original del registro ===
set "DATA_LINE="
for /f "skip=2 tokens=* delims=" %%L in ('reg query "%PPT_MRU_PATH%" /v "!ORIGINAL_NAME!" 2^>nul') do set "DATA_LINE=%%L"

if not defined DATA_LINE (
  endlocal
  exit /b 0
)

set "DATA_LINE=!DATA_LINE:*REG_SZ=!"
call :Trim DATA_LINE
set "DATA=!DATA_LINE!"

rem === Crear nuevo valor y eliminar el antiguo ===
reg add "%PPT_MRU_PATH%" /v "!NEW_NAME!" /t REG_SZ /d "!DATA!" /f >nul
reg delete "%PPT_MRU_PATH%" /v "!ORIGINAL_NAME!" /f >nul

endlocal
exit /b 0


:Trim
setlocal EnableDelayedExpansion
set "VALUE=!%~1!"
:TrimLeading
if defined VALUE if "!VALUE:~0,1!"==" " (
  set "VALUE=!VALUE:~1!"
  goto :TrimLeading
)
:TrimTrailing
if defined VALUE if "!VALUE:~-1!"==" " (
  set "VALUE=!VALUE:~0,-1!"
  goto :TrimTrailing
)
endlocal & set "%~1=%VALUE%"
exit /b 0


:DetectPowerPointMRUPath
setlocal EnableDelayedExpansion
set "PPT_MRU_PATH="
set "FOUND_ID="
for %%V in (16.0 15.0) do (
  set "BASE=HKCU\Software\Microsoft\Office\%%V\PowerPoint\Recent Templates"
  for /f "tokens=*" %%K in ('reg query "!BASE!" 2^>nul ^| findstr /R /C:"ADAL_" /C:"Livelid_"') do (
    set "FOUND_ID=%%~nK"
    goto :found
  )
)
if not defined FOUND_ID (
  set "TMP=%TEMP%\adal_search_%RANDOM%.txt"
  >"%TMP%" 2>&1 reg query "HKCU\Software\Microsoft\Office" /f "ADAL_" /s
  findstr /i "ADAL_" "%TMP%" >"%TMP%.2" 2>nul
  for /f "usebackq delims=" %%L in ("%TMP%.2") do (
    set "FOUND_ID=%%~nL"
    goto :found
  )
  >"%TMP%" 2>&1 reg query "HKCU\Software\Microsoft\Office" /f "Livelid_" /s
  findstr /i "Livelid_" "%TMP%" >"%TMP%.2" 2>nul
  for /f "usebackq delims=" %%L in ("%TMP%.2") do (
    set "FOUND_ID=%%~nL"
    goto :found
  )
  del "%TMP%" "%TMP%.2" >nul 2>&1
)
:found
if defined FOUND_ID (
  set "PPT_MRU_PATH=HKCU\Software\Microsoft\Office\16.0\PowerPoint\Recent Templates\!FOUND_ID!\File MRU"
) else (
  set "PPT_MRU_PATH=HKCU\Software\Microsoft\Office\16.0\PowerPoint\Recent Templates\File MRU"
)
endlocal & set "PPT_MRU_PATH=%PPT_MRU_PATH%"
exit /b 0
