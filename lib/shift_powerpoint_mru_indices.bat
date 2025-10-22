@echo off
setlocal EnableDelayedExpansion

rem ------------------------------------------------------
rem Script de depuracion para desplazar indices de MRU PPT
rem ------------------------------------------------------
set "DEBUG_MODE=true"
set "OFFSET=2"

echo ======================================================
echo [DEBUG] Inicio del script shift_powerpoint_mru_indices
if /I "%DEBUG_MODE%"=="true" (
  echo [DEBUG] OFFSET actual: %OFFSET%
)
echo ======================================================
pause

echo [INFO] Detectando clave MRU de PowerPoint...
if not defined PPT_MRU_PATH call :DetectPowerPointMRUPath
if not defined PPT_MRU_PATH (
  echo [ERROR] No se pudo detectar la clave MRU de PowerPoint.
  pause
  exit /b 1
)

echo [DEBUG] Ruta detectada: "%PPT_MRU_PATH%"
pause

set "TMP_FILE=%TEMP%\ppt_shift_%RANDOM%.txt"
if exist "%TMP_FILE%" del "%TMP_FILE%" >nul 2>&1

set "FOUND_VALUES="
echo [INFO] Recolectando valores existentes...
for /f "tokens=1,* delims=:" %%A in ('reg query "%PPT_MRU_PATH%" /z 2^>nul ^| findstr /C:"Value Name"') do (
  set "VALUE_NAME_RAW=%%B"
  call :Trim VALUE_NAME_RAW
  if defined VALUE_NAME_RAW (
    if /I "!VALUE_NAME_RAW!"=="(Default)" (
      if /I "%DEBUG_MODE%"=="true" echo [DEBUG] Se omite valor predeterminado.
    ) else (
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
          >>"%TMP_FILE%" echo(!PAD!^|!VALUE_NAME_RAW!
          if /I "%DEBUG_MODE%"=="true" echo [DEBUG] Valor localizado: !VALUE_NAME_RAW! (indice !INDEX!)
        ) else (
          if /I "%DEBUG_MODE%"=="true" echo [DEBUG] Se omite "!VALUE_NAME_RAW!" (indice no numerico).
        )
      ) else (
        if /I "%DEBUG_MODE%"=="true" echo [DEBUG] Se omite "!VALUE_NAME_RAW!" (sin indice).
      )
    )
  )
)

if not defined FOUND_VALUES (
  echo [INFO] No se encontraron valores "Item" para ajustar.
  pause
  exit /b 0
)

echo ------------------------------------------------------
echo [DEBUG] Contenido recolectado en "%TMP_FILE%":
type "%TMP_FILE%"
echo ------------------------------------------------------
pause

echo [INFO] Reetiquetando valores existentes con desplazamiento %OFFSET%...
pause

for /f "usebackq tokens=1* delims=|" %%A in (`sort /R "%TMP_FILE%"`) do (
  set "CURRENT_VALUE=%%B"
  call :ShiftValue "!CURRENT_VALUE!"
)

del "%TMP_FILE%" >nul 2>&1

echo [INFO] Proceso finalizado.
pause
exit /b 0

:ShiftValue
setlocal EnableDelayedExpansion
set "ORIGINAL_NAME=%~1"
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
  echo [WARN] No se pudo interpretar el nombre "!ORIGINAL_NAME!". Se omite.
  endlocal
  exit /b 0
)
set /a NEW_INDEX=INDEX+OFFSET
set "NEW_NAME=!BASE! !NEW_INDEX!"

echo [DEBUG] Renombrando "!ORIGINAL_NAME!" -> "!NEW_NAME!"

set "DATA_LINE="
for /f "skip=2 tokens=* delims=" %%L in ('reg query "%PPT_MRU_PATH%" /v "!ORIGINAL_NAME!" 2^>nul') do set "DATA_LINE=%%L"
if not defined DATA_LINE (
  echo [WARN] No se pudo leer el valor "!ORIGINAL_NAME!". Se omite.
  endlocal
  exit /b 0
)
set "DATA_LINE=!DATA_LINE:*REG_SZ=!"
call :Trim DATA_LINE
set "DATA=!DATA_LINE!"

if /I "%DEBUG_MODE%"=="true" (
  echo [DEBUG] Datos capturados: "!DATA!"
)

reg add "%PPT_MRU_PATH%" /v "!NEW_NAME!" /t REG_SZ /d "!DATA!" /f >nul 2>&1
if errorlevel 1 (
  echo [ERROR] No se pudo crear "!NEW_NAME!".
  endlocal
  exit /b 0
)
reg delete "%PPT_MRU_PATH%" /v "!ORIGINAL_NAME!" /f >nul 2>&1
if errorlevel 1 (
  echo [WARN] No se pudo eliminar "!ORIGINAL_NAME!" tras copiarlo.
) else (
  echo [OK] "!ORIGINAL_NAME!" => "!NEW_NAME!"
)
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
if /I "%DEBUG_MODE%"=="true" (
  echo [DEBUG] DetectPowerPointMRUPath => "!PPT_MRU_PATH!"
)
endlocal & set "PPT_MRU_PATH=%PPT_MRU_PATH%"
exit /b 0
