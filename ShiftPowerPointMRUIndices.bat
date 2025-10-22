@echo off
setlocal EnableDelayedExpansion

set "OFFSET=1"

if not defined PPT_MRU_PATH call :DetectPowerPointMRUPath
if not defined PPT_MRU_PATH (
  echo [ERROR] No se pudo detectar la clave MRU de PowerPoint.
  exit /b 1
)

echo [INFO] Ajustando valores en "%PPT_MRU_PATH%" con desplazamiento %OFFSET%.

set "TMP_FILE=%TEMP%\ppt_shift_%RANDOM%.txt"
if exist "%TMP_FILE%" del "%TMP_FILE%" >nul 2>&1

set "KNOWN_TYPES=REG_SZ REG_EXPAND_SZ REG_MULTI_SZ REG_BINARY REG_DWORD REG_QWORD REG_NONE"

for /f "skip=2 tokens=* delims=" %%L in ('reg query "%PPT_MRU_PATH%" 2^>nul') do (
  set "LINE=%%L"
  if not "!LINE!"=="" (
    set "VALUE_TYPE="
    set "VALUE_NAME_RAW="
    for %%T in (!KNOWN_TYPES!) do (
      if not defined VALUE_TYPE (
        echo !LINE! | findstr /C:"%%T" >nul
        if not errorlevel 1 (
          set "VALUE_TYPE=%%T"
          for /f "tokens=1* delims=%%T" %%P in ("!LINE!") do set "VALUE_NAME_RAW=%%P"
        )
      )
    )
    if defined VALUE_TYPE (
      call :Trim VALUE_NAME_RAW
      set "T1="
      set "T2="
      set "T3="
      for /f "tokens=1-3" %%a in ("!VALUE_NAME_RAW!") do (
        if not defined T1 (
          set "T1=%%a"
        ) else if not defined T2 (
          set "T2=%%a"
        ) else if not defined T3 (
          set "T3=%%a"
        )
      )
      if /I "!T1!"=="Item" (
        if defined T3 (
          set "BASE=!T1! !T2!"
          set "INDEX=!T3!"
        ) else (
          set "BASE=!T1!"
          set "INDEX=!T2!"
        )
        echo(!INDEX!| findstr /R "^[0-9][0-9]*$" >nul
        if not errorlevel 1 (
          set "PAD=0000000000!INDEX!"
          set "PAD=!PAD:~-10!"
          >>"%TMP_FILE%" echo(!PAD!^|!VALUE_NAME_RAW!^|!VALUE_TYPE!
        )
      )
    )
  )
)

if not exist "%TMP_FILE%" (
  echo [INFO] No se encontraron valores "Item" para ajustar.
  exit /b 0
)

echo [INFO] Reetiquetando valores existentes...

for /f "usebackq tokens=1-3 delims=|" %%A in (`sort /R "%TMP_FILE%"`) do (
  set "CURRENT_VALUE=%%B"
  set "CURRENT_TYPE=%%C"
  call :ShiftValue "!CURRENT_VALUE!" "!CURRENT_TYPE!"
)

del "%TMP_FILE%" >nul 2>&1

echo [INFO] Proceso finalizado.
exit /b 0

:ShiftValue
setlocal EnableDelayedExpansion
set "ORIGINAL_NAME=%~1"
set "ORIGINAL_TYPE=%~2"
set "T1="
set "T2="
set "T3="
for /f "tokens=1-3" %%a in ("!ORIGINAL_NAME!") do (
  if not defined T1 set "T1=%%a" else if not defined T2 set "T2=%%a" else if not defined T3 set "T3=%%a"
)
if defined T3 (
  set "BASE=!T1! !T2!"
  set "INDEX=!T3!"
) else (
  set "BASE=!T1!"
  set "INDEX=!T2!"
)
set /a NEW_INDEX=INDEX+OFFSET
set "NEW_NAME=!BASE! !NEW_INDEX!"

set "CSV_NAME="
set "CSV_TYPE="
set "CSV_DATA="
for /f "skip=1 tokens=1,2* delims=," %%A in ('reg query "%PPT_MRU_PATH%" /v "!ORIGINAL_NAME!" /fo csv 2^>nul') do (
  if not defined CSV_NAME set "CSV_NAME=%%~A"
  if not defined CSV_TYPE set "CSV_TYPE=%%~B"
  if not defined CSV_DATA set "CSV_DATA=%%~C"
)
if not defined CSV_NAME (
  echo [WARN] No se pudo leer el valor "!ORIGINAL_NAME!". Se omite.
  exit /b 0
)
set "VALUE_TYPE=!ORIGINAL_TYPE!"
if not defined VALUE_TYPE set "VALUE_TYPE=!CSV_TYPE!"
if not defined VALUE_TYPE set "VALUE_TYPE=REG_SZ"
set "DATA=!CSV_DATA!"
call :Trim VALUE_TYPE
call :Trim DATA

reg add "%PPT_MRU_PATH%" /v "!NEW_NAME!" /t !VALUE_TYPE! /d "!DATA!" /f >nul 2>&1
if errorlevel 1 (
  echo [ERROR] No se pudo crear "!NEW_NAME!".
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
for /f "tokens=*" %%a in ("!VALUE!") do set "VALUE=%%a"
:TrimLoop
if "!VALUE!"=="" goto :TrimDone
if not "!VALUE:~-1!"==" " goto :TrimDone
set "VALUE=!VALUE:~0,-1!"
goto :TrimLoop
:TrimDone
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
