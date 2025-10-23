@echo off
if "%~1"=="" exit /b
set "FUNC=%~1"
shift
goto %FUNC%


:DetectExcelMRUPath
rem ------------------------------------------------------
rem Detecta la ruta MRU real de Excel (ADAL_ o Livelid_)
rem ------------------------------------------------------
setlocal enabledelayedexpansion
set "EXCEL_MRU_PATH="
set "FOUND_ID="

rem 1. Buscar en rutas más probables
for %%V in (16.0 15.0) do (
  set "BASE=HKCU\Software\Microsoft\Office\%%V\Excel\Recent Templates"
  for /f "tokens=*" %%K in ('reg query "!BASE!" 2^>nul ^| findstr /R /C:"ADAL_" /C:"Livelid_"') do (
    set "FOUND_ID=%%~nK"
    goto :found
  )
)

rem 2. Si no hay resultado, intentar búsqueda global en rama Office
if not defined FOUND_ID (
  set "TMP=%TEMP%\adal_search_excel_%RANDOM%.txt"
  > "%TMP%" 2>&1 reg query "HKCU\Software\Microsoft\Office" /f "ADAL_" /s
  findstr /i "ADAL_" "%TMP%" > "%TMP%.2" 2>nul
  for /f "usebackq delims=" %%L in ("%TMP%.2") do (
    set "FOUND_ID=%%~nL"
    goto :found
  )
  > "%TMP%" 2>&1 reg query "HKCU\Software\Microsoft\Office" /f "Livelid_" /s
  findstr /i "Livelid_" "%TMP%" > "%TMP%.2" 2>nul
  for /f "usebackq delims=" %%L in ("%TMP%.2") do (
    set "FOUND_ID=%%~nL"
    goto :found
  )
  del "%TMP%" "%TMP%.2" >nul 2>&1
)

:found
rem 3. Construir ruta final según resultado
if defined FOUND_ID (
  set "EXCEL_MRU_PATH=HKCU\Software\Microsoft\Office\16.0\Excel\Recent Templates\!FOUND_ID!\File MRU"
) else (
  rem --- Fallback: usuario sin ADAL/Livelid ---
  set "EXCEL_MRU_PATH=HKCU\Software\Microsoft\Office\16.0\Excel\Recent Templates\File MRU"
)

endlocal & set "EXCEL_MRU_PATH=%EXCEL_MRU_PATH%"
exit /b


:SimulateRegEntry
rem ------------------------------------------------------
rem Args: FILE_NAME FULL_PATH LOG_FILE
rem Agrega realmente los valores "Item N" y "Item Metadata N"
rem ------------------------------------------------------
setlocal enabledelayedexpansion
set "FILE_NAME=%~1"
set "FULL_PATH=%~2"
set "LOG_FILE=%~3"

set "LOCAL_LOGGING=true"
if /I "%IsDesignModeEnabled%"=="false" set "LOCAL_LOGGING=false"

rem --- Detectar MRU real ---
if not defined EXCEL_MRU_PATH call :DetectExcelMRUPath

rem --- Fallback si sigue vacío ---
if not defined EXCEL_MRU_PATH (
  set "EXCEL_MRU_PATH=HKCU\Software\Microsoft\Office\16.0\Excel\Recent Templates\File MRU"
)


rem ------------------------------------------------------
rem === DESPLAZAR INDICES EXISTENTES ANTES DE REGISTRAR ===
rem ------------------------------------------------------
echo "Entrando en desplazamiento de índices MRU de Excel..."

set "SHIFT_SCRIPT=%~dp0lib\shift_excel_mru_indices.bat"
if exist "%SHIFT_SCRIPT%" (
    call "%SHIFT_SCRIPT%"
    if /I "%IsDesignModeEnabled%"=="true" echo [DEBUG] shift_excel_mru_indices.bat ejecutado antes del registro.
) else (
    if /I "%IsDesignModeEnabled%"=="true" echo [WARN] No se encontró shift_excel_mru_indices.bat.
)
echo "Desplazamiento de índices completado."

rem --- Inicializar contador global si no existe ---
if not defined GLOBAL_ITEM_COUNT_EXCEL set /a GLOBAL_ITEM_COUNT_EXCEL=0
set /a LOCAL_COUNT=!GLOBAL_ITEM_COUNT_EXCEL!+1


rem ------------------------------------------------------
rem === CREAR VALOR PRINCIPAL (Item N)
rem ------------------------------------------------------
set "REG_VALUE=Item 1"
set "REG_DATA=[F00000000][T01ED6D7E58D00000][O00000000]*%FULL_PATH%"

if /I "%IsDesignModeEnabled%"=="true" echo [DEBUG] Escribiendo %REG_VALUE% en "%EXCEL_MRU_PATH%"
reg add "%EXCEL_MRU_PATH%" /v "!REG_VALUE!" /t REG_SZ /d "!REG_DATA!" /f >nul 2>&1

if errorlevel 1 (
  if /I "%LOCAL_LOGGING%"=="true" echo [ERROR] Falló al escribir %REG_VALUE% >> "%LOG_FILE%"
  if /I "%IsDesignModeEnabled%"=="true" echo [ERROR] Falló al escribir %REG_VALUE%
) else (
  if /I "%IsDesignModeEnabled%"=="true" echo [OK] %REG_VALUE% agregado correctamente
)


rem ------------------------------------------------------
rem === CREAR VALOR METADATA (Item Metadata N)
rem ------------------------------------------------------
for %%N in ("%FILE_NAME%") do set "BASENAME=%%~nN"

set "META_VALUE=Item Metadata 1"
set "META_DATA=<Metadata><AppSpecific><id>%FULL_PATH%</id><nm>%BASENAME%</nm><du>%FULL_PATH%</du></AppSpecific></Metadata>"

if /I "%IsDesignModeEnabled%"=="true" echo [DEBUG] Escribiendo %META_VALUE% en "%EXCEL_MRU_PATH%"
reg add "%EXCEL_MRU_PATH%" /v "!META_VALUE!" /t REG_SZ /d "!META_DATA!" /f >nul 2>&1

if errorlevel 1 (
  if /I "%LOCAL_LOGGING%"=="true" echo [ERROR] Falló al escribir %META_VALUE% >> "%LOG_FILE%"
  if /I "%IsDesignModeEnabled%"=="true" echo [ERROR] Falló al escribir %META_VALUE%
) else (
  if /I "%IsDesignModeEnabled%"=="true" echo [OK] %META_VALUE% agregado correctamente
)


if /I "%LOCAL_LOGGING%"=="true" (
  (
    echo [REG ENTRY]
    echo REG ADD "%EXCEL_MRU_PATH%" /v "!REG_VALUE!" /t REG_SZ /d "!REG_DATA!" /f
    echo REG ADD "%EXCEL_MRU_PATH%" /v "!META_VALUE!" /t REG_SZ /d "!META_DATA!" /f
    echo [INFO] Archivo: "!FILE_NAME!"
    echo.
  ) >> "%LOG_FILE%"
)

rem --- Exportar el nuevo valor del contador global ---
endlocal & set /a GLOBAL_ITEM_COUNT_EXCEL=%LOCAL_COUNT%
exit /b
