@echo off
chcp 65001 >nul
setlocal enabledelayedexpansion

rem ==========================================================
rem === COPY_TEMPLATES.BAT - Copiado de plantillas con registro MRU
rem ==========================================================

set "DEBUG_VISIBLE=1"
if "%DEBUG_VISIBLE%"=="1" (
    title COPY_TEMPLATES DEBUG MODE
    echo [DEBUG] copy_templates.bat started
    echo Script running from: %~dp0
    echo Arguments: %*
    echo.
    pause
)

rem === Entry dispatcher ===
if "%~1"=="" exit /b
set "FUNC=%~1"
if "%FUNC:~0,1%"==":" set "FUNC=%FUNC:~1%"
goto %FUNC%


:CopyAll
rem Args: LOG_FILE BASE_DIR REG_LIB
shift
set "LOG_FILE=%~1"
set "BASE_DIR=%~2"
set "REG_LIB=%~3"

setlocal enabledelayedexpansion
set /a TOTAL_FILES=0
set /a TOTAL_ERRORS=0

rem === Detect library in same folder ===
set "DETECT_LIB=%~dp0lib\detect_office_paths.bat"
echo [DEBUG] detect_lib path resolved to: "%DETECT_LIB%" >> "%LOG_FILE%"
echo Detect library resolved to: "%DETECT_LIB%"
if not exist "%DETECT_LIB%" (
    echo [ERROR] detect_office_paths.bat NOT FOUND: "%DETECT_LIB%" >> "%LOG_FILE%"
    echo detect_office_paths.bat NOT FOUND
    pause
    endlocal
    exit /b 1
)

rem --- Ejecutar detector (sin debug interno) ---
for /f "tokens=1,* delims=:" %%A in (
    'call "%DETECT_LIB%" DetectOfficePaths ^| findstr /R /C:"WordPath:" /C:"PowerPointPath:" /C:"ExcelPath:"'
) do (
    for /f "tokens=* delims= " %%L in ("%%B") do (
        if /I "%%A"=="WordPath" set "WORD_PATH=%%L"
        if /I "%%A"=="PowerPointPath" set "PPT_PATH=%%L"
        if /I "%%A"=="ExcelPath" set "EXCEL_PATH=%%L"
    )
)

rem === Quitar barra final si existe ===
if defined WORD_PATH if "!WORD_PATH:~-1!"=="\" set "WORD_PATH=!WORD_PATH:~0,-1!"
if defined PPT_PATH  if "!PPT_PATH:~-1!"=="\"  set "PPT_PATH=!PPT_PATH:~0,-1!"
if defined EXCEL_PATH if "!EXCEL_PATH:~-1!"=="\" set "EXCEL_PATH=!EXCEL_PATH:~0,-1!"

rem === Mostrar resultado final detectado ===
echo.
echo [DEBUG] === Paths detectados (finales) ===
echo   WORD_PATH = !WORD_PATH!
echo   PPT_PATH  = !PPT_PATH!
echo   EXCEL_PATH= !EXCEL_PATH!
echo [DEBUG] ==================================
echo.
pause

rem ==========================================================
rem === ETAPA EXTRA: DETECTAR MRU DE POWERPOINT UNA VEZ ===
rem ==========================================================
set "REGISTRY_LIB=%~dp0lib\registry_tools.bat"
if exist "%REGISTRY_LIB%" (
    call "%REGISTRY_LIB%" DetectPowerPointMRUPath
    echo [DEBUG] MRU path detectado: %PPT_MRU_PATH% >> "%LOG_FILE%"
) else (
    echo [WARNING] registry_tools.bat no encontrado. No se registrarán plantillas en MRU.
    echo [WARNING] registry_tools.bat no encontrado >> "%LOG_FILE%"
)
echo.
pause

rem ==========================================================
rem === ETAPA 2: LISTADO Y VALIDACIÓN PREVIA DE ARCHIVOS ===
rem ==========================================================

echo [DEBUG] --- Iniciando exploración de archivos en BASE_DIR --- >> "%LOG_FILE%"
echo [INFO] Buscando plantillas en "%BASE_DIR%"...
echo -----------------------------------------------
dir /b "%BASE_DIR%\*.dot*" "%BASE_DIR%\*.pot*" "%BASE_DIR%\*.xlt*" 2>nul
echo -----------------------------------------------
echo.

if errorlevel 1 (
    echo [WARNING] No se encontraron archivos de plantilla en "%BASE_DIR%".
    echo [WARNING] Ningún archivo .dotx / .potx / .xltx detectado. >> "%LOG_FILE%"
    pause
)

rem ==========================================================
rem === ETAPA 3: VERIFICACIÓN DE RUTAS DE DESTINO ============
rem ==========================================================

echo [DEBUG] Verificando existencia de rutas destino...
for %%P in ("!WORD_PATH!" "!PPT_PATH!" "!EXCEL_PATH!") do (
    if exist "%%~P" (
        echo [OK] Carpeta válida: %%~P
    ) else (
        echo [ERROR] Carpeta inexistente: %%~P
        echo [ERROR] Carpeta inexistente: %%~P >> "%LOG_FILE%"
    )
)
echo.
pause

rem ==========================================================
rem === ETAPA 4: COPIADO REAL CON DEBUG Y VALIDACIÓN =========
rem ==========================================================

echo [DEBUG] Entrando a etapa de copiado...
echo [DEBUG] BASE_DIR = "%BASE_DIR%"
echo -----------------------------------------------
pause

echo [DEBUG] Iniciando iteración con ruta completa...
for %%F in ("%BASE_DIR%\*.dotx" "%BASE_DIR%\*.dotm" "%BASE_DIR%\*.potx" "%BASE_DIR%\*.potm" "%BASE_DIR%\*.xltx" "%BASE_DIR%\*.xltm") do (
    if exist "%%~fF" (
        set "FN=%%~nxF"
        set "EXT=%%~xF"

        rem === Verificar si debe omitirse ===
        set "SKIP=0"
        if /I "!FN!"=="GenericTemplate.dotm" set "SKIP=1"
        if /I "!FN!"=="GenericTemplate.potx" set "SKIP=1"
        if /I "!FN!"=="GenericTemplate.xltx" set "SKIP=1"

        rem === Determinar destino ===
        set "DEST="
        if /I "!EXT!"==".dotx" set "DEST=!WORD_PATH!"
        if /I "!EXT!"==".dotm" set "DEST=!WORD_PATH!"
        if /I "!EXT!"==".potx" set "DEST=!PPT_PATH!"
        if /I "!EXT!"==".potm" set "DEST=!PPT_PATH!"
        if /I "!EXT!"==".xltx" set "DEST=!EXCEL_PATH!"
        if /I "!EXT!"==".xltm" set "DEST=!EXCEL_PATH!"

        echo.
        echo [DEBUG] Procesando archivo: !FN!
        echo [DEBUG] Extensión detectada: !EXT!

        if "!SKIP!"=="1" (
            echo [INFO] Archivo genérico omitido: !FN!
            echo [INFO] Omitido genérico: !FN! >> "%LOG_FILE%"
        ) else (
            if defined DEST (
                echo [DEBUG] Destino asignado: !DEST!
                echo [ACTION] Copiando !FN! → !DEST!
                echo [DEBUG] Copiando !FN! a !DEST! >> "%LOG_FILE%"
                mkdir "!DEST!" 2>nul
                copy /Y "%%~fF" "!DEST!\" >> "%LOG_FILE%" 2>&1

                if exist "!DEST!\!FN!" (
                    echo [OK] Copiado exitosamente: !FN!
                    echo [RESULT] Éxito → !FN! >> "%LOG_FILE%"
                    set /a TOTAL_FILES+=1

                    rem === Registrar en el MRU (solo PowerPoint) ===
                    if /I "!EXT!"==".potx" (
                        call "%REGISTRY_LIB%" SimulateRegEntry "!FN!" "!DEST!\!FN!" "%LOG_FILE%"
                    )
                    if /I "!EXT!"==".potm" (
                        call "%REGISTRY_LIB%" SimulateRegEntry "!FN!" "!DEST!\!FN!" "%LOG_FILE%"
                    )

                ) else (
                    echo [ERROR] Falló la copia de !FN!
                    echo [RESULT] Error → !FN! >> "%LOG_FILE%"
                    set /a TOTAL_ERRORS+=1
                )
            ) else (
                echo [WARNING] No se asignó destino para !FN!
                echo [WARNING] Sin destino → !FN! >> "%LOG_FILE%"
            )
        )
        echo -----------------------------------------------
        pause
    )
)
echo [DEBUG] Fin del bucle de copiado
echo [DEBUG] TOTAL_FILES=!TOTAL_FILES! TOTAL_ERRORS=!TOTAL_ERRORS!
pause

rem ==========================================================
rem === ETAPA 5: RESUMEN FINAL ===============================
rem ==========================================================
echo.
echo [FINAL] Copiado completado.
echo   Archivos copiados: !TOTAL_FILES!
echo   Archivos con error: !TOTAL_ERRORS!
echo ----------------------------------------------------------
echo [DEBUG] Total copiados: !TOTAL_FILES!, con errores: !TOTAL_ERRORS! >> "%LOG_FILE%"
echo.
pause

endlocal
exit /b
