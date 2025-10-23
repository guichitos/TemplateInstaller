@echo off
setlocal
chcp 65001 >nul

echo ======================================================
echo [INFO] Prueba del script SHIFT_POWERPOINT_MRU_INDICES
echo ======================================================

rem === Detectar ruta actual y script objetivo ===
set "CURRENT_DIR=%~dp0"
set "TARGET_SCRIPT=%CURRENT_DIR%shift_powerpoint_mru_indices.bat"

if not exist "%TARGET_SCRIPT%" (
    echo [ERROR] No se encontró el archivo "%TARGET_SCRIPT%"
    echo Asegúrate de que este test esté en la misma carpeta que el script principal.
    pause
    exit /b 1
)

echo [INFO] Ejecutando script: "%TARGET_SCRIPT%"
echo ------------------------------------------------------

call "%TARGET_SCRIPT%"

echo ------------------------------------------------------
echo [INFO] Script finalizado. Código de salida: %errorlevel%

if %errorlevel% neq 0 (
    echo [WARN] El script devolvió un código distinto de cero.
) else (
    echo [OK] Ejecución completada sin errores.
)

echo ======================================================
pause
endlocal
exit /b 0
