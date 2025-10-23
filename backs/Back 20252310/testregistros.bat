@echo off
chcp 65001 >nul
setlocal enabledelayedexpansion

echo ============================================
echo === TEST: lectura tipo copy_templates.bat ===
echo ============================================
echo.

set "BASE_DIR=%~dp0"
set "SCRIPT_PATH=%BASE_DIR%detect_office_paths.bat"
if not exist "%SCRIPT_PATH%" (
    if exist "%BASE_DIR%lib\detect_office_paths.bat" set "SCRIPT_PATH=%BASE_DIR%lib\detect_office_paths.bat"
)

echo [INFO] Ejecutando detect_office_paths.bat desde:
echo   "%SCRIPT_PATH%"
echo.

rem Capturar rutas
for /f "tokens=1,* delims=:" %%A in (
    'call "%SCRIPT_PATH%" DetectOfficePaths ^| findstr /R /C:"WordPath:" /C:"PowerPointPath:" /C:"ExcelPath:"'
) do (
    set "NAME=%%A"
    set "RAW=%%B"
    rem eliminar espacio inicial
    for /f "tokens=* delims= " %%L in ("%%B") do (
        set "VALUE=%%L"
        if /I "%%A"=="WordPath" set "WORD_PATH=%%L"
        if /I "%%A"=="PowerPointPath" set "PPT_PATH=%%L"
        if /I "%%A"=="ExcelPath" set "EXCEL_PATH=%%L"
    )
)

rem Eliminar barra invertida final en PPT_PATH si existe
if defined PPT_PATH (
    if "!PPT_PATH:~-1!"=="\" set "PPT_PATH=!PPT_PATH:~0,-1!"
)

echo [RESULTADO] Word path: [%WORD_PATH%]
echo [RESULTADO] PowerPoint path: [%PPT_PATH%]
echo [RESULTADO] Excel path: [%EXCEL_PATH%]
echo.
pause
endlocal
exit /b
