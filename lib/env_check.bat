rem ======================================================
rem === ENVIRONMENT CHECK LIBRARY ===
rem ------------------------------------------------------
rem Ensures the script runs with sufficient privileges.
rem Detects if elevation (administrator mode) is required
rem and relaunches the main installer with admin rights
rem using PowerShell if necessary.
rem ======================================================

@echo off
if "%~1"=="" exit /b
set "FUNC=%~1"
shift
goto %FUNC%

:CheckEnvironment
rem Args: LOG_FILE [CALLER_PATH]
set "LOG_FILE=%~1"
set "CALLER_PATH=%~2"

rem Si no nos pasaron el path del script llamador, usar el llamador actual (esto será env_check.bat)
if "%CALLER_PATH%"=="" set "CALLER_PATH=%~f0"

echo [%DATE% %TIME%] Checking environment... >> "%LOG_FILE%"

openfiles >nul 2>&1
if %errorlevel% NEQ 0 (
  echo [%DATE% %TIME%] Elevation required. Attempting to relaunch as admin... >> "%LOG_FILE%"
  rem Relanzar el script CALLER_PATH con elevación (cmd.exe /c "ruta")
  powershell -NoProfile -ExecutionPolicy Bypass -Command ^
    "Start-Process 'cmd.exe' -ArgumentList '/c','\"%CALLER_PATH%\"' -Verb RunAs"
  rem Salir: el proceso actual debe terminar; el caller elevado será quien continúe.
  exit /b
)

echo [%DATE% %TIME%] Environment check passed (already running as admin). >> "%LOG_FILE%"
exit /b
