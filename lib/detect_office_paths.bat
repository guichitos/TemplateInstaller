rem detect_office_paths.bat
@echo off
if "%~1"=="" exit /b
set "FUNC=%~1"
shift
goto %FUNC%

:DetectOfficePaths
rem ======================================================
rem === DETECT OFFICE PERSONAL TEMPLATE PATHS (16.0) ===
rem Devuelve las rutas de "PersonalTemplates" en variables
rem globales: WORD_PATH, PPT_PATH y EXCEL_PATH
rem ======================================================

setlocal enabledelayedexpansion

rem --- WORD ---
for /f "tokens=1,2,*" %%A in (
  'reg query "HKCU\Software\Microsoft\Office\16.0\Word\Options" /v "PersonalTemplates" 2^>nul ^| find /I "PersonalTemplates"'
) do set "WORD_PATH=%%C"

rem --- POWERPOINT ---
for /f "tokens=1,2,*" %%A in (
  'reg query "HKCU\Software\Microsoft\Office\16.0\PowerPoint\Options" /v "PersonalTemplates" 2^>nul ^| find /I "PersonalTemplates"'
) do set "PPT_PATH=%%C"

rem --- EXCEL ---
for /f "tokens=1,2,*" %%A in (
  'reg query "HKCU\Software\Microsoft\Office\16.0\Excel\Options" /v "PersonalTemplates" 2^>nul ^| find /I "PersonalTemplates"'
) do set "EXCEL_PATH=%%C"

rem --- Limpiar rutas ---
call :CleanPath WORD_PATH
call :CleanPath PPT_PATH
call :CleanPath EXCEL_PATH

rem --- Salida final para copy_templates ---
echo WordPath: %WORD_PATH%
echo PowerPointPath: %PPT_PATH%
echo ExcelPath: %EXCEL_PATH%

endlocal & (
  set "WORD_PATH=%WORD_PATH%"
  set "PPT_PATH=%PPT_PATH%"
  set "EXCEL_PATH=%EXCEL_PATH%"
)
exit /b


:CleanPath
rem === Limpia comillas, espacios y agrega C: si falta ===
setlocal enabledelayedexpansion
set "VAR=%~1"
for /f "tokens=2 delims==" %%A in ('set !VAR! 2^>nul') do set "VAL=%%A"

if defined VAL (
  rem Quitar comillas
  set "VAL=!VAL:"=!"
  rem Quitar espacios y tabulaciones
  for /f "tokens=* delims= " %%Z in ('echo(!VAL!') do set "VAL=%%Z"
  set "VAL=!VAL:	=!"
  rem Agregar C: si empieza con \
  if "!VAL:~0,1!"=="\" set "VAL=C:!VAL!"
)
endlocal & if defined VAL set "%~1=%VAL%"
exit /b
