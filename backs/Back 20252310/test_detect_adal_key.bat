@echo off
title Test detect_adal_key
color 0A
echo ============================================================
echo  TEST DE DETECCION DE CLAVE ADAL / LIVELID
echo ============================================================
echo.

setlocal
set "LIB=detect_adal_key.bat"

if not exist "%LIB%" (
  echo ERROR: No se encontro "%LIB%" en el directorio actual.
  echo.
  goto :END
)

rem Ejecutar y capturar salida completa en archivo temporal
set "TMPFILE=%TEMP%\adal_result_%RANDOM%.txt"
call "%LIB%" > "%TMPFILE%" 2>&1

rem Mostrar TODO el contenido del log temporal (trazas DEBUG + resultado)
echo --- Salida del script detect_adal_key.bat ---
type "%TMPFILE%"
echo --- Fin de salida ---
echo.

rem Leer la primera linea util (ignorando posibles lineas debug si quieres)
set "ADAL_ID="
for /f "usebackq delims=" %%I in ("%TMPFILE%") do (
  rem La primera linea imprimida por detect_adal_key.bat que no sea vacía la guardamos
  if not defined ADAL_ID set "ADAL_ID=%%I"
)

del "%TMPFILE%" >nul 2>&1

if not defined ADAL_ID (
  echo ERROR: No se obtuvo ningun valor de salida.
  echo (Posible error de ejecucion del archivo detect_adal_key.bat)
  echo.
  goto :END
)

if /i "%ADAL_ID%"=="NO_ENCONTRADO" (
  echo No se encontro ningun identificador ADAL_ ni Livelid_ en el registro.
  echo Asegurate de haber iniciado alguna aplicacion de Office con el usuario actual.
  echo.
  goto :END
)

echo Identificador encontrado: %ADAL_ID%
echo.

:END
echo ============================================================
echo Fin del test.
echo ============================================================
echo.
pause
exit /b
