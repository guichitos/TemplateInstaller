@echo off
setlocal enabledelayedexpansion

set "FILE=%TEMP%\1-2. Main.bat"

if not exist "%FILE%" (
    echo No se encontro el archivo en la carpeta TEMP.
    echo Ruta esperada:
    echo    %FILE%
    exit /b 1
)

echo Ejecutando:
echo    "%FILE%"
call "%FILE%"

endlocal
exit /b 0
