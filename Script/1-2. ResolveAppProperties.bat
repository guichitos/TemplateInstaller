@echo off
rem ============================================================
rem ===            1-2. ResolveAppProperties.bat             ===
rem ===           Biblioteca de propiedades de la app        ===
rem ===  Uso: call "1-2. ResolveAppProperties.bat" APP_NAME  ===
rem ===       Devuelve: PROP_REG_NAME                       ===
rem ============================================================

rem APP_NAME esperado:
rem   WORD
rem   POWERPOINT
rem   EXCEL

set "APP_UP=%~1"

if /I "%APP_UP%"=="WORD" (
    set "PROP_REG_NAME=Word"
) else if /I "%APP_UP%"=="POWERPOINT" (
    set "PROP_REG_NAME=PowerPoint"
) else if /I "%APP_UP%"=="EXCEL" (
    set "PROP_REG_NAME=Excel"
) else (
    rem Si no existe, devolver vacío
    set "PROP_REG_NAME="
)

exit /b 0
