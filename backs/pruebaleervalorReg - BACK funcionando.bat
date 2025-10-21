@echo off
setlocal EnableExtensions

rem =========================
rem CONFIGURACIÓN
rem =========================
set "SID=S-1-5-21-3876285334-669651212-1223383818-1001"
set "REG_SUBKEY=Software\Microsoft\Office\16.0\PowerPoint\Recent Templates\ADAL_DA750CDDD552F36ACCDCA5039F8D25BB03D146010B3B064F371FB4F5B5CF54C9\File MRU"
set "VALUE_NAME=Item 1"
set "VALUE_DATA=[F00000000][T01DC3E24ECBDAAB0][O00000000]*C:\Users\PC\OneDrive\Documentos\Plantillas personalizadas de Office\Template 2.potx"
set "FULL_KEY=HKU\%SID%\%REG_SUBKEY%"

rem Fecha segura (YYYYMMDD) para el backup
for /f %%I in ('powershell -NoProfile -Command "Get-Date -Format yyyyMMdd"') do set "TODAY=%%I"
set "BACKUP_FILE=%TEMP%\registry_backup_%SID%_%TODAY%.reg"

rem =========================
rem ELEVACIÓN
rem =========================
openfiles >nul 2>&1 || (powershell -NoProfile -Command "Start-Process -FilePath '%~f0' -Verb RunAs" & exit /b)

rem =========================
rem INFO INICIAL
rem =========================
echo Leyendo valor actual (si existe) en HKCU:
reg query "HKCU\Software\Microsoft\Office\16.0\PowerPoint\Recent Templates" /v "Friendly1"
echo.

echo Se va a crear/actualizar:
echo   Clave: %FULL_KEY%
echo   Valor: %VALUE_NAME%
echo   Datos: %VALUE_DATA%
echo.

rem =========================
rem BACKUP (sin paréntesis, en %TEMP%)
rem =========================
echo Exportando copia del subarbol a "%BACKUP_FILE%" ...
reg export "HKU\%SID%\Software\Microsoft\Office\16.0\PowerPoint\Recent Templates" "%BACKUP_FILE%" /y >nul 2>&1
if errorlevel 1 echo No se pudo exportar el subarbol (tal vez no existe). Continuando...
if not errorlevel 1 echo Backup creado: "%BACKUP_FILE%"
echo.

rem =========================
rem CREAR CLAVE Y ESCRIBIR VALOR
rem =========================
echo Creando clave si no existe...
reg add "%FULL_KEY%" /f >nul
if errorlevel 1 echo ERROR: No se pudo crear la clave "%FULL_KEY%". & pause & exit /b 1

echo Escribiendo valor...
reg add "%FULL_KEY%" /v "%VALUE_NAME%" /t REG_SZ /d "%VALUE_DATA%" /f >nul
if errorlevel 1 echo ERROR: No se pudo escribir el valor "%VALUE_NAME%". & pause & exit /b 1

rem =========================
rem VERIFICACIÓN
rem =========================
echo.
echo Verificacion:
reg query "%FULL_KEY%" /v "%VALUE_NAME%"
echo.

echo Operacion completada.
pause
endlocal
exit /b 0
