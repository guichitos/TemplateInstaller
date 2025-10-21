@echo off
setlocal EnableExtensions

rem =====================================================
rem === CONFIGURACIÓN GENERAL
rem =====================================================
set "REG_SUBKEY=Software\Microsoft\Office\16.0\PowerPoint\Recent Templates\ADAL_DA750CDDD552F36ACCDCA5039F8D25BB03D146010B3B064F371FB4F5B5CF54C9\File MRU"
set "VALUE_NAME=Item 1"
set "VALUE_DATA=[F00000000][T01DC3E24ECBDAAB0][O00000000]*C:\Users\%USERNAME%\OneDrive\Documentos\Plantillas personalizadas de Office\Template 2.potx"

for /f %%I in ('powershell -NoProfile -Command "Get-Date -Format yyyyMMdd"') do set "TODAY=%%I"
set "BACKUP_FILE=%TEMP%\registry_backup_%USERNAME%_%TODAY%.reg"

rem =====================================================
rem === DETECTAR SI SE EJECUTA COMO ADMIN
rem =====================================================
openfiles >nul 2>&1
if %errorlevel% EQU 0 (
  set "IS_ADMIN=1"
) else (
  set "IS_ADMIN=0"
)

rem =====================================================
rem === DETECTAR SID DEL USUARIO ACTUAL (versión robusta)
rem =====================================================
for /f "usebackq tokens=*" %%S in (`powershell -NoProfile -Command "[System.Security.Principal.WindowsIdentity]::GetCurrent().User.Value"`) do (
    set "SID=%%S"
)

if "%SID%"=="" (
  echo ERROR: No se pudo obtener el SID del usuario actual.
  echo PowerShell devolvio vacío; revisa la instalación o políticas de PowerShell.
  pause
  exit /b 1
)

echo SID detectado: %SID%



rem =====================================================
rem === MOSTRAR CONTEXTO Y DATOS
rem =====================================================
echo Usuario actual: %USERNAME%
if "%IS_ADMIN%"=="1" (
  echo Contexto: Ejecutando como administrador
  echo SID detectado: %SID%
) else (
  echo Contexto: Usuario normal
)
echo.
echo Clave de destino: %FULL_KEY%
echo Valor: %VALUE_NAME%
echo Datos: %VALUE_DATA%
echo.

rem =====================================================
rem === CREAR BACKUP (si es posible)
rem =====================================================
for %%R in ("HKCU" "HKU") do (
  reg query "%%~R" >nul 2>&1
)
reg export "%FULL_KEY%\.." "%BACKUP_FILE%" /y >nul 2>&1
if errorlevel 1 (
  echo No se pudo exportar el subarbol (tal vez no existe). Continuando...
) else (
  echo Backup creado en: "%BACKUP_FILE%"
)
echo.

rem =====================================================
rem === CREAR CLAVE Y ESCRIBIR VALOR
rem =====================================================
echo Creando clave si no existe...
reg add "%FULL_KEY%" /f >nul
if errorlevel 1 (
  echo ERROR: No se pudo crear la clave "%FULL_KEY%".
  pause
  exit /b 1
)

echo Escribiendo valor...
reg add "%FULL_KEY%" /v "%VALUE_NAME%" /t REG_SZ /d "%VALUE_DATA%" /f >nul
if errorlevel 1 (
  echo ERROR: No se pudo escribir el valor "%VALUE_NAME%".
  pause
  exit /b 1
)

rem =====================================================
rem === VERIFICACIÓN
rem =====================================================
echo.
echo Verificacion:
reg query "%FULL_KEY%" /v "%VALUE_NAME%"
echo.

echo Operacion completada correctamente.
pause
endlocal
exit /b 0
