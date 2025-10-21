@echo off
if "%~1"=="" exit /b
set "FUNC=%~1"
shift
goto %FUNC%


:DetectPowerPointMRUPath
rem ------------------------------------------------------
rem Detecta la ruta MRU real de PowerPoint (ADAL_ o Livelid_)
rem ------------------------------------------------------
setlocal enabledelayedexpansion
set "PPT_MRU_PATH="
for /f "delims=" %%R in (
  'reg query "HKCU\Software\Microsoft\Office\16.0\Common\Internet\WebServiceCache" 2^>nul ^| findstr /R /C:"ADAL_" /C:"Livelid_"'
) do (
  set "ADAL_KEY=%%~nR"
  set "PPT_MRU_PATH=HKCU\Software\Microsoft\Office\16.0\Common\Internet\WebServiceCache\!ADAL_KEY!\PowerPoint\Recent Templates\File MRU"
  goto :found
)
:found
endlocal & set "PPT_MRU_PATH=%PPT_MRU_PATH%"
exit /b


:SimulateRegEntry
rem Args: FILE_NAME FULL_PATH LOG_FILE
setlocal
set "FILE_NAME=%~1"
set "FULL_PATH=%~2"
set "LOG_FILE=%~3"

rem --- Si no se detectó MRU aún, usar la función detect ---
if not defined PPT_MRU_PATH call :DetectPowerPointMRUPath

rem --- Si sigue vacío, fallback estándar ---
if not defined PPT_MRU_PATH set "PPT_MRU_PATH=HKCU\Software\Microsoft\Office\16.0\PowerPoint\Recent Templates\File MRU"

set "REG_VALUE=Item_%FILE_NAME%"
set "REG_DATA=[F00000000][T01DC3E24ECBDAAB0][O00000000]*%FULL_PATH%"

echo Simulating REG ADD "%PPT_MRU_PATH%" /v "%REG_VALUE%" /t REG_SZ /d "%REG_DATA%" /f
(
  echo [SIMULATED REG ENTRY]
  echo REG ADD "%PPT_MRU_PATH%" /v "%REG_VALUE%" /t REG_SZ /d "%REG_DATA%" /f
  echo.
) >> "%LOG_FILE%"
endlocal
exit /b
