@echo off
rem ============================================================
rem detect_adal_key.bat  (versión amplia y segura)
rem Busca ADAL_ o Livelid_ en las rutas probables y luego en toda
rem HKCU\Software\Microsoft\Office usando reg query /f (rápido).
rem Salidas:
rem   - Si encuentra: imprime SOLO el IDENTIFICADOR (p. ej. ADAL_xxx) y exit 0
rem   - Si no encuentra: imprime NO_ENCONTRADO y exit 1
rem ============================================================

setlocal enabledelayedexpansion
set "FOUND_ID="
set "FOUND_PATH="

echo [DEBUG] Inicio de búsqueda rápida en Recent Templates...
for %%V in (16.0 15.0) do (
  for %%A in (PowerPoint Word Excel) do (
    set "BASE=HKCU\Software\Microsoft\Office\%%V\%%A\Recent Templates"
    echo [DEBUG] Revisando: !BASE!
    2>nul reg query "!BASE!" | findstr /i /r "ADAL_ Livelid_" >nul
    if not errorlevel 1 (
      rem extraer la subclave que contiene ADAL_ o Livelid_
      for /f "tokens=*" %%K in ('reg query "!BASE!" 2^>nul ^| findstr /i /r "ADAL_ Livelid_"') do (
        set "FOUND_PATH=%%K"
        set "FOUND_ID=%%~nK"
        goto :FOUND
      )
    )
  )
)

echo [DEBUG] No encontrado en Recent Templates. Ahora busco en toda la rama Office (más amplia)...
rem Buscar ADAL_ a nivel de rama (use /f para acelerar)
set "TMP=%TEMP%\adal_search.txt"
> "%TMP%" 2>&1 reg query "HKCU\Software\Microsoft\Office" /f "ADAL_" /s
findstr /i "ADAL_" "%TMP%" > "%TMP%.2" 2>nul
for /f "usebackq delims=" %%L in ("%TMP%.2") do (
  set "FOUND_PATH=%%L"
  set "FOUND_ID=%%~nL"
  goto :FOUND
)

rem Si no apareció ADAL_, buscar Livelid_
> "%TMP%" 2>&1 reg query "HKCU\Software\Microsoft\Office" /f "Livelid_" /s
findstr /i "Livelid_" "%TMP%" > "%TMP%.2" 2>nul
for /f "usebackq delims=" %%L in ("%TMP%.2") do (
  set "FOUND_PATH=%%L"
  set "FOUND_ID=%%~nL"
  goto :FOUND
)

rem limpiar temporales
del "%TMP%" "%TMP%.2" >nul 2>&1

:FOUND
if defined FOUND_ID (
  echo [DEBUG] Identificador encontrado: !FOUND_ID!
  echo [DEBUG] Ruta: !FOUND_PATH!
  echo !FOUND_ID!
  endlocal & exit /b 0
) else (
  echo [DEBUG] Ninguna clave ADAL_ ni Livelid_ encontrada en HKCU\Software\Microsoft\Office.
  echo NO_ENCONTRADO
  endlocal & exit /b 1
)
