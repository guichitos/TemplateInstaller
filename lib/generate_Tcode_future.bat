@echo off
setlocal EnableExtensions EnableDelayedExpansion
chcp 65001 >nul

rem ============================================================
rem === GENERATE T[xxxxxxxxxxxxxxxx] CODE FOR DATE +10 YEARS ===
rem ============================================================

rem 1. Obtener fecha y hora actual en UTC
for /f "tokens=2 delims==." %%a in ('wmic os get localdatetime /value') do set dt=%%a
set "YYYY=%dt:~0,4%"
set "MM=%dt:~4,2%"
set "DD=%dt:~6,2%"
set "hh=%dt:~8,2%"
set "nn=%dt:~10,2%"
set "ss=%dt:~12,2%"

rem 2. Calcular la fecha dentro de 10 años usando PowerShell (solo para precisión)
for /f %%F in ('powershell -NoLogo -Command "(Get-Date).AddYears(10).ToFileTimeUtc().ToString('X16')"') do set "HEX=%%F"

rem 3. Formar el código tipo [Txxxxxxxxxxxxxxxx]
set "TCODE=[T%HEX%]"

echo ==============================================
echo Fecha actual: %YYYY%-%MM%-%DD% %hh%:%nn%:%ss%
echo Fecha futura (+10 años): (calculada por PowerShell)
echo Código T: %TCODE%
echo ==============================================

pause
