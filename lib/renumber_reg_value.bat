@echo off
chcp 65001 >nul
setlocal enabledelayedexpansion

if "%~1"=="" exit /b 0
set "TARGET_KEY=%~1"
set "SHIFT_INPUT=%~2"
if not defined SHIFT_INPUT set "SHIFT_INPUT=1"
set /a SHIFT_INT=%SHIFT_INPUT% 2>nul
if errorlevel 1 set /a SHIFT_INT=1
if !SHIFT_INT! lss 1 set /a SHIFT_INT=1
set "LOG_FILE=%~3"
set "IsDesignModeEnabled=%~4"
if not defined IsDesignModeEnabled set "IsDesignModeEnabled=false"

set "ENABLE_LOG=false"
if defined LOG_FILE (
    if not "!LOG_FILE!"=="" set "ENABLE_LOG=true"
)

if /I "!IsDesignModeEnabled!"=="true" (
    echo [DEBUG] RenumberRegValue called for "!TARGET_KEY!" (shift !SHIFT_INT!)
)

set "PS_PATH=!TARGET_KEY!"
if not "!PS_PATH:~4,1!"==":" (
    if /I "!TARGET_KEY:~0,5!"=="HKCU\" set "PS_PATH=HKCU:\!TARGET_KEY:~5!"
    if /I "!TARGET_KEY:~0,5!"=="HKLM\" set "PS_PATH=HKLM:\!TARGET_KEY:~5!"
    if /I "!TARGET_KEY:~0,5!"=="HKCR\" set "PS_PATH=HKCR:\!TARGET_KEY:~5!"
    if /I "!TARGET_KEY:~0,4!"=="HKU\" set "PS_PATH=HKU:\!TARGET_KEY:~4!"
    if /I "!TARGET_KEY:~0,5!"=="HKCC\" set "PS_PATH=HKCC:\!TARGET_KEY:~5!"
    if /I "!TARGET_KEY:~0,18!"=="HKEY_CURRENT_USER" (
        if "!TARGET_KEY:~18,1!"=="\" set "PS_PATH=HKCU:\!TARGET_KEY:~19!"
    )
    if /I "!TARGET_KEY:~0,19!"=="HKEY_LOCAL_MACHINE" (
        if "!TARGET_KEY:~19,1!"=="\" set "PS_PATH=HKLM:\!TARGET_KEY:~20!"
    )
    if /I "!TARGET_KEY:~0,17!"=="HKEY_CLASSES_ROOT" (
        if "!TARGET_KEY:~17,1!"=="\" set "PS_PATH=HKCR:\!TARGET_KEY:~18!"
    )
    if /I "!TARGET_KEY:~0,9!"=="HKEY_USERS" (
        if "!TARGET_KEY:~9,1!"=="\" set "PS_PATH=HKU:\!TARGET_KEY:~10!"
    )
    if /I "!TARGET_KEY:~0,20!"=="HKEY_CURRENT_CONFIG" (
        if "!TARGET_KEY:~20,1!"=="\" set "PS_PATH=HKCC:\!TARGET_KEY:~21!"
    )
)

set "TEMP_PS=%TEMP%\renumber_%RANDOM%%RANDOM%.ps1"
> "!TEMP_PS!" (
    echo param^(
    echo     ^[Parameter^(Mandatory=$true^)^][string]$RegistryPath,
    echo     [int]$ShiftBy = 1,
    echo     [string]$LogFile = "",
    echo     [string]$DesignMode = "false",
    echo     [string]$EnableLog = "false"
    echo )
    echo $designModeEnabled = $DesignMode -ieq "true"
    echo $logEnabled = ($EnableLog -ieq "true") -and -not [string]::IsNullOrWhiteSpace($LogFile)
    echo try {
    echo     if(-not (Test-Path -LiteralPath $RegistryPath)){
    echo         if($designModeEnabled){ Write-Output (^"[DEBUG] Registry path not found: {0}^" -f $RegistryPath) }
    echo         exit 0
    echo     }
    echo     $props = Get-ItemProperty -LiteralPath $RegistryPath -ErrorAction Stop
    echo } catch {
    echo     if($designModeEnabled){ Write-Output (^"[WARNING] Unable to access registry path: {0}^" -f $RegistryPath) }
    echo     exit 1
    echo }
    echo $entries = @()
    echo foreach($prop in $props.PSObject.Properties){
    echo     if($prop.Name -match '^Item( Metadata)? (\d+)$'){
    echo         $idx = [int][regex]::Match($prop.Name,'(\d+)$').Value
    echo         $isMetadata = $prop.Name -like 'Item Metadata *'
    echo         $entries += [pscustomobject]@{
    echo             OldName = $prop.Name
    echo             NewName = if($isMetadata){"Item Metadata " + ($idx + $ShiftBy)} else {"Item " + ($idx + $ShiftBy)}
    echo             Value = $prop.Value
    echo             Index = $idx
    echo             IsMetadata = $isMetadata
    echo         }
    echo     }
    echo }
    echo if($entries.Count -eq 0){
    echo     if($designModeEnabled){ Write-Output (^"[DEBUG] No MRU entries found to renumber at {0}^" -f $RegistryPath) }
    echo     exit 0
    echo }
    echo $entries = $entries ^| Sort-Object -Property ^@{Expression = { $_.Index }; Descending = $true }, ^@{Expression = { $_.IsMetadata }; Descending = $false }
    echo foreach($entry in $entries){
    echo     New-ItemProperty -LiteralPath $RegistryPath -Name $entry.NewName -Value $entry.Value -PropertyType String -Force ^| Out-Null
    echo }
    echo foreach($entry in $entries){
    echo     Remove-ItemProperty -LiteralPath $RegistryPath -Name $entry.OldName -ErrorAction SilentlyContinue
    echo }
    echo if($designModeEnabled){
    echo     foreach($entry in $entries){
    echo         Write-Output (^"[DEBUG] Renamed '{0}' → '{1}'^" -f $entry.OldName,$entry.NewName)
    echo     }
    echo }
    echo if($logEnabled){
    echo     $logLine = "[REG REN] " + (Get-Date -Format "yyyy-MM-dd HH:mm:ss") + " Shifted " + $entries.Count + " entries at " + $RegistryPath + " by " + $ShiftBy
    echo     Add-Content -LiteralPath $LogFile -Value $logLine
    echo }
    echo exit 0
)

powershell -NoProfile -ExecutionPolicy Bypass -File "!TEMP_PS!" -RegistryPath "!PS_PATH!" -ShiftBy !SHIFT_INT! -LogFile "!LOG_FILE!" -DesignMode "!IsDesignModeEnabled!" -EnableLog "!ENABLE_LOG!"
set "PS_EXIT=%ERRORLEVEL%"
del "!TEMP_PS!" >nul 2>&1
if /I "!IsDesignModeEnabled!"=="true" (
    if !PS_EXIT! equ 0 (
        echo [DEBUG] RenumberRegValue completed with code 0
    ) else (
        echo [WARNING] RenumberRegValue finished with code !PS_EXIT!
    )
)
endlocal & exit /b %PS_EXIT%
