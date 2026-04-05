# Полный цикл: при необходимости запуск CorelDRAW, затем deploy_direct.ps1, лог в deploy_last_log.txt
$ErrorActionPreference = 'Stop'
$root = $PSScriptRoot
if ([string]::IsNullOrEmpty($root)) {
    $root = Split-Path -Parent $MyInvocation.MyCommand.Path
}
Set-Location -LiteralPath $root

$logPath = Join-Path $root 'deploy_last_log.txt'
function Write-Log([string]$msg) {
    $line = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') $msg"
    Add-Content -LiteralPath $logPath -Value $line -Encoding UTF8
    Write-Host $msg
}

Set-Content -LiteralPath $logPath -Value "=== Shape Builder deploy_all.ps1 ===" -Encoding UTF8
Write-Log "Folder: $root"

$corelRunning = $false
try {
    $a = [Runtime.InteropServices.Marshal]::GetActiveObject('CorelDRAW.Application')
    if ($null -ne $a) {
        $corelRunning = $true
        Write-Log 'CorelDRAW already running (ROT).'
    }
} catch {
    Write-Log 'CorelDRAW not in ROT yet.'
}

if (-not $corelRunning) {
    $exeCandidates = @(
        "${env:ProgramFiles}\Corel\CorelDRAW Graphics Suite 2026\Programs64\CorelDRW.exe",
        "${env:ProgramFiles}\Corel\CorelDRAW Graphics Suite 2025\Programs64\CorelDRW.exe",
        "${env:ProgramFiles}\Corel\CorelDRAW Graphics Suite 2024\Programs64\CorelDRW.exe",
        "${env:ProgramFiles}\Corel\CorelDRAW Graphics Suite 2023\Programs64\CorelDRW.exe",
        "${env:ProgramFiles}\Corel\CorelDRAW Graphics Suite 2022\Programs64\CorelDRW.exe",
        "${env:ProgramFiles}\Corel\CorelDRAW Graphics Suite 2021\Programs64\CorelDRW.exe",
        "${env:ProgramFiles(x86)}\Corel\CorelDRAW Graphics Suite 2021\Programs\CorelDRW.exe"
    )
    $started = $false
    foreach ($exe in $exeCandidates) {
        if (Test-Path -LiteralPath $exe) {
            Write-Log "Starting CorelDRAW: $exe"
            Start-Process -FilePath $exe -ErrorAction SilentlyContinue
            Start-Sleep -Seconds 22
            $started = $true
            break
        }
    }
    if (-not $started) {
        Write-Log 'CorelDRAW.exe not found in standard paths. Open CorelDRAW manually, create/open a document, press Alt+F11 once, then run deploy_all.ps1 again.'
    }
}

$deployScript = Join-Path $root 'deploy_direct.ps1'
if (-not (Test-Path -LiteralPath $deployScript)) {
    Write-Log "ERROR: deploy_direct.ps1 missing in $root"
    exit 1
}

Write-Log 'Running deploy_direct.ps1 ...'
$code = 1
for ($attempt = 1; $attempt -le 3; $attempt++) {
    Write-Log "Attempt $attempt / 3"
    & $deployScript 2>&1 | Tee-Object -FilePath $logPath -Append
    if ($null -ne $LASTEXITCODE) { $code = [int]$LASTEXITCODE } else { $code = 0 }
    if ($code -eq 0) { break }
    if ($attempt -lt 3) {
        Write-Log 'Waiting 12 s before retry (VBE may still be loading)...'
        Start-Sleep -Seconds 12
    }
}
Write-Log ("deploy_direct final exit code: " + $code)
exit $code
