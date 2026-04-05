# Wrapper: actual deploy is deploy_direct.ps1 (UserForm2, full COM fallbacks).
$ErrorActionPreference = 'Stop'
$here = $PSScriptRoot
if ([string]::IsNullOrEmpty($here)) {
    $here = Split-Path -Parent $MyInvocation.MyCommand.Path
}
$target = Join-Path $here 'deploy_direct.ps1'
if (-not (Test-Path -LiteralPath $target)) {
    Write-Host "ERROR: deploy_direct.ps1 not found next to this script: $here"
    exit 1
}
& $target
exit $LASTEXITCODE
