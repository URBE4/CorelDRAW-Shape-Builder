$ErrorActionPreference = 'Stop'
$basePath = $PSScriptRoot
if ([string]::IsNullOrEmpty($basePath)) {
    $basePath = Split-Path -Parent $MyInvocation.MyCommand.Path
}
if ([string]::IsNullOrEmpty($basePath)) {
    $basePath = (Get-Location).Path
}

Write-Host '=== deploy_direct.ps1 (Shape Builder) ==='
Write-Host "Project folder: $basePath"
$basCheck = Join-Path $basePath 'CorelDRAW_Furniture_Facades_Macro.bas'
$ufCheck = Join-Path $basePath 'UserForm1_Code.txt'
if (-not (Test-Path -LiteralPath $basCheck)) { throw "File not found: $basCheck - run script from the project folder with .bas file." }
if (-not (Test-Path -LiteralPath $ufCheck)) { throw "File not found: $ufCheck" }
try {
    $vLine = (Select-String -LiteralPath $basCheck -Pattern 'MEBEL_MACRO_VERSION\s*=\s*"' | Select-Object -First 1).Line.Trim()
    Write-Host "Source file version: $vLine"
} catch {
    Write-Host 'Source file version: (could not read)'
}
try {
    $sLine = (Select-String -LiteralPath $ufCheck -Pattern 'SB_CODE_REV\s+As\s+String\s*=\s*"' | Select-Object -First 1).Line.Trim()
    Write-Host "Source form SB: $sLine"
} catch {
    Write-Host 'Source form SB: (could not read)'
}
Write-Host ''

function Get-CorelApplication {
    # Сначала — уже запущенный Corel: без суффикса версии (ROT), иначе часто не находит 2024/2023.
    try {
        $a = [Runtime.InteropServices.Marshal]::GetActiveObject('CorelDRAW.Application')
        if ($a) {
            Write-Host 'Corel: CorelDRAW.Application (running, no version suffix)'
            return $a
        }
    } catch {}
    # Суффикс ProgID: CorelDRAW.Application.<N>. 2026/2025 — в начале списка.
    $versions = @('27', '26', '25', '24', '23', '22', '21')
    foreach ($v in $versions) {
        $progId = "CorelDRAW.Application.$v"
        try {
            $a = [Runtime.InteropServices.Marshal]::GetActiveObject($progId)
            if ($a) {
                Write-Host "Corel: $progId (running)"
                return $a
            }
        } catch {}
    }
    try {
        $a = New-Object -ComObject CorelDRAW.Application
        if ($a) {
            Write-Host 'Corel: started CorelDRAW.Application (no version suffix)'
            $a.Visible = $true
            Start-Sleep -Seconds 3
            return $a
        }
    } catch {}
    foreach ($v in $versions) {
        $progId = "CorelDRAW.Application.$v"
        try {
            $a = New-Object -ComObject $progId
            if ($a) {
                Write-Host "Corel: started $progId"
                $a.Visible = $true
                Start-Sleep -Seconds 3
                return $a
            }
        } catch {}
    }
    return $null
}

function Wait-Vbe([object]$app, [int]$maxWaitMs = 20000) {
    $vbe = $null
    $deadline = [datetime]::UtcNow.AddMilliseconds($maxWaitMs)
    while ([datetime]::UtcNow -lt $deadline) {
        try {
            $vbe = $app.VBE
            if ($vbe) { return $vbe }
        } catch {}
        try { $null = $app.VBE } catch {}
        Start-Sleep -Milliseconds 400
    }
    return $null
}

try {
    $app = Get-CorelApplication
    if (-not $app) {
        throw 'CorelDRAW COM not found. Open CorelDRAW manually, then run this script again.'
    }

    try { if ($app.Documents.Count -eq 0) { $null = $app.CreateDocument() } } catch {}
    try { $null = $app.VBE } catch {}
    Start-Sleep -Seconds 1

    $vbe = Wait-Vbe $app 45000
    if (-not $vbe) {
        throw 'VBE unavailable. In Corel: trust VBA project access; open a document; press Alt+F11 once, close editor, retry deploy_direct.ps1'
    }

    $proj = $null
    foreach ($p in $vbe.VBProjects) {
        try { if ($p.Name -eq 'GlobalMacros') { $proj = $p; break } } catch {}
    }
    if (-not $proj) { throw 'GlobalMacros project not found in VBE' }

    function Sync-StdModule([object]$project, [string]$name, [string]$fileName) {
        $m = $null
        foreach ($c in $project.VBComponents) { if ($c.Name -eq $name) { $m = $c; break } }
        if (-not $m) { $m = $project.VBComponents.Add(1) }
        $n = $m.CodeModule.CountOfLines
        if ($n -gt 0) { $m.CodeModule.DeleteLines(1, $n) }
        $code = [System.IO.File]::ReadAllText((Join-Path $basePath $fileName), [System.Text.Encoding]::UTF8)
        $m.CodeModule.AddFromString($code)
        $m.Name = $name
        Write-Host "$name OK: $($m.CodeModule.CountOfLines)"
    }

    Sync-StdModule $proj 'Module1' 'CorelDRAW_Furniture_Facades_Macro.bas'

    # Module2 часто дубликат со старым RunModernShapeBuilder — пользователь запускал его и видел старую форму
    foreach ($killName in @('ModuleCore', 'ModuleDraw', 'Module2')) {
        $old = $null
        foreach ($c in $proj.VBComponents) { if ($c.Name -eq $killName) { $old = $c; break } }
        if ($old) { try { $proj.VBComponents.Remove($old); Write-Host "Removed $killName" } catch {} }
    }

    $uf = $null
    foreach ($c in $proj.VBComponents) { if ($c.Name -eq 'UserForm1') { $uf = $c; break } }
    if (-not $uf) { $uf = $proj.VBComponents.Add(3) }
    $cnt2 = $uf.CodeModule.CountOfLines
    if ($cnt2 -gt 0) { $uf.CodeModule.DeleteLines(1, $cnt2) }
    $ufCode = [System.IO.File]::ReadAllText((Join-Path $basePath 'UserForm1_Code.txt'), [System.Text.Encoding]::UTF8)
    $uf.CodeModule.AddFromString($ufCode)
    try { $uf.Name = 'UserForm1' } catch {}
    Write-Host "UserForm1 OK: $($uf.CodeModule.CountOfLines)"

    $uf2Path = Join-Path $basePath 'UserForm2_Code.txt'
    if (Test-Path -LiteralPath $uf2Path) {
        $uf2Len = (Get-Item -LiteralPath $uf2Path).Length
        if ($uf2Len -gt 8) {
            $uf2 = $null
            foreach ($c in $proj.VBComponents) { if ($c.Name -eq 'UserForm2') { $uf2 = $c; break } }
            if (-not $uf2) { $uf2 = $proj.VBComponents.Add(3) }
            $nUf2 = $uf2.CodeModule.CountOfLines
            if ($nUf2 -gt 0) { $uf2.CodeModule.DeleteLines(1, $nUf2) }
            $uf2Code = [System.IO.File]::ReadAllText($uf2Path, [System.Text.Encoding]::UTF8)
            $uf2.CodeModule.AddFromString($uf2Code)
            try { $uf2.Name = 'UserForm2' } catch {}
            Write-Host "UserForm2 OK: $($uf2.CodeModule.CountOfLines)"
        } else {
            Write-Host 'UserForm2 skipped (UserForm2_Code.txt too small)'
        }
    } else {
        Write-Host 'UserForm2 skipped (no UserForm2_Code.txt)'
    }

    # --- Persist helpers: crash without File-Save GMS reloads old .gms from disk ---
    $stamp = Get-Date -Format 'yyyy-MM-dd_HHmmss'
    $snapDir = Join-Path $basePath ('vba_backup\deploy_' + $stamp)
    try {
        New-Item -ItemType Directory -Path $snapDir -Force | Out-Null
        Copy-Item -LiteralPath (Join-Path $basePath 'CorelDRAW_Furniture_Facades_Macro.bas') -Destination $snapDir -Force
        Copy-Item -LiteralPath (Join-Path $basePath 'UserForm1_Code.txt') -Destination $snapDir -Force
        $uf2Snap = Join-Path $basePath 'UserForm2_Code.txt'
        if (Test-Path -LiteralPath $uf2Snap) { Copy-Item -LiteralPath $uf2Snap -Destination $snapDir -Force }
        $readme = "Sources snapshot at deploy. If GMS rolled back, run deploy_direct.ps1 from project folder.`r`nTime: $stamp"
        Set-Content -LiteralPath (Join-Path $snapDir 'README.txt') -Encoding UTF8 -Value $readme
        Write-Host "Sources snapshot: $snapDir"
    } catch {
        Write-Host ('Snapshot copy warning: ' + $_.Exception.Message)
    }

    $exportDir = Join-Path $basePath 'vba_export_last'
    try {
        if (-not (Test-Path -LiteralPath $exportDir)) { New-Item -ItemType Directory -Path $exportDir -Force | Out-Null }
        foreach ($c in $proj.VBComponents) {
            $nm = $null
            try { $nm = $c.Name } catch { continue }
            if ($nm -eq 'Module1' -or $nm -eq 'UserForm1' -or $nm -eq 'UserForm2') {
                try {
                    $ext = '.bas'
                    try { if ([int]$c.Type -eq 3) { $ext = '.frm' } } catch {}
                    $out = Join-Path $exportDir ($nm + $ext)
                    $c.Export($out)
                    Write-Host "VBE Export: $out"
                } catch {
                    Write-Host ("VBE Export skip $nm : " + $_.Exception.Message)
                }
            }
        }
    } catch {
        Write-Host ('VBE Export folder warning: ' + $_.Exception.Message)
    }

    try {
        $proj.Save() | Out-Null
        Write-Host 'GlobalMacros: VBProject.Save() OK (rare in Corel).'
    } catch {
        Write-Host 'Next step: Alt+F11, File - Save GlobalMacros (Corel COM has no Save here; avoids rollback on crash).'
    }

    Write-Host 'DIRECT DEPLOY OK.'
} catch {
    Write-Host ("DIRECT DEPLOY ERROR: " + $_.Exception.Message)
    exit 1
}
