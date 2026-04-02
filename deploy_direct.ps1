$ErrorActionPreference = 'Stop'
$basePath = $PSScriptRoot

try {
    try {
        $app = [Runtime.InteropServices.Marshal]::GetActiveObject('CorelDRAW.Application.25')
    } catch {
        $app = New-Object -ComObject CorelDRAW.Application.25
        $app.Visible = $true
        Start-Sleep -Seconds 3
    }

    try { if ($app.Documents.Count -eq 0) { $null = $app.CreateDocument() } } catch {}
    try { $null = $app.VBE } catch {}
    Start-Sleep -Seconds 1

    $vbe = $app.VBE
    if (-not $vbe) { throw 'VBE unavailable' }

    $proj = $null
    foreach ($p in $vbe.VBProjects) {
        try { if ($p.Name -eq 'GlobalMacros') { $proj = $p; break } } catch {}
    }
    if (-not $proj) { throw 'GlobalMacros not found in VBE projects' }

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

    foreach ($killName in @('ModuleCore', 'ModuleDraw')) {
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

    Write-Host 'DIRECT DEPLOY OK'
} catch {
    Write-Host ("DIRECT DEPLOY ERROR: " + $_.Exception.Message)
}
