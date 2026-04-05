# CorelDRAW Graphics Suite 2026 — основная цель; ProgID .28 / .27 (разные сборки).
# Общий список COM-суффиксов для GetActiveObject / New-Object.
$script:ShapeBuilderCorelComVersions = @('28', '27', '26', '25', '24', '23', '22', '21')

function Get-ShapeBuilderCorelAppRunningOnly {
    try {
        $a = [Runtime.InteropServices.Marshal]::GetActiveObject('CorelDRAW.Application')
        if ($a) { return $a }
    } catch {}
    foreach ($v in $script:ShapeBuilderCorelComVersions) {
        $progId = "CorelDRAW.Application.$v"
        try {
            $a = [Runtime.InteropServices.Marshal]::GetActiveObject($progId)
            if ($a) { return $a }
        } catch {}
    }
    return $null
}

function Get-ShapeBuilderCorelApplication {
    $versions = $script:ShapeBuilderCorelComVersions
    try {
        $a = [Runtime.InteropServices.Marshal]::GetActiveObject('CorelDRAW.Application')
        if ($a) {
            Write-Host 'Corel: CorelDRAW.Application (running, no version suffix)'
            return $a
        }
    } catch {}
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
