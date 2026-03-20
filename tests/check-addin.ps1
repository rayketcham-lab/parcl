Stop-Process -Name OUTLOOK -Force -ErrorAction SilentlyContinue
Start-Sleep 3

Set-ItemProperty 'HKCU:\Software\Microsoft\Office\Outlook\Addins\Parcl.Addin' -Name 'LoadBehavior' -Value 3 -Type DWord
Remove-ItemProperty 'HKCU:\Software\Microsoft\Office\16.0\Outlook\Resiliency\DisabledItems' -Name * -ErrorAction SilentlyContinue
Remove-ItemProperty 'HKCU:\Software\Microsoft\Office\16.0\Outlook\Resiliency\CrashingAddinList' -Name * -ErrorAction SilentlyContinue

Write-Host "LoadBehavior set to 3, resiliency cleared"

Start-Process 'C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE'
Write-Host "Waiting for Outlook..."
Start-Sleep 20

$ol = New-Object -ComObject Outlook.Application
Write-Host "Connected to Outlook"

Write-Host ""
Write-Host "COM Add-ins:"
foreach ($a in $ol.COMAddIns) {
    $status = if ($a.Connect) { "LOADED" } else { "DISABLED" }
    Write-Host "  [$status] $($a.ProgId)"
}

Write-Host ""
$parcl = $null
foreach ($a in $ol.COMAddIns) {
    if ($a.ProgId -eq 'Parcl.Addin') {
        $parcl = $a
        break
    }
}

if ($parcl) {
    Write-Host "Parcl found: Connected=$($parcl.Connect)"
    if (-not $parcl.Connect) {
        Write-Host "Attempting force-connect..."
        try {
            $parcl.Connect = $true
            Write-Host "After force-connect: Connected=$($parcl.Connect)"
        } catch {
            Write-Host "Force-connect failed: $($_.Exception.Message)"
            Write-Host ""
            Write-Host "Inner exception: $($_.Exception.InnerException.Message)"
        }
    }
} else {
    Write-Host "Parcl.Addin NOT found in COMAddIns"
}

$lb = (Get-ItemProperty 'HKCU:\Software\Microsoft\Office\Outlook\Addins\Parcl.Addin').LoadBehavior
Write-Host ""
Write-Host "Final LoadBehavior: $lb"
