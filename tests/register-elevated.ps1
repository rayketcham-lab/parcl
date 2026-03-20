# This script must run elevated (Run as Administrator)
param(
    [string]$Configuration = "Debug"
)

$dll = Join-Path $PSScriptRoot "..\src\Parcl.Addin\bin\$Configuration\net48\Parcl.Addin.dll"
if (-not (Test-Path $dll)) {
    Write-Error "DLL not found: $dll — build the project first."
    exit 1
}

# Unregister first
Write-Host "Unregistering..."
& 'C:\Windows\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe' /unregister $dll 2>&1

# Register with /codebase
Write-Host "Registering..."
& 'C:\Windows\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe' /codebase $dll 2>&1

# Verify
Write-Host ""
Write-Host "Verifying..."
$key = [Microsoft.Win32.Registry]::ClassesRoot.OpenSubKey('CLSID\{B8F0C3A2-7D5E-4F91-A6C8-9E1B3D5A7F42}\InprocServer32')
if ($key) {
    Write-Host "  Assembly: $($key.GetValue('Assembly'))"
    Write-Host "  CodeBase: $($key.GetValue('CodeBase'))"
    Write-Host "  Class: $($key.GetValue('Class'))"
    $key.Close()
} else {
    Write-Host "  CLSID NOT FOUND"
}

# Test COM
Write-Host ""
try {
    $type = [Type]::GetTypeFromProgID('Parcl.Addin', $true)
    $obj = [Activator]::CreateInstance($type)
    Write-Host "COM: SUCCESS - $($obj.GetType().FullName)"
} catch {
    Write-Host "COM: $($_.Exception.Message)"
}
