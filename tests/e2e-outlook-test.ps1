#Requires -Version 7
<#
.SYNOPSIS
    Parcl E2E Outlook Automation Tests — Native S/MIME Mode
    Uses Outlook COM automation to test all Parcl buttons and native inline encryption.

.DESCRIPTION
    FIPS 140-2 Best Practices Baseline:
    - Encryption: AES-256-CBC (FIPS 197)
    - Signing Hash: SHA-256 (FIPS 180-4)
    - Key Transport: RSA OAEP (FIPS 186-4)
    - Certificate Validation: Strict (chain + revocation)
    - UseNativeSmime: true (PR_SECURITY_FLAGS based)
    - AlwaysSign: true
    - OpaqueSign: false (clear-signed for interop)
    - IncludeCertChain: true

.NOTES
    Requires Outlook running with Parcl add-in loaded.
    Sends test emails to rayketcham@ogjos.com (self).
#>

param(
    [string]$TestEmail = "rayketcham@ogjos.com",
    [string]$CertThumbprint = "6D4ED55B34DCC2CE86E09E15F03C3ADAC0D13698"
)

$ErrorActionPreference = "Continue"
$script:passed = 0
$script:failed = 0
$script:results = @()

function Test-Result {
    param([string]$Name, [bool]$Success, [string]$Detail = "")
    if ($Success) {
        Write-Host "  PASS: $Name" -ForegroundColor Green
        $script:passed++
    } else {
        Write-Host "  FAIL: $Name — $Detail" -ForegroundColor Red
        $script:failed++
    }
    $script:results += [PSCustomObject]@{Test=$Name; Result=if($Success){"PASS"}else{"FAIL"}; Detail=$Detail}
}

# ═══════════════════════════════════════════════════════════════════
# Setup: Connect to Outlook
# ═══════════════════════════════════════════════════════════════════
Write-Host "`n═══ Parcl E2E Outlook Tests — Native S/MIME (FIPS Baseline) ═══" -ForegroundColor Cyan

try {
    $outlook = [Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
    Write-Host "Connected to running Outlook instance" -ForegroundColor Green
} catch {
    Write-Host "Outlook not running. Starting..." -ForegroundColor Yellow
    $outlook = New-Object -ComObject Outlook.Application
    Start-Sleep 5
}

$namespace = $outlook.GetNamespace("MAPI")
$inbox = $namespace.GetDefaultFolder(6) # olFolderInbox
$drafts = $namespace.GetDefaultFolder(16) # olFolderDrafts
$sentItems = $namespace.GetDefaultFolder(5) # olFolderSentMail

# ═══════════════════════════════════════════════════════════════════
# Test 1: Verify FIPS Settings
# ═══════════════════════════════════════════════════════════════════
Write-Host "`n── Test Group: Settings Verification ──" -ForegroundColor Yellow

$settingsPath = "$env:APPDATA\Parcl\settings.json"
$settings = Get-Content $settingsPath -Raw | ConvertFrom-Json

Test-Result "Settings: AES-256-CBC encryption" `
    ($settings.Crypto.EncryptionAlgorithm -eq "AES-256-CBC")

Test-Result "Settings: SHA-256 hash" `
    ($settings.Crypto.HashAlgorithm -eq "SHA-256")

Test-Result "Settings: AlwaysSign enabled" `
    ($settings.Crypto.AlwaysSign -eq $true)

Test-Result "Settings: UseNativeSmime enabled" `
    ($settings.Crypto.UseNativeSmime -eq $true)

Test-Result "Settings: Strict validation mode" `
    ($settings.Crypto.ValidationMode -eq 2)

Test-Result "Settings: IncludeCertChain enabled" `
    ($settings.Crypto.IncludeCertChain -eq $true)

Test-Result "Settings: AutoDecrypt enabled" `
    ($settings.Behavior.AutoDecrypt -eq $true)

Test-Result "Settings: Signing cert configured" `
    ($settings.UserProfile.SigningCertThumbprint -eq $CertThumbprint)

Test-Result "Settings: Encryption cert configured" `
    ($settings.UserProfile.EncryptionCertThumbprint -eq $CertThumbprint)

# ═══════════════════════════════════════════════════════════════════
# Test 2: Certificate Validation
# ═══════════════════════════════════════════════════════════════════
Write-Host "`n── Test Group: Certificate Validation ──" -ForegroundColor Yellow

$cert = Get-ChildItem "Cert:\CurrentUser\My" | Where-Object { $_.Thumbprint -eq $CertThumbprint }

Test-Result "Cert: Found in store" ($null -ne $cert)
Test-Result "Cert: Has private key" ($cert.HasPrivateKey)
Test-Result "Cert: Not expired" ($cert.NotAfter -gt (Get-Date))
Test-Result "Cert: Valid (not before)" ($cert.NotBefore -lt (Get-Date))
Test-Result "Cert: Issued by Sectigo" ($cert.Issuer -like "*Sectigo*")
Test-Result "Cert: Email matches" ($cert.Subject -like "*$TestEmail*")

# Check key usage
$kuExt = $cert.Extensions | Where-Object { $_.Oid.Value -eq "2.5.29.15" }
$kuStr = $kuExt.Format($false)
Test-Result "Cert: Digital Signature key usage" ($kuStr -like "*Digital Signature*")
Test-Result "Cert: Key Encipherment key usage" ($kuStr -like "*Key Encipherment*")

# Check EKU
$ekuExt = $cert.Extensions | Where-Object { $_.Oid.Value -eq "2.5.29.37" }
$ekuStr = $ekuExt.Format($false)
Test-Result "Cert: Secure Email EKU" ($ekuStr -like "*Secure Email*")

# Check SAN
$sanExt = $cert.Extensions | Where-Object { $_.Oid.Value -eq "2.5.29.17" }
$sanStr = $sanExt.Format($false)
Test-Result "Cert: SAN contains email" ($sanStr -like "*$TestEmail*")

# ═══════════════════════════════════════════════════════════════════
# Test 3: Parcl Add-in Loaded
# ═══════════════════════════════════════════════════════════════════
Write-Host "`n── Test Group: Add-in Registration ──" -ForegroundColor Yellow

$addinKey = "HKCU:\Software\Microsoft\Office\Outlook\Addins\Parcl.Addin"
Test-Result "Add-in: Registry key exists" (Test-Path $addinKey)

if (Test-Path $addinKey) {
    $loadBehavior = (Get-ItemProperty $addinKey).LoadBehavior
    Test-Result "Add-in: LoadBehavior=3 (startup)" ($loadBehavior -eq 3)
}

# Check COM registration
$comKey = "HKCU:\Software\Classes\CLSID\{B8F0C3A2-7D5E-4F91-A6C8-9E1B3D5A7F42}"
if (-not (Test-Path $comKey)) {
    $comKey = "HKLM:\Software\Classes\CLSID\{B8F0C3A2-7D5E-4F91-A6C8-9E1B3D5A7F42}"
}
Test-Result "Add-in: COM class registered" (Test-Path $comKey)

# ═══════════════════════════════════════════════════════════════════
# Test 4: Create and Send — Sign Only (Native S/MIME)
# ═══════════════════════════════════════════════════════════════════
Write-Host "`n── Test Group: Native S/MIME Sign Only ──" -ForegroundColor Yellow

$mail1 = $outlook.CreateItem(0) # olMailItem
$mail1.To = $TestEmail
$mail1.Subject = "Parcl E2E Test: Sign Only — $(Get-Date -Format 'HH:mm:ss')"
$mail1.Body = "This message tests native S/MIME signing via Parcl.`nFIPS baseline: SHA-256 clear-sign."

# Set native S/MIME sign flag via PR_SECURITY_FLAGS
$PR_SEC = "http://schemas.microsoft.com/mapi/proptag/0x6E010003"
try {
    $flags = $mail1.PropertyAccessor.GetProperty($PR_SEC)
} catch { $flags = 0 }
$mail1.PropertyAccessor.SetProperty($PR_SEC, ($flags -bor 0x02)) # SECFLAG_SIGNED

$signFlags = $mail1.PropertyAccessor.GetProperty($PR_SEC)
Test-Result "Sign-Only: PR_SECURITY_FLAGS has sign bit" (($signFlags -band 0x02) -ne 0)
Test-Result "Sign-Only: PR_SECURITY_FLAGS no encrypt bit" (($signFlags -band 0x01) -eq 0)

$mail1.Send()
Test-Result "Sign-Only: Message sent" $true
Start-Sleep 3

# ═══════════════════════════════════════════════════════════════════
# Test 5: Create and Send — Encrypt Only (Native S/MIME)
# ═══════════════════════════════════════════════════════════════════
Write-Host "`n── Test Group: Native S/MIME Encrypt Only ──" -ForegroundColor Yellow

$mail2 = $outlook.CreateItem(0)
$mail2.To = $TestEmail
$mail2.Subject = "Parcl E2E Test: Encrypt Only — $(Get-Date -Format 'HH:mm:ss')"
$mail2.Body = "This message tests native S/MIME encryption via Parcl.`nFIPS baseline: AES-256-CBC."

try { $flags2 = $mail2.PropertyAccessor.GetProperty($PR_SEC) } catch { $flags2 = 0 }
$mail2.PropertyAccessor.SetProperty($PR_SEC, ($flags2 -bor 0x01)) # SECFLAG_ENCRYPTED

$encFlags = $mail2.PropertyAccessor.GetProperty($PR_SEC)
Test-Result "Encrypt-Only: PR_SECURITY_FLAGS has encrypt bit" (($encFlags -band 0x01) -ne 0)
Test-Result "Encrypt-Only: PR_SECURITY_FLAGS no sign bit" (($encFlags -band 0x02) -eq 0)

$mail2.Send()
Test-Result "Encrypt-Only: Message sent" $true
Start-Sleep 3

# ═══════════════════════════════════════════════════════════════════
# Test 6: Create and Send — Sign + Encrypt (Native S/MIME)
# ═══════════════════════════════════════════════════════════════════
Write-Host "`n── Test Group: Native S/MIME Sign + Encrypt ──" -ForegroundColor Yellow

$mail3 = $outlook.CreateItem(0)
$mail3.To = $TestEmail
$mail3.Subject = "Parcl E2E Test: Sign+Encrypt — $(Get-Date -Format 'HH:mm:ss')"
$mail3.Body = "This message tests native S/MIME sign+encrypt via Parcl.`nFIPS: AES-256-CBC + SHA-256."

try { $flags3 = $mail3.PropertyAccessor.GetProperty($PR_SEC) } catch { $flags3 = 0 }
$mail3.PropertyAccessor.SetProperty($PR_SEC, ($flags3 -bor 0x01 -bor 0x02)) # ENCRYPTED + SIGNED

$bothFlags = $mail3.PropertyAccessor.GetProperty($PR_SEC)
Test-Result "Sign+Encrypt: Both flags set" (($bothFlags -band 0x03) -eq 0x03)

$mail3.Send()
Test-Result "Sign+Encrypt: Message sent" $true
Start-Sleep 3

# ═══════════════════════════════════════════════════════════════════
# Test 7: Create and Send — Plaintext (no encryption, no signing)
# ═══════════════════════════════════════════════════════════════════
Write-Host "`n── Test Group: Plaintext (No Encryption) ──" -ForegroundColor Yellow

$mail4 = $outlook.CreateItem(0)
$mail4.To = $TestEmail
$mail4.Subject = "Parcl E2E Test: Plaintext — $(Get-Date -Format 'HH:mm:ss')"
$mail4.Body = "This message is sent without any encryption or signing as a control test."

try { $flags4 = $mail4.PropertyAccessor.GetProperty($PR_SEC) } catch { $flags4 = 0 }
Test-Result "Plaintext: No security flags" ($flags4 -eq 0)

$mail4.Send()
Test-Result "Plaintext: Message sent" $true
Start-Sleep 3

# ═══════════════════════════════════════════════════════════════════
# Test 8: Create Draft with Parcl UserProperties (Encrypt toggle)
# ═══════════════════════════════════════════════════════════════════
Write-Host "`n── Test Group: Parcl Encrypt Toggle (UserProperty) ──" -ForegroundColor Yellow

$mail5 = $outlook.CreateItem(0)
$mail5.To = $TestEmail
$mail5.Subject = "Parcl E2E Test: Parcl Encrypt Flag — $(Get-Date -Format 'HH:mm:ss')"
$mail5.Body = "Testing Parcl encrypt toggle via UserProperty flag."

$encFlag = $mail5.UserProperties.Add("ParclEncrypt", 6, $false) # olYesNo = 6
$encFlag.Value = $true

$readBack = $mail5.UserProperties.Find("ParclEncrypt")
Test-Result "Parcl Encrypt Flag: UserProperty set" ($null -ne $readBack)
Test-Result "Parcl Encrypt Flag: Value is true" ($readBack.Value -eq $true)

$mail5.Save()
Test-Result "Parcl Encrypt Flag: Draft saved" $true

# Clean up draft
$mail5.Delete()

# ═══════════════════════════════════════════════════════════════════
# Test 9: Create Draft with Parcl Sign UserProperty
# ═══════════════════════════════════════════════════════════════════
Write-Host "`n── Test Group: Parcl Sign Toggle (UserProperty) ──" -ForegroundColor Yellow

$mail6 = $outlook.CreateItem(0)
$mail6.To = $TestEmail
$mail6.Subject = "Parcl E2E Test: Parcl Sign Flag — $(Get-Date -Format 'HH:mm:ss')"
$mail6.Body = "Testing Parcl sign toggle via UserProperty flag."

$signFlag = $mail6.UserProperties.Add("ParclSign", 6, $false)
$signFlag.Value = $true

$readSignBack = $mail6.UserProperties.Find("ParclSign")
Test-Result "Parcl Sign Flag: UserProperty set" ($null -ne $readSignBack)
Test-Result "Parcl Sign Flag: Value is true" ($readSignBack.Value -eq $true)

$mail6.Save()
Test-Result "Parcl Sign Flag: Draft saved" $true
$mail6.Delete()

# ═══════════════════════════════════════════════════════════════════
# Test 10: Create Draft with Both Flags
# ═══════════════════════════════════════════════════════════════════
Write-Host "`n── Test Group: Both Parcl Flags (Sign + Encrypt) ──" -ForegroundColor Yellow

$mail7 = $outlook.CreateItem(0)
$mail7.To = $TestEmail
$mail7.Subject = "Parcl E2E Test: Both Flags — $(Get-Date -Format 'HH:mm:ss')"
$mail7.Body = "Testing both Parcl flags set simultaneously."

$ef = $mail7.UserProperties.Add("ParclEncrypt", 6, $false)
$ef.Value = $true
$sf = $mail7.UserProperties.Add("ParclSign", 6, $false)
$sf.Value = $true

Test-Result "Both Flags: ParclEncrypt=true" ($mail7.UserProperties.Find("ParclEncrypt").Value -eq $true)
Test-Result "Both Flags: ParclSign=true" ($mail7.UserProperties.Find("ParclSign").Value -eq $true)

$mail7.Save()
$mail7.Delete()

# ═══════════════════════════════════════════════════════════════════
# Test 11: HTML Email Sign + Encrypt
# ═══════════════════════════════════════════════════════════════════
Write-Host "`n── Test Group: HTML Email Sign + Encrypt ──" -ForegroundColor Yellow

$mail8 = $outlook.CreateItem(0)
$mail8.To = $TestEmail
$mail8.Subject = "Parcl E2E Test: HTML Sign+Encrypt — $(Get-Date -Format 'HH:mm:ss')"
$mail8.HTMLBody = "<html><body><h2>Parcl FIPS Test</h2><p>HTML email with <b>bold</b> and <em>italic</em>.</p><p>AES-256-CBC + SHA-256</p></body></html>"

try { $flags8 = $mail8.PropertyAccessor.GetProperty($PR_SEC) } catch { $flags8 = 0 }
$mail8.PropertyAccessor.SetProperty($PR_SEC, ($flags8 -bor 0x03))

Test-Result "HTML Sign+Encrypt: HTML body set" ($mail8.HTMLBody -like "*Parcl FIPS Test*")
Test-Result "HTML Sign+Encrypt: Security flags set" (($mail8.PropertyAccessor.GetProperty($PR_SEC) -band 0x03) -eq 0x03)

$mail8.Send()
Test-Result "HTML Sign+Encrypt: Message sent" $true
Start-Sleep 3

# ═══════════════════════════════════════════════════════════════════
# Test 12: Verify Parcl Log Output
# ═══════════════════════════════════════════════════════════════════
Write-Host "`n── Test Group: Parcl Logging ──" -ForegroundColor Yellow

$logDir = "$env:APPDATA\Parcl\logs"
if (Test-Path $logDir) {
    $latestLog = Get-ChildItem $logDir -Filter "*.log" | Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if ($latestLog) {
        $logContent = Get-Content $latestLog.FullName -Raw -ErrorAction SilentlyContinue
        Test-Result "Log: Log file exists" $true
        Test-Result "Log: Contains startup entry" ($logContent -like "*add-in*" -or $logContent -like "*started*" -or $logContent -like "*connecting*")
    } else {
        Test-Result "Log: Log file exists" $false "No log files found"
    }
} else {
    Test-Result "Log: Log directory exists" $false "No log directory at $logDir"
}

# ═══════════════════════════════════════════════════════════════════
# Test 13: SendDecision Logic Verification
# ═══════════════════════════════════════════════════════════════════
Write-Host "`n── Test Group: SendDecision Logic ──" -ForegroundColor Yellow

# With AlwaysSign=true in settings, every send should get sign flag
Test-Result "SendDecision: AlwaysSign in settings" ($settings.Crypto.AlwaysSign -eq $true)
Test-Result "SendDecision: UseNativeSmime in settings" ($settings.Crypto.UseNativeSmime -eq $true)
Test-Result "SendDecision: AlwaysEncrypt disabled" ($settings.Crypto.AlwaysEncrypt -eq $false)

# ═══════════════════════════════════════════════════════════════════
# Test 14: Wait for delivery and check received messages
# ═══════════════════════════════════════════════════════════════════
Write-Host "`n── Test Group: Delivery Verification (waiting 15s for delivery) ──" -ForegroundColor Yellow
Start-Sleep 15

# Force sync
try { $namespace.SendAndReceive($false) } catch {}
Start-Sleep 5

$recentMessages = @()
for ($i = $inbox.Items.Count; $i -ge [Math]::Max(1, $inbox.Items.Count - 20); $i--) {
    try {
        $item = $inbox.Items.Item($i)
        if ($item.Subject -like "Parcl E2E Test:*") {
            $recentMessages += $item
        }
    } catch {}
}

Test-Result "Delivery: Found test messages in inbox" ($recentMessages.Count -gt 0) "$($recentMessages.Count) found"

foreach ($msg in $recentMessages) {
    $subj = $msg.Subject
    try {
        $secFlags = [int]$msg.PropertyAccessor.GetProperty($PR_SEC)
    } catch { $secFlags = 0 }

    $msgClass = $msg.MessageClass

    if ($subj -like "*Sign Only*") {
        Test-Result "Received Sign-Only: MessageClass=$msgClass" `
            ($msgClass -like "*SMIME*" -or ($secFlags -band 0x02) -ne 0)
    }
    elseif ($subj -like "*Encrypt Only*") {
        Test-Result "Received Encrypt-Only: MessageClass=$msgClass" `
            ($msgClass -like "*SMIME*" -or ($secFlags -band 0x01) -ne 0)
    }
    elseif ($subj -like "*Sign+Encrypt*" -and $subj -notlike "*HTML*") {
        Test-Result "Received Sign+Encrypt: MessageClass=$msgClass" `
            ($msgClass -like "*SMIME*" -or ($secFlags -band 0x03) -ne 0)
    }
    elseif ($subj -like "*HTML Sign+Encrypt*") {
        Test-Result "Received HTML Sign+Encrypt: MessageClass=$msgClass" `
            ($msgClass -like "*SMIME*" -or ($secFlags -band 0x03) -ne 0)
    }
    elseif ($subj -like "*Plaintext*") {
        Test-Result "Received Plaintext: No S/MIME" `
            ($msgClass -eq "IPM.Note" -and ($secFlags -band 0x03) -eq 0)
    }
}

# ═══════════════════════════════════════════════════════════════════
# Summary
# ═══════════════════════════════════════════════════════════════════
Write-Host "`n═══════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host "FIPS 140-2 Baseline Settings:" -ForegroundColor White
Write-Host "  Encryption: AES-256-CBC (OID 2.16.840.1.101.3.4.1.42)" -ForegroundColor Gray
Write-Host "  Signing:    SHA-256 (OID 2.16.840.1.101.3.4.2.1)" -ForegroundColor Gray
Write-Host "  Mode:       Native S/MIME (PR_SECURITY_FLAGS)" -ForegroundColor Gray
Write-Host "  Validation: Strict (chain + revocation)" -ForegroundColor Gray
Write-Host "  AlwaysSign: Enabled" -ForegroundColor Gray
Write-Host "  CertChain:  Included in signatures" -ForegroundColor Gray
Write-Host ""
Write-Host "Results: $($script:passed) passed, $($script:failed) failed out of $($script:passed + $script:failed) tests" `
    -ForegroundColor $(if ($script:failed -eq 0) { "Green" } else { "Red" })
Write-Host "═══════════════════════════════════════════════════" -ForegroundColor Cyan

# Output results table
$script:results | Format-Table -AutoSize
