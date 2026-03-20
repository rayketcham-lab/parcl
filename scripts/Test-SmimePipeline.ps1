#Requires -Version 5.1
<#
.SYNOPSIS
    Automated S/MIME pipeline verification for Parcl.
    Sends signed+encrypted email to self, verifies native S/MIME delivery.
.EXAMPLE
    .\Test-SmimePipeline.ps1
    .\Test-SmimePipeline.ps1 -TimeoutSeconds 120
#>
param(
    [int]$TimeoutSeconds = 90,
    [bool]$Cleanup = $true
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$script:failures = 0

function Write-Pass($msg)    { Write-Host "  [PASS] $msg" -ForegroundColor Green }
function Write-Fail($msg)    { Write-Host "  [FAIL] $msg" -ForegroundColor Red; $script:failures++ }
function Write-Section($msg) { Write-Host "`n== $msg ==" -ForegroundColor Cyan }

$marker = "PARCL-TEST-" + [guid]::NewGuid().ToString("N").Substring(0, 8)
Write-Host "`nParcl S/MIME Pipeline Test" -ForegroundColor Yellow
Write-Host ("=" * 40)

# ── 1. Settings ──
Write-Section "Settings"
$settingsPath = Join-Path $env:APPDATA "Parcl\settings.json"
if (-not (Test-Path $settingsPath)) { Write-Fail "Settings not found"; exit 1 }

$settings = Get-Content $settingsPath -Raw | ConvertFrom-Json
$To = $settings.UserProfile.EmailAddress
Write-Host "  To/From: $To"
Write-Host "  AlwaysSign: $($settings.Crypto.AlwaysSign)"
Write-Host "  AlwaysEncrypt: $($settings.Crypto.AlwaysEncrypt)"
Write-Host "  UseNativeSmime: $($settings.Crypto.UseNativeSmime)"
Write-Host "  Marker: $marker"

if (-not $settings.Crypto.AlwaysEncrypt) { Write-Fail "AlwaysEncrypt is false" }
if (-not $settings.Crypto.AlwaysSign)    { Write-Fail "AlwaysSign is false" }
if (-not $settings.Crypto.UseNativeSmime){ Write-Fail "UseNativeSmime is false" }

# ── 2. Outlook ──
Write-Section "Outlook"
try {
    $outlook = New-Object -ComObject Outlook.Application
    $ns = $outlook.GetNamespace("MAPI")
    $ns.Logon("Outlook", "", $false, $false)
    $userName = $ns.CurrentUser.Name
    Write-Pass "Connected: $userName"
}
catch {
    Write-Fail "Cannot connect: $($_.Exception.Message)"
    exit 1
}

# ── 3. Signing cert ──
Write-Section "Certificates"
$sigThumb = $settings.UserProfile.SigningCertThumbprint
$sigCert = Get-ChildItem Cert:\CurrentUser\My -ErrorAction SilentlyContinue |
    Where-Object { $_.Thumbprint -eq $sigThumb }
if ($sigCert -and $sigCert.NotAfter -gt (Get-Date)) {
    Write-Pass "Signing cert valid (expires $($sigCert.NotAfter.ToString('yyyy-MM-dd')))"
}
else {
    Write-Fail "Signing cert missing or expired"
}

# ── 4. Send ──
Write-Section "Send"
$mail = $outlook.CreateItem(0)
$mail.To = $To
$mail.Subject = $marker
$mail.Body = "Automated Parcl S/MIME test.`nMarker: $marker`nTime: $(Get-Date -Format o)"
$sendTime = Get-Date
$mail.Send()
Write-Pass "Sent at $($sendTime.ToString('HH:mm:ss'))"

# ── 5. Sent Items ──
Write-Section "Sent Items"
Start-Sleep -Seconds 5
$ns.SendAndReceive($false)
Start-Sleep -Seconds 5

$sentFolder = $ns.GetDefaultFolder(5)

# Search by exact subject or iterate recent items (S/MIME may alter subject)
$sentMsg = $null
try { $sentMsg = $sentFolder.Items.Find("[Subject] = '$marker'") } catch {}

if (-not $sentMsg) {
    $sentItems = $sentFolder.Items
    $sentItems.Sort("[SentOn]", $true)
    for ($i = 1; $i -le [Math]::Min(10, $sentItems.Count); $i++) {
        try {
            $item = $sentItems.Item($i)
            if ($item.Subject -match "PARCL-TEST") { $sentMsg = $item; break }
        }
        catch {}
    }
}

if ($sentMsg) {
    Write-Host "  MessageClass: $($sentMsg.MessageClass)"

    try {
        $flags = $sentMsg.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x6E010003")
        Write-Host "  PR_SECURITY_FLAGS: 0x$($flags.ToString('X'))"
        if ($flags -band 0x01) { Write-Pass "Encrypt flag set" } else { Write-Fail "Encrypt flag NOT set" }
        if ($flags -band 0x02) { Write-Pass "Sign flag set" }    else { Write-Fail "Sign flag NOT set" }
    }
    catch {
        Write-Fail "PR_SECURITY_FLAGS not readable"
    }

    $hasP7m = $false
    for ($i = 1; $i -le $sentMsg.Attachments.Count; $i++) {
        $attName = $sentMsg.Attachments.Item($i).FileName
        if ($attName -match "smime") { $hasP7m = $true }
    }
    if ($hasP7m) { Write-Fail "smime.p7m found — Parcl envelope, not native" }
    else         { Write-Pass "No .p7m attachment — native S/MIME" }
}
else {
    Write-Fail "Not found in Sent Items"
}

# ── 6. Inbox delivery ──
Write-Section "Inbox Delivery (${TimeoutSeconds}s timeout)"
$inbox = $ns.GetDefaultFolder(6)
$found = $null
$elapsed = 0

$inboxFilter = "[Subject] = '$marker'"
while ($elapsed -lt $TimeoutSeconds) {
    try { $ns.SendAndReceive($false) } catch {}
    Start-Sleep -Seconds 5
    $elapsed += 5

    try {
        $found = $inbox.Items.Find($inboxFilter)
    }
    catch {}
    if ($found) { break }
    Write-Host "." -NoNewline
}
Write-Host ""

if (-not $found) {
    Write-Fail "Not delivered within $TimeoutSeconds seconds"
}
else {
    Write-Pass "Delivered in ~${elapsed}s"

    Write-Section "Inbox Verification"
    Write-Host "  MessageClass: $($found.MessageClass)"
    Write-Host "  Attachments: $($found.Attachments.Count)"

    $bodyLen = 0
    try { $bodyLen = $found.Body.Length } catch {}
    if ($bodyLen -gt 0) { Write-Pass "Body readable ($bodyLen chars) — decrypted" }
    else                { Write-Fail "Body empty — decryption failed" }

    $hasP7mInbox = $false
    for ($i = 1; $i -le $found.Attachments.Count; $i++) {
        $attName = $found.Attachments.Item($i).FileName
        Write-Host "  Attachment: $attName"
        if ($attName -match "smime") { $hasP7mInbox = $true }
    }
    if ($hasP7mInbox) { Write-Fail "smime.p7m in inbox — not inline" }
    else              { Write-Pass "No .p7m — message is inline" }

    if ($found.Body -match $marker) { Write-Pass "Content integrity verified" }
    else                            { Write-Fail "Marker not in body" }
}

# ── 7. Parcl log ──
Write-Section "Parcl Log"
$logFile = Join-Path $env:APPDATA "Parcl\logs\parcl-$(Get-Date -Format 'yyyy-MM-dd').jsonl"
if (Test-Path $logFile) {
    $timeStr = $sendTime.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm")
    $hits = Select-String -Path $logFile -Pattern $timeStr | Where-Object { $_.Line -match "Send" }
    foreach ($hit in $hits) {
        try {
            $obj = $hit.Line | ConvertFrom-Json
            Write-Host "  [$($obj.lvl)] $($obj.msg)" -ForegroundColor Gray
        }
        catch {}
    }

    $allLog = Get-Content $logFile -Raw
    if ($allLog -match "Parcl envelope.*$timeStr|$timeStr.*Parcl envelope") {
        Write-Fail "Log shows Parcl envelope fallback"
    }
}

# ── 8. Cleanup ──
if ($Cleanup) {
    Write-Section "Cleanup"
    try { if ($found)   { $found.Delete();   Write-Host "  Inbox message deleted" } } catch {}
    try { if ($sentMsg) { $sentMsg.Delete();  Write-Host "  Sent message deleted" } }  catch {}
}

# ── Verdict ──
Write-Host ""
Write-Host ("=" * 40)
if ($script:failures -eq 0) {
    Write-Host "VERDICT: PASS" -ForegroundColor Green
    exit 0
}
else {
    Write-Host "VERDICT: FAIL ($($script:failures) failures)" -ForegroundColor Red
    exit 1
}
