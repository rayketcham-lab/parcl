# End-to-end test: Send encrypted/signed emails to self, verify round-trip
$ErrorActionPreference = "Continue"

Write-Host "=== Parcl E2E Test ===" -ForegroundColor Cyan
Write-Host ""

# Wait for Outlook to be ready
$ol = $null
try {
    $ol = [Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
    Write-Host "Outlook COM ready via ROT (v$($ol.Version))" -ForegroundColor Green
} catch {
    try {
        $ol = New-Object -ComObject Outlook.Application
        Write-Host "Outlook COM ready via New-Object (v$($ol.Version))" -ForegroundColor Green
    } catch {
        Write-Host "ERROR: Cannot connect to Outlook" -ForegroundColor Red
        exit 1
    }
}

# Check add-in loaded by looking at recent log
$logDir = Join-Path $env:APPDATA "Parcl\logs"
$today = Get-Date -Format "yyyy-MM-dd"
$logFile = Join-Path $logDir "parcl-${today}.jsonl"
Start-Sleep -Seconds 5

if (Test-Path $logFile) {
    $lastLines = Get-Content $logFile -Tail 5
    $addinLoaded = $lastLines | Where-Object { $_ -match "Parcl add-in started successfully" }
    if ($addinLoaded) {
        Write-Host "Parcl add-in loaded successfully" -ForegroundColor Green
    } else {
        Write-Host "WARNING: Add-in may not have loaded - checking log..." -ForegroundColor Yellow
        $lastLines | ForEach-Object { Write-Host "  $_" }
    }
} else {
    Write-Host "WARNING: No Parcl log found at $logFile" -ForegroundColor Yellow
}

$testEmail = "rayketcham@ogjos.com"
$ns = $ol.GetNamespace("MAPI")

# ── Test 1: Encrypt-only ─────────────────────────────────────────────
Write-Host ""
Write-Host "--- Test 1: Encrypt-only ---" -ForegroundColor Yellow
$mail1 = $ol.CreateItem(0) # olMailItem
$mail1.To = $testEmail
$mail1.Subject = "Parcl E2E Test 1 - Encrypt Only"
$mail1.Body = "This message should be encrypted but not signed. Timestamp: $(Get-Date -Format 'o')"

# Set the ParclEncrypt user property (same as clicking the Encrypt toggle)
$prop1 = $mail1.UserProperties.Add("ParclEncrypt", 6, $false) # olYesNo = 6
$prop1.Value = $true

try {
    $mail1.Send()
    Write-Host "  SENT: Encrypt-only message" -ForegroundColor Green
} catch {
    Write-Host "  FAILED: $($_.Exception.Message)" -ForegroundColor Red
}

Start-Sleep -Seconds 2

# ── Test 2: Sign-only ────────────────────────────────────────────────
Write-Host ""
Write-Host "--- Test 2: Sign-only ---" -ForegroundColor Yellow
$mail2 = $ol.CreateItem(0)
$mail2.To = $testEmail
$mail2.Subject = "Parcl E2E Test 2 - Sign Only"
$mail2.Body = "This message should be signed but not encrypted. Timestamp: $(Get-Date -Format 'o')"

$prop2 = $mail2.UserProperties.Add("ParclSign", 6, $false)
$prop2.Value = $true

try {
    $mail2.Send()
    Write-Host "  SENT: Sign-only message" -ForegroundColor Green
} catch {
    Write-Host "  FAILED: $($_.Exception.Message)" -ForegroundColor Red
}

Start-Sleep -Seconds 2

# ── Test 3: Sign + Encrypt ───────────────────────────────────────────
Write-Host ""
Write-Host "--- Test 3: Sign + Encrypt ---" -ForegroundColor Yellow
$mail3 = $ol.CreateItem(0)
$mail3.To = $testEmail
$mail3.Subject = "Parcl E2E Test 3 - Sign and Encrypt"
$mail3.HTMLBody = "<html><body><h2>Parcl E2E Test</h2><p>This message is <b>signed and encrypted</b>.</p><p>Timestamp: $(Get-Date -Format 'o')</p></body></html>"

$prop3e = $mail3.UserProperties.Add("ParclEncrypt", 6, $false)
$prop3e.Value = $true
$prop3s = $mail3.UserProperties.Add("ParclSign", 6, $false)
$prop3s.Value = $true

try {
    $mail3.Send()
    Write-Host "  SENT: Sign+Encrypt message" -ForegroundColor Green
} catch {
    Write-Host "  FAILED: $($_.Exception.Message)" -ForegroundColor Red
}

Start-Sleep -Seconds 2

# ── Test 4: Encrypt with attachment ──────────────────────────────────
Write-Host ""
Write-Host "--- Test 4: Encrypt with attachment ---" -ForegroundColor Yellow
$mail4 = $ol.CreateItem(0)
$mail4.To = $testEmail
$mail4.Subject = "Parcl E2E Test 4 - Encrypt with Attachment"
$mail4.Body = "This encrypted message has an attachment. Timestamp: $(Get-Date -Format 'o')"

# Create a temp test file
$testAttachment = Join-Path $env:TEMP "parcl-test-attachment.txt"
"This is a test attachment for Parcl E2E testing." | Set-Content $testAttachment
$mail4.Attachments.Add($testAttachment)

$prop4 = $mail4.UserProperties.Add("ParclEncrypt", 6, $false)
$prop4.Value = $true

try {
    $mail4.Send()
    Write-Host "  SENT: Encrypt+Attachment message" -ForegroundColor Green
} catch {
    Write-Host "  FAILED: $($_.Exception.Message)" -ForegroundColor Red
}

Remove-Item $testAttachment -ErrorAction SilentlyContinue

# ── Check Parcl logs for send results ─────────────────────────────────
Write-Host ""
Write-Host "--- Checking Parcl logs ---" -ForegroundColor Yellow
Start-Sleep -Seconds 3

if (Test-Path $logFile) {
    $recentLines = Get-Content $logFile -Tail 30
    $errors = $recentLines | Where-Object { $_ -match '"lvl":"ERROR"' }
    $encapsulated = $recentLines | Where-Object { $_ -match 'S/MIME encapsulated' }
    $signApplied = $recentLines | Where-Object { $_ -match 'signature flag applied' }
    $sendBlocked = $recentLines | Where-Object { $_ -match 'send blocked' }

    Write-Host "  Successful encryptions: $(@($encapsulated).Count)" -ForegroundColor $(if (@($encapsulated).Count -ge 3) { "Green" } else { "Red" })
    Write-Host "  Sign-only messages: $(@($signApplied).Count)" -ForegroundColor $(if (@($signApplied).Count -ge 1) { "Green" } else { "Yellow" })
    Write-Host "  Blocked sends: $(@($sendBlocked).Count)" -ForegroundColor $(if (@($sendBlocked).Count -eq 0) { "Green" } else { "Red" })
    Write-Host "  Errors: $(@($errors).Count)" -ForegroundColor $(if (@($errors).Count -eq 0) { "Green" } else { "Red" })

    if (@($errors).Count -gt 0) {
        Write-Host ""
        Write-Host "  Error details:" -ForegroundColor Red
        $errors | ForEach-Object {
            $json = $_ | ConvertFrom-Json
            Write-Host "    [$($json.cmp)] $($json.msg)" -ForegroundColor Red
            if ($json.err) { Write-Host "      $($json.err.type): $($json.err.msg)" -ForegroundColor Red }
        }
    }
} else {
    Write-Host "  No log file found" -ForegroundColor Red
}

# ── Wait for messages to arrive in inbox, then verify ─────────────────
Write-Host ""
Write-Host "--- Waiting 30s for messages to arrive in inbox ---" -ForegroundColor Yellow
Start-Sleep -Seconds 30

$inbox = $ns.GetDefaultFolder(6) # olFolderInbox
$sentItems = $ns.GetDefaultFolder(5) # olFolderSentMail

# Check sent items for our encrypted messages
$testMessages = @()
for ($i = $sentItems.Items.Count; $i -ge [Math]::Max(1, $sentItems.Items.Count - 10); $i--) {
    $item = $sentItems.Items.Item($i)
    if ($item.Subject -match "Parcl E2E Test" -or $item.Subject -eq "Encrypted Message") {
        $testMessages += $item
    }
}

Write-Host "  Found $($testMessages.Count) test message(s) in Sent Items" -ForegroundColor $(if ($testMessages.Count -ge 3) { "Green" } else { "Yellow" })

foreach ($msg in $testMessages) {
    $hasP7m = $false
    for ($j = 1; $j -le $msg.Attachments.Count; $j++) {
        if ($msg.Attachments.Item($j).FileName -match "\.p7m$") {
            $hasP7m = $true
            break
        }
    }
    $msgClass = "unknown"
    try { $msgClass = $msg.MessageClass } catch {}
    Write-Host "    Subject: $($msg.Subject) | Class: $msgClass | HasP7M: $hasP7m" -ForegroundColor Gray
}

# Check inbox for received copies
$inboxTests = @()
for ($i = $inbox.Items.Count; $i -ge [Math]::Max(1, $inbox.Items.Count - 15); $i--) {
    $item = $inbox.Items.Item($i)
    if ($item.Subject -match "Parcl E2E Test" -or $item.Subject -eq "Encrypted Message") {
        $inboxTests += $item
    }
}

Write-Host ""
Write-Host "  Found $($inboxTests.Count) test message(s) in Inbox" -ForegroundColor $(if ($inboxTests.Count -ge 3) { "Green" } else { "Yellow" })

foreach ($msg in $inboxTests) {
    $hasP7m = $false
    $p7mSize = 0
    for ($j = 1; $j -le $msg.Attachments.Count; $j++) {
        if ($msg.Attachments.Item($j).FileName -match "\.p7m$") {
            $hasP7m = $true
            $p7mSize = $msg.Attachments.Item($j).Size
            break
        }
    }
    $msgClass = "unknown"
    try { $msgClass = $msg.MessageClass } catch {}
    Write-Host "    Subject: $($msg.Subject) | Class: $msgClass | HasP7M: $hasP7m | P7MSize: $p7mSize" -ForegroundColor Gray
}

# ── Attempt programmatic decrypt on inbox messages with .p7m ──────────
Write-Host ""
Write-Host "--- Attempting programmatic decrypt ---" -ForegroundColor Yellow

Add-Type -Path "C:\builds\parcl\src\Parcl.Core\bin\Debug\net48\Parcl.Core.dll"

foreach ($msg in $inboxTests) {
    $hasP7m = $false
    $p7mIdx = -1
    for ($j = 1; $j -le $msg.Attachments.Count; $j++) {
        if ($msg.Attachments.Item($j).FileName -match "\.p7m$") {
            $hasP7m = $true
            $p7mIdx = $j
            break
        }
    }

    if (-not $hasP7m) {
        Write-Host "  [$($msg.Subject)] - no .p7m attachment, skipping" -ForegroundColor Gray
        continue
    }

    Write-Host "  [$($msg.Subject)] - decrypting..." -ForegroundColor Cyan

    # Save .p7m to temp, decrypt via SmimeHandler
    $tempP7m = Join-Path $env:TEMP "parcl-e2e-test.p7m"
    $msg.Attachments.Item($p7mIdx).SaveAsFile($tempP7m)
    $encData = [System.IO.File]::ReadAllBytes($tempP7m)
    Remove-Item $tempP7m -ErrorAction SilentlyContinue

    $handler = New-Object Parcl.Core.Crypto.SmimeHandler
    $decResult = $handler.Decrypt($encData)

    if ($decResult.Success -and $null -ne $decResult.Content) {
        Write-Host "    CMS decrypt OK ($($decResult.Content.Length) bytes)" -ForegroundColor Green

        # Try to unwrap SignedCms
        $mimeBytes = $decResult.Content
        $wasSigned = $false
        try {
            $signedCms = New-Object System.Security.Cryptography.Pkcs.SignedCms
            $signedCms.Decode($mimeBytes)
            $signedCms.CheckSignature($false)
            $mimeBytes = $signedCms.ContentInfo.Content
            $wasSigned = $true
            $signerSubject = $signedCms.SignerInfos[0].Certificate.Subject
            Write-Host "    Signature verified: $signerSubject" -ForegroundColor Green
        } catch {
            # Not signed - that's fine
        }

        # Parse MIME
        $mimeText = [System.Text.Encoding]::UTF8.GetString($mimeBytes)
        $extracted = [Parcl.Core.Crypto.MimeBuilder]::ExtractBody($mimeText)
        $headers = [Parcl.Core.Crypto.MimeBuilder]::ExtractProtectedHeaders($mimeText)

        if ($extracted.HasContent) {
            if ($extracted.HtmlBody) {
                $preview = $extracted.HtmlBody.Substring(0, [Math]::Min(80, $extracted.HtmlBody.Length))
                Write-Host "    HTML body extracted: $preview..." -ForegroundColor Green
            } elseif ($extracted.TextBody) {
                $preview = $extracted.TextBody.Substring(0, [Math]::Min(80, $extracted.TextBody.Length))
                Write-Host "    Text body extracted: $preview..." -ForegroundColor Green
            }
            Write-Host "    Attachments: $($extracted.Attachments.Count)" -ForegroundColor Green
        } else {
            Write-Host "    WARNING: Could not extract body from MIME" -ForegroundColor Red
        }

        if ($null -ne $headers -and $headers.Subject) {
            Write-Host "    Protected subject: $($headers.Subject)" -ForegroundColor Green
        }

        Write-Host "    RESULT: $(if ($wasSigned) {'SIGNED+'})ENCRYPTED - DECRYPT OK" -ForegroundColor Green
    } else {
        Write-Host "    DECRYPT FAILED: $($decResult.ErrorMessage)" -ForegroundColor Red
    }
}

Write-Host ""
Write-Host "=== E2E Test Complete ===" -ForegroundColor Cyan
