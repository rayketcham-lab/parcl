# Parcl

**S/MIME Certificate Manager & Encryption Add-in for Microsoft Outlook**

[![CI](https://github.com/rayketcham-lab/parcl/actions/workflows/ci.yml/badge.svg)](https://github.com/rayketcham-lab/parcl/actions/workflows/ci.yml)
[![Security](https://github.com/rayketcham-lab/parcl/actions/workflows/security.yml/badge.svg)](https://github.com/rayketcham-lab/parcl/actions/workflows/security.yml)
[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)

Parcl is a Microsoft Outlook COM add-in that provides end-to-end S/MIME email security — encryption, digital signatures, and certificate management — directly from the Outlook ribbon. Pure COM add-in — no VSTO runtime required. It handles the entire certificate lifecycle: discovery via LDAP and Outlook contacts, import from received emails, and application of RFC 5751 S/MIME cryptographic operations.

[Homepage](https://rayketcham-lab.github.io/parcl/) | [Releases](https://github.com/rayketcham-lab/parcl/releases) | [Issues](https://github.com/rayketcham-lab/parcl/issues) | [quantumnexum.com](https://quantumnexum.com) | Support: help@quantumnexum.com

---

## Why Parcl?

Most S/MIME solutions either require server-side configuration (Exchange S/MIME policies), force users to understand certificate management, or bolt encryption on as a visible `.p7m` attachment. Parcl does none of that:

- **True encapsulation** — Messages are encrypted into a proper CMS envelope per RFC 5751. The recipient's mail client handles it natively. No visible `.p7m` files.
- **Deferred encryption** — Click Encrypt while composing. Edit your message normally. Encryption happens transparently at send time.
- **Header protection** — Subject, From, To, Date are encrypted inside the envelope (RFC 7508). The outer Subject shows "Encrypted Message."
- **Sign-then-encrypt** — When both are active, the signature goes inside the encrypted envelope per RFC 5751 Section 3.7. Recipients see "Encrypted" not "Signed."
- **Zero-config cert discovery** — Finds recipient certificates from Outlook contacts, Exchange GAL, Windows cert stores, and LDAP directories. No manual cert management needed.
- **No admin required** — Per-user MSI installer. HKCU registry only.

---

## Features

### Encryption & Signing
- **Encrypt** — Toggle encryption on any compose message. S/MIME encapsulation (AES-256-CBC) happens at send time. Message stays fully editable.
- **Sign** — Toggle digital signature. Signs with your selected certificate (SHA-256 digest).
- **Sign + Encrypt** — Sign-then-encrypt layering per RFC 5751. Signature is protected inside the encrypted envelope.
- **Encrypt at Rest** — Lock received messages in your mailbox with your encryption certificate.
- **Send Blocking** — If encryption is toggled ON and fails (missing cert, expired cert, any error), the message is blocked from sending. No silent fallback.
- **Header Protection** — RFC 7508: Subject/From/To/Date inside the encrypted envelope. Outer subject replaced with "Encrypted Message."

### Certificate Management
- **Multi-source resolution** — Resolves recipient certs from: Outlook contacts > Exchange GAL > AddressEntry MAPI properties > Windows cert stores (AddressBook, My) with SAN email matching.
- **SMTP resolution** — Converts Exchange X500 internal addresses to SMTP automatically.
- **Import from email** — Manual "Import Certificates" button with per-cert user consent (subject/issuer/thumbprint displayed before import).
- **Certificate exchange** — Export your public cert as PEM or DER and attach to outgoing messages.
- **LDAP lookup** — Search LDAP directories for recipient certificates. RFC 4515 injection prevention.
- **Certificate cache** — JSON-based cache with TTL expiration keyed by email address.
- **Configurable validation** — None (expiry only), Relaxed (chain, no revocation), Strict (chain + OCSP/CRL).

### Security
- RFC 7508 header protection for S/MIME messages
- Configurable certificate validation (chain trust, revocation, expiry)
- Weak algorithms blocked — no 3DES, no SHA-1 (RFC 8551)
- Signing goes inside the encrypted envelope (sign-then-encrypt per RFC 5751)
- Send is blocked when encryption fails — no silent fallback to unencrypted
- Credentials protected via DPAPI (`CredentialProtector`)
- LDAP injection prevention (RFC 4515 escaping)
- Randomized temp file paths
- MIME header sanitization
- LDAPS by default (port 636, SSL on)
- No auto-import — every certificate import requires explicit user consent
- Structured JSONL audit logging with session correlation

### User Interface
- **Ribbon toggle buttons** — Encrypt and Sign show pressed/highlighted state reflecting the current message. State updates when switching between messages.
- **Context menus** — Right-click: Encrypt, Sign, Remove Encryption, Remove Signature, Send Certificate (compose). Encrypt (Lock), Decrypt (read).
- **Animated task pane** — WPF dashboard with security status, cert info, quick actions, LDAP lookup with spinner and results.
- **Options dialog** — LDAP directories, crypto preferences, validation mode, cache settings, behavior toggles.
- **About dialog** — Version, GitHub link, quantumnexum.com, support email.

### Logging
- **JSONL structured logging** — One JSON object per line: `ts`, `lvl`, `cmp`, `sid` (session ID), `pid`, `msg`, `err`.
- **Filter with PowerShell**: `Get-Content parcl.jsonl | ConvertFrom-Json | Where-Object cmp -eq "LDAP"`
- **Filter with jq**: `jq 'select(.lvl=="ERROR")' parcl.jsonl`
- **Excel**: Data > Get Data > From JSON > select the `.jsonl` file.
- **7-day retention** with automatic cleanup.

---

## RFC Compliance

### Implemented
| RFC | Title | Coverage |
|-----|-------|----------|
| **5652** | Cryptographic Message Syntax (CMS) | EnvelopedCms, SignedCms, ContentInfo |
| **5751** | S/MIME 3.2 Message Specification | Sign-then-encrypt, application/pkcs7-mime, IPM.Note.SMIME |
| **7508** | Securing Header Fields with S/MIME | Subject/From/To/Date inside encrypted envelope |
| **3370** | CMS Algorithms | AES-256-CBC, SHA-256 |
| **5754** | Using SHA-2 with CMS | SHA-256 digest for signatures |
| **5280** | X.509 PKI Certificate Profile | Chain building, expiry, SAN parsing, configurable revocation |
| **4515** | LDAP Search Filter Escaping | RFC 4515 special character escaping |
| **4510-4519** | LDAP v3 Protocol Suite | Search, bind auth, SSL/TLS |
| **4523** | X.509 Certificates in LDAP | userCertificate;binary attribute |
| **2045-2049** | MIME | Multipart/mixed, base64 transfer encoding |
| **2183** | Content-Disposition | Attachment headers |

### Planned
| RFC | Title | Issue |
|-----|-------|-------|
| 5083/5084 | AES-GCM Authenticated Encryption | [#33](https://github.com/rayketcham-lab/parcl/issues/33) |
| 8551 | S/MIME 4.0 Full Compliance | [#34](https://github.com/rayketcham-lab/parcl/issues/34) |
| 6211 | CMS Algorithm Attribute Protection | [#37](https://github.com/rayketcham-lab/parcl/issues/37) |
| 2231 | Non-ASCII MIME Filename Encoding | [#35](https://github.com/rayketcham-lab/parcl/issues/35) |
| 8398 | Internationalized Email in X.509 | [#39](https://github.com/rayketcham-lab/parcl/issues/39) |
| 7030 | EST Certificate Enrollment | [#38](https://github.com/rayketcham-lab/parcl/issues/38) |

---

## Architecture

```
src/
  Parcl.Addin/           # Pure COM Add-in (IDTExtensibility2 + IRibbonExtensibility)
    Animations/           #   WPF animated controls (lock, shield, badge, spinner)
    Dialogs/              #   WinForms dialogs (options, cert selector, cert exchange, about)
    TaskPane/             #   WPF security dashboard with ElementHost bridge
    ParclAddIn.cs         #   Add-in lifecycle, ItemSend encryption, event hooks
    ParclAddIn.Ribbon.cs  #   Ribbon callbacks, toggle logic, cert resolution
    ParclRibbon.xml       #   Ribbon + context menu XML
    OfficeInterop.cs      #   Inline COM interface definitions (IRibbonUI, IRibbonControl)
  Parcl.Core/            # Core library (zero Outlook dependency)
    Config/               #   ParclSettings (JSON), ParclLogger (JSONL), CredentialProtector (DPAPI)
    Crypto/               #   SmimeHandler (CMS), CertificateStore (X509), MimeBuilder, CertExchange
    Ldap/                 #   LdapCertLookup (DirectoryServices), CertificateCache
    Models/               #   CertificateInfo, LdapDirectoryEntry, UserProfile
  Parcl.Installer/       # WiX 6 per-user MSI (HKCU, no admin required)
tests/
  Parcl.Core.Tests/      # xUnit tests for the core library
```

**Stack**: .NET Framework 4.8 | COM Add-in (IDTExtensibility2) | WPF + WinForms | System.Security.Cryptography.Pkcs | System.DirectoryServices | WiX 6 | Newtonsoft.Json

---

## Requirements

- Microsoft Outlook 2016 or later (desktop, 32-bit or 64-bit)
- .NET Framework 4.8
- Windows 10 or Windows 11

---

## Installation

Download the latest MSI from [Releases](https://github.com/rayketcham-lab/parcl/releases).

Per-user install — **no administrator privileges required**. All registry entries are under `HKCU`. No `HKLM`, no `WOW6432Node`.

```
msiexec /i Parcl.Installer.msi
```

Silent install:
```
msiexec /i Parcl.Installer.msi /qn
```

Uninstall:
```
msiexec /x Parcl.Installer.msi /qn
```

---

## Building from Source

```powershell
# Restore, build, test
dotnet restore Parcl.sln
dotnet build Parcl.sln --configuration Release
dotnet test tests/Parcl.Core.Tests/Parcl.Core.Tests.csproj --configuration Release

# MSI output at:
# src/Parcl.Installer/bin/Release/Parcl.Installer.msi
```

**Build requirements**: .NET SDK 8.0+, Windows (for .NET Framework 4.8 targeting and WiX).

---

## CI/CD

| Workflow | What it does |
|----------|-------------|
| **CI** | Build, test, lint, upload MSI artifact |
| **Security** | Dependency vulnerability scan, secret detection, code analysis |

---

## Configuration

Settings stored at `%APPDATA%\Parcl\settings.json` (DPAPI-encrypted credentials).

Logs stored at `%APPDATA%\Parcl\logs\parcl-YYYY-MM-DD.jsonl`.

Quick log analysis:
```powershell
# All errors
Get-Content "$env:APPDATA\Parcl\logs\parcl-*.jsonl" | ConvertFrom-Json | Where-Object lvl -eq "ERROR" | Format-Table ts, cmp, msg

# Filter by component
Get-Content "$env:APPDATA\Parcl\logs\parcl-*.jsonl" | ConvertFrom-Json | Where-Object cmp -eq "Encrypt" | Format-Table

# Filter by session
Get-Content "$env:APPDATA\Parcl\logs\parcl-*.jsonl" | ConvertFrom-Json | Where-Object sid -eq "abc123" | Format-Table
```

---

## Releases

| Version | Date | Highlights |
|---------|------|------------|
| [v1.4.0](https://github.com/rayketcham-lab/parcl/releases/tag/v1.4.0) | 2026-03-16 | RFC 7508 header protection, configurable cert validation, weak algo removal |
| [v1.3.1](https://github.com/rayketcham-lab/parcl/releases/tag/v1.3.1) | 2026-03-16 | Fix sign-then-encrypt layering, send blocking, About dialog |
| [v1.3.0](https://github.com/rayketcham-lab/parcl/releases/tag/v1.3.0) | 2026-03-16 | Security hardening (LDAP injection, DPAPI, chain validation), security CI |
| [v1.2.0](https://github.com/rayketcham-lab/parcl/releases/tag/v1.2.0) | 2026-03-16 | Windows build, S/MIME encapsulation, JSONL logging, multi-source cert resolution |

---

## License

[MIT](LICENSE)

## Credits

Built by [Quantum Nexum](https://quantumnexum.com).
