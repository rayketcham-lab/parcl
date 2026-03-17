# Parcl

**S/MIME Encryption & Certificate Management for Microsoft Outlook**

[![CI](https://github.com/rayketcham-lab/parcl/actions/workflows/ci.yml/badge.svg)](https://github.com/rayketcham-lab/parcl/actions/workflows/ci.yml)
[![Security](https://github.com/rayketcham-lab/parcl/actions/workflows/security.yml/badge.svg)](https://github.com/rayketcham-lab/parcl/actions/workflows/security.yml)
[![License: Apache 2.0](https://img.shields.io/badge/License-Apache%202.0-blue.svg)](LICENSE)

Parcl is a Microsoft Outlook COM add-in that brings end-to-end S/MIME email encryption, digital signatures, and certificate management directly into the Outlook ribbon. Encrypted messages display inline in the reading pane — no `.p7m` attachments, no external tools, no server-side configuration.

**v0.9.3-beta** — Proof of Concept

[Homepage](https://rayketcham-lab.github.io/parcl/) | [Releases](https://github.com/rayketcham-lab/parcl/releases) | [Issues](https://github.com/rayketcham-lab/parcl/issues) | [quantumnexum.com](https://quantumnexum.com) | Support: help@quantumnexum.com

---

## Why Parcl?

- **Inline decryption** — Encrypted messages render directly in the Outlook reading pane. No separate viewer, no raw `.p7m` files.
- **Dual-mode encryption** — Native Outlook S/MIME via `PR_SECURITY_FLAGS` (compatible with any S/MIME client) and Parcl envelope mode with RFC 7508 protected headers for Parcl-to-Parcl communication.
- **Deferred encryption** — Toggle Encrypt while composing. Edit normally. Encryption happens transparently at send time.
- **Enterprise-ready** — RDN/CN-to-email matching handles enterprise certificate mismatches where the certificate CN does not match the email address.
- **Zero admin** — Per-user MSI installer. HKCU registry only. No administrator privileges required.

---

## Features

### Encryption & Signing
- **End-to-end S/MIME encryption** with inline reading pane display
- **Native Outlook S/MIME** via `PR_SECURITY_FLAGS` — interoperable with any RFC 5751 compliant S/MIME client
- **Parcl envelope mode** with RFC 7508 protected headers (Subject, From, To, Date encrypted inside the envelope)
- **Digital signatures** — sign-only, sign+encrypt, and opaque signing options
- **Sign-then-encrypt** layering per RFC 5751 Section 3.7 — signature protected inside the encrypted envelope
- **AES-256-CBC and AES-256-GCM** encryption; **SHA-256/384/512** signing
- **Send blocking** — if encryption is toggled on and fails, the message is blocked. No silent fallback.

### Certificate Management
- **Certificate management dialog** for external contacts — view, import, and remove certificates
- **Certificate exchange** — send your public certificate to contacts
- **RDN/CN-to-email matching** for enterprise certificates with name/email mismatches
- **Auto-decrypt** incoming encrypted messages
- **Inbox icons** for encrypted and signed messages (native Outlook S/MIME message classes)
- **Multi-source resolution** — Outlook contacts, Exchange GAL, MAPI properties, Windows cert stores (AddressBook, My) with SAN email matching
- **LDAP directory lookup** — optional, configurable; RFC 4515 injection prevention
- **Certificate cache** with TTL expiration and oldest-first eviction
- **Configurable validation** — None (expiry only), Relaxed (chain, no revocation), Strict (chain + OCSP/CRL)
- **BasicConstraints enforcement** — rejects CA certificates

### Security
- DPAPI credential protection
- Settings integrity via HMAC-SHA256 with DPAPI-managed key
- FIPS 140-2 mode detection and reporting (Windows CNG)
- RFC 4515 LDAP filter escaping
- RFC 2231 MIME parameter encoding for non-ASCII filenames
- RFC 8398 Unicode normalization for internationalized email addresses
- PII sanitization in logs, with CI scan enforcement
- No weak algorithms — 3DES and SHA-1 blocked per RFC 8551
- Randomized temp file paths; LDAPS by default (port 636)

### Logging & Diagnostics
- **Structured JSONL logging** with configurable log levels
- Session correlation IDs, component tags, PID tracking
- **PII sanitization** — email addresses and personal data scrubbed from log output
- 7-day automatic log retention
- GitHub issue integration — Report Issue / Suggest Feature buttons in About dialog

---

## Architecture

```
src/
  Parcl.Addin/           # COM Add-in (IDTExtensibility2 + IRibbonExtensibility)
    Dialogs/              #   WinForms dialogs (options, cert selector, cert exchange, about)
    TaskPane/             #   WPF security dashboard with ElementHost bridge
    ParclAddIn.cs         #   Add-in lifecycle, ItemSend encryption, event hooks
    ParclAddIn.Ribbon.cs  #   Ribbon callbacks, toggle logic, cert resolution
    ParclRibbon.xml       #   Ribbon + context menu XML
  Parcl.Core/            # Core library (zero Outlook dependency)
    Config/               #   ParclSettings, ParclLogger, CredentialProtector, SettingsIntegrity
    Crypto/               #   SmimeHandler (CMS), CertificateStore (X509), MimeBuilder, CertExchange
    Ldap/                 #   LdapCertLookup (DirectoryServices), CertificateCache
    Models/               #   CertificateInfo, LdapDirectoryEntry, UserProfile
  Parcl.Installer/       # WiX 6 per-user MSI (HKCU, no admin required)
tests/
  Parcl.Core.Tests/      # xUnit tests for the core library
```

**Stack**: .NET Framework 4.8 | COM Add-in (IDTExtensibility2) | WPF + WinForms | System.Security.Cryptography.Pkcs | System.DirectoryServices | WiX 6 | Newtonsoft.Json

---

## RFC Compliance

| RFC | Title | Status |
|-----|-------|--------|
| **5652** | Cryptographic Message Syntax (CMS) | Implemented |
| **5751** | S/MIME 3.2 Message Specification | Implemented |
| **7508** | Securing Header Fields with S/MIME | Implemented |
| **3370** | CMS Algorithms (AES-256-CBC, SHA-256) | Implemented |
| **5754** | Using SHA-2 with CMS | Implemented |
| **5280** | X.509 PKI Certificate Profile | Implemented |
| **4515** | LDAP Search Filter Escaping | Implemented |
| **4510-4519** | LDAP v3 Protocol Suite | Implemented |
| **2231** | Non-ASCII MIME Filename Encoding | Implemented |
| **8398** | Internationalized Email in X.509 | Implemented |
| **5083/5084** | AES-GCM Authenticated Encryption | Implemented |
| **8551** | S/MIME 4.0 (weak algo blocking) | Partial |

---

## Requirements

- Microsoft Outlook 2016 or later (desktop, 32-bit or 64-bit)
- .NET Framework 4.8
- Windows 10 or Windows 11

---

## Installation

Download the latest MSI from [Releases](https://github.com/rayketcham-lab/parcl/releases).

Per-user install — **no administrator privileges required**.

```
msiexec /i Parcl.Installer.msi
```

Silent install:
```
msiexec /i Parcl.Installer.msi /qn
```

---

## Building from Source

```powershell
dotnet restore Parcl.sln
dotnet build Parcl.sln --configuration Release
dotnet test tests/Parcl.Core.Tests/Parcl.Core.Tests.csproj --configuration Release
```

**Build requirements**: .NET SDK 8.0+, Windows (for .NET Framework 4.8 targeting and WiX).

---

## CI/CD

| Workflow | Purpose |
|----------|---------|
| **CI** | Build, test, lint, upload MSI artifact |
| **Security** | Dependency vulnerability scan, secret detection, code analysis |
| **PII Scan** | Verify no personally identifiable information leaks into logs or output |

---

## Configuration

Settings: `%APPDATA%\Parcl\settings.json` (DPAPI-encrypted credentials, HMAC integrity)

Logs: `%APPDATA%\Parcl\logs\parcl-YYYY-MM-DD.jsonl`

```powershell
# View all errors
Get-Content "$env:APPDATA\Parcl\logs\parcl-*.jsonl" | ConvertFrom-Json | Where-Object lvl -eq "ERROR"

# Filter by component
Get-Content "$env:APPDATA\Parcl\logs\parcl-*.jsonl" | ConvertFrom-Json | Where-Object cmp -eq "Encrypt"
```

---

## License

[Apache License 2.0](LICENSE)

## Credits

Built by [Quantum Nexum](https://quantumnexum.com).
