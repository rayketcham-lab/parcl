# Parcl

**S/MIME Certificate Manager & Encryption Add-in for Microsoft Outlook**

Parcl is a Microsoft Outlook add-in that provides S/MIME email encryption, decryption, and digital signing with integrated certificate management and LDAP directory lookup. Pure COM add-in — no VSTO runtime required.

[Homepage](https://rayketcham-lab.github.io/parcl/) · [Releases](https://github.com/rayketcham-lab/parcl/releases) · [Issues](https://github.com/rayketcham-lab/parcl/issues)

## Features

- **Encrypt** — Encrypt outgoing emails using recipient S/MIME certificates (AES-256-CBC)
- **Decrypt** — Decrypt received emails using your private key
- **Sign** — Digitally sign outgoing emails with SHA-256/RSA
- **Certificate Exchange** — Send your public certificate to recipients for secure communication
- **Certificate Selector** — Browse and choose signing/encryption certificates from the Windows certificate store
- **LDAP Lookup** — Automatically discover recipient certificates via directory services
- **Security Dashboard** — Animated WPF task pane showing certificate health, encryption state, and security status
- **About Dialog** — Version info, GitHub link, Quantum Nexum branding
- **Options** — Configure LDAP directories, certificate validation policy, and add-in behavior

## Security

- RFC 7508 header protection for S/MIME messages
- Configurable certificate validation (chain trust, revocation, expiry)
- Weak algorithms blocked — no MD5, no SHA-1 signatures
- Signing goes inside the encrypted envelope (encrypt-then-sign is not used)
- Send is blocked when encryption fails — no silent fallback to unencrypted
- Credentials protected via `CredentialProtector` (DPAPI)
- Structured JSONL audit logging

## Architecture

```
src/
  Parcl.Addin/           # Pure COM Add-in (IDTExtensibility2)
    Animations/          # WPF animated controls (lock, shield, spinner, badge)
    Dialogs/             # WinForms dialogs (cert selector, options, about, exchange)
    TaskPane/            # WPF security dashboard task pane
  Parcl.Core/            # Core library (zero Outlook dependency)
    Crypto/              # SmimeHandler, CertificateStore, MimeBuilder, CertExchange
    Ldap/                # LdapCertLookup, CertificateCache
    Config/              # ParclSettings, ParclLogger, CredentialProtector
    Models/              # CertificateInfo, LdapDirectoryEntry, UserProfile
  Parcl.Installer/       # WiX per-user installer (HKCU, no admin required)
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
