# Parcl

**Secure Email Certificate Manager for Microsoft Outlook**

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

## Requirements

- Microsoft Outlook 2016+ (desktop)
- .NET Framework 4.8
- Windows 10/11

## Installation

Per-user install — no administrator privileges required. Registers under `HKCU` to avoid WOW6432Node registry issues.

## Building

```powershell
dotnet restore
dotnet build --configuration Release
```

## License

MIT
