# Parcl

**Secure Email Certificate Manager for Microsoft Outlook**

Parcl is a Microsoft Outlook add-in that provides S/MIME email encryption, decryption, and digital signing with integrated certificate management and LDAP directory lookup.

## Features

- **Encrypt** — Encrypt outgoing emails using recipient certificates
- **Decrypt** — Decrypt received emails using your private key
- **Sign** — Digitally sign outgoing emails
- **Certificate Exchange** — Send your certificate to recipients for secure communication
- **Certificate Selector** — Choose signing and encryption certificates from your local store
- **LDAP Lookup** — Automatically discover recipient certificates via directory services
- **Options** — Configure LDAP directories, crypto preferences, and add-in behavior

## Architecture

```
src/
  Parcl.Addin/       # Outlook VSTO COM Add-in (Ribbon UI, dialogs)
  Parcl.Core/        # Core library (S/MIME, LDAP, cert management)
  Parcl.Installer/   # Per-user WiX installer (no admin required)
tests/
  Parcl.Core.Tests/  # Unit and integration tests
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
