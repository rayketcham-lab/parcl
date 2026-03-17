# Parcl — Outlook Certificate Manager Add-in

Outlook VSTO COM Add-in for S/MIME certificate management, encryption, signing, and LDAP lookup.

**Language:** C# / .NET Framework 4.8 | **Binary:** Parcl.Addin | **License:** Apache 2.0

---

## MANDATORY Rules

### Code Provenance

- **ALL code is original** — written from scratch
- **NEVER** copy code from GitHub, Stack Overflow, or other projects
- Dependencies (NuGet) are OK

### Architecture

```
src/
  Parcl.Addin/           # VSTO Add-in — Ribbon, dialogs, Outlook event handlers
  Parcl.Core/            # Core library — S/MIME, LDAP, cert store, models
  Parcl.Installer/       # WiX per-user installer (HKCU, no admin)
tests/
  Parcl.Core.Tests/      # xUnit tests for core library
```

### Key Design Decisions

- **Per-user install only** — all registry under HKCU, no HKLM, no WOW6432Node
- **VSTO COM Add-in** — deep Outlook integration via Microsoft.Office.Interop.Outlook
- **.NET Framework 4.8** — required by VSTO runtime
- **Ribbon XML** — custom tab with Encrypt, Decrypt, Sign, Certificate Exchange, Certificate Selector, Options buttons
- **LDAP via System.DirectoryServices** — no third-party LDAP dependencies

### Conventions

- Namespace: `Parcl.Addin`, `Parcl.Core`, `Parcl.Core.Tests`
- Use `async/await` for LDAP and crypto operations
- WinForms for dialogs (VSTO standard)
- Settings stored in user's AppData (`%APPDATA%\Parcl\`)
