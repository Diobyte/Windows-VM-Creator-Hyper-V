# Hyper-V Toolkit

## Version 1 • By Diobyte • Made with love

A polished PowerShell toolkit for creating Hyper-V virtual machines and managing GPU-P in one streamlined desktop experience.

![Version](https://img.shields.io/badge/version-1-blue?style=for-the-badge)
![Platform](https://img.shields.io/badge/platform-Windows%2010%2F11-1f6feb?style=for-the-badge)
![PowerShell](https://img.shields.io/badge/powershell-64--bit-5391fe?style=for-the-badge)
![License](https://img.shields.io/badge/license-MIT-green?style=for-the-badge)

---

## Table of Contents

- [Why this toolkit](#why-this-toolkit)
- [Feature highlights](#feature-highlights)
- [Requirements](#requirements)
- [Project layout](#project-layout)
- [Quick start](#quick-start)
- [Workflow](#workflow)
- [Troubleshooting](#troubleshooting)
- [Version](#version)
- [License](#license)

## Why this toolkit

Hyper-V setup can be repetitive, especially when combining VM creation with GPU partitioning tasks. This toolkit brings both flows into one interface so you can provision faster with fewer manual steps.

## Feature highlights

- **VM Creation:** Creates Hyper-V VMs from ISO with guided configuration.
- **Storage + Networking:** Handles disk and network settings in the same flow.
- **Unattended Setup:** Supports unattended installation options.
- **GPU-P Management:** Configures GPU partition adapter and allocation controls.
- **Runtime Safety:** Includes cleanup and reliability helpers.
- **Launcher:** Auto-prompts for Administrator permissions from `Launch.bat`.

## Requirements

- Windows 10/11 with Hyper-V support
- Hyper-V feature enabled
- Administrator privileges
- PowerShell (64-bit)
- PowerShell execution policy that allows local scripts (the launcher uses `RemoteSigned`)

## Project layout

```text
.
├─ HyperV-Toolkit.ps1   # Main GUI toolkit (VM + GPU workflows)
├─ Launch.bat           # Elevated launcher entrypoint
├─ README.md            # Project documentation
└─ LICENSE              # MIT license
```

## Quick start

1. Clone or download this repository.
2. Run `Launch.bat`.
3. Accept elevation prompt.
4. Configure VM options and optional GPU-P settings.
5. Execute and monitor logs inside the app.

Launch compatibility note:

- The launcher now forces 64-bit Windows PowerShell for best Windows 10/11 compatibility.
- If you start the script from `pwsh` (PowerShell 7), it auto-relaunches itself in Windows PowerShell 5.1.

### Optional: start from PowerShell

```powershell
Set-Location .
.\Launch.bat
```

## Workflow

```mermaid
flowchart LR
  A[Open Launch.bat] --> B[Run as Administrator]
  B --> C[Configure VM]
  C --> D[Optional GPU-P setup]
  D --> E[Create / Apply]
  E --> F[Review Logs]
```

## Troubleshooting

- Confirm Hyper-V is enabled and running.
- Use 64-bit PowerShell only.
- Ensure ISO and destination paths are accessible.
- Keep enough free disk space for VHD and drivers.
- If launch is blocked by execution policy after downloading from the internet, run `Unblock-File .\HyperV-Toolkit.ps1` in PowerShell and retry.
- Optional post-install downloads now require valid Authenticode signatures and expected signer identity; if validation fails, that install step is skipped and logged.

## Version

- Current release: **Version 1**
- Maintainer and creator: **Diobyte**

## License

MIT — see [LICENSE](LICENSE).

