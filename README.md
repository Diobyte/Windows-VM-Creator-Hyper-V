# Hyper-V Toolkit

> **Version 1** • by **Diobyte** • made with love

A focused desktop toolkit for **Hyper-V VM creation** and **GPU-P management** in a single PowerShell + WinForms workflow.

![Version](https://img.shields.io/badge/version-1-1f6feb?style=for-the-badge)
![Windows](https://img.shields.io/badge/windows-10%20%7C%2011-0078D4?style=for-the-badge)
![PowerShell](https://img.shields.io/badge/PowerShell-5.1%20(64--bit)-5391FE?style=for-the-badge)
![Hyper-V](https://img.shields.io/badge/Hyper--V-required-6f42c1?style=for-the-badge)
![License](https://img.shields.io/badge/license-MIT-2ea043?style=for-the-badge)

## Table of contents

- [Why this exists](#why-this-exists)
- [Highlights](#highlights)
- [Requirements](#requirements)
- [Quick start](#quick-start)
- [Manual usage](#manual-usage)
- [CLI parameters](#cli-parameters)
- [Project structure](#project-structure)
- [Troubleshooting](#troubleshooting)
- [Contributing](#contributing)
- [License](#license)

## Why this exists

Configuring Hyper-V guests can involve many repetitive steps, especially when mixing VM provisioning and GPU partition workflows. This toolkit keeps those tasks in one place so setup is faster, cleaner, and less error-prone.

## Highlights

| Area | What it does |
| --- | --- |
| VM creation | Guided setup from ISO with disk, memory, CPU, and network options |
| Unattended install | Optional unattended setup flow |
| GPU-P tooling | VM selection, adapter allocation, and driver copy helper actions |
| Safe launcher | `Launch.bat` enforces elevation + 64-bit Windows PowerShell |
| Startup reliability | Logging, preflight checks, and startup validation guards |

## Requirements

- Windows 10/11 with Hyper-V-capable hardware
- Hyper-V enabled
- Administrator privileges
- Windows PowerShell 5.1 (64-bit)
- Free storage for VHDX and installation media

Enable Hyper-V if needed:

```powershell
Enable-WindowsOptionalFeature -Online -FeatureName Microsoft-Hyper-V-All -All
```

Restart Windows if prompted.

## Quick start

1. Right-click **Launch.bat**.
2. Choose **Run as administrator**.
3. Complete VM configuration in the UI.
4. Optionally use the GPU tab/tools for GPU-P steps.

### Notes

- If started from `pwsh`, the script relaunches itself in Windows PowerShell 5.1.
- If files came from the internet, unblock first:

```powershell
Unblock-File .\HyperV-Toolkit.ps1
```

## Manual usage

From an elevated Windows PowerShell session in this folder:

```powershell
Set-ExecutionPolicy Bypass -Scope Process -Force
.\HyperV-Toolkit.ps1
```

Or use the launcher directly:

```powershell
.\Launch.bat
```

## CLI parameters

Use these to prefill startup values or run validation mode:

```powershell
.\HyperV-Toolkit.ps1 -VMName "Win11Lab" -ISOPath "D:\ISO\Windows11.iso"
```

```powershell
.\HyperV-Toolkit.ps1 -WhatIf
```

| Parameter | Type | Purpose |
| --- | --- | --- |
| `-VMName` | string | Prefills VM name in the Create VM tab |
| `-ISOPath` | string | Prefills ISO path in the Create VM tab |
| `-WhatIf` | switch | Runs preflight/validation paths without applying changes |
| `-Headless` | switch | Reserved for future CLI-only mode (currently not implemented) |

## Project structure

- `HyperV-Toolkit.ps1` — main WinForms toolkit script
- `Launch.bat` — elevated 64-bit launcher
- `README.md` — documentation
- `LICENSE` — MIT license

## Troubleshooting

- Confirm Hyper-V is enabled and `vmms` is running.
- Ensure you are elevated (Administrator session/UAC approved).
- Use Windows PowerShell 5.1 64-bit (not x86).
- Verify ISO path permissions and existence.
- Confirm enough free disk space for VHD operations.

### GPU-P checklist

- Use **Generation 2** VMs for GPU-P.
- Keep VM **Dynamic Memory disabled** when assigning GPU-P.
- Ensure VM has GPU virtualization settings applied (`GuestControlledCacheTypes`, MMIO space).
- Verify host reports partitionable devices:

```powershell
Get-VMHostPartitionableGpu | Format-List Name,ValidPartitionCounts
```

- If guest shows Code 12 / insufficient resources, reduce GPU allocation and verify MMIO sizing.
- Keep host and guest GPU drivers aligned to vendor guidance; for NVIDIA vGPU stacks, ensure licensing/driver branch compatibility.

If helper downloads or external binaries fail signature/validation checks, the action is skipped and logged.

## Contributing

Contributions are welcome.

1. Fork the repo.
2. Create a feature/fix branch.
3. Open a PR with a short summary and test notes.

## License

MIT — see [LICENSE](LICENSE)
