# Hyper-V Toolkit

## Version 1 • By Diobyte • Made with love

A focused PowerShell toolkit for creating Hyper-V virtual machines and managing GPU-P in one streamlined desktop workflow.

![Version](https://img.shields.io/badge/version-1-blue?style=for-the-badge)
![Platform](https://img.shields.io/badge/platform-Windows%2010%2F11-1f6feb?style=for-the-badge)
![PowerShell](https://img.shields.io/badge/powershell-64--bit-5391fe?style=for-the-badge)
![License](https://img.shields.io/badge/license-MIT-green?style=for-the-badge)

## Why this toolkit

Hyper-V setup can be repetitive, especially when combining VM creation with GPU partitioning tasks. This toolkit brings both flows into one interface so you can provision faster with fewer manual steps.

## Features

- Guided VM creation from ISO (disk, memory, CPU, network)
- Optional unattended install support
- GPU-P helper workflows for adapter allocation and driver copy
- Elevated launcher (`Launch.bat`) that ensures 64-bit Windows PowerShell and administrator privileges
- Startup/launcher logging and cleanup helpers

## Requirements

- Windows 10 or Windows 11 with Hyper-V support
- Hyper-V feature enabled
- Administrator privileges
- Windows PowerShell 5.1 (64-bit)
- Enough disk space for VHD(s) and installation media

Enable Hyper-V if needed:

```powershell
Enable-WindowsOptionalFeature -Online -FeatureName Microsoft-Hyper-V -All
```

Reboot if prompted.

## Project layout

- `HyperV-Toolkit.ps1` — Main WinForms toolkit script
- `Launch.bat` — Elevated 64-bit launcher
- `README.md` — Documentation
- `LICENSE` — MIT license

## Quick start (recommended)

1. Right-click `Launch.bat` and select **Run as administrator**.
2. Follow the toolkit UI to configure and create a VM.
3. Optionally run GPU-P workflows from the GPU tools section.

Notes:

- If launched from PowerShell 7 (`pwsh`), the script relaunches in Windows PowerShell 5.1.
- If downloaded from the internet, you may need:

```powershell
Unblock-File .\HyperV-Toolkit.ps1
```

## Running manually

From elevated Windows PowerShell in the repo folder:

```powershell
Set-ExecutionPolicy Bypass -Scope Process -Force
.\HyperV-Toolkit.ps1
```

Or simply run:

```powershell
.\Launch.bat
```

## Troubleshooting

- Confirm Hyper-V is enabled and `vmms` service is running.
- Run from an elevated session.
- Use Windows PowerShell 5.1 (64-bit).
- Verify ISO/path permissions.
- Ensure sufficient free disk space.

If a download or external helper step fails validation/signature checks, the step is skipped and logged.

## Contributing

Contributions are welcome.

1. Fork the repository.
2. Create a branch (for example: `fix/launcher-elevation`).
3. Open a PR with a short description and testing notes.

## License

MIT — see [LICENSE](LICENSE)
