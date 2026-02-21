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

# Hyper-V Toolkit

Version 1 — by Diobyte

A small, focused PowerShell toolkit to simplify Hyper-V VM creation and optional GPU‑P configuration on Windows 10/11.

![Version](https://img.shields.io/badge/version-1-blue?style=for-the-badge) ![Platform](https://img.shields.io/badge/platform-Windows%2010%2F11-1f6feb?style=for-the-badge) ![PowerShell](https://img.shields.io/badge/powershell-64--bit-5391fe?style=for-the-badge) ![License](https://img.shields.io/badge/license-MIT-green?style=for-the-badge)

--

**Contents**

- Why this toolkit
- Features
- Prerequisites
- Quick start (recommended)
- Running manually
- Troubleshooting
- Contributing
- License

## Why this toolkit

Creating Hyper-V VMs and preparing GPU‑partitioning (GPU‑P) often involves many repetitive steps. This toolkit provides a guided PowerShell flow and an elevated launcher to speed up VM provisioning while keeping common safety checks and cleanup helpers.

## Features

- Guided VM creation from ISO (disk, memory, CPU, network)
- Optional unattended install support
- GPU‑P helper workflows for GPU partitioning and adapter allocation
- Elevated launcher (`Launch.bat`) that ensures 64‑bit Windows PowerShell and administrator privileges
- Basic logging and cleanup helpers

## Prerequisites

- Windows 10 or Windows 11 with Hyper‑V support
- Hyper‑V feature enabled
- Administrator privileges (required to create VMs)
- PowerShell (Windows PowerShell 5.1 recommended for the launcher)
- Enough disk space for VHD(s) and installation media

How to enable Hyper‑V (if not already):

```powershell
Enable-WindowsOptionalFeature -Online -FeatureName Microsoft-Hyper-V -All
```

Reboot if prompted.

## Quick start (recommended)

1. Open the repository folder in File Explorer.
2. Right‑click `Launch.bat` and choose **Run as administrator** (or double‑click and accept the UAC prompt).
3. Follow the on-screen prompts in the toolkit to configure and create a VM. Optionally enable GPU‑P workflows when prompted.

Notes:
- The launcher forces the toolkit to run in 64‑bit Windows PowerShell for compatibility. If you start from PowerShell 7 (`pwsh`), the launcher will relaunch the script in Windows PowerShell 5.1.
- If you downloaded the repo from the internet, you may need to unblock the script once: `Unblock-File .\HyperV-Toolkit.ps1`.

## Running manually (PowerShell)

Open an elevated Windows PowerShell (Run as Administrator) in the repo folder and run:

```powershell
Set-ExecutionPolicy Bypass -Scope Process -Force
.\HyperV-Toolkit.ps1
```

Or use the launcher which handles elevation and execution policy automatically:

```powershell
.\Launch.bat
```

## Troubleshooting

- Ensure Hyper‑V is enabled and the Hyper‑V services are running.
- Run the toolkit from an elevated (Administrator) PowerShell session.
- Use Windows PowerShell 5.1 (64‑bit) for best compatibility with some Hyper‑V and WMI components.
- If an ISO or path is inaccessible, verify permissions and UNC/drive mappings.
- Low disk space can cause VHD creation failures — free up space or choose a different disk.
- If a step that downloads signed installers fails signature checks, the step will be skipped and logged.

If you hit an issue, collect the script log and open an issue with steps to reproduce.

## Code reference (what the script actually does)

This section maps high-level features to the important script behaviors and functions so you know where to look and what to expect.

- **Startup & host checks**: The script enforces Windows PowerShell (Desktop) and a 64‑bit process, requests elevation, and sets a friendly process execution policy. See the top of `HyperV-Toolkit.ps1` for the checks and `Write-StartupTrace`/`Write-Log` usage.
- **Hyper‑V detection & enablement**: The toolkit checks Hyper‑V via `Get-WindowsOptionalFeature` and the `vmms` service. If Hyper‑V is missing it offers to enable it (`Enable-WindowsOptionalFeature`) and can prompt for a reboot.
- **Unattended installs**: `New-UnattendXml` generates unattended install XML for ISO-based deployments and encodes passwords via `ConvertTo-UnattendPassword`.
- **ISO & image handling**: The toolkit mounts/dismounts ISOs and tracks them (cleanup helpers such as `Invoke-MountCleanup`, `Dismount-ImageRetry`).
- **VM creation UI**: The main GUI exposes an "ISO Deploy" mode plus fields for VM name, disk size, RAM, CPU, network, and VM location. Defaults come from `Get-VMHost` (VirtualMachinePath) or `C:\HyperV`.
- **Networking**: If needed the toolkit creates an internal NAT switch named `HyperV-Toolkit-NAT` (see the VMSwitch creation block).
- **GPU‑P workflows**: Functions include `Test-GpuPPreFlight`, `Test-GpuPHostReadiness`, `Get-GpuPProviders`, `Set-GpuPartitionForVM`, `Get-GpuPartitionValues`, `Copy-GpuDriverFolders`, and `Copy-GpuServiceDriver`. Notes:
	- The toolkit warns about laptop NVIDIA GPUs (commonly unsupported for GPU‑P) and AMD Polaris limitations.
	- On Windows 10 hosts GPU selection uses `AUTO` (default GPU) — explicit GPU selection requires Windows 11.
	- Driver copying is a "smart copy" (only copies GPU-relevant DriverStore folders) to minimize size and complexity.
- **Logging & guidance**: Runtime logs are written to the GUI log box; startup traces go to `%TEMP%\HyperV-Toolkit-Startup.log`. The launcher logs to `%TEMP%\HyperV-Toolkit-Launcher.log`.

## Launcher details & advanced notes

- `Launch.bat` ensures the toolkit runs elevated in 64‑bit Windows PowerShell and uses `-ExecutionPolicy Bypass` when starting the script. It prefers the Sysnative path when launched from 32‑bit hosts so the toolkit runs in a 64‑bit process.
- The launcher accepts an internal `--elevated` flag when relaunching itself; you don't need to pass this manually.
- If you prefer to run the script directly, start an elevated Windows PowerShell and run:

```powershell
Set-ExecutionPolicy Bypass -Scope Process -Force
.\HyperV-Toolkit.ps1
```

## Defaults and common values

- Default VM location: `Get-VMHost`.VirtualMachinePath or `C:\HyperV` if that call fails.
- Default NAT switch name (when created): `HyperV-Toolkit-NAT`.
- Unattend guest architecture defaults to the host architecture (amd64/arm64) unless overridden.

## Contributing

Contributions are welcome. For small fixes or documentation changes, open a pull request. For code changes, please:

1. Fork the repository.
2. Create a branch with a clear name (e.g., `fix/launcher-elevation`).
3. Submit a PR with a short description of the change and testing notes.

## License

MIT — see [LICENSE](LICENSE)

