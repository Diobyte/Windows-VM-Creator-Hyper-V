#Requires -Version 5.1

[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost', '', Justification='Colored startup/status console output is intentional for interactive desktop use.')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification='This is a WinForms application script, not exported cmdlets; ShouldProcess semantics are not user-facing here.')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseSingularNouns', '', Justification='Function names are retained for backward compatibility with existing script usage and internal call sites.')]

[CmdletBinding()]
param(
    [string]$VMName,
    [string]$ISOPath,
    [switch]$Headless,
    [switch]$WhatIf
)

$script:CliVMName = $VMName
$script:CliISOPath = $ISOPath
$script:CliHeadless = [bool]$Headless
$script:CliWhatIf = [bool]$WhatIf

# Preserve explicitly passed startup arguments when we relaunch/elevate.
$script:ForwardedCliArgs = @()
if ($PSBoundParameters.ContainsKey('VMName') -and -not [string]::IsNullOrWhiteSpace($VMName)) {
    $script:ForwardedCliArgs += @('-VMName', $VMName)
}
if ($PSBoundParameters.ContainsKey('ISOPath') -and -not [string]::IsNullOrWhiteSpace($ISOPath)) {
    $script:ForwardedCliArgs += @('-ISOPath', $ISOPath)
}
if ($PSBoundParameters.ContainsKey('Headless') -and $Headless) {
    $script:ForwardedCliArgs += '-Headless'
}
if ($PSBoundParameters.ContainsKey('WhatIf') -and $WhatIf) {
    $script:ForwardedCliArgs += '-WhatIf'
}

if ($script:CliHeadless) {
    # TODO: Headless (CLI-only) mode is planned for a future release.
    # When implemented, this block will delegate to a non-GUI code path.
    Write-Host "Headless mode is not yet implemented in this version of HyperV-Toolkit." -ForegroundColor Yellow
    Write-Host "Use GUI mode, or remove -Headless and run interactively." -ForegroundColor Yellow
    exit 1
}

# Use strict mode to catch variable issues
Set-StrictMode -Version Latest

#region ==================== INITIALIZATION ====================

$script:ToolkitVersion = "Version 1"
$script:ToolkitCreator = "Diobyte"
$script:ToolkitTagline = "Made with love"

# Startup banner (console feedback while GUI loads)
Write-Host ""
Write-Host "  =============================================" -ForegroundColor Cyan
Write-Host "   Hyper-V Toolkit | $script:ToolkitVersion | $script:ToolkitCreator" -ForegroundColor Cyan
Write-Host "   $script:ToolkitTagline" -ForegroundColor DarkCyan
Write-Host "  =============================================" -ForegroundColor Cyan
Write-Host ""

# Well-known Windows build numbers
$script:BUILD_WIN10_RS4    = 17134   # Windows 10 1803 (earliest supported for modern Secure Boot)
$script:BUILD_WIN10_RS5    = 17763   # Windows 10 1809 (Secure Boot recommended)
$script:BUILD_WIN10_MIN    = 10240   # Windows 10 RTM
$script:BUILD_WIN11_MIN    = 22000   # Windows 11 21H2
$script:HyperVFeatureName  = 'Microsoft-Hyper-V-All'
$script:EmbeddedQResSha256 = '1c21d5dea9e1ef96c00c829114fe366ed4f23be3b6ce3df89ca8125b700e5945'
$script:StartupLogPath     = Join-Path $env:TEMP 'HyperV-Toolkit-Startup.log'

function Write-StartupTrace {
    param(
        [string]$Message,
        [ValidateSet('INFO','WARN','ERROR')]
        [string]$Level = 'INFO'
    )

    try {
        Add-Content -Path $script:StartupLogPath -Value "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss.fff')] [$Level] $Message" -Encoding UTF8 -ErrorAction SilentlyContinue
    } catch {
        # Intentionally suppress startup log I/O failures; avoid recursive logging.
        [void]$PSItem
    }
}

Write-StartupTrace -Message "Startup begin | PSEdition=$($PSVersionTable.PSEdition) | PSVersion=$($PSVersionTable.PSVersion) | Is64BitProcess=$([Environment]::Is64BitProcess) | PSHome=$PSHOME"

# Rotate startup log to prevent unbounded growth across many runs
try {
    if (Test-Path $script:StartupLogPath) {
        $logSize = (Get-Item $script:StartupLogPath -ErrorAction SilentlyContinue).Length
        if ($logSize -gt 512KB) {
            $lines = Get-Content $script:StartupLogPath -Tail 200 -ErrorAction SilentlyContinue
            $lines | Set-Content $script:StartupLogPath -Encoding UTF8 -ErrorAction SilentlyContinue
        }
    }
} catch {
    Write-StartupTrace -Message "Startup log rotation failed: $($_.Exception.Message)" -Level 'WARN'
}

# Ensure script runs in Windows PowerShell (Desktop) for WinForms/WPF compatibility
if ($PSVersionTable.PSEdition -ne 'Desktop') {
    $scriptPath = if ($PSCommandPath) { $PSCommandPath } else { $MyInvocation.MyCommand.Path }
    $winPsCandidates = @(
        "$env:WINDIR\Sysnative\WindowsPowerShell\v1.0\powershell.exe",
        "$env:WINDIR\System32\WindowsPowerShell\v1.0\powershell.exe",
        "$PSHOME\powershell.exe"
    )
    $winPsExe = $winPsCandidates | Where-Object { -not [string]::IsNullOrWhiteSpace($_) -and (Test-Path $_) } | Select-Object -First 1

    if ($winPsExe -and -not [string]::IsNullOrWhiteSpace($scriptPath) -and (Test-Path $scriptPath)) {
        try {
            Write-StartupTrace -Message "Non-Desktop host detected. Relaunching with Windows PowerShell: $winPsExe"
            $relaunchArgs = @(
                '-NoProfile',
                '-ExecutionPolicy', 'Bypass',
                '-Sta',
                '-File', $scriptPath
            )
            if ($script:ForwardedCliArgs -and $script:ForwardedCliArgs.Count -gt 0) {
                $relaunchArgs += $script:ForwardedCliArgs
            }
            Start-Process -FilePath $winPsExe -ArgumentList $relaunchArgs | Out-Null
            Write-StartupTrace -Message "Relaunch command sent successfully"
            exit 0
        } catch {
            Write-StartupTrace -Message "Relaunch failed: $($_.Exception.Message)" -Level 'ERROR'
            Write-Host "Failed to relaunch in Windows PowerShell: $($_.Exception.Message)" -ForegroundColor Red
            exit 1
        }
    }

    Write-StartupTrace -Message "Windows PowerShell 5.1 executable not found for relaunch" -Level 'ERROR'
    Write-Host "Windows PowerShell 5.1 is required to run this toolkit." -ForegroundColor Red
    exit 1
}

# Execution Policy
Write-Host "  [1/7] Configuring execution policy..." -ForegroundColor DarkGray
$currentPolicy = Get-ExecutionPolicy -Scope Process
if (-not $currentPolicy -or $currentPolicy -in @('Undefined', 'Restricted')) {
    try {
        Set-ExecutionPolicy -Scope Process -ExecutionPolicy RemoteSigned -Force -ErrorAction Stop
    } catch {
        Write-Output "Could not set process execution policy to RemoteSigned. Continuing with current policy: $currentPolicy"
    }
}

# Load assemblies
Write-Host "  [2/7] Loading assemblies..." -ForegroundColor DarkGray
try {
    Add-Type -AssemblyName PresentationFramework
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    [System.Windows.Forms.Application]::EnableVisualStyles()
} catch {
    Write-StartupTrace -Message "Assembly load failed: $($_.Exception.Message)" -Level 'ERROR'
    Write-Host "  FATAL: Failed to load .NET assemblies: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "  Ensure .NET Framework 4.x or later is installed." -ForegroundColor Red
    Write-Host "  Press Enter to exit..." -ForegroundColor Red
    Read-Host
    exit 1
}

# 64-bit process check
Write-Host "  [3/7] Checking environment..." -ForegroundColor DarkGray
if (-not [Environment]::Is64BitProcess) {
    Write-StartupTrace -Message "64-bit check failed: running in 32-bit process" -Level 'ERROR'
    [System.Windows.MessageBox]::Show(
        "This tool requires 64-bit PowerShell.`n`nDo not use PowerShell (x86). Use the standard PowerShell or the Launch.bat file.",
        "64-bit Required", "OK", "Error"
    ) | Out-Null
    exit 1
}

# Admin check
Write-Host "  [4/7] Verifying administrator privileges..." -ForegroundColor DarkGray
$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    Write-StartupTrace -Message "Process is not elevated; requesting admin relaunch"
    $scriptPath = if ($PSCommandPath) { $PSCommandPath } else { $MyInvocation.MyCommand.Path }
    if (-not [string]::IsNullOrWhiteSpace($scriptPath) -and (Test-Path $scriptPath)) {
        try {
            $elevatedPowerShell = Join-Path $PSHOME 'powershell.exe'
            if (-not (Test-Path $elevatedPowerShell)) { $elevatedPowerShell = 'PowerShell.exe' }
            Write-StartupTrace -Message "Elevating with executable: $elevatedPowerShell"
            $elevationArgs = @(
                '-NoProfile',
                '-ExecutionPolicy', 'RemoteSigned',
                '-Sta',
                '-File', $scriptPath
            )
            if ($script:ForwardedCliArgs -and $script:ForwardedCliArgs.Count -gt 0) {
                $elevationArgs += $script:ForwardedCliArgs
            }
            Start-Process -FilePath $elevatedPowerShell -Verb RunAs -ArgumentList $elevationArgs | Out-Null
            Write-StartupTrace -Message "Elevation command sent successfully"
            exit 0
        } catch {
            Write-StartupTrace -Message "Elevation failed or canceled: $($_.Exception.Message)" -Level 'ERROR'
            [System.Windows.MessageBox]::Show(
                "Administrator elevation was cancelled or failed.`n`nPlease right-click Launch.bat and select 'Run as Administrator'.",
                "Administrator Required", "OK", "Warning"
            ) | Out-Null
            exit 1
        }
    }

    [System.Windows.MessageBox]::Show(
        "This tool must be run as Administrator.`n`nPlease right-click and select 'Run as Administrator', or use the Launch.bat file.",
        "Administrator Required", "OK", "Warning"
    ) | Out-Null
    exit 1
}

# Hyper-V check
Write-Host "  [5/7] Checking Hyper-V status (may take a moment)..." -ForegroundColor DarkGray
function Test-HyperVRunning {
    try {
        $svc = Get-Service -Name vmms -ErrorAction Stop
        return ($svc.Status -eq 'Running')
    } catch { return $false }
}

$feature = $null
$featureJob = $null
try {
    $featureNameForJob = $script:HyperVFeatureName
    $featureJob = Start-Job -ScriptBlock {
        Get-WindowsOptionalFeature -Online -FeatureName $Using:featureNameForJob -ErrorAction SilentlyContinue
    }
    if (Wait-Job $featureJob -Timeout 30) {
        $feature = Receive-Job $featureJob -ErrorAction SilentlyContinue
    } else {
        Write-Host "    Hyper-V query is taking longer than expected..." -ForegroundColor Yellow
        if (Wait-Job $featureJob -Timeout 60) {
            $feature = Receive-Job $featureJob -ErrorAction SilentlyContinue
        } else {
            Write-Host "    Hyper-V query timed out after 90s." -ForegroundColor Yellow
            Stop-Job $featureJob -ErrorAction SilentlyContinue
            if (Test-HyperVRunning) {
                Write-Host "    vmms service is running - proceeding." -ForegroundColor Green
                $feature = [PSCustomObject]@{ State = "Enabled" }
            }
        }
    }
} catch {
    Write-StartupTrace -Message "Hyper-V async check failed; attempting direct query" -Level 'WARN'
    Write-Host "    Async check failed, trying direct query..." -ForegroundColor Yellow
    try {
        $feature = Get-WindowsOptionalFeature -Online -FeatureName $script:HyperVFeatureName -ErrorAction SilentlyContinue
    } catch {
        Write-StartupTrace -Message "Hyper-V direct query failed: $($_.Exception.Message)" -Level 'ERROR'
        Write-Host "    Hyper-V check failed: $($_.Exception.Message)" -ForegroundColor Red
    }
} finally {
    if ($featureJob) { Remove-Job $featureJob -Force -ErrorAction SilentlyContinue }
}

# Handle pending-reboot state (Hyper-V enabled but restart not yet done)
if ($feature -and $feature.State -eq "EnablePending") {
    [System.Windows.MessageBox]::Show(
        "Hyper-V has been enabled but requires a system restart.`n`nPlease restart your computer and then run the toolkit again.",
        "Restart Required", "OK", "Warning"
    ) | Out-Null
    exit 0
}

if (-not ($feature -and $feature.State -eq "Enabled" -and (Test-HyperVRunning))) {
    $skipInstall = $false
    # If Hyper-V feature is enabled but vmms service is simply stopped, try starting it first
    if ($feature -and $feature.State -eq "Enabled" -and -not (Test-HyperVRunning)) {
        Write-Host "    Hyper-V is enabled but vmms service is not running. Attempting to start..." -ForegroundColor Yellow
        try {
            Start-Service vmms -ErrorAction Stop
            Start-Sleep -Seconds 3
            if (Test-HyperVRunning) {
                Write-Host "    vmms service started successfully." -ForegroundColor Green
                $skipInstall = $true
            }
        } catch {
            Write-Host "    Could not start vmms service: $($_.Exception.Message)" -ForegroundColor Yellow
        }
    }
    if (-not $skipInstall) {
        $installChoice = [System.Windows.MessageBox]::Show(
            "Hyper-V is not fully enabled or the hypervisor is not running.`n`nA system restart will be required after installation.`n`nDo you want to enable it now?",
            "Enable Hyper-V", "OKCancel", "Warning"
        )
        if ($installChoice -eq "OK") {
            try {
                Enable-WindowsOptionalFeature -Online -FeatureName $script:HyperVFeatureName -All -NoRestart -ErrorAction Stop *> $null
                bcdedit /set hypervisorlaunchtype auto *> $null
                if ($null -ne $LASTEXITCODE -and $LASTEXITCODE -ne 0) {
                    Write-Host "    Warning: bcdedit returned exit code $LASTEXITCODE" -ForegroundColor Yellow
                }
                $restartChoice = [System.Windows.MessageBox]::Show(
                    "Hyper-V has been enabled successfully.`n`nDo you want to restart now?",
                    "Restart Required", "OKCancel", "Question"
                )
                if ($restartChoice -eq "OK") {
                    try {
                        Restart-Computer -ErrorAction Stop
                    } catch {
                        [System.Windows.MessageBox]::Show(
                            "Automatic restart failed. Please restart your PC manually, then run the toolkit again.",
                            "Manual Restart Required", "OK", "Warning"
                        ) | Out-Null
                        exit 0
                    }
                }
                else { Write-Output "Please restart your PC and run the script again."; exit 0 }
            } catch {
                [System.Windows.MessageBox]::Show("Failed to enable Hyper-V.`n`nError: $_", "Error", "OK", "Error") | Out-Null
                exit 1
            }
        } else { exit 0 }
    } # end if (-not $skipInstall)
}

Write-Host "  [6/7] Detecting host configuration..." -ForegroundColor DarkGray
# Detect host OS and hardware (cache CIM results to avoid redundant queries during GUI setup)
$script:HostOsName = (Get-CimInstance Win32_OperatingSystem).Caption
$script:HostComputerSystem = Get-CimInstance Win32_ComputerSystem
$script:HostBuild  = 0
try {
    $rawBuild = (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion').CurrentBuild
    if (-not [int]::TryParse($rawBuild, [ref]$script:HostBuild)) {
        Write-StartupTrace -Message "Non-numeric CurrentBuild value: '$rawBuild'; defaulting to 0" -Level 'WARN'
        $script:HostBuild = 0
    }
} catch {
    Write-StartupTrace -Message "Failed to read CurrentBuild from registry: $($_.Exception.Message)" -Level 'WARN'
}
$script:HostIsWin11 = ($script:HostBuild -ge $script:BUILD_WIN11_MIN)
$script:HostIsWin11Pro = $script:HostOsName -match 'Windows 11.*(Pro|Enterprise|Education)'
$script:HostArch = if ($env:PROCESSOR_ARCHITECTURE -eq 'ARM64') { 'arm64' } else { 'amd64' }
$script:VideoControllers = @(Get-CimInstance Win32_VideoController -ErrorAction SilentlyContinue)
try {
    $addGpuCmd = Get-Command Add-VMGpuPartitionAdapter -ErrorAction SilentlyContinue
    $script:SupportsGpuInstancePath = ($addGpuCmd -and $addGpuCmd.Parameters.ContainsKey('InstancePath'))
} catch {
    $script:SupportsGpuInstancePath = $false
}

#endregion

#region ==================== GLOBALS ====================

$script:MountedISO          = $null
$script:TrackedMountedImages = @{}
$script:WimFile             = $null
$script:EditionMap          = @{}
$script:DetectedWinVersion  = ""      # "Windows 10" or "Windows 11"
$script:DetectedBuild       = 0       # e.g. 19045, 22621
$script:DetectedGuestArch   = $script:HostArch  # guest architecture for unattend (amd64/arm64)
$script:LogBox              = $null   # Set when GUI is built
$script:IsCreating          = $false  # Re-entrancy guard for VM creation
$script:IsUpdatingGPU       = $false  # Re-entrancy guard for GPU update
$script:SuppressMemEvents   = $false  # Suppress cascading Dynamic Memory ValueChanged events
$script:GpuSelectedVMs      = @{}     # Persist selected VMs across filter/refresh
$script:SuspendGpuSelectionEvents = $false # Batch-select guard to avoid event storms
$script:DoEventsWarningLogged = $false # One-time guard for DoEvents warning logging
$script:LastLogRefresh      = [DateTime]::MinValue  # Rate-limit LogBox.Refresh()
$script:LogRefreshIntervalMs = 100                  # Minimum ms between log repaints
$script:LogMaxLength        = 200000                # Trim log when exceeding ~200KB
$script:PathCache           = @{}                   # Test-Path cache: path -> [result, DateTime]
$script:PathCacheTtlMs      = 5000                  # Cache TTL for Test-PathCached (5s to reduce I/O on slow paths)
$script:NvidiaDllPatterns   = @('nv_*.dll','nvapi*.dll','nvcu*.dll','nvcuda*.dll',
                                'nvenc*.dll','nvfbc*.dll','nvml*.dll','nvopt*.dll',
                                'nvwgf2*.dll','nvidia*.dll')  # NVIDIA System32 DLL patterns (shared)

#endregion

#region ==================== UTILITY FUNCTIONS ====================

function Test-PathCached {
    <#
    .SYNOPSIS
        Cached wrapper around Test-Path to avoid UI freezes on slow paths.
        Results are cached for $script:PathCacheTtlMs milliseconds.
    #>
    param([string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) { return $false }
    $now = [DateTime]::UtcNow
    if ($script:PathCache.ContainsKey($Path)) {
        $entry = $script:PathCache[$Path]
        if (($now - $entry.Time).TotalMilliseconds -lt $script:PathCacheTtlMs) {
            return $entry.Result
        }
    }
    try {
        $result = Test-Path $Path
    } catch {
        $result = $false
    }
    # Evict stale entries when cache grows beyond 500 items to prevent unbounded memory use
    if ($script:PathCache.Count -gt 500) {
        $staleKeys = @($script:PathCache.GetEnumerator() | Where-Object { ($now - $_.Value.Time).TotalMilliseconds -ge $script:PathCacheTtlMs } | ForEach-Object { $_.Key })
        foreach ($k in $staleKeys) { $script:PathCache.Remove($k) }
        # Hard-cap fallback: if still over limit (all entries fresh), evict oldest entries regardless of TTL
        if ($script:PathCache.Count -gt 500) {
            $oldestKeys = @($script:PathCache.GetEnumerator() | Sort-Object { $_.Value.Time } | Select-Object -First ($script:PathCache.Count - 400) | ForEach-Object { $_.Key })
            foreach ($k in $oldestKeys) { $script:PathCache.Remove($k) }
        }
    }
    $script:PathCache[$Path] = @{ Result = $result; Time = $now }
    return $result
}

function Write-Log {
    param(
        [string]$Message,
        [ValidateSet("INFO","OK","WARN","ERROR")]
        [string]$Level = "INFO"
    )
    $timestamp = Get-Date -Format 'HH:mm:ss'
    if ($null -eq $script:LogBox) {
        Write-Output "$timestamp [$Level] $Message"
        return
    }

    # Trim oldest lines when text exceeds cap to prevent unbounded memory growth
    if ($script:LogBox.TextLength -gt $script:LogMaxLength) {
        $trimTo = [Math]::Max(0, $script:LogBox.TextLength - [int]($script:LogMaxLength * 0.7))
        $newlineIdx = $script:LogBox.Text.IndexOf("`n", $trimTo)
        if ($newlineIdx -gt 0) { $trimTo = $newlineIdx + 1 }
        $script:LogBox.Select(0, $trimTo)
        $script:LogBox.SelectedText = ""
        $script:LogBox.SelectionStart = $script:LogBox.TextLength
    }

    $color = switch ($Level) {
        "ERROR" { $script:theme.Danger }
        "WARN"  { $script:theme.Warning }
        "OK"    { $script:theme.Success }
        default { $script:theme.Text }
    }
    $script:LogBox.SelectionStart  = $script:LogBox.TextLength
    $script:LogBox.SelectionLength = 0
    $script:LogBox.SelectionColor  = $color
    $script:LogBox.AppendText("$timestamp [$Level] $Message`r`n")
    $script:LogBox.ScrollToCaret()
    # Rate-limited Refresh to avoid hammering UI during rapid logging (e.g. DISM, driver copy)
    $now = [DateTime]::UtcNow
    if (($now - $script:LastLogRefresh).TotalMilliseconds -ge $script:LogRefreshIntervalMs -or $Level -eq 'ERROR') {
        $script:LogBox.Refresh()
        $script:LastLogRefresh = $now
    }
}

function Write-UiWarning {
    param([string]$Message)

    if ([string]::IsNullOrWhiteSpace($Message)) { return }
    if ($script:LogBox) {
        Write-Log $Message "WARN"
    } else {
        Write-Output $Message
    }
}

function Get-ErrorGuidance {
    param([string]$Message)

    $hint = "Review the previous log lines and retry the operation."

    if ($Message -match 'Access is denied|UnauthorizedAccess|0x80070005') {
        $hint = "Run as Administrator and verify the VM folder/ISO/VHD permissions."
    } elseif ($Message -match 'not enough space|insufficient|0x80070070|There is not enough space') {
        $hint = "Free disk space, reduce VM disk size, or enable VHD auto-expand."
    } elseif ($Message -match 'used by another process|because it is being used') {
        $hint = "Close apps that may lock the VHD/ISO (Explorer, backup tools, vmconnect), then retry."
    } elseif ($Message -match 'not found|cannot find|does not exist') {
        $hint = "Verify VM name, VHD/ISO path, and virtual switch still exist."
    } elseif ($Message -match 'VMMS|Hyper-V|hypervisor') {
        $hint = "Ensure Hyper-V services are running and virtualization is enabled in BIOS/UEFI."
    } elseif ($Message -match 'partitionable|GPU-P|GpuPartition|InstancePath') {
        $hint = "Confirm GPU-P support on host and use default GPU selection if specific selection fails."
    } elseif ($Message -match 'TPM|KeyProtector|SecureBoot') {
        $hint = "For Windows 11, keep Secure Boot and TPM enabled; verify host supports vTPM."
    } elseif ($Message -match 'bcdboot|boot files|EFI') {
        $hint = "Retry with attached ISO recovery option and run Startup Repair if needed."
    }

    return $hint
}

function Write-ErrorWithGuidance {
    param(
        [string]$Context,
        [AllowNull()]$ErrorRecord
    )

    $message = if ($ErrorRecord -and $ErrorRecord.Exception -and $ErrorRecord.Exception.Message) {
        $ErrorRecord.Exception.Message
    } elseif ($ErrorRecord) {
        $ErrorRecord.ToString()
    } else {
        "Unknown error"
    }

    Write-Log "$Context failed: $message" "ERROR"
    Write-Log "$Context hint: $(Get-ErrorGuidance -Message $message)" "WARN"
}

function Register-TrackedMountedImage {
    param([string]$ImagePath)

    if ([string]::IsNullOrWhiteSpace($ImagePath)) { return }
    $script:TrackedMountedImages[$ImagePath] = $true
}

function Unregister-TrackedMountedImage {
    param([string]$ImagePath)

    if ([string]::IsNullOrWhiteSpace($ImagePath)) { return }
    if ($script:TrackedMountedImages.ContainsKey($ImagePath)) {
        [void]$script:TrackedMountedImages.Remove($ImagePath)
    }
}

function Invoke-MountCleanup {
    Write-Log "Running cleanup..."
    try {
        $trackedImagePaths = @($script:TrackedMountedImages.Keys)

        foreach ($imagePath in $trackedImagePaths) {
            try {
                $image = Get-DiskImage -ImagePath $imagePath -ErrorAction SilentlyContinue
                if ($image -and $image.Attached) {
                    Write-Log "Dismounting script-tracked image: $imagePath"
                    if (-not (Dismount-ImageRetry -ImagePath $imagePath -MaxRetries 2)) {
                        Write-Log "Cleanup dismount did not fully succeed for: $imagePath" "WARN"
                    }
                } else {
                    Unregister-TrackedMountedImage -ImagePath $imagePath
                }
            } catch {
                Write-Log "Cleanup dismount error for '$imagePath': $($_.Exception.Message)" "WARN"
            }
        }

        # Dismount ISO if still mounted (skip if already handled by tracked images loop)
        if ($script:MountedISO -and $script:MountedISO.ImagePath -and
            -not $script:TrackedMountedImages.ContainsKey($script:MountedISO.ImagePath)) {
            try {
                if (Dismount-ImageRetry -ImagePath $script:MountedISO.ImagePath -MaxRetries 2) {
                    Write-Log "Cleanup dismounted mounted ISO." "OK"
                }
            } catch {
                Write-Log "Cleanup ISO dismount error: $($_.Exception.Message)" "WARN"
            } finally {
                $script:MountedISO = $null
            }
        } elseif ($script:MountedISO) {
            # ISO was already dismounted by tracked-images loop
            $script:MountedISO = $null
        }
        Write-Log "Cleanup complete." "OK"
    } catch {
        Write-Log "Cleanup error: $($_.Exception.Message)" "WARN"
    }
}

function Set-AutoPlay {
    <#
    .SYNOPSIS
        Toggles AutoPlay on/off. Returns $true if the setting was changed (caller
        should restore it later), $false if no change was needed.
    #>
    param([bool]$Disable)
    $regPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\AutoplayHandlers"
    $regName = "DisableAutoplay"
    try {
        $current = (Get-ItemProperty -Path $regPath -Name $regName -ErrorAction Stop).$regName
    } catch { $current = 0 }

    try {
        if ($Disable -and $current -eq 0) {
            Set-ItemProperty -Path $regPath -Name $regName -Value 1 -ErrorAction Stop
            return $true   # Changed: was enabled, now disabled
        } elseif (-not $Disable -and $current -eq 1) {
            Set-ItemProperty -Path $regPath -Name $regName -Value 0 -ErrorAction Stop
            return $true   # Changed: was disabled, now restored
        }
    } catch {
        # Registry write failed (e.g. GPO-locked); report no change to avoid bad restore
        return $false
    }
    return $false      # No change needed
}

function Get-AutoPlayState {
    $regPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\AutoplayHandlers"
    $regName = "DisableAutoplay"
    try {
        return [int](Get-ItemProperty -Path $regPath -Name $regName -ErrorAction Stop).$regName
    } catch {
        return 0
    }
}

function Restore-AutoPlayState {
    param([AllowNull()][System.Nullable[int]]$State)

    if ($null -eq $State) { return }

    $regPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\AutoplayHandlers"
    $regName = "DisableAutoplay"
    try {
        Set-ItemProperty -Path $regPath -Name $regName -Value $State -ErrorAction Stop
    } catch {
        Write-Output "Could not restore AutoPlay registry value: $($_.Exception.Message)"
    }
}

function Disable-AutoPlayGuarded {
    <#
    .SYNOPSIS
        Saves the current AutoPlay state, disables AutoPlay, and returns the
        original state as a restore token.  Returns $null if no change was made.
        Pass the result to Restore-AutoPlayState when finished.
    #>
    $original = Get-AutoPlayState
    $changed  = Set-AutoPlay -Disable $true
    if ($changed) { return $original }
    return $null
}

function Dismount-ImageRetry {
    <#
    .SYNOPSIS
        Dismounts a disk image with retry logic to handle in-use locks.
    #>
    param([string]$ImagePath, [int]$MaxRetries = 5)
    if ($MaxRetries -lt 1) { $MaxRetries = 1 }
    for ($i = 1; $i -le $MaxRetries; $i++) {
        try {
            $img = Get-DiskImage -ImagePath $ImagePath -ErrorAction SilentlyContinue
            if (-not $img -or -not $img.Attached) {
                Unregister-TrackedMountedImage -ImagePath $ImagePath
                return $true
            }
            Dismount-DiskImage -ImagePath $ImagePath -ErrorAction Stop
            Unregister-TrackedMountedImage -ImagePath $ImagePath
            return $true
        } catch {
            if ($i -lt $MaxRetries) {
                Write-Log "Dismount retry $i/$MaxRetries for $(Split-Path $ImagePath -Leaf): $($_.Exception.Message)" "WARN"
                Start-Sleep -Seconds (2 * $i)
            } else {
                Write-Log "Failed to dismount after $MaxRetries retries: $ImagePath" "ERROR"
                return $false
            }
        }
    }
}

function Wait-ImageDetached {
    param(
        [string]$ImagePath,
        [int]$TimeoutSec = 20
    )

    if ([string]::IsNullOrWhiteSpace($ImagePath)) { return $true }
    if ($TimeoutSec -lt 1) { $TimeoutSec = 1 }

    $deadline = (Get-Date).AddSeconds($TimeoutSec)
    do {
        $img = Get-DiskImage -ImagePath $ImagePath -ErrorAction SilentlyContinue
        if (-not $img -or -not $img.Attached) {
            return $true
        }
        Start-Sleep -Milliseconds 500
    } while ((Get-Date) -lt $deadline)

    Write-Log "Wait-ImageDetached: '$ImagePath' still attached after ${TimeoutSec}s timeout." "WARN"
    return $false
}

function Stop-VMWithTimeout {
    param(
        [string]$VMName,
        [int]$TimeoutSec = 60
    )

    try {
        $vm = Get-VM -Name $VMName -ErrorAction Stop
        if ($vm.State -eq 'Off') { return $true }

        try {
            Stop-VM -Name $VMName -Force -ErrorAction Stop
        } catch {
            Write-Log "[$VMName] Graceful stop failed, trying turn-off fallback: $($_.Exception.Message)" "WARN"
            Stop-VM -Name $VMName -TurnOff -Force -ErrorAction SilentlyContinue | Out-Null
        }

        $elapsed = 0
        while ($elapsed -lt $TimeoutSec) {
            Start-Sleep -Seconds 1
            $elapsed++
            $vm = Get-VM -Name $VMName -ErrorAction SilentlyContinue
            if ($vm -and $vm.State -eq 'Off') { return $true }
        }

        Stop-VM -Name $VMName -TurnOff -Force -ErrorAction SilentlyContinue | Out-Null
        Start-Sleep -Seconds 3
        $vm = Get-VM -Name $VMName -ErrorAction SilentlyContinue
        return ($vm -and $vm.State -eq 'Off')
    } catch {
        Write-Log "[$VMName] Stop fallback failed: $($_.Exception.Message)" "WARN"
        return $false
    }
}

function Get-VMPrimaryVhdPath {
    param([string]$VMName)

    try {
        $hardDisks = @(Get-VMHardDiskDrive -VMName $VMName -ErrorAction Stop)
        if ($hardDisks.Count -eq 0) { return $null }

        try {
            $firmware = Get-VMFirmware -VMName $VMName -ErrorAction SilentlyContinue
            if ($firmware -and $firmware.BootOrder) {
                foreach ($bootDevice in $firmware.BootOrder) {
                    if ($bootDevice -and $bootDevice.Device -and
                        $bootDevice.BootType -eq 'Drive' -and
                        $bootDevice.Device.Path) {
                        $bootPath = $bootDevice.Device.Path
                        if (-not [string]::IsNullOrWhiteSpace($bootPath) -and (Test-Path $bootPath)) {
                            return $bootPath
                        }
                    }
                }
            }
        } catch {
            # Fall through to controller and first-disk heuristics
            Write-Verbose "Get-VMFirmware failed for '$VMName': $($PSItem.Exception.Message)"
        }

        $controllerDisk = $hardDisks | Where-Object {
            $_.ControllerNumber -eq 0 -and $_.ControllerLocation -eq 0
        } | Select-Object -First 1
        if ($controllerDisk -and -not [string]::IsNullOrWhiteSpace($controllerDisk.Path)) {
            return $controllerDisk.Path
        }

        $firstExisting = $hardDisks | Where-Object {
            -not [string]::IsNullOrWhiteSpace($_.Path) -and (Test-Path $_.Path)
        } | Select-Object -First 1
        if ($firstExisting) { return $firstExisting.Path }

        return $hardDisks[0].Path
    } catch {
        return $null
    }
}

function Mount-VhdWithFallback {
    param([string]$ImagePath)

    try {
        Mount-DiskImage -ImagePath $ImagePath -ErrorAction Stop
    } catch {
        Write-Log "Mount-DiskImage failed for $(Split-Path $ImagePath -Leaf), trying Mount-VHD fallback..." "WARN"
        try {
            Mount-VHD -Path $ImagePath -ErrorAction Stop | Out-Null
        } catch {
            Write-ErrorWithGuidance -Context "Mount VHD ($(Split-Path $ImagePath -Leaf))" -ErrorRecord $_
            return $false
        }
    }

    for ($i = 0; $i -lt 10; $i++) {
        $img = Get-DiskImage -ImagePath $ImagePath -ErrorAction SilentlyContinue
        if ($img -and $img.Attached) {
            Register-TrackedMountedImage -ImagePath $ImagePath
            return $true
        }
        Start-Sleep -Seconds 1
    }

    Write-Log "Image did not report attached state after mount: $ImagePath" "ERROR"
    return $false
}

function Start-VMWithRetry {
    param(
        [string]$VMName,
        [int]$MaxRetries = 2
    )

    # Skip retries if VM is already running
    $currentVm = Get-VM -Name $VMName -ErrorAction SilentlyContinue
    if ($currentVm -and $currentVm.State -eq 'Running') { return $true }

    for ($i = 1; $i -le $MaxRetries; $i++) {
        try {
            Start-VM -Name $VMName -ErrorAction Stop | Out-Null
            return $true
        } catch {
            if ($i -lt $MaxRetries) {
                Write-Log "[$VMName] Start retry $i/${MaxRetries}: $($_.Exception.Message)" "WARN"
                Start-Sleep -Seconds (2 * $i)
            } else {
                Write-Log "[$VMName] Failed to start VM after $MaxRetries attempts - $($_.Exception.Message)" "ERROR"
                return $false
            }
        }
    }
    return $false  # safety fallthrough (MaxRetries <= 0 guard)
}

function Get-PathAvailableSpaceGB {
    param([string]$Path)
    try {
        $resolvedPath = (Resolve-Path -Path $Path -ErrorAction Stop).Path
        $root = [System.IO.Path]::GetPathRoot($resolvedPath)
        if (-not $root) { return -1 }

        # Local drive path (e.g. C:\)
        if ($root -match '^([A-Za-z]):\\') {
            $driveInfo = New-Object System.IO.DriveInfo($matches[1])
            if ($driveInfo.IsReady) {
                return [math]::Round(($driveInfo.AvailableFreeSpace / 1GB), 2)
            }
        }

        # Fallback for mapped drives via CIM
        if ($root -match '^([A-Za-z]:)') {
            $disk = Get-CimInstance -Query "SELECT FreeSpace FROM Win32_LogicalDisk WHERE DeviceID='$($matches[1])'" -ErrorAction SilentlyContinue
            if ($disk -and $disk.FreeSpace) {
                return [math]::Round(($disk.FreeSpace / 1GB), 2)
            }
        }

        # UNC paths - space check not supported, warn user
        if ($root -match '^\\\\') {
            Write-Log "UNC path detected: $Path. Free space cannot be verified for network paths." "WARN"
        }
        return -1
    } catch {
        return -1
    }
}

function Test-DirectoryWritable {
    param([string]$Path)
    $testFile = $null
    try {
        if (-not (Test-Path -Path $Path)) { return $false }
        $testFile = Join-Path $Path ([System.Guid]::NewGuid().ToString() + '.tmp')
        Set-Content -Path $testFile -Value 'test' -Encoding ASCII -ErrorAction Stop
        Remove-Item -Path $testFile -Force -ErrorAction SilentlyContinue
        $testFile = $null
        return $true
    } catch {
        return $false
    } finally {
        if ($testFile -and (Test-Path $testFile)) {
            Remove-Item -Path $testFile -Force -ErrorAction SilentlyContinue
        }
    }
}

# Well-known GPT partition type GUIDs
$script:GptEfi = '{C12A7328-F81F-11D2-BA4B-00A0C93EC93B}'
$script:GptMsr = '{E3C9E316-0B5C-4DB8-817D-F92DF00215AE}'

function Get-DataPartitions {
    <#
    .SYNOPSIS
        Returns non-system, non-reserved partitions for a given disk number.
        Centralises the GPT filter logic used in multiple places.
    #>
    param([int]$DiskNumber)
    Get-Partition -DiskNumber $DiskNumber -ErrorAction SilentlyContinue |
        Where-Object {
            $_.Type -ne 'System' -and
            $_.Type -ne 'Reserved' -and
            $_.GptType -ne $script:GptEfi -and
            $_.GptType -ne $script:GptMsr
        }
}

function Set-ToolkitNatSwitch {
    [CmdletBinding()]
    param(
        [string]$SwitchName = "HyperV-Toolkit-NAT",
        [string]$GatewayIp = "192.168.250.1",
        [int]$PrefixLength = 24,
        [string]$NatPrefix = "192.168.250.0/24"
    )

    $switchCreatedByUs = $false
    try {
        $existingSwitch = Get-VMSwitch -Name $SwitchName -ErrorAction SilentlyContinue
        if ($existingSwitch -and $existingSwitch.SwitchType -ne 'Internal') {
            Write-Log "Existing switch '$SwitchName' is type '$($existingSwitch.SwitchType)' (expected Internal). NAT configuration may not work correctly." "WARN"
        }
        if (-not $existingSwitch) {
            Write-Log "No usable virtual switch selected. Creating internal NAT switch '$SwitchName'..." "WARN"
            New-VMSwitch -Name $SwitchName -SwitchType Internal -ErrorAction Stop | Out-Null
            $switchCreatedByUs = $true
            Start-Sleep -Seconds 1
            Write-Log "Created virtual switch: $SwitchName" "OK"
        }

        $adapterAlias = "vEthernet ($SwitchName)"
        $adapter = Get-NetAdapter -Name $adapterAlias -ErrorAction SilentlyContinue
        if (-not $adapter) {
            Write-Log "Could not find host adapter '$adapterAlias' after switch creation." "WARN"
        } else {
            $ipExists = Get-NetIPAddress -InterfaceAlias $adapterAlias -AddressFamily IPv4 -ErrorAction SilentlyContinue |
                Where-Object { $_.IPAddress -eq $GatewayIp }

            if (-not $ipExists) {
                # Check for conflicting gateway IP on a different interface
                $conflictingIp = Get-NetIPAddress -AddressFamily IPv4 -ErrorAction SilentlyContinue |
                    Where-Object { $_.IPAddress -eq $GatewayIp -and $_.InterfaceAlias -ne $adapterAlias }
                if ($conflictingIp) {
                    Write-Log "Gateway IP $GatewayIp is already assigned to interface '$($conflictingIp.InterfaceAlias)'. Skipping assignment." "WARN"
                } else {
                    try {
                        New-NetIPAddress -InterfaceAlias $adapterAlias -IPAddress $GatewayIp -PrefixLength $PrefixLength -AddressFamily IPv4 -ErrorAction Stop | Out-Null
                        Write-Log "Assigned $GatewayIp/$PrefixLength to $adapterAlias" "OK"
                    } catch {
                        Write-Log "IP assignment skipped or failed for ${adapterAlias}: $($_.Exception.Message)" "WARN"
                    }
                }
            }
        }

        $natName = "$SwitchName-NAT"
        $existingNat = Get-NetNat -Name $natName -ErrorAction SilentlyContinue
        if (-not $existingNat) {
            # Check for conflicting NAT or routes using the same subnet
            $conflictingNat = Get-NetNat -ErrorAction SilentlyContinue |
                Where-Object { $_.InternalIPInterfaceAddressPrefix -eq $NatPrefix -and $_.Name -ne $natName }
            if ($conflictingNat) {
                Write-Log "NAT subnet $NatPrefix already in use by '$($conflictingNat.Name)'. Reusing existing NAT." "WARN"
            } else {
                New-NetNat -Name $natName -InternalIPInterfaceAddressPrefix $NatPrefix -ErrorAction Stop | Out-Null
                Write-Log "Created NAT object '$natName' with prefix $NatPrefix" "OK"
            }
        }

        return $SwitchName
    } catch {
        Write-Log "Auto-create switch/NAT failed: $($_.Exception.Message)" "ERROR"
        # Clean up partially created switch if we created it in this call
        if ($switchCreatedByUs) {
            try {
                Remove-VMSwitch -Name $SwitchName -Force -ErrorAction SilentlyContinue
                Write-Log "Removed partially configured switch '$SwitchName' during cleanup." "WARN"
            } catch {
                Write-Log "Could not clean up partial switch '$SwitchName': $($_.Exception.Message)" "WARN"
            }
        }
        return $null
    }
}

function Convert-PlainTextToSecureString {
    param([AllowNull()][string]$Text)

    if ([string]::IsNullOrEmpty($Text)) {
        return $null
    }

    $secure = New-Object System.Security.SecureString
    foreach ($char in $Text.ToCharArray()) {
        $secure.AppendChar($char)
    }
    $secure.MakeReadOnly()
    return $secure
}

function Update-CreateProgress {
    param(
        [int]$Percent,
        [string]$Status = ""
    )

    if ($ctrlCreate -and $ctrlCreate.ContainsKey("CreateProgress")) {
        $safePercent = [Math]::Max(0, [Math]::Min(100, $Percent))
        $ctrlCreate["CreateProgress"].Value = $safePercent
        $ctrlCreate["CreateProgress"].Refresh()
    }
    if ($ctrlCreate -and $ctrlCreate.ContainsKey("CreateStatus")) {
        $ctrlCreate["CreateStatus"].Text = $Status
        $ctrlCreate["CreateStatus"].Refresh()
    }
    # Keep UI responsive during long sync operations. Re-entrancy-sensitive actions are
    # already guarded by $script:IsCreating / $script:IsUpdatingGPU flags.
    try {
        [System.Windows.Forms.Application]::DoEvents()
    } catch {
        if (-not $script:DoEventsWarningLogged) {
            $script:DoEventsWarningLogged = $true
            Write-Log "UI event pump warning: $($_.Exception.Message)" "WARN"
        }
    }
}

function Remove-PartialVmArtifacts {
    param(
        [string]$VMName,
        [string]$VMLoc,
        [string]$VHDPath,
        [bool]$RemoveVmFolder = $false
    )

    Write-Log "Attempting rollback cleanup for '$VMName'..." "WARN"

    try {
        $vm = Get-VM -Name $VMName -ErrorAction SilentlyContinue
        if ($vm) {
            if ($vm.State -ne 'Off') {
                Stop-VM -Name $VMName -TurnOff -Force -ErrorAction SilentlyContinue | Out-Null
            }
            Remove-VM -Name $VMName -Force -ErrorAction SilentlyContinue
            Write-Log "Removed partially created VM: $VMName" "WARN"
        }
    } catch {
        Write-Log "Rollback warning (VM remove): $($_.Exception.Message)" "WARN"
    }

    try {
        if ($VHDPath -and (Test-Path $VHDPath)) {
            if (-not (Dismount-ImageRetry -ImagePath $VHDPath -MaxRetries 3)) {
                # Fallback to Dismount-VHD if Dismount-ImageRetry failed
                Dismount-VHD -Path $VHDPath -ErrorAction SilentlyContinue
                Unregister-TrackedMountedImage -ImagePath $VHDPath
            }
            # Delete the stale VHDX file after successful dismount
            try {
                if (Test-Path $VHDPath) {
                    Remove-Item -Path $VHDPath -Force -ErrorAction Stop
                    Write-Log "Removed stale VHD file: $VHDPath" "WARN"
                }
            } catch {
                Write-Log "Rollback warning (VHD file delete): $($_.Exception.Message)" "WARN"
            }
        }
    } catch {
        Write-Log "Rollback warning (VHD dismount): $($_.Exception.Message)" "WARN"
    }

    try {
        if ($RemoveVmFolder -and $VMLoc -and (Test-Path $VMLoc)) {
            Remove-Item -Path $VMLoc -Recurse -Force -ErrorAction SilentlyContinue
            Write-Log "Removed partially created VM folder: $VMLoc" "WARN"
        } elseif ($VMLoc -and (Test-Path $VMLoc)) {
            Write-Log "Rollback safety: preserved existing VM folder (not created by this run): $VMLoc" "WARN"
        }
    } catch {
        Write-Log "Rollback warning (folder cleanup): $($_.Exception.Message)" "WARN"
    }
}

function Set-VMGuestSecureBoot {
    param(
        [string]$VMName,
        [bool]$EnableSecureBoot,
        [bool]$GuestIsWindows11,
        [int]$GuestBuild = 0,
        [AllowNull()][string[]]$TemplateOrder = $null
    )

    if (-not $EnableSecureBoot) {
        try {
            Set-VMFirmware -VMName $VMName -EnableSecureBoot Off -ErrorAction Stop
            Write-Log "  Disabling Secure Boot (Windows 10 compatibility mode)" "OK"
            return $true
        } catch {
            Write-Log "  Secure Boot disable failed: $($_.Exception.Message)" "WARN"
            return $false
        }
    }

    try {
        Set-VMFirmware -VMName $VMName -EnableSecureBoot On -ErrorAction Stop | Out-Null
    } catch {
        Write-Log "  Secure Boot enable pre-step failed: $($_.Exception.Message)" "WARN"
        Write-Log "  Cannot proceed with Secure Boot template configuration." "WARN"
        return $false
    }
    $templates = $TemplateOrder
    if (-not $templates -or $templates.Count -eq 0) {
        $templates = if ($GuestIsWindows11) {
            @('MicrosoftWindows', 'MicrosoftUEFICertificateAuthority')
        } elseif ($GuestBuild -gt 0 -and $GuestBuild -lt $script:BUILD_WIN10_RS4) {
            @('MicrosoftWindows', 'MicrosoftUEFICertificateAuthority')
        } else {
            @('MicrosoftWindows', 'MicrosoftUEFICertificateAuthority')
        }
    }

    $lastTemplateError = $null

    foreach ($template in $templates) {
        try {
            Set-VMFirmware -VMName $VMName -EnableSecureBoot On -SecureBootTemplate $template -ErrorAction Stop
            Write-Log "  Secure Boot: Enabled ($template)" "OK"
            return $true
        } catch {
            $lastTemplateError = $_.Exception.Message
        }
    }

    if ($lastTemplateError) {
        Write-Log "  Secure Boot enabled, but no known template could be explicitly applied. Last template error: $lastTemplateError" "WARN"
    } else {
        Write-Log "  Secure Boot enabled, but no known template could be explicitly applied. Using host default template." "WARN"
    }
    return $false
}

function Enable-VirtualTpmForVm {
    param(
        [Parameter(Mandatory = $true)][string]$VMName,
        [bool]$GuestIsWindows11 = $false
    )

    $requiredCmds = @('Set-VMKeyProtector', 'Enable-VMTPM')
    foreach ($cmd in $requiredCmds) {
        if (-not (Get-Command $cmd -ErrorAction SilentlyContinue)) {
            Write-Log "  TPM enable skipped: required Hyper-V command '$cmd' is unavailable on this host." "WARN"
            return $false
        }
    }

    # vTPM for Hyper-V VMs does not require a physical host TPM when using a local key protector.
    Write-Log "  TPM setup: using local key protector (host physical TPM is not required)." "INFO"

    foreach ($svcName in @('vmcompute', 'vmms')) {
        try {
            $svc = Get-Service -Name $svcName -ErrorAction Stop
            if ($svc.Status -ne 'Running') {
                Start-Service -Name $svcName -ErrorAction Stop
                Write-Log "  TPM setup: started required service '$svcName'." "INFO"
            }
        } catch {
            Write-Log "  TPM setup warning: service '$svcName' is not ready: $($_.Exception.Message)" "WARN"
        }
    }

    try {
        Set-VMKeyProtector -VMName $VMName -NewLocalKeyProtector -ErrorAction Stop
        Enable-VMTPM -VMName $VMName -ErrorAction Stop

        $verified = $false
        if (Get-Command Get-VMSecurity -ErrorAction SilentlyContinue) {
            try {
                $vmSec = Get-VMSecurity -VMName $VMName -ErrorAction Stop
                if ($vmSec -and $vmSec.TpmEnabled) {
                    $verified = $true
                }
            } catch {
                Write-Log "  TPM verification warning: $($_.Exception.Message)" "WARN"
            }
        }

        if ($verified) {
            Write-Log "  Virtual TPM: Enabled and verified" "OK"
        } else {
            Write-Log "  Virtual TPM: Enable command completed" "OK"
        }

        if ($GuestIsWindows11) {
            Write-Log "  TPM 2.0 enabled for Windows 11 compatibility" "INFO"
        }
        return $true
    } catch {
        Write-Log "  TPM setup failed: $($_.Exception.Message)" "WARN"
        if ($GuestIsWindows11) {
            Write-Log "  WARNING: Windows 11 installation may fail without TPM" "WARN"
        }
        return $false
    }
}

function Convert-SecureStringToPlainText {
    param([AllowNull()][System.Security.SecureString]$SecureString)
    if (-not $SecureString) { return "" }

    $bstr = [IntPtr]::Zero
    try {
        $bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecureString)
        return [Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr)
    } finally {
        if ($bstr -ne [IntPtr]::Zero) {
            [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)
        }
    }
}

function ConvertTo-XmlEscapedValue {
    param([AllowNull()][string]$Value)
    if ($null -eq $Value) { return "" }
    return [System.Security.SecurityElement]::Escape($Value)
}

function ConvertTo-UnattendPassword {
    <#
    .SYNOPSIS
        Encodes a password for use in Windows Unattend XML with PlainText=false.
        Windows format: base64( UTF-16LE( password + "Password" ) )
    #>
    param([AllowNull()][string]$PlainText)
    if ([string]::IsNullOrEmpty($PlainText)) { 
        Write-Log "Warning: Empty password provided to ConvertTo-UnattendPassword" "WARN"
        return "" 
    }
    $combined = $PlainText + "Password"
    $bytes = [System.Text.Encoding]::Unicode.GetBytes($combined)
    return [Convert]::ToBase64String($bytes)
}

function ConvertTo-UnattendArchitecture {
    param([AllowNull()][string]$Architecture)

    if ([string]::IsNullOrWhiteSpace($Architecture)) { return "" }

    switch -Regex ($Architecture.Trim().ToLowerInvariant()) {
        '^(amd64|x64|x86_64)$'  { return 'amd64' }
        '^(arm64|aarch64)$'      { return 'arm64' }
        '^(x86|i[3-6]86)$'      { return 'x86' }
        default {
            Write-Log "Unrecognized guest architecture: '$Architecture'. Falling back to host architecture." "WARN"
            return ""
        }
    }
}

function Test-GpuPPreFlight {
    <#
    .SYNOPSIS
        Runs pre-flight checks for GPU-P based on Diobyte Version 1 guidance.
        Returns array of warning/error messages.
    #>
    $issues = @()

    # Fetch video controllers from startup cache
    $videoControllers = $script:VideoControllers

    # Check 1: Laptop NVIDIA GPU detection (unsupported for GPU-P)
    try {
        $battery = Get-CimInstance -ClassName Win32_Battery -ErrorAction SilentlyContinue
        if ($battery) {
            $nvidiaGpu = $videoControllers | Where-Object { $_.Name -match 'NVIDIA' }
            if ($nvidiaGpu) {
                $issues += "WARNING: Laptop NVIDIA GPUs are NOT supported for GPU-P. Intel integrated GPUs on laptops may work instead."
            }
        }
    } catch {
        Write-Log "GPU preflight battery/laptop check skipped: $($_.Exception.Message)" "WARN"
    }

    # Check 2: AMD Polaris (RX 580 etc) - no hardware video encoding
    try {
        $amdGpu = $videoControllers | Where-Object { $_.Name -match 'RX\s*5[678]0|Polaris' }
        if ($amdGpu) {
            $issues += "WARNING: AMD Polaris GPUs (RX 580 etc) do not support hardware video encoding via GPU-PV."
        }
    } catch {
        Write-Log "GPU preflight AMD capability check skipped: $($_.Exception.Message)" "WARN"
    }

    # Check 3: GPU must support partitioning
    try {
        $partGpu = Get-VMHostPartitionableGpu -ErrorAction SilentlyContinue
        if (-not $partGpu -or $partGpu.Count -eq 0) {
            $issues += "ERROR: No GPU-P capable GPUs found. Your GPU may not support partitioning."
        }
    } catch {
        $issues += "ERROR: Cannot query partitionable GPUs. Ensure Hyper-V is fully enabled."
    }

    # Check 4: Win10 requires GPUName AUTO
    if ($script:HostBuild -lt $script:BUILD_WIN11_MIN) {
        $issues += "INFO: Windows 10 host detected - GPU-P will use AUTO (default GPU). Specific GPU selection requires Windows 11."
    }

    return $issues
}

function Test-GpuPHostReadiness {
    param([bool]$RequireSriov = $false)

    $errors = @()
    $warnings = @()

    try {
        $partGpu = @(Get-VMHostPartitionableGpu -ErrorAction SilentlyContinue)
        if (-not $partGpu -or $partGpu.Count -eq 0) {
            $errors += "No partitionable GPUs were reported by Hyper-V on this host."
        }
    } catch {
        $errors += "Failed to query partitionable GPUs: $($_.Exception.Message)"
    }

    try {
        $sriovAdapters = @(Get-NetAdapterSriov -ErrorAction SilentlyContinue)
        if (-not $sriovAdapters -or $sriovAdapters.Count -eq 0) {
            if ($RequireSriov) {
                $errors += "No SR-IOV-capable adapters detected (required for this host profile)."
            } else {
                $warnings += "No SR-IOV-capable adapters detected. On enterprise/clustered GPU-P hosts, SR-IOV is required."
            }
        } else {
            $ready = $sriovAdapters | Where-Object {
                $_.SriovEnabled -eq $true -or $_.SriovSupport -match 'Supported|Ready'
            }
            if (-not $ready) {
                if ($RequireSriov) {
                    $errors += "SR-IOV adapters found, but none are enabled/ready."
                } else {
                    $warnings += "SR-IOV adapters are present but not enabled/ready."
                }
            }
        }
    } catch {
        if ($RequireSriov) {
            $errors += "Failed to query SR-IOV state: $($_.Exception.Message)"
        } else {
            $warnings += "Could not validate SR-IOV state: $($_.Exception.Message)"
        }
    }

    [PSCustomObject]@{
        CanProceed = ($errors.Count -eq 0)
        Errors = $errors
        Warnings = $warnings
    }
}

function Ensure-HostGpuPartitionCountValid {
    [CmdletBinding()]
    param(
        [string]$PreferredGpuName = ""
    )

    try {
        $all = @(Get-VMHostPartitionableGpu -ErrorAction SilentlyContinue)
        if (-not $all -or $all.Count -eq 0) {
            return [PSCustomObject]@{ Success = $false; Changed = $false; Message = "No partitionable GPUs reported." }
        }

        $gpu = $null
        if (-not [string]::IsNullOrWhiteSpace($PreferredGpuName)) {
            $gpu = $all | Where-Object { $_.Name -eq $PreferredGpuName } | Select-Object -First 1
        }
        if (-not $gpu) { $gpu = $all | Select-Object -First 1 }
        if (-not $gpu) {
            return [PSCustomObject]@{ Success = $false; Changed = $false; Message = "Could not resolve target partitionable GPU." }
        }

        $validCounts = @()
        if ($gpu.PSObject.Properties['ValidPartitionCounts']) {
            $validCounts = @($gpu.ValidPartitionCounts | ForEach-Object { [int]$_ } | Where-Object { $_ -gt 0 } | Sort-Object -Unique)
        }
        $currentCount = 0
        if ($gpu.PSObject.Properties['PartitionCount']) {
            $currentCount = [int]$gpu.PartitionCount
        }

        if (-not $validCounts -or $validCounts.Count -eq 0) {
            return [PSCustomObject]@{ Success = $true; Changed = $false; Message = "ValidPartitionCounts not reported by driver; skipping partition-count normalization." }
        }

        if ($validCounts -contains $currentCount) {
            return [PSCustomObject]@{ Success = $true; Changed = $false; Message = "Host GPU partition count is valid ($currentCount)." }
        }

        $targetCount = ($validCounts | Measure-Object -Maximum).Maximum
        Set-VMHostPartitionableGpu -Name $gpu.Name -PartitionCount ([UInt16]$targetCount) -ErrorAction Stop | Out-Null
        return [PSCustomObject]@{ Success = $true; Changed = $true; Message = "Adjusted host GPU partition count from $currentCount to valid value $targetCount." }
    } catch {
        return [PSCustomObject]@{ Success = $false; Changed = $false; Message = $_.Exception.Message }
    }
}

function Get-HostNvidiaDriverVersion {
    try {
        $nvidiaDrivers = Get-CimInstance Win32_PnPSignedDriver -ErrorAction SilentlyContinue |
            Where-Object {
                $_.DeviceClass -eq 'DISPLAY' -and
                ($_.Manufacturer -match 'NVIDIA' -or $_.DeviceName -match 'NVIDIA') -and
                -not [string]::IsNullOrWhiteSpace($_.DriverVersion)
            }

        if (-not $nvidiaDrivers) { return "" }

        $latest = $nvidiaDrivers |
            Sort-Object {
                try { [version]$_.DriverVersion } catch { [version]'0.0.0.0' }
            } -Descending |
            Select-Object -First 1

        return [string]$latest.DriverVersion
    } catch {
        return ""
    }
}

function Get-NvidiaDriverVersionFromGuestStore {
    param([string]$MountLetter)

    try {
        if ([string]::IsNullOrWhiteSpace($MountLetter)) { return "" }
        $repoPath = Join-Path "$MountLetter\" 'Windows\System32\HostDriverStore\FileRepository'
        if (-not (Test-Path $repoPath)) { return "" }

        $infFiles = Get-ChildItem -Path $repoPath -Recurse -Filter '*.inf' -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -match '^nv.*\.inf$' }
        if (-not $infFiles) { return "" }

        foreach ($inf in $infFiles) {
            try {
                $driverVerLine = Select-String -Path $inf.FullName -Pattern '^\s*DriverVer\s*=\s*.+?(\d+\.\d+\.\d+\.\d+)' -CaseSensitive:$false -ErrorAction SilentlyContinue | Select-Object -First 1
                if ($driverVerLine -and $driverVerLine.Matches.Count -gt 0) {
                    return $driverVerLine.Matches[0].Groups[1].Value
                }
            } catch {
                continue
            }
        }

        return ""
    } catch {
        return ""
    }
}

function Get-GuestWindowsBuildFromMountedVolume {
    param([string]$MountLetter)

    if ([string]::IsNullOrWhiteSpace($MountLetter)) { return 0 }

    $softwareHive = Join-Path "$MountLetter\" 'Windows\System32\Config\SOFTWARE'
    if (-not (Test-Path $softwareHive)) { return 0 }

    $hiveName = "__gpu_guest_build_$([Guid]::NewGuid().ToString('N'))"
    $hiveRoot = "HKU\$hiveName"
    $loaded = $false

    try {
        & reg.exe load $hiveRoot "$softwareHive" *> $null
        if ($LASTEXITCODE -ne 0) { return 0 }
        $loaded = $true

        $buildRaw = (Get-ItemProperty -Path "Registry::$hiveRoot\Microsoft\Windows NT\CurrentVersion" -Name CurrentBuild -ErrorAction SilentlyContinue).CurrentBuild
        $parsed = 0
        if ([int]::TryParse([string]$buildRaw, [ref]$parsed)) {
            return $parsed
        }
        return 0
    } catch {
        return 0
    } finally {
        if ($loaded) {
            & reg.exe unload $hiveRoot *> $null
        }
    }
}

function Set-GuestGpuRegistryMitigations {
    <#
    .SYNOPSIS
        Applies GPU-P compatibility registry fixes to a mounted guest Windows volume
        by loading the offline SYSTEM hive, writing the required service overrides,
        then unloading it cleanly.

        Changes applied to ALL GPU-P guests
        ────────────────────────────────────
        HyperVideo (Hyper-V Synthetic Video)  Start = 4  (Disabled)
            The synthetic Hyper-V video adapter and a GPU partition adapter cannot
            both initialise during the same boot. dxgkrnl.sys serialises GPU device
            start-up; when HyperVideo is present it takes the first slot and the
            GPU-P adapter never gets a PnP start callback — manifests as a frozen
            boot spinner.  Disabling the service lets the real GPU-P adapter own
            the display stack from first boot.

        Additional changes for build >= 26100 (Win11 24H2 / 25H2)
        ────────────────────────────────────────────────────────────
        BasicDisplay                          Start = 4  (Disabled)
            On 24H2+ the basic display fallback driver races with GPU-P adapter
            initialisation during boot, occasionally winning the race and leaving
            the GPU-P adapter in a zombie state that hangs dxgkrnl.sys.
        GraphicsDrivers\TdrDelay             = 60  (seconds)
            The default TDR timeout is 2 s.  GPU-P adapter enumeration on 24H2+
            is slower post-update and the early-boot TDR fires before the adapter
            finishes starting, triggering a recovery loop that freezes the spinner.
    #>
    param(
        [string]$VMName,
        [string]$MountLetter,
        [int]$GuestBuild = 0
    )

    $systemHive = Join-Path "$MountLetter\" 'Windows\System32\Config\SYSTEM'
    if (-not (Test-Path $systemHive)) {
        Write-Log "[$VMName] Guest SYSTEM hive not found at '$systemHive'; skipping registry mitigations." "WARN"
        return
    }

    $hiveName = "__gpup_sysfix_$([Guid]::NewGuid().ToString('N'))"
    $hiveRoot = "HKU\$hiveName"
    $loaded   = $false

    try {
        & reg.exe load $hiveRoot "$systemHive" *> $null
        if ($LASTEXITCODE -ne 0) {
            Write-Log "[$VMName] Could not load guest SYSTEM hive for GPU-P registry mitigations (exit $LASTEXITCODE)." "WARN"
            return
        }
        $loaded = $true

        # Determine which control sets exist in this hive
        $controlSets = @('ControlSet001','ControlSet002') | Where-Object {
            Test-Path "Registry::$hiveRoot\$_"
        }
        if (-not $controlSets) { $controlSets = @('ControlSet001') }

        foreach ($cs in $controlSets) {
            $svcBase = "Registry::$hiveRoot\$cs\Services"

            # ── HyperVideo: disable Hyper-V synthetic video adapter ──────────────
            $hyperVideoPath = "$svcBase\HyperVideo"
            if (-not (Test-Path $hyperVideoPath)) {
                New-Item -Path $hyperVideoPath -Force | Out-Null
            }
            Set-ItemProperty -Path $hyperVideoPath -Name 'Start' -Value 4 -Type DWord -Force
            Write-Log "[$VMName] [$cs] HyperVideo service disabled (GPU-P conflict prevention)." "INFO"

            # ── Per-build mitigations for Win11 24H2 / 25H2 (build >= 26100) ────
            if ($GuestBuild -ge 26100) {

                # BasicDisplay: disable fallback display driver race
                $basicDisplayPath = "$svcBase\BasicDisplay"
                if (-not (Test-Path $basicDisplayPath)) {
                    New-Item -Path $basicDisplayPath -Force | Out-Null
                }
                Set-ItemProperty -Path $basicDisplayPath -Name 'Start' -Value 4 -Type DWord -Force
                Write-Log "[$VMName] [$cs] BasicDisplay service disabled (24H2/25H2 race mitigation)." "INFO"

                # TdrDelay: extend GPU TDR timeout to survive slow GPU-P adapter init
                $gfxDriversPath = "Registry::$hiveRoot\$cs\Control\GraphicsDrivers"
                if (-not (Test-Path $gfxDriversPath)) {
                    New-Item -Path $gfxDriversPath -Force | Out-Null
                }
                Set-ItemProperty -Path $gfxDriversPath -Name 'TdrDelay' -Value 60 -Type DWord -Force
                Write-Log "[$VMName] [$cs] TdrDelay extended to 60s (24H2/25H2 slow GPU-P init mitigation)." "INFO"
            }
        }

        Write-Log "[$VMName] Guest GPU-P registry mitigations applied successfully." "OK"
    } catch {
        Write-Log "[$VMName] Guest GPU-P registry mitigation error: $($_.Exception.Message)" "WARN"
    } finally {
        if ($loaded) {
            [GC]::Collect()
            Start-Sleep -Milliseconds 500
            & reg.exe unload $hiveRoot *> $null
            if ($LASTEXITCODE -ne 0) {
                Write-Log "[$VMName] Warning: guest SYSTEM hive unload returned exit $LASTEXITCODE. A reboot of the host may be needed to release the hive lock." "WARN"
            }
        }
    }
}

function Get-DriverVersionBranch {
    param([string]$VersionString)

    if ([string]::IsNullOrWhiteSpace($VersionString)) { return -1 }
    try {
        $segments = $VersionString.Split('.')
        # NVIDIA Windows driver versions use 4 segments: Major.Minor.Branch.Build
        # e.g. 32.0.15.7275 - the 3rd segment (index 2) is the meaningful branch.
        if ($segments.Count -ge 4) { return [int]$segments[2] }
        if ($segments.Count -ge 1) { return [int]$segments[0] }
        return -1
    } catch {
        return -1
    }
}

function Get-GpuPartitionValues {
    <#
    .SYNOPSIS
        Returns GPU partition resource-budget token values for Min/Max/Optimal.

        These values are NORMALISED RESOURCE-BUDGET TOKENS (0..1,000,000,000)
        — they are NOT physical VRAM byte counts.

        The reference implementation (bryanem32/hyperv_vm_creator v29) uses the
        proven-stable anchor:  Min = 80,000,000   Max = Optimal = 100,000,000
        at 100% allocation. We scale those anchors linearly with the UI slider so
        the user retains control while staying within the safe operating range.

        Reading adapter.MaxPartitionVRAM after Add-VMGpuPartitionAdapter returns
        the raw hardware VRAM in bytes (e.g. 8,589,934,592 for an 8 GB card) —
        using those bytes as budget tokens produces values 80x too large, causing
        dxgkrnl.sys to attempt an un-backed VRAM mapping that deadlocks on boot.
    #>
    param(
        [string]$VMName,           # reserved for future per-VM query; not used
        [ValidateRange(10,100)]
        [int]$Percentage = 100
    )

    $factor = [Math]::Max(0.1, [Math]::Min(1.0, $Percentage / 100.0))

    [UInt64]$maxVal = [UInt64][Math]::Floor(100000000 * $factor)
    [UInt64]$minVal = [UInt64][Math]::Floor( 80000000 * $factor)
    if ($minVal -lt 1)           { $minVal = 1 }
    if ($maxVal -lt $minVal)     { $maxVal = $minVal }

    $entry = @{ Supported = $true; Min = $minVal; Max = $maxVal; Optimal = $maxVal }
    return @{
        VRAM    = $entry
        Encode  = $entry
        Decode  = $entry
        Compute = $entry
    }
}

function Copy-GpuServiceDriver {
    <#
    .SYNOPSIS
        Copies the GPU kernel-mode service driver directory to the VM's
        HostDriverStore (Diobyte Version 1 method).
    #>
    [CmdletBinding()]
    param(
        [string]$MountLetter,
        [string]$GPUName = "AUTO"
    )
    try {
        $gpu = $null
        if ($GPUName -eq "AUTO") {
            $partList = Get-CimInstance -Namespace "ROOT\virtualization\v2" -ClassName "Msvm_PartitionableGpu" -ErrorAction SilentlyContinue
            if ($partList) {
                $devPath = if ($partList.Name -is [array]) { $partList.Name[0] } else { $partList.Name }
                # Extract PCI bus/device/function segment dynamically from the device path
                # Typical format: PCIP\VEN_XXXX&DEV_XXXX&... - extract the VEN&DEV segment after the first '#'
                $pciSegments = $devPath -split '#'
                $matchPattern = if ($pciSegments.Count -gt 1 -and $pciSegments[1].Length -ge 8) {
                    "*$($pciSegments[1].Substring(0, [Math]::Min($pciSegments[1].Length, 21)))*"
                } elseif ($devPath.Length -ge 24) {
                    "*$($devPath.Substring(8, 16))*"   # Legacy fallback
                } else {
                    $null
                }
                if ($matchPattern) {
                    $gpu = Get-PnpDevice -Class Display -Status OK -ErrorAction SilentlyContinue | Where-Object {
                        $_.DeviceID -like $matchPattern
                    } | Select-Object -First 1
                }
            }
        } else {
            $gpu = Get-PnpDevice -Class Display -Status OK -ErrorAction SilentlyContinue | Where-Object { $_.FriendlyName -eq $GPUName } | Select-Object -First 1
        }

        if (-not $gpu) {
            Write-Log "Could not resolve GPU PnP device for service driver copy" "WARN"
            return
        }

        $svcName = $gpu.Service
        if (-not $svcName) { return }

        $sysDriver = Get-CimInstance Win32_SystemDriver -Filter "Name='$svcName'" -ErrorAction SilentlyContinue
        if ($sysDriver -and $sysDriver.PathName) {
            # Strip common Win32_SystemDriver PathName prefixes like \??\
            $servicePath = $sysDriver.PathName -replace '^\\\?\?\\', ''
            # Robustly extract the driver folder: find the DriverStore\FileRepository segment
            $segments = $servicePath.Split('\\')
            $dsIdx = -1
            for ($si = 0; $si -lt $segments.Count; $si++) {
                if ($segments[$si] -eq 'FileRepository') { $dsIdx = $si; break }
            }
            if ($dsIdx -ge 0 -and ($dsIdx + 1) -lt $segments.Count) {
                $serviceDriverDir  = ($segments[0..($dsIdx + 1)]) -join '\'
                $relPath = ($segments[1..($dsIdx + 1)]) -join '\'
                $serviceDriverDest = Join-Path "$MountLetter\" ($relPath.Replace('DriverStore','HostDriverStore'))
            } else {
                # Fallback: use the parent folder of the driver binary
                $serviceDriverDir  = Split-Path -Parent $servicePath
                $relPath = $serviceDriverDir.Substring($serviceDriverDir.IndexOf('\') + 1)
                $serviceDriverDest = ("$MountLetter\" + $relPath).Replace('DriverStore','HostDriverStore')
            }

            if (Test-Path $serviceDriverDir) {
                if (-not (Test-Path $serviceDriverDest)) {
                    Write-Log "Copying GPU service driver: $svcName"
                    Copy-Item -Path $serviceDriverDir -Destination $serviceDriverDest -Recurse -Force -ErrorAction SilentlyContinue
                }
            }
        }
    } catch {
        Write-Log "GPU service driver copy error: $($_.Exception.Message)" "WARN"
    }
}

function Get-GpuDriverStoreFolderNamePatterns {
    param([string]$GpuVendor = "Auto")

    if ($GpuVendor -eq "NVIDIA") {
        return @('nv_*', 'nvhd*', 'nvlt*', 'nvmd*', 'nvra*', 'nvsr*', 'nvwm*', 'nvam*')
    }
    if ($GpuVendor -eq "AMD") {
        return @('u0*', 'c0*', 'amd*', 'ati*')
    }
    if ($GpuVendor -eq "Intel") {
        return @('igfx*', 'iigd*', 'cui_*', 'dch_*', 'kit_d*')
    }
    return @('nv_*', 'nvhd*', 'nvlt*', 'nvmd*', 'nvra*', 'nvsr*', 'nvwm*', 'nvam*',
             'u0*', 'c0*', 'amd*', 'ati*',
             'igfx*', 'iigd*', 'cui_*', 'dch_*', 'kit_d*')
}

function Remove-GuestGpuInjectedFilesFromManifest {
    [CmdletBinding()]
    param(
        [string]$VMName,
        [string]$MountLetter
    )

    $manifestPath = Join-Path "$MountLetter\" 'Windows\System32\HostDriverStore\GpuPvToolkit\InjectedFileManifest.txt'
    if (-not (Test-Path $manifestPath)) {
        return @{ Success = $true; Removed = 0 }
    }

    try {
        $mountRoot = [System.IO.Path]::GetFullPath((Join-Path "$MountLetter\" '.'))
        $mountRootLower = $mountRoot.ToLowerInvariant()
        $entries = @(Get-Content -Path $manifestPath -ErrorAction SilentlyContinue | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
        $removed = 0

        foreach ($entry in $entries) {
            $relative = [string]$entry
            if ($relative.StartsWith('\\') -or $relative -match '^[A-Za-z]:') { continue }

            $target = Join-Path "$MountLetter\" $relative
            $targetFull = [System.IO.Path]::GetFullPath($target)
            if (-not $targetFull.ToLowerInvariant().StartsWith($mountRootLower)) { continue }

            if (Test-Path $targetFull) {
                try {
                    Remove-Item -Path $targetFull -Force -ErrorAction Stop
                    $removed++
                } catch {
                    Write-Log "[$VMName] Failed removing prior injected file '$relative': $($_.Exception.Message)" "WARN"
                }
            }
        }

        try {
            Remove-Item -Path $manifestPath -Force -ErrorAction SilentlyContinue
        } catch {
            Write-Log "[$VMName] Could not remove old GPU injection manifest: $($_.Exception.Message)" "WARN"
        }

        if ($removed -gt 0) {
            Write-Log "[$VMName] Removed $removed previously injected GPU file(s) from manifest." "INFO"
        }
        return @{ Success = $true; Removed = $removed }
    } catch {
        Write-Log "[$VMName] Manifest cleanup error: $($_.Exception.Message)" "WARN"
        return @{ Success = $false; Removed = 0 }
    }
}

function Save-GuestGpuInjectedFilesManifest {
    [CmdletBinding()]
    param(
        [string]$VMName,
        [string]$MountLetter,
        [string[]]$RelativePaths
    )

    try {
        $manifestDir = Join-Path "$MountLetter\" 'Windows\System32\HostDriverStore\GpuPvToolkit'
        if (-not (Test-Path $manifestDir)) {
            New-Item -Path $manifestDir -ItemType Directory -Force | Out-Null
        }

        $manifestPath = Join-Path $manifestDir 'InjectedFileManifest.txt'
        $items = @($RelativePaths | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Sort-Object -Unique)
        if ($items.Count -eq 0) {
            if (Test-Path $manifestPath) {
                Remove-Item -Path $manifestPath -Force -ErrorAction SilentlyContinue
            }
            return
        }

        Set-Content -Path $manifestPath -Value $items -Encoding UTF8 -Force
        Write-Log "[$VMName] Saved GPU injection manifest with $($items.Count) file path(s)." "INFO"
    } catch {
        Write-Log "[$VMName] Failed to save GPU injection manifest: $($_.Exception.Message)" "WARN"
    }
}

function Remove-GuestGpuDriverPayload {
    [CmdletBinding()]
    param(
        [string]$VMName,
        [string]$MountLetter,
        [string]$GpuVendor = "Auto"
    )

    $repoPath = Join-Path "$MountLetter\" 'Windows\System32\HostDriverStore\FileRepository'
    $removedFolders = 0

    $manifestCleanup = Remove-GuestGpuInjectedFilesFromManifest -VMName $VMName -MountLetter $MountLetter
    if (-not $manifestCleanup.Success) {
        Write-Log "[$VMName] Proceeding despite manifest cleanup warnings." "WARN"
    }

    if (Test-Path $repoPath) {
        $patterns = Get-GpuDriverStoreFolderNamePatterns -GpuVendor $GpuVendor
        $targets = @()
        foreach ($pattern in $patterns) {
            $targets += @(Get-ChildItem -Path $repoPath -Directory -Filter $pattern -ErrorAction SilentlyContinue)
        }
        $targets = @($targets | Sort-Object FullName -Unique)

        foreach ($folder in $targets) {
            try {
                Remove-Item -Path $folder.FullName -Recurse -Force -ErrorAction Stop
                $removedFolders++
            } catch {
                Write-Log "[$VMName] Failed removing old guest GPU folder '$($folder.Name)': $($_.Exception.Message)" "WARN"
            }
        }
    }

    Write-Log "[$VMName] Guest GPU payload cleanup complete (folders removed: $removedFolders; manifest files removed: $($manifestCleanup.Removed))." "INFO"
    return @{ Success = $true; RemovedFolders = $removedFolders; RemovedFiles = $manifestCleanup.Removed }
}

function Copy-GpuReferencedFiles {
    [CmdletBinding()]
    param(
        [string]$MountLetter,
        [string]$GPUName = "AUTO"
    )

    $copied = 0
    $copiedRelativePaths = [System.Collections.Generic.List[string]]::new()
    try {
        if ([string]::IsNullOrWhiteSpace($MountLetter)) {
            return @{ Success = $false; Copied = 0 }
        }

        $gpu = $null
        if ($GPUName -eq "AUTO") {
            $partList = Get-CimInstance -Namespace "ROOT\virtualization\v2" -ClassName "Msvm_PartitionableGpu" -ErrorAction SilentlyContinue
            if ($partList) {
                $devPath = if ($partList.Name -is [array]) { $partList.Name[0] } else { $partList.Name }
                $pciSegments = $devPath -split '#'
                $matchPattern = if ($pciSegments.Count -gt 1 -and $pciSegments[1].Length -ge 8) {
                    "*$($pciSegments[1].Substring(0, [Math]::Min($pciSegments[1].Length, 21)))*"
                } elseif ($devPath.Length -ge 24) {
                    "*$($devPath.Substring(8, 16))*"
                } else {
                    $null
                }
                if ($matchPattern) {
                    $gpu = Get-PnpDevice -Class Display -Status OK -ErrorAction SilentlyContinue | Where-Object {
                        $_.DeviceID -like $matchPattern
                    } | Select-Object -First 1
                }
            }
        } else {
            $gpu = Get-PnpDevice -Class Display -Status OK -ErrorAction SilentlyContinue | Where-Object { $_.FriendlyName -eq $GPUName } | Select-Object -First 1
        }

        if (-not $gpu) {
            Write-Log "Could not resolve GPU PnP device for referenced file copy" "WARN"
            return @{ Success = $false; Copied = 0 }
        }

        $driverEntries = Get-CimInstance Win32_PnPSignedDriver -ErrorAction SilentlyContinue | Where-Object {
            $_.DeviceClass -eq 'DISPLAY' -and (
                ($gpu.InstanceId -and $_.DeviceID -eq $gpu.InstanceId) -or
                ($gpu.FriendlyName -and $_.DeviceName -eq $gpu.FriendlyName)
            )
        }

        if (-not $driverEntries) {
            Write-Log "Could not resolve Win32_PnPSignedDriver entries for $($gpu.FriendlyName). Skipping referenced-file copy and relying on HostDriverStore + service driver copy." "WARN"
            return @{ Success = $true; Copied = 0 }
        }

        $hostWindowsRoot = [System.IO.Path]::GetFullPath($env:WINDIR)
        $hostWindowsRootLower = $hostWindowsRoot.ToLowerInvariant()
        $hostDriverStoreRootLower = (Join-Path $hostWindowsRoot 'System32\DriverStore').ToLowerInvariant()

        foreach ($drv in $driverEntries) {
            $files = @()
            try {
                $files = @(Get-CimAssociatedInstance -InputObject $drv -ResultClassName CIM_DataFile -ErrorAction SilentlyContinue)
            } catch {
                $files = @()
            }

            foreach ($file in $files) {
                $srcPath = [string]$file.Name
                if ([string]::IsNullOrWhiteSpace($srcPath) -or -not (Test-Path $srcPath)) { continue }

                $srcFull = [System.IO.Path]::GetFullPath($srcPath)
                $srcLower = $srcFull.ToLowerInvariant()
                if (-not $srcLower.StartsWith($hostWindowsRootLower)) { continue }

                $destPath = $null
                if ($srcLower.StartsWith($hostDriverStoreRootLower)) {
                    $relative = $srcFull.Substring($hostWindowsRoot.Length).TrimStart('\')
                    if ($relative -match '^(?i)System32\\DriverStore\\') {
                        $relative = $relative -replace '^(?i)System32\\DriverStore\\', 'System32\\HostDriverStore\\'
                    }
                    $destPath = Join-Path "$MountLetter\Windows" $relative
                } else {
                    # Non-DriverStore files (e.g. System32 user-mode DLLs) must NOT be copied into
                    # the guest Windows directory. Transplanting host-version DLLs (OpenGL ICDs,
                    # CUDA libraries, etc.) overwrites guest system files and causes guest OS freezes.
                    Write-Log "Skipping non-DriverStore referenced GPU file (safe for guest stability): $srcFull" "INFO"
                    continue
                }

                $destDir = Split-Path -Parent $destPath
                if (-not (Test-Path $destDir)) {
                    New-Item -Path $destDir -ItemType Directory -Force | Out-Null
                }

                try {
                    Copy-Item -Path $srcFull -Destination $destPath -Force -ErrorAction Stop
                    $copied++
                    $destFull = [System.IO.Path]::GetFullPath($destPath)
                    $relativeGuest = $destFull.Substring((Join-Path "$MountLetter\" '').Length).TrimStart('\\')
                    if (-not [string]::IsNullOrWhiteSpace($relativeGuest)) {
                        $copiedRelativePaths.Add($relativeGuest)
                    }
                } catch {
                    Write-Log "Failed to copy referenced GPU file '$srcFull': $($_.Exception.Message)" "WARN"
                }
            }
        }

        if ($copied -gt 0) {
            Write-Log "Copied $copied referenced GPU driver file(s) to guest Windows paths." "OK"
        } else {
            Write-Log "No referenced GPU files were copied from Win32_PnPSignedDriver associations. Continuing with HostDriverStore and service driver payloads." "WARN"
        }

        return @{ Success = $true; Copied = $copied; Paths = @($copiedRelativePaths | Sort-Object -Unique) }
    } catch {
        Write-Log "GPU referenced file copy error: $($_.Exception.Message)" "WARN"
        return @{ Success = $false; Copied = $copied; Paths = @($copiedRelativePaths | Sort-Object -Unique) }
    }
}

#endregion

#region ==================== WINDOWS DETECTION ====================

function Get-WimVersionInfo {
    <#
    .SYNOPSIS
        Parses DISM WIM info to detect Windows version and build number.
    .OUTPUTS
        PSCustomObject with WinVersion ("Windows 10"/"Windows 11") and Build (int)
    #>
    param([string]$WimFile, [int]$Index = 1)

    $result = [PSCustomObject]@{ WinVersion = "Unknown"; Build = 0; Architecture = "" }

    try {
        $wimInfo = dism /Get-WimInfo /WimFile:"$WimFile" /Index:$Index /English 2>&1
        $dismExit = $LASTEXITCODE
        if ($dismExit -ne 0) {
            Write-Log "DISM /Get-WimInfo returned exit code $dismExit for index $Index of '$WimFile'" "WARN"
        }
        foreach ($line in $wimInfo) {
            $line = $line.ToString().Trim()
            if ($line -match '^Version\s*:\s*(\d+\.\d+\.(\d+))') {
                $result.Build = [int]$matches[2]
            }
            if ($line -match '^Name\s*:\s*(.+)$') {
                $name = $matches[1].Trim()
                if ($name -match 'Windows 11') { $result.WinVersion = "Windows 11" }
                elseif ($name -match 'Windows 10') { $result.WinVersion = "Windows 10" }
                elseif ($name -match 'Windows Server') { $result.WinVersion = "Windows Server" }
            }
            if ($line -match '^Architecture\s*:\s*(.+)$') {
                $archRaw = $matches[1].Trim()
                $normalizedArch = ConvertTo-UnattendArchitecture -Architecture $archRaw
                if ($normalizedArch) {
                    $result.Architecture = $normalizedArch
                }
            }
        }
        # Fallback: build >= 22000 is Windows 11
        if ($result.WinVersion -eq "Unknown" -and $result.Build -ge 22000) {
            $result.WinVersion = "Windows 11"
        } elseif ($result.WinVersion -eq "Unknown" -and $result.Build -ge 10240) {
            $result.WinVersion = "Windows 10"
        }
    } catch {
        Write-Log "Failed to detect Windows version from WIM: $_" "WARN"
    }
    return $result
}

function Resolve-GuestWindowsProfile {
    param(
        [string]$DetectedWinVersion,
        [int]$DetectedBuild
    )

    $isWin11 = ($DetectedWinVersion -eq 'Windows 11' -or $DetectedBuild -ge $script:BUILD_WIN11_MIN)
    $isWin10 = (-not $isWin11) -and ($DetectedWinVersion -eq 'Windows 10' -or ($DetectedBuild -ge $script:BUILD_WIN10_MIN -and $DetectedBuild -lt $script:BUILD_WIN11_MIN))
    $isLegacyWin10 = $isWin10 -and $DetectedBuild -gt 0 -and $DetectedBuild -lt $script:BUILD_WIN10_RS4

    $name = if ($isWin11) {
        'Windows 11'
    } elseif ($isWin10) {
        'Windows 10'
    } else {
        if ([string]::IsNullOrWhiteSpace($DetectedWinVersion)) { 'Unknown' } else { $DetectedWinVersion }
    }

    $defaultSecureBoot = $false
    $defaultTPM = $false
    if ($isWin11) {
        $defaultSecureBoot = $true
        $defaultTPM = $true
    } elseif ($isWin10 -and $DetectedBuild -ge $script:BUILD_WIN10_RS5) {
        $defaultSecureBoot = $true
    }

    $compatibilityNote = if ($isLegacyWin10) {
        'Legacy Windows 10 detected (pre-1803): DISM apply will prefer non-compact mode and legacy-compatible secure boot template order.'
    } elseif ($isWin10) {
        'Windows 10 detected: Secure Boot recommended, TPM optional.'
    } elseif ($isWin11) {
        'Windows 11 detected: Secure Boot and TPM are required.'
    } else {
        'Unknown/other guest detected: conservative defaults applied; review Secure Boot and TPM manually.'
    }

    [PSCustomObject]@{
        Name                    = $name
        Build                   = $DetectedBuild
        IsWindows11             = $isWin11
        IsWindows10             = $isWin10
        IsLegacyWindows10       = $isLegacyWin10
        PreferCompactApply      = (-not $isLegacyWin10)
        RequireSecureBoot       = $isWin11
        RequireTPM              = $isWin11
        DefaultSecureBoot       = $defaultSecureBoot
        DefaultTPM              = $defaultTPM
        SecureBootTemplateOrder = if ($isWin11 -or $isLegacyWin10) {
            @('MicrosoftWindows', 'MicrosoftUEFICertificateAuthority')
        } else {
            @('MicrosoftWindows', 'MicrosoftUEFICertificateAuthority')
        }
        CompatibilityNote       = $compatibilityNote
    }
}

function Set-DetectedGuestDefaults {
    param(
        [Parameter(Mandatory = $true)][hashtable]$Controls,
        [Parameter(Mandatory = $true)]$Profile,
        [switch]$EmitLog
    )

    if ($Controls.ContainsKey('OSInfo') -and $Controls['OSInfo']) {
        $guestArchLabel = if ([string]::IsNullOrWhiteSpace($script:DetectedGuestArch)) { 'unknown' } else { $script:DetectedGuestArch }
        $Controls['OSInfo'].Text = "$($Profile.Name)  (Build $($Profile.Build), $guestArchLabel)"
    }
    if ($Controls.ContainsKey('SecureBoot') -and $Controls['SecureBoot']) {
        $Controls['SecureBoot'].Checked = [bool]$Profile.DefaultSecureBoot
    }
    if ($Controls.ContainsKey('TPM') -and $Controls['TPM']) {
        $Controls['TPM'].Checked = [bool]$Profile.DefaultTPM
    }

    if ($EmitLog) {
        Write-Log $Profile.CompatibilityNote "INFO"
    }
}

function Invoke-DismApplyImage {
    <#
    .SYNOPSIS
        Applies a WIM/ESD image using DISM with timeout protection.
        Falls back from /Compact to normal mode on failure.
    #>
    [CmdletBinding()]
    param(
        [string]$ImageFile,
        [int]$Index,
        [string]$ApplyDir,
        [bool]$PreferCompactApply = $true,
        [int]$TimeoutMinutes = 60
    )

    $attempts = @()
    if ($PreferCompactApply) {
        $attempts += [PSCustomObject]@{ Label = 'with /Compact'; Args = @('/Apply-Image', "/ImageFile:$ImageFile", "/Index:$Index", "/ApplyDir:$ApplyDir", '/Compact') }
    }
    $attempts += [PSCustomObject]@{ Label = 'without /Compact'; Args = @('/Apply-Image', "/ImageFile:$ImageFile", "/Index:$Index", "/ApplyDir:$ApplyDir") }

    foreach ($attempt in $attempts) {
        Write-Log "DISM apply attempt: $($attempt.Label)"
        $dismArgs = $attempt.Args
        $dismArgsForJob = @($dismArgs)

        # Run DISM in a background job with timeout to prevent indefinite hangs.
        # Each output line is written individually so Receive-Job can stream progress.
        $dismJob = Start-Job -ScriptBlock {
            & dism @Using:dismArgsForJob 2>&1 | ForEach-Object { Write-Output $_ }
            Write-Output "__DISM_EXIT__:$LASTEXITCODE"
        }

        # Poll the job for incremental progress instead of blocking on Wait-Job
        $timeoutSec = $TimeoutMinutes * 60
        $elapsed = 0
        $pollInterval = 1  # seconds
        $lastPct = -1
        $allOutput = [System.Collections.Generic.List[string]]::new()
        $dismExitCode = $null

        while ($dismJob.State -eq 'Running' -and $elapsed -lt $timeoutSec) {
            Start-Sleep -Seconds $pollInterval
            $elapsed += $pollInterval

            # Read partial output (DISM progress lines) as they stream
            try {
                $partial = @(Receive-Job $dismJob -ErrorAction SilentlyContinue)
                foreach ($line in $partial) {
                    $lineStr = "$line".Trim()
                    if ($lineStr -match '^__DISM_EXIT__:(-?\d+)$') {
                        $dismExitCode = [int]$Matches[1]
                        continue
                    }
                    $allOutput.Add($lineStr)
                    if ($lineStr -match '(\d+)\.\d+%') {
                        $pct = [int]$Matches[1]
                        if ($pct -ne $lastPct -and ($pct % 10 -eq 0 -or $pct -ge 99)) {
                            Write-Log "  DISM progress: ${pct}%"
                            $lastPct = $pct
                        }
                    }
                }
            } catch {
                Write-Log "DISM output polling warning: $($_.Exception.Message)" "WARN"
            }
        }

        if ($dismJob.State -eq 'Running') {
            Write-Log "DISM timed out after $TimeoutMinutes minutes ($($attempt.Label))" "ERROR"
            Stop-Job $dismJob -ErrorAction SilentlyContinue
            Remove-Job $dismJob -Force -ErrorAction SilentlyContinue

            # Clean up partially applied image to avoid corrupt state on retry.
            # Guard aggressively against destructive deletions on non-target/system roots.
            if (Test-Path $ApplyDir) {
                $applyRoot = [System.IO.Path]::GetPathRoot($ApplyDir)
                $isDriveRoot = ($applyRoot -and (($ApplyDir.TrimEnd('\\') + '\\') -ieq $applyRoot))
                $windowsDir = Join-Path $ApplyDir 'Windows'
                $efiMarker = Join-Path $ApplyDir 'EFI'
                $bootMarker = Join-Path $ApplyDir 'boot'
                $safeCleanupRoot = ($isDriveRoot -and (Test-Path $windowsDir) -and ((Test-Path $efiMarker) -or (Test-Path $bootMarker)))
                $systemRoot = [System.IO.Path]::GetPathRoot($env:SystemRoot)

                if ($safeCleanupRoot -and ($applyRoot -ine $systemRoot)) {
                    Write-Log "Cleaning up known Windows install paths in $ApplyDir before retry..." "WARN"
                    foreach ($rel in @('Windows','Program Files','Program Files (x86)','ProgramData','Users','PerfLogs','Recovery','boot','Boot','EFI')) {
                        $target = Join-Path $ApplyDir $rel
                        if (Test-Path $target) {
                            Remove-Item -Path $target -Recurse -Force -ErrorAction SilentlyContinue
                        }
                    }
                    Write-Log "Partial image cleanup complete." "WARN"
                } else {
                    Write-Log "Skipping destructive cleanup for '$ApplyDir' because safety checks did not pass." "WARN"
                }
            }
            continue
        }

        # Drain any remaining output after job completes
        try {
            $remaining = @(Receive-Job $dismJob -ErrorAction SilentlyContinue)
            foreach ($line in $remaining) {
                $lineStr = "$line".Trim()
                if ($lineStr -match '^__DISM_EXIT__:(-?\d+)$') {
                    $dismExitCode = [int]$Matches[1]
                    continue
                }
                $allOutput.Add($lineStr)
            }
        } catch {
            Write-Log "DISM output finalization warning: $($_.Exception.Message)" "WARN"
        }
        Remove-Job $dismJob -Force -ErrorAction SilentlyContinue

        $dismOutput = $allOutput
        $exitCode = if ($null -ne $dismExitCode) { $dismExitCode } else { 1 }

        # Log non-progress DISM output lines (progress was already logged during polling)
        foreach ($line in $dismOutput) {
            $lineStr = "$line".Trim()
            if ($lineStr -and $lineStr -notmatch '^\s*$' -and $lineStr -notmatch '\d+\.\d+%' -and $lineStr -notmatch '^Deployment Image|^Version:') {
                Write-Log "  DISM: $lineStr"
            }
        }

        if ($exitCode -eq 0) {
            return
        }

        Write-Log "DISM apply attempt failed ($($attempt.Label)) with exit code $exitCode" "WARN"
    }

    throw "DISM /Apply-Image failed for all compatibility attempts."
}

#endregion

#region ==================== UNATTEND XML ====================

function New-UnattendXml {
    param(
        [string]$VMName,
        [string]$Username,
        [AllowNull()][System.Security.SecureString]$Password,
        [bool]$EnableAutoLogon = $true,
        [bool]$IsWindows11 = $false,
        [AllowNull()][string]$GuestArch = ""
    )

    $passwordPlain = Convert-SecureStringToPlainText -SecureString $Password

    # Culture / Locale Detection
    if (Get-Command Get-Culture -ErrorAction SilentlyContinue) {
        $culture = Get-Culture
        $uiLang  = $culture.Name
    } else {
        $culture = [System.Globalization.CultureInfo]::CurrentCulture
        $uiLang  = [System.Globalization.CultureInfo]::CurrentUICulture.Name
    }
    $systemLoc = $culture.Name
    $userLoc   = $culture.Name

    # Fallback if culture values are empty (e.g. invariant culture, minimal OS image)
    if ([string]::IsNullOrWhiteSpace($uiLang))    { $uiLang   = "en-US" }
    if ([string]::IsNullOrWhiteSpace($systemLoc))  { $systemLoc = "en-US" }
    if ([string]::IsNullOrWhiteSpace($userLoc))    { $userLoc   = "en-US" }

    # Keyboard Detection
    try { $keyboard = (Get-WinUserLanguageList)[0].InputMethodTips[0] }
    catch {
        # Fallback: derive keyboard layout from current culture instead of hard-coding US-English
        try {
            $kbLayoutId = [System.Globalization.CultureInfo]::CurrentCulture.KeyboardLayoutId
            $hexLayout  = '{0:X4}:{0:X8}' -f $kbLayoutId
            $keyboard   = $hexLayout
            Write-Log "Keyboard fallback derived from CurrentCulture: $keyboard" "WARN"
        } catch {
            $keyboard = "0409:00000409"  # ultimate fallback: US-English
        }
    }

    # Timezone Detection
    try { $timezone = (Get-TimeZone).Id }
    catch {
        try { $timezone = (Get-CimInstance Win32_TimeZone).StandardName }
        catch { $timezone = "UTC" }
    }

    # Truncate computer name from the raw VM name before XML-escaping to avoid
    # splitting XML entities (e.g. &amp;) that would occur if truncating the escaped form.
    $rawCn = $VMName
    if ($rawCn.Length -gt 15) { $rawCn = $rawCn.Substring(0,15).TrimEnd('-') }
    if ([string]::IsNullOrWhiteSpace($rawCn)) { $rawCn = $VMName -replace '[^a-zA-Z0-9]','' }
    if ([string]::IsNullOrWhiteSpace($rawCn)) { $rawCn = 'VM' }
    $xmlComputerName = ConvertTo-XmlEscapedValue -Value $rawCn
    $xmlUsername     = ConvertTo-XmlEscapedValue -Value $Username
    $xmlPasswordB64  = ConvertTo-UnattendPassword -PlainText $passwordPlain
    $passwordPlain   = $null   # Minimize plaintext password lifetime in memory
    $xmlUiLang       = ConvertTo-XmlEscapedValue -Value $uiLang
    $xmlKeyboard     = ConvertTo-XmlEscapedValue -Value $keyboard
    $xmlSystemLoc    = ConvertTo-XmlEscapedValue -Value $systemLoc
    $xmlUserLoc      = ConvertTo-XmlEscapedValue -Value $userLoc
    $xmlTimezone     = ConvertTo-XmlEscapedValue -Value $timezone
    $xmlArch = ConvertTo-UnattendArchitecture -Architecture $GuestArch
    if (-not $xmlArch) {
        $xmlArch = ConvertTo-UnattendArchitecture -Architecture $script:DetectedGuestArch
    }
    if (-not $xmlArch) {
        $xmlArch = $script:HostArch
        Write-Log "Unattend architecture: using host architecture fallback '$xmlArch'" "WARN"
    }

    # ---- Build specialize Deployment RunSynchronous commands ----
    # These run before OOBE during first-boot mini-setup (specialize pass)
    $specDeployCmds = [System.Collections.Generic.List[string]]::new()

    # Win11: bypass OOBE network requirement for offline local-account setup
    if ($IsWindows11) {
        $specDeployCmds.Add('reg.exe add "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\OOBE" /v BypassNRO /t REG_DWORD /d 1 /f')
    }

    # ContentDeliveryManager debloat: load default user hive, suppress ads/bloatware for all future profiles
    $specDeployCmds.Add('reg.exe load "HKU\mount" "C:\Users\Default\NTUSER.DAT"')
    $cdmBase = 'HKU\mount\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager'
    @('ContentDeliveryAllowed','FeatureManagementEnabled','OEMPreInstalledAppsEnabled',
      'PreInstalledAppsEnabled','PreInstalledAppsEverEnabled','SilentInstalledAppsEnabled',
      'SoftLandingEnabled','SubscribedContentEnabled','SubscribedContent-310093Enabled',
      'SubscribedContent-338387Enabled','SubscribedContent-338388Enabled',
      'SubscribedContent-338389Enabled','SubscribedContent-338393Enabled',
      'SubscribedContent-353698Enabled','SystemPaneSuggestionsEnabled') | ForEach-Object {
        $specDeployCmds.Add("reg.exe add `"$cdmBase`" /v `"$_`" /t REG_DWORD /d 0 /f")
    }
    $ucBase = 'HKU\mount\Software\Policies\Microsoft\Windows\CloudContent'
    @('DisableCloudOptimizedContent','DisableWindowsConsumerFeatures','DisableConsumerAccountStateContent') | ForEach-Object {
        $specDeployCmds.Add("reg.exe add `"$ucBase`" /v `"$_`" /t REG_DWORD /d 1 /f")
    }
    $specDeployCmds.Add('reg.exe unload "HKU\mount"')

    # Machine-level CloudContent policies
    $mcBase = 'HKLM\Software\Policies\Microsoft\Windows\CloudContent'
    @('DisableCloudOptimizedContent','DisableWindowsConsumerFeatures','DisableConsumerAccountStateContent') | ForEach-Object {
        $specDeployCmds.Add("reg.exe add `"$mcBase`" /v `"$_`" /t REG_DWORD /d 1 /f")
    }

    # Set network location to Home (enables discovery by default)
    $specDeployCmds.Add('reg.exe add "HKLM\SOFTWARE\Policies\Microsoft\Windows NT\CurrentVersion\NetworkList\Signatures\FirstNetwork" /v Category /t REG_DWORD /d 1 /f')

    # Build XML for specialize RunSynchronous commands
    $specRunSyncParts = @()
    for ($i = 0; $i -lt $specDeployCmds.Count; $i++) {
        $order = $i + 1
        $cmd = [System.Security.SecurityElement]::Escape($specDeployCmds[$i])
        $specRunSyncParts += @"
        <RunSynchronousCommand wcm:action="add">
          <Order>$order</Order>
          <Path>$cmd</Path>
        </RunSynchronousCommand>
"@
    }
    $specRunSyncXml = $specRunSyncParts -join "`n"

    # Win11 windowsPE: LabConfig bypasses (belt-and-suspenders for requirement checks)
    $windowsPeRunSync = ""
    if ($IsWindows11) {
        $windowsPeRunSync = @"

      <RunSynchronous>
        <RunSynchronousCommand wcm:action="add">
          <Order>1</Order>
          <Path>reg.exe add "HKLM\SYSTEM\Setup\LabConfig" /v BypassTPMCheck /t REG_DWORD /d 1 /f</Path>
        </RunSynchronousCommand>
        <RunSynchronousCommand wcm:action="add">
          <Order>2</Order>
          <Path>reg.exe add "HKLM\SYSTEM\Setup\LabConfig" /v BypassSecureBootCheck /t REG_DWORD /d 1 /f</Path>
        </RunSynchronousCommand>
        <RunSynchronousCommand wcm:action="add">
          <Order>3</Order>
          <Path>reg.exe add "HKLM\SYSTEM\Setup\LabConfig" /v BypassRAMCheck /t REG_DWORD /d 1 /f</Path>
        </RunSynchronousCommand>
        <RunSynchronousCommand wcm:action="add">
          <Order>4</Order>
          <Path>reg.exe add "HKLM\SYSTEM\Setup\MoSetup" /v AllowUpgradesWithUnsupportedTPMOrCPU /t REG_DWORD /d 1 /f</Path>
        </RunSynchronousCommand>
      </RunSynchronous>
"@
    }

    # AutoLogon block (high count for persistent auto-logon, matching dockur convention)
    $autoLogonBlock = ""
    if ($EnableAutoLogon) {
        $autoLogonBlock = @"
      <AutoLogon>
        <Username>$xmlUsername</Username>
        <Enabled>true</Enabled>
        <LogonCount>999</LogonCount>
        <Password>
          <Value>$xmlPasswordB64</Value>
          <PlainText>false</PlainText>
        </Password>
      </AutoLogon>
"@
    }

    return @"
<?xml version="1.0" encoding="utf-8"?>
<unattend xmlns="urn:schemas-microsoft-com:unattend" xmlns:wcm="http://schemas.microsoft.com/WMIConfig/2002/State">
  <settings pass="offlineServicing">
    <component name="Microsoft-Windows-LUA-Settings" processorArchitecture="$xmlArch" publicKeyToken="31bf3856ad364e35" language="neutral" versionScope="nonSxS">
            <EnableLUA>true</EnableLUA>
    </component>
  </settings>

  <settings pass="windowsPE">
    <component name="Microsoft-Windows-International-Core-WinPE" processorArchitecture="$xmlArch" publicKeyToken="31bf3856ad364e35" language="neutral" versionScope="nonSxS">
      <SetupUILanguage>
        <UILanguage>$xmlUiLang</UILanguage>
      </SetupUILanguage>
      <InputLocale>$xmlKeyboard</InputLocale>
      <SystemLocale>$xmlSystemLoc</SystemLocale>
      <UILanguage>$xmlUiLang</UILanguage>
      <UserLocale>$xmlUserLoc</UserLocale>
    </component>
    <component name="Microsoft-Windows-Setup" processorArchitecture="$xmlArch" publicKeyToken="31bf3856ad364e35" language="neutral" versionScope="nonSxS">
      <UserData>
        <AcceptEula>true</AcceptEula>
      </UserData>
      <EnableFirewall>true</EnableFirewall>
      <Diagnostics>
        <OptIn>false</OptIn>
      </Diagnostics>
      <UseConfigurationSet>false</UseConfigurationSet>$windowsPeRunSync
    </component>
  </settings>

  <settings pass="generalize">
    <component name="Microsoft-Windows-Security-SPP" processorArchitecture="$xmlArch" publicKeyToken="31bf3856ad364e35" language="neutral" versionScope="nonSxS">
      <SkipRearm>1</SkipRearm>
    </component>
  </settings>

  <settings pass="specialize">
    <component name="Microsoft-Windows-Shell-Setup" processorArchitecture="$xmlArch" publicKeyToken="31bf3856ad364e35" language="neutral" versionScope="nonSxS">
      <ComputerName>$xmlComputerName</ComputerName>
      <TimeZone>$xmlTimezone</TimeZone>
    </component>
    <component name="Microsoft-Windows-Security-SPP-UX" processorArchitecture="$xmlArch" publicKeyToken="31bf3856ad364e35" language="neutral" versionScope="nonSxS">
      <SkipAutoActivation>true</SkipAutoActivation>
    </component>
    <component name="Microsoft-Windows-ErrorReportingCore" processorArchitecture="$xmlArch" publicKeyToken="31bf3856ad364e35" language="neutral" versionScope="nonSxS">
      <DisableWER>1</DisableWER>
    </component>
    <component name="Microsoft-Windows-SQMApi" processorArchitecture="$xmlArch" publicKeyToken="31bf3856ad364e35" language="neutral" versionScope="nonSxS">
      <CEIPEnabled>0</CEIPEnabled>
    </component>
    <component name="Microsoft-Windows-SystemRestore-Main" processorArchitecture="$xmlArch" publicKeyToken="31bf3856ad364e35" language="neutral" versionScope="nonSxS">
      <DisableSR>1</DisableSR>
    </component>
    <component name="Microsoft-Windows-International-Core" processorArchitecture="$xmlArch" publicKeyToken="31bf3856ad364e35" language="neutral" versionScope="nonSxS">
      <InputLocale>$xmlKeyboard</InputLocale>
      <SystemLocale>$xmlSystemLoc</SystemLocale>
      <UILanguage>$xmlUiLang</UILanguage>
      <UserLocale>$xmlUserLoc</UserLocale>
    </component>
    <component name="Microsoft-Windows-Deployment" processorArchitecture="$xmlArch" publicKeyToken="31bf3856ad364e35" language="neutral" versionScope="nonSxS">
      <RunSynchronous>
$specRunSyncXml
      </RunSynchronous>
    </component>
  </settings>

  <settings pass="oobeSystem">
    <component name="Microsoft-Windows-SecureStartup-FilterDriver" processorArchitecture="$xmlArch" publicKeyToken="31bf3856ad364e35" language="neutral" versionScope="nonSxS">
      <PreventDeviceEncryption>true</PreventDeviceEncryption>
    </component>
    <component name="Microsoft-Windows-EnhancedStorage-Adm" processorArchitecture="$xmlArch" publicKeyToken="31bf3856ad364e35" language="neutral" versionScope="nonSxS">
      <TCGSecurityActivationDisabled>1</TCGSecurityActivationDisabled>
    </component>
    <component name="Microsoft-Windows-International-Core" processorArchitecture="$xmlArch" publicKeyToken="31bf3856ad364e35" language="neutral" versionScope="nonSxS">
      <InputLocale>$xmlKeyboard</InputLocale>
      <SystemLocale>$xmlSystemLoc</SystemLocale>
      <UILanguage>$xmlUiLang</UILanguage>
      <UserLocale>$xmlUserLoc</UserLocale>
    </component>
    <component name="Microsoft-Windows-Shell-Setup" processorArchitecture="$xmlArch" publicKeyToken="31bf3856ad364e35" language="neutral" versionScope="nonSxS">
      <UserAccounts>
        <LocalAccounts>
          <LocalAccount wcm:action="add">
            <Name>$xmlUsername</Name>
            <DisplayName>$xmlUsername</DisplayName>
            <Group>Administrators</Group>
            <Password>
              <Value>$xmlPasswordB64</Value>
              <PlainText>false</PlainText>
            </Password>
          </LocalAccount>
        </LocalAccounts>
      </UserAccounts>
$autoLogonBlock
      <Display>
        <ColorDepth>32</ColorDepth>
        <HorizontalResolution>1920</HorizontalResolution>
        <VerticalResolution>1080</VerticalResolution>
      </Display>
      <OOBE>
        <ProtectYourPC>3</ProtectYourPC>
        <HideEULAPage>true</HideEULAPage>
        <HideLocalAccountScreen>true</HideLocalAccountScreen>
        <HideOEMRegistrationScreen>true</HideOEMRegistrationScreen>
        <HideOnlineAccountScreens>true</HideOnlineAccountScreens>
        <HideWirelessSetupInOOBE>true</HideWirelessSetupInOOBE>
        <NetworkLocation>Home</NetworkLocation>
        <SkipUserOOBE>true</SkipUserOOBE>
        <SkipMachineOOBE>true</SkipMachineOOBE>
      </OOBE>
      <FirstLogonCommands>
        <SynchronousCommand wcm:action="add">
          <Order>1</Order>
          <CommandLine>powershell.exe -NoProfile -NonInteractive -Command "try { Set-LocalUser -Name '$xmlUsername' -PasswordNeverExpires 1 } catch { net.exe user '$xmlUsername' /expires:never }"</CommandLine>
          <Description>Disable Password Expiration</Description>
        </SynchronousCommand>
        <SynchronousCommand wcm:action="add">
          <Order>2</Order>
          <CommandLine>cmd /C POWERCFG -H OFF</CommandLine>
          <Description>Disable Hibernation</Description>
        </SynchronousCommand>
        <SynchronousCommand wcm:action="add">
          <Order>3</Order>
          <CommandLine>cmd /C POWERCFG -X -monitor-timeout-ac 0</CommandLine>
          <Description>Disable monitor blanking</Description>
        </SynchronousCommand>
        <SynchronousCommand wcm:action="add">
          <Order>4</Order>
          <CommandLine>cmd /C POWERCFG -X -standby-timeout-ac 0</CommandLine>
          <Description>Disable Sleep</Description>
        </SynchronousCommand>
        <SynchronousCommand wcm:action="add">
          <Order>5</Order>
          <CommandLine>reg.exe add "HKLM\SYSTEM\CurrentControlSet\Control\Power" /v "HibernateFileSizePercent" /t REG_DWORD /d 0 /f</CommandLine>
          <Description>Zero Hibernation File</Description>
        </SynchronousCommand>
        <SynchronousCommand wcm:action="add">
          <Order>6</Order>
          <CommandLine>reg.exe add "HKLM\SYSTEM\CurrentControlSet\Control\Power" /v "HibernateEnabled" /t REG_DWORD /d 0 /f</CommandLine>
          <Description>Disable Hibernation Registry</Description>
        </SynchronousCommand>
        <SynchronousCommand wcm:action="add">
          <Order>7</Order>
          <CommandLine>reg.exe add "HKLM\SOFTWARE\Policies\Microsoft\Edge" /v "HideFirstRunExperience" /t REG_DWORD /d 1 /f</CommandLine>
          <Description>Disable Edge first-run experience</Description>
        </SynchronousCommand>
        <SynchronousCommand wcm:action="add">
          <Order>8</Order>
          <CommandLine>reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced" /v "HideFileExt" /t REG_DWORD /d 0 /f</CommandLine>
          <Description>Show file extensions in Explorer</Description>
        </SynchronousCommand>
        <SynchronousCommand wcm:action="add">
          <Order>9</Order>
          <CommandLine>reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced" /v "ShowCopilotButton" /t REG_DWORD /d 0 /f</CommandLine>
          <Description>Hide Copilot button</Description>
        </SynchronousCommand>
        <SynchronousCommand wcm:action="add">
          <Order>10</Order>
          <CommandLine>reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced" /v "ShowTaskViewButton" /t REG_DWORD /d 0 /f</CommandLine>
          <Description>Remove Task View from Taskbar</Description>
        </SynchronousCommand>
        <SynchronousCommand wcm:action="add">
          <Order>11</Order>
          <CommandLine>reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced" /v "TaskbarDa" /t REG_DWORD /d 0 /f</CommandLine>
          <Description>Remove Widgets from Taskbar</Description>
        </SynchronousCommand>
        <SynchronousCommand wcm:action="add">
          <Order>12</Order>
          <CommandLine>reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced" /v "TaskbarMn" /t REG_DWORD /d 0 /f</CommandLine>
          <Description>Remove Chat from Taskbar</Description>
        </SynchronousCommand>
        <SynchronousCommand wcm:action="add">
          <Order>13</Order>
          <CommandLine>reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Search" /v "SearchboxTaskbarMode" /t REG_DWORD /d 0 /f</CommandLine>
          <Description>Remove Search from Taskbar</Description>
        </SynchronousCommand>
        <SynchronousCommand wcm:action="add">
          <Order>14</Order>
          <CommandLine>netsh advfirewall firewall set rule group="@FirewallAPI.dll,-32752" new enable=Yes</CommandLine>
          <Description>Enable Network Discovery</Description>
        </SynchronousCommand>
        <SynchronousCommand wcm:action="add">
          <Order>15</Order>
          <CommandLine>reg.exe add "HKLM\SYSTEM\CurrentControlSet\Control\Network\NewNetworkWindowOff" /f</CommandLine>
          <Description>Disable Network Discovery popup</Description>
        </SynchronousCommand>
        <SynchronousCommand wcm:action="add">
          <Order>16</Order>
          <CommandLine>netsh advfirewall firewall set rule group="@FirewallAPI.dll,-28502" new enable=Yes</CommandLine>
          <Description>Enable File Sharing</Description>
        </SynchronousCommand>
        <SynchronousCommand wcm:action="add">
          <Order>17</Order>
          <CommandLine>reg.exe add "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\PasswordLess\Device" /v "DevicePasswordLessBuildVersion" /t REG_DWORD /d 0 /f</CommandLine>
          <Description>Enable passwordless sign-in option</Description>
        </SynchronousCommand>
        <SynchronousCommand wcm:action="add">
          <Order>18</Order>
          <CommandLine>reg.exe add "HKCU\Control Panel\UnsupportedHardwareNotificationCache" /v SV1 /d 0 /t REG_DWORD /f</CommandLine>
          <Description>Disable unsupported hardware notification</Description>
        </SynchronousCommand>
        <SynchronousCommand wcm:action="add">
          <Order>19</Order>
          <CommandLine>reg.exe add "HKCU\Control Panel\UnsupportedHardwareNotificationCache" /v SV2 /d 0 /t REG_DWORD /f</CommandLine>
          <Description>Disable unsupported hardware notification</Description>
        </SynchronousCommand>
      </FirstLogonCommands>
    </component>
  </settings>
</unattend>
"@
}

#endregion

#region ==================== GPU FUNCTIONS ====================

function Get-GpuPProviders {
    $gpus = Get-VMHostPartitionableGpu -ErrorAction SilentlyContinue
    if (!$gpus) { return @() }

    $displayDevices = Get-PnpDevice -Class Display -Status OK -ErrorAction SilentlyContinue

    $list = @()
    foreach ($gpu in $gpus) {
        $friendlyName = $gpu.Name
        $pnp = $null
        $pciShort = ($gpu.Name -split "#")[1]
        if (-not [string]::IsNullOrWhiteSpace($pciShort) -and $displayDevices) {
            $normalizedPciShort = ($pciShort -replace '[^a-zA-Z0-9]', '').ToLowerInvariant()
            $pnp = $displayDevices | Where-Object {
                $instanceId = $_.InstanceId
                if ([string]::IsNullOrWhiteSpace($instanceId)) { return $false }
                $normalizedInstanceId = ($instanceId -replace '[^a-zA-Z0-9]', '').ToLowerInvariant()
                return $normalizedInstanceId -like "*$normalizedPciShort*"
            } | Select-Object -First 1

            if ($pnp -and -not [string]::IsNullOrWhiteSpace($pnp.FriendlyName)) {
                $friendlyName = $pnp.FriendlyName
            }
        }

        # Skip dedicated AI/NPU accelerators (but not GPUs with 'AI' in branding)
        # Also skip if PnP resolved to a non-Display class device
        if ($friendlyName -match '(?i)\b(NPU|Neural Processing|VPU|Coral|Myriad)\b') { continue }
        if ($pnp -and $pnp.Class -and $pnp.Class -ne 'Display') { continue }

        $list += [PSCustomObject]@{
            Friendly = $friendlyName
            Name     = $gpu.Name
            Provider = $gpu
        }
    }

    $removeOption = [PSCustomObject]@{
        Friendly = "NONE - Remove GPU Adapter"
        Name     = $null
        Provider = $null
    }
    return @(,$removeOption) + $list
}

function Get-GpuDriverStoreFolders {
    <#
    .SYNOPSIS
        Returns only GPU-relevant driver folders from the host DriverStore.
        This dramatically reduces copy size vs copying the entire DriverStore.
    #>
    param(
        [string]$DriverStore = "$env:SystemRoot\System32\DriverStore\FileRepository",
        [string]$GpuVendor = "Auto"
    )

    $folders = [System.Collections.Generic.List[System.IO.DirectoryInfo]]::new()

    # Method 1: Query actual display adapter PnP devices
    try {
        $displayDevices = Get-PnpDevice -Class Display -Status OK -ErrorAction Stop
        foreach ($dev in $displayDevices) {
            $infProp = Get-PnpDeviceProperty -InstanceId $dev.InstanceId `
                -KeyName 'DEVPKEY_Device_DriverInfPath' -ErrorAction SilentlyContinue
            if ($infProp -and $infProp.Data) {
                $infBase = [IO.Path]::GetFileNameWithoutExtension($infProp.Data)
                Get-ChildItem $DriverStore -Directory -Filter "$infBase*" -ErrorAction SilentlyContinue |
                    ForEach-Object { $folders.Add($_) }
            }
        }
        # GPU-related audio drivers (HDMI/DP audio) - exclude generic Realtek/Conexant HD Audio
        $audioDevs = Get-PnpDevice -Class MEDIA -Status OK -ErrorAction SilentlyContinue |
            Where-Object { $_.FriendlyName -match 'NVIDIA|AMD|Radeon|Intel.*(Display|Graphics)' }
        foreach ($dev in $audioDevs) {
            $infProp = Get-PnpDeviceProperty -InstanceId $dev.InstanceId `
                -KeyName 'DEVPKEY_Device_DriverInfPath' -ErrorAction SilentlyContinue
            if ($infProp -and $infProp.Data) {
                $infBase = [IO.Path]::GetFileNameWithoutExtension($infProp.Data)
                Get-ChildItem $DriverStore -Directory -Filter "$infBase*" -ErrorAction SilentlyContinue |
                    ForEach-Object { $folders.Add($_) }
            }
        }
    } catch {
        Write-Log "PnP device query failed, using pattern matching" "WARN"
    }

    # Method 2: Pattern-based matching (always runs as safety net)
    $patterns = Get-GpuDriverStoreFolderNamePatterns -GpuVendor $GpuVendor
    foreach ($p in $patterns) {
        Get-ChildItem $DriverStore -Directory -Filter $p -ErrorAction SilentlyContinue |
            ForEach-Object { $folders.Add($_) }
    }

    # Deduplicate by full path
    return $folders | Sort-Object FullName -Unique
}

function Copy-DriversToVhd {
    <#
    .SYNOPSIS
        Copies files from source to a destination inside the mounted VHD.
        Uses robocopy for reliability when available.
    #>
    param(
        [string]$VMName,
        [string]$MountLetter,
        [string]$Source,
        [string]$Destination,
        [string]$FileMask = "*",
        [switch]$ForceDelete
    )

    $copySucceeded = $true

    $target = Join-Path "$($MountLetter)\" $Destination

    # If target exists and ForceDelete, remove it first
    if ($ForceDelete -and (Test-Path $target)) {
        # Safety: verify target is under the mounted VHD drive
        $targetRoot = [System.IO.Path]::GetPathRoot($target)
        $expectedRoot = [System.IO.Path]::GetPathRoot("$($MountLetter)\")
        if ($targetRoot -ne $expectedRoot) {
            Write-Log "[$VMName] Refusing ForceDelete: target root '$targetRoot' does not match mount root '$expectedRoot'" "ERROR"
            return $false
        }
        Write-Log "[$VMName] Cleaning existing $target"
        try {
            $driveLetter = ($target -split ':')[0]
            $volume = Get-Volume -DriveLetter $driveLetter -ErrorAction SilentlyContinue
            if ($volume -and $volume.FileSystem -eq 'NTFS') {
                & takeown.exe /F "$target" /R /D Y 2>$null | Out-Null
                & icacls.exe "$target" /grant Administrators:F /T /C 2>$null | Out-Null
            }
            Remove-Item -Path $target -Recurse -Force -ErrorAction Stop
        } catch {
            Write-Log "[$VMName] Error cleaning directory: $($_.Exception.Message)" "WARN"
        }
    }

    if (-not (Test-Path $target)) {
        New-Item -Path $target -ItemType Directory -Force | Out-Null
    }

    Write-Log "[$VMName] Copying $Source -> $target ($FileMask)"
    try {
        if ($FileMask -eq "*" -and (Get-Command robocopy -ErrorAction SilentlyContinue)) {
            & robocopy "$Source" "$target" /E /MT:8 /NP /R:2 /W:1 /XJ /NFL /NDL /NJH /NJS 2>&1 | Out-Null
            if ($LASTEXITCODE -ge 8) {
                Write-Log "[$VMName] Robocopy returned exit code $LASTEXITCODE" "WARN"
                # Fallback to Copy-Item
                $srcPath = Join-Path $Source $FileMask
                if (Test-Path $srcPath) {
                    Copy-Item -Path $srcPath -Destination $target -Recurse -Force -ErrorAction Stop
                } else {
                    Write-Log "[$VMName] Source path not found for fallback copy: $srcPath" "WARN"
                    if ($FileMask -eq "*") { $copySucceeded = $false }
                }
            }
        } else {
            $srcPath = Join-Path $Source $FileMask
            if ($FileMask -ne "*" -and -not (Get-ChildItem -Path $srcPath -ErrorAction SilentlyContinue)) {
                Write-Log "[$VMName] No files matching '$FileMask' in $Source" "WARN"
            } elseif ($FileMask -eq "*" -and -not (Test-Path $srcPath)) {
                Write-Log "[$VMName] No files matching '$FileMask' in $Source" "WARN"
                $copySucceeded = $false
            } else {
                Copy-Item -Path $srcPath -Destination $target -Recurse -Force -ErrorAction Stop
            }
        }
    } catch {
        Write-Log "[$VMName] ERROR copying files: $($_.Exception.Message)" "ERROR"
        $copySucceeded = $false
    }

    return $copySucceeded
}

function Copy-GpuDriverFolders {
    <#
    .SYNOPSIS
        Smart GPU driver copy - only copies GPU-relevant DriverStore folders.
        Checks for sufficient disk space and auto-expands VHD if needed.
    #>
    [CmdletBinding()]
    param(
        [string]$VMName,
        [string]$MountLetter,
        [string]$VhdPath,
        [string]$GpuVendor = "Auto",
        [bool]$SmartCopy = $true,
        [bool]$AutoExpand = $true
    )

    if ([string]::IsNullOrWhiteSpace($MountLetter)) {
        Write-Log "[$VMName] Mount letter is null or empty. Cannot copy GPU drivers." "ERROR"
        return @{ Success = $false; MountLetter = $MountLetter }
    }

    $HostDriverStore = "$env:SystemRoot\System32\DriverStore\FileRepository"
    $VMDriverStore   = "Windows\System32\HostDriverStore\FileRepository"
    $targetBase      = Join-Path "$($MountLetter)\" $VMDriverStore

    if ($SmartCopy) {
        # ---- SMART COPY: Only GPU-relevant folders ----
        Write-Log "[$VMName] Using smart GPU driver copy (GPU folders only)"
        $gpuFolders = Get-GpuDriverStoreFolders -GpuVendor $GpuVendor

        if ($gpuFolders.Count -eq 0) {
            Write-Log "[$VMName] No GPU driver folders found!" "ERROR"
            return @{ Success = $false; MountLetter = $MountLetter }
        }

        # Calculate total size
        $totalSize = 0
        foreach ($f in $gpuFolders) {
            $totalSize += (Get-ChildItem $f.FullName -Recurse -ErrorAction SilentlyContinue |
                Measure-Object -Property Length -Sum).Sum
        }
        Write-Log "[$VMName] Found $($gpuFolders.Count) GPU driver folders ($([math]::Round($totalSize / 1GB, 2)) GB)"
    } else {
        # ---- FULL COPY: Entire DriverStore ----
        Write-Log "[$VMName] Using full DriverStore copy"
        $totalSize = (Get-ChildItem $HostDriverStore -Recurse -ErrorAction SilentlyContinue |
            Measure-Object -Property Length -Sum).Sum
        Write-Log "[$VMName] Full DriverStore size: $([math]::Round($totalSize / 1GB, 2)) GB"
    }

    # Check available space
    $freeSpace = (Get-Volume -DriveLetter $MountLetter[0] -ErrorAction SilentlyContinue).SizeRemaining
    if ($null -eq $freeSpace -or $freeSpace -le 0) {
        Write-Log "[$VMName] Could not determine free space for mounted VHD volume '$MountLetter'." "ERROR"
        return @{ Success = $false; MountLetter = $MountLetter }
    }
    Write-Log "[$VMName] VHD free space: $([math]::Round($freeSpace / 1GB, 2)) GB"

    if ($totalSize -gt ($freeSpace - 1GB)) {
        if (-not $AutoExpand) {
            Write-Log "[$VMName] Insufficient space and auto-expand is disabled." "ERROR"
            return @{ Success = $false; MountLetter = $MountLetter }
        }
        Write-Log "[$VMName] Insufficient space. Attempting VHD auto-expand..." "WARN"

        # Dismount VHD
        [void](Dismount-ImageRetry -ImagePath $VhdPath -MaxRetries 2)
        Start-Sleep -Seconds 2

        try {
            $currentMaxSize  = (Get-VHD -Path $VhdPath).Size
            $additionalNeeded = $totalSize - $freeSpace + 5GB
            $newSize = $currentMaxSize + $additionalNeeded

            Write-Log "[$VMName] Expanding VHD from $([math]::Round($currentMaxSize/1GB)) GB to $([math]::Round($newSize/1GB)) GB"
            Resize-VHD -Path $VhdPath -SizeBytes $newSize -ErrorAction Stop

            # Remount
            if (-not (Mount-VhdWithFallback -ImagePath $VhdPath)) {
                Write-Log "[$VMName] Could not remount VHD after expansion" "ERROR"
                return @{ Success = $false; MountLetter = $MountLetter }
            }
            Start-Sleep -Seconds 2

            $disk = Get-DiskImage -ImagePath $VhdPath | Get-Disk
            # Find the NTFS partition and extend it
            $partition = Get-DataPartitions -DiskNumber $disk.Number |
                Select-Object -Last 1

            if ($partition) {
                $maxPartSize = (Get-PartitionSupportedSize -DiskNumber $disk.Number -PartitionNumber $partition.PartitionNumber).SizeMax
                Resize-Partition -DiskNumber $disk.Number -PartitionNumber $partition.PartitionNumber -Size $maxPartSize
                Write-Log "[$VMName] VHD partition extended successfully" "OK"
            }

            # Re-check mounted Windows volume and free space
            Start-Sleep -Seconds 1
            $updatedDriveLetter = $null
            $partitions = Get-DataPartitions -DiskNumber $disk.Number
            foreach ($part in $partitions) {
                $vol = Get-Volume -Partition $part -ErrorAction SilentlyContinue
                if ($vol -and $vol.FileSystem -eq 'NTFS' -and $vol.DriveLetter) {
                    $updatedDriveLetter = [string]$vol.DriveLetter
                    break
                }
            }
            if ($updatedDriveLetter) {
                $MountLetter = "${updatedDriveLetter}:"
            }

            $freeSpace = (Get-Volume -DriveLetter $MountLetter[0] -ErrorAction SilentlyContinue).SizeRemaining
            if ($null -eq $freeSpace -or $freeSpace -le 0) {
                Write-Log "[$VMName] Could not determine free space after VHD expansion." "ERROR"
                return @{ Success = $false; MountLetter = $MountLetter }
            }
            Write-Log "[$VMName] New free space: $([math]::Round($freeSpace / 1GB, 2)) GB"

            if ($totalSize -gt ($freeSpace - 1GB)) {
                Write-Log "[$VMName] Still insufficient space after expansion!" "ERROR"
                return @{ Success = $false; MountLetter = $MountLetter }
            }
        } catch {
            Write-Log "[$VMName] VHD expansion failed: $($_.Exception.Message)" "ERROR"
            # Try to remount for cleanup
            try {
                [void](Mount-VhdWithFallback -ImagePath $VhdPath)
            } catch {
                Write-Log "[$VMName] VHD remount after expansion failure also failed: $($_.Exception.Message)" "WARN"
            }
            return @{ Success = $false; MountLetter = $MountLetter }
        }
    }

    # Ensure target directory exists
    if (-not (Test-Path $targetBase)) {
        New-Item -Path $targetBase -ItemType Directory -Force | Out-Null
    }

    # Perform the copy
    if ($SmartCopy) {
        $copied = 0
        foreach ($folder in $gpuFolders) {
            $dest = Join-Path $targetBase $folder.Name
            try {
                if (Get-Command robocopy -ErrorAction SilentlyContinue) {
                    & robocopy "$($folder.FullName)" "$dest" /E /MT:8 /NP /R:2 /W:1 /XJ /NFL /NDL /NJH /NJS 2>&1 | Out-Null
                    if ($LASTEXITCODE -lt 8) { $copied++ }
                    else {
                        Copy-Item -Path $folder.FullName -Destination $dest -Recurse -Force -ErrorAction Stop
                        $copied++
                    }
                } else {
                    Copy-Item -Path $folder.FullName -Destination $dest -Recurse -Force -ErrorAction Stop
                    $copied++
                }
            } catch {
                Write-Log "[$VMName] Error copying $($folder.Name): $($_.Exception.Message)" "WARN"
            }
        }
        Write-Log "[$VMName] Copied $copied of $($gpuFolders.Count) GPU driver folders" "OK"
        if ($copied -eq 0 -and $gpuFolders.Count -gt 0) {
            Write-Log "[$VMName] No GPU driver folders were copied successfully." "ERROR"
            return @{ Success = $false; MountLetter = $MountLetter }
        }
    } else {
        # Full copy using robocopy or Copy-Item
        Write-Log "[$VMName] Copying entire DriverStore (this may take a while)..."
        $fullCopyOk = Copy-DriversToVhd -VMName $VMName -MountLetter $MountLetter -Source $HostDriverStore -Destination $VMDriverStore -ForceDelete
        if (-not $fullCopyOk) {
            Write-Log "[$VMName] Full DriverStore copy did not complete successfully." "ERROR"
            return @{ Success = $false; MountLetter = $MountLetter }
        }
    }

    return @{ Success = $true; MountLetter = $MountLetter }
}

function Test-HostHasNvidiaGpu {
    $gpus = $script:VideoControllers
    $nvidiaGpus = $gpus | Where-Object { $_.Name -match "NVIDIA" }
    if ($nvidiaGpus) {
        Write-Log "Host NVIDIA GPU(s): $(($nvidiaGpus | ForEach-Object { $_.Name }) -join ', ')"
        return $true
    }
    return $false
}

function Set-GpuPartitionForVM {
    [CmdletBinding()]
    param(
        [string]$VMName,
        [ValidateRange(10,100)]
        [int]$AllocationPercent = 100,
        [bool]$ConservativeProfile = $true
    )

    try {
        $vmMemory = Get-VMMemory -VMName $VMName -ErrorAction Stop
        if ($vmMemory -and $vmMemory.DynamicMemoryEnabled) {
            $startupBytes = [UInt64]$vmMemory.Startup
            if ($startupBytes -le 0) {
                $startupBytes = [UInt64]$vmMemory.Minimum
            }
            if ($startupBytes -le 0) {
                $startupBytes = 4GB
            }
            Set-VMMemory -VMName $VMName -DynamicMemoryEnabled $false -StartupBytes $startupBytes -ErrorAction Stop
            Write-Log "[$VMName] Dynamic Memory was enabled; switched to static memory for GPU-P stability." "WARN"
        }
    } catch {
        Write-Log "[$VMName] Could not normalize VM memory mode before GPU-P apply: $($_.Exception.Message)" "WARN"
    }

    $partitionValues = Get-GpuPartitionValues -VMName $VMName -Percentage $AllocationPercent

    $setParams = @{
        VMName = $VMName
        ErrorAction = 'Stop'
    }

    if ($ConservativeProfile) {
        # Conservative profile: Min/Max/Optimal must all be set together.
        # If Max is omitted the hypervisor leaves it at the GPU's full hardware
        # maximum. The guest GPU driver reads Max during early-boot init and tries
        # to map that entire VRAM/compute aperture via SLAT. Because only Min and
        # Optimal are backed the mapping requests fault, the hypervisor spinlock
        # never releases, and the Windows boot animation freezes before login.
        # Use .Max (the slider-scaled value) for both MaxPartition* and OptimalPartition*.
        # .Optimal was previously the raw hardware cap — unaffected by the slider — so
        # the guest driver would map the full VRAM aperture and deadlock on boot.
        if ($partitionValues.VRAM.Supported) {
            $setParams['MinPartitionVRAM']     = $partitionValues.VRAM.Min
            $setParams['MaxPartitionVRAM']     = $partitionValues.VRAM.Max
            $setParams['OptimalPartitionVRAM'] = $partitionValues.VRAM.Max
        }
        if ($partitionValues.Encode.Supported) {
            $setParams['MinPartitionEncode']     = $partitionValues.Encode.Min
            $setParams['MaxPartitionEncode']     = $partitionValues.Encode.Max
            $setParams['OptimalPartitionEncode'] = $partitionValues.Encode.Max
        }
        if ($partitionValues.Decode.Supported) {
            $setParams['MinPartitionDecode']     = $partitionValues.Decode.Min
            $setParams['MaxPartitionDecode']     = $partitionValues.Decode.Max
            $setParams['OptimalPartitionDecode'] = $partitionValues.Decode.Max
        }
        if ($partitionValues.Compute.Supported) {
            $setParams['MinPartitionCompute']     = $partitionValues.Compute.Min
            $setParams['MaxPartitionCompute']     = $partitionValues.Compute.Max
            $setParams['OptimalPartitionCompute'] = $partitionValues.Compute.Max
        }
    } else {
        if ($partitionValues.VRAM.Supported) {
            $setParams['MinPartitionVRAM'] = $partitionValues.VRAM.Min
            $setParams['MaxPartitionVRAM'] = $partitionValues.VRAM.Max
            $setParams['OptimalPartitionVRAM'] = $partitionValues.VRAM.Optimal
        }
        if ($partitionValues.Encode.Supported) {
            $setParams['MinPartitionEncode'] = $partitionValues.Encode.Min
            $setParams['MaxPartitionEncode'] = $partitionValues.Encode.Max
            $setParams['OptimalPartitionEncode'] = $partitionValues.Encode.Optimal
        }
        if ($partitionValues.Decode.Supported) {
            $setParams['MinPartitionDecode'] = $partitionValues.Decode.Min
            $setParams['MaxPartitionDecode'] = $partitionValues.Decode.Max
            $setParams['OptimalPartitionDecode'] = $partitionValues.Decode.Optimal
        }
        if ($partitionValues.Compute.Supported) {
            $setParams['MinPartitionCompute'] = $partitionValues.Compute.Min
            $setParams['MaxPartitionCompute'] = $partitionValues.Compute.Max
            $setParams['OptimalPartitionCompute'] = $partitionValues.Compute.Optimal
        }
    }

    try {
        if ($setParams.Count -gt 2) {
            Set-VMGpuPartitionAdapter @setParams
        } else {
            Write-Log "[$VMName] GPU capability limits were not fully reported; applying default GPU partition settings." "WARN"
            Set-VMGpuPartitionAdapter -VMName $VMName -ErrorAction Stop
        }
    } catch {
        Write-Log "[$VMName] Explicit GPU partition sizing failed; retrying with host defaults. Error: $($_.Exception.Message)" "WARN"
        Set-VMGpuPartitionAdapter -VMName $VMName -ErrorAction Stop
    }

    $vmConfig = Get-VM -Name $VMName -ErrorAction SilentlyContinue
    # Reference: bryanem32/hyperv_vm_creator v29 uses LowMMIO=1GB, HighMMIO=32GB.
    # We keep HighMMIO at 128GB (better for large-VRAM cards) but align Low to 1GB.
    [UInt64]$targetLowMmio  = 1GB
    [UInt64]$targetHighMmio = 128GB
    if ($vmConfig) {
        if ([UInt64]$vmConfig.LowMemoryMappedIoSpace -gt $targetLowMmio) {
            $targetLowMmio = [UInt64]$vmConfig.LowMemoryMappedIoSpace
        }
        if ([UInt64]$vmConfig.HighMemoryMappedIoSpace -gt $targetHighMmio) {
            $targetHighMmio = [UInt64]$vmConfig.HighMemoryMappedIoSpace
        }
    }

    Set-VM -VMName $VMName -GuestControlledCacheTypes $true -ErrorAction Stop
    Set-VM -VMName $VMName -LowMemoryMappedIoSpace $targetLowMmio -ErrorAction Stop
    Set-VM -VMName $VMName -HighMemoryMappedIoSpace $targetHighMmio -ErrorAction Stop
}

#endregion

#region ==================== GUI CONSTRUCTION ====================

Write-Host "  [7/7] Building interface..." -ForegroundColor DarkGray

# ============================================================
#  FONT CACHE
# ============================================================
$script:FontMain          = New-Object System.Drawing.Font("Segoe UI", 9.5)
$script:FontTabHeader     = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$script:FontHeader        = New-Object System.Drawing.Font("Segoe UI", 8.75)
$script:FontBoldButton    = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$script:FontSmall         = New-Object System.Drawing.Font("Segoe UI", 8.25)
$script:FontBoldLabel     = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$script:FontConsolas      = New-Object System.Drawing.Font("Consolas", 9.5)
$script:FontSidebarNav    = New-Object System.Drawing.Font("Segoe UI", 10.5)
$script:FontAppTitle      = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)
$script:ThemeFontGroupBox = New-Object System.Drawing.Font("Segoe UI", 9.5, [System.Drawing.FontStyle]::Bold)

# ============================================================
#  COLOUR PALETTE  (VS Code-inspired charcoal dark theme)
# ============================================================
try {
    $theme = @{
        Bg            = [System.Drawing.Color]::FromArgb(30,  30,  30)
        Sidebar       = [System.Drawing.Color]::FromArgb(37,  37,  38)
        SidebarHov    = [System.Drawing.Color]::FromArgb(50,  50,  52)
        SidebarActive = [System.Drawing.Color]::FromArgb(0,   100, 177)
        SidebarAccent = [System.Drawing.Color]::FromArgb(0,   120, 215)
        Card          = [System.Drawing.Color]::FromArgb(43,  43,  45)
        Surface       = [System.Drawing.Color]::FromArgb(55,  55,  58)
        Input         = [System.Drawing.Color]::FromArgb(60,  60,  63)
        InputFocus    = [System.Drawing.Color]::FromArgb(74,  74,  78)
        Border        = [System.Drawing.Color]::FromArgb(68,  68,  72)
        BorderFocus   = [System.Drawing.Color]::FromArgb(0,   122, 204)
        HeaderBar     = [System.Drawing.Color]::FromArgb(0,   40,  80)
        Text          = [System.Drawing.Color]::FromArgb(204, 204, 204)
        TextHigh      = [System.Drawing.Color]::FromArgb(255, 255, 255)
        TextSecondary = [System.Drawing.Color]::FromArgb(180, 190, 210)
        TextMuted     = [System.Drawing.Color]::FromArgb(133, 133, 140)
        Muted         = [System.Drawing.Color]::FromArgb(133, 133, 140)
        Accent        = [System.Drawing.Color]::FromArgb(0,   120, 212)
        AccentHover   = [System.Drawing.Color]::FromArgb(28,  140, 228)
        AccentPressed = [System.Drawing.Color]::FromArgb(0,   100, 190)
        Success       = [System.Drawing.Color]::FromArgb(78,  201, 176)
        SuccessHover  = [System.Drawing.Color]::FromArgb(58,  175, 152)
        Warning       = [System.Drawing.Color]::FromArgb(220, 170, 20)
        WarningHover  = [System.Drawing.Color]::FromArgb(190, 148, 10)
        Danger        = [System.Drawing.Color]::FromArgb(241, 76,  76)
        DangerHover   = [System.Drawing.Color]::FromArgb(205, 49,  49)
        Info          = [System.Drawing.Color]::FromArgb(86,  156, 214)
        InfoHover     = [System.Drawing.Color]::FromArgb(60,  130, 190)
    }
    Write-StartupTrace -Message "Theme palette created successfully"
} catch {
    Write-StartupTrace -Message "Theme palette creation failed: $($_.Exception.Message)" -Level 'ERROR'
    $theme = @{
        Bg            = [System.Drawing.SystemColors]::ControlDarkDark
        Sidebar       = [System.Drawing.SystemColors]::ControlDark
        SidebarHov    = [System.Drawing.SystemColors]::Control
        SidebarActive = [System.Drawing.SystemColors]::Highlight
        SidebarAccent = [System.Drawing.SystemColors]::Highlight
        Card          = [System.Drawing.SystemColors]::Control
        Surface       = [System.Drawing.SystemColors]::Control
        Input         = [System.Drawing.SystemColors]::Window
        InputFocus    = [System.Drawing.SystemColors]::Window
        Border        = [System.Drawing.SystemColors]::ControlDark
        BorderFocus   = [System.Drawing.SystemColors]::Highlight
        HeaderBar     = [System.Drawing.SystemColors]::ControlDarkDark
        Text          = [System.Drawing.SystemColors]::ControlText
        TextHigh      = [System.Drawing.SystemColors]::ControlText
        TextSecondary = [System.Drawing.SystemColors]::ControlText
        TextMuted     = [System.Drawing.SystemColors]::GrayText
        Muted         = [System.Drawing.SystemColors]::GrayText
        Accent        = [System.Drawing.SystemColors]::Highlight
        AccentHover   = [System.Drawing.SystemColors]::Highlight
        AccentPressed = [System.Drawing.SystemColors]::Highlight
        Success       = [System.Drawing.Color]::Green
        SuccessHover  = [System.Drawing.Color]::DarkGreen
        Warning       = [System.Drawing.Color]::Orange
        WarningHover  = [System.Drawing.Color]::DarkOrange
        Danger        = [System.Drawing.Color]::Red
        DangerHover   = [System.Drawing.Color]::DarkRed
        Info          = [System.Drawing.Color]::CornflowerBlue
        InfoHover     = [System.Drawing.Color]::Blue
    }
}

# ============================================================
#  MAIN FORM
# ============================================================
$form = New-Object System.Windows.Forms.Form
$form.Text            = "Hyper-V Toolkit  —  Version 1  —  Diobyte"
$form.Size            = New-Object System.Drawing.Size(1280, 880)
$form.MinimumSize     = New-Object System.Drawing.Size(1100, 760)
$form.FormBorderStyle = 'Sizable'
$form.MaximizeBox     = $true
$form.StartPosition   = "CenterScreen"
$form.Font            = $script:FontMain
$form.BackColor       = $theme.Bg
$form.ForeColor       = $theme.Text
$form.AutoScaleMode   = [System.Windows.Forms.AutoScaleMode]::Font
$form.KeyPreview      = $true
$form.Padding         = New-Object System.Windows.Forms.Padding(0)

# ============================================================
#  MENU STRIP (minimal)
# ============================================================
$menuStrip = New-Object System.Windows.Forms.MenuStrip
$menuStrip.BackColor = $theme.Sidebar
$menuStrip.ForeColor = $theme.Text
$menuStrip.RenderMode = [System.Windows.Forms.ToolStripRenderMode]::Professional

$fileMenu = New-Object System.Windows.Forms.ToolStripMenuItem
$fileMenu.Text = "&File"
$exitItem = New-Object System.Windows.Forms.ToolStripMenuItem
$exitItem.Text = "E&xit"
$exitItem.ShortcutKeys = [System.Windows.Forms.Keys]::Alt -bor [System.Windows.Forms.Keys]::F4
$exitItem.Add_Click({ $form.Close() })
$fileMenu.DropDownItems.Add($exitItem)
$menuStrip.Items.Add($fileMenu)

$viewMenu = New-Object System.Windows.Forms.ToolStripMenuItem
$viewMenu.Text = "&View"

$clearLogItem = New-Object System.Windows.Forms.ToolStripMenuItem
$clearLogItem.Text = "&Clear Log"
$clearLogItem.Add_Click({ $script:LogBox.Clear() })
$viewMenu.DropDownItems.Add($clearLogItem)
$menuStrip.Items.Add($viewMenu)

$helpMenu = New-Object System.Windows.Forms.ToolStripMenuItem
$helpMenu.Text = "&Help"
$aboutItem = New-Object System.Windows.Forms.ToolStripMenuItem
$aboutItem.Text = "&About"
$aboutItem.Add_Click({
    [System.Windows.Forms.MessageBox]::Show(
        "Hyper-V Toolkit v1`nCreated by Diobyte`nMade with love",
        "About", "OK", "Information"
    )
})
$helpMenu.DropDownItems.Add($aboutItem)
$menuStrip.Items.Add($helpMenu)

$form.MainMenuStrip = $menuStrip
$form.Controls.Add($menuStrip)

# ============================================================
#  STATUS STRIP
# ============================================================
$statusStrip = New-Object System.Windows.Forms.StatusStrip
$statusStrip.BackColor = $theme.Card
$statusStrip.ForeColor = $theme.TextMuted

$statusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
$statusLabel.Text    = "Ready"
$statusLabel.AutoSize = $true
$statusStrip.Items.Add($statusLabel)

$progressBar = New-Object System.Windows.Forms.ToolStripProgressBar
$progressBar.Visible  = $false
$progressBar.Minimum  = 0
$progressBar.Maximum  = 100
$progressBar.Width    = 200
$statusStrip.Items.Add($progressBar)

$script:StatusLabel       = $statusLabel
$script:StatusProgressBar = $progressBar

function Update-StatusBar {
    param([string]$Message = "Ready", [int]$Progress = -1)
    if ($script:StatusLabel)       { $script:StatusLabel.Text = $Message }
    if ($script:StatusProgressBar) {
        if ($Progress -ge 0) {
            $script:StatusProgressBar.Visible = $true
            $script:StatusProgressBar.Value   = [Math]::Min(100, [Math]::Max(0, $Progress))
        } else {
            $script:StatusProgressBar.Visible = $false
        }
    }
    if ($script:HeaderStatusLabel) {
        $script:HeaderStatusLabel.Text = "  $Message"
        $script:HeaderStatusLabel.ForeColor = `
            $(if ($Message -match 'error|fail')            { $theme.Danger }
              elseif ($Message -match 'warn')              { $theme.Warning }
              elseif ($Message -match 'ok|ready|complete|success|done') { $theme.Success }
              else                                    { $theme.Info })
    }
}

$form.Controls.Add($statusStrip)

# ============================================================
#  LAYOUT – Docked shell panels
#
#   [MenuStrip - auto at top by WinForms]
#   [HeaderBar  48 px  Dock=Top     ]
#   [ContentRow          Dock=Fill  ]
#     [Sidebar 170 px  Dock=Left    ]
#     [SidebarDivider 1px Dock=Left ]
#     [ContentArea     Dock=Fill    ]
#   [LogFooter 155 px  Dock=Bottom  ]
#   [StatusStrip - auto at bottom   ]
# ============================================================

# ---- HEADER BAR ----
$pnlHeader = New-Object System.Windows.Forms.Panel
$pnlHeader.Dock      = [System.Windows.Forms.DockStyle]::Top
$pnlHeader.Height    = 52
$pnlHeader.BackColor = $theme.HeaderBar
$pnlHeader.Padding   = New-Object System.Windows.Forms.Padding(0)
$form.Controls.Add($pnlHeader)

# Accent left-edge stripe on header
$pnlHeaderAccent = New-Object System.Windows.Forms.Panel
$pnlHeaderAccent.Dock      = [System.Windows.Forms.DockStyle]::Left
$pnlHeaderAccent.Width     = 4
$pnlHeaderAccent.BackColor = $theme.SidebarAccent
$pnlHeader.Controls.Add($pnlHeaderAccent)

$lblAppTitle = New-Object System.Windows.Forms.Label
$lblAppTitle.Text      = "HYPER-V TOOLKIT"
$lblAppTitle.Font      = $script:FontAppTitle
$lblAppTitle.ForeColor = [System.Drawing.Color]::White
$lblAppTitle.AutoSize  = $true
$lblAppTitle.Location  = New-Object System.Drawing.Point(18, 9)
$pnlHeader.Controls.Add($lblAppTitle)

$lblAppMeta = New-Object System.Windows.Forms.Label
$lblAppMeta.Text      = "Version 1  |  Diobyte  |  Made with love"
$lblAppMeta.Font      = $script:FontSmall
$lblAppMeta.ForeColor = [System.Drawing.Color]::FromArgb(160, 185, 215)
$lblAppMeta.AutoSize  = $true
$lblAppMeta.Location  = New-Object System.Drawing.Point(22, 35)
$pnlHeader.Controls.Add($lblAppMeta)

$lblHeaderStatus = New-Object System.Windows.Forms.Label
$lblHeaderStatus.Text      = "  Ready"
$lblHeaderStatus.Font      = $script:FontBoldLabel
$lblHeaderStatus.ForeColor = $theme.Success
$lblHeaderStatus.AutoSize  = $true
$lblHeaderStatus.Location  = New-Object System.Drawing.Point(420, 19)
$pnlHeader.Controls.Add($lblHeaderStatus)
$script:HeaderStatusLabel = $lblHeaderStatus

# ---- LOG FOOTER (before content row so Dock=Bottom takes effect first) ----
$pnlLogFooter = New-Object System.Windows.Forms.Panel
$pnlLogFooter.Dock      = [System.Windows.Forms.DockStyle]::Bottom
$pnlLogFooter.Height    = 155
$pnlLogFooter.BackColor = $theme.Card
$pnlLogFooter.Padding   = New-Object System.Windows.Forms.Padding(0)
$form.Controls.Add($pnlLogFooter)

# Top border on log footer
$pnlLogTopBorder = New-Object System.Windows.Forms.Panel
$pnlLogTopBorder.Dock      = [System.Windows.Forms.DockStyle]::Top
$pnlLogTopBorder.Height    = 1
$pnlLogTopBorder.BackColor = $theme.Border
$pnlLogFooter.Controls.Add($pnlLogTopBorder)

# Button column on the right of the log footer
$pnlLogButtons = New-Object System.Windows.Forms.Panel
$pnlLogButtons.Dock      = [System.Windows.Forms.DockStyle]::Right
$pnlLogButtons.Width     = 100
$pnlLogButtons.BackColor = $theme.Card
$pnlLogFooter.Controls.Add($pnlLogButtons)

$script:LogBox           = New-Object System.Windows.Forms.RichTextBox
$script:LogBox.Dock      = [System.Windows.Forms.DockStyle]::Fill
$script:LogBox.ReadOnly  = $true
$script:LogBox.BackColor = $theme.Input
$script:LogBox.ForeColor = $theme.Text
$script:LogBox.Font      = $script:FontConsolas
$script:LogBox.WordWrap  = $false
$script:LogBox.ScrollBars = 'Both'
$script:LogBox.BorderStyle = 'None'
$pnlLogFooter.Controls.Add($script:LogBox)
$pnlLogFooter.Controls.SetChildIndex($script:LogBox, 0)

function New-LogButton {
    param([string]$Text, [int]$Top, [System.Drawing.Color]$Bg, [System.Drawing.Color]$Hover)
    $btn = New-Object System.Windows.Forms.Button
    $btn.Text      = $Text
    $btn.Size      = New-Object System.Drawing.Size(84, 30)
    $btn.Location  = New-Object System.Drawing.Point(8, $Top)
    $btn.FlatStyle = 'Flat'
    $btn.BackColor = $Bg
    $btn.ForeColor = [System.Drawing.Color]::White
    $btn.Font      = $script:FontSmall
    $btn.Cursor    = [System.Windows.Forms.Cursors]::Hand
    $btn.FlatAppearance.BorderColor       = $theme.Border
    $btn.FlatAppearance.MouseOverBackColor = $Hover
    $btn.FlatAppearance.MouseDownBackColor = $Hover
    $pnlLogButtons.Controls.Add($btn)
    return $btn
}

$btnClearLog = New-LogButton "Clear Log" 10  $theme.Surface  $theme.SidebarHov
$btnSaveLog  = New-LogButton "Save Log"  46  $theme.Surface  $theme.SidebarHov
$btnExit     = New-LogButton "Exit"      82  $theme.Danger   $theme.DangerHover

$btnSaveLog.Add_Click({
    $saveDlg = New-Object System.Windows.Forms.SaveFileDialog
    try {
        $saveDlg.Filter   = "Log files (*.log)|*.log|Text files (*.txt)|*.txt|All files (*.*)|*.*"
        $saveDlg.FileName = "HyperV-Toolkit_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
        if ($saveDlg.ShowDialog() -eq 'OK') {
            try {
                $script:LogBox.Text | Out-File -FilePath $saveDlg.FileName -Encoding UTF8 -Force
                Write-Log "Log saved to: $($saveDlg.FileName)" "OK"
            } catch {
                Write-Log "Failed to save log: $($_.Exception.Message)" "ERROR"
            }
        }
    } finally { $saveDlg.Dispose() }
})
$btnClearLog.Add_Click({ $script:LogBox.Clear() })
$btnExit.Add_Click({ $form.Close() })

# ---- CONTENT ROW (fills between header and footer) ----
$pnlContentRow = New-Object System.Windows.Forms.Panel
$pnlContentRow.Dock      = [System.Windows.Forms.DockStyle]::Fill
$pnlContentRow.BackColor = $theme.Bg
$pnlContentRow.Padding   = New-Object System.Windows.Forms.Padding(0)
$form.Controls.Add($pnlContentRow)
# Move Fill control to front of z-order so the dock layout engine processes it
# LAST (after Top/Bottom edge controls claim their space)
$form.Controls.SetChildIndex($pnlContentRow, 0)

# ---- CONTENT AREA ----
$pnlContent = New-Object System.Windows.Forms.Panel
$pnlContent.Dock      = [System.Windows.Forms.DockStyle]::Fill
$pnlContent.BackColor = $theme.Bg
$pnlContent.Padding   = New-Object System.Windows.Forms.Padding(0)

# ---- SIDEBAR ----
$pnlSidebar = New-Object System.Windows.Forms.Panel
$pnlSidebar.Dock      = [System.Windows.Forms.DockStyle]::Left
$pnlSidebar.Width     = 172
$pnlSidebar.BackColor = $theme.Sidebar

# Sidebar right-edge divider
$pnlSidebarDiv = New-Object System.Windows.Forms.Panel
$pnlSidebarDiv.Dock      = [System.Windows.Forms.DockStyle]::Left
$pnlSidebarDiv.Width     = 1
$pnlSidebarDiv.BackColor = $theme.Border

# Add in correct z-order for docking: Fill first (index 0, processed last),
# then edge controls (higher indices, processed first by layout engine)
$pnlContentRow.Controls.Add($pnlContent)
$pnlContentRow.Controls.Add($pnlSidebarDiv)
$pnlContentRow.Controls.Add($pnlSidebar)

# ---- SIDEBAR: Branding block ----
$pnlBrand = New-Object System.Windows.Forms.Panel
$pnlBrand.Dock      = [System.Windows.Forms.DockStyle]::Top
$pnlBrand.Height    = 64
$pnlBrand.BackColor = [System.Drawing.Color]::FromArgb(25, 25, 26)
$pnlSidebar.Controls.Add($pnlBrand)

$lblSidebarHV = New-Object System.Windows.Forms.Label
$lblSidebarHV.Text      = "HV"
$script:FontSidebarBrand = New-Object System.Drawing.Font("Segoe UI", 22, [System.Drawing.FontStyle]::Bold)
$lblSidebarHV.Font      = $script:FontSidebarBrand
$lblSidebarHV.ForeColor = $theme.SidebarAccent
$lblSidebarHV.AutoSize  = $false
$lblSidebarHV.Size      = New-Object System.Drawing.Size(172, 64)
$lblSidebarHV.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
$pnlBrand.Controls.Add($lblSidebarHV)

# ---- SIDEBAR: Nav helper ----
$script:NavButtons  = @()
$script:NavPanelIdx = 0

function New-NavButton {
    param([string]$Text, [int]$TopOffset)
    $btn = New-Object System.Windows.Forms.Button
    $btn.Text      = $Text
    $btn.Size      = New-Object System.Drawing.Size(172, 52)
    $btn.Location  = New-Object System.Drawing.Point(0, $TopOffset)
    $btn.FlatStyle = 'Flat'
    $btn.BackColor = $theme.Sidebar
    $btn.ForeColor = $theme.TextMuted
    $btn.Font      = $script:FontSidebarNav
    $btn.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $btn.Padding   = New-Object System.Windows.Forms.Padding(18, 0, 0, 0)
    $btn.Cursor    = [System.Windows.Forms.Cursors]::Hand
    $btn.FlatAppearance.BorderSize             = 0
    $btn.FlatAppearance.MouseOverBackColor     = $theme.SidebarHov
    $btn.FlatAppearance.MouseDownBackColor     = $theme.SidebarActive
    $pnlSidebar.Controls.Add($btn)
    return $btn
}

$btnNavCreate = New-NavButton "+    Create VM"    64
$btnNavGPU    = New-NavButton "#    GPU Setup"    120

$script:NavButtons = @($btnNavCreate, $btnNavGPU)

# Sidebar separator and host status indicator
$pnlSidebarSepLine = New-Object System.Windows.Forms.Label
$pnlSidebarSepLine.BorderStyle = 'Fixed3D'
$pnlSidebarSepLine.Size        = New-Object System.Drawing.Size(140, 2)
$pnlSidebarSepLine.Location    = New-Object System.Drawing.Point(16, 184)
$pnlSidebar.Controls.Add($pnlSidebarSepLine)

$lblSidebarHVState = New-Object System.Windows.Forms.Label
$lblSidebarHVState.Text      = "Hyper-V  Active"
$lblSidebarHVState.Font      = $script:FontSmall
$lblSidebarHVState.ForeColor = $theme.Success
$lblSidebarHVState.AutoSize  = $false
$lblSidebarHVState.Size      = New-Object System.Drawing.Size(172, 20)
$lblSidebarHVState.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
$lblSidebarHVState.Location  = New-Object System.Drawing.Point(0, 194)
$pnlSidebar.Controls.Add($lblSidebarHVState)

$lblSidebarOs = New-Object System.Windows.Forms.Label
$lblSidebarOs.Text      = if ($script:HostIsWin11) { "Windows 11 Host" } else { "Windows 10 Host" }
$lblSidebarOs.Font      = $script:FontSmall
$lblSidebarOs.ForeColor = $theme.TextMuted
$lblSidebarOs.AutoSize  = $false
$lblSidebarOs.Size      = New-Object System.Drawing.Size(172, 18)
$lblSidebarOs.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
$lblSidebarOs.Location  = New-Object System.Drawing.Point(0, 272)
$pnlSidebar.Controls.Add($lblSidebarOs)

# ============================================================
#  CONTENT PANELS  (one per section, only one visible at a time)
# ============================================================

# ── Helper used across all panels ──────────────────────────
function New-LabeledControl {
    param(
        [System.Windows.Forms.Control]$Parent,
        [int]$X, [int]$Y,
        [string]$LabelText,
        [int]$LabelWidth = 120,
        [string]$ControlType = "TextBox",
        [int]$ControlWidth = 260,
        [hashtable]$ControlProps = @{}
    )
    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text      = $LabelText
    $lbl.ForeColor = $theme.Text
    $lbl.Location  = New-Object System.Drawing.Point($X, ($Y + 3))
    $lbl.AutoSize  = $false
    $lbl.Size      = New-Object System.Drawing.Size(($LabelWidth - 4), 22)
    $lbl.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $Parent.Controls.Add($lbl)

    $ctrl = switch ($ControlType) {
        "TextBox"       { New-Object System.Windows.Forms.TextBox }
        "ComboBox"      { $c = New-Object System.Windows.Forms.ComboBox; $c.DropDownStyle = 'DropDownList'; $c }
        "NumericUpDown" { New-Object System.Windows.Forms.NumericUpDown }
        "CheckBox"      { New-Object System.Windows.Forms.CheckBox }
        "Label"         { $l = New-Object System.Windows.Forms.Label; $l.ForeColor = $theme.Info; $l }
        default         { New-Object System.Windows.Forms.TextBox }
    }
    $ctrl.Location = New-Object System.Drawing.Point(($X + $LabelWidth), $Y)
    if ($ControlType -ne "CheckBox" -and $ControlType -ne "Label") {
        $ctrl.Width  = $ControlWidth
        $ctrl.BackColor = $theme.Input
        $ctrl.ForeColor = $theme.Text
    }
    if ($ControlType -eq "TextBox")  { $ctrl.BorderStyle = 'FixedSingle' }
    if ($ControlType -eq "ComboBox") { $ctrl.FlatStyle   = 'Flat' }
    foreach ($k in $ControlProps.Keys) { $ctrl.$k = $ControlProps[$k] }
    $Parent.Controls.Add($ctrl)
    return $ctrl
}

# ── GroupBox factory ────────────────────────────────────────
function New-ThemedGroupBox {
    param([string]$Title, [System.Windows.Forms.Control]$Parent)
    $gb = New-Object System.Windows.Forms.GroupBox
    $gb.Text      = $Title
    $gb.ForeColor = $theme.Text
    $gb.BackColor = $theme.Surface
    $gb.Font      = $script:ThemeFontGroupBox
    $Parent.Controls.Add($gb)
    return $gb
}

# ── Horizontal divider ──────────────────────────────────────
function New-Divider {
    param([System.Windows.Forms.Control]$Parent, [int]$X, [int]$Y, [int]$Width)
    $div = New-Object System.Windows.Forms.Label
    $div.Text        = ""
    $div.BorderStyle = 'Fixed3D'
    $div.Size        = New-Object System.Drawing.Size($Width, 2)
    $div.Location    = New-Object System.Drawing.Point($X, $Y)
    $Parent.Controls.Add($div)
}

# ── Themed section-title label ──────────────────────────────
function New-SectionLabel {
    param([string]$Text, [System.Windows.Forms.Control]$Parent, [int]$X, [int]$Y)
    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text      = $Text
    $lbl.Font      = $script:FontBoldLabel
    $lbl.ForeColor = $theme.Muted
    $lbl.AutoSize  = $true
    $lbl.Location  = New-Object System.Drawing.Point($X, $Y)
    $Parent.Controls.Add($lbl)
    return $lbl
}

# ============================================================
#  PANEL 1:  CREATE VM
# ============================================================
$tabCreate = New-Object System.Windows.Forms.Panel
$tabCreate.Dock       = [System.Windows.Forms.DockStyle]::Fill
$tabCreate.BackColor  = $theme.Bg
$tabCreate.AutoScroll = $true
$tabCreate.Visible    = $true       # default view
$pnlContent.Controls.Add($tabCreate)

$ctrlCreate = @{}

# ── Panel header ────────────────────────────────────────────
$pnlCreateHeader = New-Object System.Windows.Forms.Panel
$pnlCreateHeader.Dock      = [System.Windows.Forms.DockStyle]::Top
$pnlCreateHeader.Height    = 44
$pnlCreateHeader.BackColor = $theme.Card

$lblCreateTitle = New-Object System.Windows.Forms.Label
$lblCreateTitle.Text      = "Create Virtual Machine"
$lblCreateTitle.Font      = $script:FontTabHeader
$lblCreateTitle.ForeColor = $theme.TextHigh
$lblCreateTitle.AutoSize  = $true
$lblCreateTitle.Location  = New-Object System.Drawing.Point(14, 8)
$pnlCreateHeader.Controls.Add($lblCreateTitle)

$lblCreateSub = New-Object System.Windows.Forms.Label
$lblCreateSub.Text      = "Configure and provision a new Hyper-V virtual machine with automated unattended setup"
$lblCreateSub.Font      = $script:FontSmall
$lblCreateSub.ForeColor = $theme.TextMuted
$lblCreateSub.AutoSize  = $true
$lblCreateSub.Location  = New-Object System.Drawing.Point(14, 26)
$pnlCreateHeader.Controls.Add($lblCreateSub)

$tabCreate.Controls.Add($pnlCreateHeader)

# ── Scrollable body ─────────────────────────────────────────
$pnlCreateBody = New-Object System.Windows.Forms.Panel
$pnlCreateBody.AutoScroll  = $true
$pnlCreateBody.BackColor   = $theme.Bg
$pnlCreateBody.Location    = New-Object System.Drawing.Point(0, 44)
$tabCreate.Controls.Add($pnlCreateBody)

# ── LEFT COLUMN ─────────────────────────────────────────────
$grpConfig           = New-ThemedGroupBox "VM Configuration" $pnlCreateBody
$grpConfig.Location  = New-Object System.Drawing.Point(12, 12)
$grpConfig.Size      = New-Object System.Drawing.Size(470, 570)
$grpConfig.Anchor    = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left

$rowY = 26

$ctrlCreate["VMName"] = New-LabeledControl $grpConfig 12 $rowY "VM Name:"            -LabelWidth 130 -ControlWidth 310
$rowY += 36

$ctrlCreate["VMLocation"] = New-LabeledControl $grpConfig 12 $rowY "VM Location:" -LabelWidth 130 -ControlWidth 240
$btnBrowseVM = New-Object System.Windows.Forms.Button
$btnBrowseVM.Text      = "Browse"
$btnBrowseVM.Size      = New-Object System.Drawing.Size(60, 24)
$btnBrowseVM.Location  = New-Object System.Drawing.Point(390, $rowY)
$btnBrowseVM.FlatStyle = 'Flat'
$btnBrowseVM.BackColor = $theme.Surface
$btnBrowseVM.ForeColor = $theme.Text
$btnBrowseVM.Cursor    = [System.Windows.Forms.Cursors]::Hand
$grpConfig.Controls.Add($btnBrowseVM)
try {
    $ctrlCreate["VMLocation"].Text = (Get-VMHost).VirtualMachinePath
} catch {
    $ctrlCreate["VMLocation"].Text = Join-Path $env:PUBLIC "Hyper-V"
    Write-Log "Using fallback VM location." "WARN"
}
$rowY += 36

$ctrlCreate["ISOPath"] = New-LabeledControl $grpConfig 12 $rowY "ISO File:" -LabelWidth 130 -ControlWidth 240
$btnBrowseISO = New-Object System.Windows.Forms.Button
$btnBrowseISO.Text      = "Browse"
$btnBrowseISO.Size      = New-Object System.Drawing.Size(60, 24)
$btnBrowseISO.Location  = New-Object System.Drawing.Point(390, $rowY)
$btnBrowseISO.FlatStyle = 'Flat'
$btnBrowseISO.BackColor = $theme.Surface
$btnBrowseISO.ForeColor = $theme.Text
$btnBrowseISO.Cursor    = [System.Windows.Forms.Cursors]::Hand
$grpConfig.Controls.Add($btnBrowseISO)
$rowY += 36

$ctrlCreate["Edition"] = New-LabeledControl $grpConfig 12 $rowY "Edition:" -LabelWidth 130 -ControlType ComboBox -ControlWidth 310
$rowY += 34

$ctrlCreate["OSInfo"] = New-LabeledControl $grpConfig 12 $rowY "Detected OS:" -LabelWidth 130 -ControlType Label -ControlWidth 300 `
    -ControlProps @{ AutoSize = $true; Text = "(select an ISO to detect)" }
$rowY += 30

New-Divider $grpConfig 12 $rowY 440
$rowY += 14

$ctrlCreate["Username"] = New-LabeledControl $grpConfig 12 $rowY "Local Username:" -LabelWidth 130 -ControlWidth 200
$ctrlCreate["Username"].Text = "User"
$rowY += 36

$ctrlCreate["Password"] = New-LabeledControl $grpConfig 12 $rowY "Password:" -LabelWidth 130 -ControlWidth 200
$ctrlCreate["Password"].UseSystemPasswordChar = $true
$rowY += 36

New-Divider $grpConfig 12 $rowY 440
$rowY += 14

$ctrlCreate["vCPU"] = New-LabeledControl $grpConfig 12 $rowY "vCPUs:" -LabelWidth 130 -ControlType NumericUpDown -ControlWidth 80 `
    -ControlProps @{ Minimum = 1; Maximum = [Environment]::ProcessorCount; Value = [Math]::Min(4, [Environment]::ProcessorCount); DecimalPlaces = 0 }
$rowY += 34

$totalRamGB = [math]::Round($script:HostComputerSystem.TotalPhysicalMemory / 1GB)
$ctrlCreate["Memory"] = New-LabeledControl $grpConfig 12 $rowY "Memory (GB):" -LabelWidth 130 -ControlType NumericUpDown -ControlWidth 80 `
    -ControlProps @{ Minimum = 1; Maximum = $totalRamGB; Value = [Math]::Min(8, $totalRamGB); DecimalPlaces = 0 }
$rowY += 34

$ctrlCreate["DiskSize"] = New-LabeledControl $grpConfig 12 $rowY "Disk (GB):" -LabelWidth 130 -ControlType NumericUpDown -ControlWidth 80 `
    -ControlProps @{ Minimum = 20; Maximum = 2048; Value = 80; DecimalPlaces = 0 }
$rowY += 34

$ctrlCreate["Switch"] = New-LabeledControl $grpConfig 12 $rowY "Virtual Switch:" -LabelWidth 130 -ControlType ComboBox -ControlWidth 200
try {
    Get-VMSwitch | Select-Object -ExpandProperty Name | ForEach-Object { [void]$ctrlCreate["Switch"].Items.Add($_) }
    if ($ctrlCreate["Switch"].Items.Count -gt 0) { $ctrlCreate["Switch"].SelectedIndex = 0 }
} catch {
    Write-Log "Could not enumerate Hyper-V switches: $($_.Exception.Message)" "WARN"
}
$rowY += 34

$ctrlCreate["Resolution"] = New-LabeledControl $grpConfig 12 $rowY "Resolution:" -LabelWidth 130 -ControlType ComboBox -ControlWidth 150
@("800x600","1024x768","1280x720","1280x800","1280x1024","1366x768","1440x900","1600x900","1680x1050","1920x1080") |
    ForEach-Object { [void]$ctrlCreate["Resolution"].Items.Add($_) }
$ctrlCreate["Resolution"].SelectedItem = "1920x1080"
$rowY += 34

$ctrlCreate["CheckpointMode"] = New-LabeledControl $grpConfig 12 $rowY "Checkpoints:" -LabelWidth 130 -ControlType ComboBox -ControlWidth 150
@("Disabled","Production","ProductionOnly","Standard") | ForEach-Object { [void]$ctrlCreate["CheckpointMode"].Items.Add($_) }
$ctrlCreate["CheckpointMode"].SelectedItem = "Disabled"
$rowY += 34

$ctrlCreate["DynamicMemMin"] = New-LabeledControl $grpConfig 12 $rowY "Dyn. Min (GB):"  -LabelWidth 130 -ControlType NumericUpDown -ControlWidth 80 `
    -ControlProps @{ Minimum = 1; Maximum = $totalRamGB; Value = 1; DecimalPlaces = 0 }
$rowY += 34

$ctrlCreate["DynamicMemMax"] = New-LabeledControl $grpConfig 12 $rowY "Dyn. Max (GB):"  -LabelWidth 130 -ControlType NumericUpDown -ControlWidth 80 `
    -ControlProps @{ Minimum = 1; Maximum = $totalRamGB; Value = [Math]::Min(16, $totalRamGB); DecimalPlaces = 0 }

# Resize grpConfig to fit content
$configBottom = ($grpConfig.Controls | ForEach-Object { $_.Bottom } | Measure-Object -Maximum).Maximum
$grpConfig.Height = [Math]::Max(560, $configBottom + 14)

# ── RIGHT COLUMN ────────────────────────────────────────────
$grpBoot = New-ThemedGroupBox "Boot && Hardware" $pnlCreateBody
$grpBoot.Location = New-Object System.Drawing.Point(496, 12)
$grpBoot.Size     = New-Object System.Drawing.Size(480, 120)
$grpBoot.Anchor   = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left

$chkBootY = 26
foreach ($chkDef in @(
    @{ Key = "SecureBoot"; Text = "Secure Boot  (auto: ON for Win11 and modern Win10)"; Default = $true  },
    @{ Key = "TPM";        Text = "Virtual TPM  (required for Windows 11)";            Default = $true  },
    @{ Key = "VHDType";    Text = "Fixed-size VHD  (default: dynamic / expanding)";   Default = $false }
)) {
    $cb = New-Object System.Windows.Forms.CheckBox
    $cb.Text      = $chkDef.Text
    $cb.AutoSize  = $true
    $cb.Checked   = $chkDef.Default
    $cb.Location  = New-Object System.Drawing.Point(14, $chkBootY)
    $cb.ForeColor = $theme.Text
    $cb.FlatStyle = 'Flat'
    $grpBoot.Controls.Add($cb)
    $ctrlCreate[$chkDef.Key] = $cb
    $chkBootY += 30
}

$grpOpts = New-ThemedGroupBox "VM Options" $pnlCreateBody
$grpOpts.Location = New-Object System.Drawing.Point(496, 144)
$grpOpts.Size     = New-Object System.Drawing.Size(480, 200)
$grpOpts.Anchor   = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left

$optsLeft  = @(
    @{ Key = "DynamicMem";      Text = "Dynamic Memory";             X = 14;  Y = 26 },
    @{ Key = "EnhancedSession"; Text = "Enhanced Session Mode";      X = 14;  Y = 56 },
    @{ Key = "StartVM";         Text = "Start VM after creation";    X = 14;  Y = 86 },
    @{ Key = "StrictLegacyMode";Text = "Strict Legacy Mode (Win10)"; X = 14;  Y = 116}
)
$optsRight = @(
    @{ Key = "AutoCreateSwitch";Text = "Auto-create NAT switch";     X = 248; Y = 26 },
    @{ Key = "EnableMetering";  Text = "Enable Resource Metering";   X = 248; Y = 56 },
    @{ Key = "EnableAutoLogon"; Text = "Enable Auto Logon";          X = 248; Y = 86 }
)
foreach ($chkDef in ($optsLeft + $optsRight)) {
    $cb = New-Object System.Windows.Forms.CheckBox
    $cb.Text      = $chkDef.Text
    $cb.AutoSize  = $true
    $cb.Checked   = ($chkDef.Key -in @("StartVM","AutoCreateSwitch","EnableMetering","EnableAutoLogon"))
    $cb.Location  = New-Object System.Drawing.Point($chkDef.X, $chkDef.Y)
    $cb.ForeColor = $theme.Text
    $cb.FlatStyle = 'Flat'
    $grpOpts.Controls.Add($cb)
    $ctrlCreate[$chkDef.Key] = $cb
}

$ctrlCreate["RoutingHint"] = New-Object System.Windows.Forms.Label
$ctrlCreate["RoutingHint"].Text      = "Network: Auto-create NAT provides fallback networking when no switch is selected."
$ctrlCreate["RoutingHint"].Size      = New-Object System.Drawing.Size(450, 40)
$ctrlCreate["RoutingHint"].Location  = New-Object System.Drawing.Point(14, 148)
$ctrlCreate["RoutingHint"].ForeColor = $theme.Muted
$ctrlCreate["RoutingHint"].AutoEllipsis = $true
$grpOpts.Controls.Add($ctrlCreate["RoutingHint"])

# ── Create VM action strip ───────────────────────────────────
$pnlCreateAction = New-Object System.Windows.Forms.Panel
$pnlCreateAction.BackColor = $theme.Card
$pnlCreateAction.Location  = New-Object System.Drawing.Point(0, 0)  # positioned by layout
$pnlCreateAction.Height    = 80
$pnlCreateBody.Controls.Add($pnlCreateAction)

$ctrlCreate["ValidationHint"] = New-Object System.Windows.Forms.Label
$ctrlCreate["ValidationHint"].Text      = "Checks:  Name pending  |  Source pending  |  Network pending  |  User pending  |  Password pending"
$ctrlCreate["ValidationHint"].Size      = New-Object System.Drawing.Size(700, 20)
$ctrlCreate["ValidationHint"].Location  = New-Object System.Drawing.Point(14, 10)
$ctrlCreate["ValidationHint"].ForeColor = $theme.Muted
$ctrlCreate["ValidationHint"].AutoEllipsis = $true
$pnlCreateAction.Controls.Add($ctrlCreate["ValidationHint"])

$ctrlCreate["ModeHint"] = New-Object System.Windows.Forms.Label
$ctrlCreate["ModeHint"].Text      = "Mode: ISO Deploy — uses ISO, selected edition, and unattended setup."
$ctrlCreate["ModeHint"].Size      = New-Object System.Drawing.Size(700, 18)
$ctrlCreate["ModeHint"].Location  = New-Object System.Drawing.Point(14, 32)
$ctrlCreate["ModeHint"].ForeColor = $theme.Muted
$pnlCreateAction.Controls.Add($ctrlCreate["ModeHint"])

$ctrlCreate["CreateStatus"] = New-Object System.Windows.Forms.Label
$ctrlCreate["CreateStatus"].Text      = "Ready to create VM"
$ctrlCreate["CreateStatus"].Size      = New-Object System.Drawing.Size(440, 22)
$ctrlCreate["CreateStatus"].Location  = New-Object System.Drawing.Point(14, 52)
$ctrlCreate["CreateStatus"].ForeColor = $theme.Info
$ctrlCreate["CreateStatus"].AutoEllipsis = $true
$pnlCreateAction.Controls.Add($ctrlCreate["CreateStatus"])

$ctrlCreate["CreateProgress"] = New-Object System.Windows.Forms.ProgressBar
$ctrlCreate["CreateProgress"].Minimum = 0
$ctrlCreate["CreateProgress"].Maximum = 100
$ctrlCreate["CreateProgress"].Value   = 0
$ctrlCreate["CreateProgress"].Style   = 'Continuous'
$ctrlCreate["CreateProgress"].Size    = New-Object System.Drawing.Size(440, 12)
$ctrlCreate["CreateProgress"].Location = New-Object System.Drawing.Point(14, 56)
$pnlCreateAction.Controls.Add($ctrlCreate["CreateProgress"])

$btnCreateVM = New-Object System.Windows.Forms.Button
$btnCreateVM.Text      = "Create VM"
$btnCreateVM.Size      = New-Object System.Drawing.Size(150, 50)
$btnCreateVM.FlatStyle = 'Flat'
$btnCreateVM.BackColor = $theme.Accent
$btnCreateVM.ForeColor = [System.Drawing.Color]::White
$btnCreateVM.Font      = $script:FontBoldButton
$btnCreateVM.Cursor    = [System.Windows.Forms.Cursors]::Hand
$btnCreateVM.FlatAppearance.BorderColor        = $theme.AccentHover
$btnCreateVM.FlatAppearance.MouseOverBackColor = $theme.AccentHover
$btnCreateVM.FlatAppearance.MouseDownBackColor = $theme.AccentPressed
$pnlCreateAction.Controls.Add($btnCreateVM)

# ============================================================
#  PANEL 2:  GPU MANAGER
# ============================================================
$tabGPU = New-Object System.Windows.Forms.Panel
$tabGPU.Dock       = [System.Windows.Forms.DockStyle]::Fill
$tabGPU.BackColor  = $theme.Bg
$tabGPU.AutoScroll = $true
$tabGPU.Visible    = $false
$pnlContent.Controls.Add($tabGPU)

$ctrlGPU = @{}

# ── Panel header ────────────────────────────────────────────
$pnlGpuHeader = New-Object System.Windows.Forms.Panel
$pnlGpuHeader.Dock      = [System.Windows.Forms.DockStyle]::Top
$pnlGpuHeader.Height    = 44
$pnlGpuHeader.BackColor = $theme.Card

$lblGpuTitle = New-Object System.Windows.Forms.Label
$lblGpuTitle.Text      = "GPU Partition Manager"
$lblGpuTitle.Font      = $script:FontTabHeader
$lblGpuTitle.ForeColor = $theme.TextHigh
$lblGpuTitle.AutoSize  = $true
$lblGpuTitle.Location  = New-Object System.Drawing.Point(14, 8)
$pnlGpuHeader.Controls.Add($lblGpuTitle)

$lblGpuSub = New-Object System.Windows.Forms.Label
$lblGpuSub.Text      = "Select VMs, set target GPU share, and inject / update driver stacks"
$lblGpuSub.Font      = $script:FontSmall
$lblGpuSub.ForeColor = $theme.TextMuted
$lblGpuSub.AutoSize  = $true
$lblGpuSub.Location  = New-Object System.Drawing.Point(14, 26)
$pnlGpuHeader.Controls.Add($lblGpuSub)
$tabGPU.Controls.Add($pnlGpuHeader)

# ── Scrollable body ─────────────────────────────────────────
$pnlGpuBody = New-Object System.Windows.Forms.Panel
$pnlGpuBody.AutoScroll = $true
$pnlGpuBody.BackColor  = $theme.Bg
$pnlGpuBody.Location   = New-Object System.Drawing.Point(0, 44)
$tabGPU.Controls.Add($pnlGpuBody)

# ── Left: VM list ───────────────────────────────────────────
$grpVMs = New-ThemedGroupBox "Select VMs to Update" $pnlGpuBody
$grpVMs.Location = New-Object System.Drawing.Point(12, 12)
$grpVMs.Size     = New-Object System.Drawing.Size(360, 430)

$lblVmSearch = New-Object System.Windows.Forms.Label
$lblVmSearch.Text      = "Search:"
$lblVmSearch.AutoSize  = $true
$lblVmSearch.ForeColor = $theme.Text
$lblVmSearch.Location  = New-Object System.Drawing.Point(12, 26)
$grpVMs.Controls.Add($lblVmSearch)

$ctrlGPU["VmSearch"] = New-Object System.Windows.Forms.TextBox
$ctrlGPU["VmSearch"].Location   = New-Object System.Drawing.Point(58, 23)
$ctrlGPU["VmSearch"].Size       = New-Object System.Drawing.Size(190, 24)
$ctrlGPU["VmSearch"].BackColor  = $theme.Input
$ctrlGPU["VmSearch"].ForeColor  = $theme.Text
$ctrlGPU["VmSearch"].BorderStyle = 'FixedSingle'
$grpVMs.Controls.Add($ctrlGPU["VmSearch"])

$btnClearVmSearch = New-Object System.Windows.Forms.Button
$btnClearVmSearch.Text      = "Clear"
$btnClearVmSearch.Size      = New-Object System.Drawing.Size(60, 24)
$btnClearVmSearch.Location  = New-Object System.Drawing.Point(254, 23)
$btnClearVmSearch.FlatStyle = 'Flat'
$btnClearVmSearch.BackColor = $theme.Surface
$btnClearVmSearch.ForeColor = $theme.Text
$btnClearVmSearch.Cursor    = [System.Windows.Forms.Cursors]::Hand
$grpVMs.Controls.Add($btnClearVmSearch)

$vmPanel = New-Object System.Windows.Forms.Panel
$vmPanel.Location   = New-Object System.Drawing.Point(10, 55)
$vmPanel.Size       = New-Object System.Drawing.Size(336, 320)
$vmPanel.AutoScroll = $true
$vmPanel.BackColor  = $theme.Input
$vmPanel.BorderStyle = 'FixedSingle'
$grpVMs.Controls.Add($vmPanel)

$ctrlGPU["VMCheckboxes"] = @()

function Update-VMList {
    foreach ($existingCb in $ctrlGPU["VMCheckboxes"]) {
        if ($existingCb -and $existingCb.Text) {
            $script:GpuSelectedVMs[$existingCb.Text] = [bool]$existingCb.Checked
        }
    }
    $vmPanel.SuspendLayout()
    try {
        $vmPanel.Controls.Clear()
        $checkboxes = @()
        $filterText = ""
        if ($ctrlGPU.ContainsKey("VmSearch") -and $ctrlGPU["VmSearch"]) {
            $filterText = [string]$ctrlGPU["VmSearch"].Text
        }
        $allVms = @(Get-VM -ErrorAction SilentlyContinue | Sort-Object Name)
        $vms = if (-not [string]::IsNullOrWhiteSpace($filterText)) {
            $escapedFilter = [System.Management.Automation.WildcardPattern]::Escape($filterText)
            $allVms | Where-Object { $_.Name -like "*$escapedFilter*" }
        } else {
            $allVms
        }

        $currentVmNames = @($allVms | ForEach-Object { $_.Name })
        $staleKeys = @($script:GpuSelectedVMs.Keys | Where-Object { $_ -notin $currentVmNames })
        foreach ($k in $staleKeys) { [void]$script:GpuSelectedVMs.Remove($k) }

        $y = 6
        foreach ($vm in $vms) {
            $cb = New-Object System.Windows.Forms.CheckBox
            $cb.Text      = $vm.Name
            $cb.AutoSize  = $true
            $cb.Location  = New-Object System.Drawing.Point(8, $y)
            $cb.ForeColor = $theme.Text
            $cb.FlatStyle = 'Flat'
            if ($script:GpuSelectedVMs.ContainsKey($vm.Name)) {
                $cb.Checked = [bool]$script:GpuSelectedVMs[$vm.Name]
            }
            $cb.Add_CheckedChanged({
                if ($script:SuspendGpuSelectionEvents) { return }
                if ($this -and $this.Text) { $script:GpuSelectedVMs[$this.Text] = [bool]$this.Checked }
                if (Get-Command Update-GpuActionState -ErrorAction SilentlyContinue) { Update-GpuActionState }
            })
            if ($toolTip) {
                $toolTip.SetToolTip($cb, "Select this VM for GPU driver/adapter update.")
            }
            $vmPanel.Controls.Add($cb)
            $checkboxes += $cb
            $y += 26
        }
        $ctrlGPU["VMCheckboxes"] = $checkboxes
        if (Get-Command Update-GpuActionState -ErrorAction SilentlyContinue) { Update-GpuActionState }
    } finally {
        $vmPanel.ResumeLayout($true)
    }
}

$script:VmFilterTimer = New-Object System.Windows.Forms.Timer
$script:VmFilterTimer.Interval = 300
$script:VmFilterTimer.Add_Tick({ $script:VmFilterTimer.Stop(); Update-VMList })
$ctrlGPU["VmSearch"].Add_TextChanged({ $script:VmFilterTimer.Stop(); $script:VmFilterTimer.Start() })
$btnClearVmSearch.Add_Click({ $ctrlGPU["VmSearch"].Text = ""; $script:VmFilterTimer.Stop(); Update-VMList })

# Select All / None / Refresh
$btnSelectAll = New-Object System.Windows.Forms.Button
$btnSelectAll.Text      = "All"
$btnSelectAll.Size      = New-Object System.Drawing.Size(60, 26)
$btnSelectAll.Location  = New-Object System.Drawing.Point(10, 384)
$btnSelectAll.FlatStyle = 'Flat'
$btnSelectAll.BackColor = $theme.Surface
$btnSelectAll.ForeColor = $theme.Text
$btnSelectAll.Cursor    = [System.Windows.Forms.Cursors]::Hand
$btnSelectAll.Add_Click({
    $script:SuspendGpuSelectionEvents = $true
    try { foreach ($cb in $ctrlGPU["VMCheckboxes"]) { $cb.Checked = $true; if ($cb -and $cb.Text) { $script:GpuSelectedVMs[$cb.Text] = $true } } }
    finally { $script:SuspendGpuSelectionEvents = $false }
    Update-GpuActionState
})
$grpVMs.Controls.Add($btnSelectAll)

$btnSelectNone = New-Object System.Windows.Forms.Button
$btnSelectNone.Text      = "None"
$btnSelectNone.Size      = New-Object System.Drawing.Size(60, 26)
$btnSelectNone.Location  = New-Object System.Drawing.Point(76, 384)
$btnSelectNone.FlatStyle = 'Flat'
$btnSelectNone.BackColor = $theme.Surface
$btnSelectNone.ForeColor = $theme.Text
$btnSelectNone.Cursor    = [System.Windows.Forms.Cursors]::Hand
$btnSelectNone.Add_Click({
    $script:SuspendGpuSelectionEvents = $true
    try { foreach ($cb in $ctrlGPU["VMCheckboxes"]) { $cb.Checked = $false; if ($cb -and $cb.Text) { $script:GpuSelectedVMs[$cb.Text] = $false } } }
    finally { $script:SuspendGpuSelectionEvents = $false }
    Update-GpuActionState
})
$grpVMs.Controls.Add($btnSelectNone)

$btnRefreshVMs = New-Object System.Windows.Forms.Button
$btnRefreshVMs.Text      = "Refresh"
$btnRefreshVMs.Size      = New-Object System.Drawing.Size(70, 26)
$btnRefreshVMs.Location  = New-Object System.Drawing.Point(142, 384)
$btnRefreshVMs.FlatStyle = 'Flat'
$btnRefreshVMs.BackColor = $theme.Surface
$btnRefreshVMs.ForeColor = $theme.Text
$btnRefreshVMs.Cursor    = [System.Windows.Forms.Cursors]::Hand
$btnRefreshVMs.Add_Click({ Update-VMList })
$grpVMs.Controls.Add($btnRefreshVMs)

# ── Right: GPU Settings ──────────────────────────────────────
$grpGPUSettings = New-ThemedGroupBox "GPU-P Settings" $pnlGpuBody
$grpGPUSettings.Location = New-Object System.Drawing.Point(384, 12)
$grpGPUSettings.Size     = New-Object System.Drawing.Size(480, 210)

$lblGpuSelector = New-Object System.Windows.Forms.Label
$lblGpuSelector.Text      = "GPU:"
$lblGpuSelector.AutoSize  = $true
$lblGpuSelector.ForeColor = $theme.Text
$lblGpuSelector.Location  = New-Object System.Drawing.Point(14, 28)
$grpGPUSettings.Controls.Add($lblGpuSelector)

$ctrlGPU["GpuSelector"] = New-Object System.Windows.Forms.ComboBox
$ctrlGPU["GpuSelector"].DropDownStyle = 'DropDownList'
$ctrlGPU["GpuSelector"].Width         = 360
$ctrlGPU["GpuSelector"].Location      = New-Object System.Drawing.Point(54, 25)
$ctrlGPU["GpuSelector"].BackColor     = $theme.Input
$ctrlGPU["GpuSelector"].ForeColor     = $theme.Text
$ctrlGPU["GpuSelector"].FlatStyle     = 'Flat'
$grpGPUSettings.Controls.Add($ctrlGPU["GpuSelector"])
$ctrlGPU["GpuSelector"].Add_SelectedIndexChanged({
    if (Get-Command Update-GpuActionState -ErrorAction SilentlyContinue) { Update-GpuActionState }
})

$script:GpuPList = Get-GpuPProviders
$script:GpuPList | ForEach-Object { [void]$ctrlGPU["GpuSelector"].Items.Add($_.Friendly) }
if ($ctrlGPU["GpuSelector"].Items.Count -gt 1) {
    # Default to first actual GPU (index 1) instead of "NONE - Remove GPU Adapter" (index 0)
    $ctrlGPU["GpuSelector"].SelectedIndex = 1
} elseif ($ctrlGPU["GpuSelector"].Items.Count -gt 0) {
    $ctrlGPU["GpuSelector"].SelectedIndex = 0
}

if (-not $script:SupportsGpuInstancePath) {
    $lblGpuWarn = New-Object System.Windows.Forms.Label
    $lblGpuWarn.Text      = "Note: This host uses default GPU selection for GPU-P."
    $lblGpuWarn.AutoSize  = $true
    $lblGpuWarn.ForeColor = $theme.Warning
    $lblGpuWarn.Font      = $script:FontSmall
    $lblGpuWarn.Location  = New-Object System.Drawing.Point(14, 56)
    $grpGPUSettings.Controls.Add($lblGpuWarn)
}

New-Divider $grpGPUSettings 14 72 446

$lblCopyMode = New-Object System.Windows.Forms.Label
$lblCopyMode.Text      = "Driver Copy Mode:"
$lblCopyMode.AutoSize  = $true
$lblCopyMode.ForeColor = $theme.Text
$lblCopyMode.Location  = New-Object System.Drawing.Point(14, 82)
$grpGPUSettings.Controls.Add($lblCopyMode)

$ctrlGPU["SmartCopy"] = New-Object System.Windows.Forms.RadioButton
$ctrlGPU["SmartCopy"].Text      = "Smart  (GPU folders only — recommended)"
$ctrlGPU["SmartCopy"].AutoSize  = $true
$ctrlGPU["SmartCopy"].Checked   = $true
$ctrlGPU["SmartCopy"].Location  = New-Object System.Drawing.Point(22, 104)
$ctrlGPU["SmartCopy"].ForeColor = $theme.Text
$grpGPUSettings.Controls.Add($ctrlGPU["SmartCopy"])

$ctrlGPU["FullCopy"] = New-Object System.Windows.Forms.RadioButton
$ctrlGPU["FullCopy"].Text      = "Full DriverStore  (larger, may need big VHD)"
$ctrlGPU["FullCopy"].AutoSize  = $true
$ctrlGPU["FullCopy"].Location  = New-Object System.Drawing.Point(22, 128)
$ctrlGPU["FullCopy"].ForeColor = $theme.Text
$grpGPUSettings.Controls.Add($ctrlGPU["FullCopy"])

New-Divider $grpGPUSettings 14 156 446

$lblGpuAlloc = New-Object System.Windows.Forms.Label
$lblGpuAlloc.Text      = "Target GPU Share:"
$lblGpuAlloc.AutoSize  = $true
$lblGpuAlloc.ForeColor = $theme.Text
$lblGpuAlloc.Location  = New-Object System.Drawing.Point(14, 166)
$grpGPUSettings.Controls.Add($lblGpuAlloc)

$ctrlGPU["GpuAllocSlider"] = New-Object System.Windows.Forms.TrackBar
$ctrlGPU["GpuAllocSlider"].Minimum       = 10
$ctrlGPU["GpuAllocSlider"].Maximum       = 100
$ctrlGPU["GpuAllocSlider"].Value         = 80
$ctrlGPU["GpuAllocSlider"].TickFrequency = 10
$ctrlGPU["GpuAllocSlider"].SmallChange   = 5
$ctrlGPU["GpuAllocSlider"].LargeChange   = 10
$ctrlGPU["GpuAllocSlider"].Location      = New-Object System.Drawing.Point(158, 158)
$ctrlGPU["GpuAllocSlider"].Size          = New-Object System.Drawing.Size(240, 34)
$ctrlGPU["GpuAllocSlider"].BackColor     = $theme.Surface
$grpGPUSettings.Controls.Add($ctrlGPU["GpuAllocSlider"])

$ctrlGPU["GpuAllocLabel"] = New-Object System.Windows.Forms.Label
$ctrlGPU["GpuAllocLabel"].Text      = "80%"
$ctrlGPU["GpuAllocLabel"].AutoSize  = $true
$ctrlGPU["GpuAllocLabel"].Location  = New-Object System.Drawing.Point(406, 166)
$ctrlGPU["GpuAllocLabel"].ForeColor = $theme.Success
$ctrlGPU["GpuAllocLabel"].Font      = $script:FontBoldLabel
$grpGPUSettings.Controls.Add($ctrlGPU["GpuAllocLabel"])
$ctrlGPU["GpuAllocSlider"].Add_ValueChanged({ $ctrlGPU["GpuAllocLabel"].Text = "$($ctrlGPU['GpuAllocSlider'].Value)%" })

# ── Right: GPU Options ───────────────────────────────────────
$grpGPUOpts = New-ThemedGroupBox "GPU Update Options" $pnlGpuBody
$grpGPUOpts.Location = New-Object System.Drawing.Point(384, 234)
$grpGPUOpts.Size     = New-Object System.Drawing.Size(480, 140)

foreach ($chkDef in @(
    @{ Key = "StartVM";       Text = "Start VM after update";                         X = 14; Y = 26; Default = $false },
    @{ Key = "AutoExpand";    Text = "Auto-expand VHD if insufficient space";         X = 14; Y = 56; Default = $true  },
    @{ Key = "CopySvcDriver"; Text = "Copy GPU service driver (recommended)";         X = 14; Y = 86; Default = $true  },
    @{ Key = "StrictChecks";  Text = "Strict GPU safety checks";                     X = 246;Y = 26; Default = $true  }
)) {
    $cb = New-Object System.Windows.Forms.CheckBox
    $cb.Text      = $chkDef.Text
    $cb.AutoSize  = $true
    $cb.Checked   = $chkDef.Default
    $cb.Location  = New-Object System.Drawing.Point($chkDef.X, $chkDef.Y)
    $cb.ForeColor = $theme.Text
    $cb.FlatStyle = 'Flat'
    $grpGPUOpts.Controls.Add($cb)
    $ctrlGPU[$chkDef.Key] = $cb
}

# ── GPU action row ───────────────────────────────────────────
$btnUpdateGPU = New-Object System.Windows.Forms.Button
$btnUpdateGPU.Text      = "Update GPU Drivers"
$btnUpdateGPU.Size      = New-Object System.Drawing.Size(180, 42)
$btnUpdateGPU.Location  = New-Object System.Drawing.Point(384, 388)
$btnUpdateGPU.FlatStyle = 'Flat'
$btnUpdateGPU.BackColor = $theme.Success
$btnUpdateGPU.ForeColor = [System.Drawing.Color]::White
$btnUpdateGPU.Font      = $script:FontBoldButton
$btnUpdateGPU.Cursor    = [System.Windows.Forms.Cursors]::Hand
$btnUpdateGPU.FlatAppearance.BorderColor        = $theme.SuccessHover
$btnUpdateGPU.FlatAppearance.MouseOverBackColor = $theme.SuccessHover
$btnUpdateGPU.FlatAppearance.MouseDownBackColor = $theme.SuccessHover
$pnlGpuBody.Controls.Add($btnUpdateGPU)

$ctrlGPU["SelectionHint"] = New-Object System.Windows.Forms.Label
$ctrlGPU["SelectionHint"].Text      = "Select at least one VM to enable GPU update."
$ctrlGPU["SelectionHint"].Size      = New-Object System.Drawing.Size(480, 28)
$ctrlGPU["SelectionHint"].Location  = New-Object System.Drawing.Point(384, 440)
$ctrlGPU["SelectionHint"].ForeColor = $theme.Muted
$ctrlGPU["SelectionHint"].AutoEllipsis = $true
$pnlGpuBody.Controls.Add($ctrlGPU["SelectionHint"])

$ctrlGPU["GpuStatus"] = New-Object System.Windows.Forms.Label
$ctrlGPU["GpuStatus"].Text      = ""
$ctrlGPU["GpuStatus"].Size      = New-Object System.Drawing.Size(480, 28)
$ctrlGPU["GpuStatus"].Location  = New-Object System.Drawing.Point(384, 470)
$ctrlGPU["GpuStatus"].ForeColor = $theme.Info
$ctrlGPU["GpuStatus"].AutoEllipsis = $true
$pnlGpuBody.Controls.Add($ctrlGPU["GpuStatus"])

# ── Software & Settings (on Create VM tab, right column) ─────
$grpSoft = New-ThemedGroupBox "Software && Settings" $pnlCreateBody
$grpSoft.Location = New-Object System.Drawing.Point(496, ($grpOpts.Bottom + 10))
$grpSoft.Size     = New-Object System.Drawing.Size(480, 280)
$grpSoft.Anchor   = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left

$softItems = @(
    @{ Key = "Parsec";            Text = "Parsec (Per Computer)";          X = 14;  Y = 28  },
    @{ Key = "VBCable";           Text = "VB-Audio Cable";                 X = 270; Y = 28  },
    @{ Key = "USBMMIDD";          Text = "Virtual Display Driver";         X = 14;  Y = 58  },
    @{ Key = "RDP";               Text = "Remote Desktop";                 X = 270; Y = 58  },
    @{ Key = "Share";             Text = "Share Folder";                   X = 14;  Y = 88  },
    @{ Key = "PauseUpdate";       Text = "Pause Windows Updates";          X = 270; Y = 88  },
    @{ Key = "FullUpdate";        Text = "Full Windows Updates";           X = 14;  Y = 118 },
    @{ Key = "NestedVirt";        Text = "Nested Virtualization";          X = 270; Y = 118 },
    @{ Key = "NestedNetFollowup"; Text = "Nested Net (MAC spoofing)";      X = 14;  Y = 148 },
    @{ Key = "ResetBootOrder";    Text = "Reset boot order after recovery";X = 270; Y = 148 },
    @{ Key = "GoldenImage";       Text = "Create from Golden VHDX";       X = 14;  Y = 178 }
)
foreach ($sw in $softItems) {
    $cb = New-Object System.Windows.Forms.CheckBox
    $cb.Text      = $sw.Text
    $cb.AutoSize  = $true
    $cb.Location  = New-Object System.Drawing.Point($sw.X, $sw.Y)
    $cb.ForeColor = $theme.Text
    $cb.FlatStyle = 'Flat'
    $grpSoft.Controls.Add($cb)
    $ctrlCreate[$sw.Key] = $cb
}

$goldenLabel = New-Object System.Windows.Forms.Label
$goldenLabel.Text      = "Parent VHDX:"
$goldenLabel.ForeColor = $theme.Text
$goldenLabel.AutoSize  = $false
$goldenLabel.Size      = New-Object System.Drawing.Size(100, 22)
$goldenLabel.Location  = New-Object System.Drawing.Point(14, 216)
$goldenLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$grpSoft.Controls.Add($goldenLabel)

$ctrlCreate["GoldenParentVHD"] = New-Object System.Windows.Forms.TextBox
$ctrlCreate["GoldenParentVHD"].Location   = New-Object System.Drawing.Point(118, 214)
$ctrlCreate["GoldenParentVHD"].Size       = New-Object System.Drawing.Size(260, 24)
$ctrlCreate["GoldenParentVHD"].BackColor  = $theme.Input
$ctrlCreate["GoldenParentVHD"].ForeColor  = $theme.Text
$ctrlCreate["GoldenParentVHD"].BorderStyle = 'FixedSingle'
$grpSoft.Controls.Add($ctrlCreate["GoldenParentVHD"])

$btnBrowseGoldenEnv = New-Object System.Windows.Forms.Button
$btnBrowseGoldenEnv.Text      = "Browse"
$btnBrowseGoldenEnv.Size      = New-Object System.Drawing.Size(64, 24)
$btnBrowseGoldenEnv.Location  = New-Object System.Drawing.Point(388, 214)
$btnBrowseGoldenEnv.FlatStyle = 'Flat'
$btnBrowseGoldenEnv.BackColor = $theme.Surface
$btnBrowseGoldenEnv.ForeColor = $theme.Text
$btnBrowseGoldenEnv.Cursor    = [System.Windows.Forms.Cursors]::Hand
$grpSoft.Controls.Add($btnBrowseGoldenEnv)
$btnBrowseGoldenEnv.Add_Click({
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    try {
        $dlg.Filter = "Virtual Hard Disk|*.vhdx;*.vhd"
        $dlg.Title  = "Select parent Golden VHDX/VHD"
        if ($dlg.ShowDialog() -eq 'OK') { $ctrlCreate["GoldenParentVHD"].Text = $dlg.FileName }
    } finally { $dlg.Dispose() }
})

# Alias: keep $btnBrowseGolden pointing to the browse button used by existing handlers
$btnBrowseGolden = $btnBrowseGoldenEnv

# ============================================================
#  NAVIGATION SWITCH LOGIC
# ============================================================
$script:NavPanels    = @($tabCreate, $tabGPU)
$script:ActiveNavIdx = 0

function Switch-NavPanel {
    param([int]$Index)
    for ($i = 0; $i -lt $script:NavPanels.Count; $i++) {
        $script:NavPanels[$i].Visible = ($i -eq $Index)
    }
    $script:ActiveNavIdx = $Index

    for ($i = 0; $i -lt $script:NavButtons.Count; $i++) {
        if ($i -eq $Index) {
            $script:NavButtons[$i].BackColor = $theme.SidebarActive
            $script:NavButtons[$i].ForeColor = [System.Drawing.Color]::White
        } else {
            $script:NavButtons[$i].BackColor = $theme.Sidebar
            $script:NavButtons[$i].ForeColor = $theme.TextMuted
        }
    }

    $form.AcceptButton = switch ($Index) {
        0 { $btnCreateVM }
        1 { $btnUpdateGPU }
        default { $null }
    }
}

$btnNavCreate.Add_Click({ Switch-NavPanel -Index 0 })
$btnNavGPU.Add_Click({    Switch-NavPanel -Index 1 })

# Activate Create VM by default
Switch-NavPanel -Index 0

# ============================================================
#  TABCONTROL COMPATIBILITY SHIM
#  (Event handlers reference $tabControl.Enabled,
#   .SelectedIndex, and .SelectedTab — shim keeps them working)
# ============================================================
$tabControl = New-Object PSObject
$tabControl | Add-Member -MemberType NoteProperty -Name "_Enabled"   -Value $true
$tabControl | Add-Member -MemberType NoteProperty -Name "_SelIdx"    -Value 0
$tabControl | Add-Member -MemberType NoteProperty -Name "SelectedTab" -Value $tabCreate
$tabControl | Add-Member -MemberType ScriptProperty -Name "Enabled" `
    -Value       { return $this._Enabled } `
    -SecondValue {
        param($v)
        $this._Enabled = $v
        foreach ($p in $script:NavPanels) { $p.Enabled = $v }
        foreach ($nb in $script:NavButtons) { $nb.Enabled = $v }
    }
$tabControl | Add-Member -MemberType ScriptProperty -Name "SelectedIndex" `
    -Value       { return $this._SelIdx } `
    -SecondValue {
        param($v)
        $this._SelIdx = $v
        $this.SelectedTab = $script:NavPanels[[Math]::Max(0, [Math]::Min($v, $script:NavPanels.Count - 1))]
        Switch-NavPanel -Index $v
    }

# ============================================================
#  ADAPTIVE BODY PANELS (resize to fill the scrollable body)
# ============================================================
function Update-MainLayout {
    param([System.Windows.Forms.Form]$RootForm)
    if (-not $RootForm -or $RootForm.IsDisposed) { return }
    try {
        $RootForm.SuspendLayout()

        # Header status label: keep near right side of header
        if ($script:HeaderStatusLabel -and $pnlHeader) {
            $script:HeaderStatusLabel.Location = New-Object System.Drawing.Point(
                [Math]::Max(360, ($pnlHeader.Width - $script:HeaderStatusLabel.Width - 24)), 18)
        }

        # Body panels fill the tab panel's client area
        $contentW = $pnlContent.ClientSize.Width
        $contentH = $pnlContent.ClientSize.Height

        foreach ($panel in @($tabCreate, $tabGPU)) {
            $panel.Size = New-Object System.Drawing.Size($contentW, $contentH)
        }

        # Create VM body fills remaining below its header
        if ($pnlCreateBody) {
            $pnlCreateBody.Size = New-Object System.Drawing.Size($contentW, [Math]::Max(100, $contentH - 44))
        }
        # GPU body
        if ($pnlGpuBody) {
            $pnlGpuBody.Size = New-Object System.Drawing.Size($contentW, [Math]::Max(100, $contentH - 44))
        }

        # Reposition action strip at bottom of create body
        if ($pnlCreateAction -and $pnlCreateBody) {
            $actionY = [Math]::Max([Math]::Max($grpSoft.Bottom, $grpOpts.Bottom), $grpConfig.Bottom) + 14
            $pnlCreateAction.Location = New-Object System.Drawing.Point(0, $actionY)
            $pnlCreateAction.Width    = [Math]::Max(600, $pnlCreateBody.ClientSize.Width)

            # Create VM button: right side of action panel
            $btnW = $btnCreateVM.Width
            $btnCreateVM.Location = New-Object System.Drawing.Point(
                [Math]::Max(500, $pnlCreateAction.Width - $btnW - 20), 14)
        }

        # Adapt grpConfig width to available space
        if ($pnlCreateBody -and $pnlCreateBody.ClientSize.Width -gt 0) {
            $availW = $pnlCreateBody.ClientSize.Width
            $cfgRight = [Math]::Min(480, [int]($availW * 0.48))
            $grpConfig.Width = [Math]::Max(440, $cfgRight)

            $rightX   = $grpConfig.Right + 12
            $rightW   = [Math]::Max(440, $availW - $rightX - 12)
            $grpBoot.Location = New-Object System.Drawing.Point($rightX, 12)
            $grpBoot.Width    = $rightW
            $grpOpts.Location = New-Object System.Drawing.Point($rightX, ($grpBoot.Bottom + 10))
            $grpOpts.Width    = $rightW
            $grpSoft.Location = New-Object System.Drawing.Point($rightX, ($grpOpts.Bottom + 10))
            $grpSoft.Width    = $rightW
            $ctrlCreate["GoldenParentVHD"].Width = [Math]::Max(200, $grpSoft.Width - 220)
            $btnBrowseGoldenEnv.Location = New-Object System.Drawing.Point(
                [Math]::Max(320, $ctrlCreate["GoldenParentVHD"].Right + 8), $btnBrowseGoldenEnv.Top)

            # Resize browse buttons to stay near right edge of grpConfig
            $browseRightX = [Math]::Max(360, $grpConfig.Width - 72)
            if ($btnBrowseVM) {
                $btnBrowseVM.Location  = New-Object System.Drawing.Point($browseRightX, $btnBrowseVM.Top)
                $ctrlCreate["VMLocation"].Width = [Math]::Max(140, $browseRightX - $ctrlCreate["VMLocation"].Left - 8)
            }
            if ($btnBrowseISO) {
                $btnBrowseISO.Location = New-Object System.Drawing.Point($browseRightX, $btnBrowseISO.Top)
                $ctrlCreate["ISOPath"].Width    = [Math]::Max(140, $browseRightX - $ctrlCreate["ISOPath"].Left - 8)
            }
        }

        # Adapt GPU panel widths
        if ($pnlGpuBody -and $pnlGpuBody.ClientSize.Width -gt 0) {
            $gpuW      = $pnlGpuBody.ClientSize.Width
            $vmListW   = [Math]::Min(360, [int]($gpuW * 0.37))
            $grpVMs.Width = $vmListW
            $vmPanel.Width = [Math]::Max(180, $vmListW - 24)

            $gpuRightX = $grpVMs.Right + 12
            $gpuRightW = [Math]::Max(380, $gpuW - $gpuRightX - 12)
            foreach ($gpuCtrl in @($grpGPUSettings, $grpGPUOpts)) {
                $gpuCtrl.Location = New-Object System.Drawing.Point($gpuRightX, $gpuCtrl.Top)
                $gpuCtrl.Width    = $gpuRightW
            }
            $ctrlGPU["GpuSelector"].Width = [Math]::Max(180, $grpGPUSettings.Width - 70)
            $allocLabelX = [Math]::Max(8, $grpGPUSettings.Width - 62)
            $ctrlGPU["GpuAllocLabel"].Location = New-Object System.Drawing.Point($allocLabelX, $ctrlGPU["GpuAllocLabel"].Top)
            $sliderRight = $ctrlGPU["GpuAllocLabel"].Left - 14
            $ctrlGPU["GpuAllocSlider"].Width = [Math]::Max(120, $sliderRight - $ctrlGPU["GpuAllocSlider"].Left)

            $btnUpdateGPU.Location = New-Object System.Drawing.Point($gpuRightX, 388)
            if ($ctrlGPU["SelectionHint"]) { $ctrlGPU["SelectionHint"].Location = New-Object System.Drawing.Point($gpuRightX, 440) }
            if ($ctrlGPU["GpuStatus"])     { $ctrlGPU["GpuStatus"].Location     = New-Object System.Drawing.Point($gpuRightX, 470) }
        }

    } catch {
        Write-UiWarning "Layout adjustment warning: $($_.Exception.Message)"
    } finally {
        $RootForm.ResumeLayout($true)
    }
}

$script:LayoutTimer = New-Object System.Windows.Forms.Timer
$script:LayoutTimer.Interval = 100
$script:LayoutTimer.Add_Tick({
    $script:LayoutTimer.Stop()
    try { Update-MainLayout -RootForm $form } catch { Write-UiWarning "Deferred layout: $($_.Exception.Message)" }
})

$form.Add_Shown({
    try { Update-MainLayout -RootForm $form } catch { Write-UiWarning "Shown layout: $($_.Exception.Message)" }
})
$form.Add_Resize({
    $script:LayoutTimer.Stop()
    $script:LayoutTimer.Start()
})
try {
    $form.Add_DpiChanged({
        $script:LayoutTimer.Stop()
        $script:LayoutTimer.Start()
    })
} catch {
    Write-StartupTrace -Message "DpiChanged event not available on this .NET runtime (requires .NET 4.7+)" -Level 'WARN'
}

# ============================================================
#  BUTTON HOVER HELPERS  (kept for compatibility)
# ============================================================
function Set-ButtonHover {
    param(
        [System.Windows.Forms.Button]$Button,
        [System.Drawing.Color]$Normal,
        [System.Drawing.Color]$Hover
    )
    if (-not $Button) { return }
    [System.Drawing.Color]$normalColor = $Normal
    [System.Drawing.Color]$hoverColor  = $Hover
    if ($null -eq $normalColor -or $normalColor.IsEmpty) { $normalColor = [System.Drawing.SystemColors]::ControlDarkDark }
    if ($null -eq $hoverColor  -or $hoverColor.IsEmpty)  { $hoverColor  = $normalColor }
    $Button.BackColor = $normalColor
    $Button.FlatStyle = 'Flat'
    $Button.FlatAppearance.BorderSize             = 0
    $Button.FlatAppearance.MouseDownBackColor     = $hoverColor
    $Button.FlatAppearance.MouseOverBackColor     = $hoverColor
    $Button.Cursor = [System.Windows.Forms.Cursors]::Hand
    $Button.Add_MouseEnter({ if (-not $hoverColor.IsEmpty)  { $this.BackColor = $hoverColor  } }.GetNewClosure())
    $Button.Add_MouseLeave({ if (-not $normalColor.IsEmpty) { $this.BackColor = $normalColor } }.GetNewClosure())
}

function Set-ModernTheme {
    param([System.Windows.Forms.Control]$Root)
    foreach ($control in $Root.Controls) {
        switch ($control.GetType().Name) {
            'GroupBox' {
                $control.BackColor = $theme.Surface
                $control.ForeColor = $theme.Text
                $control.Font      = $script:ThemeFontGroupBox
            }
            'Label' {
                if ($control.BackColor -ne $theme.Surface -and $control.BackColor -ne $theme.Card) {
                    if ($control.ForeColor -eq [System.Drawing.Color]::White) { $control.ForeColor = $theme.TextHigh }
                }
            }
            'TextBox' {
                $control.BackColor  = $theme.Input
                $control.ForeColor  = $theme.Text
                $control.BorderStyle = 'FixedSingle'
            }
            'ComboBox' {
                $control.BackColor = $theme.Input
                $control.ForeColor = $theme.Text
                $control.FlatStyle = 'Flat'
            }
            'NumericUpDown' {
                $control.BackColor = $theme.Input
                $control.ForeColor = $theme.Text
            }
            'Panel' {
                # Only override if it's still default
                if ($control.BackColor -eq [System.Drawing.SystemColors]::Control) {
                    $control.BackColor = $theme.Surface
                }
            }
            'CheckBox' {
                $control.ForeColor = $theme.Text
                $control.FlatStyle = 'Flat'
            }
            'RadioButton' {
                $control.ForeColor = $theme.Text
            }
            'TrackBar' {
                $control.BackColor = $theme.Surface
            }
            'ProgressBar' {
                $control.ForeColor = $theme.Accent
            }
        }
        if ($control.HasChildren) { Set-ModernTheme -Root $control }
    }
}

# Apply theme
Set-ModernTheme -Root $form

# Double-buffer key containers to reduce flicker
function Enable-ControlDoubleBuffer {
    param([System.Windows.Forms.Control]$Control)
    try {
        $prop = $Control.GetType().GetProperty('DoubleBuffered',
            [System.Reflection.BindingFlags]'NonPublic,Instance')
        if ($prop) { $prop.SetValue($Control, $true, $null) }
    } catch {
        Write-StartupTrace -Message "DoubleBuffer not supported on $($Control.GetType().Name): $($_.Exception.Message)" -Level 'WARN'
    }
}

Enable-ControlDoubleBuffer -Control $form
Enable-ControlDoubleBuffer -Control $pnlContent
Enable-ControlDoubleBuffer -Control $pnlCreateBody
Enable-ControlDoubleBuffer -Control $pnlGpuBody
Enable-ControlDoubleBuffer -Control $vmPanel
Enable-ControlDoubleBuffer -Control $script:LogBox

# Accept/Cancel buttons
$form.AcceptButton = $btnCreateVM

# ============================================================
#  TOOLTIP SETUP
# ============================================================
$toolTip = New-Object System.Windows.Forms.ToolTip
$toolTip.AutoPopDelay = 8000
$toolTip.InitialDelay = 300
$toolTip.ReshowDelay  = 120
$toolTip.ShowAlways   = $true

$toolTip.SetToolTip($ctrlCreate["VMName"],            "Short unique VM name. Letters, numbers, hyphens only. Cannot start a digit or end with a hyphen.")
$toolTip.SetToolTip($ctrlCreate["ISOPath"],           "Windows installation ISO for unattended deploy mode.")
$toolTip.SetToolTip($ctrlCreate["GoldenParentVHD"],   "Parent image used when Golden mode is enabled (differencing disk).")
$toolTip.SetToolTip($ctrlCreate["CheckpointMode"],    "Production/ProductionOnly recommended over Standard for stable rollback.")
$toolTip.SetToolTip($ctrlCreate["SecureBoot"],        "Enables UEFI Secure Boot. Required for Windows 11, recommended for modern images.")
$toolTip.SetToolTip($ctrlCreate["TPM"],               "Adds a virtual TPM. Required for Windows 11.")
$toolTip.SetToolTip($ctrlCreate["VHDType"],           "Fixed-size VHD: better I/O predictability at cost of upfront disk space.")
$toolTip.SetToolTip($ctrlCreate["DynamicMem"],        "Enable Hyper-V Dynamic Memory ballooning.")
$toolTip.SetToolTip($ctrlCreate["DynamicMemMin"],     "Lowest RAM the VM can shrink to with Dynamic Memory.")
$toolTip.SetToolTip($ctrlCreate["DynamicMemMax"],     "Highest RAM the VM can grow to with Dynamic Memory.")
$toolTip.SetToolTip($ctrlCreate["EnhancedSession"],   "Enhanced Session: clipboard, dynamic display, device redirection.")
$toolTip.SetToolTip($ctrlCreate["StartVM"],           "Starts the VM automatically after provisioning.")
$toolTip.SetToolTip($ctrlCreate["StrictLegacyMode"],  "Forces legacy-safe DISM apply for older/custom Windows images.")
$toolTip.SetToolTip($ctrlCreate["AutoCreateSwitch"],  "Auto-creates an internal NAT switch when no switch is selected.")
$toolTip.SetToolTip($ctrlCreate["EnableMetering"],    "Enables Hyper-V resource metering.")
$toolTip.SetToolTip($ctrlCreate["EnableAutoLogon"],   "Configures automatic logon during initial setup reboots (LogonCount=999).")
$toolTip.SetToolTip($ctrlCreate["Parsec"],            "Downloads and installs Parsec silently in SetupComplete.")
$toolTip.SetToolTip($ctrlCreate["VBCable"],           "Installs VB-Audio virtual cable for guest audio routing.")
$toolTip.SetToolTip($ctrlCreate["USBMMIDD"],          "Installs a virtual display driver for headless remote sessions.")
$toolTip.SetToolTip($ctrlCreate["RDP"],               "Enables Remote Desktop and firewall rules in the guest.")
$toolTip.SetToolTip($ctrlCreate["Share"],             "Creates a Desktop\\share folder in the guest. Lab use only.")
$toolTip.SetToolTip($ctrlCreate["PauseUpdate"],       "Pauses Windows Updates for an extended period.")
$toolTip.SetToolTip($ctrlCreate["FullUpdate"],        "Runs full Windows Update at first logon via PSWindowsUpdate.")
$toolTip.SetToolTip($ctrlCreate["NestedVirt"],        "Exposes virtualization extensions for Hyper-V/WSL2 nesting.")
$toolTip.SetToolTip($ctrlCreate["NestedNetFollowup"], "Enables MAC spoofing for nested networking. Requires Nested Virt.")
$toolTip.SetToolTip($ctrlCreate["ResetBootOrder"],    "Resets firmware boot order back to VHD after ISO recovery boot.")
$toolTip.SetToolTip($ctrlCreate["GoldenImage"],       "Creates the VM from a parent Golden VHDX differencing disk.")
$toolTip.SetToolTip($ctrlGPU["StartVM"],              "Starts selected VMs after GPU adapter update.")
$toolTip.SetToolTip($ctrlGPU["AutoExpand"],           "Automatically expands guest VHD if insufficient space during driver copy.")
$toolTip.SetToolTip($ctrlGPU["CopySvcDriver"],        "Copies GPU service driver components to guest HostDriverStore.")
$toolTip.SetToolTip($ctrlGPU["StrictChecks"],         "Enforces additional GPU-P safety checks.")
$toolTip.SetToolTip($ctrlGPU["GpuAllocSlider"],       "Controls target GPU share. Toolkit applies this as conservative Optimal partition values for stability.")
$toolTip.SetToolTip($btnUpdateGPU,                    "Inject/update GPU-P drivers into selected VMs.")
$toolTip.SetToolTip($ctrlGPU["VmSearch"],             "Type to filter VMs by name.")
$toolTip.SetToolTip($btnCreateVM,                     "Start VM provisioning with validated settings.")
$toolTip.SetToolTip($btnClearLog,                     "Clear the on-screen log output.")
$toolTip.SetToolTip($btnSaveLog,                      "Export the current log output to a file.")
$toolTip.SetToolTip($btnExit,                         "Close the toolkit and run image mount cleanup.")


# Populate VM checkbox list now that tooltips are registered
Update-VMList

# ============================================================
#  KEYBOARD SHORTCUTS
# ============================================================
function Get-ActiveFocusedControl {
    param([System.Windows.Forms.Control]$RootControl)
    $current = $RootControl
    while ($current -and $current.ContainsFocus -and
           $current -is [System.Windows.Forms.ContainerControl] -and
           $current.ActiveControl) {
        $current = $current.ActiveControl
    }
    return $current
}

$form.Add_KeyDown({
    param($eventSourceControl, $e)
    [void]$eventSourceControl

    if ($e.KeyCode -eq [System.Windows.Forms.Keys]::Escape) {
        $result = [System.Windows.Forms.MessageBox]::Show(
            "Are you sure you want to close the toolkit?",
            "Confirm Exit",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question)
        if ($result -eq [System.Windows.Forms.DialogResult]::Yes) { $form.Close() }
        $e.SuppressKeyPress = $true
        return
    }

    if (-not $e.Control) { return }

    $focused = Get-ActiveFocusedControl -RootControl $form
    if ($focused -is [System.Windows.Forms.TextBox] -or
        $focused -is [System.Windows.Forms.TextBoxBase]) { return }

    if ($e.KeyCode -eq [System.Windows.Forms.Keys]::L) {
        if ($btnClearLog -and $btnClearLog.Enabled) { $btnClearLog.PerformClick() }
        $e.SuppressKeyPress = $true
    } elseif ($e.KeyCode -eq [System.Windows.Forms.Keys]::S) {
        if ($btnSaveLog -and $btnSaveLog.Enabled) { $btnSaveLog.PerformClick() }
        $e.SuppressKeyPress = $true
    } elseif ($e.KeyCode -eq [System.Windows.Forms.Keys]::F -and $script:ActiveNavIdx -eq 1) {
        if ($ctrlGPU.ContainsKey("VmSearch") -and $ctrlGPU["VmSearch"]) {
            $ctrlGPU["VmSearch"].Focus()
            $ctrlGPU["VmSearch"].SelectAll()
        }
        $e.SuppressKeyPress = $true
    } elseif ($e.KeyCode -eq [System.Windows.Forms.Keys]::D1) {
        $tabControl.SelectedIndex = 0
        $e.SuppressKeyPress = $true
    } elseif ($e.KeyCode -eq [System.Windows.Forms.Keys]::D2) {
        $tabControl.SelectedIndex = 1
        $e.SuppressKeyPress = $true
    } elseif ($e.KeyCode -eq [System.Windows.Forms.Keys]::D3) {
        $tabControl.SelectedIndex = 2
        $e.SuppressKeyPress = $true
    }
})

# ============================================================
#  BRUSH STUBS  (kept so FormClosing disposal code works)
# ============================================================
$script:TabBrushSelected = New-Object System.Drawing.SolidBrush($theme.SidebarActive)
$script:TabBrushNormal   = New-Object System.Drawing.SolidBrush($theme.Sidebar)

#endregion
#region ==================== EVENT HANDLERS ====================

# ---- Browse VM Location ----
$btnBrowseVM.Add_Click({
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    try {
        $dlg.Description = "Select the folder where VMs will be stored"
        if ($dlg.ShowDialog() -eq 'OK') { $ctrlCreate["VMLocation"].Text = $dlg.SelectedPath }
    } finally {
        $dlg.Dispose()
    }
})

# ---- Browse ISO + Detect Editions + Detect Version ----
$btnBrowseISO.Add_Click({
    # Prevent ISO browse/dismount while VM creation is in progress
    if ($script:IsCreating) { return }

    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Filter = "ISO Files|*.iso"
    $dlg.Title  = "Select a Windows Installation ISO"
    # Set initial directory to previous ISO path or user's Downloads folder
    try {
        $prevIso = $ctrlCreate["ISOPath"].Text
        if (-not [string]::IsNullOrWhiteSpace($prevIso)) {
            $prevDir = Split-Path $prevIso -Parent -ErrorAction SilentlyContinue
            if ($prevDir -and (Test-Path $prevDir)) { $dlg.InitialDirectory = $prevDir }
        }
        if (-not $dlg.InitialDirectory) {
            $downloadsDir = [Environment]::GetFolderPath('UserProfile') + '\Downloads'
            if (Test-Path $downloadsDir) { $dlg.InitialDirectory = $downloadsDir }
        }
    } catch {
        Write-Log "Could not determine initial ISO browse folder: $($_.Exception.Message)" "WARN"
    }
    try {
    if ($dlg.ShowDialog() -eq 'OK') {
        $ctrlCreate["ISOPath"].Text = $dlg.FileName

        # Dismount previous ISO if any
        if ($script:MountedISO) {
            try {
                [void](Dismount-ImageRetry -ImagePath $script:MountedISO.ImagePath -MaxRetries 2)
            } catch {
                Write-Log "Previous ISO dismount warning: $($_.Exception.Message)" "WARN"
            } finally {
                $script:MountedISO = $null
            }
        }

        try {
            Write-Log "Mounting ISO: $($dlg.FileName)"
            $script:MountedISO = Mount-DiskImage -ImagePath $dlg.FileName -PassThru -ErrorAction Stop
            Register-TrackedMountedImage -ImagePath $dlg.FileName

            # Poll for a drive letter (up to ~10 s) instead of a fixed sleep
            $isoVolume = $null
            for ($pollAttempt = 0; $pollAttempt -lt 10; $pollAttempt++) {
                Start-Sleep -Milliseconds 1000
                $isoVolume = $script:MountedISO | Get-Volume | Where-Object { $_.DriveLetter } | Select-Object -First 1
                if ($isoVolume -and $isoVolume.DriveLetter) { break }
            }
            if (-not $isoVolume -or -not $isoVolume.DriveLetter) {
                Write-Log "ISO mounted but no drive letter was assigned after 10 s. Try dismounting and re-mounting." "ERROR"
                return
            }
            $isoDrive = $isoVolume.DriveLetter + ":"

            # Find WIM or ESD
            $script:WimFile = Join-Path "$isoDrive\sources" "install.wim"
            if (-not (Test-Path $script:WimFile)) { $script:WimFile = Join-Path "$isoDrive\sources" "install.esd" }
            if (-not (Test-Path $script:WimFile)) {
                Write-Log "Cannot find install.wim or install.esd in ISO" "ERROR"
                $script:WimFile = $null
                $script:EditionMap = @{}
                $ctrlCreate["Edition"].Items.Clear()
                return
            }

            # Parse editions (this can take several seconds for large WIM/ESD files)
            Update-CreateProgress -Percent 0 -Status "Scanning ISO editions (DISM)..."
            $wimInfo = dism /Get-WimInfo /WimFile:"$($script:WimFile)" /English 2>&1
            $dismEditionExitCode = $LASTEXITCODE
            if ($dismEditionExitCode -ne 0) {
                Write-Log "DISM failed to read WIM info (exit code $dismEditionExitCode). The ISO image may be corrupt or unsupported." "ERROR"
            }
            $editions = @()
            $script:EditionMap = @{}
            $currentIndex = $null
            foreach ($line in $wimInfo) {
                $lineStr = $line.ToString().Trim()
                if ($lineStr -match "^Index\s*:\s*(\d+)$") { $currentIndex = [int]$matches[1] }
                elseif ($lineStr -match "^Name\s*:\s*(.+)$" -and $currentIndex) {
                    $name = $matches[1].Trim()
                    $editions += $name
                    $script:EditionMap[$name] = $currentIndex
                    $currentIndex = $null
                }
            }

            $ctrlCreate["Edition"].Items.Clear()
            $ctrlCreate["Edition"].Items.AddRange($editions)
            if ($editions.Count -gt 0) { $ctrlCreate["Edition"].SelectedIndex = 0 }
            if ($editions.Count -eq 0 -and $dismEditionExitCode -eq 0) {
                Write-Log "No Windows editions found in the WIM/ESD file. The ISO may not contain a standard install image." "WARN"
            }
            Write-Log "Found $($editions.Count) edition(s): $($editions -join ', ')" "OK"

            # Detect Windows version from first edition
            if ($editions.Count -gt 0) {
                $firstIndex = $script:EditionMap[$editions[0]]
                $versionInfo = Get-WimVersionInfo -WimFile $script:WimFile -Index $firstIndex
                $script:DetectedWinVersion = $versionInfo.WinVersion
                $script:DetectedBuild      = $versionInfo.Build
                $script:DetectedGuestArch  = if ($versionInfo.Architecture) { $versionInfo.Architecture } else { $script:HostArch }
                Write-Log "Detected guest architecture: $($script:DetectedGuestArch)" "INFO"

                $detectedProfile = Resolve-GuestWindowsProfile -DetectedWinVersion $script:DetectedWinVersion -DetectedBuild $script:DetectedBuild
                Set-DetectedGuestDefaults -Controls $ctrlCreate -Profile $detectedProfile -EmitLog

                # Auto-suggest VM name from ISO filename
                if ([string]::IsNullOrWhiteSpace($ctrlCreate["VMName"].Text)) {
                    $suggestedName = [IO.Path]::GetFileNameWithoutExtension($dlg.FileName) -replace '[^a-zA-Z0-9-]', ''
                    # Strip leading non-letter chars so the name passes '^[a-zA-Z]...' validation
                    $suggestedName = $suggestedName -replace '^[^a-zA-Z]+', ''
                    if ($suggestedName.Length -gt 15) { $suggestedName = $suggestedName.Substring(0, 15).TrimEnd('-') }
                    # Avoid suggesting a name that conflicts with an existing VM
                    if ($suggestedName -and (Get-VM -Name $suggestedName -ErrorAction SilentlyContinue)) {
                        $suffix = 2
                        $baseName = if ($suggestedName.Length -gt 13) { $suggestedName.Substring(0, 13) } else { $suggestedName }
                        while (Get-VM -Name "${baseName}-${suffix}" -ErrorAction SilentlyContinue) { $suffix++ }
                        $suggestedName = "${baseName}-${suffix}"
                    }
                    $ctrlCreate["VMName"].Text = $suggestedName
                }
            }

        } catch {
            Write-Log "Failed to read ISO: $_" "ERROR"
            # Clear stale edition/WIM data so validation doesn't show old values
            $script:WimFile = $null
            $script:EditionMap = @{}
            $ctrlCreate["Edition"].Items.Clear()
        }
    }
    } finally {
        $dlg.Dispose()
    }
})

# NOTE: $btnBrowseGolden.Add_Click is registered inline in the GUI construction section.

function Update-CreateModeUi {
    $isGoldenMode = $false
    if ($ctrlCreate.ContainsKey("GoldenImage") -and $ctrlCreate["GoldenImage"]) {
        $isGoldenMode = [bool]$ctrlCreate["GoldenImage"].Checked
    }

    if ($ctrlCreate.ContainsKey("ISOPath") -and $ctrlCreate["ISOPath"]) {
        $ctrlCreate["ISOPath"].Enabled = -not $isGoldenMode
    }
    if ($ctrlCreate.ContainsKey("Edition") -and $ctrlCreate["Edition"]) {
        $ctrlCreate["Edition"].Enabled = -not $isGoldenMode
    }
    if ($ctrlCreate.ContainsKey("GoldenParentVHD") -and $ctrlCreate["GoldenParentVHD"]) {
        $ctrlCreate["GoldenParentVHD"].Enabled = $isGoldenMode
    }
    if ($btnBrowseISO) { $btnBrowseISO.Enabled = -not $isGoldenMode }
    if ($btnBrowseGolden) { $btnBrowseGolden.Enabled = $isGoldenMode }

    # Resolution is irrelevant in golden mode (QRes.exe is not injected)
    if ($ctrlCreate.ContainsKey('Resolution') -and $ctrlCreate['Resolution']) {
        $ctrlCreate['Resolution'].Enabled = -not $isGoldenMode
    }

    # Golden mode skips unattend, so disable user-credential and post-install controls
    $goldenDisableKeys = @('Username','Password','EnableAutoLogon','Parsec','VBCable','USBMMIDD','RDP','Share','PauseUpdate','FullUpdate','StrictLegacyMode')
    foreach ($key in $goldenDisableKeys) {
        if ($ctrlCreate.ContainsKey($key) -and $ctrlCreate[$key]) {
            $ctrlCreate[$key].Enabled = -not $isGoldenMode
        }
    }

    if ($ctrlCreate.ContainsKey("ModeHint") -and $ctrlCreate["ModeHint"]) {
        if ($isGoldenMode) {
            $ctrlCreate["ModeHint"].Text = "Mode: Golden Image - Uses parent VHDX differencing disk. ISO and edition are ignored."
        } else {
            $ctrlCreate["ModeHint"].Text = "Mode: ISO Deploy - Uses ISO, selected edition, and unattended setup."
        }
    }
}

function Update-RoutingHint {
    if (-not $ctrlCreate.ContainsKey("RoutingHint") -or -not $ctrlCreate["RoutingHint"]) { return }

    $hasSwitch = ($ctrlCreate.ContainsKey("Switch") -and $ctrlCreate["Switch"] -and $ctrlCreate["Switch"].SelectedItem)
    $autoSwitch = ($ctrlCreate.ContainsKey("AutoCreateSwitch") -and $ctrlCreate["AutoCreateSwitch"] -and $ctrlCreate["AutoCreateSwitch"].Checked)
    $nestedVirt = ($ctrlCreate.ContainsKey("NestedVirt") -and $ctrlCreate["NestedVirt"] -and $ctrlCreate["NestedVirt"].Checked)
    $nestedNet  = ($ctrlCreate.ContainsKey("NestedNetFollowup") -and $ctrlCreate["NestedNetFollowup"] -and $ctrlCreate["NestedNetFollowup"].Checked)

    $networkText = if ($hasSwitch) {
        "Network: Selected switch is set."
    } elseif ($autoSwitch) {
        "Network: Auto-create NAT fallback is enabled."
    } else {
        "Network: No switch selected and auto-create is off. VM creation may fail."
    }

    $nestedText = if ($nestedNet -and $nestedVirt) {
        "Nested routing: Nested virtualization + MAC spoofing follow-up are enabled."
    } elseif ($nestedVirt) {
        "Nested routing: Nested virtualization is enabled."
    } else {
        "Nested routing: Disabled."
    }

    $ctrlCreate["RoutingHint"].Text = "$networkText $nestedText"
    if ($hasSwitch -or $autoSwitch) {
        $ctrlCreate["RoutingHint"].ForeColor = $theme.Muted
    } else {
        $ctrlCreate["RoutingHint"].ForeColor = [System.Drawing.Color]::Gold
    }
}

function Update-CreateValidationHint {
    if (-not $ctrlCreate.ContainsKey("ValidationHint") -or -not $ctrlCreate["ValidationHint"]) { return }

    $vmName = if ($ctrlCreate.ContainsKey("VMName") -and $ctrlCreate["VMName"]) { [string]$ctrlCreate["VMName"].Text } else { "" }
    $isoPath = if ($ctrlCreate.ContainsKey("ISOPath") -and $ctrlCreate["ISOPath"]) { [string]$ctrlCreate["ISOPath"].Text } else { "" }
    $userName = if ($ctrlCreate.ContainsKey("Username") -and $ctrlCreate["Username"]) { [string]$ctrlCreate["Username"].Text } else { "" }
    $passwordText = if ($ctrlCreate.ContainsKey("Password") -and $ctrlCreate["Password"]) { [string]$ctrlCreate["Password"].Text } else { "" }
    $switchSelected = ($ctrlCreate.ContainsKey("Switch") -and $ctrlCreate["Switch"] -and $ctrlCreate["Switch"].SelectedItem)
    $autoSwitch = ($ctrlCreate.ContainsKey("AutoCreateSwitch") -and $ctrlCreate["AutoCreateSwitch"] -and $ctrlCreate["AutoCreateSwitch"].Checked)
    $useGolden = ($ctrlCreate.ContainsKey("GoldenImage") -and $ctrlCreate["GoldenImage"] -and $ctrlCreate["GoldenImage"].Checked)
    $goldenPath = if ($ctrlCreate.ContainsKey("GoldenParentVHD") -and $ctrlCreate["GoldenParentVHD"]) { [string]$ctrlCreate["GoldenParentVHD"].Text } else { "" }

    $nameOk = (-not [string]::IsNullOrWhiteSpace($vmName)) -and ($vmName -match '^[a-zA-Z][a-zA-Z0-9-]*$') -and ($vmName -notmatch '-$') -and ($vmName.Length -le 64)
    $sourceOk = if ($useGolden) {
        -not [string]::IsNullOrWhiteSpace($goldenPath) -and (Test-PathCached $goldenPath)
    } else {
        $isoPathOk = (-not [string]::IsNullOrWhiteSpace($isoPath)) -and (Test-PathCached $isoPath)
        $isoMountedMatches = $false
        if ($script:MountedISO -and $script:MountedISO.ImagePath) {
            $mountedIsoPath = [string]$script:MountedISO.ImagePath
            try {
                $resolvedMounted = [System.IO.Path]::GetFullPath($mountedIsoPath.Trim()).ToLowerInvariant()
                $resolvedSelected = [System.IO.Path]::GetFullPath($isoPath.Trim()).ToLowerInvariant()
                $isoMountedMatches = ($resolvedMounted -eq $resolvedSelected)
            } catch {
                $isoMountedMatches = ($mountedIsoPath.Trim().ToLowerInvariant() -eq $isoPath.Trim().ToLowerInvariant())
            }
        }
        $editionSelected = ($ctrlCreate.ContainsKey("Edition") -and $ctrlCreate["Edition"] -and $ctrlCreate["Edition"].SelectedItem)
        $wimReady = (-not [string]::IsNullOrWhiteSpace($script:WimFile)) -and (Test-PathCached $script:WimFile) -and $editionSelected
        ($isoPathOk -and $isoMountedMatches -and $wimReady)
    }
    $networkOk = ($switchSelected -or $autoSwitch)
    $userOk = (-not [string]::IsNullOrWhiteSpace($userName)) -and ($userName -match '^[a-zA-Z0-9]+$') -and ($userName -ine $vmName)
    $passwordOk = (-not [string]::IsNullOrWhiteSpace($passwordText))
    $passwordWeak = $false
    if ($passwordOk) {
        $hasMinLength = ($passwordText.Length -ge 8)
        $hasUpper = ($passwordText -cmatch '[A-Z]')
        $hasLower = ($passwordText -cmatch '[a-z]')
        $hasDigit = ($passwordText -match '\d')
        $hasSpecial = ($passwordText -match '[^a-zA-Z0-9]')
        $complexityScore = @($hasUpper, $hasLower, $hasDigit, $hasSpecial) | Where-Object { $_ } | Measure-Object | Select-Object -ExpandProperty Count
        $passwordWeak = (-not $hasMinLength -or $complexityScore -lt 3)
    }

    # Golden image mode uses the embedded user profile - skip user/password checks
    if ($useGolden) {
        $userOk = $true
        $passwordOk = $true
        $passwordWeak = $false
    }

    $tick = [char]0x2713
    $cross = [char]0x2717
    $isReady = ($nameOk -and $sourceOk -and $networkOk -and $userOk -and $passwordOk)
    $pwLabel = if ($passwordWeak) { "Password(weak)" } else { "Password" }
    $ctrlCreate["ValidationHint"].Text = "Checks: Name $(if($nameOk){$tick}else{$cross}) | Source $(if($sourceOk){$tick}else{$cross}) | Network $(if($networkOk){$tick}else{$cross}) | User $(if($userOk){$tick}else{$cross}) | $pwLabel $(if($passwordOk){$tick}else{$cross})"

    if ($isReady) {
        $ctrlCreate["ValidationHint"].ForeColor = $theme.Success
    } elseif ($nameOk -or $sourceOk -or $networkOk -or $userOk -or $passwordOk) {
        $ctrlCreate["ValidationHint"].ForeColor = [System.Drawing.Color]::Gold
    } else {
        $ctrlCreate["ValidationHint"].ForeColor = $theme.Muted
    }

    if ($ctrlCreate.ContainsKey("CreateStatus") -and $ctrlCreate["CreateStatus"] -and -not $script:IsCreating) {
        if ($isReady) {
            $ctrlCreate["CreateStatus"].Text = "Ready to create VM"
            $ctrlCreate["CreateStatus"].ForeColor = $theme.Success
        } else {
            $missing = @()
            if (-not $nameOk) { $missing += "VM name" }
            if (-not $sourceOk) { $missing += $(if ($useGolden) { "parent VHD" } else { "ISO/source" }) }
            if (-not $networkOk) { $missing += "network" }
            if (-not $userOk) { $missing += "user" }
            if (-not $passwordOk) { $missing += "password" }
            $ctrlCreate["CreateStatus"].Text = "Fix required: $($missing -join ', ')"
            $ctrlCreate["CreateStatus"].ForeColor = [System.Drawing.Color]::Gold
        }
    }

    # Visual feedback on the VM Name textbox itself
    if ($ctrlCreate.ContainsKey("VMName") -and $ctrlCreate["VMName"]) {
        $vmNameBox = $ctrlCreate["VMName"]
        if ([string]::IsNullOrWhiteSpace($vmName)) {
            $vmNameBox.ForeColor = $theme.Text  # neutral when empty
        } elseif ($nameOk) {
            $vmNameBox.ForeColor = $theme.Text
        } else {
            $vmNameBox.ForeColor = [System.Drawing.Color]::FromArgb(255, 100, 100)  # red-ish for invalid
        }
    }

    if ($btnCreateVM -and -not $script:IsCreating) {
        $btnCreateVM.Enabled = $isReady
    }

    return $isReady
}

function Update-GpuActionState {
    if (-not $ctrlGPU -or -not $btnUpdateGPU) { return }

    $selectedNames = @()
    if ($script:GpuSelectedVMs) {
        $selectedNames += @(
            $script:GpuSelectedVMs.GetEnumerator() |
                Where-Object { $_.Value } |
                ForEach-Object { $_.Key }
        )
    }
    $selectedCount = @($selectedNames | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Sort-Object -Unique).Count
    $hasGpuChoice = ($ctrlGPU.ContainsKey("GpuSelector") -and $ctrlGPU["GpuSelector"] -and $ctrlGPU["GpuSelector"].SelectedIndex -ge 0)
    $canRun = ($selectedCount -gt 0 -and $hasGpuChoice)

    if ($ctrlGPU.ContainsKey("SelectionHint") -and $ctrlGPU["SelectionHint"]) {
        if ($selectedCount -eq 0) {
            $ctrlGPU["SelectionHint"].Text = "Select at least one VM to enable GPU update."
            $ctrlGPU["SelectionHint"].ForeColor = $theme.Muted
        } elseif (-not $hasGpuChoice) {
            $ctrlGPU["SelectionHint"].Text = "Select a GPU provider to continue."
            $ctrlGPU["SelectionHint"].ForeColor = [System.Drawing.Color]::Gold
        } else {
            $ctrlGPU["SelectionHint"].Text = "Ready: $selectedCount VM(s) selected for GPU update."
            $ctrlGPU["SelectionHint"].ForeColor = $theme.Success
        }
    }

    if (-not $script:IsUpdatingGPU) {
        $btnUpdateGPU.Enabled = $canRun
    }
}

function Update-DynamicMemoryUi {
    if (-not $ctrlCreate.ContainsKey("DynamicMem") -or -not $ctrlCreate["DynamicMem"]) { return }

    $enabled = [bool]$ctrlCreate["DynamicMem"].Checked
    foreach ($key in @("DynamicMemMin", "DynamicMemMax")) {
        if ($ctrlCreate.ContainsKey($key) -and $ctrlCreate[$key]) {
            $ctrlCreate[$key].Enabled = $enabled
        }
    }
}

$ctrlCreate["GoldenImage"].Add_CheckedChanged({
    Update-CreateModeUi
    Update-CreateValidationHint
})

Update-CreateModeUi
Update-RoutingHint
Update-CreateValidationHint
Update-DynamicMemoryUi

# Keep NAT/routing-related checkboxes in a valid state
$ctrlCreate["AutoCreateSwitch"].Add_CheckedChanged({
    if (-not $ctrlCreate["AutoCreateSwitch"].Checked) {
        $selectedSwitch = $ctrlCreate["Switch"].SelectedItem
        if (-not $selectedSwitch -or [string]::IsNullOrWhiteSpace([string]$selectedSwitch)) {
            Write-Log "Auto-create NAT switch was disabled with no selected switch. Re-enabling to avoid network provisioning failure." "WARN"
            $ctrlCreate["AutoCreateSwitch"].Checked = $true
        }
    }
    Update-RoutingHint
    Update-CreateValidationHint
})

$ctrlCreate["DynamicMem"].Add_CheckedChanged({
    # Hyper-V does not support dynamic memory with nested virtualization
    if ($ctrlCreate["DynamicMem"].Checked -and $ctrlCreate["NestedVirt"].Checked) {
        $ctrlCreate["DynamicMem"].Checked = $false
        Write-Log "Dynamic Memory cannot be enabled while Nested Virtualization is active (Hyper-V limitation)." "WARN"
    }
    Update-DynamicMemoryUi
})

$ctrlCreate["DynamicMemMin"].Add_ValueChanged({
    if ($script:SuppressMemEvents) { return }
    $script:SuppressMemEvents = $true
    try {
        if ([int]$ctrlCreate["DynamicMemMin"].Value -gt [int]$ctrlCreate["Memory"].Value) {
            $ctrlCreate["DynamicMemMin"].Value = $ctrlCreate["Memory"].Value
        }
        if ([int]$ctrlCreate["DynamicMemMin"].Value -gt [int]$ctrlCreate["DynamicMemMax"].Value) {
            $ctrlCreate["DynamicMemMax"].Value = $ctrlCreate["DynamicMemMin"].Value
        }
    } finally { $script:SuppressMemEvents = $false }
})

$ctrlCreate["DynamicMemMax"].Add_ValueChanged({
    if ($script:SuppressMemEvents) { return }
    $script:SuppressMemEvents = $true
    try {
        if ([int]$ctrlCreate["DynamicMemMax"].Value -lt [int]$ctrlCreate["Memory"].Value) {
            $ctrlCreate["DynamicMemMax"].Value = $ctrlCreate["Memory"].Value
        }
        if ([int]$ctrlCreate["DynamicMemMax"].Value -lt [int]$ctrlCreate["DynamicMemMin"].Value) {
            $ctrlCreate["DynamicMemMin"].Value = $ctrlCreate["DynamicMemMax"].Value
        }
    } finally { $script:SuppressMemEvents = $false }
})

$ctrlCreate["NestedVirt"].Add_CheckedChanged({
    if (-not $ctrlCreate["NestedVirt"].Checked -and $ctrlCreate["NestedNetFollowup"].Checked) {
        $ctrlCreate["NestedNetFollowup"].Checked = $false
        Write-Log "Nested Net follow-up was disabled because Nested Virtualization is off." "INFO"
    }
    # Hyper-V does not support dynamic memory with nested virtualization
    if ($ctrlCreate["NestedVirt"].Checked -and $ctrlCreate["DynamicMem"].Checked) {
        $ctrlCreate["DynamicMem"].Checked = $false
        Write-Log "Dynamic Memory was disabled because Hyper-V does not support it with Nested Virtualization." "WARN"
    }
    Update-RoutingHint
})

$ctrlCreate["NestedNetFollowup"].Add_CheckedChanged({
    if ($ctrlCreate["NestedNetFollowup"].Checked -and -not $ctrlCreate["NestedVirt"].Checked) {
        $ctrlCreate["NestedVirt"].Checked = $true
        Write-Log "Nested Virtualization was automatically enabled because Nested Net follow-up requires it." "INFO"
    }
    Update-RoutingHint
})

# Mutual exclusion: PauseUpdate and FullUpdate conflict
$ctrlCreate["PauseUpdate"].Add_CheckedChanged({
    if ($ctrlCreate["PauseUpdate"].Checked -and $ctrlCreate["FullUpdate"].Checked) {
        $ctrlCreate["FullUpdate"].Checked = $false
        Write-Log "Full Windows Updates was disabled because Pause Updates is enabled (they conflict)." "INFO"
    }
})

$ctrlCreate["FullUpdate"].Add_CheckedChanged({
    if ($ctrlCreate["FullUpdate"].Checked -and $ctrlCreate["PauseUpdate"].Checked) {
        $ctrlCreate["PauseUpdate"].Checked = $false
        Write-Log "Pause Updates was disabled because Full Windows Updates is enabled (they conflict)." "INFO"
    }
})

# ---- Update OS info when edition changes ----
$ctrlCreate["Edition"].Add_SelectedIndexChanged({
    $selectedEdition = $ctrlCreate["Edition"].SelectedItem
    if ($selectedEdition -and $script:EditionMap.ContainsKey($selectedEdition)) {
        if (-not $script:WimFile -or -not (Test-Path $script:WimFile)) {
            Write-Log "WIM file no longer accessible (ISO may have been dismounted). Re-select the ISO." "WARN"
            return
        }
        $idx = $script:EditionMap[$selectedEdition]
        $versionInfo = Get-WimVersionInfo -WimFile $script:WimFile -Index $idx
        $script:DetectedWinVersion = $versionInfo.WinVersion
        $script:DetectedBuild      = $versionInfo.Build
        $script:DetectedGuestArch  = if ($versionInfo.Architecture) { $versionInfo.Architecture } else { $script:HostArch }
        Write-Log "Detected guest architecture: $($script:DetectedGuestArch)" "INFO"
        $detectedProfile = Resolve-GuestWindowsProfile -DetectedWinVersion $script:DetectedWinVersion -DetectedBuild $script:DetectedBuild
        Set-DetectedGuestDefaults -Controls $ctrlCreate -Profile $detectedProfile
    }
    Update-CreateValidationHint
})

$ctrlCreate["Memory"].Add_ValueChanged({
    if ($script:SuppressMemEvents) { return }
    if (-not $ctrlCreate.ContainsKey("DynamicMemMin") -or -not $ctrlCreate.ContainsKey("DynamicMemMax")) { return }
    $script:SuppressMemEvents = $true
    try {
        $startup = [int]$ctrlCreate["Memory"].Value
        if ([int]$ctrlCreate["DynamicMemMin"].Value -gt $startup) {
            $ctrlCreate["DynamicMemMin"].Value = $startup
        }
        if ([int]$ctrlCreate["DynamicMemMax"].Value -lt $startup) {
            $ctrlCreate["DynamicMemMax"].Value = $startup
        }
    } finally { $script:SuppressMemEvents = $false }
})

# Debounce timer: coalesce rapid TextChanged events into a single validation pass (300 ms)
$script:ValidationTimer = New-Object System.Windows.Forms.Timer
$script:ValidationTimer.Interval = 300
$script:ValidationTimer.Add_Tick({
    $script:ValidationTimer.Stop()
    Update-CreateValidationHint
})

$ctrlCreate["VMName"].Add_TextChanged({ $script:ValidationTimer.Stop(); $script:ValidationTimer.Start() })
$ctrlCreate["ISOPath"].Add_TextChanged({ $script:ValidationTimer.Stop(); $script:ValidationTimer.Start() })
$ctrlCreate["GoldenParentVHD"].Add_TextChanged({ $script:ValidationTimer.Stop(); $script:ValidationTimer.Start() })
$ctrlCreate["Username"].Add_TextChanged({ $script:ValidationTimer.Stop(); $script:ValidationTimer.Start() })
$ctrlCreate["Password"].Add_TextChanged({ $script:ValidationTimer.Stop(); $script:ValidationTimer.Start() })
$ctrlCreate["Switch"].Add_SelectedIndexChanged({ Update-RoutingHint; $script:ValidationTimer.Stop(); $script:ValidationTimer.Start() })

Update-GpuActionState

# ================================================================
#  CREATE VM - Main Logic
# ================================================================
$btnCreateVM.Add_Click({
    # Re-entrancy guard: prevent double-click during VM creation
    if ($script:IsCreating) { return }

    # Confirmation dialog before starting VM creation
    $confirmResult = [System.Windows.Forms.MessageBox]::Show(
        "Are you sure you want to create this VM?`n`nThis will allocate disk space and configure Hyper-V resources.",
        "Confirm VM Creation", [System.Windows.Forms.MessageBoxButtons]::OKCancel, [System.Windows.Forms.MessageBoxIcon]::Question)
    if ($confirmResult -ne [System.Windows.Forms.DialogResult]::OK) { return }

    $script:IsCreating = $true
    $validationError = $null
    $isValidationError = $false

    $VMName = ""
    $VMLoc = ""
    $VHDPath = ""
    $preflightLines = @()
    $rollbackNeeded = $false
    $autoPlayGuard = $null
    $vhdMountedForDeploy = $false
    $vmFolderCreatedByScript = $false
    $tempExe = $null
    $cmdFile = $null
    $UnattendXMLPath = $null

    try {
        # Wrap the entire validation and creation logic in a nested try-catch
        # so we can handle early exits without skipping cleanup
        try {
        Update-CreateProgress -Percent 2 -Status "Validating VM settings..."

        # ---- Gather inputs ----
        $VMName          = $ctrlCreate["VMName"].Text.Trim()
        $VMLocBase       = $ctrlCreate["VMLocation"].Text.Trim()
        $ISOPath         = $ctrlCreate["ISOPath"].Text.Trim()
        $Username        = $ctrlCreate["Username"].Text.Trim()
        $PasswordText    = $ctrlCreate["Password"].Text
        $SelectedResolution = [string]$ctrlCreate["Resolution"].SelectedItem
        if ([string]::IsNullOrWhiteSpace($SelectedResolution)) { $SelectedResolution = "1920x1080" }
        $Password        = $null
        $vCPU            = [int]$ctrlCreate["vCPU"].Value
        $MemGB           = [int]$ctrlCreate["Memory"].Value
        $DiskGB          = [int]$ctrlCreate["DiskSize"].Value
        $VMSwitch        = $ctrlCreate["Switch"].SelectedItem
        $CheckpointMode  = [string]$ctrlCreate["CheckpointMode"].SelectedItem
        if ([string]::IsNullOrWhiteSpace($CheckpointMode)) { $CheckpointMode = "Disabled" }
        $DynamicMemMinGB = [int]$ctrlCreate["DynamicMemMin"].Value
        $DynamicMemMaxGB = [int]$ctrlCreate["DynamicMemMax"].Value
        $SelectedEdition = $ctrlCreate["Edition"].SelectedItem
        $SelectedIndex   = $null
        if ($SelectedEdition -and $script:EditionMap.ContainsKey($SelectedEdition)) {
            $SelectedIndex = [int]$script:EditionMap[$SelectedEdition]
        }
        $EnableSecureBoot     = $ctrlCreate["SecureBoot"].Checked
        $EnableTPM            = $ctrlCreate["TPM"].Checked
        $FixedVHD             = $ctrlCreate["VHDType"].Checked
        $EnableDynamicMem     = $ctrlCreate["DynamicMem"].Checked
        $EnableEnhancedSession = $ctrlCreate["EnhancedSession"].Checked
        $StartVM              = $ctrlCreate["StartVM"].Checked
        $StrictLegacyMode     = $ctrlCreate["StrictLegacyMode"].Checked
        $AutoCreateSwitch     = $ctrlCreate["AutoCreateSwitch"].Checked
        $EnableMetering       = $ctrlCreate["EnableMetering"].Checked
        $EnableAutoLogon      = if ($ctrlCreate.ContainsKey("EnableAutoLogon") -and $ctrlCreate["EnableAutoLogon"]) { [bool]$ctrlCreate["EnableAutoLogon"].Checked } else { $true }
        $EnableNestedVirt     = $ctrlCreate["NestedVirt"].Checked
        $EnableNestedNetFollowup = $ctrlCreate["NestedNetFollowup"].Checked
        $ResetBootOrder       = $ctrlCreate["ResetBootOrder"].Checked
        $UseGoldenImage       = $ctrlCreate["GoldenImage"].Checked
        $GoldenParentVHD      = $ctrlCreate["GoldenParentVHD"].Text.Trim()
        $EnableVmNotes        = $true
        $guestProfile = if ($UseGoldenImage) {
            Resolve-GuestWindowsProfile -DetectedWinVersion "Unknown" -DetectedBuild 0
        } else {
            Resolve-GuestWindowsProfile -DetectedWinVersion $script:DetectedWinVersion -DetectedBuild $script:DetectedBuild
        }
        $IsWin11 = $guestProfile.IsWindows11
        $effectiveDetectedBuild = if ($UseGoldenImage) { [int]$guestProfile.Build } else { [int]$script:DetectedBuild }
        $effectiveDetectedOsName = if ($UseGoldenImage) { $guestProfile.Name } else { $script:DetectedWinVersion }

        if ([string]::IsNullOrWhiteSpace($VMLocBase)) { throw "VM Location is required." }
        try {
            $VMLocBase = [System.IO.Path]::GetFullPath($VMLocBase)
        } catch {
            throw "VM Location path is invalid: $($PSItem.Exception.Message)"
        }

        if (-not $UseGoldenImage -and -not [string]::IsNullOrWhiteSpace($ISOPath)) {
            try {
                $ISOPath = [System.IO.Path]::GetFullPath($ISOPath)
            } catch {
                Write-Log "ISO path normalization warning: $($_.Exception.Message)" "WARN"
            }
        }
        if ($UseGoldenImage -and -not [string]::IsNullOrWhiteSpace($GoldenParentVHD)) {
            try {
                $GoldenParentVHD = [System.IO.Path]::GetFullPath($GoldenParentVHD)
            } catch {
                Write-Log "Golden parent path normalization warning: $($_.Exception.Message)" "WARN"
            }
        }

        if ($StrictLegacyMode) {
            if ($IsWin11) {
                Write-Log "Strict Legacy Mode is ignored for Windows 11 compatibility requirements." "WARN"
            } else {
                $guestProfile.IsLegacyWindows10 = $true
                $guestProfile.PreferCompactApply = $false
                $guestProfile.SecureBootTemplateOrder = @('MicrosoftWindows', 'MicrosoftUEFICertificateAuthority')
                $guestProfile.CompatibilityNote = "Strict Legacy Mode enabled: forcing legacy-safe image apply and secure boot template order."
            }
        }

        if ($UseGoldenImage) {
            Write-Log "Golden image mode: guest OS build is not auto-detected from VHD. Secure Boot/TPM defaults remain user-controlled." "INFO"
        } else {
            Write-Log "Guest compatibility profile: $($guestProfile.Name) build $($guestProfile.Build)" "INFO"
            Write-Log "  $($guestProfile.CompatibilityNote)" "INFO"
        }

        if ($IsWin11 -and -not $EnableSecureBoot) {
            Write-Log "Windows 11 detected - Secure Boot is required, enabling automatically." "WARN"
            $EnableSecureBoot = $true
            $ctrlCreate["SecureBoot"].Checked = $true
        }
        if ($IsWin11 -and -not $EnableTPM) {
            Write-Log "Windows 11 detected - TPM is required, enabling automatically." "WARN"
            $EnableTPM = $true
            $ctrlCreate["TPM"].Checked = $true
        }

        # ---- Validate ----
        $isValidationError = $true   # flag: throws in this section are user-facing validation errors
        if ([string]::IsNullOrWhiteSpace($VMName)) { throw "VM Name is required!" }
        if ($VMName -notmatch '^[a-zA-Z][a-zA-Z0-9-]*$') { 
            throw "VM Name must start with a letter and contain only letters, numbers, and hyphens."
        }
        if ($VMName -match '-$') { throw "VM Name cannot end with a hyphen." }
        if ($VMName.Length -gt 64) { throw "VM Name must be 64 characters or less." }
        $reservedDeviceNames = @('CON','PRN','AUX','NUL','COM1','COM2','COM3','COM4','COM5','COM6','COM7','COM8','COM9','LPT1','LPT2','LPT3','LPT4','LPT5','LPT6','LPT7','LPT8','LPT9')
        if ($VMName.ToUpperInvariant() -in $reservedDeviceNames) {
            throw "VM Name cannot be a Windows reserved device name ('$VMName')."
        }
        if ($VMName.Length -gt 15) { Write-Log "VM Name is longer than 15 characters. This is valid for Hyper-V, but may reduce compatibility with legacy NetBIOS-dependent workflows." "WARN" }
        if (Get-VM -Name $VMName -ErrorAction SilentlyContinue) { throw "A VM named '$VMName' already exists!" }

        if (-not $UseGoldenImage) {
            if ([string]::IsNullOrWhiteSpace($ISOPath) -or -not (Test-Path $ISOPath)) { throw "Valid ISO file is required." }
            if (-not $script:MountedISO -or [string]::IsNullOrWhiteSpace($script:MountedISO.ImagePath)) {
                throw "Selected ISO is not mounted. Please re-select the ISO before creating the VM."
            }
            $mountedIsoPath = [string]$script:MountedISO.ImagePath
            $isoMatch = $false
            try {
                $resolvedMounted = [System.IO.Path]::GetFullPath($mountedIsoPath.Trim()).ToLowerInvariant()
                $resolvedSelected = [System.IO.Path]::GetFullPath($ISOPath.Trim()).ToLowerInvariant()
                $isoMatch = ($resolvedMounted -eq $resolvedSelected)
            } catch {
                $isoMatch = ($mountedIsoPath.Trim().ToLowerInvariant() -eq $ISOPath.Trim().ToLowerInvariant())
            }
            if (-not $isoMatch) {
                throw "Selected ISO does not match the currently mounted image. Re-select the ISO to refresh source metadata."
            }
        } else {
            if ([string]::IsNullOrWhiteSpace($GoldenParentVHD) -or -not (Test-Path $GoldenParentVHD)) {
                throw "Golden Image mode is enabled. Please select a valid parent VHDX/VHD."
            }
            $goldenExt = [System.IO.Path]::GetExtension($GoldenParentVHD)
            if ($goldenExt -notin @('.vhd', '.vhdx')) {
                throw "Golden parent disk must be a .vhd or .vhdx file."
            }
        }

        if (-not (Test-Path $VMLocBase)) {
            Write-Log "VM location does not exist. Creating folder: $VMLocBase" "WARN"
            try {
                New-Item -Path $VMLocBase -ItemType Directory -Force -ErrorAction Stop | Out-Null
            } catch {
                throw "Failed to create VM location: $($PSItem.Exception.Message)"
            }
        }
        if (-not (Test-DirectoryWritable -Path $VMLocBase)) { throw "VM Location is not writable: $VMLocBase" }

        if ($UseGoldenImage) {
            # Golden mode creates only a differencing disk at this stage.
            $requiredDiskGB = 2
        } elseif ($FixedVHD) {
            # Fixed VHD allocates full disk size up front.
            $requiredDiskGB = $DiskGB + 8
        } else {
            # Dynamic VHD baseline plus selected post-install workload overhead.
            $dynamicRequiredDiskGB = 24
            $dynamicOverheadGB = 0
            if ($ctrlCreate["Parsec"].Checked)   { $dynamicOverheadGB += 2 }
            if ($ctrlCreate["VBCable"].Checked)  { $dynamicOverheadGB += 2 }
            if ($ctrlCreate["USBMMIDD"].Checked) { $dynamicOverheadGB += 2 }
            if ($ctrlCreate["Share"].Checked)    { $dynamicOverheadGB += 1 }
            if ($ctrlCreate["FullUpdate"].Checked) { $dynamicOverheadGB += 8 }

            # Keep estimate realistic for dynamic disks: no more than VHD max + modest temp overhead.
            $requiredDiskGB = [math]::Min(($DiskGB + 4), ($dynamicRequiredDiskGB + $dynamicOverheadGB))
        }
        $freeDiskGB = Get-PathAvailableSpaceGB -Path $VMLocBase
        if ($freeDiskGB -ge 0 -and $freeDiskGB -lt $requiredDiskGB) {
            throw "Insufficient disk space in VM location. Need about $requiredDiskGB GB free, found $freeDiskGB GB."
        }

        if (-not $UseGoldenImage) {
            if (-not $script:WimFile -or -not (Test-Path $script:WimFile)) { throw "No WIM/ESD file found. Please select an ISO first." }
        }
        if (-not $UseGoldenImage) {
            if ([string]::IsNullOrWhiteSpace($Username)) { throw "Username is required!" }
            if ($Username -notmatch '^[a-zA-Z0-9]+$') { throw "Username cannot contain special characters." }
            if ($Username -ieq $VMName) { throw "Username cannot be the same as VM Name (causes admin permission issues in the VM)." }
            $reservedUserNames = @('Administrator','Guest','DefaultAccount','WDAGUtilityAccount','SYSTEM','NetworkService','LocalService')
            if ($Username -in $reservedUserNames) {
                throw "Username cannot be a Windows built-in account name ('$Username')."
            }
            if ([string]::IsNullOrWhiteSpace($PasswordText)) { throw "Password is required for unattended setup." }
            # Convert password immediately to minimize plaintext lifetime
            $Password = Convert-PlainTextToSecureString -Text $PasswordText
            $PasswordText = $null   # Clear plaintext immediately
        }
        if ($EnableDynamicMem) {
            if ($EnableNestedVirt) {
                Write-Log "Dynamic Memory is incompatible with Nested Virtualization. Disabling dynamic memory." "WARN"
                $EnableDynamicMem = $false
                $ctrlCreate["DynamicMem"].Checked = $false
            } else {
                if ($DynamicMemMinGB -gt $DynamicMemMaxGB) {
                    throw "Dynamic Memory minimum ($DynamicMemMinGB GB) cannot be greater than maximum ($DynamicMemMaxGB GB)."
                }
                if ($DynamicMemMinGB -gt $MemGB) {
                    throw "Dynamic Memory minimum ($DynamicMemMinGB GB) cannot be greater than startup memory ($MemGB GB)."
                }
                if ($DynamicMemMaxGB -lt $MemGB) {
                    throw "Dynamic Memory maximum ($DynamicMemMaxGB GB) cannot be less than startup memory ($MemGB GB)."
                }
            }
        }

        if ($VMSwitch -and (Get-VMSwitch -Name $VMSwitch -ErrorAction SilentlyContinue)) {
            # selected switch is valid
        } else {
            if (-not $AutoCreateSwitch) {
                if (-not $VMSwitch) { throw "No Virtual Switch selected!" }
                else { throw "Selected Virtual Switch no longer exists: $VMSwitch" }
            }

            $createdSwitch = Set-ToolkitNatSwitch
            if (-not $createdSwitch) {
                throw "Could not auto-create a fallback virtual switch."
            }

            $VMSwitch = $createdSwitch
            $ctrlCreate["Switch"].Items.Clear()
            Get-VMSwitch | Select-Object -ExpandProperty Name | ForEach-Object { [void]$ctrlCreate["Switch"].Items.Add($_) }
            $ctrlCreate["Switch"].SelectedItem = $VMSwitch
        }

        if (-not $UseGoldenImage -and ($null -eq $SelectedIndex -or $SelectedIndex -lt 1)) { throw "No Windows Edition selected!" }
        $isValidationError = $false  # end of validation section

        $hostSupportedVersions = "Unavailable"
        try {
            $supported = Get-VMHostSupportedVersion -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Version
            if ($supported) { $hostSupportedVersions = ($supported -join ', ') }
        } catch {
            Write-Log "Preflight note: unable to query host supported VM versions: $($_.Exception.Message)" "WARN"
        }
        $sriovReady = $false
        try {
            $sriovAdapters = Get-NetAdapterSriov -ErrorAction SilentlyContinue
            if ($sriovAdapters) {
                $sriovReady = ($sriovAdapters | Where-Object { $_.SriovEnabled -eq $true -or $_.SriovSupport -match 'Supported|Ready' }).Count -gt 0
            }
        } catch {
            Write-Log "Preflight note: unable to query SR-IOV readiness: $($_.Exception.Message)" "WARN"
        }
        $freeHostMemGB = -1
        try {
            $osInfo = Get-CimInstance Win32_OperatingSystem -ErrorAction SilentlyContinue
            if ($osInfo) { $freeHostMemGB = [math]::Round((($osInfo.FreePhysicalMemory * 1KB) / 1GB), 2) }
        } catch {
            Write-Log "Preflight note: unable to query host free memory: $($_.Exception.Message)" "WARN"
        }
        $tpmCmdAvailable = [bool](Get-Command Enable-VMTPM -ErrorAction SilentlyContinue)

        $preflightLines = @(
            "ProfileTime=$((Get-Date).ToString('yyyy-MM-dd HH:mm:ss'))",
            "VMName=$VMName",
            "Mode=$(if ($UseGoldenImage) { 'GoldenImage' } else { 'ISODeploy' })",
            "StrictLegacyMode=$StrictLegacyMode",
            "GuestProfile=$($guestProfile.Name)",
            "GuestBuild=$($guestProfile.Build)",
            "GuestLegacyWin10=$($guestProfile.IsLegacyWindows10)",
            "PreferCompactApply=$($guestProfile.PreferCompactApply)",
            "HostBuild=$($script:HostBuild)",
            "HostSupportedVMVersions=$hostSupportedVersions",
            "SRIOVReady=$sriovReady",
            "FreeHostMemoryGB=$freeHostMemGB",
            "TargetFreeDiskGB=$freeDiskGB",
            "vTPMCommandAvailable=$tpmCmdAvailable",
            "SelectedSwitch=$VMSwitch"
        )
        Write-Log "Preflight Profile:" "INFO"
        foreach ($line in $preflightLines) { Write-Log "  $line" "INFO" }

        if ($script:CliWhatIf) {
            Write-Log "WhatIf mode: preflight completed. No VM changes will be made." "WARN"
            Update-CreateProgress -Percent 100 -Status "WhatIf complete: no changes were applied."
            return
        }

        # ---- Host/VM version match check ----
        if ($effectiveDetectedBuild -gt 0 -and $script:HostBuild -gt 0) {
            if ([math]::Abs($effectiveDetectedBuild - $script:HostBuild) -gt 5000) {
                Write-Log "WARNING: Large version gap between host (Build $($script:HostBuild)) and VM image (Build $effectiveDetectedBuild)." "WARN"
                Write-Log "Mismatched Windows versions can cause GPU-P driver issues or BSODs. Matching versions recommended." "WARN"
            }
        }

        # ---- Disable UI ----
        $tabControl.Enabled = $false
        $btnCreateVM.Enabled = $false
        Update-CreateProgress -Percent 6 -Status "Preparing VM creation workflow..."
        Write-Log "========================================" "INFO"
        Write-Log "Starting VM creation: $VMName" "INFO"
        Write-Log "OS: $effectiveDetectedOsName Build $effectiveDetectedBuild" "INFO"
        Write-Log "Secure Boot: $EnableSecureBoot | TPM: $EnableTPM" "INFO"
        Write-Log "========================================" "INFO"
        if (-not $UseGoldenImage) {
            Write-Log "  Security baseline: UAC and Windows Firewall remain enabled in unattend defaults." "INFO"
        }

        # ---- Disable AutoPlay ----
        $autoPlayGuard = Disable-AutoPlayGuarded

        # ---- Create VM directory ----
        Update-CreateProgress -Percent 10 -Status "Creating VM workspace..."
        $VMLoc = Join-Path $VMLocBase $VMName
        $VHDPath = Join-Path $VMLoc "$VMName.vhdx"  # Set early so rollback always knows the VHD path
        if (Test-Path $VMLoc) { Write-Log "Directory $VMLoc already exists; files may be overwritten" "WARN" }
        else {
            New-Item -Path $VMLoc -ItemType Directory -Force -ErrorAction Stop | Out-Null
            $vmFolderCreatedByScript = $true
        }
        try {
            $preflightPath = Join-Path $VMLoc "Preflight-Profile.txt"
            $preflightLines | Out-File -FilePath $preflightPath -Encoding UTF8 -Force
            Write-Log "Preflight profile exported: $preflightPath" "OK"
        } catch {
            Write-Log "Could not export preflight profile: $($_.Exception.Message)" "WARN"
        }
        $rollbackNeeded = $true

        $attachISOForRecovery = $false

        if (-not $UseGoldenImage) {
            # ---- Create QRes.exe ----
            # QRes v1.1 by Anders Kjersem - open-source display resolution changer
            # Source: https://sourceforge.net/projects/qres/ (embedded for offline VM provisioning)
            Update-CreateProgress -Percent 15 -Status "Preparing setup tools..."
            Write-Log "Creating QRes.exe..."
            $tempExe = Join-Path $VMLoc "QRes.exe"
        $base64 = @"
TVqQAAMAAAAEAAAA//8AALgAAAAAAAAAQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA0AAAAA4fug4AtAnNIbgBTM0hVGhpcyBwcm9ncmFtIGNhbm5vdCBiZSBydW4gaW4gRE9TIG1vZGUuDQ0KJAAAAAAAAACDEdcDx3C5UMdwuVDHcLlQRGy3UMZwuVAvb71QxXC5UMdwuFDZcLlQpW+qUM5wuVAvb7NQy3C5UFJpY2jHcLlQAAAAAAAAAAAAAAAAAAAAAFBFAABMAQEASP76PgAAAAAAAAAA4AAPAQsBBgAAAAAAABAAAAAAAABIGwAAABAAAAAQAAAAAEAAABAAAAACAAAEAAAAAAAAAAQAAAAAAAAAACAAAAACAAD2EAEAAwAAAAAAEAAAEAAAAAAQAAAQAAAAAAAAEAAAAAAAAAAAAAAAsBwAAHgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAQAACEAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAALmRhdGEAAACKDwAAABAAAAAQAAAAAgAAAAAAAAAAAAAAAAAAQAAAwAAAAAAAAAAAAAAAAAAAAABqHgAAeB4AAIweAAAAAAAAUB4AAAAAAADSHQAAxB0AALgdAACsHQAAAAAAAKoeAAD4HgAACB8AAHwfAABoHwAAtB4AAMoeAADSHgAA4B4AAOgeAABWHwAAFB8AACgfAAA4HwAASB8AAAAAAAAYHgAA8h0AAP4dAAAkHgAALB4AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAARXJyb3I6ICVzCgAAICAlcwklcwoAAAAACSAlcwoAAAAlcy4KAAAAACBAICVkIEh6AAAAAEFkYXB0ZXIgRGVmYXVsdABPcHRpbWFsAHVua25vd24AJWR4JWQsICVkIGJpdHMAACBAIAAKRXg6ICJRUmVzLmV4ZSAveDo2NDAgL2M6OCIgQ2hhbmdlcyByZXNvbHV0aW9uIHRvIDY0MCB4IDQ4MCBhbmQgdGhlIGNvbG9yIGRlcHRoIHRvIDI1NiBjb2xvcnMuCgAvSAAARGlzcGxheXMgbW9yZSBoZWxwLgAvPwAARGlzcGxheXMgdXNhZ2UgaW5mb3JtYXRpb24uAC9WAABEb2VzIE5PVCBkaXNwbGF5IHZlcnNpb24gaW5mb3JtYXRpb24uAAAAL0QAAERvZXMgTk9UIHNhdmUgZGlzcGxheSBzZXR0aW5ncyBpbiB0aGUgcmVnaXN0cnkuLgAAAAAvTAAATGlzdCBhbGwgZGlzcGxheSBtb2Rlcy4AL1MAAFNob3cgY3VycmVudCBkaXNwbGF5IHNldHRpbmdzLgAALTE9IE9wdGltYWwuAAAAADAgPSBBZGFwdGVyIERlZmF1bHQuAAAAAC9SAABSZWZyZXNoIHJhdGUuAAAAMzI9IFRydWUgY29sb3IuADI0PSBUcnVlIGNvbG9yLgAxNj0gSGlnaCBjb2xvci4AOCA9IDI1NiBjb2xvcnMuADQgPSAxNiBjb2xvcnMuAAAvQwAAQ29sb3IgZGVwdGguAAAAAC9ZAABIZWlnaHQgaW4gcGl4ZWxzLgAAAC9YAABXaWR0aCBpbiBwaXhlbHMuAAAAAFFSRVMgWy9YOltweF1dIFsvWTpbcHhdXSBbL0M6W2JpdHNdIFsvUjpbcnJdXSBbL1NdIFsvTF0gWy9EXSBbL1ZdIFsvP10gWy9IXQoKAAAAU2V0dGluZ3MgY291bGQgbm90IGJlIHNhdmVkLGdyYXBoaWNzIG1vZGUgd2lsbCBiZSBjaGFuZ2VkIGR5bmFtaWNhbGx5Li4uAAAAAFRoZSBjb21wdXRlciBtdXN0IGJlIHJlc3RhcnRlZCBpbiBvcmRlciBmb3IgdGhlIGdyYXBoaWNzIG1vZGUgdG8gd29yay4uLgAAAABUaGUgZ3JhcGhpY3MgbW9kZSBpcyBub3Qgc3VwcG9ydGVkIQBNb2RlIE9rLi4uCgBSZWZyZXNoUmF0ZQBEaXNwbGF5XFNldHRpbmdzAAAAAFFSZXMgdjEuMQpDb3B5cmlnaHQgKEMpIEFuZGVycyBLamVyc2VtLgoKAAAAAQAAAAAAAAD/////OBxAAEwcQAAAAAAA/3QkBGigEEAA/xUsEEAAWTPAWcP/dCQI/3QkCGisEEAA/xUsEEAAg8QMw/90JARouBBAAP8VLBBAAFlZw4tMJARWM/YzwIA5LXUEagFeQYoRgPowfBGA+jl/DA++0o0EgI1EQtDr54X2XnQC99jDi0QkBIA4AHQBQIoIgPk6dAiA+SB0AzPAw0BQ6K7///9Zi0wkCGoBiQFYw1WL7IPsZFaLdQihBBFAAFdqGIlFnFkzwP92aI19oPOr/3Zwiz0sEEAA/3ZsaPQQQAD/14PEEIM9oBxAAAB1NItGeIXAdgeD+P91KIXAdB2D+P90B2jsEEAA6wVo5BBAAI1FnFD/FSAQQADrHGjUEEAA6+3/dniNRZxoyBBAAFD/FXAQQACDxAz2RQwBdA+NRZxojBxAAFD/FSQQQACNRZxQaMAQQAD/11lZX17Jw1WL7IHsvAAAAFNWM8BXiUX8iUX4iUX0iUXsx0Xw/v////8VGBBAAIvw/xUcEEAAPQAAAIAbwPfYo6AcQACKBjwidQ6KRgFGhMB0FDwidBDr8oTAdAo8IHQGikYBRuvygD4AdAFGgD4gdPpqBF9qAluKBjwvdAg8LQ+FNwEAAITAD4QvAQAAD75GAUaD+Fl/RA+EygAAAIP4TH8XdFKD6D90WSvHdGdIdFsrx3RL6fcAAACD6FJ0eUgPhOcAAACD6AMPhNcAAAArww+EsAAAAOnVAAAAg/hyf3Z0VYPoY3QtSHQhK8d0ESvHD4W6AAAAg038IOmxAAAACX38Rgld/OmlAAAAg038COmcAAAAjUXsUFboEP7//1mFwFl0AgPzigY8IA+EgAAAAITAdHxG6++NRfBQVujt/f//WYXAWXQCA/OKBjwgdGGEwHRdRuvzg+hzdFGD6AN0RSvDdCJIdUmNRfRQVui9/f//WYXAWXQCA/OKBjwgdDGEwHQtRuvzjUX4UFbonv3//1mFwFl0AgPzigY8IHQShMB0Dkbr80aDTfwB6wSDTfwQgD4gD4W+/v//Ruv09kX8AYsdLBBAAHUIaEwUQAD/01n2RfwgdFOLNXwQQABqAV+NhUT///9QM9tXU//WhcAPhHUDAABHg728AXQjgX2wgAIAAHIaM8A5HaAcQAAPlMBQjYVE////UOg9/f//WVmNhUT///9QV1PrwfZF/BAPhMoAAABqAP8VeBBAAIv4hf8PhCQDAACLNRAQQABqCFf/1moKiUWwW1NX/9ZqDFeJRbT/1mp0V4lFrP/WhcCJRbx1bjP2OTWgHEAAdWaNRehQaBkAAgBWaDgUQABoBQAAgP8VCBBAAIXAdUiNReSJXeRQjUXYUI1F/FBWaCwUQAD/dej/FQQQQACFwHUZg338AXQGg338AnUNjUXYUOgs/P//WYlFvP916P8VABBAAOsCM/aNhUT///9WUOhr/P//WVlXVv8VbBBAAOlsAgAA9kX8Ag+FTgEAAIN9+ACLRfR1E4XAdQ85Rex1D4N98P4PhDIBAACD+AF9DotF+Jn3/40EQIlF9OsSg334AX0MweACagOZWff5iUX4aJQAAACNhUT///9qAFDoFgIAAItF9ItN+ItV8IlFtItF7IPEDIXAZseFaP///5QAiU2wiUWsiVW8fgqBhWz///8AAAQAhcl+CoGFbP///wAAGACD+v6+AABAAHU4gz2gHEAAAHQ1agD/FXgQQACL+IX/dBsBtWz///9qdFf/FRAQQABXagCJRbz/FWwQQACDffD+dAYBtWz///+LNXQQQACNhUT///9qAlD/1ov4hf91IItF/PfQwegDg+ABUI2FRP///1D/1mggFEAAi/j/0+sWi8dIdAdo/BNAAOsFaLATQADoj/r//4P//VkPhS8BAABoZBNAAOh7+v//WY2FRP///2oAUP/W6RQBAABoFBNAAP/TxwQkABNAAGj8EkAA6Gb6//9o6BJAAGjkEkAA6Ff6//9o1BJAAGjQEkAA6Ej6//+LdfyDxBgj93Q7aMASQADoS/r//8cEJLASQADoP/r//8cEJKASQADoM/r//8cEJJASQADoJ/r//8cEJIASQADoG/r//1locBJAAGhsEkAA6PT5//+DPaAcQAAAWVl1G4X2dBdoVBJAAOjy+f//xwQkRBJAAOjm+f//WWgkEkAAaCASQADov/n//2gIEkAAaAQSQADosPn//2jQEUAAaMwRQADoofn//2ikEUAAaKARQADokvn//2iEEUAAaIARQADog/n//2hsEUAAaGgRQADodPn//2gIEUAA/9ODxDRfXjPAW8nDzP8lQBBAAFWL7Gr/aIAUQABogBxAAGShAAAAAFBkiSUAAAAAg+wgU1ZXiWXog2X8AGoB/xVUEEAAWYMNpBxAAP+DDagcQAD//xVkEEAAiw2cHEAAiQj/FWAQQACLDZgcQACJCKFcEEAAiwCjrBxAAOjDAAAAgz14FEAAAHUMaHYcQAD/FVgQQABZ6JQAAABokBBAAGiMEEAA6H8AAAChlBxAAIlF2I1F2FD/NZAcQACNReBQjUXUUI1F5FD/FTAQQABoiBBAAGiEEEAA6EwAAAD/FVAQQACLTeCJCP914P911P915Oit+f//g8QwiUXcUP8VTBBAAItF7IsIiwmJTdBQUegPAAAAWVnDi2Xo/3XQ/xVEEEAA/yVIEEAA/yU0EEAAaAAAAwBoAAABAOgTAAAAWVnDM8DDw8zMzMzMzP8lPBBAAP8lOBBAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAEAdAAAAAAAAAAAAAOQdAAAYEAAAlB0AAAAAAAAAAAAARB4AAGwQAAA4HQAAAAAAAAAAAABgHgAAEBAAACgdAAAAAAAAAAAAAJweAAAAEAAAVB0AAAAAAAAAAAAAvh4AACwQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAGoeAAB4HgAAjB4AAAAAAABQHgAAAAAAANIdAADEHQAAuB0AAKwdAAAAAAAAqh4AAPgeAAAIHwAAfB8AAGgfAAC0HgAAyh4AANIeAADgHgAA6B4AAFYfAAAUHwAAKB8AADgfAABIHwAAAAAAABgeAADyHQAA/h0AACQeAAAsHgAAAAAAAAIDbHN0cmNweUEAAPkCbHN0cmNhdEEAAHQBR2V0VmVyc2lvbgAAygBHZXRDb21tYW5kTGluZUEAS0VSTkVMMzIuZGxsAACsAndzcHJpbnRmQQAbAENoYW5nZURpc3BsYXlTZXR0aW5nc0EAAAMCUmVsZWFzZURDAP0AR2V0REMAxQBFbnVtRGlzcGxheVNldHRpbmdzQQAAVVNFUjMyLmRsbAAAJQFHZXREZXZpY2VDYXBzAEdESTMyLmRsbABbAVJlZ0Nsb3NlS2V5AHsBUmVnUXVlcnlWYWx1ZUV4QQAAcgFSZWdPcGVuS2V5RXhBAEFEVkFQSTMyLmRsbAAAngJwcmludGYAAJkCbWVtc2V0AABNU1ZDUlQuZGxsAADTAF9leGl0AEgAX1hjcHRGaWx0ZXIASQJleGl0AABkAF9fcF9fX2luaXRlbnYAWABfX2dldG1haW5hcmdzAA8BX2luaXR0ZXJtAIMAX19zZXR1c2VybWF0aGVycgAAnQBfYWRqdXN0X2ZkaXYAAGoAX19wX19jb21tb2RlAABvAF9fcF9fZm1vZGUAAIEAX19zZXRfYXBwX3R5cGUAAMoAX2V4Y2VwdF9oYW5kbGVyMwAAtwBfY29udHJvbGZwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
"@
            [IO.File]::WriteAllBytes($tempExe, [Convert]::FromBase64String($base64))
            $qresHash = (Get-FileHash -Path $tempExe -Algorithm SHA256).Hash.ToLowerInvariant()
            if ($qresHash -ne $script:EmbeddedQResSha256) {
                throw "Embedded QRes payload hash mismatch. Expected $($script:EmbeddedQResSha256), got $qresHash."
            }

        # ---- Create SetupComplete.cmd ----
        Write-Log "Creating SetupComplete.cmd..."
        $cmdFile = Join-Path $VMLoc "SetupComplete.cmd"
        $lines = @(
            '@echo off'
            ':: SetupComplete.cmd - runs after Windows Setup, before first logon'
            'setlocal enableextensions'
            'set WORKDIR=C:\Windows\Temp'
            'set LOGFILE=%WORKDIR%\SetupComplete.log'
            'echo [%date% %time%] Starting SetupComplete.cmd >> %LOGFILE%'
            ''
        )

        if ($SelectedResolution -match '^(\d+)x(\d+)$') {
            $resX = $matches[1]
            $resY = $matches[2]
            $lines += @(
                ':: --- Set display resolution ---'
                'if /i "%PROCESSOR_ARCHITECTURE%"=="ARM64" (echo [%date% %time%] QRes skipped: not supported on ARM64 >> %LOGFILE% & goto :AfterQRes)'
                'echo [%date% %time%] Setting display resolution... >> %LOGFILE%'
                ('if exist "%WORKDIR%\QRes.exe" start /wait "" "%WORKDIR%\QRes.exe" /x:{0} /y:{1}' -f $resX, $resY)
                ':AfterQRes'
                ''
            )
        }

        if ($ctrlCreate["Parsec"].Checked) {
            $lines += @(
                ':: --- Parsec ---'
                'if /i "%PROCESSOR_ARCHITECTURE%"=="ARM64" (echo [%date% %time%] Parsec install skipped on ARM64 guest architecture >> %LOGFILE% & goto :AfterParsec)'
                'echo [%date% %time%] Downloading Parsec... >> %LOGFILE%'
                'powershell -NoProfile -Command "$ErrorActionPreference = ''Stop''; [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; $ProgressPreference = ''SilentlyContinue''; Invoke-WebRequest -UseBasicParsing -Uri ''https://builds.parsecgaming.com/package/parsec-windows.exe'' -OutFile ''%WORKDIR%\parsec.exe''; $sig = Get-AuthenticodeSignature ''%WORKDIR%\parsec.exe''; if ($sig.Status -ne ''Valid'' -or $sig.SignerCertificate.Subject -notmatch ''Parsec'') { throw ''Parsec signature validation failed (status/signer mismatch)'' }" >> %LOGFILE% 2>&1'
                'if errorlevel 1 (echo [%date% %time%] Parsec download/signature validation failed >> %LOGFILE% & goto :AfterParsec)'
                'echo [%date% %time%] Installing Parsec... >> %LOGFILE%'
                'start /wait %WORKDIR%\parsec.exe /silent /percomputer /norun /vdd'
                ':AfterParsec'
                ''
            )
        }
        if ($ctrlCreate["VBCable"].Checked) {
            $lines += @(
                ':: --- VB Cable ---'
                'echo [%date% %time%] Downloading VB Cable... >> %LOGFILE%'
                'powershell -NoProfile -Command "$ErrorActionPreference = ''Stop''; [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; $ProgressPreference = ''SilentlyContinue''; Invoke-WebRequest -UseBasicParsing -Uri ''https://download.vb-audio.com/Download_CABLE/VBCABLE_Driver_Pack45.zip'' -OutFile ''%WORKDIR%\vb.zip''" >> %LOGFILE% 2>&1'
                'if errorlevel 1 (echo [%date% %time%] VB Cable download failed >> %LOGFILE% & goto :AfterVBCable)'
                'if not exist "%WORKDIR%\VB" mkdir "%WORKDIR%\VB"'
                'powershell -NoProfile -Command "$ErrorActionPreference = ''Stop''; Expand-Archive -Path ''%WORKDIR%\vb.zip'' -DestinationPath ''%WORKDIR%\VB'' -Force" >> %LOGFILE% 2>&1'
                'if errorlevel 1 (echo [%date% %time%] VB Cable archive extraction failed >> %LOGFILE% & goto :AfterVBCable)'
                'if /i "%PROCESSOR_ARCHITECTURE%"=="ARM64" (echo [%date% %time%] VB Cable install skipped on ARM64 guest architecture >> %LOGFILE% & goto :AfterVBCable)'
                'powershell -NoProfile -Command "$ErrorActionPreference = ''Stop''; $sig = Get-AuthenticodeSignature ''%WORKDIR%\VB\VBCABLE_Setup_x64.exe''; if ($sig.Status -ne ''Valid'' -or $sig.SignerCertificate.Subject -notmatch ''VB[- ]Audio|Vincent Burel'') { throw ''VB Cable signature validation failed (status/signer mismatch)'' }" >> %LOGFILE% 2>&1'
                'if errorlevel 1 (echo [%date% %time%] VB Cable signature validation failed >> %LOGFILE% & goto :AfterVBCable)'
                'start /wait "" "%WORKDIR%\VB\VBCABLE_Setup_x64.exe" -h -i -H -n'
                ':AfterVBCable'
                ''
            )
        }
        if ($ctrlCreate["USBMMIDD"].Checked) {
            $lines += @(
                ':: --- Virtual Display Driver ---'
                'echo [%date% %time%] Downloading USBMMIDD... >> %LOGFILE%'
                'powershell -NoProfile -Command "$ErrorActionPreference = ''Stop''; [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; $ProgressPreference = ''SilentlyContinue''; Invoke-WebRequest -UseBasicParsing -Uri ''https://www.amyuni.com/downloads/usbmmidd_v2.zip'' -OutFile ''%WORKDIR%\usbmmidd_v2.zip''" >> %LOGFILE% 2>&1'
                'if errorlevel 1 (echo [%date% %time%] USBMMIDD download failed >> %LOGFILE% & goto :AfterUSBMMIDD)'
                'if not exist "%WORKDIR%\usbmmidd_v2" mkdir "%WORKDIR%\usbmmidd_v2"'
                'powershell -NoProfile -Command "$ErrorActionPreference = ''Stop''; Expand-Archive -Path ''%WORKDIR%\usbmmidd_v2.zip'' -DestinationPath ''%WORKDIR%'' -Force" >> %LOGFILE% 2>&1'
                'if errorlevel 1 (echo [%date% %time%] USBMMIDD archive extraction failed >> %LOGFILE% & goto :AfterUSBMMIDD)'
                'if /i "%PROCESSOR_ARCHITECTURE%"=="ARM64" (echo [%date% %time%] USBMMIDD install skipped on ARM64 guest architecture >> %LOGFILE% & goto :AfterUSBMMIDD)'
                'powershell -NoProfile -Command "$ErrorActionPreference = ''Stop''; $sig = Get-AuthenticodeSignature ''%WORKDIR%\usbmmidd_v2\deviceinstaller64.exe''; if ($sig.Status -ne ''Valid'' -or $sig.SignerCertificate.Subject -notmatch ''Amyuni'') { throw ''USBMMIDD signature validation failed (status/signer mismatch)'' }" >> %LOGFILE% 2>&1'
                'if errorlevel 1 (echo [%date% %time%] USBMMIDD signature validation failed >> %LOGFILE% & goto :AfterUSBMMIDD)'
                '@echo off'
                'setlocal DisableDelayedExpansion'
                'echo @cd /d "%%~dp0" > "%WORKDIR%\usbmmidd_v2\usbmmidd2.bat"'
                'echo @goto %%PROCESSOR_ARCHITECTURE%% >> "%WORKDIR%\usbmmidd_v2\usbmmidd2.bat"'
                'echo @exit >> "%WORKDIR%\usbmmidd_v2\usbmmidd2.bat"'
                'echo. >> "%WORKDIR%\usbmmidd_v2\usbmmidd2.bat"'
                'echo :AMD64 >> "%WORKDIR%\usbmmidd_v2\usbmmidd2.bat"'
                'echo @cmd /c deviceinstaller64.exe install usbmmidd.inf usbmmidd >> "%WORKDIR%\usbmmidd_v2\usbmmidd2.bat"'
                'echo deviceinstaller64.exe enableidd 1 ^&^& exit >> "%WORKDIR%\usbmmidd_v2\usbmmidd2.bat"'
                'echo @goto end >> "%WORKDIR%\usbmmidd_v2\usbmmidd2.bat"'
                'echo. >> "%WORKDIR%\usbmmidd_v2\usbmmidd2.bat"'
                'echo :x86 >> "%WORKDIR%\usbmmidd_v2\usbmmidd2.bat"'
                'echo @cmd /c deviceinstaller.exe install usbmmidd.inf usbmmidd >> "%WORKDIR%\usbmmidd_v2\usbmmidd2.bat"'
                'echo deviceinstaller.exe enableidd 1 ^&^& exit >> "%WORKDIR%\usbmmidd_v2\usbmmidd2.bat"'
                'echo. >> "%WORKDIR%\usbmmidd_v2\usbmmidd2.bat"'
                'echo :end >> "%WORKDIR%\usbmmidd_v2\usbmmidd2.bat"'
                'start /wait %WORKDIR%\usbmmidd_v2\usbmmidd2.bat'
                ':AfterUSBMMIDD'
                ''
            )
        }
        if ($ctrlCreate["RDP"].Checked) {
            $lines += @(
                ':: --- Remote Desktop ---'
                'reg add "HKLM\SYSTEM\CurrentControlSet\Control\Terminal Server" /v fDenyTSConnections /t REG_DWORD /d 0 /f'
                'netsh advfirewall firewall set rule group="remote desktop" new enable=Yes'
                ''
            )
        }
        if ($ctrlCreate["Share"].Checked) {
            $lines += @(
                ':: --- Share Folder ---'
                'reg add "HKLM\Software\Microsoft\Windows\CurrentVersion\RunOnce" /v CreateShare /d "cmd /c C:\Windows\Temp\CreateShare.cmd" /f'
                'echo @echo off > C:\Windows\Temp\CreateShare.cmd'
                'echo set "SHAREFOLDER=%%USERPROFILE%%\Desktop\share" >> C:\Windows\Temp\CreateShare.cmd'
                'echo if not exist "%%SHAREFOLDER%%" mkdir "%%SHAREFOLDER%%" >> C:\Windows\Temp\CreateShare.cmd'
                'echo icacls "%%SHAREFOLDER%%" /inheritance:r /grant:r "%%USERNAME%%:(OI)(CI)M" "Administrators:(OI)(CI)F" /T >> C:\Windows\Temp\CreateShare.cmd'
                'echo powershell -Command "Set-NetFirewallRule -DisplayGroup ''File and Printer Sharing'' -Enabled True" >> C:\Windows\Temp\CreateShare.cmd'
                'echo net share share /delete /y ^>nul 2^>nul >> C:\Windows\Temp\CreateShare.cmd'
                'echo net share share="%%USERPROFILE%%\Desktop\share" /grant:%%USERNAME%%,FULL /grant:Administrators,FULL >> C:\Windows\Temp\CreateShare.cmd'
                'echo exit >> C:\Windows\Temp\CreateShare.cmd'
                ''
            )
        }
        if ($ctrlCreate["PauseUpdate"].Checked) {
            $lines += @(
                ':: --- Pause Windows Updates 1 year ---'
                'powershell -NoProfile -Command "$now = (Get-Date).ToString(''yyyy-MM-ddTHH:mm:ssZ''); $future = (Get-Date).AddDays(365).ToString(''yyyy-MM-ddTHH:mm:ssZ''); $wuPath = ''HKLM:\SOFTWARE\Microsoft\WindowsUpdate\UX\Settings''; if (-not (Test-Path $wuPath)) { New-Item -Path $wuPath -Force | Out-Null }; Set-ItemProperty -Path $wuPath -Name ''PauseFeatureUpdatesStartTime'' -Value $now; Set-ItemProperty -Path $wuPath -Name ''PauseFeatureUpdatesEndTime'' -Value $future; Set-ItemProperty -Path $wuPath -Name ''PauseFeatureUpdates'' -Value 1 -Type DWord; Set-ItemProperty -Path $wuPath -Name ''PauseQualityUpdatesStartTime'' -Value $now; Set-ItemProperty -Path $wuPath -Name ''PauseQualityUpdatesEndTime'' -Value $future; Set-ItemProperty -Path $wuPath -Name ''PauseQualityUpdates'' -Value 1 -Type DWord"'
                ''
            )
        }
        if ($ctrlCreate["FullUpdate"].Checked) {
            $lines += @(
                ':: --- Full Windows Updates at first logon ---'
                'reg add "HKLM\Software\Microsoft\Windows\CurrentVersion\RunOnce" /v RunUpdates /d "cmd /c C:\Windows\Temp\RunUpdates.cmd" /f'
                'echo @echo off > C:\Windows\Temp\RunUpdates.cmd'
                'echo powershell -NoProfile -Command "try { Install-PackageProvider -Name NuGet -Force; [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; if (-not (Get-PSRepository -Name ''PSGallery'' -ErrorAction SilentlyContinue)) { throw ''PSGallery repository not found'' }; Set-PSRepository -Name ''PSGallery'' -InstallationPolicy Trusted; Install-Module PSWindowsUpdate -Repository PSGallery -RequiredVersion 2.2.1.5 -Force -Scope AllUsers; $mod = Get-Module -ListAvailable -Name PSWindowsUpdate | Sort-Object Version -Descending | Select-Object -First 1; if (-not $mod) { throw ''PSWindowsUpdate module missing after install'' }; $sig = Get-AuthenticodeSignature (Join-Path $mod.ModuleBase ''PSWindowsUpdate.psd1''); if ($sig.Status -ne ''Valid'') { throw ''PSWindowsUpdate signature validation failed'' }; Import-Module PSWindowsUpdate -RequiredVersion 2.2.1.5 -Force; Add-WUServiceManager -MicrosoftUpdate -Confirm:$false -ErrorAction SilentlyContinue; Get-WindowsUpdate -MicrosoftUpdate -Install -AcceptAll -IgnoreReboot | Out-File -FilePath C:\\Windows\\Temp\\WUOutput.log -Encoding UTF8; } catch { Write-Output $_.Exception.Message }" >> C:\Windows\Temp\RunUpdates.cmd'
                'echo shutdown /r /t 30 /c "Windows Updates complete. Rebooting in 30 seconds." >> C:\Windows\Temp\RunUpdates.cmd'
                'echo exit >> C:\Windows\Temp\RunUpdates.cmd'
                ''
            )
        }
        # ---- Cleanup unattend files that may contain plaintext credentials (#12) ----
        $lines += @(
            ''
            ':SetupCompleteCleanup'
            ':: --- Remove unattend XML files (may contain plaintext passwords) ---'
            'echo [%date% %time%] Cleaning up unattend files >> %LOGFILE%'
            'del /f /q C:\Autounattend.xml 2>nul'
            'del /f /q C:\Windows\Panther\Unattend\Unattend.xml 2>nul'
            'del /f /q C:\Windows\System32\Sysprep\Unattend.xml 2>nul'
            'echo [%date% %time%] Unattend cleanup complete >> %LOGFILE%'
        )
        $lines += @(
            'echo [%date% %time%] SetupComplete.cmd finished >> %LOGFILE%'
            'endlocal'
            'exit /b 0'
        )
        $lines | Out-File -FilePath $cmdFile -Encoding ASCII -Force

            # ---- Generate Unattend XML ----
            Update-CreateProgress -Percent 22 -Status "Generating unattended setup..."
            Write-Log "Generating Autounattend.xml..."
            $UnattendXMLPath = Join-Path $VMLoc "Autounattend.xml"
            $guestArchForUnattend = if ($UseGoldenImage) { $script:HostArch } else { $script:DetectedGuestArch }
            $unattendContent = New-UnattendXml -VMName $VMName -Username $Username -Password $Password -EnableAutoLogon $EnableAutoLogon -IsWindows11 $IsWin11 -GuestArch $guestArchForUnattend
            [IO.File]::WriteAllText($UnattendXMLPath, $unattendContent, [System.Text.UTF8Encoding]::new($false))

            # Minimize plaintext password lifetime in memory
            $PasswordText = $null
            if ($Password) {
                try {
                    $Password.Dispose()
                } catch {
                    Write-Log "SecureString dispose warning: $($_.Exception.Message)" "WARN"
                }
                $Password = $null
            }

            # ---- Create VHDX ----
            Update-CreateProgress -Percent 30 -Status "Creating virtual disk..."
            $VHDPath = Join-Path $VMLoc "$VMName.vhdx"
            if ($FixedVHD) {
                Write-Log "Creating fixed VHDX ($DiskGB GB)..."
                New-VHD -Path $VHDPath -SizeBytes ($DiskGB * 1GB) -Fixed -ErrorAction Stop | Out-Null
            } else {
                Write-Log "Creating dynamic VHDX ($DiskGB GB max)..."
                New-VHD -Path $VHDPath -SizeBytes ($DiskGB * 1GB) -Dynamic -ErrorAction Stop | Out-Null
            }
            if (-not (Test-Path $VHDPath)) {
                throw "VHD creation command completed but disk file was not found at '$VHDPath'."
            }

        # ---- Mount VHD and partition ----
            Update-CreateProgress -Percent 38 -Status "Partitioning virtual disk..."
            Write-Log "Mounting VHD and creating GPT/EFI/MSR/Windows partitions..."
            $mountedVHD = Mount-VHD -Path $VHDPath -Passthru -ErrorAction Stop
            Register-TrackedMountedImage -ImagePath $VHDPath
            $vhdMountedForDeploy = $true
            try {
            $diskNumber = $mountedVHD.DiskNumber
            Initialize-Disk -Number $diskNumber -PartitionStyle GPT -PassThru | Out-Null

            # EFI System Partition (260MB for better compatibility with old Win10)
            # NOTE: -AssignDriveLetter is attempted here, but Windows often auto-removes
            # drive letters from ESP partitions after a short delay.  The boot configuration
            # section below robustly re-acquires a letter right before bcdboot runs.
            $efi = New-Partition -DiskNumber $diskNumber -Size 260MB -GptType "{C12A7328-F81F-11D2-BA4B-00A0C93EC93B}" -AssignDriveLetter
            if (-not $efi.DriveLetter) {
                Write-Log "EFI partition auto-letter not assigned (will acquire before boot config)." "INFO"
            }
            Format-Volume -Partition $efi -FileSystem FAT32 -NewFileSystemLabel "System" -Confirm:$false -ErrorAction Stop | Out-Null

            # MSR Partition
            New-Partition -DiskNumber $diskNumber -Size 16MB -GptType "{E3C9E316-0B5C-4DB8-817D-F92DF00215AE}" | Out-Null

            # Windows Partition (remaining space)
            $winPart = New-Partition -DiskNumber $diskNumber -UseMaximumSize -AssignDriveLetter
            if (-not $winPart.DriveLetter) {
                Write-Log "Windows partition created but no drive letter was assigned. Aborting." "ERROR"
                throw "Windows partition has no drive letter"
            }
            $driveLetter = $winPart.DriveLetter + ":"
            Format-Volume -Partition $winPart -FileSystem NTFS -NewFileSystemLabel $VMName -Confirm:$false -ErrorAction Stop | Out-Null

            # Wait for drive to be ready
            $driveReady = $false
            for ($w = 0; $w -lt 15; $w++) {
                if (Test-Path "$driveLetter\") { $driveReady = $true; break }
                Start-Sleep -Seconds 1
            }
            if (-not $driveReady) {
                Write-Log "Windows partition drive not ready at $driveLetter after 15 seconds. Aborting." "ERROR"
                throw "Windows partition drive not accessible at $driveLetter"
            }

            # ---- Apply Windows image ----
            Update-CreateProgress -Percent 48 -Status "Applying Windows image (this can take a while)..."
            Write-Log "Applying Windows image (Edition: $SelectedEdition, Index: $SelectedIndex)..."
        Write-Log "This may take several minutes..."
        # Trailing backslash is required: 'E:' targets 'current dir on E' vs 'E:\' targets the root.
        # Microsoft docs always use the trailing form (e.g., W:\) for /ApplyDir.
        $applyDirRoot = $driveLetter + "\"
        Invoke-DismApplyImage -ImageFile $script:WimFile -Index $SelectedIndex -ApplyDir $applyDirRoot -PreferCompactApply $guestProfile.PreferCompactApply
        Write-Log "Windows image applied." "OK"

        # ---- Inject files into VHD ----
        # For an offline DISM deploy (no setup.exe), Windows performs a "mini-setup"
        # on first boot (specialize + oobeSystem passes).  The implicit answer file
        # search order that matters here is:
        #   Priority 2: %WINDIR%\Panther\Unattend  (downlevel only — may NOT be searched)
        #   Priority 3: %WINDIR%\Panther            (cached — MOST RELIABLE for offline deploy)
        #   Priority 6: %WINDIR%\System32\Sysprep
        #   Priority 7: %SYSTEMDRIVE% root
        # We place the file in ALL locations so it is found regardless of which
        # search paths Windows evaluates on first boot.

        # Unattend to root of Windows drive (priority 7)
        Copy-Item -Path $UnattendXMLPath -Destination (Join-Path "$driveLetter\" "Autounattend.xml") -Force

        # Unattend directly into Panther (priority 3 — highest reliable priority)
        $PantherDirDirect = Join-Path "$driveLetter\Windows" "Panther"
        New-Item -Path $PantherDirDirect -ItemType Directory -Force | Out-Null
        Copy-Item -Path $UnattendXMLPath -Destination (Join-Path $PantherDirDirect "Unattend.xml") -Force

        # Unattend into Panther\Unattend subdirectory (priority 2 — downlevel fallback)
        $PantherSubDir = Join-Path $PantherDirDirect "Unattend"
        New-Item -Path $PantherSubDir -ItemType Directory -Force | Out-Null
        Copy-Item -Path $UnattendXMLPath -Destination (Join-Path $PantherSubDir "Unattend.xml") -Force

        # Also to Sysprep for maximum compatibility (priority 6)
        $SysprepDir = Join-Path "$driveLetter\Windows\System32" "Sysprep"
        if (Test-Path $SysprepDir) {
            Copy-Item -Path $UnattendXMLPath -Destination (Join-Path $SysprepDir "Unattend.xml") -Force
        }
        Write-Log "Unattend.xml injected (Panther, Panther\Unattend, Sysprep, root)" "OK"

        # QRes.exe
        $QresDestDir = Join-Path "$driveLetter\" "Windows\Temp"
        if (-not (Test-Path $QresDestDir)) { New-Item -Path $QresDestDir -ItemType Directory -Force | Out-Null }
        Copy-Item -Path $tempExe -Destination (Join-Path $QresDestDir "QRes.exe") -Force
        Write-Log "QRes.exe injected"

        # SetupComplete.cmd
        $VMSetupScriptsDir = Join-Path "$driveLetter\" "Windows\Setup\Scripts"
        New-Item -Path $VMSetupScriptsDir -ItemType Directory -Force | Out-Null
        Copy-Item -Path $cmdFile -Destination (Join-Path $VMSetupScriptsDir "SetupComplete.cmd") -Force
        Write-Log "SetupComplete.cmd injected" "OK"

        # ---- Create boot files (robust with retry and fallback) ----
        Update-CreateProgress -Percent 70 -Status "Configuring boot files..."

        # ---- Robust EFI partition drive letter acquisition ----
        # Windows auto-removes drive letters from ESP partitions.  The letter captured at
        # New-Partition time is typically gone after the long DISM apply.  We therefore
        # re-acquire a working letter here through multiple methods before bcdboot.
        $efiDrive = $null
        $efiLetterAssignedByUs = $false
        $efiRefreshed = $null

        # Method 1: Check if the originally-assigned letter still works
        if ($efi.DriveLetter -and $efi.DriveLetter -ne [char]0) {
            $candidateDrive = "$($efi.DriveLetter):"
            if (Test-Path "$candidateDrive\") {
                $efiDrive = $candidateDrive
                Write-Log "EFI partition accessible at original letter $efiDrive" "OK"
            }
        }

        # Method 2: Re-query the partition — Windows may have reassigned a different letter
        if (-not $efiDrive) {
            try {
                $efiRefreshed = Get-Partition -DiskNumber $diskNumber |
                    Where-Object { $_.GptType -eq '{c12a7328-f81f-11d2-ba4b-00a0c93ec93b}' } |
                    Select-Object -First 1
            } catch {
                Write-Log "Could not refresh EFI partition metadata: $($_.Exception.Message)" "WARN"
            }
            if ($efiRefreshed -and $efiRefreshed.DriveLetter -and $efiRefreshed.DriveLetter -ne [char]0) {
                $candidateDrive = "$($efiRefreshed.DriveLetter):"
                if (Test-Path "$candidateDrive\") {
                    $efiDrive = $candidateDrive
                    Write-Log "EFI partition accessible at refreshed letter $efiDrive" "OK"
                }
            }
        }

        # Method 3: Assign a fresh drive letter via Add-PartitionAccessPath
        if (-not $efiDrive) {
            Write-Log "EFI drive letter lost (common for ESP partitions). Assigning fresh letter..." "WARN"
            $usedLetters = @()
            Get-Volume | Where-Object { $_.DriveLetter } | ForEach-Object { $usedLetters += $_.DriveLetter }
            Get-CimInstance Win32_MappedLogicalDisk -ErrorAction SilentlyContinue |
                ForEach-Object { if ($_.DeviceID) { $usedLetters += $_.DeviceID[0] } }
            $freeLetter = $null
            # Prefer higher letters (S-Z then G-R) to avoid collisions with common mounts
            foreach ($c in [char[]]@(83,84,85,86,87,88,89,90,71,72,73,74,75,76,77,78,79,80,81,82)) {
                if ($c -notin $usedLetters) { $freeLetter = $c; break }
            }
            if (-not $freeLetter) {
                throw "No free drive letter available for EFI partition (all letters S-Z, G-R in use)."
            }
            $targetPart = if ($efiRefreshed) { $efiRefreshed } else { $efi }
            $accessPathWorked = $false
            try {
                $targetPart | Add-PartitionAccessPath -AccessPath "$($freeLetter):" -ErrorAction Stop
                $efiLetterAssignedByUs = $true
                Start-Sleep -Seconds 3
                if (Test-Path "$($freeLetter):\") {
                    $efiDrive = "$($freeLetter):"
                    $accessPathWorked = $true
                    Write-Log "EFI partition mounted at $efiDrive via Add-PartitionAccessPath" "OK"
                } else {
                    # Sometimes needs a bit longer; wait and retry
                    Start-Sleep -Seconds 3
                    if (Test-Path "$($freeLetter):\") {
                        $efiDrive = "$($freeLetter):"
                        $accessPathWorked = $true
                        Write-Log "EFI partition mounted at $efiDrive via Add-PartitionAccessPath (delayed)" "OK"
                    }
                }
            } catch {
                Write-Log "Add-PartitionAccessPath failed: $($_.Exception.Message). Trying diskpart..." "WARN"
            }

            # Method 4: diskpart — most reliable for ESP partitions on all Windows versions
            if (-not $accessPathWorked) {
                $diskpartExe = Join-Path $env:WINDIR 'System32\diskpart.exe'
                if (-not (Test-Path $diskpartExe)) { $diskpartExe = 'diskpart.exe' }
                $dpScript = @"
select disk $diskNumber
select partition $($targetPart.PartitionNumber)
assign letter=$freeLetter
"@
                $dpFile = Join-Path $env:TEMP "ht_efi_$([guid]::NewGuid().ToString('N').Substring(0,8)).txt"
                try {
                    $dpScript | Out-File -FilePath $dpFile -Encoding ASCII -Force
                    $dpResult = & $diskpartExe /s $dpFile 2>&1
                    $dpExitCode = $LASTEXITCODE
                    Remove-Item $dpFile -Force -ErrorAction SilentlyContinue
                    if ($dpExitCode -eq 0) {
                        $efiLetterAssignedByUs = $true
                        Start-Sleep -Seconds 3
                        if (Test-Path "$($freeLetter):\") {
                            $efiDrive = "$($freeLetter):"
                            Write-Log "EFI partition mounted at $efiDrive via diskpart" "OK"
                        } else {
                            Write-Log "diskpart reported success but $($freeLetter): not accessible." "WARN"
                        }
                    } else {
                        Write-Log "diskpart failed (exit $dpExitCode): $($dpResult | Out-String)" "WARN"
                    }
                } catch {
                    Write-Log "diskpart exception: $($_.Exception.Message)" "WARN"
                    Remove-Item $dpFile -Force -ErrorAction SilentlyContinue
                }
            }
        }

        if (-not $efiDrive) {
            throw "EFI partition inaccessible after all mount attempts (original letter, refresh, Add-PartitionAccessPath, diskpart). Cannot write boot files."
        }

        Write-Log "Creating UEFI boot files on $efiDrive..."

        $bootSuccess = $false
        $attachISOForRecovery = $false

        # Attempt 1: Use the applied image's bcdboot (most compatible for old Win10)
        $imageBcdboot = Join-Path "$driveLetter\Windows\System32" "bcdboot.exe"
        if (Test-Path $imageBcdboot) {
            for ($retry = 1; $retry -le 3; $retry++) {
                try {
                    Write-Log "Boot attempt $retry/3: Using image's bcdboot.exe..."
                    Start-Sleep -Seconds 2
                    $result = & $imageBcdboot "$driveLetter\Windows" /s $efiDrive /f UEFI 2>&1
                    if ($LASTEXITCODE -eq 0) {
                        Write-Log "Boot files created successfully (image bcdboot)" "OK"
                        $bootSuccess = $true
                        break
                    } else {
                        Write-Log "  bcdboot exit code $LASTEXITCODE : $($result | Out-String)" "WARN"
                    }
                } catch {
                    Write-Log "  Exception: $_" "WARN"
                }
            }
        } else {
            Write-Log "Image bcdboot.exe not found at $imageBcdboot; skipping image-native attempt." "INFO"
        }

        # Attempt 2: Use host's bcdboot
        if (-not $bootSuccess) {
            $hostBcdboot = Join-Path $env:WINDIR 'System32\bcdboot.exe'
            if (-not (Test-Path $hostBcdboot)) { $hostBcdboot = 'bcdboot.exe' }
            for ($retry = 1; $retry -le 3; $retry++) {
                try {
                    Write-Log "Boot attempt $retry/3: Using host bcdboot..."
                    Start-Sleep -Seconds 2
                    $result = & $hostBcdboot "$driveLetter\Windows" /s $efiDrive /f UEFI 2>&1
                    if ($LASTEXITCODE -eq 0) {
                        Write-Log "Boot files created successfully (host bcdboot)" "OK"
                        $bootSuccess = $true
                        break
                    } else {
                        Write-Log "  bcdboot exit code $LASTEXITCODE : $($result | Out-String)" "WARN"
                    }
                } catch {
                    Write-Log "  Exception: $_" "WARN"
                }
            }
        }

        # Attempt 3: Host bcdboot with /f ALL for broader compatibility
        if (-not $bootSuccess) {
            $hostBcdboot = Join-Path $env:WINDIR 'System32\bcdboot.exe'
            if (-not (Test-Path $hostBcdboot)) { $hostBcdboot = 'bcdboot.exe' }
            for ($retry = 1; $retry -le 2; $retry++) {
                try {
                    Write-Log "Boot attempt $retry/2: Using host bcdboot with /f ALL fallback..."
                    Start-Sleep -Seconds 2
                    $result = & $hostBcdboot "$driveLetter\Windows" /s $efiDrive /f ALL 2>&1
                    if ($LASTEXITCODE -eq 0) {
                        Write-Log "Boot files created successfully (host bcdboot /f ALL)" "OK"
                        $bootSuccess = $true
                        break
                    } else {
                        Write-Log "  bcdboot exit code $LASTEXITCODE : $($result | Out-String)" "WARN"
                    }
                } catch {
                    Write-Log "  Exception: $_" "WARN"
                }
            }
        }

        # ---- Verify boot files actually landed on the EFI partition ----
        $bootmgrEfi = Join-Path $efiDrive "EFI\Microsoft\Boot\bootmgfw.efi"
        $bcdStore   = Join-Path $efiDrive "EFI\Microsoft\Boot\BCD"
        if ($bootSuccess) {
            if (-not (Test-Path $bootmgrEfi)) {
                Write-Log "bcdboot reported success but bootmgfw.efi NOT found at $bootmgrEfi — marking as failed." "ERROR"
                $bootSuccess = $false
            } elseif (-not (Test-Path $bcdStore)) {
                Write-Log "bcdboot reported success but BCD store NOT found at $bcdStore — marking as failed." "ERROR"
                $bootSuccess = $false
            } else {
                Write-Log "Boot file verification passed: bootmgfw.efi and BCD confirmed on EFI partition." "OK"
            }
        }

        # ---- Always create UEFI fallback boot file ----
        # bcdboot with /s writes \EFI\Microsoft\Boot\ but does NOT create the
        # standard UEFI fallback path \EFI\Boot\bootx64.efi (or bootaa64.efi).
        # Because we run bcdboot offline (no VM NVRAM to write), the Hyper-V Gen 2
        # firmware may not find the boot entry without this fallback file.
        # Per UEFI 2.3.1 spec, firmware looks for \EFI\Boot\boot<arch>.efi when
        # no explicit NVRAM entry exists.
        if ($bootSuccess -and (Test-Path $bootmgrEfi)) {
            try {
                $efiFallbackDir = Join-Path $efiDrive "EFI\Boot"
                New-Item -Path $efiFallbackDir -ItemType Directory -Force | Out-Null
                $fallbackName = if ($script:HostArch -eq 'arm64') { 'bootaa64.efi' } else { 'bootx64.efi' }
                $fallbackPath = Join-Path $efiFallbackDir $fallbackName
                if (-not (Test-Path $fallbackPath)) {
                    Copy-Item -Path $bootmgrEfi -Destination $fallbackPath -Force
                    Write-Log "Created UEFI fallback boot file: $fallbackPath" "OK"
                } else {
                    Write-Log "UEFI fallback boot file already exists: $fallbackPath" "OK"
                }
            } catch {
                Write-Log "Failed to create UEFI fallback boot file: $($_.Exception.Message)" "WARN"
            }
        }

        # Attempt 4 (last resort): Manually copy boot files from the applied Windows image
        if (-not $bootSuccess) {
            Write-Log "Attempting manual boot file copy as last resort..." "WARN"
            try {
                $efiBootDir = Join-Path $efiDrive "EFI\Microsoft\Boot"
                New-Item -Path $efiBootDir -ItemType Directory -Force | Out-Null
                # Copy bootmgfw.efi from the applied image's Windows\Boot\EFI folder
                $srcBootmgr = Join-Path "$driveLetter\Windows\Boot\EFI" "bootmgfw.efi"
                if (Test-Path $srcBootmgr) {
                    Copy-Item -Path $srcBootmgr -Destination $efiBootDir -Force
                    # Copy to fallback UEFI path \EFI\Boot\bootx64.efi (or bootaa64.efi on ARM64)
                    $efiFallbackDir = Join-Path $efiDrive "EFI\Boot"
                    New-Item -Path $efiFallbackDir -ItemType Directory -Force | Out-Null
                    $fallbackName = if ($script:HostArch -eq 'arm64') { 'bootaa64.efi' } else { 'bootx64.efi' }
                    Copy-Item -Path $srcBootmgr -Destination (Join-Path $efiFallbackDir $fallbackName) -Force
                    Write-Log "Copied bootmgfw.efi to EFI partition + fallback $fallbackName" "OK"
                    # Copy remaining boot resources (fonts, locales, memtest, etc.)
                    $srcBootDir = Join-Path "$driveLetter\Windows\Boot\EFI" ""
                    if (Test-Path $srcBootDir) {
                        Get-ChildItem -Path $srcBootDir -Recurse -ErrorAction SilentlyContinue |
                            ForEach-Object {
                                $relPath = $_.FullName.Substring($srcBootDir.Length)
                                $destPath = Join-Path $efiBootDir $relPath
                                if ($_.PSIsContainer) {
                                    New-Item -Path $destPath -ItemType Directory -Force -ErrorAction SilentlyContinue | Out-Null
                                } else {
                                    Copy-Item -Path $_.FullName -Destination $destPath -Force -ErrorAction SilentlyContinue
                                }
                            }
                    }
                    # Re-run bcdboot to create the BCD store referencing these files
                    Start-Sleep -Seconds 1
                    $hostBcdboot = Join-Path $env:WINDIR 'System32\bcdboot.exe'
                    if (-not (Test-Path $hostBcdboot)) { $hostBcdboot = 'bcdboot.exe' }
                    $result = & $hostBcdboot "$driveLetter\Windows" /s $efiDrive /f UEFI 2>&1
                    if ($LASTEXITCODE -eq 0 -and (Test-Path $bcdStore)) {
                        Write-Log "BCD store created after manual boot file copy." "OK"
                        $bootSuccess = $true
                    } else {
                        Write-Log "BCD store creation still failed after manual copy. VM will need repair boot." "ERROR"
                    }
                } else {
                    Write-Log "bootmgfw.efi not found in applied image at $srcBootmgr" "ERROR"
                }
            } catch {
                Write-Log "Manual boot file copy failed: $($_.Exception.Message)" "ERROR"
            }
        }

        # Clean up the temporarily assigned EFI drive letter (will also vanish on VHD dismount)
        if ($efiLetterAssignedByUs -and $efiDrive) {
            try {
                $cleanupPart = if ($efiRefreshed) { $efiRefreshed } else { $efi }
                $cleanupPart | Remove-PartitionAccessPath -AccessPath $efiDrive -ErrorAction SilentlyContinue
            } catch {
                Write-Log "EFI access-path cleanup failed for ${efiDrive}: $($_.Exception.Message)" "WARN"
            }
        }

        # ---- Configure Windows Recovery Environment ----
        # Per Microsoft deployment docs, Winre.wim should be registered so that
        # Windows Update and recovery boot work correctly.  Since we don't create
        # a separate Recovery partition (Gen 2 Hyper-V VMs can re-download from
        # Windows Update), we register WinRE on the Windows partition itself.
        try {
            $winreSource = Join-Path "$driveLetter\Windows\System32\Recovery" "Winre.wim"
            if (Test-Path $winreSource) {
                $winreTargetDir = Join-Path "$driveLetter\" "Recovery\WindowsRE"
                New-Item -Path $winreTargetDir -ItemType Directory -Force | Out-Null
                Copy-Item -Path $winreSource -Destination (Join-Path $winreTargetDir "Winre.wim") -Force
                # Register recovery image with ReAgentC (uses the image's own copy)
                $reagentc = Join-Path "$driveLetter\Windows\System32" "ReAgentC.exe"
                if (Test-Path $reagentc) {
                    $reagentResult = & $reagentc /setreimage /path "$winreTargetDir" /target "$driveLetter\Windows" 2>&1
                    if ($LASTEXITCODE -eq 0) {
                        Write-Log "Windows Recovery Environment registered." "OK"
                    } else {
                        Write-Log "ReAgentC returned code $LASTEXITCODE (non-fatal): $($reagentResult | Out-String)" "WARN"
                    }
                } else {
                    Write-Log "ReAgentC.exe not found in image; WinRE copied but not registered." "WARN"
                }
            } else {
                Write-Log "Winre.wim not found in applied image (non-fatal, Windows Update can fetch recovery later)." "WARN"
            }
        } catch {
            Write-Log "WinRE setup warning (non-fatal): $($_.Exception.Message)" "WARN"
        }

            if (-not $bootSuccess) {
                Write-Log "Boot file creation failed after all attempts! ISO will be attached for Windows Setup repair boot." "ERROR"
                $attachISOForRecovery = $true
            }

            } finally {
            # ---- Dismount VHD (in finally to guarantee cleanup) ----
            if ($vhdMountedForDeploy -and $VHDPath) {
            Update-CreateProgress -Percent 80 -Status "Finalizing disk images..."
            if (Dismount-ImageRetry -ImagePath $VHDPath) {
                Write-Log "VHD dismounted." "OK"
            } else {
                try {
                    Dismount-VHD -Path $VHDPath -ErrorAction SilentlyContinue
                    Unregister-TrackedMountedImage -ImagePath $VHDPath
                } catch {
                    Write-Log "VHD dismount fallback failed: $($_.Exception.Message)" "WARN"
                }
                Write-Log "VHD dismount required fallback and may still be attached." "WARN"
            }
            }
            }

            if ($script:MountedISO) {
                try {
                    if (Dismount-ImageRetry -ImagePath $script:MountedISO.ImagePath) {
                        Write-Log "ISO dismounted."
                    }
                } catch {
                    Write-Log "ISO dismount warning: $($_.Exception.Message)" "WARN"
                } finally {
                    $script:MountedISO = $null
                }
            }
        } else {
            # Golden Image mode: dismount any previously-mounted ISO to avoid a lingering mount
            if ($script:MountedISO) {
                try {
                    if (Dismount-ImageRetry -ImagePath $script:MountedISO.ImagePath -MaxRetries 2) {
                        Write-Log "Dismounted previously-mounted ISO (not needed for golden image mode)."
                    }
                } catch {
                    Write-Log "ISO dismount warning (golden mode): $($_.Exception.Message)" "WARN"
                } finally {
                    $script:MountedISO = $null
                }
            }

            Update-CreateProgress -Percent 30 -Status "Creating differencing disk from golden image..."
            $VHDPath = Join-Path $VMLoc "$VMName.vhdx"

            # Validate parent VHDX exists and destination has enough space for the diff disk
            if (-not (Test-Path $GoldenParentVHD)) {
                throw "Golden parent VHDX not found: $GoldenParentVHD"
            }
            $destFreeGB = Get-PathAvailableSpaceGB -Path $VMLoc
            # Differencing disk starts small but grows; ensure at least 2 GB headroom
            if ($destFreeGB -ge 0 -and $destFreeGB -lt 2) {
                throw "Insufficient disk space for differencing VHDX ($destFreeGB GB free, need >= 2 GB)."
            }

            Write-Log "Creating differencing VHDX from parent: $GoldenParentVHD"
            try {
                # Validate parent VHD accessibility and format
                $parentVhd = Get-VHD -Path $GoldenParentVHD
                if ($parentVhd.VhdFormat -eq 'VHD' -and $parentVhd.Size -gt 127GB) {
                    throw "Parent VHD exceeds 127GB limit for VHD format. Use VHDX or smaller VHD."
                }
                
                New-VHD -Path $VHDPath -ParentPath $GoldenParentVHD -Differencing -ErrorAction Stop | Out-Null
                if (-not (Test-Path $VHDPath)) {
                    throw "Differencing VHDX was not found after creation at '$VHDPath'."
                }
                Write-Log "Golden image differencing disk created." "OK"
            } catch {
                throw "Failed to create differencing VHDX: $($_.Exception.Message)"
            }
            Write-Log "WARNING: Golden Image mode cannot auto-detect the guest OS version." "WARN"
            Write-Log "  - Secure Boot / TPM settings use your selections; verify they match the parent image OS." "WARN"
            Write-Log "  - If the parent is Windows 11, ensure Secure Boot and TPM are both enabled." "WARN"
            # Attempt lightweight OS detection from parent VHD to warn about mismatched settings
            $goldenHiveLoaded = $false
            $goldenParentMounted = $false
            $regExe = Join-Path $env:WINDIR 'System32\reg.exe'
            if (-not (Test-Path $regExe)) { $regExe = 'reg.exe' }
            try {
                $parentMounted = $null
                for ($mountTry = 1; $mountTry -le 3; $mountTry++) {
                    try {
                        $parentMounted = Mount-VHD -Path $GoldenParentVHD -ReadOnly -Passthru -ErrorAction Stop
                        Register-TrackedMountedImage -ImagePath $GoldenParentVHD
                        $goldenParentMounted = $true
                        if (Wait-ImageDetached -ImagePath $GoldenParentVHD -TimeoutSec 1) {
                            throw "Parent VHD did not remain attached after mount attempt."
                        }
                        break
                    } catch {
                        if ($mountTry -lt 3) {
                            Write-Log "Golden parent mount retry ${mountTry}/3 failed: $($_.Exception.Message)" "WARN"
                            Start-Sleep -Seconds (2 * $mountTry)
                        } else {
                            throw
                        }
                    }
                }

                $parentDisk = $parentMounted | Get-Disk -ErrorAction SilentlyContinue
                if ($parentDisk) {
                    $parentParts = Get-DataPartitions -DiskNumber $parentDisk.Number
                    foreach ($pp in $parentParts) {
                        $pVol = Get-Volume -Partition $pp -ErrorAction SilentlyContinue
                        if ($pVol -and $pVol.FileSystem -eq 'NTFS' -and $pVol.DriveLetter) {
                            $regPath = "$($pVol.DriveLetter):\Windows\System32\config\SOFTWARE"
                            if (Test-Path $regPath) {
                                try {
                                    & $regExe load "HKU\__golden_detect" $regPath 2>&1 | Out-Null
                                    if ($LASTEXITCODE -ne 0) { throw "reg load failed with code $LASTEXITCODE" }
                                    $goldenHiveLoaded = $true
                                    $goldenBuild = (Get-ItemProperty -Path 'Registry::HKU\__golden_detect\Microsoft\Windows NT\CurrentVersion' -Name 'CurrentBuild' -ErrorAction SilentlyContinue).CurrentBuild
                                    if ($goldenBuild -and [int]$goldenBuild -ge 22000 -and -not $EnableSecureBoot) {
                                        Write-Log "Golden parent appears to be Windows 11 (Build $goldenBuild) but Secure Boot is disabled. Enable it to avoid boot failures." "WARN"
                                    }
                                    if ($goldenBuild -and [int]$goldenBuild -ge 22000 -and -not $EnableTPM) {
                                        Write-Log "Golden parent appears to be Windows 11 (Build $goldenBuild) but TPM is disabled. Enable it to avoid boot failures." "WARN"
                                    }
                                    if ($goldenBuild) { Write-Log "Golden parent detected build: $goldenBuild" "INFO" }
                                } catch {
                                    Write-Log "Golden parent build detection warning: $($_.Exception.Message)" "WARN"
                                } finally {
                                    if ($goldenHiveLoaded) {
                                        & $regExe unload "HKU\__golden_detect" 2>&1 | Out-Null
                                        $goldenHiveLoaded = $false
                                    }
                                }
                                break
                            }
                        }
                    }
                }
            } catch {
                Write-Log "Could not auto-detect golden parent OS (non-critical): $($_.Exception.Message)" "WARN"
            } finally {
                if ($goldenHiveLoaded) {
                    try {
                        & $regExe unload "HKU\__golden_detect" 2>&1 | Out-Null
                        $goldenHiveLoaded = $false
                    } catch {
                        Write-Log "Golden detect hive unload warning: $($_.Exception.Message)" "WARN"
                    }
                }
                if ($goldenParentMounted) {
                    if (-not (Dismount-ImageRetry -ImagePath $GoldenParentVHD -MaxRetries 3)) {
                        try {
                            Dismount-VHD -Path $GoldenParentVHD -ErrorAction SilentlyContinue
                            if (-not (Wait-ImageDetached -ImagePath $GoldenParentVHD -TimeoutSec 12)) {
                                Write-Log "Golden parent VHD still appears attached after fallback dismount." "WARN"
                            }
                        } catch {
                            Write-Log "Golden parent cleanup warning: $($_.Exception.Message)" "WARN"
                        }
                    }
                    Unregister-TrackedMountedImage -ImagePath $GoldenParentVHD
                }
            }
        }

        # ---- Create Hyper-V VM ----
        Update-CreateProgress -Percent 88 -Status "Creating Hyper-V VM..."
        Write-Log "Creating Generation 2 Hyper-V VM..."
        New-VM -Name $VMName -MemoryStartupBytes ($MemGB * 1GB) -Generation 2 -VHDPath $VHDPath -Path $VMLoc -SwitchName $VMSwitch -ErrorAction Stop | Out-Null
        if (-not (Get-VM -Name $VMName -ErrorAction SilentlyContinue)) {
            throw "Hyper-V did not report VM '$VMName' after New-VM completed."
        }

        # Processor
        Set-VM -Name $VMName -ProcessorCount $vCPU -ErrorAction Stop
        Write-Log "  vCPUs: $vCPU"

        if ($EnableVmNotes) {
            try {
                $sourceRef = if ($UseGoldenImage) { "GoldenParent=$GoldenParentVHD" } else { "ISO=$ISOPath; Edition=$SelectedEdition; Build=$($script:DetectedBuild)" }
                $noteText = "CreatedBy=HyperV-Toolkit; Created=$(Get-Date -Format s); HostBuild=$($script:HostBuild); $sourceRef"
                Set-VM -Name $VMName -Notes $noteText
                Write-Log "  VM Notes: Added provenance metadata" "OK"
            } catch {
                Write-Log "  VM Notes setup failed: $($_.Exception.Message)" "WARN"
            }
        }

        # Secure Boot
        $secureBootResult = Set-VMGuestSecureBoot -VMName $VMName -EnableSecureBoot $EnableSecureBoot -GuestIsWindows11 $IsWin11 -GuestBuild $guestProfile.Build -TemplateOrder $guestProfile.SecureBootTemplateOrder
        if ($EnableSecureBoot -and -not $secureBootResult) {
            Write-Log "  Secure Boot was requested but template configuration could not be fully applied. Verify firmware settings in the VM manually." "WARN"
        }

        # TPM setup (host physical TPM not required)
        if ($EnableTPM) {
            [void](Enable-VirtualTpmForVm -VMName $VMName -GuestIsWindows11 $IsWin11)
        }

        # Checkpoints
        Set-VM -Name $VMName -CheckpointType $CheckpointMode -ErrorAction Stop
        Write-Log "  Checkpoint Mode: $CheckpointMode"

        # Dynamic Memory
        try {
            if ($EnableDynamicMem) {
                Set-VMMemory -VMName $VMName -DynamicMemoryEnabled $true -StartupBytes ($MemGB * 1GB) -MinimumBytes ($DynamicMemMinGB * 1GB) -MaximumBytes ($DynamicMemMaxGB * 1GB) -ErrorAction Stop
            } else {
                Set-VMMemory -VMName $VMName -DynamicMemoryEnabled $false -StartupBytes ($MemGB * 1GB) -ErrorAction Stop
            }
            Write-Log "  Dynamic Memory: $(if ($EnableDynamicMem){"Enabled (min ${DynamicMemMinGB}GB, max ${DynamicMemMaxGB}GB)"}else{"Disabled"})"
        } catch {
            Write-Log "  Dynamic Memory configuration failed: $($_.Exception.Message)" "WARN"
        }

        # Enhanced Session
        if ($EnableEnhancedSession) {
            try {
                $enableEnhancedSession = $true
                try {
                    $hostInfo = Get-VMHost -ErrorAction Stop
                    if ($hostInfo -and -not $hostInfo.EnableEnhancedSessionMode) {
                        $confirmEnhancedSession = [System.Windows.Forms.MessageBox]::Show(
                            "Enhanced Session mode is a host-wide setting and affects all VMs on this host.`n`nEnable host Enhanced Session mode now?",
                            "Confirm Host-Wide Change",
                            [System.Windows.Forms.MessageBoxButtons]::YesNo,
                            [System.Windows.Forms.MessageBoxIcon]::Warning)
                        if ($confirmEnhancedSession -ne [System.Windows.Forms.DialogResult]::Yes) {
                            $enableEnhancedSession = $false
                            Write-Log "  Enhanced Session: skipped host-wide change by user choice." "WARN"
                        }
                    }
                } catch {
                    Write-Log "  Enhanced Session pre-check warning: $($_.Exception.Message)" "WARN"
                }

                if (-not $enableEnhancedSession) {
                    throw "Enhanced Session host-wide change was not approved"
                }
                Set-VMHost -EnableEnhancedSessionMode $true -ErrorAction Stop
                Write-Log "  Enhanced Session: Enabled on host" "OK"
            } catch {
                Write-Log "  Enhanced Session enable failed: $($_.Exception.Message)" "WARN"
            }
        } else {
            Write-Log "  Enhanced Session: Host setting unchanged (not forced off)"
        }

        if ($EnableMetering) {
            try {
                Enable-VMResourceMetering -VMName $VMName -ErrorAction Stop
                Reset-VMResourceMetering -VMName $VMName -ErrorAction SilentlyContinue
                $meterPath = Join-Path $VMLoc "ResourceMetering-Initial.csv"
                Measure-VM -Name $VMName -ErrorAction SilentlyContinue |
                    Select-Object VMName, MeteringDuration, AverageProcessorUsage, AverageMemoryUsage, MaximumMemoryUsage, MinimumMemoryUsage, TotalDiskAllocation |
                    Export-Csv -Path $meterPath -NoTypeInformation
                Write-Log "  Resource Metering: Enabled (initial snapshot: $meterPath)" "OK"
            } catch {
                Write-Log "  Resource Metering setup failed: $($_.Exception.Message)" "WARN"
            }
        }

        # Nested Virtualization
        if ($EnableNestedVirt) {
            try {
                Set-VMProcessor -VMName $VMName -ExposeVirtualizationExtensions $true
                Write-Log "  Nested Virtualization: Enabled" "OK"
                if ($EnableNestedNetFollowup) {
                    try {
                        Set-VMNetworkAdapter -VMName $VMName -MacAddressSpoofing On -ErrorAction Stop
                        Write-Log "  Nested Networking Follow-up: MAC spoofing enabled" "OK"
                    } catch {
                        Write-Log "  Nested Networking follow-up failed: $($_.Exception.Message)" "WARN"
                    }
                }
            } catch {
                Write-Log "  Nested Virtualization failed: $($_.Exception.Message)" "WARN"
            }
        }

        # Attach ISO for recovery boot if bcdboot failed
        if ($attachISOForRecovery -and $ctrlCreate["ISOPath"].Text) {
            try {
                Add-VMDvdDrive -VMName $VMName -Path $ctrlCreate["ISOPath"].Text
                # Set DVD as first boot device
                $dvd = Get-VMDvdDrive -VMName $VMName | Select-Object -First 1
                Set-VMFirmware -VMName $VMName -FirstBootDevice $dvd
                Write-Log "ISO attached as DVD drive for recovery boot. Boot from ISO and select 'Repair your computer'." "WARN"
            } catch {
                Write-Log "Failed to attach ISO for recovery: $($_.Exception.Message)" "ERROR"
            }
        }

        # ---- Start VM ----
        $vmStarted = $false
        if ($StartVM) {
            if (Start-VMWithRetry -VMName $VMName -MaxRetries 2) {
                $vmStarted = $true
                Write-Log "VM started." "OK"
            } elseif ($EnableSecureBoot) {
                Write-Log "VM start failed with Secure Boot enabled. Retrying with alternate Secure Boot template order..." "WARN"
                $altTemplateOrder = @('MicrosoftUEFICertificateAuthority', 'MicrosoftWindows')
                if ($guestProfile -and $guestProfile.SecureBootTemplateOrder -and ($guestProfile.SecureBootTemplateOrder -join '|') -eq ($altTemplateOrder -join '|')) {
                    $altTemplateOrder = @('MicrosoftWindows', 'MicrosoftUEFICertificateAuthority')
                }

                if (Set-VMGuestSecureBoot -VMName $VMName -EnableSecureBoot $true -GuestIsWindows11 $IsWin11 -GuestBuild $guestProfile.Build -TemplateOrder $altTemplateOrder) {
                    if (Start-VMWithRetry -VMName $VMName -MaxRetries 1) {
                        $vmStarted = $true
                        Write-Log "VM started after Secure Boot template fallback." "OK"
                    } else {
                        Write-Log "VM still failed to start after Secure Boot template fallback. Try disabling Secure Boot for this image or review firmware template manually." "WARN"
                    }
                } else {
                    Write-Log "Secure Boot template fallback could not be applied." "WARN"
                }
            }
        }

        # Open vmconnect only when VM starts
        if ($vmStarted) {
            try {
                $vmconnectExe = Join-Path $env:WINDIR 'System32\vmconnect.exe'
                if (-not (Test-Path $vmconnectExe)) { $vmconnectExe = 'vmconnect.exe' }
                & $vmconnectExe localhost $VMName
            } catch {
                Write-Log "vmconnect launch failed for '$VMName': $($_.Exception.Message)" "WARN"
            }

            if ($attachISOForRecovery -and $ResetBootOrder) {
                try {
                    $hdd = Get-VMHardDiskDrive -VMName $VMName | Where-Object { $_.Path -eq $VHDPath } | Select-Object -First 1
                    if (-not $hdd) {
                        $hdd = Get-VMHardDiskDrive -VMName $VMName | Select-Object -First 1
                    }
                    if ($hdd) {
                        Set-VMFirmware -VMName $VMName -FirstBootDevice $hdd
                        Write-Log "Recovery boot reset: boot order switched back to VHD." "OK"
                    }
                } catch {
                    Write-Log "Recovery boot reset failed: $($_.Exception.Message)" "WARN"
                }
            }
        }

        Write-Log "========================================" "OK"
        Write-Log "VM '$VMName' creation completed successfully!" "OK"
        Write-Log "========================================" "OK"
        
        # Summary of VM configuration
        Write-Log "VM Configuration Summary:" "INFO"
        Write-Log "  Name: $VMName" "INFO"
        Write-Log "  Location: $VMLoc" "INFO"
        Write-Log "  Memory: $MemGB GB" "INFO"
        Write-Log "  Disk: $DiskGB GB" "INFO"
        Write-Log "  OS Type: $(if ($UseGoldenImage) { 'Golden Image' } else { $script:DetectedWinVersion })" "INFO"
        Write-Log "  TPM: $(if ($EnableTPM) { 'Enabled' } else { 'Disabled' })" "INFO"
        Write-Log "  Secure Boot: $(if ($EnableSecureBoot) { 'Enabled' } else { 'Disabled' })" "INFO"
        
        Update-CreateProgress -Percent 100 -Status "VM creation completed successfully"

        # Refresh GPU tab VM list
        Update-VMList

        $rollbackNeeded = $false

        } catch {
            # Handle validation errors vs other unexpected errors
            $errorMessage = $_.Exception.Message
            if ($isValidationError) {
                # This is a validation error - show user-friendly message
                $validationError = $errorMessage
                Write-Log $errorMessage "ERROR"
            } else {
                # This is an unexpected error during validation or VM creation
                Update-CreateProgress -Percent 0 -Status "VM creation failed"
                Write-ErrorWithGuidance -Context "Create VM ($VMName)" -ErrorRecord $_
                Write-Log "Stack: $($_.ScriptStackTrace)" "ERROR"
                if ($rollbackNeeded) {
                    Remove-PartialVmArtifacts -VMName $VMName -VMLoc $VMLoc -VHDPath $VHDPath -RemoveVmFolder $vmFolderCreatedByScript
                }
            }
        }
    } finally {
        foreach ($artifact in @($UnattendXMLPath, $cmdFile, $tempExe)) {
            if (-not [string]::IsNullOrWhiteSpace($artifact) -and (Test-Path $artifact)) {
                try {
                    Remove-Item -Path $artifact -Force -ErrorAction Stop
                } catch {
                    Write-Log "Could not remove setup artifact '$artifact': $($_.Exception.Message)" "WARN"
                }
            }
        }
        if ($null -ne $autoPlayGuard) {
            try {
                Restore-AutoPlayState -State $autoPlayGuard
            } catch {
                Write-Log "Could not restore AutoPlay state: $($_.Exception.Message)" "WARN"
            }
        }
        # Dispose SecureString password if still alive (e.g. golden mode or early exception)
        if ($Password) {
            try { 
                $Password.Dispose() 
            } catch { 
                Write-Log "Password disposal warning: $($_.Exception.Message)" "WARN"
            }
            $Password = $null
        }
        $PasswordText = $null
        # Clear password from the UI text box so it doesn't persist in memory
        if ($ctrlCreate.ContainsKey("Password") -and $ctrlCreate["Password"]) {
            $ctrlCreate["Password"].Text = ""
        }
        $script:IsCreating = $false
        $tabControl.Enabled  = $true
        [void](Update-CreateValidationHint)
        
        # Show validation error to user if one occurred
        if ($validationError) {
            [System.Windows.Forms.MessageBox]::Show(
                $validationError,
                "Validation Error", 
                [System.Windows.Forms.MessageBoxButtons]::OK, 
                [System.Windows.Forms.MessageBoxIcon]::Warning)
        }
    }
})

# ================================================================
#  UPDATE GPU - Main Logic
# ================================================================
$btnUpdateGPU.Add_Click({
    # Re-entrancy guard: prevent double-click during GPU update
    if ($script:IsUpdatingGPU) { return }
    $script:IsUpdatingGPU = $true

    $autoPlayGuard = $null

    try {
        # Gather selections from single source of truth
        $selectedVMs = @()
        if ($script:GpuSelectedVMs) {
            $selectedVMs += @(
                $script:GpuSelectedVMs.GetEnumerator() |
                    Where-Object { $_.Value } |
                    ForEach-Object { $_.Key }
            )
        }
        $selectedVMs = @($selectedVMs | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Sort-Object -Unique)
        if ($selectedVMs.Count -eq 0) { Write-Log "No VMs selected!" "ERROR"; return }

        $providerObj = $null
        $gpuVendor   = "Auto"
        if ($script:GpuPList.Count -gt 0) {
            $providerObj = $script:GpuPList | Where-Object { $_.Friendly -eq $ctrlGPU["GpuSelector"].SelectedItem } | Select-Object -First 1
            if ($providerObj -and $providerObj.Provider) {
                $gpuVendor = switch -Wildcard ($providerObj.Friendly) {
                    "*NVIDIA*" { "NVIDIA" }
                    "*AMD*"    { "AMD" }
                    "*Intel*"  { "Intel" }
                    default    { "Auto" }
                }
            }
        }

        $smartCopy = $ctrlGPU["SmartCopy"].Checked
        $autoExpand = $ctrlGPU["AutoExpand"].Checked
        $copySvcDriver = $ctrlGPU["CopySvcDriver"].Checked
        $gpuAllocPercent = [int]$ctrlGPU["GpuAllocSlider"].Value
        $startAfterUpdate = $ctrlGPU["StartVM"].Checked
        $strictChecks = if ($ctrlGPU.ContainsKey("StrictChecks") -and $ctrlGPU["StrictChecks"]) { [bool]$ctrlGPU["StrictChecks"].Checked } else { $true }
        $removeOnlyMode = ($providerObj -and $null -eq $providerObj.Provider)

        if ($removeOnlyMode) {
            $confirmRemove = [System.Windows.Forms.MessageBox]::Show(
                "You selected 'NONE - Remove GPU Adapter'.`n`nThis will remove the GPU-P adapter from $($selectedVMs.Count) VM(s).`n`nContinue?",
                "Confirm GPU Adapter Removal",
                [System.Windows.Forms.MessageBoxButtons]::YesNo,
                [System.Windows.Forms.MessageBoxIcon]::Warning)
            if ($confirmRemove -ne [System.Windows.Forms.DialogResult]::Yes) {
                Write-Log "GPU adapter removal cancelled by user." "INFO"
                return
            }
        }

        if (-not $removeOnlyMode) {
            $preflightIssues = @(Test-GpuPPreFlight)
            foreach ($issue in $preflightIssues) {
                if ($issue -match '^ERROR:') { Write-Log $issue "ERROR" }
                elseif ($issue -match '^WARNING:') { Write-Log $issue "WARN" }
                else { Write-Log $issue "INFO" }
            }

            $preferredHostGpuName = if ($providerObj -and $providerObj.Provider) { [string]$providerObj.Name } else { "" }
            $partitionCountFix = Ensure-HostGpuPartitionCountValid -PreferredGpuName $preferredHostGpuName
            if (-not $partitionCountFix.Success) {
                if ($strictChecks) {
                    Write-Log "GPU host partition-count validation failed: $($partitionCountFix.Message). Aborting due to strict checks." "ERROR"
                    return
                }
                Write-Log "GPU host partition-count validation warning: $($partitionCountFix.Message)" "WARN"
            } elseif ($partitionCountFix.Changed) {
                Write-Log "GPU host partition-count normalized: $($partitionCountFix.Message)" "WARN"
            } else {
                Write-Log "GPU host partition-count check: $($partitionCountFix.Message)" "INFO"
            }

            $requireSriov = ($script:HostOsName -match 'Windows Server')
            $hostReadiness = Test-GpuPHostReadiness -RequireSriov:$requireSriov
            foreach ($warn in $hostReadiness.Warnings) { Write-Log "GPU host readiness: $warn" "WARN" }
            if (-not $hostReadiness.CanProceed) {
                foreach ($err in $hostReadiness.Errors) { Write-Log "GPU host readiness: $err" "ERROR" }
                if ($strictChecks) {
                    Write-Log "Strict GPU safety checks are enabled. Aborting GPU update." "ERROR"
                    return
                }
                Write-Log "Continuing despite host readiness errors because strict checks are disabled." "WARN"
            }
        }

        $hostNvidiaVersion = Get-HostNvidiaDriverVersion
        $hostNvidiaBranch = Get-DriverVersionBranch -VersionString $hostNvidiaVersion
        if ($hostNvidiaVersion) {
            Write-Log "Host NVIDIA driver version detected: $hostNvidiaVersion (branch $hostNvidiaBranch)" "INFO"
        }

        if ($script:CliWhatIf) {
            Write-Log "WhatIf mode: GPU update preflight complete. No VM GPU/driver changes will be made." "WARN"
            return
        }

        # Disable UI
        $tabControl.Enabled  = $false
        $btnUpdateGPU.Enabled = $false

        Write-Log "========================================" "INFO"
        Write-Log "Starting GPU driver update for $($selectedVMs.Count) VM(s)" "INFO"
        Write-Log "GPU Vendor: $gpuVendor | Smart Copy: $smartCopy | AutoExpand: $autoExpand | Allocation: $gpuAllocPercent%" "INFO"
        Write-Log "========================================" "INFO"

        if ($ctrlGPU.ContainsKey("GpuStatus") -and $ctrlGPU["GpuStatus"]) {
            $ctrlGPU["GpuStatus"].Text = "Running GPU update for $($selectedVMs.Count) VM(s)..."
            $ctrlGPU["GpuStatus"].ForeColor = $theme.Info
            $ctrlGPU["GpuStatus"].Refresh()
        }

        # Disable AutoPlay
        $autoPlayGuard = Disable-AutoPlayGuarded

        foreach ($VMName in $selectedVMs) {
            $vhdPath = $null
            $mountLetter = $null
            $mountedByScript = $false
            $skipDriverInjection = $false
            $skipHostDriverInjection = $false
            $gpuAdapterConfigured = $false
            $keepGpuAdapter = $false
            $startVmPending = $false

            try {
                $vm = Get-VM -Name $VMName -ErrorAction Stop
                Write-Log "Processing VM: $VMName"
                if ($ctrlGPU.ContainsKey("GpuStatus") -and $ctrlGPU["GpuStatus"]) {
                    $ctrlGPU["GpuStatus"].Text = "[$VMName] Processing..."
                    $ctrlGPU["GpuStatus"].ForeColor = $theme.Info
                    $ctrlGPU["GpuStatus"].Refresh()
                }

                if (-not $removeOnlyMode -and [int]$vm.Generation -ne 2) {
                    Write-Log "[$VMName] Generation $($vm.Generation) VM detected. GPU-P requires Generation 2 VM. Skipping." "ERROR"
                    continue
                }

                # Shutdown VM if not already Off (handles Running, Saved, Paused, etc.)
                if ($vm.State -ne 'Off') {
                    Write-Log "[$VMName] VM state is '$($vm.State)'. Shutting down..."
                    if (-not (Stop-VMWithTimeout -VMName $VMName -TimeoutSec 60)) {
                        Write-Log "[$VMName] VM did not stop cleanly. Skipping to avoid corruption." "ERROR"
                        continue
                    }
                    Write-Log "[$VMName] VM stopped." "OK"
                }

                if (-not $removeOnlyMode) {
                    # Nested virtualisation is incompatible with GPU-P for ALL GPU vendors.
                    # Set-VMProcessor requires the VM to be Off, so this must run after shutdown.
                    try {
                        $vmProcessor = Get-VMProcessor -VMName $VMName -ErrorAction SilentlyContinue
                        if ($vmProcessor -and $vmProcessor.ExposeVirtualizationExtensions) {
                            Write-Log "[$VMName] Nested virtualisation is enabled. Disabling ExposeVirtualizationExtensions for GPU-P stability." "WARN"
                            Set-VMProcessor -VMName $VMName -ExposeVirtualizationExtensions $false -ErrorAction Stop
                        }
                    } catch {
                        Write-Log "[$VMName] Could not disable nested virtualisation: $($_.Exception.Message)" "WARN"
                    }

                    # Automatic checkpoints are incompatible with GPU-P: Hyper-V cannot save
                    # GPU partition state into a checkpoint file, causing the VM to stall or
                    # corrupt when a background checkpoint is attempted while GPU-P is active.
                    try {
                        Set-VM -VMName $VMName -AutomaticCheckpointsEnabled $false -ErrorAction Stop
                        Write-Log "[$VMName] Automatic checkpoints disabled for GPU-P compatibility." "INFO"
                    } catch {
                        Write-Log "[$VMName] Could not disable automatic checkpoints: $($_.Exception.Message)" "WARN"
                    }
                }

                # Remove existing GPU-P adapter
                Start-Sleep -Seconds 2
                $removeAttempt = 0
                do {
                    Remove-VMGpuPartitionAdapter -VMName $VMName -Confirm:$false -ErrorAction SilentlyContinue
                    Start-Sleep -Seconds 1
                    $remainingAdapters = @(Get-VMGpuPartitionAdapter -VMName $VMName -ErrorAction SilentlyContinue)
                    if ($remainingAdapters.Count -eq 0) { break }
                    $removeAttempt++
                } while ($removeAttempt -lt 3)

                if ($remainingAdapters.Count -gt 0) {
                    Write-Log "[$VMName] Could not fully remove existing GPU-P adapter/allocation entries after retries. Skipping for clean install safety." "ERROR"
                    continue
                }

                # Add GPU-P adapter (unless "Remove" was selected)
                if ($providerObj -and $null -ne $providerObj.Provider) {
                    try {
                        if ($script:SupportsGpuInstancePath) {
                            Write-Log "[$VMName] Adding GPU-P adapter: $($providerObj.Friendly)"
                            try {
                                Add-VMGpuPartitionAdapter -VMName $VMName -InstancePath $providerObj.Name -ErrorAction Stop
                            } catch {
                                Write-Log "[$VMName] Specific GPU path failed, falling back to default GPU adapter." "WARN"
                                Add-VMGpuPartitionAdapter -VMName $VMName -ErrorAction Stop
                            }
                        } else {
                            Write-Log "[$VMName] Host does not support specific GPU selection, adding default GPU-P adapter"
                            Add-VMGpuPartitionAdapter -VMName $VMName -ErrorAction Stop
                        }

                        Set-GpuPartitionForVM -VMName $VMName -AllocationPercent $gpuAllocPercent -ConservativeProfile $true
                        Write-Log "[$VMName] GPU-P adapter configured." "OK"
                        $gpuAdapterConfigured = $true
                    } catch {
                        Write-Log "[$VMName] GPU-P adapter error: $($_.Exception.Message)" "ERROR"
                        continue
                    }
                } elseif ($providerObj -and $null -eq $providerObj.Provider) {
                    Write-Log "[$VMName] GPU-P adapter removed only (NONE selected)." "OK"
                    $skipDriverInjection = $true
                } else {
                    # No provider obj - try default for Win10
                    try {
                        Write-Log "[$VMName] Adding GPU-P adapter (default selection)"
                        Add-VMGpuPartitionAdapter -VMName $VMName -ErrorAction Stop
                        Set-GpuPartitionForVM -VMName $VMName -AllocationPercent $gpuAllocPercent -ConservativeProfile $true
                        Write-Log "[$VMName] GPU-P adapter configured." "OK"
                        $gpuAdapterConfigured = $true
                    } catch {
                        Write-Log "[$VMName] GPU-P adapter error: $($_.Exception.Message)" "ERROR"
                        continue
                    }
                }

                if ($skipDriverInjection) {
                    if ($startAfterUpdate) {
                        $startVmPending = $true
                        Write-Log "[$VMName] Remove-only mode: VM start is scheduled by option." "INFO"
                    }
                    Write-Log "[$VMName] Driver injection skipped (remove-only mode)." "INFO"
                    continue
                }

                if (-not $gpuAdapterConfigured) {
                    Write-Log "[$VMName] GPU-P adapter was not configured successfully. Skipping driver injection." "ERROR"
                    continue
                }

                # Get VHD path
                $vhdPath = Get-VMPrimaryVhdPath -VMName $VMName
                if (-not $vhdPath -or -not (Test-Path $vhdPath)) {
                    Write-Log "[$VMName] No primary VHD path found." "ERROR"
                    continue
                }
                Write-Log "[$VMName] Found VHDX: $vhdPath"

                # Mount VHD
                if (-not (Mount-VhdWithFallback -ImagePath $vhdPath)) {
                    Write-Log "[$VMName] Error mounting VHD. Skipping this VM." "ERROR"
                    continue
                }
                $mountedByScript = $true

                # Wait for NTFS volume
                $disk = $null
                for ($d = 0; $d -lt 10; $d++) {
                    try {
                        $disk = Get-DiskImage -ImagePath $vhdPath -ErrorAction Stop | Get-Disk -ErrorAction Stop
                        if ($disk) { break }
                    } catch {
                        if ($d -ge 9) {
                            Write-Log "[$VMName] Unable to resolve mounted disk object after retries: $($_.Exception.Message)" "WARN"
                        }
                    }
                    Start-Sleep -Seconds 1
                }
                if (-not $disk) {
                    Write-Log "[$VMName] Could not resolve mounted disk object." "ERROR"
                    continue
                }

                $maxWait = 20
                $waited = 0
                while (-not $mountLetter -and $waited -lt $maxWait) {
                    Start-Sleep -Seconds 1
                    $partitions = Get-DataPartitions -DiskNumber $disk.Number
                    foreach ($part in $partitions) {
                        $vol = Get-Volume -Partition $part -ErrorAction SilentlyContinue
                        if ($vol -and $vol.FileSystem -eq 'NTFS' -and $vol.DriveLetter) {
                            $mountLetter = "$($vol.DriveLetter):"
                            break
                        }
                    }
                    $waited++
                }

                if (-not $mountLetter) {
                    Write-Log "[$VMName] Could not find NTFS volume on VHD." "ERROR"
                    continue
                }
                Write-Log "[$VMName] Mounted at $mountLetter"

                $guestBuild = Get-GuestWindowsBuildFromMountedVolume -MountLetter $mountLetter
                if ($guestBuild -gt 0 -and $script:HostBuild -gt 0) {
                    Write-Log "[$VMName] Build check: host=$($script:HostBuild), guest=$guestBuild" "INFO"
                    $hostIsWin11 = ($script:HostBuild -ge $script:BUILD_WIN11_MIN)
                    $guestIsWin11 = ($guestBuild -ge $script:BUILD_WIN11_MIN)
                    if ($hostIsWin11 -ne $guestIsWin11) {
                        $msg = "[$VMName] Host/guest major OS generation mismatch detected (Win10 vs Win11). This is known to destabilize GPU-P driver injection."
                        if ($strictChecks) {
                            Write-Log "$msg Strict checks enabled, skipping VM." "ERROR"
                            continue
                        }
                        Write-Log "$msg Strict checks disabled: skipping host GPU file injection for boot stability." "WARN"
                        $skipHostDriverInjection = $true
                    }

                    # Windows 11 24H2 / 25H2 (build >= 26100) contains a known dxgkrnl.sys
                    # bug where GPU-P causes the boot animation to freeze and the Hyper-V
                    # console to hang until a cumulative Windows Update (including Preview /
                    # optional updates) is installed inside the guest VM. This is the same
                    # issue documented in bryanem32/hyperv_vm_creator README (v21 changelog).
                    # GPU drivers will still be injected, but the guest MUST run Windows
                    # Update (including previews) before GPU-P will function without freezing.
                    if ($guestBuild -ge 26100) {
                        Write-Log "[$VMName] WARNING: Guest is Windows 11 24H2/25H2 (build $guestBuild). A known Microsoft bug in dxgkrnl.sys causes GPU-P to freeze the boot animation on this build. FIX: Boot the guest WITHOUT GPU-P first, run Windows Update (Settings > Windows Update > Advanced > Optional Updates — install ALL), reboot the guest, then re-run GPU Setup." "WARN"
                    }
                }

                # Apply GPU-P registry mitigations to the guest SYSTEM hive while the VHD
                # is mounted. This disables HyperVideo (synthetic video / GPU-P conflict)
                # and on 24H2/25H2 also patches the TDR timeout and BasicDisplay race.
                Set-GuestGpuRegistryMitigations -VMName $VMName -MountLetter $mountLetter -GuestBuild $guestBuild

                if (-not $removeOnlyMode -and $gpuVendor -eq 'NVIDIA' -and $hostNvidiaVersion) {
                    $guestNvidiaVersion = Get-NvidiaDriverVersionFromGuestStore -MountLetter $mountLetter
                    if ($guestNvidiaVersion) {
                        $guestNvidiaBranch = Get-DriverVersionBranch -VersionString $guestNvidiaVersion
                        Write-Log "[$VMName] Guest NVIDIA driver version detected: $guestNvidiaVersion (branch $guestNvidiaBranch)" "INFO"
                        if ($guestNvidiaBranch -ge 0 -and $hostNvidiaBranch -ge 0 -and $guestNvidiaBranch -ne $hostNvidiaBranch) {
                            Write-Log "[$VMName] NVIDIA host/guest driver branch mismatch detected (host=$hostNvidiaBranch, guest=$guestNvidiaBranch). Proceeding with clean guest GPU payload purge and host-branch reinjection." "WARN"
                        }
                    } else {
                        Write-Log "[$VMName] No existing NVIDIA guest driver metadata found in HostDriverStore (fresh or non-NVIDIA guest state)." "INFO"
                    }
                }

                if ($skipHostDriverInjection) {
                    if ($startAfterUpdate) {
                        $startVmPending = $true
                    }
                    Write-Log "[$VMName] Host GPU file injection skipped to reduce boot-freeze risk. Keep guest and host GPU drivers on matching OS generation/branch, then rerun update." "WARN"
                    continue
                }

                $cleanupResult = Remove-GuestGpuDriverPayload -VMName $VMName -MountLetter $mountLetter -GpuVendor $gpuVendor
                if (-not $cleanupResult.Success) {
                    Write-Log "[$VMName] Guest GPU payload cleanup failed; skipping VM to preserve clean-install guarantees." "ERROR"
                    continue
                }

                # ---- Copy GPU drivers (smart or full) ----
                $copyResult = Copy-GpuDriverFolders -VMName $VMName -MountLetter $mountLetter `
                    -VhdPath $vhdPath -GpuVendor $gpuVendor -SmartCopy $smartCopy -AutoExpand $autoExpand

                # Update mount letter in case VHD auto-expand changed the drive letter
                if ($copyResult.MountLetter) { $mountLetter = $copyResult.MountLetter }

                if (-not $copyResult.Success) {
                    Write-Log "[$VMName] GPU driver copy failed." "ERROR"
                    continue
                }

                # Copy vendor-specific kernel-mode GPU-P interface files from host System32.
                # These files (nv*, amdkmd*, igfx*) are required for the guest dxgkrnl.sys
                # to attach to the GPU partition during early boot. Without them the guest
                # GPU driver stack is incomplete and the VM freezes at the boot animation.
                # Reference: bryanem32/hyperv_vm_creator v29 copies these for all vendors.
                $sys32Mask = switch -Wildcard ($gpuVendor) {
                    'NVIDIA' { 'nv*' }
                    'AMD'    { 'amdkmd*' }
                    'Intel'  { 'igfx*' }
                    default  { $null }
                }
                if ($sys32Mask) {
                    Write-Log "[$VMName] Copying vendor GPU-P kernel files ($sys32Mask) from host System32 to guest System32..." "INFO"
                    Copy-DriversToVhd -VMName $VMName -MountLetter $mountLetter `
                        -Source "$env:SystemRoot\System32" -Destination 'Windows\System32' -FileMask $sys32Mask | Out-Null
                }

                # Copy precise file set referenced by active display driver package(s)
                $gpuCopyName = if ($providerObj -and $providerObj.Provider) { $providerObj.Friendly } else { "AUTO" }
                $refCopy = Copy-GpuReferencedFiles -MountLetter $mountLetter -GPUName $gpuCopyName
                if (-not $refCopy.Success) {
                    Write-Log "[$VMName] Referenced GPU file copy step reported failure; proceeding with HostDriverStore/service-driver-only path." "WARN"
                }
                Save-GuestGpuInjectedFilesManifest -VMName $VMName -MountLetter $mountLetter -RelativePaths $refCopy.Paths

                if ($copySvcDriver) {
                    Copy-GpuServiceDriver -MountLetter $mountLetter -GPUName $(if ($providerObj -and $providerObj.Provider) { $providerObj.Friendly } else { "AUTO" })
                }

                Write-Log "[$VMName] GPU drivers injected." "OK"
                $keepGpuAdapter = $true

                # Defer VM start until after VHD dismount in finally
                if ($startAfterUpdate) {
                    $startVmPending = $true
                }

                Write-Log "[$VMName] Done." "OK"
                if ($ctrlGPU.ContainsKey("GpuStatus") -and $ctrlGPU["GpuStatus"]) {
                    $ctrlGPU["GpuStatus"].Text = "[$VMName] Complete."
                    $ctrlGPU["GpuStatus"].ForeColor = $theme.Success
                    $ctrlGPU["GpuStatus"].Refresh()
                }

            } catch {
                Write-ErrorWithGuidance -Context "GPU update [$VMName]" -ErrorRecord $_
            } finally {
                $canStartVm = $true
                if ($mountedByScript -and $vhdPath) {
                    if (Dismount-ImageRetry -ImagePath $vhdPath) {
                        Write-Log "[$VMName] VHD dismounted." "OK"
                    } else {
                        try {
                            Dismount-VHD -Path $vhdPath -ErrorAction SilentlyContinue
                            Unregister-TrackedMountedImage -ImagePath $vhdPath
                        } catch {
                            Write-Log "[$VMName] VHD dismount fallback failed: $($_.Exception.Message)" "WARN"
                        }
                        Write-Log "[$VMName] VHD dismount required fallback and may still be attached." "WARN"
                        $canStartVm = $false
                    }

                    if (-not (Wait-ImageDetached -ImagePath $vhdPath -TimeoutSec 20)) {
                        Write-Log "[$VMName] VHD still appears attached after dismount wait; skipping auto-start to avoid file-lock failure." "WARN"
                        $canStartVm = $false
                    }
                }

                if ($gpuAdapterConfigured -and -not $keepGpuAdapter) {
                    try {
                        Remove-VMGpuPartitionAdapter -VMName $VMName -Confirm:$false -ErrorAction SilentlyContinue
                        Start-Sleep -Seconds 1
                        $leftoverAdapter = @(Get-VMGpuPartitionAdapter -VMName $VMName -ErrorAction SilentlyContinue)
                        if ($leftoverAdapter.Count -eq 0) {
                            Write-Log "[$VMName] Rolled back GPU-P adapter due to incomplete/failed GPU update path." "WARN"
                        } else {
                            Write-Log "[$VMName] GPU-P rollback attempted but adapter still present." "WARN"
                        }
                    } catch {
                        Write-Log "[$VMName] GPU-P rollback failed: $($_.Exception.Message)" "WARN"
                    }
                }

                if ($startVmPending) {
                    if (-not $canStartVm) {
                        Write-Log "[$VMName] Start skipped because VHD could not be cleanly released." "WARN"
                    } elseif (Start-VMWithRetry -VMName $VMName -MaxRetries 2) {
                        Write-Log "[$VMName] VM started." "OK"
                        # Open vmconnect if not already open
                        $escapedVMName = [regex]::Escape($VMName)
                        $existing = Get-CimInstance Win32_Process -Filter "Name = 'vmconnect.exe'" -ErrorAction SilentlyContinue |
                            Where-Object { $_.CommandLine -match "(\s|`")${escapedVMName}(`"|\s|$)" }
                        if (-not $existing) { vmconnect.exe localhost $VMName }
                    }
                }
            }
        }

        Write-Log "========================================" "OK"
        Write-Log "GPU driver update complete." "OK"
        Write-Log "========================================" "OK"

    } catch {
        Write-ErrorWithGuidance -Context "GPU update" -ErrorRecord $_
    } finally {
        if ($null -ne $autoPlayGuard) {
            try {
                Restore-AutoPlayState -State $autoPlayGuard
            } catch {
                Write-Log "Could not restore AutoPlay state after GPU update: $($_.Exception.Message)" "WARN"
            }
        }
        $script:IsUpdatingGPU = $false
        $tabControl.Enabled   = $true
        if ($ctrlGPU.ContainsKey("GpuStatus") -and $ctrlGPU["GpuStatus"]) {
            $ctrlGPU["GpuStatus"].Text = ""
        }
        Update-GpuActionState
    }
})

# ---- Form Closing Cleanup ----
$form.Add_FormClosing({
    param($eventSource, $e)
    [void]$eventSource
    # Warn if a long-running operation is still in progress
    if ($script:IsCreating -or $script:IsUpdatingGPU) {
        $closeConfirm = [System.Windows.Forms.MessageBox]::Show(
            "An operation is still in progress. Are you sure you want to close?`n`nClosing now may leave resources in an inconsistent state.",
            "Operation In Progress",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning)
        if ($closeConfirm -ne [System.Windows.Forms.DialogResult]::Yes) {
            $e.Cancel = $true
            return
        }
    }
    Invoke-MountCleanup
    # Dispose timers
    if ($script:ValidationTimer)  { $script:ValidationTimer.Stop();  $script:ValidationTimer.Dispose() }
    if ($script:VmFilterTimer)    { $script:VmFilterTimer.Stop();    $script:VmFilterTimer.Dispose() }
    if ($script:LayoutTimer)      { $script:LayoutTimer.Stop();      $script:LayoutTimer.Dispose() }
    # Dispose cached brushes
    if ($script:TabBrushSelected) { $script:TabBrushSelected.Dispose() }
    if ($script:TabBrushNormal)   { $script:TabBrushNormal.Dispose() }
    # Dispose tooltip
    if ($toolTip)                 { $toolTip.Dispose() }
    # Dispose cached fonts (moved here so they survive any final paint events)
    foreach ($f in @($script:FontMain, $script:FontTabHeader, $script:FontHeader, $script:FontBoldButton,
                     $script:FontSmall, $script:FontBoldLabel, $script:FontConsolas, $script:FontSidebarNav,
                     $script:FontAppTitle, $script:ThemeFontGroupBox, $script:FontSidebarBrand)) {
        if ($f) {
            try {
                $f.Dispose()
            } catch {
                Write-StartupTrace -Message "Font dispose warning during shutdown: $($_.Exception.Message)" -Level 'WARN'
            }
        }
    }
})

#endregion

#region ==================== MAIN ====================

# Log header
Write-Log "Hyper-V Toolkit $($script:ToolkitVersion) by $($script:ToolkitCreator) - $($script:ToolkitTagline)" "OK"
Write-Log "Host OS: $($script:HostOsName)" "INFO"
Write-Log "GPU-P Specific GPU Selection: $(if ($script:SupportsGpuInstancePath) {'Available'} else {'Not available on this host build'})" "INFO"
if (-not [string]::IsNullOrWhiteSpace($script:CliVMName)) {
    $ctrlCreate["VMName"].Text = $script:CliVMName.Trim()
    Write-Log "CLI prefill: VM Name set from parameter." "INFO"
}
if (-not [string]::IsNullOrWhiteSpace($script:CliISOPath)) {
    $prefillIsoPath = $script:CliISOPath.Trim()
    $ctrlCreate["ISOPath"].Text = $prefillIsoPath
    if (Test-PathCached $prefillIsoPath) {
        Write-Log "CLI prefill: ISO path set from parameter. Use Browse ISO to load editions and mount source." "INFO"
    } else {
        Write-Log "CLI prefill warning: provided ISO path does not exist: $prefillIsoPath" "WARN"
    }
}
if ($script:CliWhatIf) {
    Write-Log "WhatIf mode is enabled: creation/GPU actions will run preflight only and skip host/guest changes." "WARN"
}
Write-Log "Ready." "OK"

Write-Host "  Startup complete." -ForegroundColor Green
Write-Host ""

try {
    Update-MainLayout -RootForm $form
} catch {
    Write-UiWarning "Initial layout warning: $($_.Exception.Message)"
}

Update-StatusBar -Message "Ready to create VMs"

try {
    [void]$form.ShowDialog()
} catch {
    Write-StartupTrace -Message "GUI ShowDialog fatal error: $($_.Exception.Message)" -Level 'ERROR'
    Write-Host "`n  FATAL: GUI error: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "  $($_.ScriptStackTrace)" -ForegroundColor DarkRed
    Write-Host "`n  Press Enter to exit..." -ForegroundColor Red
    Read-Host
} finally {
    if ($form) { $form.Dispose() }
}

exit 0

#endregion
