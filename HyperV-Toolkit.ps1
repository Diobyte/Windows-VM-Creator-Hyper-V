#region ==================== INITIALIZATION ====================

$script:ToolkitVersion = "Version 1"
$script:ToolkitCreator = "Diobyte"
$script:ToolkitTagline = "Made with love"

# Execution Policy
$currentPolicy = Get-ExecutionPolicy -Scope Process
if (-not $currentPolicy -or $currentPolicy -in @('Undefined', 'Restricted')) {
    try {
        Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force -ErrorAction Stop
    } catch {
        Write-Output "Could not set process execution policy to Bypass. Continuing with current policy: $currentPolicy"
    }
}

# Load assemblies
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# 64-bit process check
if (-not [Environment]::Is64BitProcess) {
    [System.Windows.MessageBox]::Show(
        "This tool requires 64-bit PowerShell.`n`nDo not use PowerShell (x86). Use the standard PowerShell or the Launch.bat file.",
        "64-bit Required", "OK", "Error"
    ) | Out-Null
    exit 1
}

# Admin check
$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    [System.Windows.MessageBox]::Show(
        "This tool must be run as Administrator.`n`nPlease right-click and select 'Run as Administrator', or use the Launch.bat file.",
        "Administrator Required", "OK", "Warning"
    ) | Out-Null
    exit 1
}

# Hyper-V check
function Test-HyperVRunning {
    try {
        $svc = Get-Service -Name vmms -ErrorAction Stop
        return ($svc.Status -eq 'Running')
    } catch { return $false }
}

$feature = Get-WindowsOptionalFeature -Online -FeatureName Microsoft-Hyper-V-All -ErrorAction SilentlyContinue
if (-not ($feature.State -eq "Enabled" -and (Test-HyperVRunning))) {
    $installChoice = [System.Windows.MessageBox]::Show(
        "Hyper-V is not fully enabled or the hypervisor is not running.`n`nA system restart will be required after installation.`n`nDo you want to enable it now?",
        "Enable Hyper-V", "OKCancel", "Warning"
    )
    if ($installChoice -eq "OK") {
        try {
            Enable-WindowsOptionalFeature -Online -FeatureName Microsoft-Hyper-V -All -NoRestart -ErrorAction Stop *> $null
            bcdedit /set hypervisorlaunchtype auto *> $null
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
}

# Detect host OS
$script:HostOsName = (Get-CimInstance Win32_OperatingSystem).Caption
$script:HostBuild  = [int](Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion').CurrentBuild
$script:HostIsWin11 = ($script:HostBuild -ge 22000)
$script:HostIsWin11Pro = $script:HostOsName -match 'Windows 11.*(Pro|Enterprise|Education)'
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
$script:LogBox              = $null   # Set when GUI is built

#endregion

#region ==================== UTILITY FUNCTIONS ====================

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
    $color = switch ($Level) {
        "ERROR" { [System.Drawing.Color]::Tomato }
        "WARN"  { [System.Drawing.Color]::Gold }
        "OK"    { [System.Drawing.Color]::LimeGreen }
        default { [System.Drawing.Color]::White }
    }
    $script:LogBox.SelectionStart  = $script:LogBox.TextLength
    $script:LogBox.SelectionLength = 0
    $script:LogBox.SelectionColor  = $color
    $script:LogBox.AppendText("$timestamp [$Level] $Message`r`n")
    $script:LogBox.ScrollToCaret()
    [System.Windows.Forms.Application]::DoEvents()
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

        # Dismount ISO if still mounted
        if ($script:MountedISO -and $script:MountedISO.ImagePath) {
            try {
                if (Dismount-ImageRetry -ImagePath $script:MountedISO.ImagePath -MaxRetries 2) {
                    Write-Log "Cleanup dismounted mounted ISO." "OK"
                }
            } catch {
                Write-Log "Cleanup ISO dismount error: $($_.Exception.Message)" "WARN"
            } finally {
                $script:MountedISO = $null
            }
        }
        Write-Log "Cleanup complete." "OK"
    } catch {
        Write-Log "Cleanup error: $($_.Exception.Message)" "WARN"
    }
}

function Set-AutoPlay {
    param([bool]$Disable)
    $regPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\AutoplayHandlers"
    $regName = "DisableAutoplay"
    try {
        $current = (Get-ItemProperty -Path $regPath -Name $regName -ErrorAction Stop).$regName
    } catch { $current = 0 }

    if ($Disable -and $current -eq 0) {
        Set-ItemProperty -Path $regPath -Name $regName -Value 1
        return 0  # was enabled
    } elseif (-not $Disable -and $current -eq 1) {
        # Only restore if we disabled it
    }
    return $current
}

function Dismount-ImageRetry {
    <#
    .SYNOPSIS
        Dismounts a disk image with retry logic to handle in-use locks.
    #>
    param([string]$ImagePath, [int]$MaxRetries = 5)
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
}

function Get-PathAvailableSpaceGB {
    param([string]$Path)
    try {
        $resolvedPath = (Resolve-Path -Path $Path -ErrorAction Stop).Path
        $root = [System.IO.Path]::GetPathRoot($resolvedPath)
        if (-not $root) { return -1 }
        $driveName = $root.TrimEnd('\\').TrimEnd(':')
        $drive = Get-PSDrive -Name $driveName -ErrorAction Stop
        return [math]::Round(($drive.Free / 1GB), 2)
    } catch {
        return -1
    }
}

function Test-DirectoryWritable {
    param([string]$Path)
    try {
        if (-not (Test-Path -Path $Path)) { return $false }
        $testFile = Join-Path $Path ([System.Guid]::NewGuid().ToString() + '.tmp')
        Set-Content -Path $testFile -Value 'test' -Encoding ASCII -ErrorAction Stop
        Remove-Item -Path $testFile -Force -ErrorAction Stop
        return $true
    } catch {
        return $false
    }
}

function Ensure-ToolkitNatSwitch {
    param(
        [string]$SwitchName = "HyperV-Toolkit-NAT",
        [string]$GatewayIp = "192.168.250.1",
        [int]$PrefixLength = 24,
        [string]$NatPrefix = "192.168.250.0/24"
    )

    try {
        $existingSwitch = Get-VMSwitch -Name $SwitchName -ErrorAction SilentlyContinue
        if (-not $existingSwitch) {
            Write-Log "No usable virtual switch selected. Creating internal NAT switch '$SwitchName'..." "WARN"
            New-VMSwitch -Name $SwitchName -SwitchType Internal -ErrorAction Stop | Out-Null
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
                try {
                    New-NetIPAddress -InterfaceAlias $adapterAlias -IPAddress $GatewayIp -PrefixLength $PrefixLength -AddressFamily IPv4 -ErrorAction Stop | Out-Null
                    Write-Log "Assigned $GatewayIp/$PrefixLength to $adapterAlias" "OK"
                } catch {
                    Write-Log "IP assignment skipped or failed for ${adapterAlias}: $($_.Exception.Message)" "WARN"
                }
            }
        }

        $natName = "$SwitchName-NAT"
        $existingNat = Get-NetNat -Name $natName -ErrorAction SilentlyContinue
        if (-not $existingNat) {
            New-NetNat -Name $natName -InternalIPInterfaceAddressPrefix $NatPrefix -ErrorAction Stop | Out-Null
            Write-Log "Created NAT object '$natName' with prefix $NatPrefix" "OK"
        }

        return $SwitchName
    } catch {
        Write-Log "Auto-create switch/NAT failed: $($_.Exception.Message)" "ERROR"
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
    }
    if ($ctrlCreate -and $ctrlCreate.ContainsKey("CreateStatus")) {
        $ctrlCreate["CreateStatus"].Text = $Status
    }
    [System.Windows.Forms.Application]::DoEvents()
}

function Remove-PartialVmArtifacts {
    param(
        [string]$VMName,
        [string]$VMLoc,
        [string]$VHDPath
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
            $img = Get-DiskImage -ImagePath $VHDPath -ErrorAction SilentlyContinue
            if ($img -and $img.Attached) {
                Dismount-VHD -Path $VHDPath -ErrorAction SilentlyContinue
            }
        }
    } catch {
        Write-Log "Rollback warning (VHD dismount): $($_.Exception.Message)" "WARN"
    }

    try {
        if ($VMLoc -and (Test-Path $VMLoc)) {
            Remove-Item -Path $VMLoc -Recurse -Force -ErrorAction SilentlyContinue
            Write-Log "Removed partially created VM folder: $VMLoc" "WARN"
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
        Set-VMFirmware -VMName $VMName -EnableSecureBoot Off
        Write-Log "  Disabling Secure Boot (Windows 10 compatibility mode)" "OK"
        return $true
    }

    Set-VMFirmware -VMName $VMName -EnableSecureBoot On -ErrorAction SilentlyContinue | Out-Null
    $templates = $TemplateOrder
    if (-not $templates -or $templates.Count -eq 0) {
        $templates = if ($GuestIsWindows11) {
            @('MicrosoftUEFICertificateAuthority', 'MicrosoftWindows')
        } elseif ($GuestBuild -gt 0 -and $GuestBuild -lt 17134) {
            @('MicrosoftUEFICertificateAuthority', 'MicrosoftWindows')
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

function Escape-XmlValue {
    param([AllowNull()][string]$Value)
    if ($null -eq $Value) { return "" }
    return [System.Security.SecurityElement]::Escape($Value)
}

function Test-GpuPPreFlight {
    <#
    .SYNOPSIS
        Runs pre-flight checks for GPU-P based on Diobyte Version 1 guidance.
        Returns array of warning/error messages.
    #>
    $issues = @()

    # Check 1: Laptop NVIDIA GPU detection (unsupported for GPU-P)
    try {
        $battery = Get-CimInstance -ClassName Win32_Battery -ErrorAction SilentlyContinue
        if ($battery) {
            $nvidiaGpu = Get-CimInstance Win32_VideoController | Where-Object { $_.Name -match 'NVIDIA' }
            if ($nvidiaGpu) {
                $issues += "WARNING: Laptop NVIDIA GPUs are NOT supported for GPU-P. Intel integrated GPUs on laptops may work instead."
            }
        }
    } catch {
        Write-Log "GPU preflight battery/laptop check skipped: $($_.Exception.Message)" "WARN"
    }

    # Check 2: AMD Polaris (RX 580 etc) - no hardware video encoding
    try {
        $amdGpu = Get-CimInstance Win32_VideoController | Where-Object { $_.Name -match 'RX\s*5[678]0|Polaris' }
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
    if ($script:HostBuild -lt 22000) {
        $issues += "INFO: Windows 10 host detected - GPU-P will use AUTO (default GPU). Specific GPU selection requires Windows 11."
    }

    return $issues
}

function Get-GpuPartitionValues {
    <#
    .SYNOPSIS
        Calculates GPU partition adapter VRAM/Encode/Decode/Compute values
        based on a resource allocation percentage (Diobyte Version 1 approach).
    #>
    param(
        [ValidateRange(1,100)]
        [int]$Percentage = 100
    )
    [float]$divider = [math]::Round(100 / $Percentage, 2)
    return @{
        VRAM    = [math]::Round(1000000000 / $divider)
        Encode  = [math]::Round([decimal]18446744073709551615 / $divider)
        Decode  = [math]::Round(1000000000 / $divider)
        Compute = [math]::Round(1000000000 / $divider)
    }
}

function Copy-GpuServiceDriver {
    <#
    .SYNOPSIS
        Copies the GPU kernel-mode service driver directory to the VM's
        HostDriverStore (Diobyte Version 1 method).
    #>
    param(
        [string]$MountLetter,
        [string]$GPUName = "AUTO"
    )
    try {
        $gpu = $null
        if ($GPUName -eq "AUTO") {
            $partList = Get-WmiObject -Class "Msvm_PartitionableGpu" -ComputerName $env:COMPUTERNAME -Namespace "ROOT\virtualization\v2" -ErrorAction SilentlyContinue
            if ($partList) {
                $devPath = if ($partList.Name -is [array]) { $partList.Name[0] } else { $partList.Name }
                $gpu = Get-PnpDevice | Where-Object {
                    ($_.DeviceID -like "*$($devPath.Substring(8,16))*") -and ($_.Status -eq 'OK')
                } | Select-Object -First 1
            }
        } else {
            $gpu = Get-PnpDevice | Where-Object { ($_.Name -eq $GPUName) -and ($_.Status -eq 'OK') } | Select-Object -First 1
        }

        if (-not $gpu) {
            Write-Log "Could not resolve GPU PnP device for service driver copy" "WARN"
            return
        }

        $svcName = $gpu.Service
        if (-not $svcName) { return }

        $sysDriver = Get-WmiObject Win32_SystemDriver | Where-Object { $_.Name -eq $svcName }
        if ($sysDriver -and $sysDriver.PathName) {
            $servicePath = $sysDriver.PathName
            $serviceDriverDir  = ($servicePath.Split('\')[0..5]) -join '\'
            $serviceDriverDest = ("$MountLetter\" + ($servicePath.Split('\')[1..5] -join '\')).Replace('DriverStore','HostDriverStore')

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

    $result = [PSCustomObject]@{ WinVersion = "Unknown"; Build = 0 }

    try {
        $wimInfo = dism /Get-WimInfo /WimFile:"$WimFile" /Index:$Index /English 2>&1
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

    $isWin11 = ($DetectedWinVersion -eq 'Windows 11' -or $DetectedBuild -ge 22000)
    $isWin10 = (-not $isWin11) -and ($DetectedWinVersion -eq 'Windows 10' -or ($DetectedBuild -ge 10240 -and $DetectedBuild -lt 22000))
    $isLegacyWin10 = $isWin10 -and $DetectedBuild -gt 0 -and $DetectedBuild -lt 17134

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
    } elseif ($isWin10 -and $DetectedBuild -ge 17763) {
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
            @('MicrosoftUEFICertificateAuthority', 'MicrosoftWindows')
        } else {
            @('MicrosoftWindows', 'MicrosoftUEFICertificateAuthority')
        }
        CompatibilityNote       = $compatibilityNote
    }
}

function Apply-DetectedGuestDefaults {
    param(
        [Parameter(Mandatory = $true)][hashtable]$Controls,
        [Parameter(Mandatory = $true)]$Profile,
        [switch]$EmitLog
    )

    if ($Controls.ContainsKey('OSInfo') -and $Controls['OSInfo']) {
        $Controls['OSInfo'].Text = "$($Profile.Name)  (Build $($Profile.Build))"
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
    param(
        [string]$ImageFile,
        [int]$Index,
        [string]$ApplyDir,
        [bool]$PreferCompactApply = $true
    )

    $attempts = @()
    if ($PreferCompactApply) {
        $attempts += [PSCustomObject]@{ Label = 'with /Compact'; Args = @('/Apply-Image', "/ImageFile:$ImageFile", "/Index:$Index", "/ApplyDir:$ApplyDir", '/Compact') }
    }
    $attempts += [PSCustomObject]@{ Label = 'without /Compact'; Args = @('/Apply-Image', "/ImageFile:$ImageFile", "/Index:$Index", "/ApplyDir:$ApplyDir") }

    foreach ($attempt in $attempts) {
        Write-Log "DISM apply attempt: $($attempt.Label)"
        $dismOutput = & dism @($attempt.Args) 2>&1

        foreach ($line in $dismOutput) {
            $lineStr = $line.ToString().Trim()
            if ($lineStr -and $lineStr -notmatch '^\s*$') {
                if ($lineStr -match '\d+\.\d+%') {
                    $pctMatch = [regex]::Match($lineStr, '(\d+)\.\d+%')
                    if ($pctMatch.Success) {
                        $pct = [int]$pctMatch.Groups[1].Value
                        if ($pct % 25 -eq 0 -or $pct -ge 99) { Write-Log "  DISM: $lineStr" }
                    }
                } elseif ($lineStr -notmatch '^Deployment Image|^Version:') {
                    Write-Log "  DISM: $lineStr"
                }
            }
        }

        if ($LASTEXITCODE -eq 0) {
            return
        }

        Write-Log "DISM apply attempt failed ($($attempt.Label)) with exit code $LASTEXITCODE" "WARN"
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
        [int]$ResWidth,
        [int]$ResHeight,
        [bool]$IsWindows11 = $false
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

    # Keyboard Detection
    try { $keyboard = (Get-WinUserLanguageList)[0].InputMethodTips[0] }
    catch { $keyboard = "0409:00000409" }

    # Timezone Detection
    try { $timezone = (Get-TimeZone).Id }
    catch { $timezone = (Get-WmiObject Win32_TimeZone).StandardName }

    $xmlVMName       = Escape-XmlValue -Value $VMName
    $xmlUsername     = Escape-XmlValue -Value $Username
    $xmlPassword     = Escape-XmlValue -Value $passwordPlain
    $xmlUiLang       = Escape-XmlValue -Value $uiLang
    $xmlKeyboard     = Escape-XmlValue -Value $keyboard
    $xmlSystemLoc    = Escape-XmlValue -Value $systemLoc
    $xmlUserLoc      = Escape-XmlValue -Value $userLoc
    $xmlTimezone     = Escape-XmlValue -Value $timezone

    # Build the BypassNRO command for specialize pass (helps Win11 offline setup)
    $bypassBlock = ""
    if ($IsWindows11) {
        $bypassBlock = @"
    <component name="Microsoft-Windows-Deployment" processorArchitecture="amd64" publicKeyToken="31bf3856ad364e35" language="neutral" versionScope="nonSxS">
      <RunSynchronous>
        <RunSynchronousCommand wcm:action="add">
          <Order>1</Order>
          <Path>reg add "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\OOBE" /v BypassNRO /t REG_DWORD /d 1 /f</Path>
        </RunSynchronousCommand>
      </RunSynchronous>
    </component>
"@
    }

    return @"
<?xml version="1.0" encoding="utf-8"?>
<unattend xmlns="urn:schemas-microsoft-com:unattend" xmlns:wcm="http://schemas.microsoft.com/WMIConfig/2002/State">
  <settings pass="offlineServicing"></settings>

  <settings pass="windowsPE">
    <component name="Microsoft-Windows-International-Core-WinPE" processorArchitecture="amd64" publicKeyToken="31bf3856ad364e35" language="neutral" versionScope="nonSxS">
      <SetupUILanguage>
                <UILanguage>$xmlUiLang</UILanguage>
      </SetupUILanguage>
            <InputLocale>$xmlKeyboard</InputLocale>
            <SystemLocale>$xmlSystemLoc</SystemLocale>
            <UILanguage>$xmlUiLang</UILanguage>
            <UserLocale>$xmlUserLoc</UserLocale>
    </component>
    <component name="Microsoft-Windows-Setup" processorArchitecture="amd64" publicKeyToken="31bf3856ad364e35" language="neutral" versionScope="nonSxS">
      <UserData>
        <AcceptEula>true</AcceptEula>
      </UserData>
      <UseConfigurationSet>false</UseConfigurationSet>
    </component>
  </settings>

  <settings pass="generalize"></settings>

  <settings pass="specialize">
    <component name="Microsoft-Windows-Shell-Setup" processorArchitecture="amd64" publicKeyToken="31bf3856ad364e35" language="neutral" versionScope="nonSxS">
            <ComputerName>$xmlVMName</ComputerName>
            <TimeZone>$xmlTimezone</TimeZone>
    </component>
$bypassBlock
  </settings>

  <settings pass="oobeSystem">
    <component name="Microsoft-Windows-International-Core" processorArchitecture="amd64" publicKeyToken="31bf3856ad364e35" language="neutral" versionScope="nonSxS">
            <InputLocale>$xmlKeyboard</InputLocale>
            <SystemLocale>$xmlSystemLoc</SystemLocale>
            <UILanguage>$xmlUiLang</UILanguage>
            <UserLocale>$xmlUserLoc</UserLocale>
    </component>
    <component name="Microsoft-Windows-Shell-Setup" processorArchitecture="amd64" publicKeyToken="31bf3856ad364e35" language="neutral" versionScope="nonSxS">
      <UserAccounts>
        <LocalAccounts>
          <LocalAccount wcm:action="add">
                        <Name>$xmlUsername</Name>
                        <DisplayName>$xmlUsername</DisplayName>
            <Group>Administrators</Group>
            <Password>
                            <Value>$xmlPassword</Value>
              <PlainText>true</PlainText>
            </Password>
          </LocalAccount>
        </LocalAccounts>
      </UserAccounts>
      <AutoLogon>
                <Username>$xmlUsername</Username>
        <Enabled>true</Enabled>
        <LogonCount>9999</LogonCount>
        <Password>
                    <Value>$xmlPassword</Value>
          <PlainText>true</PlainText>
        </Password>
      </AutoLogon>
      <OOBE>
        <ProtectYourPC>3</ProtectYourPC>
        <HideEULAPage>true</HideEULAPage>
        <HideLocalAccountScreen>true</HideLocalAccountScreen>
        <HideOEMRegistrationScreen>true</HideOEMRegistrationScreen>
        <HideOnlineAccountScreens>true</HideOnlineAccountScreens>
        <HideWirelessSetupInOOBE>true</HideWirelessSetupInOOBE>
      </OOBE>
      <FirstLogonCommands>
        <SynchronousCommand wcm:action="add">
          <Order>1</Order>
          <CommandLine>cmd /c C:\Windows\Temp\QRes.exe /x:$ResWidth /y:$ResHeight</CommandLine>
          <Description>Set Display Resolution</Description>
        </SynchronousCommand>
        <SynchronousCommand wcm:action="add">
          <Order>2</Order>
          <CommandLine>cmd /c wmic useraccount where name="$Username" set PasswordExpires=False</CommandLine>
          <Description>Disable Password Expiration</Description>
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

    $list = @()
    foreach ($gpu in $gpus) {
        $pciShort    = ($gpu.Name -split "#")[1]
        $pnp         = Get-PnpDevice -InstanceId ("PCI\" + $pciShort + "*") -ErrorAction SilentlyContinue
        $friendlyName = if ($pnp) { $pnp.FriendlyName } else { $gpu.Name }

        # Skip AI accelerators
        if ($friendlyName -match '(?i)\bAI\b') { continue }

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
        [string]$DriverStore = "C:\Windows\System32\DriverStore\FileRepository",
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
        # GPU-related audio drivers (HDMI/DP audio)
        $audioDevs = Get-PnpDevice -Class MEDIA -Status OK -ErrorAction SilentlyContinue |
            Where-Object { $_.FriendlyName -match 'NVIDIA|AMD|Radeon|Intel.*Display|High Definition Audio' }
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
    $patterns = switch ($GpuVendor) {
        "NVIDIA" { @('nv_*', 'nvhd*', 'nvlt*', 'nvmd*', 'nvra*', 'nvsr*', 'nvwm*', 'nvam*') }
        "AMD"    { @('u0*', 'c0*', 'amd*', 'ati*') }
        "Intel"  { @('igfx*', 'iigd*', 'cui_*', 'dch_*', 'kit*') }
        default  { @('nv_*', 'nvhd*', 'nvlt*', 'nvmd*', 'nvra*', 'nvsr*', 'nvwm*',
                      'u0*', 'c0*', 'amd*', 'ati*',
                      'igfx*', 'iigd*', 'cui_*', 'dch_*', 'kit*') }
    }
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

    $target = Join-Path "$($MountLetter)\" $Destination

    # If target exists and ForceDelete, remove it first
    if ($ForceDelete -and (Test-Path $target)) {
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
            if ($LASTEXITCODE -gt 7) {
                Write-Log "[$VMName] Robocopy returned exit code $LASTEXITCODE" "WARN"
                # Fallback to Copy-Item
                Copy-Item -Path (Join-Path $Source $FileMask) -Destination $target -Recurse -Force -ErrorAction Stop
            }
        } else {
            Copy-Item -Path (Join-Path $Source $FileMask) -Destination $target -Recurse -Force -ErrorAction Stop
        }
    } catch {
        Write-Log "[$VMName] ERROR copying files: $($_.Exception.Message)" "ERROR"
    }
}

function Copy-GpuDriverFolders {
    <#
    .SYNOPSIS
        Smart GPU driver copy - only copies GPU-relevant DriverStore folders.
        Checks for sufficient disk space and auto-expands VHD if needed.
    #>
    param(
        [string]$VMName,
        [string]$MountLetter,
        [string]$VhdPath,
        [string]$GpuVendor = "Auto",
        [bool]$SmartCopy = $true,
        [bool]$AutoExpand = $true
    )

    $HostDriverStore = "C:\Windows\System32\DriverStore\FileRepository"
    $VMDriverStore   = "Windows\System32\HostDriverStore\FileRepository"
    $targetBase      = Join-Path "$($MountLetter)\" $VMDriverStore

    if ($SmartCopy) {
        # ---- SMART COPY: Only GPU-relevant folders ----
        Write-Log "[$VMName] Using smart GPU driver copy (GPU folders only)"
        $gpuFolders = Get-GpuDriverStoreFolders -GpuVendor $GpuVendor

        if ($gpuFolders.Count -eq 0) {
            Write-Log "[$VMName] No GPU driver folders found!" "ERROR"
            return $false
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
    Write-Log "[$VMName] VHD free space: $([math]::Round($freeSpace / 1GB, 2)) GB"

    if ($totalSize -gt ($freeSpace - 1GB)) {
        if (-not $AutoExpand) {
            Write-Log "[$VMName] Insufficient space and auto-expand is disabled." "ERROR"
            return $false
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
            Mount-DiskImage -ImagePath $VhdPath -ErrorAction Stop
            Register-TrackedMountedImage -ImagePath $VhdPath
            Start-Sleep -Seconds 2

            $disk = Get-DiskImage -ImagePath $VhdPath | Get-Disk
            # Find the NTFS partition and extend it
            $partition = Get-Partition -DiskNumber $disk.Number |
                Where-Object { $_.Type -ne 'System' -and $_.Type -ne 'Reserved' -and $_.GptType -ne '{C12A7328-F81F-11D2-BA4B-00A0C93EC93B}' -and $_.GptType -ne '{E3C9E316-0B5C-4DB8-817D-F92DF00215AE}' } |
                Select-Object -Last 1

            if ($partition) {
                $maxPartSize = (Get-PartitionSupportedSize -DiskNumber $disk.Number -PartitionNumber $partition.PartitionNumber).SizeMax
                Resize-Partition -DiskNumber $disk.Number -PartitionNumber $partition.PartitionNumber -Size $maxPartSize
                Write-Log "[$VMName] VHD partition extended successfully" "OK"
            }

            # Re-check mounted Windows volume and free space
            Start-Sleep -Seconds 1
            $updatedDriveLetter = $null
            $partitions = Get-Partition -DiskNumber $disk.Number -ErrorAction SilentlyContinue |
                Where-Object { $_.Type -ne 'System' -and $_.Type -ne 'Reserved' -and $_.GptType -ne '{C12A7328-F81F-11D2-BA4B-00A0C93EC93B}' -and $_.GptType -ne '{E3C9E316-0B5C-4DB8-817D-F92DF00215AE}' }
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
            Write-Log "[$VMName] New free space: $([math]::Round($freeSpace / 1GB, 2)) GB"

            if ($totalSize -gt ($freeSpace - 1GB)) {
                Write-Log "[$VMName] Still insufficient space after expansion!" "ERROR"
                return $false
            }
        } catch {
            Write-Log "[$VMName] VHD expansion failed: $($_.Exception.Message)" "ERROR"
            # Try to remount for cleanup
            try {
                Mount-DiskImage -ImagePath $VhdPath -ErrorAction SilentlyContinue
                Register-TrackedMountedImage -ImagePath $VhdPath
            } catch {
                Write-Log "[$VMName] VHD remount after expansion failure also failed: $($_.Exception.Message)" "WARN"
            }
            return $false
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
                    if ($LASTEXITCODE -le 7) { $copied++ }
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
            return $false
        }
    } else {
        # Full copy using robocopy or Copy-Item
        Write-Log "[$VMName] Copying entire DriverStore (this may take a while)..."
        Copy-DriversToVhd -VMName $VMName -MountLetter $MountLetter -Source $HostDriverStore -Destination $VMDriverStore -ForceDelete
    }

    return $true
}

function Test-HostHasNvidiaGpu {
    $gpus = Get-CimInstance Win32_VideoController
    $nvidiaGpus = $gpus | Where-Object { $_.Name -match "NVIDIA" }
    if ($nvidiaGpus) {
        Write-Log "Host NVIDIA GPU(s): $(($nvidiaGpus | ForEach-Object { $_.Name }) -join ', ')"
        return $true
    }
    return $false
}

function Set-GpuPartitionForVM {
    param(
        [string]$VMName,
        [ValidateRange(10,100)]
        [int]$AllocationPercent = 100
    )

    $partitionValues = Get-GpuPartitionValues -Percentage $AllocationPercent
    Set-VMGpuPartitionAdapter -VMName $VMName `
        -MinPartitionVRAM $partitionValues.VRAM -MaxPartitionVRAM $partitionValues.VRAM -OptimalPartitionVRAM $partitionValues.VRAM `
        -MinPartitionEncode $partitionValues.Encode -MaxPartitionEncode $partitionValues.Encode -OptimalPartitionEncode $partitionValues.Encode `
        -MinPartitionDecode $partitionValues.Decode -MaxPartitionDecode $partitionValues.Decode -OptimalPartitionDecode $partitionValues.Decode `
        -MinPartitionCompute $partitionValues.Compute -MaxPartitionCompute $partitionValues.Compute -OptimalPartitionCompute $partitionValues.Compute
    Set-VM -VMName $VMName -GuestControlledCacheTypes $true
    Set-VM -VMName $VMName -LowMemoryMappedIoSpace 1GB
    Set-VM -VMName $VMName -HighMemoryMappedIoSpace 32GB
}

#endregion

#region ==================== GUI CONSTRUCTION ====================

# ---- Main Form ----
$form = New-Object System.Windows.Forms.Form
$form.Text              = "Hyper-V Toolkit • Version 1 • Diobyte • Made with love"
$form.Size              = New-Object System.Drawing.Size(1240, 790)
$form.FormBorderStyle   = 'FixedSingle'
$form.MaximizeBox       = $false
$form.StartPosition     = "CenterScreen"
$form.Font              = New-Object System.Drawing.Font("Segoe UI", 9.75)
$form.BackColor         = [System.Drawing.Color]::FromArgb(24, 26, 31)
$form.ForeColor         = [System.Drawing.Color]::White
$form.Padding           = New-Object System.Windows.Forms.Padding(6)

# Theme palette
$theme = @{
    Bg          = [System.Drawing.Color]::FromArgb(24, 26, 31)
    Card        = [System.Drawing.Color]::FromArgb(33, 37, 46)
    Surface     = [System.Drawing.Color]::FromArgb(40, 44, 54)
    Input       = [System.Drawing.Color]::FromArgb(31, 35, 43)
    Border      = [System.Drawing.Color]::FromArgb(62, 68, 82)
    Text        = [System.Drawing.Color]::FromArgb(239, 242, 248)
    Muted       = [System.Drawing.Color]::FromArgb(166, 176, 195)
    Accent      = [System.Drawing.Color]::FromArgb(59, 130, 246)
    AccentHover = [System.Drawing.Color]::FromArgb(37, 99, 235)
    Success     = [System.Drawing.Color]::FromArgb(22, 163, 74)
    SuccessHover= [System.Drawing.Color]::FromArgb(21, 128, 61)
    Danger      = [System.Drawing.Color]::FromArgb(185, 28, 28)
    DangerHover = [System.Drawing.Color]::FromArgb(153, 27, 27)
}

# ---- Tab Control ----
$tabControl            = New-Object System.Windows.Forms.TabControl
$tabControl.Location   = New-Object System.Drawing.Point(10, 8)
$tabControl.Size       = New-Object System.Drawing.Size(1205, 490)
$tabControl.Font       = New-Object System.Drawing.Font("Segoe UI Semibold", 10)
$tabControl.DrawMode   = [System.Windows.Forms.TabDrawMode]::OwnerDrawFixed
$tabControl.ItemSize   = New-Object System.Drawing.Size(150, 32)
$tabControl.SizeMode   = [System.Windows.Forms.TabSizeMode]::Fixed
$tabControl.Padding    = New-Object System.Drawing.Point(14, 4)
$tabControl.BackColor  = $theme.Card
$tabControl.Add_DrawItem({
    param($sender, $e)
    $isSelected = ($e.State -band [System.Windows.Forms.DrawItemState]::Selected) -eq [System.Windows.Forms.DrawItemState]::Selected
    $tabPage = $sender.TabPages[$e.Index]
    $rect = $e.Bounds

    $bg = if ($isSelected) { $theme.Accent } else { $theme.Surface }
    $fg = if ($isSelected) { [System.Drawing.Color]::White } else { $theme.Text }

    $brush = New-Object System.Drawing.SolidBrush($bg)
    $textBrush = New-Object System.Drawing.SolidBrush($fg)
    $sf = New-Object System.Drawing.StringFormat
    $sf.Alignment = [System.Drawing.StringAlignment]::Center
    $sf.LineAlignment = [System.Drawing.StringAlignment]::Center

    $e.Graphics.FillRectangle($brush, $rect)
    $e.Graphics.DrawString($tabPage.Text.Trim(), $sender.Font, $textBrush, $rect, $sf)

    $brush.Dispose()
    $textBrush.Dispose()
    $sf.Dispose()
})
$form.Controls.Add($tabControl)

# ============================================================
#  TAB 1: CREATE VM
# ============================================================
$tabCreate             = New-Object System.Windows.Forms.TabPage
$tabCreate.Text        = "  Create VM  "
$tabCreate.BackColor   = $theme.Card
$tabCreate.ForeColor   = $theme.Text
$tabControl.TabPages.Add($tabCreate)

$lblCreateHeader = New-Object System.Windows.Forms.Label
$lblCreateHeader.Text = "Build, configure, and launch Hyper-V VMs with secure defaults and automation"
$lblCreateHeader.AutoSize = $true
$lblCreateHeader.Location = New-Object System.Drawing.Point(10, 2)
$lblCreateHeader.ForeColor = $theme.Muted
$lblCreateHeader.Font = New-Object System.Drawing.Font("Segoe UI", 8.75)
$tabCreate.Controls.Add($lblCreateHeader)

# Vars to hold controls
$ctrlCreate = @{}

# --- Helper: Add Label+Control Row ---
function New-LabeledControl {
    param(
        [System.Windows.Forms.Control]$Parent,
        [int]$X, [int]$Y,
        [string]$LabelText,
        [int]$LabelWidth = 110,
        [string]$ControlType = "TextBox",
        [int]$ControlWidth = 280,
        [hashtable]$ControlProps = @{}
    )
    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text     = $LabelText
    $lbl.AutoSize = $true
    $lbl.Location = New-Object System.Drawing.Point($X, ($Y + 3))
    $lbl.ForeColor = [System.Drawing.Color]::White
    $Parent.Controls.Add($lbl)

    $ctrl = switch ($ControlType) {
        "TextBox"       { New-Object System.Windows.Forms.TextBox }
        "ComboBox"      { $c = New-Object System.Windows.Forms.ComboBox; $c.DropDownStyle = 'DropDownList'; $c }
        "NumericUpDown" { New-Object System.Windows.Forms.NumericUpDown }
        "CheckBox"      { New-Object System.Windows.Forms.CheckBox }
        "Label"         { $l = New-Object System.Windows.Forms.Label; $l.ForeColor = [System.Drawing.Color]::Cyan; $l }
    }
    $ctrl.Location = New-Object System.Drawing.Point(($X + $LabelWidth), $Y)
    if ($ControlType -ne "CheckBox") { $ctrl.Width = $ControlWidth }
    foreach ($k in $ControlProps.Keys) { $ctrl.$k = $ControlProps[$k] }
    $Parent.Controls.Add($ctrl)
    return $ctrl
}

# --- Left Column: VM Configuration ---
$grpConfig           = New-Object System.Windows.Forms.GroupBox
$grpConfig.Text      = "VM Configuration • Core Settings"
$grpConfig.ForeColor = $theme.Text
$grpConfig.Location  = New-Object System.Drawing.Point(8, 18)
$grpConfig.Size      = New-Object System.Drawing.Size(460, 445)
$grpConfig.BackColor = $theme.Surface
$tabCreate.Controls.Add($grpConfig)

$rowY = 22
$ctrlCreate["VMName"] = New-LabeledControl $grpConfig 12 $rowY "VM Name:" -ControlWidth 300
$rowY += 32

$ctrlCreate["VMLocation"] = New-LabeledControl $grpConfig 12 $rowY "VM Location:" -ControlWidth 240
$btnBrowseVM = New-Object System.Windows.Forms.Button
$btnBrowseVM.Text     = "Browse"
$btnBrowseVM.Size     = New-Object System.Drawing.Size(52, 24)
$btnBrowseVM.Location = New-Object System.Drawing.Point(365, $rowY)
$btnBrowseVM.FlatStyle = 'Flat'
$grpConfig.Controls.Add($btnBrowseVM)
# Default VM location
try { $ctrlCreate["VMLocation"].Text = (Get-VMHost).VirtualMachinePath } catch { $ctrlCreate["VMLocation"].Text = "C:\HyperV" }
$rowY += 32

$ctrlCreate["ISOPath"] = New-LabeledControl $grpConfig 12 $rowY "ISO File:" -ControlWidth 240
$btnBrowseISO = New-Object System.Windows.Forms.Button
$btnBrowseISO.Text     = "Browse"
$btnBrowseISO.Size     = New-Object System.Drawing.Size(52, 24)
$btnBrowseISO.Location = New-Object System.Drawing.Point(365, $rowY)
$btnBrowseISO.FlatStyle = 'Flat'
$grpConfig.Controls.Add($btnBrowseISO)
$rowY += 32

$ctrlCreate["Edition"] = New-LabeledControl $grpConfig 12 $rowY "Win Edition:" -ControlType ComboBox -ControlWidth 300
$rowY += 30

# OS Detection label (auto-filled)
$ctrlCreate["OSInfo"] = New-LabeledControl $grpConfig 12 $rowY "Detected OS:" -ControlType Label -ControlWidth 300 -ControlProps @{ AutoSize = $true; Text = "(select an ISO to detect)" }
$rowY += 28

# Separator
$sep1 = New-Object System.Windows.Forms.Label
$sep1.Text      = ""
$sep1.BorderStyle = 'Fixed3D'
$sep1.Size      = New-Object System.Drawing.Size(430, 2)
$sep1.Location  = New-Object System.Drawing.Point(12, $rowY)
$grpConfig.Controls.Add($sep1)
$rowY += 10

$ctrlCreate["Username"] = New-LabeledControl $grpConfig 12 $rowY "Local User:" -ControlWidth 200
$ctrlCreate["Username"].Text = "User"
$rowY += 32

$ctrlCreate["Password"] = New-LabeledControl $grpConfig 12 $rowY "Password:" -ControlWidth 200
$ctrlCreate["Password"].Text = "Password1"
$ctrlCreate["Password"].UseSystemPasswordChar = $true
$rowY += 32

# Separator
$sep2 = New-Object System.Windows.Forms.Label
$sep2.Text      = ""
$sep2.BorderStyle = 'Fixed3D'
$sep2.Size      = New-Object System.Drawing.Size(430, 2)
$sep2.Location  = New-Object System.Drawing.Point(12, $rowY)
$grpConfig.Controls.Add($sep2)
$rowY += 10

$ctrlCreate["vCPU"] = New-LabeledControl $grpConfig 12 $rowY "vCPUs:" -ControlType NumericUpDown -ControlWidth 80 `
    -ControlProps @{ Minimum = 1; Maximum = [Environment]::ProcessorCount; Value = [Math]::Min(4, [Environment]::ProcessorCount) }
$rowY += 30

$totalRamGB = [math]::Round((Get-CimInstance Win32_ComputerSystem).TotalPhysicalMemory / 1GB)
$ctrlCreate["Memory"] = New-LabeledControl $grpConfig 12 $rowY "Memory (GB):" -ControlType NumericUpDown -ControlWidth 80 `
    -ControlProps @{ Minimum = 1; Maximum = $totalRamGB; Value = [Math]::Min(8, $totalRamGB) }
$rowY += 30

$ctrlCreate["DiskSize"] = New-LabeledControl $grpConfig 12 $rowY "Disk (GB):" -ControlType NumericUpDown -ControlWidth 80 `
    -ControlProps @{ Minimum = 20; Maximum = 2048; Value = 80 }
$rowY += 30

$ctrlCreate["Switch"] = New-LabeledControl $grpConfig 12 $rowY "Virtual Switch:" -ControlType ComboBox -ControlWidth 200
try {
    Get-VMSwitch | Select-Object -ExpandProperty Name | ForEach-Object { [void]$ctrlCreate["Switch"].Items.Add($_) }
    if ($ctrlCreate["Switch"].Items.Count -gt 0) { $ctrlCreate["Switch"].SelectedIndex = 0 }
} catch {
    Write-Log "Could not enumerate Hyper-V virtual switches in UI: $($_.Exception.Message)" "WARN"
}
$rowY += 30

$ctrlCreate["Resolution"] = New-LabeledControl $grpConfig 12 $rowY "Resolution:" -ControlType ComboBox -ControlWidth 150
@("800x600","1024x768","1280x720","1280x800","1280x1024","1366x768","1440x900",
  "1600x900","1680x1050","1920x1080") | ForEach-Object { [void]$ctrlCreate["Resolution"].Items.Add($_) }
$ctrlCreate["Resolution"].SelectedItem = "1920x1080"
$rowY += 30

$ctrlCreate["CheckpointMode"] = New-LabeledControl $grpConfig 12 $rowY "Checkpoint Mode:" -ControlType ComboBox -ControlWidth 150
@("Disabled","Production","ProductionOnly","Standard") | ForEach-Object { [void]$ctrlCreate["CheckpointMode"].Items.Add($_) }
$ctrlCreate["CheckpointMode"].SelectedItem = "Disabled"
$rowY += 30

$ctrlCreate["DynamicMemMin"] = New-LabeledControl $grpConfig 12 $rowY "Dynamic Min (GB):" -ControlType NumericUpDown -ControlWidth 80 `
        -ControlProps @{ Minimum = 1; Maximum = $totalRamGB; Value = 1 }
$rowY += 30

$ctrlCreate["DynamicMemMax"] = New-LabeledControl $grpConfig 12 $rowY "Dynamic Max (GB):" -ControlType NumericUpDown -ControlWidth 80 `
        -ControlProps @{ Minimum = 1; Maximum = $totalRamGB; Value = [Math]::Min(16, $totalRamGB) }

# --- Right Column: Options ---

# GroupBox: Boot & Hardware
$grpBoot           = New-Object System.Windows.Forms.GroupBox
$grpBoot.Text      = "Boot && Hardware • Security"
$grpBoot.ForeColor = [System.Drawing.Color]::White
$grpBoot.Location  = New-Object System.Drawing.Point(478, 18)
$grpBoot.Size      = New-Object System.Drawing.Size(350, 105)
$tabCreate.Controls.Add($grpBoot)

$ctrlCreate["SecureBoot"] = New-Object System.Windows.Forms.CheckBox
$ctrlCreate["SecureBoot"].Text     = "Secure Boot (auto: ON for Win11, OFF for Win10)"
$ctrlCreate["SecureBoot"].AutoSize = $true
$ctrlCreate["SecureBoot"].Checked  = $true
$ctrlCreate["SecureBoot"].Location = New-Object System.Drawing.Point(12, 24)
$ctrlCreate["SecureBoot"].ForeColor = [System.Drawing.Color]::White
$grpBoot.Controls.Add($ctrlCreate["SecureBoot"])

$ctrlCreate["TPM"] = New-Object System.Windows.Forms.CheckBox
$ctrlCreate["TPM"].Text     = "Virtual TPM (required for Win11)"
$ctrlCreate["TPM"].AutoSize = $true
$ctrlCreate["TPM"].Checked  = $true
$ctrlCreate["TPM"].Location = New-Object System.Drawing.Point(12, 52)
$ctrlCreate["TPM"].ForeColor = [System.Drawing.Color]::White
$grpBoot.Controls.Add($ctrlCreate["TPM"])

$ctrlCreate["VHDType"] = New-Object System.Windows.Forms.CheckBox
$ctrlCreate["VHDType"].Text     = "Fixed-size VHD (default: Dynamic / expanding)"
$ctrlCreate["VHDType"].AutoSize = $true
$ctrlCreate["VHDType"].Checked  = $false
$ctrlCreate["VHDType"].Location = New-Object System.Drawing.Point(12, 80)
$ctrlCreate["VHDType"].ForeColor = [System.Drawing.Color]::White
$grpBoot.Controls.Add($ctrlCreate["VHDType"])

# GroupBox: VM Options
$grpOpts           = New-Object System.Windows.Forms.GroupBox
$grpOpts.Text      = "VM Options • Runtime"
$grpOpts.ForeColor = [System.Drawing.Color]::White
$grpOpts.Location  = New-Object System.Drawing.Point(478, 128)
$grpOpts.Size      = New-Object System.Drawing.Size(350, 134)
$tabCreate.Controls.Add($grpOpts)

$chkNames = @(
    @{ Key = "DynamicMem";       Text = "Enable Dynamic Memory";       X = 12;  Y = 22; Default = $false },
    @{ Key = "EnhancedSession";  Text = "Enable Enhanced Session Mode"; X = 12;  Y = 48; Default = $false },
    @{ Key = "StartVM";          Text = "Start VM after creation";     X = 12;  Y = 74; Default = $true },
    @{ Key = "StrictLegacyMode"; Text = "Strict Legacy Mode (Win10 fallback)"; X = 12;  Y = 100; Default = $false },
    @{ Key = "AutoCreateSwitch"; Text = "Auto-create NAT switch";      X = 185; Y = 22; Default = $true },
    @{ Key = "EnableMetering";   Text = "Enable Resource Metering";    X = 185; Y = 48; Default = $true }
)
foreach ($chk in $chkNames) {
    $cb = New-Object System.Windows.Forms.CheckBox
    $cb.Text     = $chk.Text
    $cb.AutoSize = $true
    $cb.Checked  = $chk.Default
    $cb.Location = New-Object System.Drawing.Point($chk.X, $chk.Y)
    $cb.ForeColor = [System.Drawing.Color]::White
    $grpOpts.Controls.Add($cb)
    $ctrlCreate[$chk.Key] = $cb
}

# GroupBox: Post-Install Software
$grpSoft           = New-Object System.Windows.Forms.GroupBox
$grpSoft.Text      = "Post-Install Software && Advanced"
$grpSoft.ForeColor = [System.Drawing.Color]::White
$grpSoft.Location  = New-Object System.Drawing.Point(478, 262)
$grpSoft.Size      = New-Object System.Drawing.Size(350, 195)
$tabCreate.Controls.Add($grpSoft)

$softwareChecks = @(
    @{ Key = "Parsec";       Text = "Parsec (Per Computer)";  X = 12;  Y = 22 },
    @{ Key = "VBCable";      Text = "VB-Audio Cable";         X = 185; Y = 22 },
    @{ Key = "USBMMIDD";     Text = "Virtual Display Driver"; X = 12;  Y = 48 },
    @{ Key = "RDP";          Text = "Remote Desktop";         X = 185; Y = 48 },
    @{ Key = "Share";        Text = "Share Folder";           X = 12;  Y = 74 },
    @{ Key = "PauseUpdate";  Text = "Pause Win Updates";      X = 185; Y = 74 },
    @{ Key = "FullUpdate";   Text = "Full Win Updates";       X = 12;  Y = 100 },
    @{ Key = "NestedVirt";   Text = "Nested Virtualization";  X = 185; Y = 100 },
    @{ Key = "NestedNetFollowup"; Text = "Nested Net (MAC spoofing)"; X = 12;  Y = 126 },
    @{ Key = "ResetBootOrder";    Text = "Reset boot order after recovery"; X = 185; Y = 126 },
    @{ Key = "GoldenImage";       Text = "Create from Golden VHDX"; X = 12; Y = 152 }
)
foreach ($sw in $softwareChecks) {
    $cb = New-Object System.Windows.Forms.CheckBox
    $cb.Text     = $sw.Text
    $cb.AutoSize = $true
    $cb.Checked  = $false
    $cb.Location = New-Object System.Drawing.Point($sw.X, $sw.Y)
    $cb.ForeColor = [System.Drawing.Color]::White
    $grpSoft.Controls.Add($cb)
    $ctrlCreate[$sw.Key] = $cb
}

$ctrlCreate["GoldenParentVHD"] = New-LabeledControl $grpSoft 12 172 "Parent VHDX:" -ControlWidth 230
$btnBrowseGolden = New-Object System.Windows.Forms.Button
$btnBrowseGolden.Text     = "Browse"
$btnBrowseGolden.Size     = New-Object System.Drawing.Size(52, 24)
$btnBrowseGolden.Location = New-Object System.Drawing.Point(290, 170)
$btnBrowseGolden.FlatStyle = 'Flat'
$grpSoft.Controls.Add($btnBrowseGolden)

$ctrlCreate["ModeHint"] = New-Object System.Windows.Forms.Label
$ctrlCreate["ModeHint"].Text = "Mode: ISO Deploy - Uses ISO, selected edition, and unattended setup."
$ctrlCreate["ModeHint"].Size = New-Object System.Drawing.Size(820, 14)
$ctrlCreate["ModeHint"].Location = New-Object System.Drawing.Point(8, 432)
$ctrlCreate["ModeHint"].ForeColor = [System.Drawing.Color]::Silver
$tabCreate.Controls.Add($ctrlCreate["ModeHint"])

# Create VM Button
$btnCreateVM           = New-Object System.Windows.Forms.Button
$btnCreateVM.Text      = "Create VM  →"
$btnCreateVM.Size      = New-Object System.Drawing.Size(140, 36)
$btnCreateVM.Location  = New-Object System.Drawing.Point(540, 448)
$btnCreateVM.FlatStyle = 'Flat'
$btnCreateVM.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
$btnCreateVM.ForeColor = [System.Drawing.Color]::White
$btnCreateVM.Font      = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$tabCreate.Controls.Add($btnCreateVM)

# Create VM status + progress
$ctrlCreate["CreateStatus"] = New-Object System.Windows.Forms.Label
$ctrlCreate["CreateStatus"].Text = "Ready to create VM"
$ctrlCreate["CreateStatus"].Size = New-Object System.Drawing.Size(450, 18)
$ctrlCreate["CreateStatus"].Location = New-Object System.Drawing.Point(8, 452)
$ctrlCreate["CreateStatus"].ForeColor = [System.Drawing.Color]::Cyan
$tabCreate.Controls.Add($ctrlCreate["CreateStatus"])

$ctrlCreate["CreateProgress"] = New-Object System.Windows.Forms.ProgressBar
$ctrlCreate["CreateProgress"].Minimum = 0
$ctrlCreate["CreateProgress"].Maximum = 100
$ctrlCreate["CreateProgress"].Value = 0
$ctrlCreate["CreateProgress"].Style = 'Continuous'
$ctrlCreate["CreateProgress"].Size = New-Object System.Drawing.Size(450, 14)
$ctrlCreate["CreateProgress"].Location = New-Object System.Drawing.Point(8, 470)
$tabCreate.Controls.Add($ctrlCreate["CreateProgress"])

# ============================================================
#  TAB 2: GPU MANAGER
# ============================================================
$tabGPU             = New-Object System.Windows.Forms.TabPage
$tabGPU.Text        = "  GPU Manager  "
$tabGPU.BackColor   = $theme.Card
$tabGPU.ForeColor   = $theme.Text
$tabControl.TabPages.Add($tabGPU)

$lblGpuHeader = New-Object System.Windows.Forms.Label
$lblGpuHeader.Text = "Select VMs, configure GPU-P allocation, and inject/update driver stacks"
$lblGpuHeader.AutoSize = $true
$lblGpuHeader.Location = New-Object System.Drawing.Point(10, 2)
$lblGpuHeader.ForeColor = $theme.Muted
$lblGpuHeader.Font = New-Object System.Drawing.Font("Segoe UI", 8.75)
$tabGPU.Controls.Add($lblGpuHeader)

$ctrlGPU = @{}

# --- Left: VM Selection ---
$grpVMs           = New-Object System.Windows.Forms.GroupBox
$grpVMs.Text      = "Select VMs to Update"
$grpVMs.ForeColor = [System.Drawing.Color]::White
$grpVMs.Location  = New-Object System.Drawing.Point(8, 18)
$grpVMs.Size      = New-Object System.Drawing.Size(360, 400)
$tabGPU.Controls.Add($grpVMs)

$lblVmSearch = New-Object System.Windows.Forms.Label
$lblVmSearch.Text = "Filter:"
$lblVmSearch.AutoSize = $true
$lblVmSearch.Location = New-Object System.Drawing.Point(8, 24)
$lblVmSearch.ForeColor = [System.Drawing.Color]::White
$grpVMs.Controls.Add($lblVmSearch)

$ctrlGPU["VmSearch"] = New-Object System.Windows.Forms.TextBox
$ctrlGPU["VmSearch"].Location = New-Object System.Drawing.Point(52, 21)
$ctrlGPU["VmSearch"].Size = New-Object System.Drawing.Size(210, 24)
$ctrlGPU["VmSearch"].Text = ""
$grpVMs.Controls.Add($ctrlGPU["VmSearch"])

$btnClearVmSearch = New-Object System.Windows.Forms.Button
$btnClearVmSearch.Text = "Clear"
$btnClearVmSearch.Size = New-Object System.Drawing.Size(64, 24)
$btnClearVmSearch.Location = New-Object System.Drawing.Point(270, 21)
$btnClearVmSearch.FlatStyle = 'Flat'
$btnClearVmSearch.ForeColor = [System.Drawing.Color]::White
$grpVMs.Controls.Add($btnClearVmSearch)

# Scrollable panel for VM checkboxes
$vmPanel = New-Object System.Windows.Forms.Panel
$vmPanel.Location   = New-Object System.Drawing.Point(8, 52)
$vmPanel.Size       = New-Object System.Drawing.Size(340, 308)
$vmPanel.AutoScroll = $true
$vmPanel.BackColor  = [System.Drawing.Color]::FromArgb(30, 30, 30)
$grpVMs.Controls.Add($vmPanel)

$ctrlGPU["VMCheckboxes"] = @()

function Update-VMList {
    $vmPanel.Controls.Clear()
    $script:gpuVMCheckboxes = @()
    $filterText = ""
    if ($ctrlGPU.ContainsKey("VmSearch") -and $ctrlGPU["VmSearch"]) {
        $filterText = [string]$ctrlGPU["VmSearch"].Text
    }
    $vms = Get-VM -ErrorAction SilentlyContinue
    if (-not [string]::IsNullOrWhiteSpace($filterText)) {
        $vms = $vms | Where-Object { $_.Name -match [regex]::Escape($filterText) }
    }
    $y = 5
    foreach ($vm in $vms) {
        $cb = New-Object System.Windows.Forms.CheckBox
        $cb.Text     = $vm.Name
        $cb.AutoSize = $true
        $cb.Location = New-Object System.Drawing.Point(8, $y)
        $cb.ForeColor = [System.Drawing.Color]::White
        $vmPanel.Controls.Add($cb)
        $script:gpuVMCheckboxes += $cb
        $y += 26
    }
    $ctrlGPU["VMCheckboxes"] = $script:gpuVMCheckboxes
}
Update-VMList

$ctrlGPU["VmSearch"].Add_TextChanged({ Update-VMList })
$btnClearVmSearch.Add_Click({ $ctrlGPU["VmSearch"].Text = ""; Update-VMList })

# Select All / None buttons
$btnSelectAll = New-Object System.Windows.Forms.Button
$btnSelectAll.Text     = "All"
$btnSelectAll.Size     = New-Object System.Drawing.Size(60, 24)
$btnSelectAll.Location = New-Object System.Drawing.Point(8, 368)
$btnSelectAll.FlatStyle = 'Flat'
$btnSelectAll.ForeColor = [System.Drawing.Color]::White
$btnSelectAll.Add_Click({ foreach ($cb in $ctrlGPU["VMCheckboxes"]) { $cb.Checked = $true } })
$grpVMs.Controls.Add($btnSelectAll)

$btnSelectNone = New-Object System.Windows.Forms.Button
$btnSelectNone.Text     = "None"
$btnSelectNone.Size     = New-Object System.Drawing.Size(60, 24)
$btnSelectNone.Location = New-Object System.Drawing.Point(74, 368)
$btnSelectNone.FlatStyle = 'Flat'
$btnSelectNone.ForeColor = [System.Drawing.Color]::White
$btnSelectNone.Add_Click({ foreach ($cb in $ctrlGPU["VMCheckboxes"]) { $cb.Checked = $false } })
$grpVMs.Controls.Add($btnSelectNone)

$btnRefreshVMs = New-Object System.Windows.Forms.Button
$btnRefreshVMs.Text     = "Refresh"
$btnRefreshVMs.Size     = New-Object System.Drawing.Size(70, 24)
$btnRefreshVMs.Location = New-Object System.Drawing.Point(140, 368)
$btnRefreshVMs.FlatStyle = 'Flat'
$btnRefreshVMs.ForeColor = [System.Drawing.Color]::White
$btnRefreshVMs.Add_Click({ Update-VMList })
$grpVMs.Controls.Add($btnRefreshVMs)

# --- Right: GPU Settings ---
$grpGPUSettings           = New-Object System.Windows.Forms.GroupBox
$grpGPUSettings.Text      = "GPU-P Settings"
$grpGPUSettings.ForeColor = [System.Drawing.Color]::White
$grpGPUSettings.Location  = New-Object System.Drawing.Point(378, 6)
$grpGPUSettings.Size      = New-Object System.Drawing.Size(450, 190)
$tabGPU.Controls.Add($grpGPUSettings)

# GPU Selector
$lblGpu = New-Object System.Windows.Forms.Label
$lblGpu.Text     = "GPU-P GPU:"
$lblGpu.AutoSize = $true
$lblGpu.Location = New-Object System.Drawing.Point(12, 28)
$lblGpu.ForeColor = [System.Drawing.Color]::White
$grpGPUSettings.Controls.Add($lblGpu)

$ctrlGPU["GpuSelector"] = New-Object System.Windows.Forms.ComboBox
$ctrlGPU["GpuSelector"].DropDownStyle = 'DropDownList'
$ctrlGPU["GpuSelector"].Width    = 340
$ctrlGPU["GpuSelector"].Location = New-Object System.Drawing.Point(95, 25)
$grpGPUSettings.Controls.Add($ctrlGPU["GpuSelector"])

# Populate GPU list
$script:GpuPList = Get-GpuPProviders
$script:GpuPList | ForEach-Object { [void]$ctrlGPU["GpuSelector"].Items.Add($_.Friendly) }
if ($ctrlGPU["GpuSelector"].Items.Count -gt 0) { $ctrlGPU["GpuSelector"].SelectedIndex = 0 }

if (-not $script:SupportsGpuInstancePath) {
    $lblGpuWarn = New-Object System.Windows.Forms.Label
    $lblGpuWarn.Text     = "Note: This host build uses default GPU selection for GPU-P."
    $lblGpuWarn.AutoSize = $true
    $lblGpuWarn.Location = New-Object System.Drawing.Point(12, 56)
    $lblGpuWarn.ForeColor = [System.Drawing.Color]::Gold
    $lblGpuWarn.Font     = New-Object System.Drawing.Font("Segoe UI", 8)
    $grpGPUSettings.Controls.Add($lblGpuWarn)
}

# Driver copy mode
$lblCopyMode = New-Object System.Windows.Forms.Label
$lblCopyMode.Text     = "Driver Copy Mode:"
$lblCopyMode.AutoSize = $true
$lblCopyMode.Location = New-Object System.Drawing.Point(12, 82)
$lblCopyMode.ForeColor = [System.Drawing.Color]::White
$grpGPUSettings.Controls.Add($lblCopyMode)

$ctrlGPU["SmartCopy"] = New-Object System.Windows.Forms.RadioButton
$ctrlGPU["SmartCopy"].Text     = "Smart (GPU folders only - recommended)"
$ctrlGPU["SmartCopy"].AutoSize = $true
$ctrlGPU["SmartCopy"].Checked  = $true
$ctrlGPU["SmartCopy"].Location = New-Object System.Drawing.Point(20, 104)
$ctrlGPU["SmartCopy"].ForeColor = [System.Drawing.Color]::White
$grpGPUSettings.Controls.Add($ctrlGPU["SmartCopy"])

$ctrlGPU["FullCopy"] = New-Object System.Windows.Forms.RadioButton
$ctrlGPU["FullCopy"].Text     = "Full DriverStore (larger, may need big VHD)"
$ctrlGPU["FullCopy"].AutoSize = $true
$ctrlGPU["FullCopy"].Location = New-Object System.Drawing.Point(20, 128)
$ctrlGPU["FullCopy"].ForeColor = [System.Drawing.Color]::White
$grpGPUSettings.Controls.Add($ctrlGPU["FullCopy"])

# GPU Resource Allocation % (Diobyte Version 1 best practices)
$lblGpuAlloc = New-Object System.Windows.Forms.Label
$lblGpuAlloc.Text     = "GPU Resource Allocation %:"
$lblGpuAlloc.AutoSize = $true
$lblGpuAlloc.Location = New-Object System.Drawing.Point(12, 158)
$lblGpuAlloc.ForeColor = [System.Drawing.Color]::White
$grpGPUSettings.Controls.Add($lblGpuAlloc)

$ctrlGPU["GpuAllocSlider"] = New-Object System.Windows.Forms.TrackBar
$ctrlGPU["GpuAllocSlider"].Minimum  = 10
$ctrlGPU["GpuAllocSlider"].Maximum  = 100
$ctrlGPU["GpuAllocSlider"].Value    = 100
$ctrlGPU["GpuAllocSlider"].TickFrequency = 10
$ctrlGPU["GpuAllocSlider"].SmallChange   = 5
$ctrlGPU["GpuAllocSlider"].LargeChange   = 10
$ctrlGPU["GpuAllocSlider"].Location = New-Object System.Drawing.Point(190, 153)
$ctrlGPU["GpuAllocSlider"].Size     = New-Object System.Drawing.Size(200, 30)
$ctrlGPU["GpuAllocSlider"].BackColor = [System.Drawing.Color]::FromArgb(40, 40, 40)
$grpGPUSettings.Controls.Add($ctrlGPU["GpuAllocSlider"])

$ctrlGPU["GpuAllocLabel"] = New-Object System.Windows.Forms.Label
$ctrlGPU["GpuAllocLabel"].Text     = "100%"
$ctrlGPU["GpuAllocLabel"].AutoSize = $true
$ctrlGPU["GpuAllocLabel"].Location = New-Object System.Drawing.Point(395, 158)
$ctrlGPU["GpuAllocLabel"].ForeColor = [System.Drawing.Color]::Cyan
$ctrlGPU["GpuAllocLabel"].Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$grpGPUSettings.Controls.Add($ctrlGPU["GpuAllocLabel"])

$ctrlGPU["GpuAllocSlider"].Add_ValueChanged({
    $ctrlGPU["GpuAllocLabel"].Text = "$($ctrlGPU['GpuAllocSlider'].Value)%"
})

# GPU Options
$grpGPUOpts           = New-Object System.Windows.Forms.GroupBox
$grpGPUOpts.Text      = "Options"
$grpGPUOpts.ForeColor = [System.Drawing.Color]::White
$grpGPUOpts.Location  = New-Object System.Drawing.Point(378, 200)
$grpGPUOpts.Size      = New-Object System.Drawing.Size(450, 95)
$tabGPU.Controls.Add($grpGPUOpts)

$ctrlGPU["StartVM"] = New-Object System.Windows.Forms.CheckBox
$ctrlGPU["StartVM"].Text     = "Start VM after update"
$ctrlGPU["StartVM"].AutoSize = $true
$ctrlGPU["StartVM"].Location = New-Object System.Drawing.Point(12, 24)
$ctrlGPU["StartVM"].ForeColor = [System.Drawing.Color]::White
$grpGPUOpts.Controls.Add($ctrlGPU["StartVM"])

$ctrlGPU["AutoExpand"] = New-Object System.Windows.Forms.CheckBox
$ctrlGPU["AutoExpand"].Text     = "Auto-expand VHD if insufficient space"
$ctrlGPU["AutoExpand"].AutoSize = $true
$ctrlGPU["AutoExpand"].Checked  = $true
$ctrlGPU["AutoExpand"].Location = New-Object System.Drawing.Point(12, 48)
$ctrlGPU["AutoExpand"].ForeColor = [System.Drawing.Color]::White
$grpGPUOpts.Controls.Add($ctrlGPU["AutoExpand"])

$ctrlGPU["CopySvcDriver"] = New-Object System.Windows.Forms.CheckBox
$ctrlGPU["CopySvcDriver"].Text     = "Copy GPU service driver (recommended)"
$ctrlGPU["CopySvcDriver"].AutoSize = $true
$ctrlGPU["CopySvcDriver"].Checked  = $true
$ctrlGPU["CopySvcDriver"].Location = New-Object System.Drawing.Point(12, 72)
$ctrlGPU["CopySvcDriver"].ForeColor = [System.Drawing.Color]::White
$grpGPUOpts.Controls.Add($ctrlGPU["CopySvcDriver"])

# Update GPU Button
$btnUpdateGPU           = New-Object System.Windows.Forms.Button
$btnUpdateGPU.Text      = "Update GPU Drivers  →"
$btnUpdateGPU.Size      = New-Object System.Drawing.Size(160, 36)
$btnUpdateGPU.Location  = New-Object System.Drawing.Point(540, 448)
$btnUpdateGPU.FlatStyle = 'Flat'
$btnUpdateGPU.BackColor = [System.Drawing.Color]::FromArgb(0, 153, 51)
$btnUpdateGPU.ForeColor = [System.Drawing.Color]::White
$btnUpdateGPU.Font      = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$tabGPU.Controls.Add($btnUpdateGPU)

# ============================================================
#  SHARED LOG PANEL
# ============================================================
$script:LogBox           = New-Object System.Windows.Forms.RichTextBox
$script:LogBox.Location  = New-Object System.Drawing.Point(10, 505)
$script:LogBox.Size      = New-Object System.Drawing.Size(1095, 210)
$script:LogBox.ReadOnly  = $true
$script:LogBox.BackColor = [System.Drawing.Color]::FromArgb(17, 19, 24)
$script:LogBox.ForeColor = [System.Drawing.Color]::FromArgb(166, 243, 160)
$script:LogBox.Font      = New-Object System.Drawing.Font("Consolas", 9.5)
$script:LogBox.WordWrap  = $true
$script:LogBox.BorderStyle = 'FixedSingle'
$form.Controls.Add($script:LogBox)

# Clear Log button
$btnClearLog           = New-Object System.Windows.Forms.Button
$btnClearLog.Text      = "Clear Log"
$btnClearLog.Size      = New-Object System.Drawing.Size(85, 30)
$btnClearLog.Location  = New-Object System.Drawing.Point(1115, 505)
$btnClearLog.FlatStyle = 'Flat'
$btnClearLog.ForeColor = [System.Drawing.Color]::White
$btnClearLog.Add_Click({ $script:LogBox.Clear() })
$form.Controls.Add($btnClearLog)

# Save Log button
$btnSaveLog            = New-Object System.Windows.Forms.Button
$btnSaveLog.Text       = "Save Log"
$btnSaveLog.Size       = New-Object System.Drawing.Size(85, 30)
$btnSaveLog.Location   = New-Object System.Drawing.Point(1115, 540)
$btnSaveLog.FlatStyle  = 'Flat'
$btnSaveLog.ForeColor  = [System.Drawing.Color]::White
$btnSaveLog.Add_Click({
    $saveDlg = New-Object System.Windows.Forms.SaveFileDialog
    $saveDlg.Filter = "Log files (*.log)|*.log|Text files (*.txt)|*.txt|All files (*.*)|*.*"
    $saveDlg.FileName = "HyperV-Toolkit_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
    if ($saveDlg.ShowDialog() -eq 'OK') {
        try {
            $script:LogBox.Text | Out-File -FilePath $saveDlg.FileName -Encoding UTF8 -Force
            Write-Log "Log saved to: $($saveDlg.FileName)" "OK"
        } catch {
            Write-Log "Failed to save log: $($_.Exception.Message)" "ERROR"
        }
    }
})
$form.Controls.Add($btnSaveLog)

# Exit Button
$btnExit           = New-Object System.Windows.Forms.Button
$btnExit.Text      = "EXIT"
$btnExit.Size      = New-Object System.Drawing.Size(85, 30)
$btnExit.Location  = New-Object System.Drawing.Point(1115, 575)
$btnExit.FlatStyle = 'Flat'
$btnExit.BackColor = [System.Drawing.Color]::FromArgb(180, 40, 40)
$btnExit.ForeColor = [System.Drawing.Color]::White
$btnExit.Add_Click({ $form.Close() })
$form.Controls.Add($btnExit)

# ---- Modern UI styling helpers ----
function Set-ButtonHover {
    param(
        [System.Windows.Forms.Button]$Button,
        [System.Drawing.Color]$Normal,
        [System.Drawing.Color]$Hover
    )
    if (-not $Button) { return }
    $Button.BackColor = $Normal
    $Button.FlatStyle = 'Flat'
    $Button.FlatAppearance.BorderSize = 1
    $Button.FlatAppearance.BorderColor = $theme.Border
    $Button.FlatAppearance.MouseDownBackColor = $Hover
    $Button.FlatAppearance.MouseOverBackColor = $Hover
    $Button.Cursor = [System.Windows.Forms.Cursors]::Hand
    $Button.Add_MouseEnter({ $this.BackColor = $Hover })
    $Button.Add_MouseLeave({ $this.BackColor = $Normal })
}

function Apply-ModernTheme {
    param([System.Windows.Forms.Control]$Root)

    foreach ($control in $Root.Controls) {
        switch ($control.GetType().Name) {
            'GroupBox' {
                $control.BackColor = $theme.Surface
                $control.ForeColor = $theme.Text
                $control.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 9.5)
            }
            'Label' {
                if ($control.ForeColor -eq [System.Drawing.Color]::White) {
                    $control.ForeColor = $theme.Text
                }
                if ($control.ForeColor -eq [System.Drawing.Color]::Silver) {
                    $control.ForeColor = $theme.Muted
                }
            }
            'TextBox' {
                $control.BackColor = $theme.Input
                $control.ForeColor = $theme.Text
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
                if ($control.BackColor -eq [System.Drawing.Color]::FromArgb(30, 30, 30)) {
                    $control.BackColor = $theme.Input
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

        if ($control.HasChildren) {
            Apply-ModernTheme -Root $control
        }
    }
}

# Primary action emphasis
Set-ButtonHover -Button $btnCreateVM -Normal $theme.Accent -Hover $theme.AccentHover
Set-ButtonHover -Button $btnUpdateGPU -Normal $theme.Success -Hover $theme.SuccessHover
Set-ButtonHover -Button $btnExit -Normal $theme.Danger -Hover $theme.DangerHover
Set-ButtonHover -Button $btnClearLog -Normal $theme.Surface -Hover $theme.Border
Set-ButtonHover -Button $btnSaveLog -Normal $theme.Surface -Hover $theme.Border
Set-ButtonHover -Button $btnBrowseVM -Normal $theme.Surface -Hover $theme.Border
Set-ButtonHover -Button $btnBrowseISO -Normal $theme.Surface -Hover $theme.Border
Set-ButtonHover -Button $btnBrowseGolden -Normal $theme.Surface -Hover $theme.Border
Set-ButtonHover -Button $btnSelectAll -Normal $theme.Surface -Hover $theme.Border
Set-ButtonHover -Button $btnSelectNone -Normal $theme.Surface -Hover $theme.Border
Set-ButtonHover -Button $btnRefreshVMs -Normal $theme.Surface -Hover $theme.Border
Set-ButtonHover -Button $btnClearVmSearch -Normal $theme.Surface -Hover $theme.Border

Apply-ModernTheme -Root $form

# Keyboard-first behavior: Enter runs primary action for active tab
$form.CancelButton = $btnExit
$form.AcceptButton = $btnCreateVM
$tabControl.Add_SelectedIndexChanged({
    if ($tabControl.SelectedTab -eq $tabCreate) {
        $form.AcceptButton = $btnCreateVM
    } elseif ($tabControl.SelectedTab -eq $tabGPU) {
        $form.AcceptButton = $btnUpdateGPU
    }
})

# Contextual tooltips for intuitiveness
$toolTip = New-Object System.Windows.Forms.ToolTip
$toolTip.AutoPopDelay = 8000
$toolTip.InitialDelay = 300
$toolTip.ReshowDelay = 120
$toolTip.ShowAlways = $true
$toolTip.SetToolTip($ctrlCreate["VMName"], "Use a short unique VM name (letters, numbers, -, _).")
$toolTip.SetToolTip($ctrlCreate["ISOPath"], "Windows installation ISO for unattended deployment mode.")
$toolTip.SetToolTip($ctrlCreate["GoldenParentVHD"], "Parent image used when Golden mode is enabled (differencing disk).")
$toolTip.SetToolTip($ctrlCreate["CheckpointMode"], "Production/ProductionOnly are recommended over Standard for stable rollback.")
$toolTip.SetToolTip($ctrlCreate["DynamicMem"], "Enable memory ballooning. Configure min/startup/max below.")
$toolTip.SetToolTip($ctrlCreate["DynamicMemMin"], "Lowest RAM the VM can shrink to when Dynamic Memory is enabled.")
$toolTip.SetToolTip($ctrlCreate["DynamicMemMax"], "Highest RAM the VM can grow to when Dynamic Memory is enabled.")
$toolTip.SetToolTip($ctrlCreate["StrictLegacyMode"], "Forces legacy-safe deployment behavior (non-compact DISM + legacy template order) for older/custom Windows 10 images.")
$toolTip.SetToolTip($ctrlCreate["AutoCreateSwitch"], "Automatically create an internal NAT switch if selected switch is missing.")
$toolTip.SetToolTip($ctrlCreate["EnableMetering"], "Collect CPU, memory, network, and disk telemetry for this VM.")
$toolTip.SetToolTip($ctrlGPU["GpuAllocSlider"], "Controls GPU partition resource share assigned to the VM.")
$toolTip.SetToolTip($btnUpdateGPU, "Inject/update GPU-P drivers and optional services into selected VMs.")
$toolTip.SetToolTip($ctrlGPU["VmSearch"], "Type to quickly filter VMs by name.")

#endregion

#region ==================== EVENT HANDLERS ====================

# ---- Browse VM Location ----
$btnBrowseVM.Add_Click({
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    $dlg.Description = "Select the folder where VMs will be stored"
    if ($dlg.ShowDialog() -eq 'OK') { $ctrlCreate["VMLocation"].Text = $dlg.SelectedPath }
})

# ---- Browse ISO + Detect Editions + Detect Version ----
$btnBrowseISO.Add_Click({
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Filter = "ISO Files|*.iso"
    $dlg.Title  = "Select a Windows Installation ISO"
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
            $script:MountedISO = Mount-DiskImage -ImagePath $dlg.FileName -PassThru
            Register-TrackedMountedImage -ImagePath $dlg.FileName
            Start-Sleep -Seconds 2
            $isoDrive = ($script:MountedISO | Get-DiskImage | Get-Volume | Where-Object DriveLetter).DriveLetter + ":"

            # Find WIM or ESD
            $script:WimFile = Join-Path "$isoDrive\sources" "install.wim"
            if (-not (Test-Path $script:WimFile)) { $script:WimFile = Join-Path "$isoDrive\sources" "install.esd" }
            if (-not (Test-Path $script:WimFile)) {
                Write-Log "Cannot find install.wim or install.esd in ISO" "ERROR"
                return
            }

            # Parse editions
            $wimInfo = dism /Get-WimInfo /WimFile:"$($script:WimFile)" /English 2>&1
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
            Write-Log "Found $($editions.Count) edition(s): $($editions -join ', ')" "OK"

            # Detect Windows version from first edition
            if ($editions.Count -gt 0) {
                $firstIndex = $script:EditionMap[$editions[0]]
                $versionInfo = Get-WimVersionInfo -WimFile $script:WimFile -Index $firstIndex
                $script:DetectedWinVersion = $versionInfo.WinVersion
                $script:DetectedBuild      = $versionInfo.Build

                $detectedProfile = Resolve-GuestWindowsProfile -DetectedWinVersion $script:DetectedWinVersion -DetectedBuild $script:DetectedBuild
                Apply-DetectedGuestDefaults -Controls $ctrlCreate -Profile $detectedProfile -EmitLog

                # Auto-suggest VM name from ISO filename
                if ([string]::IsNullOrWhiteSpace($ctrlCreate["VMName"].Text)) {
                    $suggestedName = [IO.Path]::GetFileNameWithoutExtension($dlg.FileName) -replace '[^a-zA-Z0-9_-]', ''
                    if ($suggestedName.Length -gt 15) { $suggestedName = $suggestedName.Substring(0, 15) }
                    $ctrlCreate["VMName"].Text = $suggestedName
                }
            }

        } catch {
            Write-Log "Failed to read ISO: $_" "ERROR"
        }
    }
})

$btnBrowseGolden.Add_Click({
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Filter = "Virtual Hard Disk|*.vhdx;*.vhd"
    $dlg.Title  = "Select parent Golden VHDX/VHD"
    if ($dlg.ShowDialog() -eq 'OK') {
        $ctrlCreate["GoldenParentVHD"].Text = $dlg.FileName
    }
})

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

    if ($ctrlCreate.ContainsKey("ModeHint") -and $ctrlCreate["ModeHint"]) {
        if ($isGoldenMode) {
            $ctrlCreate["ModeHint"].Text = "Mode: Golden Image - Uses parent VHDX differencing disk. ISO and edition are ignored."
        } else {
            $ctrlCreate["ModeHint"].Text = "Mode: ISO Deploy - Uses ISO, selected edition, and unattended setup."
        }
    }
}

$ctrlCreate["GoldenImage"].Add_CheckedChanged({
    Update-CreateModeUi
})

Update-CreateModeUi

# ---- Update OS info when edition changes ----
$ctrlCreate["Edition"].Add_SelectedIndexChanged({
    $selectedEdition = $ctrlCreate["Edition"].SelectedItem
    if ($selectedEdition -and $script:EditionMap.ContainsKey($selectedEdition)) {
        $idx = $script:EditionMap[$selectedEdition]
        $versionInfo = Get-WimVersionInfo -WimFile $script:WimFile -Index $idx
        $script:DetectedWinVersion = $versionInfo.WinVersion
        $script:DetectedBuild      = $versionInfo.Build
        $detectedProfile = Resolve-GuestWindowsProfile -DetectedWinVersion $script:DetectedWinVersion -DetectedBuild $script:DetectedBuild
        Apply-DetectedGuestDefaults -Controls $ctrlCreate -Profile $detectedProfile
    }
})

$ctrlCreate["Memory"].Add_ValueChanged({
    if (-not $ctrlCreate.ContainsKey("DynamicMemMin") -or -not $ctrlCreate.ContainsKey("DynamicMemMax")) { return }
    $startup = [int]$ctrlCreate["Memory"].Value
    if ([int]$ctrlCreate["DynamicMemMin"].Value -gt $startup) {
        $ctrlCreate["DynamicMemMin"].Value = $startup
    }
    if ([int]$ctrlCreate["DynamicMemMax"].Value -lt $startup) {
        $ctrlCreate["DynamicMemMax"].Value = $startup
    }
})

# ================================================================
#  CREATE VM - Main Logic
# ================================================================
$btnCreateVM.Add_Click({
    $VMName = ""
    $VMLoc = ""
    $VHDPath = ""
    $preflightLines = @()
    $rollbackNeeded = $false
    $originalAutoPlay = $null
    $autoPlayChanged = $false

    try {
        Update-CreateProgress -Percent 2 -Status "Validating VM settings..."

        # ---- Gather inputs ----
        $VMName          = $ctrlCreate["VMName"].Text.Trim()
        $VMLocBase       = $ctrlCreate["VMLocation"].Text.Trim()
        $ISOPath         = $ctrlCreate["ISOPath"].Text.Trim()
        $Username        = $ctrlCreate["Username"].Text.Trim()
        $PasswordText    = $ctrlCreate["Password"].Text
        $Password        = Convert-PlainTextToSecureString -Text $PasswordText
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

        if ($StrictLegacyMode) {
            if ($IsWin11) {
                Write-Log "Strict Legacy Mode is ignored for Windows 11 compatibility requirements." "WARN"
            } else {
                $guestProfile.IsLegacyWindows10 = $true
                $guestProfile.PreferCompactApply = $false
                $guestProfile.SecureBootTemplateOrder = @('MicrosoftUEFICertificateAuthority', 'MicrosoftWindows')
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

        # Parse resolution
        $ResWidth = 1920; $ResHeight = 1080
        $raw = $ctrlCreate["Resolution"].SelectedItem
        if ($raw) {
            $m = [regex]::Match([string]$raw, '^(\d+)\s*[xX×]\s*(\d+)$')
            if ($m.Success) { $ResWidth = [int]$m.Groups[1].Value; $ResHeight = [int]$m.Groups[2].Value }
        }

        # ---- Validate ----
        if ([string]::IsNullOrWhiteSpace($VMName)) { Write-Log "VM Name is required!" "ERROR"; return }
        if ($VMName -notmatch '^[a-zA-Z0-9_-]+$') { Write-Log "VM Name contains invalid characters. Use only letters, numbers, hyphens, underscores." "ERROR"; return }
        if ($VMName.Length -gt 15) { Write-Log "VM Name must be 15 characters or less for NetBIOS compatibility" "ERROR"; return }
        if (Get-VM -Name $VMName -ErrorAction SilentlyContinue) { Write-Log "A VM named '$VMName' already exists!" "ERROR"; return }

        if (-not $UseGoldenImage) {
            if ([string]::IsNullOrWhiteSpace($ISOPath) -or -not (Test-Path $ISOPath)) { Write-Log "Valid ISO file is required." "ERROR"; return }
        } else {
            if ([string]::IsNullOrWhiteSpace($GoldenParentVHD) -or -not (Test-Path $GoldenParentVHD)) {
                Write-Log "Golden Image mode is enabled. Please select a valid parent VHDX/VHD." "ERROR"
                return
            }
        }

        if (-not (Test-Path $VMLocBase)) {
            Write-Log "VM location does not exist. Creating folder: $VMLocBase" "WARN"
            try {
                New-Item -Path $VMLocBase -ItemType Directory -Force -ErrorAction Stop | Out-Null
            } catch {
                Write-Log "Failed to create VM location: $($PSItem.Exception.Message)" "ERROR"
                return
            }
        }
        if (-not (Test-DirectoryWritable -Path $VMLocBase)) { Write-Log "VM Location is not writable: $VMLocBase" "ERROR"; return }

        $requiredDiskGB = $DiskGB + 8
        $freeDiskGB = Get-PathAvailableSpaceGB -Path $VMLocBase
        if ($freeDiskGB -ge 0 -and $freeDiskGB -lt $requiredDiskGB) {
            Write-Log "Insufficient disk space in VM location. Need about $requiredDiskGB GB free, found $freeDiskGB GB." "ERROR"
            return
        }

        if (-not $UseGoldenImage) {
            if (-not $script:WimFile -or -not (Test-Path $script:WimFile)) { Write-Log "No WIM/ESD file found. Please select an ISO first." "ERROR"; return }
        }
        if ([string]::IsNullOrWhiteSpace($Username)) { Write-Log "Username is required!" "ERROR"; return }
        if ($Username -notmatch '^[a-zA-Z0-9]+$') { Write-Log "Username cannot contain special characters." "ERROR"; return }
        if ($Username -eq $VMName) { Write-Log "Username cannot be the same as VM Name (causes admin permission issues in the VM)." "ERROR"; return }
        if (-not [string]::IsNullOrWhiteSpace($PasswordText) -and $PasswordText.Length -lt 8) {
            Write-Log "Password must be at least 8 characters." "ERROR"
            return
        }
        if ($EnableDynamicMem) {
            if ($DynamicMemMinGB -gt $DynamicMemMaxGB) {
                Write-Log "Dynamic Memory minimum cannot be greater than maximum." "ERROR"
                return
            }
            if ($DynamicMemMinGB -gt $MemGB) {
                Write-Log "Dynamic Memory minimum cannot be greater than startup memory." "ERROR"
                return
            }
            if ($DynamicMemMaxGB -lt $MemGB) {
                Write-Log "Dynamic Memory maximum cannot be less than startup memory." "ERROR"
                return
            }
        }

        if ($VMSwitch -and (Get-VMSwitch -Name $VMSwitch -ErrorAction SilentlyContinue)) {
            # selected switch is valid
        } else {
            if (-not $AutoCreateSwitch) {
                if (-not $VMSwitch) { Write-Log "No Virtual Switch selected!" "ERROR" }
                else { Write-Log "Selected Virtual Switch no longer exists: $VMSwitch" "ERROR" }
                return
            }

            $createdSwitch = Ensure-ToolkitNatSwitch
            if (-not $createdSwitch) {
                Write-Log "Could not auto-create a fallback virtual switch." "ERROR"
                return
            }

            $VMSwitch = $createdSwitch
            $ctrlCreate["Switch"].Items.Clear()
            Get-VMSwitch | Select-Object -ExpandProperty Name | ForEach-Object { [void]$ctrlCreate["Switch"].Items.Add($_) }
            $ctrlCreate["Switch"].SelectedItem = $VMSwitch
        }

        if (-not $UseGoldenImage -and ($null -eq $SelectedIndex -or $SelectedIndex -lt 1)) { Write-Log "No Windows Edition selected!" "ERROR"; return }

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

        # ---- Host/VM version match check ----
        if ($script:DetectedBuild -gt 0 -and $script:HostBuild -gt 0) {
            if ([math]::Abs($script:DetectedBuild - $script:HostBuild) -gt 5000) {
                Write-Log "WARNING: Large version gap between host (Build $($script:HostBuild)) and VM image (Build $($script:DetectedBuild))." "WARN"
                Write-Log "Mismatched Windows versions can cause GPU-P driver issues or BSODs. Matching versions recommended." "WARN"
            }
        }

        # ---- Disable UI ----
        $tabControl.Enabled = $false
        $btnCreateVM.Enabled = $false
        Update-CreateProgress -Percent 6 -Status "Preparing VM creation workflow..."
        Write-Log "========================================" "INFO"
        Write-Log "Starting VM creation: $VMName" "INFO"
        Write-Log "OS: $($script:DetectedWinVersion) Build $($script:DetectedBuild)" "INFO"
        Write-Log "Secure Boot: $EnableSecureBoot | TPM: $EnableTPM" "INFO"
        Write-Log "========================================" "INFO"

        # ---- Disable AutoPlay ----
        $originalAutoPlay = Set-AutoPlay -Disable $true
        $autoPlayChanged = $true

        # ---- Create VM directory ----
        Update-CreateProgress -Percent 10 -Status "Creating VM workspace..."
        $VMLoc = Join-Path $VMLocBase $VMName
        if (Test-Path $VMLoc) { Write-Log "Directory $VMLoc already exists; files may be overwritten" "WARN" }
        else { New-Item -Path $VMLoc -ItemType Directory -Force | Out-Null }
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
            Update-CreateProgress -Percent 15 -Status "Preparing setup tools..."
            Write-Log "Creating QRes.exe..."
            $tempExe = Join-Path $VMLoc "QRes.exe"
        $base64 = @"
TVqQAAMAAAAEAAAA//8AALgAAAAAAAAAQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA0AAAAA4fug4AtAnNIbgBTM0hVGhpcyBwcm9ncmFtIGNhbm5vdCBiZSBydW4gaW4gRE9TIG1vZGUuDQ0KJAAAAAAAAACDEdcDx3C5UMdwuVDHcLlQRGy3UMZwuVAvb71QxXC5UMdwuFDZcLlQpW+qUM5wuVAvb7NQy3C5UFJpY2jHcLlQAAAAAAAAAAAAAAAAAAAAAFBFAABMAQEASP76PgAAAAAAAAAA4AAPAQsBBgAAAAAAABAAAAAAAABIGwAAABAAAAAQAAAAAEAAABAAAAACAAAEAAAAAAAAAAQAAAAAAAAAACAAAAACAAD2EAEAAwAAAAAAEAAAEAAAAAAQAAAQAAAAAAAAEAAAAAAAAAAAAAAAsBwAAHgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAQAACEAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAALmRhdGEAAACKDwAAABAAAAAQAAAAAgAAAAAAAAAAAAAAAAAAQAAAwAAAAAAAAAAAAAAAAAAAAABqHgAAeB4AAIweAAAAAAAAUB4AAAAAAADSHQAAxB0AALgdAACsHQAAAAAAAKoeAAD4HgAACB8AAHwfAABoHwAAtB4AAMoeAADSHgAA4B4AAOgeAABWHwAAFB8AACgfAAA4HwAASB8AAAAAAAAYHgAA8h0AAP4dAAAkHgAALB4AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAARXJyb3I6ICVzCgAAICAlcwklcwoAAAAACSAlcwoAAAAlcy4KAAAAACBAICVkIEh6AAAAAEFkYXB0ZXIgRGVmYXVsdABPcHRpbWFsAHVua25vd24AJWR4JWQsICVkIGJpdHMAACBAIAAKRXg6ICJRUmVzLmV4ZSAveDo2NDAgL2M6OCIgQ2hhbmdlcyByZXNvbHV0aW9uIHRvIDY0MCB4IDQ4MCBhbmQgdGhlIGNvbG9yIGRlcHRoIHRvIDI1NiBjb2xvcnMuCgAvSAAARGlzcGxheXMgbW9yZSBoZWxwLgAvPwAARGlzcGxheXMgdXNhZ2UgaW5mb3JtYXRpb24uAC9WAABEb2VzIE5PVCBkaXNwbGF5IHZlcnNpb24gaW5mb3JtYXRpb24uAAAAL0QAAERvZXMgTk9UIHNhdmUgZGlzcGxheSBzZXR0aW5ncyBpbiB0aGUgcmVnaXN0cnkuLgAAAAAvTAAATGlzdCBhbGwgZGlzcGxheSBtb2Rlcy4AL1MAAFNob3cgY3VycmVudCBkaXNwbGF5IHNldHRpbmdzLgAALTE9IE9wdGltYWwuAAAAADAgPSBBZGFwdGVyIERlZmF1bHQuAAAAAC9SAABSZWZyZXNoIHJhdGUuAAAAMzI9IFRydWUgY29sb3IuADI0PSBUcnVlIGNvbG9yLgAxNj0gSGlnaCBjb2xvci4AOCA9IDI1NiBjb2xvcnMuADQgPSAxNiBjb2xvcnMuAAAvQwAAQ29sb3IgZGVwdGguAAAAAC9ZAABIZWlnaHQgaW4gcGl4ZWxzLgAAAC9YAABXaWR0aCBpbiBwaXhlbHMuAAAAAFFSRVMgWy9YOltweF1dIFsvWTpbcHhdXSBbL0M6W2JpdHNdIFsvUjpbcnJdXSBbL1NdIFsvTF0gWy9EXSBbL1ZdIFsvP10gWy9IXQoKAAAAU2V0dGluZ3MgY291bGQgbm90IGJlIHNhdmVkLGdyYXBoaWNzIG1vZGUgd2lsbCBiZSBjaGFuZ2VkIGR5bmFtaWNhbGx5Li4uAAAAAFRoZSBjb21wdXRlciBtdXN0IGJlIHJlc3RhcnRlZCBpbiBvcmRlciBmb3IgdGhlIGdyYXBoaWNzIG1vZGUgdG8gd29yay4uLgAAAABUaGUgZ3JhcGhpY3MgbW9kZSBpcyBub3Qgc3VwcG9ydGVkIQBNb2RlIE9rLi4uCgBSZWZyZXNoUmF0ZQBEaXNwbGF5XFNldHRpbmdzAAAAAFFSZXMgdjEuMQpDb3B5cmlnaHQgKEMpIEFuZGVycyBLamVyc2VtLgoKAAAAAQAAAAAAAAD/////OBxAAEwcQAAAAAAA/3QkBGigEEAA/xUsEEAAWTPAWcP/dCQI/3QkCGisEEAA/xUsEEAAg8QMw/90JARouBBAAP8VLBBAAFlZw4tMJARWM/YzwIA5LXUEagFeQYoRgPowfBGA+jl/DA++0o0EgI1EQtDr54X2XnQC99jDi0QkBIA4AHQBQIoIgPk6dAiA+SB0AzPAw0BQ6K7///9Zi0wkCGoBiQFYw1WL7IPsZFaLdQihBBFAAFdqGIlFnFkzwP92aI19oPOr/3Zwiz0sEEAA/3ZsaPQQQAD/14PEEIM9oBxAAAB1NItGeIXAdgeD+P91KIXAdB2D+P90B2jsEEAA6wVo5BBAAI1FnFD/FSAQQADrHGjUEEAA6+3/dniNRZxoyBBAAFD/FXAQQACDxAz2RQwBdA+NRZxojBxAAFD/FSQQQACNRZxQaMAQQAD/11lZX17Jw1WL7IHsvAAAAFNWM8BXiUX8iUX4iUX0iUXsx0Xw/v////8VGBBAAIvw/xUcEEAAPQAAAIAbwPfYo6AcQACKBjwidQ6KRgFGhMB0FDwidBDr8oTAdAo8IHQGikYBRuvygD4AdAFGgD4gdPpqBF9qAluKBjwvdAg8LQ+FNwEAAITAD4QvAQAAD75GAUaD+Fl/RA+EygAAAIP4TH8XdFKD6D90WSvHdGdIdFsrx3RL6fcAAACD6FJ0eUgPhOcAAACD6AMPhNcAAAArww+EsAAAAOnVAAAAg/hyf3Z0VYPoY3QtSHQhK8d0ESvHD4W6AAAAg038IOmxAAAACX38Rgld/OmlAAAAg038COmcAAAAjUXsUFboEP7//1mFwFl0AgPzigY8IA+EgAAAAITAdHxG6++NRfBQVujt/f//WYXAWXQCA/OKBjwgdGGEwHRdRuvzg+hzdFGD6AN0RSvDdCJIdUmNRfRQVui9/f//WYXAWXQCA/OKBjwgdDGEwHQtRuvzjUX4UFbonv3//1mFwFl0AgPzigY8IHQShMB0Dkbr80aDTfwB6wSDTfwQgD4gD4W+/v//Ruv09kX8AYsdLBBAAHUIaEwUQAD/01n2RfwgdFOLNXwQQABqAV+NhUT///9QM9tXU//WhcAPhHUDAABHg728AXQjgX2wgAIAAHIaM8A5HaAcQAAPlMBQjYVE////UOg9/f//WVmNhUT///9QV1PrwfZF/BAPhMoAAABqAP8VeBBAAIv4hf8PhCQDAACLNRAQQABqCFf/1moKiUWwW1NX/9ZqDFeJRbT/1mp0V4lFrP/WhcCJRbx1bjP2OTWgHEAAdWaNRehQaBkAAgBWaDgUQABoBQAAgP8VCBBAAIXAdUiNReSJXeRQjUXYUI1F/FBWaCwUQAD/dej/FQQQQACFwHUZg338AXQGg338AnUNjUXYUOgs/P//WYlFvP916P8VABBAAOsCM/aNhUT///9WUOhr/P//WVlXVv8VbBBAAOlsAgAA9kX8Ag+FTgEAAIN9+ACLRfR1E4XAdQ85Rex1D4N98P4PhDIBAACD+AF9DotF+Jn3/40EQIlF9OsSg334AX0MweACagOZWff5iUX4aJQAAACNhUT///9qAFDoFgIAAItF9ItN+ItV8IlFtItF7IPEDIXAZseFaP///5QAiU2wiUWsiVW8fgqBhWz///8AAAQAhcl+CoGFbP///wAAGACD+v6+AABAAHU4gz2gHEAAAHQ1agD/FXgQQACL+IX/dBsBtWz///9qdFf/FRAQQABXagCJRbz/FWwQQACDffD+dAYBtWz///+LNXQQQACNhUT///9qAlD/1ov4hf91IItF/PfQwegDg+ABUI2FRP///1D/1mggFEAAi/j/0+sWi8dIdAdo/BNAAOsFaLATQADoj/r//4P//VkPhS8BAABoZBNAAOh7+v//WY2FRP///2oAUP/W6RQBAABoFBNAAP/TxwQkABNAAGj8EkAA6Gb6//9o6BJAAGjkEkAA6Ff6//9o1BJAAGjQEkAA6Ej6//+LdfyDxBgj93Q7aMASQADoS/r//8cEJLASQADoP/r//8cEJKASQADoM/r//8cEJJASQADoJ/r//8cEJIASQADoG/r//1locBJAAGhsEkAA6PT5//+DPaAcQAAAWVl1G4X2dBdoVBJAAOjy+f//xwQkRBJAAOjm+f//WWgkEkAAaCASQADov/n//2gIEkAAaAQSQADosPn//2jQEUAAaMwRQADoofn//2ikEUAAaKARQADokvn//2iEEUAAaIARQADog/n//2hsEUAAaGgRQADodPn//2gIEUAA/9ODxDRfXjPAW8nDzP8lQBBAAFWL7Gr/aIAUQABogBxAAGShAAAAAFBkiSUAAAAAg+wgU1ZXiWXog2X8AGoB/xVUEEAAWYMNpBxAAP+DDagcQAD//xVkEEAAiw2cHEAAiQj/FWAQQACLDZgcQACJCKFcEEAAiwCjrBxAAOjDAAAAgz14FEAAAHUMaHYcQAD/FVgQQABZ6JQAAABokBBAAGiMEEAA6H8AAAChlBxAAIlF2I1F2FD/NZAcQACNReBQjUXUUI1F5FD/FTAQQABoiBBAAGiEEEAA6EwAAAD/FVAQQACLTeCJCP914P911P915Oit+f//g8QwiUXcUP8VTBBAAItF7IsIiwmJTdBQUegPAAAAWVnDi2Xo/3XQ/xVEEEAA/yVIEEAA/yU0EEAAaAAAAwBoAAABAOgTAAAAWVnDM8DDw8zMzMzMzP8lPBBAAP8lOBBAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAEAdAAAAAAAAAAAAAOQdAAAYEAAAlB0AAAAAAAAAAAAARB4AAGwQAAA4HQAAAAAAAAAAAABgHgAAEBAAACgdAAAAAAAAAAAAAJweAAAAEAAAVB0AAAAAAAAAAAAAvh4AACwQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAGoeAAB4HgAAjB4AAAAAAABQHgAAAAAAANIdAADEHQAAuB0AAKwdAAAAAAAAqh4AAPgeAAAIHwAAfB8AAGgfAAC0HgAAyh4AANIeAADgHgAA6B4AAFYfAAAUHwAAKB8AADgfAABIHwAAAAAAABgeAADyHQAA/h0AACQeAAAsHgAAAAAAAAIDbHN0cmNweUEAAPkCbHN0cmNhdEEAAHQBR2V0VmVyc2lvbgAAygBHZXRDb21tYW5kTGluZUEAS0VSTkVMMzIuZGxsAACsAndzcHJpbnRmQQAbAENoYW5nZURpc3BsYXlTZXR0aW5nc0EAAAMCUmVsZWFzZURDAP0AR2V0REMAxQBFbnVtRGlzcGxheVNldHRpbmdzQQAAVVNFUjMyLmRsbAAAJQFHZXREZXZpY2VDYXBzAEdESTMyLmRsbABbAVJlZ0Nsb3NlS2V5AHsBUmVnUXVlcnlWYWx1ZUV4QQAAcgFSZWdPcGVuS2V5RXhBAEFEVkFQSTMyLmRsbAAAngJwcmludGYAAJkCbWVtc2V0AABNU1ZDUlQuZGxsAADTAF9leGl0AEgAX1hjcHRGaWx0ZXIASQJleGl0AABkAF9fcF9fX2luaXRlbnYAWABfX2dldG1haW5hcmdzAA8BX2luaXR0ZXJtAIMAX19zZXR1c2VybWF0aGVycgAAnQBfYWRqdXN0X2ZkaXYAAGoAX19wX19jb21tb2RlAABvAF9fcF9fZm1vZGUAAIEAX19zZXRfYXBwX3R5cGUAAMoAX2V4Y2VwdF9oYW5kbGVyMwAAtwBfY29udHJvbGZwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
"@
            [IO.File]::WriteAllBytes($tempExe, [Convert]::FromBase64String($base64))

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

        if ($ctrlCreate["Parsec"].Checked) {
            $lines += @(
                ':: --- Parsec ---'
                'echo [%date% %time%] Downloading Parsec... >> %LOGFILE%'
                'bitsadmin /transfer "DownloadParsec" https://builds.parsecgaming.com/package/parsec-windows.exe %WORKDIR%\parsec.exe >> %LOGFILE% 2>&1'
                'echo [%date% %time%] Installing Parsec... >> %LOGFILE%'
                'start /wait %WORKDIR%\parsec.exe /silent /percomputer /norun /vdd'
                ''
            )
        }
        if ($ctrlCreate["VBCable"].Checked) {
            $lines += @(
                ':: --- VB Cable ---'
                'echo [%date% %time%] Downloading VB Cable... >> %LOGFILE%'
                'bitsadmin /transfer "DownloadVBCable" https://download.vb-audio.com/Download_CABLE/VBCABLE_Driver_Pack45.zip %WORKDIR%\vb.zip >> %LOGFILE% 2>&1'
                'if not exist "%WORKDIR%\VB" mkdir "%WORKDIR%\VB"'
                'tar -xf %WORKDIR%\vb.zip -C %WORKDIR%\VB >> %LOGFILE% 2>&1'
                'start /wait %WORKDIR%\VB\VBCABLE_Setup_x64 -h -i -H -n'
                ''
            )
        }
        if ($ctrlCreate["USBMMIDD"].Checked) {
            $lines += @(
                ':: --- Virtual Display Driver ---'
                'echo [%date% %time%] Downloading USBMMIDD... >> %LOGFILE%'
                'bitsadmin /transfer "DownloadUSBMMIDD" https://www.amyuni.com/downloads/usbmmidd_v2.zip %WORKDIR%\usbmmidd_v2.zip >> %LOGFILE% 2>&1'
                'if not exist "%WORKDIR%\usbmmidd_v2" mkdir "%WORKDIR%\usbmmidd_v2"'
                'tar -xf %WORKDIR%\usbmmidd_v2.zip -C %WORKDIR% >> %LOGFILE% 2>&1'
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
                'echo icacls "%%SHAREFOLDER%%" /grant Everyone:(OI)(CI)F /T >> C:\Windows\Temp\CreateShare.cmd'
                'echo powershell -Command "Set-NetFirewallRule -DisplayGroup ''File and Printer Sharing'' -Enabled True" >> C:\Windows\Temp\CreateShare.cmd'
                'echo powershell -Command "if (Get-SmbShare -Name ''share'' -ErrorAction SilentlyContinue) { Remove-SmbShare -Name ''share'' -Force }" >> C:\Windows\Temp\CreateShare.cmd'
                'echo powershell -Command "New-SmbShare -Name ''share'' -Path \"${env:USERPROFILE}\Desktop\share\" -FullAccess ''Everyone''" >> C:\Windows\Temp\CreateShare.cmd'
                'echo exit >> C:\Windows\Temp\CreateShare.cmd'
                ''
            )
        }
        if ($ctrlCreate["PauseUpdate"].Checked) {
            $lines += @(
                ':: --- Pause Windows Updates 1 year ---'
                'powershell -ExecutionPolicy Bypass -NoProfile -Command "$now = (Get-Date).ToString(''yyyy-MM-ddTHH:mm:ssZ''); $future = (Get-Date).AddDays(365).ToString(''yyyy-MM-ddTHH:mm:ssZ''); $wuPath = ''HKLM:\SOFTWARE\Microsoft\WindowsUpdate\UX\Settings''; if (-not (Test-Path $wuPath)) { New-Item -Path $wuPath -Force | Out-Null }; Set-ItemProperty -Path $wuPath -Name ''PauseFeatureUpdatesStartTime'' -Value $now; Set-ItemProperty -Path $wuPath -Name ''PauseFeatureUpdatesEndTime'' -Value $future; Set-ItemProperty -Path $wuPath -Name ''PauseFeatureUpdates'' -Value 1 -Type DWord; Set-ItemProperty -Path $wuPath -Name ''PauseQualityUpdatesStartTime'' -Value $now; Set-ItemProperty -Path $wuPath -Name ''PauseQualityUpdatesEndTime'' -Value $future; Set-ItemProperty -Path $wuPath -Name ''PauseQualityUpdates'' -Value 1 -Type DWord"'
                ''
            )
        }
        if ($ctrlCreate["FullUpdate"].Checked) {
            $lines += @(
                ':: --- Full Windows Updates at first logon ---'
                'reg add "HKLM\Software\Microsoft\Windows\CurrentVersion\RunOnce" /v RunUpdates /d "cmd /c C:\Windows\Temp\RunUpdates.cmd" /f'
                'echo @echo off > C:\Windows\Temp\RunUpdates.cmd'
                'echo powershell -ExecutionPolicy Bypass -NoProfile -Command "try { Install-PackageProvider -Name NuGet -Force; [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; Install-Module PSWindowsUpdate -Force -Scope AllUsers; Import-Module PSWindowsUpdate; Add-WUServiceManager -MicrosoftUpdate -Confirm:$false -ErrorAction SilentlyContinue; Get-WindowsUpdate -MicrosoftUpdate -Install -AcceptAll -IgnoreReboot | Out-File -FilePath C:\\Windows\\Temp\\WUOutput.log -Encoding UTF8; } catch { Write-Output $_.Exception.Message }" >> C:\Windows\Temp\RunUpdates.cmd'
                'echo shutdown /r /t 30 /c "Windows Updates complete. Rebooting in 30 seconds." >> C:\Windows\Temp\RunUpdates.cmd'
                'echo exit >> C:\Windows\Temp\RunUpdates.cmd'
                ''
            )
        }
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
            New-UnattendXml -VMName $VMName -Username $Username -Password $Password `
                -ResWidth $ResWidth -ResHeight $ResHeight -IsWindows11 $IsWin11 |
                Out-File -FilePath $UnattendXMLPath -Encoding UTF8

            # Minimize plaintext password lifetime in memory
            $PasswordText = $null

            # ---- Create VHDX ----
            Update-CreateProgress -Percent 30 -Status "Creating virtual disk..."
            $VHDPath = Join-Path $VMLoc "$VMName.vhdx"
            if ($FixedVHD) {
                Write-Log "Creating fixed VHDX ($DiskGB GB)..."
                New-VHD -Path $VHDPath -SizeBytes ($DiskGB * 1GB) -Fixed | Out-Null
            } else {
                Write-Log "Creating dynamic VHDX ($DiskGB GB max)..."
                New-VHD -Path $VHDPath -SizeBytes ($DiskGB * 1GB) -Dynamic | Out-Null
            }

        # ---- Mount VHD and partition ----
        Update-CreateProgress -Percent 38 -Status "Partitioning virtual disk..."
        Write-Log "Mounting VHD and creating GPT/EFI/MSR/Windows partitions..."
        $mountedVHD = Mount-VHD -Path $VHDPath -Passthru
        Register-TrackedMountedImage -ImagePath $VHDPath
        $diskNumber = $mountedVHD.DiskNumber
        Initialize-Disk -Number $diskNumber -PartitionStyle GPT -PassThru | Out-Null

        # EFI System Partition (260MB for better compatibility with old Win10)
        $efi = New-Partition -DiskNumber $diskNumber -Size 260MB -GptType "{C12A7328-F81F-11D2-BA4B-00A0C93EC93B}" -AssignDriveLetter
        Format-Volume -Partition $efi -FileSystem FAT32 -NewFileSystemLabel "System" -Confirm:$false | Out-Null

        # MSR Partition
        New-Partition -DiskNumber $diskNumber -Size 16MB -GptType "{E3C9E316-0B5C-4DB8-817D-F92DF00215AE}" | Out-Null

        # Windows Partition (remaining space)
        $winPart = New-Partition -DiskNumber $diskNumber -UseMaximumSize -AssignDriveLetter
        $driveLetter = $winPart.DriveLetter + ":"
        Format-Volume -Partition $winPart -FileSystem NTFS -NewFileSystemLabel $VMName -Confirm:$false | Out-Null

        # Wait for drive to be ready
        $driveReady = $false
        for ($w = 0; $w -lt 15; $w++) {
            if (Test-Path "$driveLetter\") { $driveReady = $true; break }
            Start-Sleep -Seconds 1
        }
        if (-not $driveReady) { Write-Log "Windows partition drive not ready at $driveLetter" "WARN" }

        # ---- Apply Windows image ----
        Update-CreateProgress -Percent 48 -Status "Applying Windows image (this can take a while)..."
        Write-Log "Applying Windows image (Edition: $SelectedEdition, Index: $SelectedIndex)..."
        Write-Log "This may take several minutes..."
        Invoke-DismApplyImage -ImageFile $script:WimFile -Index $SelectedIndex -ApplyDir $driveLetter -PreferCompactApply $guestProfile.PreferCompactApply
        Write-Log "Windows image applied." "OK"

        # ---- Inject files into VHD ----
        # Unattend to root
        Copy-Item -Path $UnattendXMLPath -Destination (Join-Path "$driveLetter\" "Autounattend.xml") -Force
        # Unattend to Panther
        $PantherDir = Join-Path "$driveLetter\Windows" "Panther\Unattend"
        New-Item -Path $PantherDir -ItemType Directory -Force | Out-Null
        Copy-Item -Path $UnattendXMLPath -Destination (Join-Path $PantherDir "Unattend.xml") -Force
        # Also to Sysprep for maximum compatibility
        $SysprepDir = Join-Path "$driveLetter\Windows\System32" "Sysprep"
        if (Test-Path $SysprepDir) {
            Copy-Item -Path $UnattendXMLPath -Destination (Join-Path $SysprepDir "Unattend.xml") -Force
        }
        Write-Log "Autounattend.xml injected (root + Panther + Sysprep)" "OK"

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
        $efiDrive = $efi.DriveLetter + ":"
        Write-Log "Creating UEFI boot files on $efiDrive..."

        # Ensure EFI partition is accessible
        $efiReady = $false
        for ($w = 0; $w -lt 10; $w++) {
            if (Test-Path "$efiDrive\") { $efiReady = $true; break }
            Start-Sleep -Seconds 1
        }
        if (-not $efiReady) {
            Write-Log "EFI partition not accessible at $efiDrive, attempting drive letter reassignment..." "WARN"
            try {
                $efi | Set-Partition -NewDriveLetter $efi.DriveLetter
                Start-Sleep -Seconds 2
            } catch { Write-Log "Drive letter reassignment failed: $_" "WARN" }
        }

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
                        Write-Log "  Exit code $LASTEXITCODE : $($result | Out-String)" "WARN"
                    }
                } catch {
                    Write-Log "  Exception: $_" "WARN"
                }
            }
        }

        # Attempt 2: Use host's bcdboot
        if (-not $bootSuccess) {
            for ($retry = 1; $retry -le 3; $retry++) {
                try {
                    Write-Log "Boot attempt $retry/3: Using host bcdboot..."
                    Start-Sleep -Seconds 2
                    $result = & bcdboot "$driveLetter\Windows" /s $efiDrive /f UEFI 2>&1
                    if ($LASTEXITCODE -eq 0) {
                        Write-Log "Boot files created successfully (host bcdboot)" "OK"
                        $bootSuccess = $true
                        break
                    } else {
                        Write-Log "  Exit code $LASTEXITCODE : $($result | Out-String)" "WARN"
                    }
                } catch {
                    Write-Log "  Exception: $_" "WARN"
                }
            }
        }

        # Attempt 3: Host bcdboot with /f ALL for broader compatibility
        if (-not $bootSuccess) {
            for ($retry = 1; $retry -le 2; $retry++) {
                try {
                    Write-Log "Boot attempt $retry/2: Using host bcdboot with /f ALL fallback..."
                    Start-Sleep -Seconds 2
                    $result = & bcdboot "$driveLetter\Windows" /s $efiDrive /f ALL 2>&1
                    if ($LASTEXITCODE -eq 0) {
                        Write-Log "Boot files created successfully (host bcdboot /f ALL)" "OK"
                        $bootSuccess = $true
                        break
                    } else {
                        Write-Log "  Exit code $LASTEXITCODE : $($result | Out-String)" "WARN"
                    }
                } catch {
                    Write-Log "  Exception: $_" "WARN"
                }
            }
        }

            if (-not $bootSuccess) {
                Write-Log "Boot file creation failed! ISO will be attached for Windows Setup repair boot." "ERROR"
                $attachISOForRecovery = $true
            }

            # ---- Dismount VHD and ISO ----
            Update-CreateProgress -Percent 80 -Status "Finalizing disk images..."
            if (Dismount-ImageRetry -ImagePath $VHDPath) {
                Write-Log "VHD dismounted."
            } else {
                try {
                    Dismount-VHD -Path $VHDPath -ErrorAction SilentlyContinue
                    Unregister-TrackedMountedImage -ImagePath $VHDPath
                } catch {
                    Write-Log "VHD dismount fallback failed: $($_.Exception.Message)" "WARN"
                }
                Write-Log "VHD dismount required fallback and may still be attached." "WARN"
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
            Update-CreateProgress -Percent 30 -Status "Creating differencing disk from golden image..."
            $VHDPath = Join-Path $VMLoc "$VMName.vhdx"
            Write-Log "Creating differencing VHDX from parent: $GoldenParentVHD"
            New-VHD -Path $VHDPath -ParentPath $GoldenParentVHD -Differencing | Out-Null
            Write-Log "Golden image differencing disk created." "OK"
        }

        # ---- Create Hyper-V VM ----
        Update-CreateProgress -Percent 88 -Status "Creating Hyper-V VM..."
        Write-Log "Creating Generation 2 Hyper-V VM..."
        New-VM -Name $VMName -MemoryStartupBytes ($MemGB * 1GB) -Generation 2 -VHDPath $VHDPath -Path $VMLoc -SwitchName $VMSwitch | Out-Null

        # Processor
        Set-VM -Name $VMName -ProcessorCount $vCPU
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
        Set-VMGuestSecureBoot -VMName $VMName -EnableSecureBoot $EnableSecureBoot -GuestIsWindows11 $IsWin11 -GuestBuild $guestProfile.Build -TemplateOrder $guestProfile.SecureBootTemplateOrder | Out-Null

        # TPM
        if ($EnableTPM) {
            try {
                Set-VMKeyProtector -VMName $VMName -NewLocalKeyProtector -ErrorAction Stop
                Enable-VMTPM -VMName $VMName -ErrorAction Stop
                Write-Log "  Virtual TPM: Enabled" "OK"
            } catch {
                Write-Log "  TPM setup failed: $($_.Exception.Message)" "WARN"
            }
        }

        # Checkpoints
        Set-VM -Name $VMName -CheckpointType $CheckpointMode
        Write-Log "  Checkpoint Mode: $CheckpointMode"

        # Dynamic Memory
        if ($EnableDynamicMem) {
            Set-VMMemory -VMName $VMName -DynamicMemoryEnabled $true -StartupBytes ($MemGB * 1GB) -MinimumBytes ($DynamicMemMinGB * 1GB) -MaximumBytes ($DynamicMemMaxGB * 1GB)
        } else {
            Set-VMMemory -VMName $VMName -DynamicMemoryEnabled $false -StartupBytes ($MemGB * 1GB)
        }
        Write-Log "  Dynamic Memory: $(if ($EnableDynamicMem){"Enabled (min ${DynamicMemMinGB}GB, max ${DynamicMemMaxGB}GB)"}else{"Disabled"})"

        # Enhanced Session
        if ($EnableEnhancedSession) {
            try {
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
            }
        }

        # Open vmconnect only when VM starts
        if ($vmStarted) {
            try {
                vmconnect.exe localhost $VMName
            } catch {
                Write-Log "vmconnect launch failed for '$VMName': $($_.Exception.Message)" "WARN"
            }

            if ($attachISOForRecovery -and $ResetBootOrder) {
                try {
                    $hdd = Get-VMHardDiskDrive -VMName $VMName | Select-Object -First 1
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
        Write-Log "VM '$VMName' creation completed!" "OK"
        Write-Log "========================================" "OK"
        Update-CreateProgress -Percent 100 -Status "VM creation completed successfully"

        # Refresh GPU tab VM list
        Update-VMList

        $rollbackNeeded = $false

    } catch {
        Update-CreateProgress -Percent 0 -Status "VM creation failed"
        Write-ErrorWithGuidance -Context "Create VM ($VMName)" -ErrorRecord $_
        Write-Log "Stack: $($_.ScriptStackTrace)" "ERROR"
        if ($rollbackNeeded) {
            Remove-PartialVmArtifacts -VMName $VMName -VMLoc $VMLoc -VHDPath $VHDPath
        }
    } finally {
        if ($autoPlayChanged -and $originalAutoPlay -eq 0) {
            try {
                Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\AutoplayHandlers" -Name "DisableAutoplay" -Value 0
            } catch {
                Write-Log "Could not restore AutoPlay state: $($_.Exception.Message)" "WARN"
            }
        }
        $tabControl.Enabled  = $true
        $btnCreateVM.Enabled = $true
        if ($ctrlCreate.ContainsKey("CreateStatus") -and $ctrlCreate.ContainsKey("CreateProgress") -and $ctrlCreate["CreateProgress"].Value -lt 100) {
            $ctrlCreate["CreateStatus"].Text = "Ready to create VM"
        }
    }
})

# ================================================================
#  UPDATE GPU - Main Logic
# ================================================================
$btnUpdateGPU.Add_Click({
    $originalAutoPlay = $null
    $autoPlayChanged = $false

    try {
        # Gather selections
        $selectedVMs = @()
        foreach ($cb in $ctrlGPU["VMCheckboxes"]) {
            if ($cb.Checked) { $selectedVMs += $cb.Text }
        }
        if ($selectedVMs.Count -eq 0) { Write-Log "No VMs selected!" "ERROR"; return }

        $providerObj = $null
        $gpuVendor   = "Auto"
        if ($script:GpuPList.Count -gt 0) {
            $providerObj = $script:GpuPList | Where-Object { $_.Friendly -eq $ctrlGPU["GpuSelector"].SelectedItem }
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

        # Disable UI
        $tabControl.Enabled  = $false
        $btnUpdateGPU.Enabled = $false

        Write-Log "========================================" "INFO"
        Write-Log "Starting GPU driver update for $($selectedVMs.Count) VM(s)" "INFO"
        Write-Log "GPU Vendor: $gpuVendor | Smart Copy: $smartCopy | AutoExpand: $autoExpand | Allocation: $gpuAllocPercent%" "INFO"
        Write-Log "========================================" "INFO"

        # Disable AutoPlay
        $originalAutoPlay = Set-AutoPlay -Disable $true
        $autoPlayChanged = $true

        foreach ($VMName in $selectedVMs) {
            $vhdPath = $null
            $mountLetter = $null
            $mountedByScript = $false
            $skipDriverInjection = $false

            try {
                $vm = Get-VM -Name $VMName -ErrorAction Stop
                Write-Log "Processing VM: $VMName"

                # Shutdown VM if running
                if ($vm.State -eq 'Running') {
                    Write-Log "[$VMName] Shutting down..."
                    if (-not (Stop-VMWithTimeout -VMName $VMName -TimeoutSec 60)) {
                        Write-Log "[$VMName] VM did not stop cleanly. Skipping to avoid corruption." "ERROR"
                        continue
                    }
                    Write-Log "[$VMName] VM stopped." "OK"
                }

                # Remove existing GPU-P adapter
                Start-Sleep -Seconds 2
                Remove-VMGpuPartitionAdapter -VMName $VMName -Confirm:$false -ErrorAction SilentlyContinue
                Start-Sleep -Seconds 1

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

                        Set-GpuPartitionForVM -VMName $VMName -AllocationPercent $gpuAllocPercent
                        Write-Log "[$VMName] GPU-P adapter configured." "OK"
                    } catch {
                        Write-Log "[$VMName] GPU-P adapter error: $($_.Exception.Message)" "ERROR"
                    }
                } elseif ($providerObj -and $null -eq $providerObj.Provider) {
                    Write-Log "[$VMName] GPU-P adapter removed only (NONE selected)." "OK"
                    $skipDriverInjection = $true
                } else {
                    # No provider obj - try default for Win10
                    try {
                        Write-Log "[$VMName] Adding GPU-P adapter (default selection)"
                        Add-VMGpuPartitionAdapter -VMName $VMName -ErrorAction Stop
                        Set-GpuPartitionForVM -VMName $VMName -AllocationPercent $gpuAllocPercent
                        Write-Log "[$VMName] GPU-P adapter configured." "OK"
                    } catch {
                        Write-Log "[$VMName] GPU-P adapter error: $($_.Exception.Message)" "ERROR"
                    }
                }

                if ($skipDriverInjection) {
                    Write-Log "[$VMName] Driver injection skipped (remove-only mode)." "INFO"
                    continue
                }

                # Get VHD path
                $vhdPath = (Get-VMHardDiskDrive -VMName $VMName | Select-Object -First 1).Path
                if (-not $vhdPath) {
                    Write-Log "[$VMName] No VHDX found!" "ERROR"
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
                    $partitions = Get-Partition -DiskNumber $disk.Number -ErrorAction SilentlyContinue |
                        Where-Object { $_.Type -ne 'System' -and $_.Type -ne 'Reserved' -and $_.GptType -ne '{C12A7328-F81F-11D2-BA4B-00A0C93EC93B}' -and $_.GptType -ne '{E3C9E316-0B5C-4DB8-817D-F92DF00215AE}' }
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

                # ---- Copy GPU drivers (smart or full) ----
                $copyResult = Copy-GpuDriverFolders -VMName $VMName -MountLetter $mountLetter `
                    -VhdPath $vhdPath -GpuVendor $gpuVendor -SmartCopy $smartCopy -AutoExpand $autoExpand

                if (-not $copyResult) {
                    Write-Log "[$VMName] GPU driver copy failed." "ERROR"
                }

                # Copy vendor-specific System32 DLLs
                if ($script:SupportsGpuInstancePath -and $providerObj -and $providerObj.Provider) {
                    switch ($gpuVendor) {
                        "NVIDIA" {
                            Write-Log "[$VMName] Copying NVIDIA System32 DLLs (nv*)..."
                            Copy-DriversToVhd -VMName $VMName -MountLetter $mountLetter `
                                -Source "C:\Windows\System32" -Destination "Windows\System32" -FileMask "nv*"
                        }
                        "AMD" {
                            Write-Log "[$VMName] Copying AMD System32 DLLs..."
                            Copy-DriversToVhd -VMName $VMName -MountLetter $mountLetter `
                                -Source "C:\Windows\System32" -Destination "Windows\System32" -FileMask "amdkmd*"
                        }
                        "Intel" {
                            Write-Log "[$VMName] Copying Intel System32 DLLs..."
                            Copy-DriversToVhd -VMName $VMName -MountLetter $mountLetter `
                                -Source "C:\Windows\System32" -Destination "Windows\System32" -FileMask "igfx*"
                        }
                    }
                } else {
                    # Win10 default - check for NVIDIA
                    if (Test-HostHasNvidiaGpu) {
                        Write-Log "[$VMName] Host has NVIDIA GPU, copying nv* DLLs..."
                        Copy-DriversToVhd -VMName $VMName -MountLetter $mountLetter `
                            -Source "C:\Windows\System32" -Destination "Windows\System32" -FileMask "nv*"
                    }
                }

                if ($copySvcDriver) {
                    Copy-GpuServiceDriver -MountLetter $mountLetter -GPUName $(if ($providerObj -and $providerObj.Provider) { $providerObj.Friendly } else { "AUTO" })
                }

                Write-Log "[$VMName] GPU drivers injected." "OK"

                # Start VM if requested
                if ($startAfterUpdate) {
                    if (Start-VMWithRetry -VMName $VMName -MaxRetries 2) {
                        Write-Log "[$VMName] VM started." "OK"
                        # Open vmconnect if not already open
                        $existing = Get-CimInstance Win32_Process -Filter "Name = 'vmconnect.exe'" -ErrorAction SilentlyContinue |
                            Where-Object { $_.CommandLine -match [regex]::Escape($VMName) }
                        if (-not $existing) { vmconnect.exe localhost $VMName }
                    }
                }

                Write-Log "[$VMName] Done." "OK"

            } catch {
                Write-ErrorWithGuidance -Context "GPU update [$VMName]" -ErrorRecord $_
            } finally {
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
        if ($autoPlayChanged -and $originalAutoPlay -eq 0) {
            try {
                Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\AutoplayHandlers" -Name "DisableAutoplay" -Value 0
            } catch {
                Write-Log "Could not restore AutoPlay state after GPU update: $($_.Exception.Message)" "WARN"
            }
        }
        $tabControl.Enabled   = $true
        $btnUpdateGPU.Enabled = $true
    }
})

# ---- Form Closing Cleanup ----
$form.Add_FormClosing({
    Invoke-MountCleanup
})

#endregion

#region ==================== MAIN ====================

# Log header
Write-Log "Hyper-V Toolkit $($script:ToolkitVersion) by $($script:ToolkitCreator) - $($script:ToolkitTagline)" "OK"
Write-Log "Host OS: $($script:HostOsName)" "INFO"
Write-Log "GPU-P Specific GPU Selection: $(if ($script:SupportsGpuInstancePath) {'Available'} else {'Not available on this host build'})" "INFO"
Write-Log "Ready." "OK"

[void]$form.ShowDialog()

#endregion
