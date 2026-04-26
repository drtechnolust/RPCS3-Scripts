<#
.SYNOPSIS
    Detects whether Windows devices have completed the Secure Boot certificate transition before June 2026 expiration.

.DESCRIPTION
    This detection script checks if devices have fully transitioned to the new Secure Boot certificates
    before the June 2026 expiration of Microsoft Corporation KEK CA 2011 and UEFI CA 2011.
    Microsoft Windows Production PCA 2011 expires in October 2026.

    The script uses a tiered compliance model:
    - Stage 0: Secure Boot disabled (exit 1)
    - Stage 1: AvailableUpdates not configured and no deployment timestamp (exit 1 - triggers remediation)
    - Stage 2: Deployment configured, awaiting processing (exit 1)
    - Stage 3: Certificate updates in progress (exit 1)
    - Stage 3b: AvailableUpdates cleared, cert DB write pending (exit 1)
    - Stage 4: CA2023 certificate in UEFI DB but not yet booting from it (exit 1)
    - Stage 5: COMPLIANT (exit 0) — any one of three paths:
        Path A: WindowsUEFICA2023Capable = 2 (firmware attestation — authoritative, all builds)
        Path B: UEFICA2023Status = Updated    (pre-24H2 servicing pipeline)
        Path C: Build 26100+ + CA2023 in DB + bootmgfw.efi >= 10.0.19041.4522 (24H2+ OS image)

.PARAMETER FallbackDays
    Number of days to wait after managed opt-in before falling back to direct AvailableUpdates method.
    Default: 30

.PARAMETER TimestampRegPath
    Registry path where the ManagedOptInDate timestamp is stored by the remediation script.
    Default: HKLM:\SOFTWARE\Mindcore\Secureboot

.EXAMPLE
    .\Detect-SecureBootCertificateUpdate.ps1

.EXAMPLE
    .\Detect-SecureBootCertificateUpdate.ps1 -FallbackDays 45 -TimestampRegPath "HKLM:\SOFTWARE\Contoso\Secureboot"

.NOTES
    Version:        6.1
    Original Author: Mattias Melkersen
    Updated:        2026-04-15

    CHANGELOG
    ---------------
    2026-04-15 - v6.1 - Added KEK cert verification section to diagnostic block (UTF-16LE + ASCII dual decode)
                        Added UEFICA2023Status meaning interpretation to diagnostic log
                        Added ConfidenceLevel to Intune portal Write-Host output on Path A/B
                        Added 'Updated' -> code 5 mapping in status switch for pre-24H2 serviced state
                        Fixed Path A/B/C Write-Host output to include ConfidenceLevel when Under Observation
    2026-04-15 - v6.0 - Added WindowsUEFICA2023Capable = 2 as authoritative compliance indicator (Path A)
                        Added 24H2+ (Build 26100+) compliance bypass — Path C
                        Fixed ASCII-only UEFI DB decode — now UTF-16LE + ASCII dual pass
                        Fixed Stage 3->2 fallthrough when AvailableUpdates cleared to 0 (Stage 3b)
                        Fixed Stage 1 early compliance exit missing TimestampRegPath cleanup
                        Fixed $lastBoot/$uptime null reference risk in Stage 4
                        Fixed WinCsFlags.exe argument style with fallback and output validation
                        Fixed Get-AvailableUpdatesStatus missing 0x4104 stuck state label
                        Fixed Get-WindowsUEFICA2023Status value 0 ambiguity vs null
                        Fixed Get-FallbackStatus Parse() -> TryParse() for locale safety
                        Fixed Remove-Item -Recurse -> Remove-ItemProperty (scoped)
                        Fixed event log query — removed System channel noise
                        Added 64-bit process enforcement guard
                        Added ConfidenceLevel to registry read and diagnostic output
                        Added Get-BootManagerCompliance helper function
    2026-04-13 - v5.0 - Aligned with Microsoft Secure Boot playbook (MM)
    2026-04-06 - v4.0 - Added SecureBootUpdates payload folder validation (MM)
    2026-04-05 - v3.1 - Buffered logging (MM)
    2026-03-27 - v3.0 - Added configurable fallback timer (MM)
    2026-02-19 - v2.2 - Fixed misleading Updates label (MM)
    2026-02-18 - v2.1 - Enhanced local device logging (MM)
    2026-02-18 - v2.0 - Tiered compliance model (MM)
    2026-01-15 - v1.0 - Initial version (MM)

    References:
    - https://aka.ms/getsecureboot
    - https://techcommunity.microsoft.com/blog/windows-itpro-blog/act-now-secure-boot-certificates-expire-in-june-2026/4426856
    - https://support.microsoft.com/topic/enterprise-deployment-guidance-for-cve-2023-24932-88b8f034-20b7-4a45-80cb-c6049b0f9967
    - https://blog.mindcore.dk/2026/04/secure-boot-certificate-update-intune/
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [int]$FallbackDays = 30,

    [Parameter(Mandatory = $false)]
    [string]$TimestampRegPath = "HKLM:\SOFTWARE\Mindcore\Secureboot"
)

#region 64-bit Enforcement
# Confirm-SecureBootUEFI, WMI security namespaces, and COM objects require 64-bit PowerShell.
if (-not [Environment]::Is64BitProcess) {
    $sysnativePosh = "$env:SystemRoot\SysNative\WindowsPowerShell\v1.0\powershell.exe"
    if (Test-Path $sysnativePosh) {
        $scriptPath = $MyInvocation.MyCommand.Path
        $argList    = "-NonInteractive -ExecutionPolicy Bypass -File `"$scriptPath`""
        if ($FallbackDays -ne 30) {
            $argList += " -FallbackDays $FallbackDays"
        }
        if ($TimestampRegPath -ne "HKLM:\SOFTWARE\Mindcore\Secureboot") {
            $argList += " -TimestampRegPath `"$TimestampRegPath`""
        }
        $proc = Start-Process -FilePath $sysnativePosh -ArgumentList $argList `
                    -Wait -NoNewWindow -PassThru
        exit $proc.ExitCode
    }
}
#endregion

#region Logging
[string]$LogFile    = "$env:ProgramData\Microsoft\IntuneManagementExtension\Logs\SecureBootCertificateUpdate.log"
[string]$ScriptName = "DETECT"
[int]$MaxLogSizeMB  = 4
$script:LogBuffer   = [System.Collections.Generic.List[string]]::new()

function Write-Log {
    param(
        [Parameter(Mandatory = $true)][string]$Message,
        [Parameter(Mandatory = $false)]
        [ValidateSet("INFO","WARNING","ERROR","SUCCESS")]
        [string]$Level = "INFO"
    )
    $script:LogBuffer.Add("$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') [$ScriptName] [$Level] $Message")
}

function Flush-Log {
    if ($script:LogBuffer.Count -eq 0) { return }
    try {
        $LogDir = Split-Path $LogFile -Parent
        if (-not (Test-Path $LogDir)) {
            New-Item -Path $LogDir -ItemType Directory -Force | Out-Null
        }
        if (Test-Path $LogFile) {
            if ((Get-Item $LogFile).Length / 1MB -ge $MaxLogSizeMB) {
                $BackupLog = "$LogFile.old"
                if (Test-Path $BackupLog) { Remove-Item $BackupLog -Force -ErrorAction SilentlyContinue }
                Rename-Item $LogFile $BackupLog -Force -ErrorAction SilentlyContinue
                $script:LogBuffer.Insert(0, "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') [SYSTEM] [INFO] Log rotated — archived to: $BackupLog")
            }
        }
        Add-Content -Path $LogFile -Value $script:LogBuffer.ToArray() -ErrorAction SilentlyContinue
        $script:LogBuffer.Clear()
    }
    catch { }
}
#endregion

#region Helper Functions

function Get-SecureBootStatus {
    try { return Confirm-SecureBootUEFI -ErrorAction Stop }
    catch {
        try {
            $regPath = "HKLM:\SYSTEM\CurrentControlSet\Control\SecureBoot\State"
            if (Test-Path $regPath) {
                $value = (Get-ItemProperty $regPath -Name "UEFISecureBootEnabled" `
                    -ErrorAction SilentlyContinue).UEFISecureBootEnabled
                return ($value -eq 1)
            }
        }
        catch { }
    }
    return $false
}

function Get-AvailableUpdatesStatus {
    param([int]$Value)
    switch ($Value) {
        22852  { return "Not Started — All updates pending (0x5944)" }
        16384  { return "Complete — All certificates applied (0x4000)" }
        16644  { return "STUCK — KEK cert not clearing (0x4104) — OEM firmware update may be required" }
        0      { return "Cleared — No pending updates (0x0)" }
        default { return "In Progress (0x$($Value.ToString('X')))" }
    }
}

function Get-WindowsUEFICA2023Status {
    param([int]$Value)
    switch ($Value) {
        0 { return "Checked — cert not yet in DB (0)" }
        1 { return "Cert present in UEFI DB (1)" }
        2 { return "Cert in DB AND booting from CA 2023 signed boot manager (2) — COMPLIANT" }
        default { return "Unknown ($Value)" }
    }
}

function Get-UEFIStatusMeaning {
    param([object]$Status)
    if ($null -eq $Status) { return "Key not present — WU has not processed yet" }
    switch ($Status.ToString()) {
        'NotStarted' { return "Expected on 24H2+ — cert ships in OS image" }
        'Updated'    { return "Fully serviced — pre-24H2 compliance state" }
        'CertInDb'   { return "Cert in DB only — boot manager not yet updated" }
        default      { return $Status.ToString() }
    }
}

function Get-FallbackStatus {
    param([string]$RegPath, [int]$Threshold)
    $result = @{
        TimestampExists = $false
        OptInDate       = $null
        DaysElapsed     = 0
        DaysRemaining   = $Threshold
        IsActive        = $false
    }
    try {
        if (Test-Path $RegPath) {
            $dateStr = (Get-ItemProperty $RegPath -Name "ManagedOptInDate" `
                -ErrorAction SilentlyContinue).ManagedOptInDate
            if ($dateStr) {
                $parsed = $null
                if ([DateTime]::TryParse($dateStr, [ref]$parsed)) {
                    $elapsed = ((Get-Date) - $parsed).TotalDays
                    $result.TimestampExists = $true
                    $result.OptInDate       = $parsed.ToString("yyyy-MM-dd HH:mm:ss")
                    $result.DaysElapsed     = [Math]::Floor($elapsed)
                    $result.DaysRemaining   = [Math]::Max(0, $Threshold - [Math]::Floor($elapsed))
                    $result.IsActive        = ($elapsed -ge $Threshold)
                }
            }
        }
    }
    catch { }
    return $result
}

function Get-SecureBootPayloadStatus {
    $payloadPath = "$env:SystemRoot\System32\SecureBootUpdates"
    $result = @{
        FolderExists = $false
        FileCount    = 0
        Files        = @()
        HasBinFiles  = $false
        IsHealthy    = $false
        State        = "Unknown"
    }
    try {
        if (Test-Path $payloadPath) {
            $result.FolderExists = $true
            $files = Get-ChildItem $payloadPath -File -ErrorAction SilentlyContinue
            if ($files) {
                $result.FileCount   = $files.Count
                $result.Files       = $files | ForEach-Object { "$($_.Name) ($([Math]::Round($_.Length/1KB,1))KB)" }
                $result.HasBinFiles = ($files | Where-Object { $_.Extension -eq '.bin' }).Count -gt 0
                $result.IsHealthy   = $result.HasBinFiles
                $result.State       = if ($result.HasBinFiles) { "Healthy" } else { "Empty" }
            }
            else { $result.State = "Empty" }
        }
        else { $result.State = "Missing" }
    }
    catch { }
    return $result
}

function Get-SecureBootTaskStatus {
    $result = @{
        TaskExists     = $false
        LastRunTime    = $null
        LastTaskResult = $null
        NextRunTime    = $null
        ResultHex      = $null
        IsMissingFiles = $false
    }
    try {
        $task = Get-ScheduledTask -TaskPath "\Microsoft\Windows\PI\" `
            -TaskName "Secure-Boot-Update" -ErrorAction SilentlyContinue
        if ($task) {
            $result.TaskExists = $true
            $taskInfo = Get-ScheduledTaskInfo -TaskPath "\Microsoft\Windows\PI\" `
                -TaskName "Secure-Boot-Update" -ErrorAction SilentlyContinue
            if ($taskInfo) {
                $result.LastRunTime    = $taskInfo.LastRunTime
                $result.LastTaskResult = $taskInfo.LastTaskResult
                $result.ResultHex      = "0x$($taskInfo.LastTaskResult.ToString('X'))"
                $result.NextRunTime    = $taskInfo.NextRunTime
                $result.IsMissingFiles = ($taskInfo.LastTaskResult -eq 0x80070002)
            }
        }
    }
    catch { }
    return $result
}

function Get-BootManagerCompliance {
    $result = @{
        Compliant    = $false
        Version      = $null
        CleanVersion = $null
        Path         = "$env:SystemRoot\Boot\EFI\bootmgfw.efi"
        Found        = $false
    }
    if (Test-Path $result.Path) {
        $result.Found        = $true
        $result.Version      = (Get-Item $result.Path).VersionInfo.FileVersion
        # Strip trailing build label e.g. " (WinBuild.160101.0800)" before version parse
        $result.CleanVersion = ($result.Version -split ' ')[0].Trim()
        $parsedVer           = $null
        $result.Compliant    = [System.Version]::TryParse($result.CleanVersion, [ref]$parsedVer) -and
                               ($parsedVer -ge [System.Version]'10.0.19041.4522')
    }
    return $result
}

function Remove-TrackingTimestamp {
    param([string]$RegPath)
    if (Test-Path $RegPath) {
        try {
            Remove-ItemProperty -Path $RegPath -Name "ManagedOptInDate" -Force -ErrorAction Stop
            Write-Log -Message "  Cleanup: Removed ManagedOptInDate from $RegPath" -Level "SUCCESS"
        }
        catch {
            Write-Log -Message "  Cleanup: Could not remove ManagedOptInDate — $($_.Exception.Message)" -Level "WARNING"
        }
    }
}

function Get-ConfidenceNote {
    # Returns a portal-safe note string when device is in Under Observation cohort
    param([object]$ConfidenceLevel)
    if ($null -ne $ConfidenceLevel -and $ConfidenceLevel -match 'Under Observation') {
        return " | Confidence:UnderObservation(WU-Held-NotAFailure)"
    }
    return ""
}

function Set-ExitCode {
    # ISE-safe exit — uses return (caller must follow with 'return') in ISE,
    # uses exit in non-ISE (Intune IME, powershell.exe -File).
    param([int]$Code)
    Flush-Log
    if ($Host.Name -eq 'Windows PowerShell ISE Host') {
        $global:LASTEXITCODE = $Code
        Write-Host "(ISE mode — LASTEXITCODE set to $Code, host kept alive)" -ForegroundColor Cyan
    }
    else {
        exit $Code
    }
}

#endregion

#region Main Detection Logic
try {
    Write-Log -Message "========== DETECTION STARTED ==========" -Level "INFO"
    Write-Log -Message "Script Version: 6.1" -Level "INFO"
    Write-Log -Message "Computer: $env:COMPUTERNAME | User: $env:USERNAME" -Level "INFO"
    Write-Log -Message "PowerShell: $($PSVersionTable.PSVersion) | Process: $(if ([Environment]::Is64BitProcess) {'64-bit'} else {'32-bit'})" -Level "INFO"

    # ── Secure Boot state ──────────────────────────────────────────────────
    Write-Log -Message "Checking Secure Boot status..." -Level "INFO"
    $secureBootEnabled = Get-SecureBootStatus

    # ══════════════════════════════════════════════════════════════════════
    # STAGE 0 — Secure Boot must be enabled
    # ══════════════════════════════════════════════════════════════════════
    if (-not $secureBootEnabled) {
        Write-Log -Message "Secure Boot is DISABLED — Cannot apply certificate updates" -Level "ERROR"
        Write-Log -Message "--- Stage 0 Diagnostics ---" -Level "INFO"
        try {
            $sbStatePath = "HKLM:\SYSTEM\CurrentControlSet\Control\SecureBoot\State"
            if (Test-Path $sbStatePath) {
                $sbStateValue = (Get-ItemProperty $sbStatePath -Name "UEFISecureBootEnabled" `
                    -ErrorAction SilentlyContinue).UEFISecureBootEnabled
                Write-Log -Message "  Firmware Mode: UEFI (SecureBoot State key exists)" -Level "INFO"
                Write-Log -Message "  UEFISecureBootEnabled: $sbStateValue" -Level "INFO"
                Write-Log -Message "  WHY: Device supports Secure Boot but it is DISABLED in BIOS/UEFI settings" -Level "WARNING"
                Write-Log -Message "  NEXT STEPS: Enter BIOS/UEFI setup and enable Secure Boot under Security settings" -Level "WARNING"
            }
            else {
                Write-Log -Message "  Firmware Mode: Likely Legacy BIOS (SecureBoot State key does not exist)" -Level "WARNING"
                Write-Log -Message "  WHY: Legacy BIOS firmware does not support Secure Boot" -Level "ERROR"
                Write-Log -Message "  NEXT STEPS: Convert disk to GPT and switch firmware from Legacy to UEFI" -Level "ERROR"
                Write-Log -Message "  Reference: https://learn.microsoft.com/en-us/windows/deployment/mbr-to-gpt" -Level "INFO"
                $osDisk = Get-Disk -Number 0 -ErrorAction SilentlyContinue
                if ($osDisk) {
                    Write-Log -Message "  OS Disk Partition Style: $($osDisk.PartitionStyle)" -Level "INFO"
                    if ($osDisk.PartitionStyle -eq "MBR") {
                        Write-Log -Message "  MBR disk detected — MBR2GPT conversion required before enabling UEFI" -Level "WARNING"
                    }
                }
                else {
                    Write-Log -Message "  OS Disk: Unable to determine partition style" -Level "WARNING"
                }
            }
            $biosInfo = Get-CimInstance -ClassName Win32_BIOS -ErrorAction SilentlyContinue
            if ($biosInfo) {
                Write-Log -Message "  BIOS Manufacturer: $($biosInfo.Manufacturer)" -Level "INFO"
                Write-Log -Message "  BIOS Version: $($biosInfo.SMBIOSBIOSVersion)" -Level "INFO"
                Write-Log -Message "  BIOS Release Date: $($biosInfo.ReleaseDate)" -Level "INFO"
            }
        }
        catch {
            Write-Log -Message "  Stage 0 supplemental diagnostics failed: $($_.Exception.Message)" -Level "WARNING"
        }
        Write-Log -Message "--- End Stage 0 Diagnostics ---" -Level "INFO"
        Write-Host "SECURE_BOOT_DISABLED | Action: Enable Secure Boot in BIOS/UEFI"
        Write-Log -Message "Detection Result: NON-COMPLIANT - Stage 0 (exit 1)" -Level "WARNING"
        Write-Log -Message "========== DETECTION COMPLETED ==========" -Level "INFO"
        Set-ExitCode 1
        return
    }
    Write-Log -Message "Secure Boot is ENABLED" -Level "SUCCESS"

    # ── Read all core registry values up front ─────────────────────────────
    $regPath       = "HKLM:\SYSTEM\CurrentControlSet\Control\Secureboot"
    $servicingPath = "HKLM:\SYSTEM\CurrentControlSet\Control\SecureBoot\Servicing"

    $availableUpdatesValue = $null
    if (Test-Path $regPath) {
        $availableUpdatesValue = (Get-ItemProperty $regPath -Name "AvailableUpdates" `
            -ErrorAction SilentlyContinue).AvailableUpdates
    }

    $svcProps        = Get-ItemProperty $servicingPath -ErrorAction SilentlyContinue
    $uefiCA2023Status = $svcProps.UEFICA2023Status
    $ca2023Capable   = $svcProps.WindowsUEFICA2023Capable
    $uefiError       = $svcProps.UEFICA2023Error
    $uefiErrorEvent  = $svcProps.UEFICA2023ErrorEvent
    $confidenceLevel = $svcProps.ConfidenceLevel

    # Log all key values with meanings
    Write-Log -Message "UEFICA2023Status         : $(if ($null -ne $uefiCA2023Status) {$uefiCA2023Status} else {'<not set>'})" -Level "INFO"
    Write-Log -Message "UEFICA2023Status meaning  : $(Get-UEFIStatusMeaning -Status $uefiCA2023Status)" -Level "INFO"
    Write-Log -Message "WindowsUEFICA2023Capable  : $(if ($null -ne $ca2023Capable) {$ca2023Capable} else {'<not set>'})" -Level "INFO"
    Write-Log -Message "AvailableUpdates          : $(if ($null -ne $availableUpdatesValue) {"0x$($availableUpdatesValue.ToString('X')) ($availableUpdatesValue)"} else {'<not set>'})" -Level "INFO"

    if ($null -ne $confidenceLevel) {
        Write-Log -Message "ConfidenceLevel           : $confidenceLevel" -Level "INFO"
        if ($confidenceLevel -match 'Under Observation') {
            Write-Log -Message "  NOTE: Microsoft is withholding WU auto-push for this device cohort" -Level "INFO"
            Write-Log -Message "  NOTE: This is NOT a compliance failure — firmware attestation is authoritative" -Level "INFO"
        }
    }

    if ($null -ne $uefiError -and $uefiError -ne 0) {
        Write-Log -Message "UEFICA2023Error           : 0x$($uefiError.ToString('X')) ($uefiError)" -Level "ERROR"
        if ($null -ne $uefiErrorEvent) {
            Write-Log -Message "UEFICA2023ErrorEvent      : $uefiErrorEvent (check System Event Log)" -Level "ERROR"
        }
        Write-Log -Message "  Reference: https://support.microsoft.com/topic/37e47cf8-608b-4a87-8175-bdead630eb69" -Level "WARNING"
    }

    # ── OS build ───────────────────────────────────────────────────────────
    $osInfo     = Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction SilentlyContinue
    $osBuildInt = if ($osInfo) { [int]$osInfo.BuildNumber } else { 0 }
    Write-Log -Message "OS Build: $osBuildInt | Caption: $($osInfo.Caption)" -Level "INFO"

    # ── UEFI DB — dual UTF-16LE + ASCII decode ─────────────────────────────
    # EFI_SIGNATURE_LIST cert subjects are stored as UTF-16LE.
    # ASCII-only decode (used by some scripts) produces false negatives — do not use alone.
    $hasCA2023InDB = $false
    try {
        $dbBytes     = (Get-SecureBootUEFI db -ErrorAction Stop).bytes
        $dbText      = [System.Text.Encoding]::Unicode.GetString($dbBytes) +
                       [System.Text.Encoding]::ASCII.GetString($dbBytes)
        $hasCA2023InDB = $dbText -match 'Windows UEFI CA 2023'
        Write-Log -Message "UEFI DB CA2023 Cert       : $(if ($hasCA2023InDB) {'PRESENT'} else {'NOT FOUND'})" `
            -Level $(if ($hasCA2023InDB) { 'SUCCESS' } else { 'WARNING' })
    }
    catch {
        Write-Log -Message "UEFI DB query failed: $($_.Exception.Message)" -Level "WARNING"
    }

    # ── Boot manager ───────────────────────────────────────────────────────
    $bootMgr = Get-BootManagerCompliance
    if ($bootMgr.Found) {
        Write-Log -Message "Boot Manager             : $($bootMgr.CleanVersion) — $(if ($bootMgr.Compliant) {'Meets minimum'} else {'BELOW MINIMUM (10.0.19041.4522)'})" `
            -Level $(if ($bootMgr.Compliant) { 'SUCCESS' } else { 'WARNING' })
    }
    else {
        Write-Log -Message "Boot Manager             : bootmgfw.efi not found at expected path" -Level "WARNING"
    }

    # ══════════════════════════════════════════════════════════════════════
    # STAGE 5 — COMPLIANT (three independent paths — any one = exit 0)
    # Evaluated before all non-compliant stages
    # ══════════════════════════════════════════════════════════════════════

    $confNote = Get-ConfidenceNote -ConfidenceLevel $confidenceLevel

    # ── Path A: WindowsUEFICA2023Capable = 2 ──────────────────────────────
    # Authoritative firmware attestation signal — valid on ALL builds.
    # Value 2 = firmware has confirmed device is booting from CA 2023 signed boot manager.
    # Takes priority over UEFICA2023Status, OS build, and DB string matching.
    if ($null -ne $ca2023Capable -and [int]$ca2023Capable -eq 2) {
        Write-Log -Message "--- COMPLIANT: Path A — WindowsUEFICA2023Capable=2 ---" -Level "SUCCESS"
        Write-Log -Message "  Firmware confirmed: device is booting from CA 2023 signed boot manager chain" -Level "SUCCESS"
        Write-Log -Message "  This signal is authoritative regardless of UEFICA2023Status or OS build" -Level "SUCCESS"
        if ($null -ne $confidenceLevel) {
            Write-Log -Message "  ConfidenceLevel: $confidenceLevel" -Level "INFO"
            if ($confidenceLevel -match 'Under Observation') {
                Write-Log -Message "  Under Observation = WU auto-push withheld for this cohort — not a compliance issue" -Level "INFO"
            }
        }
        Remove-TrackingTimestamp -RegPath $TimestampRegPath
        Write-Host "COMPLIANT | Path:FirmwareAttestation | CA2023Capable=2$(if ($hasCA2023InDB) {' | DB:Present'})$confNote"
        Write-Log -Message "Detection Result: COMPLIANT - Stage 5 Path A (exit 0)" -Level "SUCCESS"
        Write-Log -Message "========== DETECTION COMPLETED ==========" -Level "INFO"
        Set-ExitCode 0
        return
    }

    # ── Path B: UEFICA2023Status = "Updated" ──────────────────────────────
    # Pre-24H2 servicing pipeline fully completed.
    if ($uefiCA2023Status -eq "Updated") {
        Write-Log -Message "--- COMPLIANT: Path B — UEFICA2023Status=Updated ---" -Level "SUCCESS"
        Write-Log -Message "  Pre-24H2 servicing pipeline completed — cert + boot manager fully updated" -Level "SUCCESS"
        if ($null -ne $ca2023Capable) {
            Write-Log -Message "  WindowsUEFICA2023Capable: $ca2023Capable" -Level "SUCCESS"
        }
        Remove-TrackingTimestamp -RegPath $TimestampRegPath
        Write-Host "COMPLIANT | Path:ServicingUpdated | UEFICA2023Status=Updated$(if ($null -ne $ca2023Capable) {" | CA2023Capable=$ca2023Capable"})$confNote"
        Write-Log -Message "Detection Result: COMPLIANT - Stage 5 Path B (exit 0)" -Level "SUCCESS"
        Write-Log -Message "========== DETECTION COMPLETED ==========" -Level "INFO"
        Set-ExitCode 0
        return
    }

    # ── Path C: Build 26100+ + CA2023 in DB + boot manager version ─────────
    # On Windows 11 24H2+/25H2, the CA 2023 cert ships pre-loaded in the OS image.
    # UEFICA2023Status is never written as Updated on this path.
    if ($osBuildInt -ge 26100) {
        Write-Log -Message "--- Build $osBuildInt = Windows 11 24H2+/25H2 — evaluating Path C ---" -Level "INFO"
        Write-Log -Message "  CA2023 cert ships in OS image on this build — UEFICA2023Status=Updated not written" -Level "INFO"
        Write-Log -Message "  Compliance check: CA2023 in UEFI DB + bootmgfw.efi >= 10.0.19041.4522" -Level "INFO"

        if ($hasCA2023InDB -and $bootMgr.Compliant) {
            Write-Log -Message "--- COMPLIANT: Path C — 24H2+/25H2 OS image path ---" -Level "SUCCESS"
            Write-Log -Message "  CA2023 cert: Present in UEFI DB" -Level "SUCCESS"
            Write-Log -Message "  Boot manager: $($bootMgr.CleanVersion) meets minimum" -Level "SUCCESS"
            Remove-TrackingTimestamp -RegPath $TimestampRegPath
            Write-Host "COMPLIANT | Path:24H2OsImage | Build:$osBuildInt | CA2023:InDB | BootMgr:$($bootMgr.CleanVersion)$confNote"
            Write-Log -Message "Detection Result: COMPLIANT - Stage 5 Path C (exit 0)" -Level "SUCCESS"
            Write-Log -Message "========== DETECTION COMPLETED ==========" -Level "INFO"
            Set-ExitCode 0
            return
        }
        else {
            Write-Log -Message "  Path C NOT met — CA2023InDB=$hasCA2023InDB | BootMgrOK=$($bootMgr.Compliant) ($($bootMgr.CleanVersion))" -Level "WARNING"
            Write-Log -Message "  24H2+/25H2 device non-compliant — cert missing from DB or boot manager below minimum" -Level "WARNING"
            Write-Host "REVIEW_REQUIRED | Build:$osBuildInt(24H2+) | CA2023InDB:$hasCA2023InDB | BootMgr:$(if ($bootMgr.Found) {$bootMgr.CleanVersion} else {'NotFound'})$confNote"
            Write-Log -Message "Detection Result: NON-COMPLIANT - 24H2+ Path C failed (exit 1)" -Level "WARNING"
            Write-Log -Message "========== DETECTION COMPLETED ==========" -Level "INFO"
            Set-ExitCode 1
            return
        }
    }

    # ══════════════════════════════════════════════════════════════════════
    # PRE-24H2 NON-COMPLIANT STAGES
    # Reached only when all three compliance paths above were not met
    # ══════════════════════════════════════════════════════════════════════

    $deploymentTimestamp = $null
    if (Test-Path $TimestampRegPath) {
        $deploymentTimestamp = (Get-ItemProperty $TimestampRegPath -Name "ManagedOptInDate" `
            -ErrorAction SilentlyContinue).ManagedOptInDate
    }

    $deploymentTriggered = ($null -ne $availableUpdatesValue -and $availableUpdatesValue -ne 0) -or
                           ($null -ne $deploymentTimestamp)

    $fallback = Get-FallbackStatus -RegPath $TimestampRegPath -Threshold $FallbackDays
    if ($fallback.TimestampExists) {
        Write-Log -Message "Fallback Timer: OptIn=$($fallback.OptInDate) | Elapsed=$($fallback.DaysElapsed)d | Threshold=$($FallbackDays)d | Remaining=$($fallback.DaysRemaining)d | Active=$($fallback.IsActive)" -Level "INFO"
    }

    # ── Stage 1: Deployment not triggered ─────────────────────────────────
    if (-not $deploymentTriggered) {
        Write-Log -Message "Deployment NOT TRIGGERED — Remediation required" -Level "WARNING"
        Write-Log -Message "--- Stage 1 Analysis ---" -Level "INFO"
        Write-Log -Message "  AvailableUpdates: $(if ($null -eq $availableUpdatesValue) {'<does not exist>'} else {"$availableUpdatesValue (0x$($availableUpdatesValue.ToString('X')))"})" -Level "INFO"
        Write-Log -Message "  Deployment Timestamp: $(if ($null -eq $deploymentTimestamp) {'<does not exist>'} else {$deploymentTimestamp})" -Level "INFO"
        Write-Log -Message "  Expected: AvailableUpdates = 0x5944 (22852)" -Level "INFO"
        Write-Log -Message "  WHY: AvailableUpdates has not been set to trigger Secure Boot certificate deployment" -Level "INFO"
        Write-Log -Message "  NEXT STEPS: The remediation script will automatically set AvailableUpdates = 0x5944" -Level "INFO"
        Write-Log -Message "--- End Stage 1 Analysis ---" -Level "INFO"
        Write-Host "DEPLOYMENT_NOT_TRIGGERED | Action: Remediation will set AvailableUpdates"
        Write-Log -Message "Detection Result: NON-COMPLIANT - Stage 1 (exit 1)" -Level "WARNING"
        Write-Log -Message "========== DETECTION COMPLETED ==========" -Level "INFO"
        Set-ExitCode 1
        return
    }
    Write-Log -Message "Deployment TRIGGERED (AvailableUpdates: $(if ($null -ne $availableUpdatesValue) {"0x$($availableUpdatesValue.ToString('X'))"} else {'cleared'}) | Timestamp: $(if ($null -ne $deploymentTimestamp) {'exists'} else {'none'}))" -Level "SUCCESS"

    # Build progressive detail string
    $detailList = [System.Collections.Generic.List[string]]::new()
    if ($null -ne $availableUpdatesValue -and ($null -eq $ca2023Capable -or [int]$ca2023Capable -lt 1)) {
        $detailList.Add("Updates:$(Get-AvailableUpdatesStatus -Value $availableUpdatesValue)")
    }
    if ($null -ne $uefiCA2023Status)  { $detailList.Add("Status:$uefiCA2023Status") }
    if ($null -ne $ca2023Capable)     { $detailList.Add("CA2023:$(Get-WindowsUEFICA2023Status -Value ([int]$ca2023Capable))") }
    if ($null -ne $uefiError -and $uefiError -ne 0) { $detailList.Add("Error:0x$($uefiError.ToString('X'))") }

    # ══════════════════════════════════════════════════════════════════════
    # DIAGNOSTIC DATA COLLECTION
    # Only reached for pre-24H2 non-compliant devices
    # ══════════════════════════════════════════════════════════════════════
    Write-Log -Message "---------- DIAGNOSTIC DATA ----------" -Level "INFO"

    # ── Payload folder ─────────────────────────────────────────────────────
    Write-Log -Message "--- SecureBootUpdates Payload Check ---" -Level "INFO"
    $payload = Get-SecureBootPayloadStatus
    Write-Log -Message "  Payload State: $($payload.State)" `
        -Level $(if ($payload.State -eq 'Healthy') { 'SUCCESS' } else { 'WARNING' })
    if ($payload.FolderExists -and $payload.FileCount -gt 0) {
        foreach ($f in $payload.Files) { Write-Log -Message "  File: $f" -Level "INFO" }
    }
    if ($payload.State -ne 'Healthy') {
        Write-Log -Message "  May cause Secure-Boot-Update task to fail with 0x80070002" -Level "WARNING"
        Write-Log -Message "  FIX: Install latest cumulative update or use WinCsFlags.exe if available" -Level "WARNING"
    }
    Write-Log -Message "--- End Payload Check ---" -Level "INFO"

    # ── Scheduled task ─────────────────────────────────────────────────────
    Write-Log -Message "--- Secure-Boot-Update Task Status ---" -Level "INFO"
    $taskStatus = Get-SecureBootTaskStatus
    if ($taskStatus.TaskExists) {
        $lastRunStr = if ($null -ne $taskStatus.LastRunTime -and $taskStatus.LastRunTime.Year -gt 2000) {
            $taskStatus.LastRunTime.ToString('yyyy-MM-dd HH:mm:ss')
        } else { "Never" }
        Write-Log -Message "  Task: Found | Last Run: $lastRunStr | Result: $($taskStatus.ResultHex)" -Level "INFO"
        if ($taskStatus.IsMissingFiles) {
            Write-Log -Message "  ALERT: Task failed 0x80070002 — missing payload files" -Level "ERROR"
            Write-Log -Message "  FIX: Install latest cumulative update or use WinCsFlags.exe" -Level "ERROR"
        }
        elseif ($taskStatus.LastTaskResult -ne 0) {
            Write-Log -Message "  WARNING: Task exited with non-zero result $($taskStatus.ResultHex)" -Level "WARNING"
        }
        else {
            Write-Log -Message "  Task result: Success (0x0)" -Level "SUCCESS"
        }
    }
    else {
        Write-Log -Message "  Task: NOT FOUND (requires July 2024+ cumulative update)" -Level "WARNING"
    }
    Write-Log -Message "--- End Task Status ---" -Level "INFO"

    # ── WinCsFlags.exe ─────────────────────────────────────────────────────
    Write-Log -Message "--- WinCS Availability ---" -Level "INFO"
    $winCsPath      = "$env:SystemRoot\System32\WinCsFlags.exe"
    $winCsAvailable = Test-Path $winCsPath
    if ($winCsAvailable) {
        Write-Log -Message "  WinCsFlags.exe: AVAILABLE" -Level "SUCCESS"
        try {
            # Try Windows-style args first, fall back to GNU-style
            $winCsOutput = & $winCsPath /query /key F33E0C8E002 2>&1
            if (-not $winCsOutput) {
                $winCsOutput = & $winCsPath --query --key F33E0C8E002 2>&1
            }
            if ($winCsOutput) {
                foreach ($line in (($winCsOutput | Out-String).Trim() -split "`n")) {
                    $trimmed = $line.Trim()
                    if ($trimmed) { Write-Log -Message "  WinCS: $trimmed" -Level "INFO" }
                }
            }
            else { Write-Log -Message "  WinCS: No output returned from query" -Level "WARNING" }
        }
        catch { Write-Log -Message "  WinCS query failed: $($_.Exception.Message)" -Level "WARNING" }
    }
    else {
        Write-Log -Message "  WinCsFlags.exe: NOT AVAILABLE (requires Oct/Nov 2025+ cumulative update)" -Level "INFO"
    }
    Write-Log -Message "--- End WinCS Availability ---" -Level "INFO"

    # ── UEFI DB full cert check ────────────────────────────────────────────
    Write-Log -Message "--- UEFI DB Certificate Verification ---" -Level "INFO"
    try {
        $dbBytes   = (Get-SecureBootUEFI db -ErrorAction Stop).bytes
        $dbFull    = [System.Text.Encoding]::Unicode.GetString($dbBytes) +
                     [System.Text.Encoding]::ASCII.GetString($dbBytes)
        $db2023    = $dbFull -match 'Windows UEFI CA 2023'
        $db2011Win = $dbFull -match 'Microsoft Windows Production PCA 2011'
        $db2011MS  = $dbFull -match 'Microsoft Corporation UEFI CA 2011'
        Write-Log -Message "  Windows UEFI CA 2023               : $(if ($db2023)    {'PRESENT'} else {'NOT FOUND'})" `
            -Level $(if ($db2023)    { 'SUCCESS' } else { 'WARNING' })
        Write-Log -Message "  Windows Production PCA 2011        : $(if ($db2011Win) {'Present'} else {'Not found'})" `
            -Level $(if ($db2011Win) { 'SUCCESS' } else { 'INFO' })
        Write-Log -Message "  Microsoft Corporation UEFI CA 2011 : $(if ($db2011MS)  {'Present'} else {'Not found'})" `
            -Level $(if ($db2011MS)  { 'SUCCESS' } else { 'INFO' })
        Write-Log -Message "  Note: DB inspection is supporting evidence — registry attestation is authoritative" -Level "INFO"
    }
    catch { Write-Log -Message "  UEFI DB: Unable to query — $($_.Exception.Message)" -Level "WARNING" }
    Write-Log -Message "--- End UEFI DB Verification ---" -Level "INFO"

    # ── UEFI KEK cert check ────────────────────────────────────────────────
    # Uses dual UTF-16LE + ASCII decode — ASCII-only decode misses certs stored as UTF-16LE.
    # KEK string matching is heuristic only — OEM implementations vary.
    Write-Log -Message "--- UEFI KEK Certificate Verification ---" -Level "INFO"
    try {
        $kekBytes   = (Get-SecureBootUEFI KEK -ErrorAction Stop).bytes
        $kekText    = [System.Text.Encoding]::Unicode.GetString($kekBytes) +
                      [System.Text.Encoding]::ASCII.GetString($kekBytes)
        $kekHas2023 = $kekText -match 'Microsoft Corporation KEK 2K CA 2023'
        $kekHas2011 = $kekText -match 'Microsoft Corporation KEK CA 2011'
        Write-Log -Message "  KEK 2K CA 2023 : $(if ($kekHas2023) {'Present'} else {'Not detected'})" `
            -Level $(if ($kekHas2023) { 'SUCCESS' } else { 'WARNING' })
        Write-Log -Message "  KEK CA 2011    : $(if ($kekHas2011) {'Present'} else {'Not detected'})" `
            -Level $(if ($kekHas2011) { 'SUCCESS' } else { 'INFO' })
        Write-Log -Message "  Note: KEK string matching is heuristic — OEM implementations vary" -Level "INFO"
    }
    catch { Write-Log -Message "  KEK: Unable to query — $($_.Exception.Message)" -Level "WARNING" }
    Write-Log -Message "--- End KEK Verification ---" -Level "INFO"

    # ── OS info and uptime ─────────────────────────────────────────────────
    $lastBoot    = $null
    $uptime      = $null
    $lastBootStr = "Unknown"
    $uptimeDays  = "?"
    if ($osInfo) {
        $lastBoot    = $osInfo.LastBootUpTime
        $uptime      = (Get-Date) - $lastBoot
        $lastBootStr = $lastBoot.ToString('yyyy-MM-dd HH:mm:ss')
        $uptimeDays  = [Math]::Floor($uptime.TotalDays)
        Write-Log -Message "OS: $($osInfo.Caption) (Build $osBuildInt)" -Level "INFO"
        if ($osInfo.Caption -like "*Windows 10*") {
            Write-Log -Message "WARNING: Windows 10 support ended October 2025 — upgrade to Windows 11 or ESU" -Level "WARNING"
        }
        Write-Log -Message "Last Boot: $lastBootStr | Uptime: $($uptimeDays)d $($uptime.Hours)h $($uptime.Minutes)m" -Level "INFO"
    }

    # ── TPM ────────────────────────────────────────────────────────────────
    try {
        $tpm = Get-CimInstance -Namespace "root\cimv2\Security\MicrosoftTpm" `
            -ClassName Win32_Tpm -ErrorAction Stop
        if ($tpm) {
            Write-Log -Message "TPM: Present | Enabled: $($tpm.IsEnabled_InitialValue) | Activated: $($tpm.IsActivated_InitialValue) | Spec: $($tpm.SpecVersion)" -Level "INFO"
        }
        else { Write-Log -Message "TPM: Not found" -Level "WARNING" }
    }
    catch { Write-Log -Message "TPM: Unable to query — $($_.Exception.Message)" -Level "WARNING" }

    # ── BitLocker ──────────────────────────────────────────────────────────
    try {
        $blVolume = Get-CimInstance -Namespace "root\cimv2\Security\MicrosoftVolumeEncryption" `
            -ClassName Win32_EncryptableVolume `
            -Filter "DriveLetter='$env:SystemDrive'" -ErrorAction Stop
        if ($blVolume) {
            $blProtection = switch ($blVolume.ProtectionStatus) {
                0 { "OFF" } 1 { "ON" } 2 { "UNKNOWN" } default { "Unknown ($($blVolume.ProtectionStatus))" }
            }
            $blConversion = switch ($blVolume.ConversionStatus) {
                0 { "FullyDecrypted" } 1 { "FullyEncrypted" } 2 { "EncryptionInProgress" }
                3 { "DecryptionInProgress" } 4 { "EncryptionPaused" } 5 { "DecryptionPaused" }
                default { "Unknown ($($blVolume.ConversionStatus))" }
            }
            Write-Log -Message "BitLocker ($env:SystemDrive): Protection=$blProtection | Status=$blConversion" -Level "INFO"
            if ($blProtection -eq "ON") {
                Write-Log -Message "BitLocker NOTE: Secure Boot cert changes may trigger recovery key prompt on next reboot" -Level "WARNING"
            }
        }
    }
    catch { Write-Log -Message "BitLocker: Unable to query — $($_.Exception.Message)" -Level "WARNING" }

    # ── Windows Update service ─────────────────────────────────────────────
    try {
        $wuService = Get-Service -Name wuauserv -ErrorAction Stop
        Write-Log -Message "Windows Update Service: Status=$($wuService.Status) | StartType=$($wuService.StartType)" -Level "INFO"
        if ($wuService.Status -ne 'Running' -and $wuService.Status -ne 'Stopped') {
            Write-Log -Message "WU Service WARNING: Unexpected state '$($wuService.Status)'" -Level "WARNING"
        }
    }
    catch { Write-Log -Message "Windows Update Service: Unable to query — $($_.Exception.Message)" -Level "WARNING" }

    # ── WU last scan / install ─────────────────────────────────────────────
    try {
        $autoUpdate = New-Object -ComObject Microsoft.Update.AutoUpdate -ErrorAction Stop
        $lastSearch = $autoUpdate.Results.LastSearchSuccessDate
        if ($lastSearch -and $lastSearch.Year -gt 2000) {
            $searchAge = (Get-Date) - $lastSearch
            Write-Log -Message "Last WU Scan: $($lastSearch.ToString('yyyy-MM-dd HH:mm:ss')) ($([Math]::Floor($searchAge.TotalHours))h ago)" -Level "INFO"
            if ($searchAge.TotalDays -gt 7) {
                Write-Log -Message "WU Scan WARNING: Last successful scan was over 7 days ago" -Level "WARNING"
            }
        }
        else { Write-Log -Message "Last WU Scan: No successful scan on record" -Level "WARNING" }
        $lastInstall = $autoUpdate.Results.LastInstallationSuccessDate
        if ($lastInstall -and $lastInstall.Year -gt 2000) {
            Write-Log -Message "Last WU Install: $($lastInstall.ToString('yyyy-MM-dd HH:mm:ss'))" -Level "INFO"
        }
    }
    catch { Write-Log -Message "Windows Update COM: Unable to query — $($_.Exception.Message)" -Level "WARNING" }

    # ── Pending reboot ─────────────────────────────────────────────────────
    $pendingRebootReasons = @()
    if (Test-Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending") {
        $pendingRebootReasons += "CBS-RebootPending"
    }
    if (Test-Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired") {
        $pendingRebootReasons += "WU-RebootRequired"
    }
    if ((Get-ItemProperty "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager" `
            -Name "PendingFileRenameOperations" -ErrorAction SilentlyContinue).PendingFileRenameOperations) {
        $pendingRebootReasons += "PendingFileRename"
    }
    if (Test-Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\PostRebootReporting") {
        $pendingRebootReasons += "WU-PostRebootReporting"
    }
    Write-Log -Message "Pending Reboot: $(if ($pendingRebootReasons.Count -gt 0) {"YES — $($pendingRebootReasons -join ', ')"} else {'None detected'})" `
        -Level $(if ($pendingRebootReasons.Count -gt 0) { 'WARNING' } else { 'INFO' })

    # ── Full registry dump ─────────────────────────────────────────────────
    Write-Log -Message "--- Secure Boot Registry Dump ---" -Level "INFO"
    foreach ($dumpPath in @($regPath, $servicingPath)) {
        if (Test-Path $dumpPath) {
            $props = Get-ItemProperty $dumpPath -ErrorAction SilentlyContinue
            $props.PSObject.Properties | Where-Object { $_.Name -notlike "PS*" } | ForEach-Object {
                $val = if ($_.Value -is [int])        { "$($_.Value) (0x$($_.Value.ToString('X')))" }
                       elseif ($_.Value -is [byte[]]) { "[byte[]] Length=$($_.Value.Length)" }
                       else                           { "$($_.Value)" }
                Write-Log -Message "  $($dumpPath.Split('\')[-1])\$($_.Name) = $val" -Level "INFO"
            }
        }
    }
    Write-Log -Message "--- End Registry Dump ---" -Level "INFO"

    # ── Event log — Kernel-Boot/Operational only ───────────────────────────
    Write-Log -Message "--- Secure Boot Event Log ---" -Level "INFO"
    $sbEventIds = @(1036, 1043, 1044, 1045, 1795, 1801, 1808)
    try {
        $sbEvents = Get-WinEvent -FilterHashtable @{
            LogName = 'Microsoft-Windows-Kernel-Boot/Operational'
            Id      = $sbEventIds
        } -MaxEvents 20 -ErrorAction SilentlyContinue
        if ($sbEvents -and $sbEvents.Count -gt 0) {
            $sbEvents | Group-Object -Property Id | ForEach-Object {
                $latest     = $_.Group | Sort-Object TimeCreated -Descending | Select-Object -First 1
                $msgPreview = ($latest.Message -split "`n")[0]
                if ($msgPreview.Length -gt 120) { $msgPreview = $msgPreview.Substring(0, 120) + "..." }
                Write-Log -Message "  Event $($_.Name): Last=$($latest.TimeCreated.ToString('yyyy-MM-dd HH:mm:ss')) Count=$($_.Count) [$msgPreview]" -Level "INFO"
            }
        }
        else {
            Write-Log -Message "  No Secure Boot events (IDs: $($sbEventIds -join ',')) found" -Level "INFO"
        }
    }
    catch { Write-Log -Message "  Event Log query failed: $($_.Exception.Message)" -Level "WARNING" }
    Write-Log -Message "--- End Event Log ---" -Level "INFO"
    Write-Log -Message "---------- END DIAGNOSTIC DATA ----------" -Level "INFO"

    # Append payload / task / WinCS / fallback to detail string
    if ($payload.State -ne 'Healthy') { $detailList.Add("Payload:$($payload.State)") }
    if ($taskStatus.TaskExists -and $taskStatus.IsMissingFiles) { $detailList.Add("Task:0x80070002") }
    elseif ($taskStatus.TaskExists -and $taskStatus.LastTaskResult -eq 0) { $detailList.Add("Task:OK") }
    if ($winCsAvailable) { $detailList.Add("WinCS:Available") }
    if ($fallback.TimestampExists) {
        $detailList.Add($(if ($fallback.IsActive) { "Fallback:ACTIVE($($fallback.DaysElapsed)d)" }
                          else { "Fallback:$($fallback.DaysRemaining)d remaining" }))
    }
    $details = $detailList -join " | "

    # ══════════════════════════════════════════════════════════════════════
    # STAGE 4 — CA2023 in DB, awaiting reboot
    # ══════════════════════════════════════════════════════════════════════
    if ($null -ne $ca2023Capable -and [int]$ca2023Capable -eq 1) {
        Write-Log -Message "--- Stage 4 Analysis ---" -Level "INFO"
        Write-Log -Message "  WHY: CA2023 cert written to UEFI DB but device has not rebooted to load new boot manager" -Level "INFO"
        Write-Log -Message "  Last Boot: $lastBootStr ($uptimeDays days ago)" -Level "INFO"
        if ($pendingRebootReasons.Count -gt 0) {
            Write-Log -Message "  Pending Reboot: $($pendingRebootReasons -join ', ')" -Level "WARNING"
        }
        else {
            Write-Log -Message "  Pending Reboot: No indicators — reboot still required to activate new boot manager" -Level "INFO"
        }
        Write-Log -Message "  NEXT STEPS: Reboot device. If still Stage 4 after reboot, verify BitLocker recovery key." -Level "WARNING"
        if ($fallback.TimestampExists -and $fallback.IsActive) {
            Write-Log -Message "  FALLBACK: Timer exceeded ($($fallback.DaysElapsed)d > $($FallbackDays)d) — remediation will use direct method" -Level "WARNING"
        }
        elseif ($fallback.TimestampExists) {
            Write-Log -Message "  FALLBACK: $($fallback.DaysRemaining) days until direct method fallback activates" -Level "INFO"
        }
        Write-Log -Message "--- End Stage 4 Analysis ---" -Level "INFO"
        Write-Host "CONFIGURED_CA2023_IN_DB | $details | Action: Reboot to complete transition"
        Write-Log -Message "Detection Result: NON-COMPLIANT - Stage 4 (exit 1)" -Level "WARNING"
        Write-Log -Message "========== DETECTION COMPLETED ==========" -Level "INFO"
        Set-ExitCode 1
        return
    }

    # ══════════════════════════════════════════════════════════════════════
    # STAGE 3b — AvailableUpdates cleared, cert DB write pending
    # ══════════════════════════════════════════════════════════════════════
    if ($null -ne $availableUpdatesValue -and $availableUpdatesValue -eq 0 -and $null -eq $uefiCA2023Status) {
        Write-Log -Message "--- Stage 3b Analysis ---" -Level "INFO"
        Write-Log -Message "  WHY: AvailableUpdates cleared (task completed) but UEFICA2023Status not yet written" -Level "INFO"
        Write-Log -Message "  WHY: Transient state — cert DB write is pending" -Level "INFO"
        Write-Log -Message "  NEXT STEPS: Allow Windows Update to complete the cert DB write" -Level "INFO"
        Write-Log -Message "--- End Stage 3b Analysis ---" -Level "INFO"
        Write-Host "CONFIGURED_PROCESSING_COMPLETE | $details | Action: Awaiting cert DB write"
        Write-Log -Message "Detection Result: NON-COMPLIANT - Stage 3b (exit 1)" -Level "WARNING"
        Write-Log -Message "========== DETECTION COMPLETED ==========" -Level "INFO"
        Set-ExitCode 1
        return
    }

    # ══════════════════════════════════════════════════════════════════════
    # STAGE 3 — Certificate updates actively in progress
    # ══════════════════════════════════════════════════════════════════════
    if ($null -ne $availableUpdatesValue -and $availableUpdatesValue -ne 0 -and $availableUpdatesValue -ne 22852) {
        Write-Log -Message "--- Stage 3 Analysis ---" -Level "INFO"
        Write-Log -Message "  WHY: Windows Update actively processing Secure Boot certificate updates" -Level "INFO"
        Write-Log -Message "  WHY: AvailableUpdates = 0x$($availableUpdatesValue.ToString('X')) ($availableUpdatesValue)" -Level "INFO"
        Write-Log -Message "  WHY: Target value is 0x4000 (16384) = all certificates applied" -Level "INFO"
        if ($availableUpdatesValue -eq 0x4104) {
            Write-Log -Message "  STUCK STATE: 0x4104 — device cannot progress past KEK cert deployment" -Level "ERROR"
            Write-Log -Message "  ROOT CAUSE: 0x0004 bit (KEK) not clearing — check OEM for firmware update" -Level "ERROR"
            Write-Log -Message "  Reference: https://learn.microsoft.com/windows-hardware/manufacture/desktop/windows-secure-boot-key-creation-and-management-guidance" -Level "ERROR"
        }
        if ($pendingRebootReasons.Count -gt 0) {
            Write-Log -Message "  Pending Reboot: $($pendingRebootReasons -join ', ') — reboot may be needed to continue" -Level "WARNING"
        }
        Write-Log -Message "  NEXT STEPS: Allow Windows Update to complete (typically 1-2 quality update cycles)" -Level "INFO"
        Write-Log -Message "  NEXT STEPS: If stuck >30 days run 'usoclient StartScan' or check WSUS/WU policy" -Level "WARNING"
        if ($fallback.TimestampExists -and $fallback.IsActive) {
            Write-Log -Message "  FALLBACK: Timer exceeded ($($fallback.DaysElapsed)d > $($FallbackDays)d) — remediation will use direct method" -Level "WARNING"
        }
        elseif ($fallback.TimestampExists) {
            Write-Log -Message "  FALLBACK: $($fallback.DaysRemaining) days until direct method fallback activates" -Level "INFO"
        }
        Write-Log -Message "--- End Stage 3 Analysis ---" -Level "INFO"
        Write-Host "CONFIGURED_UPDATE_IN_PROGRESS | $details | Action: Waiting for Windows Update"
        Write-Log -Message "Detection Result: NON-COMPLIANT - Stage 3 (exit 1)" -Level "WARNING"
        Write-Log -Message "========== DETECTION COMPLETED ==========" -Level "INFO"
        Set-ExitCode 1
        return
    }

    # ══════════════════════════════════════════════════════════════════════
    # STAGE 2 — Deployment triggered, WU scan not yet started
    # ══════════════════════════════════════════════════════════════════════
    Write-Log -Message "--- Stage 2 Analysis ---" -Level "INFO"
    Write-Log -Message "  WHY: AvailableUpdates set but Secure-Boot-Update task has not yet processed" -Level "INFO"
    if ($null -eq $availableUpdatesValue) {
        Write-Log -Message "  WHY: AvailableUpdates key does not exist — Windows Update has not yet scanned" -Level "INFO"
    }
    else {
        Write-Log -Message "  WHY: AvailableUpdates = 0x$($availableUpdatesValue.ToString('X')) — waiting for WU to begin" -Level "INFO"
    }
    if ($null -eq $ca2023Capable) {
        Write-Log -Message "  WHY: WindowsUEFICA2023Capable not present — normal before WU processes" -Level "INFO"
    }
    Write-Log -Message "  NEXT STEPS: Ensure device is internet connected and Windows Update service is running" -Level "INFO"
    Write-Log -Message "  NEXT STEPS: If stuck >14 days run 'usoclient StartScan' or check WU policy/WSUS" -Level "WARNING"
    if ($fallback.TimestampExists -and $fallback.IsActive) {
        Write-Log -Message "  FALLBACK: Timer exceeded ($($fallback.DaysElapsed)d > $($FallbackDays)d) — remediation will use direct method" -Level "WARNING"
    }
    elseif ($fallback.TimestampExists) {
        Write-Log -Message "  FALLBACK: $($fallback.DaysRemaining) days until direct method fallback activates" -Level "INFO"
    }
    Write-Log -Message "--- End Stage 2 Analysis ---" -Level "INFO"
    Write-Host "CONFIGURED_AWAITING_UPDATE | $details | Action: Waiting for Windows Update scan"
    Write-Log -Message "Detection Result: NON-COMPLIANT - Stage 2 (exit 1)" -Level "WARNING"
    Write-Log -Message "========== DETECTION COMPLETED ==========" -Level "INFO"
    Set-ExitCode 1
    return
}
catch {
    Write-Log -Message "Unexpected error: $($_.Exception.Message)" -Level "ERROR"
    Write-Log -Message "Stack Trace: $($_.ScriptStackTrace)" -Level "ERROR"
    Write-Host "ERROR: $($_.Exception.Message)"
    Write-Log -Message "Detection Result: ERROR (exit 1)" -Level "ERROR"
    Write-Log -Message "========== DETECTION COMPLETED ==========" -Level "INFO"
    Set-ExitCode 1
    return
}
#endregion