<#
.SYNOPSIS
    Remediates Windows devices to enable Secure Boot certificate updates before June 2026 expiration.

.DESCRIPTION
    This remediation script deploys Secure Boot certificate updates before the June 2026 expiration
    using the AvailableUpdates registry key (Microsoft playbook Option 2 - direct deployment).
    Microsoft Corporation KEK CA 2011 and Microsoft Corporation UEFI CA 2011 certificates expire in June 2026.
    Microsoft Windows Production PCA 2011 expires in October 2026.
    
    The script performs the following actions:
    - Checks UEFICA2023Status, WindowsUEFICA2023Capable, and 24H2+ OS image path for existing compliance
    - Checks UEFICA2023Error for previous deployment errors
    - Sets AvailableUpdates to 0x5944 to deploy all certificates and 2023 boot manager
    - Triggers the Secure-Boot-Update scheduled task to begin processing immediately
    - Writes a ManagedOptInDate timestamp for fallback timer tracking
    - Outputs detailed status for Intune logging
    
    After setting AvailableUpdates, the Secure-Boot-Update task processes the bits every 12 hours.
    One or more reboots may be required to fully complete the certificate transition.
    
    Fallback Timer: If the device has been configured for more than FallbackDays days without
    reaching full compliance, the script automatically falls back to a direct method. The preferred
    fallback is WinCS (WinCsFlags.exe /apply) which bypasses the SecureBootUpdates payload folder
    entirely. If WinCS is not available, the legacy step-by-step AvailableUpdates method (0x40 then
    0x100) is used, provided the payload files exist in C:\Windows\System32\SecureBootUpdates\.
    If neither WinCS nor payloads are available, the script exits with an error indicating a
    cumulative update is needed.
    
    Exit codes:
    - 0: Remediation successful (or already configured)
    - 1: Remediation failed

.PARAMETER FallbackDays
    Number of days to wait after managed opt-in before falling back to the direct
    AvailableUpdates method if the device has not reached compliance. Default: 30

.PARAMETER TimestampRegPath
    Registry path where the ManagedOptInDate timestamp is stored.
    Default: HKLM:\SOFTWARE\Mindcore\Secureboot

.EXAMPLE
    .\Remediate-SecureBootCertificateUpdate.ps1
    
    Configures the device to receive Secure Boot certificate updates and verifies the configuration.

.EXAMPLE
    .\Remediate-SecureBootCertificateUpdate.ps1 -FallbackDays 45 -TimestampRegPath "HKLM:\SOFTWARE\Contoso\Secureboot"
    
    Uses a 45-day fallback threshold and a custom registry path for the opt-in timestamp.

.NOTES
    Version:        6.0
    Author:         Mattias Melkersen
    Creation Date:  2026-01-15
    
    CHANGELOG
    ---------------
    2026-04-15 - v6.0 - Added 64-bit process enforcement guard (parity with detection v6.0)
                        Added ISE-safe Set-ExitCode helper (return instead of exit in ISE)
                        Aligned compliance check with detection v6.1 three-path model:
                          Path A: WindowsUEFICA2023Capable = 2 (authoritative firmware attestation)
                          Path B: UEFICA2023Status = Updated (pre-24H2 servicing pipeline)
                          Path C: Build 26100+ + CA2023 in UEFI DB + bootmgfw.efi >= 10.0.19041.4522
                        Fixed Remove-Item -Recurse on TimestampRegPath -> Remove-ItemProperty (scoped)
                        Fixed [datetime]::Parse() -> [DateTime]::TryParse() for locale safety
                        Changed firmware age gate from blocking to WARNING (too aggressive for fleet deployment)
                        Added Get-BootManagerCompliance helper (parity with detection v6.0)
    2026-04-13 - v5.0 - Aligned with Microsoft Secure Boot playbook (MM)
                        Switched primary deployment from MicrosoftUpdateManagedOptIn (CFR) to AvailableUpdates = 0x5944 (Option 2)
                        Added UEFICA2023Status = Updated as sole compliance check
                        Added UEFICA2023Error and UEFICA2023ErrorEvent pre-check before remediation
                        Triggers Secure-Boot-Update scheduled task immediately after setting AvailableUpdates
                        Idempotency now based on timestamp + AvailableUpdates state
                        Accelerates legacy CFR-configured devices by setting AvailableUpdates if not present
    2026-04-13 - v4.1 - Fixed PS 5.1 compatibility: replaced ?? null-coalescing operators in Get-FirmwareAgeStatus (MM)
    2026-04-06 - v4.0 - WinCS (WinCsFlags.exe /apply) is now the preferred fallback method over AvailableUpdates (MM)
                        Added SecureBootUpdates payload folder pre-flight validation before task trigger
                        Legacy AvailableUpdates path retained only when WinCS unavailable and payloads present
                        Added post-task LastTaskResult validation to detect 0x80070002 in legacy path
                        Devices without WinCS or payloads now exit 1 with explicit "needs cumulative update" message
    2026-04-05 - v3.1 - Buffered logging (MM)
    2026-03-27 - v3.0 - Added configurable fallback timer with FallbackDays and TimestampRegPath parameters (MM)
    2026-02-18 - v2.1 - Enhanced logging in ALREADY_CONFIGURED path (MM)
    2026-02-18 - v2.0 - Made idempotent: skip registry write if OptIn already set correctly (MM)
    2026-01-15 - v1.0 - Initial version for June 2026 certificate expiration preparation (MM)
    
    References:
    - https://aka.ms/getsecureboot
    - https://techcommunity.microsoft.com/blog/windows-itpro-blog/act-now-secure-boot-certificates-expire-in-june-2026/4426856
    - https://support.microsoft.com/topic/enterprise-deployment-guidance-for-cve-2023-24932-88b8f034-20b7-4a45-80cb-c6049b0f9967
    
    Prerequisites:
    - Secure Boot must be enabled in BIOS/UEFI (this script cannot enable it)
    - Latest OEM firmware should be installed before certificate updates apply
    - Device must receive Windows Updates from Microsoft
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

#region Logging Configuration
[string]$LogFile = "$env:ProgramData\Microsoft\IntuneManagementExtension\Logs\SecureBootCertificateUpdate.log"
[string]$ScriptName = "REMEDIATE"
[int]$MaxLogSizeMB = 4
$script:LogBuffer = [System.Collections.Generic.List[string]]::new()

function Write-Log {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,
        [Parameter(Mandatory = $false)]
        [ValidateSet("INFO", "WARNING", "ERROR", "SUCCESS")]
        [string]$Level = "INFO"
    )
    $TimeStamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $script:LogBuffer.Add("$TimeStamp [$ScriptName] [$Level] $Message")
}

function Flush-Log {
    if ($script:LogBuffer.Count -eq 0) { return }
    try {
        $LogDir = Split-Path -Path $LogFile -Parent
        if (-not (Test-Path $LogDir)) {
            New-Item -Path $LogDir -ItemType Directory -Force | Out-Null
        }
        if (Test-Path $LogFile) {
            $LogFileSizeMB = (Get-Item $LogFile).Length / 1MB
            if ($LogFileSizeMB -ge $MaxLogSizeMB) {
                $BackupLog = "$LogFile.old"
                if (Test-Path $BackupLog) {
                    Remove-Item -Path $BackupLog -Force -ErrorAction SilentlyContinue
                }
                Rename-Item -Path $LogFile -NewName $BackupLog -Force -ErrorAction SilentlyContinue
                $TimeStamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                $script:LogBuffer.Insert(0, "$TimeStamp [SYSTEM] [INFO] Log rotated - Previous log archived to: $BackupLog")
            }
        }
        Add-Content -Path $LogFile -Value $script:LogBuffer.ToArray() -ErrorAction SilentlyContinue
        $script:LogBuffer.Clear()
    }
    catch { }
}
#endregion

#region Functions
function Get-SecureBootStatus {
    try {
        return Confirm-SecureBootUEFI -ErrorAction Stop
    }
    catch {
        try {
            $regPath = "HKLM:\SYSTEM\CurrentControlSet\Control\SecureBoot\State"
            if (Test-Path $regPath) {
                $value = (Get-ItemProperty -Path $regPath -Name "UEFISecureBootEnabled" -ErrorAction SilentlyContinue).UEFISecureBootEnabled
                return ($value -eq 1)
            }
        }
        catch { }
    }
    return $false
}

function Get-FirmwareAgeStatus {
    $result = @{
        Manufacturer    = "Unknown"
        Model           = "Unknown"
        BiosVersion     = "Unknown"
        ReleaseDate     = $null
        AgeDays         = $null
        IsStale         = $false
        UpdateGuidance  = "Check your device manufacturer's support site for firmware updates."
    }

    try {
        $bios = Get-CimInstance -ClassName Win32_BIOS -ErrorAction Stop
        $cs = Get-CimInstance -ClassName Win32_ComputerSystem -ErrorAction Stop
        $result.Manufacturer = $(if ($cs.Manufacturer) { $cs.Manufacturer } else { "Unknown" }).Trim()
        $result.Model = $(if ($cs.Model) { $cs.Model } else { "Unknown" }).Trim()
        $result.BiosVersion = $bios.SMBIOSBIOSVersion

        if ($null -ne $bios.ReleaseDate) {
            $result.ReleaseDate = $bios.ReleaseDate
            $result.AgeDays = [Math]::Floor(((Get-Date) - $bios.ReleaseDate).TotalDays)
            $result.IsStale = ($result.AgeDays -gt 365)
        }

        # OEM-specific update guidance
        $mfr = $result.Manufacturer.ToUpper()
        switch -Wildcard ($mfr) {
            "LENOVO*" {
                $result.UpdateGuidance = "Update firmware via Lenovo Vantage, SCCM driver packs, or https://support.lenovo.com"
            }
            "DELL*" {
                $result.UpdateGuidance = "Update firmware via Dell Command Update, Dell Support Assist, or https://www.dell.com/support"
            }
            { $_ -like "HP*" -or $_ -like "HEWLETT*" } {
                $result.UpdateGuidance = "Update firmware via HP Support Assistant, HP Image Assistant, or https://support.hp.com"
            }
            "MICROSOFT*" {
                $result.UpdateGuidance = "Update firmware via Windows Update (Surface drivers) or https://www.microsoft.com/surface/support"
            }
        }
    }
    catch { }

    return $result
}

function Get-SecureBootPayloadStatus {
    $payloadPath = "$env:SystemRoot\System32\SecureBootUpdates"
    $result = @{
        FolderExists  = $false
        FileCount     = 0
        Files         = @()
        HasBinFiles   = $false
        IsHealthy     = $false
    }

    try {
        if (Test-Path $payloadPath) {
            $result.FolderExists = $true
            $files = Get-ChildItem -Path $payloadPath -File -ErrorAction SilentlyContinue
            if ($files) {
                $result.FileCount = $files.Count
                $result.Files = $files | ForEach-Object { "$($_.Name) ($([Math]::Round($_.Length / 1KB, 1))KB)" }
                $result.HasBinFiles = ($files | Where-Object { $_.Extension -eq '.bin' }).Count -gt 0
                $result.IsHealthy = $result.HasBinFiles
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

#region Main Remediation Logic
try {
    Write-Log -Message "========== REMEDIATION STARTED ==========" -Level "INFO"
    Write-Log -Message "Script Version: 6.0" -Level "INFO"
    Write-Log -Message "Computer: $env:COMPUTERNAME | User: $env:USERNAME" -Level "INFO"
    Write-Log -Message "PowerShell: $($PSVersionTable.PSVersion) | Process: $(if ([Environment]::Is64BitProcess) {'64-bit'} else {'32-bit'})" -Level "INFO"
    
    # ── Secure Boot state ──────────────────────────────────────────────────
    Write-Log -Message "Verifying Secure Boot status..." -Level "INFO"
    $secureBootEnabled = Get-SecureBootStatus
    
    if (-not $secureBootEnabled) {
        Write-Log -Message "Secure Boot is DISABLED - Cannot apply remediation" -Level "ERROR"
        Write-Log -Message "Action Required: Enable Secure Boot in BIOS/UEFI firmware settings manually" -Level "ERROR"
        Write-Host "FAILED: Secure Boot DISABLED - Enable in BIOS/UEFI manually"
        Write-Log -Message "Remediation Result: FAILED (exit 1)" -Level "ERROR"
        Write-Log -Message "========== REMEDIATION COMPLETED ==========" -Level "INFO"
        Set-ExitCode 1
        return
    }
    Write-Log -Message "Secure Boot is ENABLED - Proceeding with remediation" -Level "SUCCESS"
    
    # ── Firmware age check (warning only — does not block) ─────────────────
    Write-Log -Message "Checking firmware age..." -Level "INFO"
    $firmware = Get-FirmwareAgeStatus
    Write-Log -Message "  Manufacturer: $($firmware.Manufacturer)" -Level "INFO"
    Write-Log -Message "  Model: $($firmware.Model)" -Level "INFO"
    Write-Log -Message "  BIOS Version: $($firmware.BiosVersion)" -Level "INFO"
    if ($null -ne $firmware.ReleaseDate) {
        Write-Log -Message "  Firmware Release Date: $($firmware.ReleaseDate.ToString('yyyy-MM-dd'))" -Level "INFO"
        Write-Log -Message "  Firmware Age: $($firmware.AgeDays) days" -Level "INFO"
    }
    else {
        Write-Log -Message "  Firmware Release Date: Unable to determine" -Level "WARNING"
    }

    if ($firmware.IsStale) {
        Write-Log -Message "FIRMWARE WARNING: Firmware is $($firmware.AgeDays) days old (>365 days)" -Level "WARNING"
        Write-Log -Message "  Outdated firmware increases risk of boot issues after Secure Boot cert updates" -Level "WARNING"
        Write-Log -Message "  $($firmware.UpdateGuidance)" -Level "WARNING"
        Write-Log -Message "  Proceeding with remediation — update firmware at earliest opportunity" -Level "WARNING"
    }
    else {
        Write-Log -Message "Firmware age check: PASSED" -Level "SUCCESS"
    }
    
    # ── Pre-check: UEFICA2023Error ─────────────────────────────────────────
    $servicingPath = "HKLM:\SYSTEM\CurrentControlSet\Control\SecureBoot\Servicing"
    $svcProps      = Get-ItemProperty -Path $servicingPath -ErrorAction SilentlyContinue
    $uefiError     = $svcProps.UEFICA2023Error
    $uefiErrorEvent = $svcProps.UEFICA2023ErrorEvent
    if ($null -ne $uefiError -and $uefiError -ne 0) {
        Write-Log -Message "WARNING: UEFICA2023Error detected: 0x$($uefiError.ToString('X')) ($uefiError)" -Level "WARNING"
        if ($null -ne $uefiErrorEvent) {
            Write-Log -Message "  UEFICA2023ErrorEvent: $uefiErrorEvent (check Windows System Event Log)" -Level "WARNING"
        }
        Write-Log -Message "  A previous Secure Boot certificate update attempt encountered an error" -Level "WARNING"
        Write-Log -Message "  Reference: https://support.microsoft.com/topic/37e47cf8-608b-4a87-8175-bdead630eb69" -Level "WARNING"
    }
    
    # ══════════════════════════════════════════════════════════════════════
    # COMPLIANCE CHECK — three paths, aligned with detection script v6.1
    # ══════════════════════════════════════════════════════════════════════
    $uefiCA2023Status = $svcProps.UEFICA2023Status
    $ca2023Capable    = $svcProps.WindowsUEFICA2023Capable

    # ── Path A: WindowsUEFICA2023Capable = 2 (authoritative firmware attestation) ──
    if ($null -ne $ca2023Capable -and [int]$ca2023Capable -eq 2) {
        Write-Log -Message "Device is ALREADY COMPLIANT — Path A: WindowsUEFICA2023Capable=2 (firmware attestation)" -Level "SUCCESS"
        if ($null -ne $uefiCA2023Status) {
            Write-Log -Message "  UEFICA2023Status: $uefiCA2023Status" -Level "INFO"
        }
        Remove-TrackingTimestamp -RegPath $TimestampRegPath
        Write-Host "ALREADY_COMPLIANT | Path:FirmwareAttestation | CA2023Capable=2"
        Write-Log -Message "Remediation Result: ALREADY_COMPLIANT (exit 0)" -Level "SUCCESS"
        Write-Log -Message "========== REMEDIATION COMPLETED ==========" -Level "INFO"
        Set-ExitCode 0
        return
    }

    # ── Path B: UEFICA2023Status = "Updated" (pre-24H2 servicing pipeline) ──
    if ($uefiCA2023Status -eq "Updated") {
        Write-Log -Message "Device is ALREADY COMPLIANT — Path B: UEFICA2023Status=Updated (servicing pipeline)" -Level "SUCCESS"
        if ($null -ne $ca2023Capable) {
            Write-Log -Message "  WindowsUEFICA2023Capable: $ca2023Capable" -Level "INFO"
        }
        Remove-TrackingTimestamp -RegPath $TimestampRegPath
        Write-Host "ALREADY_COMPLIANT | Path:ServicingUpdated | UEFICA2023Status=Updated$(if ($null -ne $ca2023Capable) {" | CA2023Capable=$ca2023Capable"})"
        Write-Log -Message "Remediation Result: ALREADY_COMPLIANT (exit 0)" -Level "SUCCESS"
        Write-Log -Message "========== REMEDIATION COMPLETED ==========" -Level "INFO"
        Set-ExitCode 0
        return
    }

    # ── Path C: Build 26100+ + CA2023 in UEFI DB + boot manager version ──
    $osInfo     = Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction SilentlyContinue
    $osBuildInt = if ($osInfo) { [int]$osInfo.BuildNumber } else { 0 }

    if ($osBuildInt -ge 26100) {
        $hasCA2023InDB = $false
        try {
            $dbBytes       = (Get-SecureBootUEFI db -ErrorAction Stop).bytes
            $dbText        = [System.Text.Encoding]::Unicode.GetString($dbBytes) +
                             [System.Text.Encoding]::ASCII.GetString($dbBytes)
            $hasCA2023InDB = $dbText -match 'Windows UEFI CA 2023'
        }
        catch {
            Write-Log -Message "UEFI DB query failed: $($_.Exception.Message)" -Level "WARNING"
        }

        $bootMgr = Get-BootManagerCompliance

        if ($hasCA2023InDB -and $bootMgr.Compliant) {
            Write-Log -Message "Device is ALREADY COMPLIANT — Path C: 24H2+ OS image path" -Level "SUCCESS"
            Write-Log -Message "  Build: $osBuildInt | CA2023: In DB | BootMgr: $($bootMgr.CleanVersion)" -Level "SUCCESS"
            Remove-TrackingTimestamp -RegPath $TimestampRegPath
            Write-Host "ALREADY_COMPLIANT | Path:24H2OsImage | Build:$osBuildInt | CA2023:InDB | BootMgr:$($bootMgr.CleanVersion)"
            Write-Log -Message "Remediation Result: ALREADY_COMPLIANT (exit 0)" -Level "SUCCESS"
            Write-Log -Message "========== REMEDIATION COMPLETED ==========" -Level "INFO"
            Set-ExitCode 0
            return
        }
        else {
            Write-Log -Message "24H2+ device (Build $osBuildInt) — Path C NOT met: CA2023InDB=$hasCA2023InDB | BootMgrOK=$($bootMgr.Compliant)" -Level "WARNING"
        }
    }
    
    # Log informational note when DB key is present but full transition is not complete
    if ($null -ne $ca2023Capable -and [int]$ca2023Capable -ge 1 -and $uefiCA2023Status -ne "Updated") {
        Write-Log -Message "NOTE: WindowsUEFICA2023Capable=$ca2023Capable but UEFICA2023Status=$uefiCA2023Status - DB key present but transition incomplete" -Level "WARNING"
    }
    
    # ══════════════════════════════════════════════════════════════════════
    # REMEDIATION — device is not compliant
    # ══════════════════════════════════════════════════════════════════════

    # Define registry configuration
    $regPath = "HKLM:\SYSTEM\CurrentControlSet\Control\Secureboot"
    $regName = "AvailableUpdates"
    $regValue = 0x5944  # Deploy all certificates and 2023 boot manager (per Microsoft playbook Option 2)
    $regType = "DWord"
    
    # ── Idempotency check: skip if deployment already triggered ────────────
    $existingValue = $null
    if (Test-Path $regPath) {
        $existingValue = (Get-ItemProperty -Path $regPath -Name $regName -ErrorAction SilentlyContinue).$regName
    }
    
    # Check timestamp as durable marker of previous configuration (AvailableUpdates bits clear as task processes)
    $optInDateStr = $null
    if (Test-Path $TimestampRegPath) {
        $optInDateStr = (Get-ItemProperty -Path $TimestampRegPath -Name "ManagedOptInDate" -ErrorAction SilentlyContinue).ManagedOptInDate
    }
    
    # Deployment was triggered if: AvailableUpdates is non-zero OR timestamp exists from a prior run
    $deploymentTriggered = ($null -ne $existingValue -and $existingValue -ne 0) -or ($null -ne $optInDateStr)
    
    if ($deploymentTriggered) {
        $ca2023Text = switch ($ca2023Capable) {
            0 { "Not in DB" }
            1 { "In DB" }
            2 { "In DB and booting from 2023 cert" }
            default { "Pending" }
        }
        
        Write-Log -Message "--- Idempotency: Deployment Already Triggered ---" -Level "INFO"
        if ($null -ne $existingValue) {
            Write-Log -Message "  AvailableUpdates: 0x$($existingValue.ToString('X')) ($existingValue)" -Level "INFO"
        }
        else {
            Write-Log -Message "  AvailableUpdates: Key not present (bits fully processed or set by prior CFR method)" -Level "INFO"
        }
        Write-Log -Message "  CA2023Capable: $ca2023Capable ($ca2023Text)" -Level "INFO"
        if ($null -ne $uefiCA2023Status) {
            Write-Log -Message "  UEFICA2023Status: $uefiCA2023Status" -Level "INFO"
        }
        try {
            $remLastBoot = (Get-CimInstance -ClassName Win32_OperatingSystem).LastBootUpTime
            Write-Log -Message "  Last Boot: $($remLastBoot.ToString('yyyy-MM-dd HH:mm:ss'))" -Level "INFO"
        }
        catch {}
        Write-Log -Message "--- End Idempotency Check ---" -Level "INFO"
        
        # If AvailableUpdates was not set (legacy CFR path or bits fully cleared), set it now to accelerate
        if ($null -eq $existingValue -or $existingValue -eq 0) {
            Write-Log -Message "AvailableUpdates not set - setting to 0x$($regValue.ToString('X')) to accelerate deployment" -Level "INFO"
            Set-ItemProperty -Path $regPath -Name $regName -Value $regValue -Type $regType -Force -ErrorAction Stop
            $task = Get-ScheduledTask -TaskPath "\Microsoft\Windows\PI\" -TaskName "Secure-Boot-Update" -ErrorAction SilentlyContinue
            if ($task) {
                Start-ScheduledTask -TaskPath "\Microsoft\Windows\PI\" -TaskName "Secure-Boot-Update" -ErrorAction SilentlyContinue
                Write-Log -Message "Secure-Boot-Update task triggered to begin processing" -Level "SUCCESS"
            }
        }
        
        # Check fallback timer
        Write-Log -Message "--- Fallback Timer Check ---" -Level "INFO"
        
        if (-not $optInDateStr) {
            # Backfill timestamp for devices configured before v3.0 or by external methods
            Write-Log -Message "  ManagedOptInDate not found - backfilling timestamp (clock starts now)" -Level "WARNING"
            try {
                if (-not (Test-Path $TimestampRegPath)) {
                    New-Item -Path $TimestampRegPath -Force | Out-Null
                }
                $backfillDate = (Get-Date).ToString("o")
                Set-ItemProperty -Path $TimestampRegPath -Name "ManagedOptInDate" -Value $backfillDate -Type String -Force
                Write-Log -Message "  ManagedOptInDate backfilled: $backfillDate" -Level "INFO"
            }
            catch {
                Write-Log -Message "  WARNING: Could not backfill ManagedOptInDate: $($_.Exception.Message)" -Level "WARNING"
            }
            Write-Log -Message "--- End Fallback Timer Check ---" -Level "INFO"
            Write-Host "ALREADY_CONFIGURED: AvailableUpdates deployment triggered. CA2023: $ca2023Text. Fallback timer started."
            Write-Log -Message "Remediation Result: ALREADY_CONFIGURED (exit 0)" -Level "SUCCESS"
            Write-Log -Message "========== REMEDIATION COMPLETED ==========" -Level "INFO"
            Set-ExitCode 0
            return
        }
        
        # Timestamp exists - calculate elapsed days (locale-safe)
        $optInDate = $null
        if (-not [DateTime]::TryParse($optInDateStr, [ref]$optInDate)) {
            Write-Log -Message "  WARNING: Could not parse ManagedOptInDate '$optInDateStr' — resetting timestamp" -Level "WARNING"
            $optInDate = Get-Date
            Set-ItemProperty -Path $TimestampRegPath -Name "ManagedOptInDate" -Value ($optInDate.ToString("o")) -Type String -Force -ErrorAction SilentlyContinue
        }
        $daysElapsed = [Math]::Floor(((Get-Date) - $optInDate).TotalDays)
        $daysRemaining = [Math]::Max(0, $FallbackDays - $daysElapsed)
        Write-Log -Message "  ManagedOptInDate: $($optInDate.ToString('yyyy-MM-dd HH:mm:ss'))" -Level "INFO"
        Write-Log -Message "  Days Elapsed: $daysElapsed | Threshold: $FallbackDays | Remaining: $daysRemaining" -Level "INFO"
        
        if ($daysElapsed -lt $FallbackDays) {
            # Threshold not yet reached
            Write-Log -Message "  Fallback not yet active - $daysRemaining days remaining" -Level "INFO"
            Write-Log -Message "--- End Fallback Timer Check ---" -Level "INFO"
            Write-Host "ALREADY_CONFIGURED: AvailableUpdates deployment triggered. CA2023: $ca2023Text. Fallback in $($daysRemaining)d."
            Write-Log -Message "Remediation Result: ALREADY_CONFIGURED (exit 0)" -Level "SUCCESS"
            Write-Log -Message "========== REMEDIATION COMPLETED ==========" -Level "INFO"
            Set-ExitCode 0
            return
        }
        
        # --- FALLBACK: WinCS preferred, legacy AvailableUpdates as backup ---
        Write-Log -Message "--- FALLBACK: Activating direct method ---" -Level "WARNING"
        Write-Log -Message "  Deployment has been active for $daysElapsed days (threshold: $FallbackDays days)" -Level "WARNING"
        Write-Log -Message "  CA2023Capable: $ca2023Capable ($ca2023Text) - switching to direct method" -Level "WARNING"
        
        # Check WinCS availability (preferred method - bypasses payload folder dependency)
        $winCsPath = "$env:SystemRoot\System32\WinCsFlags.exe"
        $winCsAvailable = Test-Path $winCsPath
        
        if ($winCsAvailable) {
            # --- WinCS path (preferred) ---
            Write-Log -Message "  WinCsFlags.exe: AVAILABLE - using WinCS method (preferred)" -Level "SUCCESS"
            try {
                Write-Log -Message "  Applying WinCS key: WinCsFlags.exe /apply --key F33E0C8E002" -Level "INFO"
                $winCsOutput = & $winCsPath /apply --key "F33E0C8E002" 2>&1
                $winCsOutputStr = ($winCsOutput | Out-String).Trim()
                foreach ($line in ($winCsOutputStr -split "`n")) {
                    $trimmed = $line.Trim()
                    if ($trimmed) {
                        Write-Log -Message "  WinCS: $trimmed" -Level "INFO"
                    }
                }
                # WinCS sets the flags; the Secure-Boot-Update task runs every 12h to process them
                # Manually trigger the task to expedite
                $task = Get-ScheduledTask -TaskPath "\Microsoft\Windows\PI\" -TaskName "Secure-Boot-Update" -ErrorAction SilentlyContinue
                if ($task) {
                    Write-Log -Message "  Triggering Secure-Boot-Update task to expedite processing" -Level "INFO"
                    Start-ScheduledTask -TaskPath "\Microsoft\Windows\PI\" -TaskName "Secure-Boot-Update" -ErrorAction SilentlyContinue
                    Write-Log -Message "  Task triggered successfully" -Level "SUCCESS"
                }
                else {
                    Write-Log -Message "  Secure-Boot-Update task not found - WinCS will process on next TPMTasks cycle" -Level "INFO"
                }
                Write-Log -Message "--- End Fallback (WinCS) ---" -Level "INFO"
                Write-Host "FALLBACK_WINCS: WinCS applied key F33E0C8E002. CA2023 was: $ca2023Text. Reboot required."
                Write-Log -Message "Remediation Result: FALLBACK_WINCS (exit 0)" -Level "SUCCESS"
                Write-Log -Message "========== REMEDIATION COMPLETED ==========" -Level "INFO"
                Set-ExitCode 0
                return
            }
            catch {
                Write-Log -Message "  WinCS apply FAILED: $($_.Exception.Message)" -Level "ERROR"
                Write-Log -Message "  Falling through to legacy AvailableUpdates method" -Level "WARNING"
            }
        }
        else {
            Write-Log -Message "  WinCsFlags.exe: NOT AVAILABLE (requires Oct/Nov 2025+ cumulative update)" -Level "INFO"
        }
        
        # --- Legacy path: AvailableUpdates + scheduled task ---
        # Pre-flight: check payload folder
        $payload = Get-SecureBootPayloadStatus
        if (-not $payload.IsHealthy) {
            $payloadDetail = if ($payload.FolderExists) { "folder exists but empty" } else { "folder missing" }
            Write-Log -Message "  Payload Pre-flight: FAILED ($payloadDetail)" -Level "ERROR"
            Write-Log -Message "  SecureBootUpdates folder has no .bin payload files" -Level "ERROR"
            Write-Log -Message "  The Secure-Boot-Update task will fail with 0x80070002 without these files" -Level "ERROR"
            if (-not $winCsAvailable) {
                Write-Log -Message "  Neither WinCS nor payload files are available on this device" -Level "ERROR"
                Write-Log -Message "  ACTION: Install the latest cumulative update to get WinCsFlags.exe or payload files" -Level "ERROR"
            }
            Write-Log -Message "--- End Fallback (No Method Available) ---" -Level "INFO"
            Write-Host "FALLBACK_BLOCKED: No WinCS and no payload files. Install latest cumulative update. CA2023: $ca2023Text"
            Write-Log -Message "Remediation Result: FALLBACK_BLOCKED (exit 1)" -Level "ERROR"
            Write-Log -Message "========== REMEDIATION COMPLETED ==========" -Level "INFO"
            Set-ExitCode 1
            return
        }
        Write-Log -Message "  Payload Pre-flight: PASSED ($($payload.FileCount) files, .bin present)" -Level "SUCCESS"
        
        # Check scheduled task prerequisite
        $task = Get-ScheduledTask -TaskPath "\Microsoft\Windows\PI\" -TaskName "Secure-Boot-Update" -ErrorAction SilentlyContinue
        if (-not $task) {
            Write-Log -Message "  FALLBACK BLOCKED: Scheduled task '\Microsoft\Windows\PI\Secure-Boot-Update' not found" -Level "ERROR"
            Write-Log -Message "  The required Windows Update (July 2024+) may not be installed" -Level "ERROR"
            Write-Log -Message "--- End Fallback ---" -Level "INFO"
            Write-Host "FALLBACK_BLOCKED: Scheduled task not found. CA2023: $ca2023Text"
            Write-Log -Message "Remediation Result: FALLBACK_BLOCKED (exit 1)" -Level "ERROR"
            Write-Log -Message "========== REMEDIATION COMPLETED ==========" -Level "INFO"
            Set-ExitCode 1
            return
        }
        
        Write-Log -Message "  Using legacy AvailableUpdates + scheduled task method (KB5025885)" -Level "INFO"
        $fallbackSuccess = $true
        
        # Step 1: DB cert (Mitigation 1) - if cert not yet in DB
        if ($null -eq $ca2023Capable -or $ca2023Capable -lt 1) {
            Write-Log -Message "  Step 1: Setting AvailableUpdates = 0x40 (DB cert update)" -Level "INFO"
            try {
                Set-ItemProperty -Path $regPath -Name "AvailableUpdates" -Value 0x40 -Type DWord -Force -ErrorAction Stop
                $verifyAU = (Get-ItemProperty -Path $regPath -Name "AvailableUpdates" -ErrorAction Stop).AvailableUpdates
                if ($verifyAU -eq 0x40) {
                    Write-Log -Message "  Step 1: AvailableUpdates set to 0x40 - verified" -Level "SUCCESS"
                }
                else {
                    Write-Log -Message "  Step 1: Verification failed (got 0x$($verifyAU.ToString('X')))" -Level "ERROR"
                    $fallbackSuccess = $false
                }
                
                Write-Log -Message "  Step 1: Triggering scheduled task" -Level "INFO"
                Start-ScheduledTask -TaskPath "\Microsoft\Windows\PI\" -TaskName "Secure-Boot-Update" -ErrorAction Stop
                Write-Log -Message "  Step 1: Scheduled task triggered successfully" -Level "SUCCESS"
                
                # Poll task completion rather than using a fixed sleep
                $taskDeadline = (Get-Date).AddSeconds(60)
                do {
                    Start-Sleep -Seconds 2
                    $taskState = (Get-ScheduledTask -TaskPath "\Microsoft\Windows\PI\" -TaskName "Secure-Boot-Update" -ErrorAction SilentlyContinue).State
                } while ($taskState -eq 'Running' -and (Get-Date) -lt $taskDeadline)
                if ($taskState -eq 'Running') {
                    Write-Log -Message "  Step 1: Task still running after 60s timeout - proceeding to Step 2" -Level "WARNING"
                }
                else {
                    Write-Log -Message "  Step 1: Task completed (State: $taskState)" -Level "INFO"
                    # Post-task result validation
                    $step1Info = Get-ScheduledTaskInfo -TaskPath "\Microsoft\Windows\PI\" -TaskName "Secure-Boot-Update" -ErrorAction SilentlyContinue
                    if ($step1Info -and $step1Info.LastTaskResult -eq 0x80070002) {
                        Write-Log -Message "  Step 1: Task failed with 0x80070002 (ERROR_FILE_NOT_FOUND)" -Level "ERROR"
                        Write-Log -Message "  ROOT CAUSE: Payload files missing despite pre-flight check passing (possible race condition)" -Level "ERROR"
                        $fallbackSuccess = $false
                    }
                    elseif ($step1Info -and $step1Info.LastTaskResult -ne 0) {
                        Write-Log -Message "  Step 1: Task returned non-zero result: 0x$($step1Info.LastTaskResult.ToString('X'))" -Level "WARNING"
                    }
                }
            }
            catch {
                Write-Log -Message "  Step 1 FAILED: $($_.Exception.Message)" -Level "ERROR"
                $fallbackSuccess = $false
            }
        }
        else {
            Write-Log -Message "  Step 1: Skipped - CA2023 cert already in DB (CA2023Capable=$ca2023Capable)" -Level "INFO"
        }
        
        # Step 2: Boot manager (Mitigation 2)
        if ($fallbackSuccess) {
            Write-Log -Message "  Step 2: Setting AvailableUpdates = 0x100 (boot manager update)" -Level "INFO"
            try {
                Set-ItemProperty -Path $regPath -Name "AvailableUpdates" -Value 0x100 -Type DWord -Force -ErrorAction Stop
                $verifyAU2 = (Get-ItemProperty -Path $regPath -Name "AvailableUpdates" -ErrorAction Stop).AvailableUpdates
                if ($verifyAU2 -eq 0x100) {
                    Write-Log -Message "  Step 2: AvailableUpdates set to 0x100 - verified" -Level "SUCCESS"
                }
                else {
                    Write-Log -Message "  Step 2: Verification failed (got 0x$($verifyAU2.ToString('X')))" -Level "ERROR"
                    $fallbackSuccess = $false
                }
                
                Write-Log -Message "  Step 2: Triggering scheduled task" -Level "INFO"
                Start-ScheduledTask -TaskPath "\Microsoft\Windows\PI\" -TaskName "Secure-Boot-Update" -ErrorAction Stop
                Write-Log -Message "  Step 2: Scheduled task triggered successfully" -Level "SUCCESS"
                
                # Post-task result validation for Step 2
                $step2Deadline = (Get-Date).AddSeconds(60)
                do {
                    Start-Sleep -Seconds 2
                    $taskState2 = (Get-ScheduledTask -TaskPath "\Microsoft\Windows\PI\" -TaskName "Secure-Boot-Update" -ErrorAction SilentlyContinue).State
                } while ($taskState2 -eq 'Running' -and (Get-Date) -lt $step2Deadline)
                $step2Info = Get-ScheduledTaskInfo -TaskPath "\Microsoft\Windows\PI\" -TaskName "Secure-Boot-Update" -ErrorAction SilentlyContinue
                if ($step2Info -and $step2Info.LastTaskResult -eq 0x80070002) {
                    Write-Log -Message "  Step 2: Task failed with 0x80070002 (ERROR_FILE_NOT_FOUND)" -Level "ERROR"
                    $fallbackSuccess = $false
                }
                elseif ($step2Info -and $step2Info.LastTaskResult -ne 0) {
                    Write-Log -Message "  Step 2: Task returned non-zero result: 0x$($step2Info.LastTaskResult.ToString('X'))" -Level "WARNING"
                }
                else {
                    Write-Log -Message "  Step 2: Task completed successfully" -Level "SUCCESS"
                }
            }
            catch {
                Write-Log -Message "  Step 2 FAILED: $($_.Exception.Message)" -Level "ERROR"
                $fallbackSuccess = $false
            }
        }
        else {
            Write-Log -Message "  Step 2: Skipped due to Step 1 failure" -Level "WARNING"
        }
        
        Write-Log -Message "--- End Fallback (Legacy) ---" -Level "INFO"
        
        if ($fallbackSuccess) {
            Write-Host "FALLBACK_APPLIED: Direct method triggered. CA2023 was: $ca2023Text. Reboot may be required."
            Write-Log -Message "Remediation Result: FALLBACK_APPLIED (exit 0)" -Level "SUCCESS"
            Write-Log -Message "========== REMEDIATION COMPLETED ==========" -Level "INFO"
            Set-ExitCode 0
            return
        }
        else {
            Write-Host "FALLBACK_FAILED: Direct method encountered errors. CA2023: $ca2023Text"
            Write-Log -Message "Remediation Result: FALLBACK_FAILED (exit 1)" -Level "ERROR"
            Write-Log -Message "========== REMEDIATION COMPLETED ==========" -Level "INFO"
            Set-ExitCode 1
            return
        }
    }
    
    # ══════════════════════════════════════════════════════════════════════
    # FIRST-TIME DEPLOYMENT — device not previously configured
    # ══════════════════════════════════════════════════════════════════════

    Write-Log -Message "Registry Configuration:" -Level "INFO"
    Write-Log -Message "  Path: $regPath" -Level "INFO"
    Write-Log -Message "  Name: $regName" -Level "INFO"
    Write-Log -Message "  Value: 0x$($regValue.ToString('X')) ($regValue)" -Level "INFO"
    Write-Log -Message "  Type: $regType" -Level "INFO"
    
    # Create registry path if it doesn't exist
    if (-not (Test-Path $regPath)) {
        Write-Log -Message "Registry path does not exist, creating: $regPath" -Level "INFO"
        New-Item -Path $regPath -Force | Out-Null
        Write-Log -Message "Registry path created successfully" -Level "SUCCESS"
    }
    else {
        Write-Log -Message "Registry path already exists" -Level "INFO"
    }
    
    # Set the registry value
    Write-Log -Message "Setting AvailableUpdates registry value..." -Level "INFO"
    Set-ItemProperty -Path $regPath -Name $regName -Value $regValue -Type $regType -Force -ErrorAction Stop
    Write-Log -Message "Registry value set successfully" -Level "SUCCESS"
    
    # Verify the registry value was set correctly
    Write-Log -Message "Verifying registry configuration..." -Level "INFO"
    $currentValue = (Get-ItemProperty -Path $regPath -Name $regName -ErrorAction Stop).$regName
    
    if ($currentValue -eq $regValue) {
        Write-Log -Message "Registry value verified: 0x$($currentValue.ToString('X')) ($currentValue)" -Level "SUCCESS"
        Write-Log -Message "Device is now configured for Secure Boot certificate deployment" -Level "SUCCESS"
        
        # Trigger the scheduled task to begin processing immediately
        $task = Get-ScheduledTask -TaskPath "\Microsoft\Windows\PI\" -TaskName "Secure-Boot-Update" -ErrorAction SilentlyContinue
        if ($task) {
            Start-ScheduledTask -TaskPath "\Microsoft\Windows\PI\" -TaskName "Secure-Boot-Update" -ErrorAction SilentlyContinue
            Write-Log -Message "Secure-Boot-Update task triggered to begin processing" -Level "SUCCESS"
        }
        else {
            Write-Log -Message "Secure-Boot-Update task not found - updates will process on next 12h cycle" -Level "INFO"
        }
        
        # Write fallback timer timestamp
        try {
            if (-not (Test-Path $TimestampRegPath)) {
                New-Item -Path $TimestampRegPath -Force | Out-Null
            }
            $optInDate = (Get-Date).ToString("o")
            Set-ItemProperty -Path $TimestampRegPath -Name "ManagedOptInDate" -Value $optInDate -Type String -Force
            Write-Log -Message "Fallback timer started: ManagedOptInDate = $optInDate (threshold: $FallbackDays days)" -Level "INFO"
        }
        catch {
            Write-Log -Message "WARNING: Could not write ManagedOptInDate: $($_.Exception.Message)" -Level "WARNING"
        }
        Write-Host "SUCCESS: AvailableUpdates set to 0x$($currentValue.ToString('X')). Certificate deployment initiated."
        Write-Log -Message "Remediation Result: SUCCESS (exit 0)" -Level "SUCCESS"
        Write-Log -Message "========== REMEDIATION COMPLETED ==========" -Level "INFO"
    }
    else {
        Write-Log -Message "Registry value mismatch detected!" -Level "ERROR"
        Write-Log -Message "  Expected: 0x$($regValue.ToString('X')) ($regValue)" -Level "ERROR"
        Write-Log -Message "  Actual: 0x$($currentValue.ToString('X')) ($currentValue)" -Level "ERROR"
        Write-Host "FAILED: Registry mismatch - Expected 0x$($regValue.ToString('X')), Got 0x$($currentValue.ToString('X'))"
        Write-Log -Message "Remediation Result: FAILED (exit 1)" -Level "ERROR"
        Write-Log -Message "========== REMEDIATION COMPLETED ==========" -Level "INFO"
        Set-ExitCode 1
        return
    }
    
    Set-ExitCode 0
    return
}
catch {
    Write-Log -Message "Unexpected error during remediation: $($_.Exception.Message)" -Level "ERROR"
    Write-Log -Message "Stack Trace: $($_.ScriptStackTrace)" -Level "ERROR"
    Write-Host "ERROR: $($_.Exception.Message)"
    Write-Log -Message "Remediation Result: ERROR (exit 1)" -Level "ERROR"
    Write-Log -Message "========== REMEDIATION COMPLETED ==========" -Level "INFO"
    Set-ExitCode 1
    return
}
#endregion