#Requires -Version 7.0
<#
.SYNOPSIS
    RPCS3 Batch PKG Installer — PowerShell 7 edition
.DESCRIPTION
    Extracts ZIPs containing PS3 PKG/RAP files, installs them via RPCS3,
    tracks state by SHA-256 hash, organises archives, and auto-unchecks cache options.
    Successfully installed folders (DIR mode) are moved to _Installed just like ZIPs.

.VERSION HISTORY
    v1.0  — Initial release. ZIP + folder support, state tracking, RAP copy/verify.
    v1.1  — Fixed StrictMode Count error (@() array wrapping for Get-ChildItem).
    v1.2  — Fixed exit-0 = success regardless of new game folder (DLC/patch support).
            Folders now move to _Installed/_Failed same as ZIPs.
    v1.3  — Replaced SendKeys Install with UI Automation InvokePattern (no focus bugs).
            Added $progressConfirmed to prevent premature grace-period exit.
    v1.4  — Added corrupt PKG detection: Cancel-only dialog returns -2 immediately
            instead of hanging forever.
    v1.5  — Fixed ESC race condition: ESC after Loading games cancel is now only sent
            if UI Automation failed. Unconditional ESC was landing on the PKG
            Installation dialog that opens immediately after, dismissing it.
    v2.0  — Fixed post-install hang: added break after Loading games cancel inside
            the watcher loop. Window needs time to close; without the break the next
            poll immediately saw it still visible, reset emptyChecks to 0, and looped
            forever leaving RPCS3 open doing nothing after a successful install.
    v2.1  — Simplified post-install watcher: kill RPCS3 the instant the PKG
            Installation progress bar disappears. Loading games interaction removed.
    v2.2  — Fixed fast-install miss: if the PKG installs faster than the 2s poll
            interval, the progress bar is never seen and $progressConfirmed stays
            false. Loading games only appears after a successful install, so seeing
            it now also sets $progressConfirmed. RPCS3 is killed once both the
            progress bar and Loading games are gone, covering all timing scenarios.

.PARAMETER SourceRoot
    Folder containing the source ZIP archives / PKG folders.
.PARAMETER Rpcs3Exe
    Full path to rpcs3.exe.
.PARAMETER Rpcs3Root
    Root folder of the RPCS3 installation (contains dev_hdd0).
.PARAMETER TempRoot
    Scratch folder used for ZIP extraction (cleaned up after each archive).
.PARAMETER SevenZipExe
    Path to 7z.exe.
.PARAMETER FolderPollSeconds
    How long (seconds) to keep polling for a new game folder after each PKG install.

.NOTES
    PowerShell 7 or later required.

.LICENSE
    MIT License
    Copyright (c) Paul Mardis
#>
param(
    [string] $SourceRoot        = 'D:\Arcade\System roms\Sony Playstation 3\Sony - PlayStation 3 (PSN) (Content)',
    [string] $Rpcs3Exe          = 'C:\Arcade\LaunchBox\Emulators\RPCS3\rpcs3.exe',
    [string] $Rpcs3Root         = 'C:\Arcade\LaunchBox\Emulators\RPCS3',
    [string] $TempRoot          = 'D:\Arcade\RPCS3_Temp',
    [string] $SevenZipExe       = 'C:\Program Files\7-Zip\7z.exe',
    [int]    $FolderPollSeconds = 20
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ── UTF-8 console output ─────────────────────────────────────────────────────
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)
$OutputEncoding           = [System.Text.UTF8Encoding]::new($false)

# ════════════════════════════════════════════════════════════════════════════
#  DERIVED PATHS
# ════════════════════════════════════════════════════════════════════════════
$GameRoot      = Join-Path $Rpcs3Root 'dev_hdd0\game'
$RapTargetRoot = Join-Path $Rpcs3Root 'dev_hdd0\home\00000001\exdata'
$PS3Root       = 'D:\Arcade\System roms\Sony Playstation 3'
$StateRoot     = Join-Path $PS3Root   '_Automation'
$InstalledRoot = Join-Path $PS3Root   '_Installed'
$FailedRoot    = Join-Path $PS3Root   '_Failed'
$LogFile       = Join-Path $StateRoot 'install_log.csv'
$StateFile     = Join-Path $StateRoot 'installed_state.json'

# ════════════════════════════════════════════════════════════════════════════
#  BANNER + INTERACTIVE RUN-TYPE PROMPT
# ════════════════════════════════════════════════════════════════════════════
function Show-Banner {
    Write-Host ''
    Write-Host '╔══════════════════════════════════════════════════════╗' -ForegroundColor Cyan
    Write-Host '║       RPCS3 Batch PKG Installer  (PS7)  v2.2        ║' -ForegroundColor Cyan
    Write-Host '╚══════════════════════════════════════════════════════╝' -ForegroundColor Cyan
    Write-Host ''
}

function Get-RunType {
    Write-Host '  Select run type:' -ForegroundColor Yellow
    Write-Host ''
    Write-Host '  [1]  Normal     — Install new items, skip already-installed' -ForegroundColor White
    Write-Host '  [2]  Dry Run    — Preview only, no changes made'             -ForegroundColor White
    Write-Host '  [3]  Force      — Reinstall ALL items (ignore state file)'   -ForegroundColor White
    Write-Host '  [4]  Failed     — Retry items currently in _Failed folder'   -ForegroundColor White
    Write-Host '  [Q]  Quit'                                                   -ForegroundColor White
    Write-Host ''

    do { $key = (Read-Host '  Your choice').Trim().ToUpper() }
    while ($key -notin @('1','2','3','4','Q'))

    return $key
}

Show-Banner
$runChoice = Get-RunType

if ($runChoice -eq 'Q') {
    Write-Host "`n  Aborted." -ForegroundColor Red
    return
}

Write-Host ''
Write-Host '  [Optional Limits]' -ForegroundColor Yellow
$typeChoice = Read-Host '  Scan for [1] Both  [2] ZIPs only  [3] Folders only? (Press Enter for Both)'
$batchInput = Read-Host '  Process how many items? (Press Enter for ALL)'

$scanZips    = $typeChoice -ne '3'
$scanFolders = $typeChoice -ne '2'
$batchLimit  = if ([string]::IsNullOrWhiteSpace($batchInput)) { 0 } else { [int]$batchInput }

$DryRun         = $runChoice -eq '2'
$ForceReinstall = $runChoice -eq '3'
$RetryFailed    = $runChoice -eq '4'

$modeLabel = @{ '1'='Normal'; '2'='Dry Run'; '3'='Force Reinstall'; '4'='Retry Failed' }[$runChoice]

Write-Host ''
Write-Host '  Mode     : ' -NoNewline -ForegroundColor Gray; Write-Host $modeLabel   -ForegroundColor Cyan
Write-Host '  Source   : ' -NoNewline -ForegroundColor Gray; Write-Host $SourceRoot  -ForegroundColor Cyan
Write-Host '  RPCS3    : ' -NoNewline -ForegroundColor Gray; Write-Host $Rpcs3Exe    -ForegroundColor Cyan
Write-Host '  7-Zip    : ' -NoNewline -ForegroundColor Gray; Write-Host $SevenZipExe -ForegroundColor Cyan
Write-Host ''

# ════════════════════════════════════════════════════════════════════════════
#  HELPER FUNCTIONS
# ════════════════════════════════════════════════════════════════════════════

function Initialize-Folder ([string] $Path) {
    if (-not (Test-Path -LiteralPath $Path)) {
        New-Item -ItemType Directory -Path $Path -Force | Out-Null
    }
}

function Write-TextUtf8 ([string] $Path, [string] $Text) {
    [System.IO.File]::WriteAllText($Path, $Text, [System.Text.UTF8Encoding]::new($false))
}

function Add-LineUtf8 ([string] $Path, [string] $Line) {
    $sw = [System.IO.StreamWriter]::new($Path, $true, [System.Text.UTF8Encoding]::new($false))
    try   { $sw.WriteLine($Line) }
    finally { $sw.Dispose() }
}

function Write-LogRow {
    param(
        [string] $ArchivePath,
        [string] $PkgPath,
        [string] $RapPath,
        [string] $Status,
        [string] $Details,
        [string] $DetectedGameFolder,
        [int]    $ExitCode
    )
    $row = [pscustomobject] @{
        Timestamp          = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
        ArchivePath        = $ArchivePath
        PkgPath            = $PkgPath
        RapPath            = $RapPath
        Status             = $Status
        Details            = $Details
        DetectedGameFolder = $DetectedGameFolder
        ExitCode           = $ExitCode
    }
    if (-not (Test-Path -LiteralPath $LogFile)) {
        Add-LineUtf8 -Path $LogFile -Line ($row | ConvertTo-Csv -NoTypeInformation)[0]
    }
    Add-LineUtf8 -Path $LogFile -Line ($row | ConvertTo-Csv -NoTypeInformation)[1]
}

function Import-State {
    if (Test-Path -LiteralPath $StateFile) {
        try {
            $raw = [System.IO.File]::ReadAllText($StateFile, [System.Text.Encoding]::UTF8)
            if ([string]::IsNullOrWhiteSpace($raw)) { return @{} }
            return $raw | ConvertFrom-Json -AsHashtable
        }
        catch { return @{} }
    }
    return @{}
}

function Save-State ([hashtable] $State) {
    Write-TextUtf8 -Path $StateFile -Text ($State | ConvertTo-Json -Depth 10)
}

function Get-FileHashSafe ([string] $Path) {
    try   { return (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash }
    catch { return $null }
}

function Get-GameSnapshot ([string] $Path) {
    $snapshot = @{}
    if (Test-Path -LiteralPath $Path) {
        Get-ChildItem -LiteralPath $Path -Directory | ForEach-Object {
            $sfo = Join-Path $_.FullName 'PARAM.SFO'
            $snapshot[$_.FullName] = (Test-Path -LiteralPath $sfo) `
                ? (Get-Item -LiteralPath $sfo).LastWriteTimeUtc `
                : $_.LastWriteTimeUtc
        }
    }
    return $snapshot
}

function Wait-ForNewGameFolder {
    param(
        [hashtable] $Before,
        [datetime]  $InstallStartUtc,
        [int]       $TimeoutSeconds = 20
    )
    $deadline = (Get-Date).AddSeconds($TimeoutSeconds)
    while ((Get-Date) -lt $deadline) {
        Start-Sleep -Milliseconds 1000
        $after = Get-GameSnapshot -Path $GameRoot
        foreach ($p in $after.Keys) {
            if (-not $Before.ContainsKey($p)) { return $p }
            if ($after[$p] -gt $Before[$p] -and $after[$p] -ge $InstallStartUtc.AddMinutes(-1)) { return $p }
        }
    }
    return $null
}

function Move-ToFolder ([string] $Path, [string] $Destination) {
    if (-not (Test-Path -LiteralPath $Path)) { return $null }
    Initialize-Folder -Path $Destination
    $leaf = [System.IO.Path]::GetFileName($Path)
    $dest = Join-Path $Destination $leaf
    if (Test-Path -LiteralPath $dest) {
        $base = [System.IO.Path]::GetFileNameWithoutExtension($Path)
        $ext  = [System.IO.Path]::GetExtension($Path)
        $dest = Join-Path $Destination ('{0}_{1}{2}' -f $base, (Get-Date -Format 'yyyyMMdd_HHmmss'), $ext)
    }
    Move-Item -LiteralPath $Path -Destination $dest -Force
    return $dest
}

function Remove-FolderSafe ([string] $Path) {
    if (Test-Path -LiteralPath $Path) {
        Remove-Item -LiteralPath $Path -Recurse -Force -ErrorAction SilentlyContinue
    }
}

function Get-MatchingRap ([System.IO.FileInfo] $Pkg, [System.IO.FileInfo[]] $Raps) {
    $pkgBase = [System.IO.Path]::GetFileNameWithoutExtension($Pkg.Name)
    $exact   = $Raps | Where-Object {
        [System.IO.Path]::GetFileNameWithoutExtension($_.Name) -eq $pkgBase
    } | Select-Object -First 1
    if ($exact) { return $exact }

    return $Raps | Where-Object {
        $rapBase = [System.IO.Path]::GetFileNameWithoutExtension($_.Name)
        $pkgBase.StartsWith($rapBase, [System.StringComparison]::OrdinalIgnoreCase)
    } | Sort-Object { $_.Name.Length } -Descending | Select-Object -First 1
}

function Expand-WithSevenZip ([string] $ArchivePath, [string] $DestinationPath) {
    & $SevenZipExe x $ArchivePath "-o$DestinationPath" -bso0 -bsp1 -y
    if ($LASTEXITCODE -ne 0) { throw "7-Zip exited with code $LASTEXITCODE" }
}

# ── Load UI Automation + Win32 APIs ──────────────────────────────────────────
Add-Type -AssemblyName UIAutomationClient   -ErrorAction SilentlyContinue
Add-Type -AssemblyName UIAutomationTypes    -ErrorAction SilentlyContinue
Add-Type -AssemblyName System.Windows.Forms -ErrorAction SilentlyContinue

if (-not ('RPCS3Win32' -as [type])) {
    Add-Type -TypeDefinition @'
using System;
using System.Text;
using System.Collections.Generic;
using System.Runtime.InteropServices;
public static class RPCS3Win32 {
    [DllImport("user32.dll", CharSet=CharSet.Unicode)]
    public static extern IntPtr FindWindow(IntPtr lpClassName, string lpWindowName);
    [DllImport("user32.dll")]
    public static extern bool SetForegroundWindow(IntPtr hWnd);
    [DllImport("user32.dll")]
    public static extern bool IsWindowVisible(IntPtr hWnd);

    public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);
    [DllImport("user32.dll")]
    public static extern bool EnumWindows(EnumWindowsProc enumProc, IntPtr lParam);
    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);
    [DllImport("user32.dll", CharSet=CharSet.Unicode)]
    public static extern int GetWindowText(IntPtr hWnd, StringBuilder strText, int maxCount);

    public static string[] GetProcessWindowTitles(int processId) {
        List<string> titles = new List<string>();
        EnumWindows((hWnd, lParam) => {
            if (IsWindowVisible(hWnd)) {
                uint pid;
                GetWindowThreadProcessId(hWnd, out pid);
                if (pid == (uint)processId) {
                    StringBuilder sb = new StringBuilder(256);
                    GetWindowText(hWnd, sb, 256);
                    titles.Add(sb.ToString());
                }
            }
            return true;
        }, IntPtr.Zero);
        return titles.ToArray();
    }
}
'@ -ErrorAction SilentlyContinue
}

# ════════════════════════════════════════════════════════════════════════════
#  RPCS3 INSTALL
#
#  Return codes:
#    0   = success
#   -1   = timed out waiting for dialog
#   -2   = corrupt / uninstallable PKG (RPCS3 rejected it)
#
#  Post-install watcher logic (v2.2):
#  RPCS3 always shows "Loading games" after a successful install to rebuild
#  its game list cache. This is the most reliable signal we have.
#
#  Three signals set $progressConfirmed:
#    1. PKG Installation progress bar seen (normal path)
#    2. Loading games appears (fast-install path — bar gone before first poll)
#    3. Either of the above (belt-and-suspenders)
#
#  RPCS3 is killed the moment $progressConfirmed is true AND both
#  $progressVisible and $loadingVisible are false — meaning everything
#  is done and RPCS3 is just sitting idle with the game list open.
# ════════════════════════════════════════════════════════════════════════════
function Invoke-RPCS3Install {
    param(
        [string] $PkgPath,
        [int]    $DialogTimeoutSeconds = 600
    )

    $outLog  = Join-Path $TempRoot 'rpcs3_stdout.log'
    $errLog  = Join-Path $TempRoot 'rpcs3_stderr.log'

    $proc = Start-Process -FilePath $Rpcs3Exe `
                          -ArgumentList '--installpkg', "`"$PkgPath`"" `
                          -WorkingDirectory $Rpcs3Root `
                          -RedirectStandardOutput $outLog `
                          -RedirectStandardError  $errLog `
                          -PassThru

    $deadline = (Get-Date).AddSeconds($DialogTimeoutSeconds)
    $clicked  = $false
    $wshell   = New-Object -ComObject wscript.shell

    Write-Host '    Waiting for Install dialog...' -NoNewline -ForegroundColor DarkGray

    while ((Get-Date) -lt $deadline) {
        Start-Sleep -Milliseconds 600
        if ($proc.HasExited) { break }

        # ── Pre-install: auto-cancel "Loading games" ─────────────────────────
        # ESC only sent if UI Automation failed — unconditional ESC would land
        # on the PKG Installation dialog that opens immediately after.
        $hwndLoading = [RPCS3Win32]::FindWindow([IntPtr]::Zero, 'Loading games')
        if ($hwndLoading -ne [IntPtr]::Zero -and [RPCS3Win32]::IsWindowVisible($hwndLoading)) {
            Write-Host "`r    [+] Auto-canceling 'Loading games' to save time...             " `
                -NoNewline -ForegroundColor DarkYellow
            $canceledViaAutomation = $false
            try {
                $loadElem   = [System.Windows.Automation.AutomationElement]::FromHandle($hwndLoading)
                $cancelCond = [System.Windows.Automation.PropertyCondition]::new(
                    [System.Windows.Automation.AutomationElement]::NameProperty, 'Cancel')
                $cancelBtn  = $loadElem.FindFirst(
                    [System.Windows.Automation.TreeScope]::Descendants, $cancelCond)
                if ($cancelBtn) {
                    $inv = $cancelBtn.GetCurrentPattern(
                        [System.Windows.Automation.InvokePattern]::Pattern) `
                        -as [System.Windows.Automation.InvokePattern]
                    if ($inv) {
                        $inv.Invoke()
                        $canceledViaAutomation = $true
                        Start-Sleep -Milliseconds 800
                    }
                }
            } catch {}

            if (-not $canceledViaAutomation) {
                $wshell.AppActivate($proc.Id) | Out-Null
                $wshell.SendKeys('{ESC}')
                Start-Sleep -Milliseconds 800
            }

            $deadline = (Get-Date).AddSeconds($DialogTimeoutSeconds)
            continue
        }

        # ── PKG Installation dialog ──────────────────────────────────────────
        $hwndInstall = [RPCS3Win32]::FindWindow([IntPtr]::Zero, 'PKG Installation')
        if ($hwndInstall -ne [IntPtr]::Zero -and [RPCS3Win32]::IsWindowVisible($hwndInstall)) {

            [RPCS3Win32]::SetForegroundWindow($hwndInstall) | Out-Null
            Start-Sleep -Milliseconds 500

            try {
                $winElem = [System.Windows.Automation.AutomationElement]::FromHandle($hwndInstall)

                # Check for Install button first
                $btnCond    = [System.Windows.Automation.PropertyCondition]::new(
                    [System.Windows.Automation.AutomationElement]::NameProperty, 'Install')
                $installBtn = $winElem.FindFirst(
                    [System.Windows.Automation.TreeScope]::Descendants, $btnCond)

                # No Install button = corrupt/error dialog (Cancel only) — bail immediately
                if (-not $installBtn) {
                    Write-Host "`r    [!] Corrupt PKG error dialog detected — no Install button found.  " `
                        -ForegroundColor Red
                    $cancelCond = [System.Windows.Automation.PropertyCondition]::new(
                        [System.Windows.Automation.AutomationElement]::NameProperty, 'Cancel')
                    $cancelBtn  = $winElem.FindFirst(
                        [System.Windows.Automation.TreeScope]::Descendants, $cancelCond)
                    if ($cancelBtn) {
                        $inv = $cancelBtn.GetCurrentPattern(
                            [System.Windows.Automation.InvokePattern]::Pattern) `
                            -as [System.Windows.Automation.InvokePattern]
                        if ($inv) {
                            $inv.Invoke()
                            Write-Host "`r    [+] Clicked Cancel on corrupt package dialog.                  " `
                                -ForegroundColor DarkYellow
                            Start-Sleep -Seconds 2
                        }
                    } else {
                        $wshell.AppActivate($proc.Id) | Out-Null
                        $wshell.SendKeys('{ESC}')
                        Start-Sleep -Seconds 1
                    }
                    $proc | Stop-Process -Force -ErrorAction SilentlyContinue
                    return -2
                }

                # Uncheck Precompile caches if ticked
                $cbCond   = [System.Windows.Automation.PropertyCondition]::new(
                    [System.Windows.Automation.AutomationElement]::NameProperty, 'Precompile caches')
                $checkbox = $winElem.FindFirst(
                    [System.Windows.Automation.TreeScope]::Descendants, $cbCond)
                if ($checkbox) {
                    $toggle = $checkbox.GetCurrentPattern(
                        [System.Windows.Automation.TogglePattern]::Pattern) `
                        -as [System.Windows.Automation.TogglePattern]
                    if ($toggle -and
                        $toggle.Current.ToggleState -eq [System.Windows.Automation.ToggleState]::On) {
                        $toggle.Toggle()
                        Write-Host "`r    [+] Auto-unchecked 'Precompile caches'.                        " `
                            -ForegroundColor DarkCyan
                        Start-Sleep -Milliseconds 300
                    }
                }

                # Invoke Install button via UI Automation
                $invoke = $installBtn.GetCurrentPattern(
                    [System.Windows.Automation.InvokePattern]::Pattern) `
                    -as [System.Windows.Automation.InvokePattern]
                if ($invoke) {
                    $invoke.Invoke()
                    $clicked = $true
                    Write-Host "`r    [+] Clicked Install via UI Automation.                         " `
                        -ForegroundColor DarkCyan
                }
            }
            catch {
                # Automation failed entirely — last-resort SendKeys
                Write-Host "`r    [!] Automation error — falling back to Alt+I SendKeys.         " `
                    -ForegroundColor DarkYellow
                try {
                    $wshell.AppActivate($proc.Id) | Out-Null
                    Start-Sleep -Milliseconds 300
                    $wshell.SendKeys('%i')
                    $clicked = $true
                } catch {}
            }

            if (-not $clicked) {
                $proc | Stop-Process -Force -ErrorAction SilentlyContinue
                return -1
            }

            # ── POST-INSTALL WATCHER (v2.2) ──────────────────────────────────
            # Three signals that confirm install success:
            #   1. PKG Installation progress bar seen        (normal timing)
            #   2. Loading games appears after install       (fast install —
            #      bar disappeared before first 2s poll)
            #   3. Both seen                                 (belt-and-suspenders)
            #
            # Kill RPCS3 the moment $progressConfirmed is true AND both
            # $progressVisible and $loadingVisible are false — all activity done.
            Write-Host "`r    Installing — waiting for PKG to unpack...                      " `
                -NoNewline -ForegroundColor DarkGreen
            Start-Sleep -Seconds 3

            $installTimeout    = (Get-Date).AddMinutes(45)
            $progressConfirmed = $false

            while ((Get-Date) -lt $installTimeout) {
                if ($proc.HasExited) { break }

                $titles          = [RPCS3Win32]::GetProcessWindowTitles($proc.Id)
                $progressVisible = $false
                $loadingVisible  = $false

                foreach ($t in $titles) {
                    if ($t -match '^RPCS3 \d+\.\d+') { continue }
                    if ([string]::IsNullOrWhiteSpace($t)) { continue }

                    # PKG Installation — progress bar or corrupt error dialog
                    if ($t -match 'PKG Installation') {
                        try {
                            $hwndPkg = [RPCS3Win32]::FindWindow([IntPtr]::Zero, 'PKG Installation')
                            $pkgElem = [System.Windows.Automation.AutomationElement]::FromHandle($hwndPkg)

                            $midCancelCond  = [System.Windows.Automation.PropertyCondition]::new(
                                [System.Windows.Automation.AutomationElement]::NameProperty, 'Cancel')
                            $midInstallCond = [System.Windows.Automation.PropertyCondition]::new(
                                [System.Windows.Automation.AutomationElement]::NameProperty, 'Install')

                            $hasCancel  = $null -ne $pkgElem.FindFirst(
                                [System.Windows.Automation.TreeScope]::Descendants, $midCancelCond)
                            $hasInstall = $null -ne $pkgElem.FindFirst(
                                [System.Windows.Automation.TreeScope]::Descendants, $midInstallCond)

                            if ($hasCancel -and -not $hasInstall) {
                                # Corrupt PKG error dialog — cancel and bail
                                Write-Host "`r    [!] Corrupt PKG error detected mid-install — canceling.        " `
                                    -ForegroundColor Red
                                $errCancel = $pkgElem.FindFirst(
                                    [System.Windows.Automation.TreeScope]::Descendants, $midCancelCond)
                                if ($errCancel) {
                                    $inv = $errCancel.GetCurrentPattern(
                                        [System.Windows.Automation.InvokePattern]::Pattern) `
                                        -as [System.Windows.Automation.InvokePattern]
                                    if ($inv) { $inv.Invoke() }
                                }
                                Start-Sleep -Seconds 1
                                $proc | Stop-Process -Force -ErrorAction SilentlyContinue
                                return -2
                            } else {
                                # Genuine progress bar
                                $progressConfirmed = $true
                                $progressVisible   = $true
                            }
                        } catch {
                            $progressConfirmed = $true
                            $progressVisible   = $true
                        }
                    }

                    # Loading games only appears after a successful install.
                    # If we see it, install is confirmed even if we missed the
                    # progress bar during a fast install (v2.2 fix).
                    if ($t -match 'Loading games') {
                        $progressConfirmed = $true
                        $loadingVisible    = $true
                        Write-Host "`r    [~] Post-install cache rebuild running — waiting to finish...   " `
                            -NoNewline -ForegroundColor DarkGray
                    }
                }

                # Install confirmed AND all post-install windows gone — we are done.
                # Kill RPCS3 immediately; no need to wait for anything else.
                if ($progressConfirmed -and -not $progressVisible -and -not $loadingVisible) {
                    Write-Host "`r    Install complete! Closing RPCS3...                             " `
                        -ForegroundColor DarkGreen
                    Start-Sleep -Seconds 2   # let filesystem flush writes
                    $proc | Stop-Process -Force -ErrorAction SilentlyContinue
                    return 0
                }

                Start-Sleep -Seconds 2
            }

            # Fell through 45-min timeout — close cleanly anyway
            Write-Host "`r    Install timeout reached. Closing RPCS3...                      " `
                -ForegroundColor DarkYellow
            $proc | Stop-Process -Force -ErrorAction SilentlyContinue
            return 0
        }
    }

    if (-not $clicked -and -not $proc.HasExited) {
        Write-Host "`r    Timed out waiting for dialog — killing RPCS3.                  " -ForegroundColor Red
        $proc | Stop-Process -Force -ErrorAction SilentlyContinue
        return -1
    }

    if ($proc.HasExited) { return $proc.ExitCode }
    return 0
}

# ════════════════════════════════════════════════════════════════════════════
#  OUTPUT HELPERS
# ════════════════════════════════════════════════════════════════════════════
function Write-ProgressLine ([int] $Current, [int] $Total, [string] $ZipName) {
    $pct  = [math]::Floor(($Current / $Total) * 100)
    $done = [math]::Floor($pct / 5)
    $bar  = ('█' * $done).PadRight(20, '░')
    Write-Host ("`r  [{0}] {1,3}%  ({2}/{3})  {4}" -f $bar, $pct, $Current, $Total, $ZipName) `
        -NoNewline -ForegroundColor Cyan
}

function Write-StatusLine ([string] $Icon, [string] $Color, [string] $Label, [string] $Detail) {
    Write-Host ''
    Write-Host ('  {0} {1,-12} {2}' -f $Icon, $Label, $Detail) -ForegroundColor $Color
}

# ════════════════════════════════════════════════════════════════════════════
#  PRE-FLIGHT CHECKS
# ════════════════════════════════════════════════════════════════════════════
foreach ($check in @(
    @{ Path = $Rpcs3Exe;    Label = 'RPCS3 executable' }
    @{ Path = $SevenZipExe; Label = '7-Zip executable'  }
    @{ Path = $GameRoot;    Label = 'RPCS3 game folder'  }
)) {
    if (-not (Test-Path -LiteralPath $check.Path)) {
        Write-Host ("  ERROR: {0} not found:`n         {1}" -f $check.Label, $check.Path) -ForegroundColor Red
        return
    }
}

foreach ($dir in @($StateRoot, $InstalledRoot, $FailedRoot, $TempRoot, $RapTargetRoot)) {
    Initialize-Folder -Path $dir
}

# ════════════════════════════════════════════════════════════════════════════
#  LOAD STATE & COLLECT WORK ITEMS
# ════════════════════════════════════════════════════════════════════════════
$state    = Import-State
$scanRoot = $RetryFailed ? $FailedRoot : $SourceRoot

$skipFolderNames = @('_DeDuplication','_Installed','_Failed','_Automation','Complete')
$workItems = [System.Collections.Generic.List[pscustomobject]]::new()

if ($scanZips) {
    Write-Host '  🔍 Scanning for ZIP archives...' -ForegroundColor DarkYellow
    $zips = Get-ChildItem -LiteralPath $scanRoot -Recurse -File -Filter *.zip |
        Where-Object {
            $_.FullName -notmatch '\\_installed\\'  -and
            $_.FullName -notmatch '\\_automation\\' -and
            $_.FullName -notmatch '\\_failed\\'
        }

    foreach ($z in $zips) {
        if ($batchLimit -gt 0 -and $workItems.Count -ge $batchLimit) {
            Write-Host '  ⚠ Batch limit reached. Stopping ZIP scan early.' -ForegroundColor DarkYellow
            break
        }
        $workItems.Add([pscustomobject] @{
            Type        = 'zip'
            SourcePath  = $z.FullName
            DisplayName = $z.BaseName
            SizeBytes   = $z.Length
        })
    }
}

if ($scanFolders -and ($batchLimit -eq 0 -or $workItems.Count -lt $batchLimit)) {
    Write-Host '  🔍 Scanning for folders with PKGs...' -ForegroundColor DarkYellow
    $dirs = Get-ChildItem -LiteralPath $scanRoot -Directory |
        Where-Object { $_.Name -notin $skipFolderNames -and $_.Name -notmatch '^_' }

    foreach ($d in $dirs) {
        if ($batchLimit -gt 0 -and $workItems.Count -ge $batchLimit) {
            Write-Host '  ⚠ Batch limit reached. Stopping folder scan early.' -ForegroundColor DarkYellow
            break
        }
        $hasPkg = [bool](Get-ChildItem -LiteralPath $d.FullName -Recurse -File -Filter *.pkg `
                            -ErrorAction SilentlyContinue | Select-Object -First 1)
        if ($hasPkg) {
            $workItems.Add([pscustomobject] @{
                Type        = 'folder'
                SourcePath  = $d.FullName
                DisplayName = $d.Name
                SizeBytes   = 0
            })
        }
    }
}

$workItems = [pscustomobject[]] ($workItems | Sort-Object DisplayName)

if ($workItems.Count -eq 0) {
    Write-Host ('  No installable items found in {0}.' -f ($RetryFailed ? '_Failed' : 'source folder')) `
        -ForegroundColor Yellow
    return
}

$zipCount    = @($workItems | Where-Object Type -eq 'zip').Count
$folderCount = @($workItems | Where-Object Type -eq 'folder').Count
$total       = $workItems.Count

Write-Host ("  Found {0} item(s) to process  ({1} ZIP, {2} folder)." -f $total, $zipCount, $folderCount) `
    -ForegroundColor White
Write-Host ''

$existingRpcs3 = Get-Process -Name 'rpcs3' -ErrorAction SilentlyContinue
if ($existingRpcs3) {
    Write-Host '  ⚠  RPCS3 is already running.' -ForegroundColor Yellow
    Write-Host '     The installer cannot run alongside an open RPCS3 instance.' -ForegroundColor Yellow
    $kill = (Read-Host '     Kill it now and continue? [Y/N]').Trim().ToUpper()
    if ($kill -eq 'Y') {
        $existingRpcs3 | Stop-Process -Force
        Start-Sleep -Seconds 2
        Write-Host '     RPCS3 closed.' -ForegroundColor Green
    } else {
        Write-Host '     Aborted. Close RPCS3 manually and re-run.' -ForegroundColor Red
        return
    }
}
Write-Host ''

$counters    = @{ Success = 0; Skipped = 0; Failed = 0; Error = 0 }
$failedItems = [System.Collections.Generic.List[pscustomobject]]::new()

# ════════════════════════════════════════════════════════════════════════════
#  MAIN LOOP
# ════════════════════════════════════════════════════════════════════════════
$idx = 0
foreach ($item in $workItems) {
    $idx++
    $isZip      = $item.Type -eq 'zip'
    $sourcePath = $item.SourcePath
    $dispName   = $item.DisplayName
    $typeLabel  = $isZip ? 'ZIP' : 'DIR'

    Write-ProgressLine -Current $idx -Total $total -ZipName $dispName

    $itemHash = $null
    if ($isZip) {
        $itemHash = Get-FileHashSafe -Path $sourcePath
    } else {
        $firstPkg = Get-ChildItem -LiteralPath $sourcePath -Recurse -File -Filter *.pkg |
                        Select-Object -First 1
        if ($firstPkg) {
            $keyString = '{0}|{1}|{2}' -f $dispName, $firstPkg.Name, $firstPkg.Length
            $bytes     = [System.Text.Encoding]::UTF8.GetBytes($keyString)
            $itemHash  = [System.Convert]::ToHexString(
                            [System.Security.Cryptography.SHA256]::HashData($bytes))
        }
    }

    if (-not $ForceReinstall -and -not $RetryFailed -and $itemHash -and $state.ContainsKey($itemHash)) {
        $cachedFolder = $state[$itemHash]?['DetectedGameFolder'] ?? ''
        Write-StatusLine -Icon '⏭' -Color DarkGray -Label 'SKIPPED' -Detail $dispName
        $counters.Skipped++
        Write-LogRow -ArchivePath $sourcePath -PkgPath '' -RapPath '' -Status 'Skipped' `
            -Details 'Already in state file' -DetectedGameFolder $cachedFolder -ExitCode 0
        continue
    }

    $workFolder   = $sourcePath
    $needsCleanup = $false

    try {
        Write-Host ''
        Write-Host ("  ── [{0}/{1}] {2}  [{3}]" -f $idx, $total, $dispName, $typeLabel) -ForegroundColor White

        if ($isZip) {
            $workFolder   = Join-Path $TempRoot $dispName
            $needsCleanup = $true
            $sizeMB       = [math]::Round($item.SizeBytes / 1MB, 1)
            Write-Host ("  ⏳ STAGE 1/3  Extracting  ({0} MB)..." -f $sizeMB) -ForegroundColor DarkYellow
            Remove-FolderSafe   -Path $workFolder
            Initialize-Folder   -Path $workFolder
            Expand-WithSevenZip -ArchivePath $sourcePath -DestinationPath $workFolder
            Write-Host '  ✔  Extraction complete' -ForegroundColor DarkGreen
        } else {
            Write-Host '  ✔  STAGE 1/3  Using folder directly (no extraction needed)' -ForegroundColor DarkGreen
        }

        $pkgFiles = [System.IO.FileInfo[]] @(Get-ChildItem -LiteralPath $workFolder -Recurse -File -Filter *.pkg)
        $rapFiles = [System.IO.FileInfo[]] @(Get-ChildItem -LiteralPath $workFolder -Recurse -File -Filter *.rap)

        Write-Host ("    Found: {0} PKG, {1} RAP" -f $pkgFiles.Count, $rapFiles.Count) -ForegroundColor DarkGray

        if ($pkgFiles.Count -eq 0) {
            Write-StatusLine -Icon '✖' -Color Red -Label 'NO PKG' -Detail $dispName
            $counters.Failed++
            $failedItems.Add([pscustomobject] @{ Name = $dispName; Reason = 'No PKG file found'; Exit = 'N/A' })
            Write-LogRow -ArchivePath $sourcePath -PkgPath '' -RapPath '' -Status 'Failed' `
                -Details 'No PKG found' -DetectedGameFolder '' -ExitCode -1
            if (Test-Path -LiteralPath $sourcePath) {
                Move-ToFolder -Path $sourcePath -Destination $FailedRoot | Out-Null
            }
            if ($needsCleanup) { Remove-FolderSafe -Path $workFolder }
            continue
        }

        if ($DryRun) {
            foreach ($pkg in $pkgFiles) {
                $rap     = Get-MatchingRap -Pkg $pkg -Raps $rapFiles
                $rapNote = $rap ? "RAP: $($rap.Name)" : 'No RAP'
                Write-StatusLine -Icon '🔍' -Color DarkCyan -Label 'DRY RUN' `
                    -Detail "$($pkg.Name)  [$rapNote]"
                Write-LogRow -ArchivePath $sourcePath -PkgPath $pkg.FullName `
                    -RapPath ($rap?.FullName ?? '') `
                    -Status 'DryRun' -Details 'No changes made' -DetectedGameFolder '' -ExitCode 0
            }
            if ($needsCleanup) { Remove-FolderSafe -Path $workFolder }
            continue
        }

        $itemSuccess     = $true
        $detectedFolders = [System.Collections.Generic.List[string]]::new()

        foreach ($pkg in $pkgFiles) {
            $rap = Get-MatchingRap -Pkg $pkg -Raps $rapFiles

            $rapDestPath = ''
            if ($rap) {
                $rapDestPath = Join-Path $RapTargetRoot $rap.Name
                Copy-Item -LiteralPath $rap.FullName -Destination $rapDestPath -Force

                if (-not (Test-Path -LiteralPath $rapDestPath)) {
                    throw "RAP copy failed — file missing at destination: $rapDestPath"
                }
                $srcHash  = Get-FileHashSafe -Path $rap.FullName
                $destHash = Get-FileHashSafe -Path $rapDestPath
                if ($srcHash -ne $destHash) { throw "RAP hash mismatch after copy: $($rap.Name)" }

                Write-Host ("    RAP      : {0}  ✔ verified" -f $rap.Name) -ForegroundColor DarkGreen
            } else {
                Write-Host '    RAP      : none found (game may not need one)' -ForegroundColor DarkGray
            }

            $before          = Get-GameSnapshot -Path $GameRoot
            $installStartUtc = (Get-Date).ToUniversalTime()

            Write-Host ("  ▶ STAGE 2/3  Installing: {0}" -f $pkg.Name) -ForegroundColor White

            $exitCode = Invoke-RPCS3Install -PkgPath $pkg.FullName

            Write-Host '  ⏳ STAGE 3/3  Waiting for game folder...' -NoNewline -ForegroundColor DarkGray
            $detectedFolder = Wait-ForNewGameFolder -Before $before `
                                                    -InstallStartUtc $installStartUtc `
                                                    -TimeoutSeconds $FolderPollSeconds
            Write-Host "`r                                                             `r" -NoNewline

            if ($exitCode -eq 0) {
                if ($detectedFolder) { $detectedFolders.Add($detectedFolder) }
                $folderNote = $detectedFolder `
                    ? [System.IO.Path]::GetFileName($detectedFolder) `
                    : 'DLC/Patch (existing folder updated)'
                Write-StatusLine -Icon '✔' -Color Green -Label 'OK' `
                    -Detail ('{0}  →  {1}' -f $pkg.Name, $folderNote)
                Write-LogRow -ArchivePath $sourcePath -PkgPath $pkg.FullName `
                    -RapPath $rapDestPath -Status 'Success' -Details 'RPCS3 exit 0' `
                    -DetectedGameFolder ($detectedFolder ?? 'DLC/Patch') -ExitCode $exitCode
            } else {
                $itemSuccess = $false
                $reason = $exitCode -eq -2 `
                    ? 'Corrupt or uninstallable PKG (RPCS3 rejected it)' `
                    : "RPCS3 exit code $exitCode"
                Write-StatusLine -Icon '✖' -Color Red -Label 'FAILED' `
                    -Detail ('{0}  —  {1}' -f $pkg.Name, $reason)
                $failedItems.Add([pscustomobject] @{ Name = $dispName; Reason = $reason; Exit = $exitCode })
                Write-LogRow -ArchivePath $sourcePath -PkgPath $pkg.FullName `
                    -RapPath $rapDestPath -Status 'Failed' -Details $reason `
                    -DetectedGameFolder '' -ExitCode $exitCode
            }
        }

        if ($itemSuccess) {
            $counters.Success++
            $movedTo = Move-ToFolder -Path $sourcePath -Destination $InstalledRoot
            if ($movedTo) {
                Write-Host ("  📦 {0} moved  →  _Installed" -f ($isZip ? 'ZIP' : 'Folder')) `
                    -ForegroundColor DarkGreen
            }
            if ($itemHash) {
                $state[$itemHash] = @{
                    Timestamp          = (Get-Date -Format 'o')
                    Type               = $typeLabel
                    SourcePath         = $sourcePath
                    ArchiveMovedTo     = $movedTo
                    DetectedGameFolder = ($detectedFolders -join '; ')
                }
                Save-State -State $state
            }
        } else {
            $counters.Failed++
            if (Test-Path -LiteralPath $sourcePath) {
                Move-ToFolder -Path $sourcePath -Destination $FailedRoot | Out-Null
                Write-Host ("  📦 {0} moved  →  _Failed" -f ($isZip ? 'ZIP' : 'Folder')) `
                    -ForegroundColor DarkYellow
            }
        }

        if ($needsCleanup) { Remove-FolderSafe -Path $workFolder }
    }
    catch {
        $errMsg = $_.Exception.Message
        Write-StatusLine -Icon '⚠' -Color Red -Label 'ERROR' -Detail ('{0}  —  {1}' -f $dispName, $errMsg)
        $counters.Error++
        $failedItems.Add([pscustomobject] @{ Name = $dispName; Reason = $errMsg; Exit = 'exception' })
        Write-LogRow -ArchivePath $sourcePath -PkgPath '' -RapPath '' -Status 'Error' `
            -Details $errMsg -DetectedGameFolder '' -ExitCode -1

        if (Test-Path -LiteralPath $sourcePath) {
            Move-ToFolder -Path $sourcePath -Destination $FailedRoot | Out-Null
            Write-Host ("  📦 {0} moved  →  _Failed" -f ($isZip ? 'ZIP' : 'Folder')) `
                -ForegroundColor DarkYellow
        }
        if ($needsCleanup) { Remove-FolderSafe -Path $workFolder }
    }
}

# ════════════════════════════════════════════════════════════════════════════
#  SESSION SUMMARY
# ════════════════════════════════════════════════════════════════════════════
$totalFailed = $counters.Failed + $counters.Error

Write-Host ''
Write-Host ''
Write-Host '╔══════════════════════════════════════════════════════╗' -ForegroundColor Cyan
Write-Host '║                    SESSION SUMMARY                   ║' -ForegroundColor Cyan
Write-Host '╚══════════════════════════════════════════════════════╝' -ForegroundColor Cyan
Write-Host ''
Write-Host ("  Total processed : {0}" -f $total)            -ForegroundColor White
Write-Host ("  ✔  Succeeded    : {0}" -f $counters.Success) -ForegroundColor Green
Write-Host ("  ⏭  Skipped      : {0}" -f $counters.Skipped) -ForegroundColor DarkGray
Write-Host ("  ✖  Failed       : {0}" -f $totalFailed) `
    -ForegroundColor ($totalFailed -gt 0 ? 'Red' : 'Green')

if ($failedItems.Count -gt 0) {
    Write-Host ''
    Write-Host '  ── Failed Items ──────────────────────────────────────' -ForegroundColor Red
    foreach ($fi in $failedItems) {
        Write-Host ("  • {0}"           -f $fi.Name)   -ForegroundColor Yellow
        Write-Host ("    Reason : {0}" -f $fi.Reason) -ForegroundColor Gray
        Write-Host ("    Exit   : {0}" -f $fi.Exit)   -ForegroundColor Gray
    }
}

Write-Host ''
Write-Host "  Log   : $LogFile"   -ForegroundColor DarkGray
Write-Host "  State : $StateFile" -ForegroundColor DarkGray
Write-Host ''