#Requires -Version 7.0
<#
.SYNOPSIS
    RPCS3 Batch PKG Installer - PowerShell 7 edition
.DESCRIPTION
    Processes folders and ZIPs containing PS3 PKG/RAP files, installs via RPCS3,
    tracks state, and organises sources into _Installed / _Failed.
.PARAMETER SourceRoot
    Folder containing game subfolders and/or ZIP archives.
.PARAMETER Rpcs3Exe
    Full path to rpcs3.exe
.PARAMETER Rpcs3Root
    Root folder of the RPCS3 installation (contains dev_hdd0).
.PARAMETER TempRoot
    Scratch folder used for ZIP extraction.
.PARAMETER SevenZipExe
    Path to 7z.exe
.PARAMETER FolderPollSeconds
    How long to poll for a new game folder after each install. Default: 20.
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
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)
$OutputEncoding           = [System.Text.UTF8Encoding]::new($false)

# ============================================================================
#  DERIVED PATHS
# ============================================================================
$GameRoot      = Join-Path $Rpcs3Root  'dev_hdd0\game'
$RapTargetRoot = Join-Path $Rpcs3Root  'dev_hdd0\home\00000001\exdata'
$PS3Root       = 'D:\Arcade\System roms\Sony Playstation 3'
$StateRoot     = Join-Path $PS3Root    '_Automation'
$InstalledRoot = Join-Path $PS3Root    '_Installed'
$FailedRoot    = Join-Path $PS3Root    '_Failed'
$LogFile       = Join-Path $StateRoot  'install_log.csv'
$StateFile     = Join-Path $StateRoot  'installed_state.json'

# ============================================================================
#  BANNER + RUN TYPE PROMPT
# ============================================================================
function Show-Banner {
    Write-Host ''
    Write-Host '======================================================' -ForegroundColor Cyan
    Write-Host '         RPCS3 Batch PKG Installer  (PS7)            ' -ForegroundColor Cyan
    Write-Host '======================================================' -ForegroundColor Cyan
    Write-Host ''
}

function Get-RunType {
    Write-Host '  Select run type:' -ForegroundColor Yellow
    Write-Host ''
    Write-Host '  [1]  Normal   - Install new items, skip already-installed' -ForegroundColor White
    Write-Host '  [2]  Dry Run  - Preview only, no changes made'             -ForegroundColor White
    Write-Host '  [3]  Force    - Reinstall ALL items (ignore state file)'   -ForegroundColor White
    Write-Host '  [4]  Failed   - Retry items currently in _Failed folder'   -ForegroundColor White
    Write-Host '  [Q]  Quit'                                                  -ForegroundColor White
    Write-Host ''
    do { $key = (Read-Host '  Your choice').Trim().ToUpper() }
    while ($key -notin @('1','2','3','4','Q'))
    return $key
}

Show-Banner
$runChoice = Get-RunType

if ($runChoice -eq 'Q') {
    Write-Host ''
    Write-Host '  Aborted.' -ForegroundColor Red
    return
}

$DryRun         = $runChoice -eq '2'
$ForceReinstall = $runChoice -eq '3'
$RetryFailed    = $runChoice -eq '4'
$modeLabel      = @{ '1'='Normal'; '2'='Dry Run'; '3'='Force Reinstall'; '4'='Retry Failed' }[$runChoice]

Write-Host ''
Write-Host "  Mode   : $modeLabel"  -ForegroundColor Cyan
Write-Host "  Source : $SourceRoot" -ForegroundColor Cyan
Write-Host "  RPCS3  : $Rpcs3Exe"   -ForegroundColor Cyan
Write-Host "  7-Zip  : $SevenZipExe" -ForegroundColor Cyan
Write-Host ''

# ============================================================================
#  HELPER FUNCTIONS
# ============================================================================

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
        [string] $SourcePath, [string] $PkgPath, [string] $RapPath,
        [string] $Status, [string] $Details, [string] $DetectedGameFolder,
        [int]    $ExitCode
    )
    $row = [pscustomobject] @{
        Timestamp          = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
        SourcePath         = $SourcePath
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
    param([hashtable] $Before, [datetime] $InstallStartUtc, [int] $TimeoutSeconds = 20)
    $deadline = (Get-Date).AddSeconds($TimeoutSeconds)
    while ((Get-Date) -lt $deadline) {
        Start-Sleep -Milliseconds 1000
        $after = Get-GameSnapshot -Path $GameRoot
        foreach ($p in $after.Keys) {
            if (-not $Before.ContainsKey($p)) { return $p }
            if ($after[$p] -gt $Before[$p] -and $after[$p] -ge $InstallStartUtc.AddMinutes(-1)) {
                return $p
            }
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
    if ($LASTEXITCODE -ne 0) {
        throw "7-Zip failed with exit code $LASTEXITCODE"
    }
}

# ============================================================================
#  P/INVOKE + SENDKEYS - auto-click the RPCS3 PKG Installation dialog
# ============================================================================
Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class User32 {
    [DllImport("user32.dll", CharSet=CharSet.Unicode)]
    public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);
    [DllImport("user32.dll")]
    public static extern bool SetForegroundWindow(IntPtr hWnd);
    [DllImport("user32.dll")]
    public static extern bool IsWindowVisible(IntPtr hWnd);
}
'@
Add-Type -AssemblyName System.Windows.Forms

function Invoke-RPCS3Install {
    param([string] $PkgPath, [int] $DialogTimeoutSeconds = 60)

    $proc     = Start-Process -FilePath $Rpcs3Exe `
                              -ArgumentList '--installpkg', $PkgPath `
                              -WorkingDirectory $Rpcs3Root -PassThru
    $deadline = (Get-Date).AddSeconds($DialogTimeoutSeconds)
    $clicked  = $false

    Write-Host '    Waiting for dialog...' -NoNewline -ForegroundColor DarkGray

    while ((Get-Date) -lt $deadline) {
        Start-Sleep -Milliseconds 600
        if ($proc.HasExited) { break }

        $hwnd = [User32]::FindWindow($null, 'PKG Installation')
        if ($hwnd -ne [IntPtr]::Zero -and [User32]::IsWindowVisible($hwnd)) {
            Start-Sleep -Milliseconds 500
            [User32]::SetForegroundWindow($hwnd) | Out-Null
            Start-Sleep -Milliseconds 300
            [System.Windows.Forms.SendKeys]::SendWait('{ENTER}')
            $clicked = $true
            Write-Host "`r    Installing...                    " -ForegroundColor DarkGreen
            break
        }
    }

    if (-not $clicked -and -not $proc.HasExited) {
        Write-Host "`r    Timed out - killing RPCS3.          " -ForegroundColor Red
        $proc | Stop-Process -Force -ErrorAction SilentlyContinue
        return -1
    }

    $proc.WaitForExit()
    return $proc.ExitCode
}

# ============================================================================
#  PRE-FLIGHT CHECKS
# ============================================================================
foreach ($check in @(
    @{ Path = $Rpcs3Exe;    Label = 'RPCS3 executable' }
    @{ Path = $SevenZipExe; Label = '7-Zip executable'  }
    @{ Path = $GameRoot;    Label = 'RPCS3 game folder'  }
)) {
    if (-not (Test-Path -LiteralPath $check.Path)) {
        Write-Host ("  ERROR: {0} not found: {1}" -f $check.Label, $check.Path) -ForegroundColor Red
        return
    }
}

foreach ($dir in @($StateRoot, $InstalledRoot, $FailedRoot, $TempRoot, $RapTargetRoot)) {
    Initialize-Folder -Path $dir
}

# ============================================================================
#  COLLECT WORK ITEMS  (plain folders with PKGs  +  ZIP archives)
# ============================================================================
$scanRoot        = $RetryFailed ? $FailedRoot : $SourceRoot
$skipFolderNames = @('_DeDuplication','_Installed','_Failed','_Automation','Complete')
$workItems       = [System.Collections.Generic.List[pscustomobject]]::new()

# Plain folders containing PKGs
Get-ChildItem -LiteralPath $scanRoot -Directory |
    Where-Object { $_.Name -notin $skipFolderNames -and $_.Name -notmatch '^_' } |
    ForEach-Object {
        $hasPkg = [bool](Get-ChildItem -LiteralPath $_.FullName -Recurse -File -Filter *.pkg `
                            -ErrorAction SilentlyContinue | Select-Object -First 1)
        if ($hasPkg) {
            $workItems.Add([pscustomobject] @{
                Type        = 'folder'
                SourcePath  = $_.FullName
                DisplayName = $_.Name
                SizeBytes   = 0
            })
        }
    }

# ZIP archives
Get-ChildItem -LiteralPath $scanRoot -Recurse -File -Filter *.zip |
    Where-Object { $_.FullName -notmatch '\\_installed\\' -and $_.FullName -notmatch '\\_failed\\' } |
    ForEach-Object {
        $workItems.Add([pscustomobject] @{
            Type        = 'zip'
            SourcePath  = $_.FullName
            DisplayName = $_.BaseName
            SizeBytes   = $_.Length
        })
    }

$workItems = [pscustomobject[]]($workItems | Sort-Object DisplayName)

if (-not $workItems) {
    Write-Host '  No installable items found.' -ForegroundColor Yellow
    return
}

$zipCount    = ($workItems | Where-Object Type -eq 'zip').Count
$folderCount = ($workItems | Where-Object Type -eq 'folder').Count
$total       = $workItems.Count

Write-Host ("  Found {0} items  ({1} folder, {2} ZIP)" -f $total, $folderCount, $zipCount) `
    -ForegroundColor White
Write-Host ''

# ============================================================================
#  KILL ANY EXISTING RPCS3 INSTANCE
# ============================================================================
$existingRpcs3 = Get-Process -Name 'rpcs3' -ErrorAction SilentlyContinue
if ($existingRpcs3) {
    Write-Host '  WARNING: RPCS3 is already running.' -ForegroundColor Yellow
    $kill = (Read-Host '  Kill it and continue? [Y/N]').Trim().ToUpper()
    if ($kill -eq 'Y') {
        $existingRpcs3 | Stop-Process -Force
        Start-Sleep -Seconds 2
        Write-Host '  RPCS3 closed.' -ForegroundColor Green
    } else {
        Write-Host '  Aborted.' -ForegroundColor Red
        return
    }
}

$state       = Import-State
$counters    = @{ Success = 0; Skipped = 0; Failed = 0; Error = 0 }
$failedItems = [System.Collections.Generic.List[pscustomobject]]::new()

# ============================================================================
#  MAIN LOOP
# ============================================================================
$idx = 0
foreach ($item in $workItems) {
    $idx++
    $isZip      = $item.Type -eq 'zip'
    $sourcePath = $item.SourcePath
    $dispName   = $item.DisplayName
    $typeLabel  = $isZip ? 'ZIP' : 'DIR'

    Write-Host ''
    Write-Host ('--- [{0}/{1}] {2} [{3}]' -f $idx, $total, $dispName, $typeLabel) -ForegroundColor Cyan

    # -- State key --
    # ZIPs    : SHA256 of the ZIP file
    # Folders : hash of "name|pkgfile|pkgsize" - never hash a 13GB PKG file
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

    # -- Skip check --
    if (-not $ForceReinstall -and -not $RetryFailed -and $itemHash -and $state.ContainsKey($itemHash)) {
        Write-Host '  SKIPPED (already installed)' -ForegroundColor DarkGray
        $counters.Skipped++
        Write-LogRow -SourcePath $sourcePath -PkgPath '' -RapPath '' -Status 'Skipped' `
            -Details 'Already in state file' `
            -DetectedGameFolder ($state[$itemHash]?['DetectedGameFolder'] ?? '') -ExitCode 0
        continue
    }

    # -- Resolve working folder --
    $workFolder   = $sourcePath
    $needsCleanup = $false

    try {
        if ($isZip) {
            $workFolder   = Join-Path $TempRoot $dispName
            $needsCleanup = $true
            $sizeMB       = [math]::Round($item.SizeBytes / 1MB, 1)
            Write-Host ("  STAGE 1/3 - Extracting ({0} MB)..." -f $sizeMB) -ForegroundColor Yellow
            Remove-FolderSafe -Path $workFolder
            Initialize-Folder -Path $workFolder
            Expand-WithSevenZip -ArchivePath $sourcePath -DestinationPath $workFolder
            Write-Host '  Extraction complete.' -ForegroundColor Green
        } else {
            Write-Host '  STAGE 1/3 - Using folder directly.' -ForegroundColor Green
        }

        $pkgFiles = [System.IO.FileInfo[]](Get-ChildItem -LiteralPath $workFolder -Recurse -File -Filter *.pkg)
        $rapFiles = [System.IO.FileInfo[]](Get-ChildItem -LiteralPath $workFolder -Recurse -File -Filter *.rap)

        Write-Host ("  Found: {0} PKG, {1} RAP" -f $pkgFiles.Count, $rapFiles.Count) -ForegroundColor DarkGray

        if (-not $pkgFiles) {
            Write-Host '  FAILED - No PKG found.' -ForegroundColor Red
            $counters.Failed++
            $failedItems.Add([pscustomobject] @{ Name = $dispName; Reason = 'No PKG found'; Exit = 'N/A' })
            Write-LogRow -SourcePath $sourcePath -PkgPath '' -RapPath '' -Status 'Failed' `
                -Details 'No PKG found' -DetectedGameFolder '' -ExitCode -1
            if ($isZip) { Move-ToFolder -Path $sourcePath -Destination $FailedRoot | Out-Null }
            if ($needsCleanup) { Remove-FolderSafe -Path $workFolder }
            continue
        }

        if ($DryRun) {
            foreach ($pkg in $pkgFiles) {
                $rap = Get-MatchingRap -Pkg $pkg -Raps $rapFiles
                $rapNote = $rap ? "RAP: $($rap.Name)" : 'No RAP'
                Write-Host ("  DRY RUN - {0} [{1}]" -f $pkg.Name, $rapNote) -ForegroundColor DarkCyan
                Write-LogRow -SourcePath $sourcePath -PkgPath $pkg.FullName `
                    -RapPath ($rap?.FullName ?? '') -Status 'DryRun' `
                    -Details 'No changes made' -DetectedGameFolder '' -ExitCode 0
            }
            if ($needsCleanup) { Remove-FolderSafe -Path $workFolder }
            continue
        }

        $itemSuccess     = $true
        $detectedFolders = [System.Collections.Generic.List[string]]::new()

        foreach ($pkg in $pkgFiles) {
            $rap = Get-MatchingRap -Pkg $pkg -Raps $rapFiles

            # Copy RAP with verification
            $rapDestPath = ''
            if ($rap) {
                $rapDestPath = Join-Path $RapTargetRoot $rap.Name
                Copy-Item -LiteralPath $rap.FullName -Destination $rapDestPath -Force
                if (-not (Test-Path -LiteralPath $rapDestPath)) {
                    throw "RAP copy failed - file missing: $rapDestPath"
                }
                $srcHash  = Get-FileHashSafe -Path $rap.FullName
                $destHash = Get-FileHashSafe -Path $rapDestPath
                if ($srcHash -ne $destHash) {
                    throw "RAP hash mismatch: $($rap.Name)"
                }
                Write-Host ("  RAP OK : {0}" -f $rap.Name) -ForegroundColor Green
            } else {
                Write-Host '  RAP    : none' -ForegroundColor DarkGray
            }

            $before          = Get-GameSnapshot -Path $GameRoot
            $installStartUtc = (Get-Date).ToUniversalTime()

            Write-Host ("  STAGE 2/3 - Installing: {0}" -f $pkg.Name) -ForegroundColor Yellow

            $exitCode = Invoke-RPCS3Install -PkgPath $pkg.FullName

            Write-Host '  STAGE 3/3 - Waiting for game folder...' -NoNewline -ForegroundColor DarkGray
            $detectedFolder = Wait-ForNewGameFolder -Before $before `
                                                    -InstallStartUtc $installStartUtc `
                                                    -TimeoutSeconds $FolderPollSeconds
            Write-Host "`r                                              `r" -NoNewline

            if ($exitCode -eq 0 -and $detectedFolder) {
                $detectedFolders.Add($detectedFolder)
                $folderName = [System.IO.Path]::GetFileName($detectedFolder)
                Write-Host ("  OK - {0} -> {1}" -f $pkg.Name, $folderName) -ForegroundColor Green
                Write-LogRow -SourcePath $sourcePath -PkgPath $pkg.FullName `
                    -RapPath $rapDestPath -Status 'Success' `
                    -Details 'RPCS3 exit 0, folder detected' `
                    -DetectedGameFolder $detectedFolder -ExitCode $exitCode
            } else {
                $itemSuccess = $false
                $reason = ($exitCode -ne 0) `
                    ? "RPCS3 exit code $exitCode" `
                    : "No game folder after ${FolderPollSeconds}s"
                Write-Host ("  FAILED - {0}" -f $reason) -ForegroundColor Red
                $failedItems.Add([pscustomobject] @{ Name = $dispName; Reason = $reason; Exit = $exitCode })
                Write-LogRow -SourcePath $sourcePath -PkgPath $pkg.FullName `
                    -RapPath $rapDestPath -Status 'Failed' -Details $reason `
                    -DetectedGameFolder ($detectedFolder ?? '') -ExitCode $exitCode
            }
        }

        if ($itemSuccess) {
            $counters.Success++
            $movedTo = ''
            if ($isZip) { $movedTo = Move-ToFolder -Path $sourcePath -Destination $InstalledRoot }
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
            if ($isZip) { Move-ToFolder -Path $sourcePath -Destination $FailedRoot | Out-Null }
        }

        if ($needsCleanup) { Remove-FolderSafe -Path $workFolder }
    }
    catch {
        $errMsg = $_.Exception.Message
        Write-Host ("  ERROR - {0}" -f $errMsg) -ForegroundColor Red
        $counters.Error++
        $failedItems.Add([pscustomobject] @{ Name = $dispName; Reason = $errMsg; Exit = 'exception' })
        Write-LogRow -SourcePath $sourcePath -PkgPath '' -RapPath '' -Status 'Error' `
            -Details $errMsg -DetectedGameFolder '' -ExitCode -1
        if ($isZip -and (Test-Path -LiteralPath $sourcePath)) {
            Move-ToFolder -Path $sourcePath -Destination $FailedRoot | Out-Null
        }
        if ($needsCleanup) { Remove-FolderSafe -Path $workFolder }
    }
}

# ============================================================================
#  SESSION SUMMARY
# ============================================================================
$totalFailed = $counters.Failed + $counters.Error

Write-Host ''
Write-Host '======================================================' -ForegroundColor Cyan
Write-Host '                  SESSION SUMMARY                    ' -ForegroundColor Cyan
Write-Host '======================================================' -ForegroundColor Cyan
Write-Host ''
Write-Host ("  Total     : {0}" -f $total)            -ForegroundColor White
Write-Host ("  Succeeded : {0}" -f $counters.Success) -ForegroundColor Green
Write-Host ("  Skipped   : {0}" -f $counters.Skipped) -ForegroundColor DarkGray
Write-Host ("  Failed    : {0}" -f $totalFailed) `
    -ForegroundColor ($totalFailed -gt 0 ? 'Red' : 'Green')

if ($failedItems.Count -gt 0) {
    Write-Host ''
    Write-Host '  -- Failed Items --' -ForegroundColor Red
    foreach ($fi in $failedItems) {
        Write-Host ("  * {0}"          -f $fi.Name)   -ForegroundColor Yellow
        Write-Host ("    Reason : {0}" -f $fi.Reason) -ForegroundColor Gray
        Write-Host ("    Exit   : {0}" -f $fi.Exit)   -ForegroundColor Gray
    }
}

Write-Host ''
Write-Host "  Log   : $LogFile"   -ForegroundColor DarkGray
Write-Host "  State : $StateFile" -ForegroundColor DarkGray
Write-Host ''
