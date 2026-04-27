#Requires -Version 5.1
<#
.SYNOPSIS
    Scans PS3 game folders for EBOOT.BIN files and creates .m3u playlist files.

.DESCRIPTION
    Performs a single recursive scan for EBOOT.BIN files under the source folder,
    derives the game name from the directory structure, cleans the name for use
    as a filename, and writes one .m3u file per game pointing to the EBOOT.BIN.

    Existing .m3u files that already point to the correct EBOOT.BIN path are
    skipped unless -Force is specified.

    Optimized for performance with a single recursive scan rather than multiple
    per-directory passes. Compatible with both PowerShell 5.1 and PowerShell 7.
    ISE-compatible progress reporting (milestone percentages instead of progress bar).

.PARAMETER SourceDir
    Source folder containing PS3 game directories (e.g. dev_hdd0\game or dev_hdd0\disc).
    If not provided, prompted interactively.

.PARAMETER DestDir
    Output folder for .m3u files.
    If not provided, prompted interactively.

.PARAMETER Force
    When set, overwrites existing .m3u files even if they already point to the
    correct EBOOT.BIN path.

.PARAMETER Quiet
    Suppresses per-file output. Only shows summary.

.PARAMETER ShowSkipped
    When set, shows each file that was skipped because the .m3u already exists.

.PARAMETER ProgressInterval
    Show progress every N files. Default: 50.

.EXAMPLE
    .\PS3-M3U-Generator.ps1

.EXAMPLE
    .\PS3-M3U-Generator.ps1 -SourceDir "D:\PS3\dev_hdd0\game" -DestDir "D:\PS3\M3U" -Force

.VERSION
    2.0.0 - MIT license, config block with known PS3 defaults, replaced bare
            exit calls with Set-ExitCode, ISE-compatible progress reporting.
            Logic unchanged.
    1.0.0 - Initial release. Single-scan optimization, ISE compatibility.

.LICENSE
    MIT License
    Copyright (c) Paul Mardis
#>

param(
    [string]$SourceDir       = "D:\Arcade\System roms\Sony Playstation 3\dev_hdd0\game",
    [string]$DestDir         = "",
    [switch]$Force,
    [switch]$Quiet,
    [switch]$ShowSkipped,
    [int]$ProgressInterval   = 50
)

# ==============================================================================
# CONFIG -- Edit default paths here if needed.
# SourceDir and DestDir can also be passed as parameters or entered interactively.
# ==============================================================================
# $SourceDir already defaulted in param() above.
# $DestDir will be prompted if left empty.
# ==============================================================================

function Set-ExitCode {
    param([int]$Code)
    $global:LASTEXITCODE = $Code
}

$Config = @{
    LogFile   = ""
    StartTime = Get-Date
    IsISE     = $host.Name -eq "Windows PowerShell ISE Host"
    Stats     = @{ Found = 0; Created = 0; Skipped = 0; Errors = 0 }
}

# ------------------------------------------------------------------------------
# Output helpers
# ------------------------------------------------------------------------------
function Write-Msg {
    param([string]$Message, [string]$Color = "White")
    if (-not $Quiet) { Write-Host $Message -ForegroundColor $Color }
}

function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    if ($Config.LogFile -and (Test-Path (Split-Path $Config.LogFile -Parent))) {
        try {
            $Ts    = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            $Entry = "[$Ts] [$Level] $Message"
            Add-Content -Path $Config.LogFile -Value $Entry -Encoding UTF8 -ErrorAction SilentlyContinue
        } catch {}
    }
}

function Show-Progress {
    param([string]$Activity, [string]$Status, [int]$Pct = 0)

    if ($Config.IsISE) {
        if ($Pct -in @(0, 25, 50, 75, 100)) {
            Write-Msg ("  [{0}%] {1} - {2}" -f $Pct, $Activity, $Status) -Color Yellow
        }
    } else {
        Write-Progress -Activity $Activity -Status $Status -PercentComplete $Pct
    }
}

# ------------------------------------------------------------------------------
# Prompt for a folder path (creates destination if user confirms)
# ------------------------------------------------------------------------------
function Read-FolderPath {
    param([string]$Prompt, [string]$DefaultPath = "", [bool]$AllowCreate = $false)

    do {
        Write-Host ""
        if ($DefaultPath) {
            Write-Host $Prompt -ForegroundColor Cyan
            Write-Host ("  Default: {0}" -f $DefaultPath) -ForegroundColor Gray
            Write-Host "  Press Enter to use default, or type new path: " -ForegroundColor White -NoNewline
            $UserInput = Read-Host
            $Path = if ([string]::IsNullOrWhiteSpace($UserInput)) { $DefaultPath } else { $UserInput }
        } else {
            Write-Host $Prompt -ForegroundColor Cyan
            Write-Host "  Enter path: " -ForegroundColor White -NoNewline
            $Path = Read-Host
        }

        if ([string]::IsNullOrWhiteSpace($Path)) {
            Write-Msg "  Please enter a valid path." -Color Red
            continue
        }

        $Path = $Path.Trim().Trim('"').Trim("'")

        try {
            $Path = [System.Environment]::ExpandEnvironmentVariables($Path)
            $Path = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path)
        } catch {
            Write-Msg ("  Invalid path format: {0}" -f $Path) -Color Red
            continue
        }

        if (-not (Test-Path -LiteralPath $Path -PathType Container)) {
            if ($AllowCreate) {
                $Create = Read-Host "  Directory does not exist. Create it? (y/N)"
                if ($Create -eq 'y' -or $Create -eq 'Y') {
                    try {
                        New-Item -Path $Path -ItemType Directory -Force | Out-Null
                        Write-Msg ("  Created: {0}" -f $Path) -Color Green
                        Write-Log ("Created directory: {0}" -f $Path)
                        return $Path
                    } catch {
                        Write-Msg ("  Failed to create: {0}" -f $_) -Color Red
                        continue
                    }
                }
            } else {
                Write-Msg ("  Directory not found: {0}" -f $Path) -Color Red
                continue
            }
        } else {
            return $Path
        }
    } while ($true)
}

# ------------------------------------------------------------------------------
# Sanitize game name for use as a filename
# ------------------------------------------------------------------------------
function Get-CleanFileName {
    param([string]$FileName)

    if ([string]::IsNullOrWhiteSpace($FileName)) { return "UnknownGame" }

    $Clean = $FileName -replace '[<>:"/\\|?*]', '_'
    $Clean = $Clean -replace '\s+', '_'
    $Clean = $Clean -replace '_+', '_'
    $Clean = $Clean.Trim('._')

    $Reserved = @('CON','PRN','AUX','NUL','COM1','COM2','COM3','COM4','COM5',
                  'COM6','COM7','COM8','COM9','LPT1','LPT2','LPT3','LPT4',
                  'LPT5','LPT6','LPT7','LPT8','LPT9')
    if ($Reserved -contains $Clean.ToUpper()) { $Clean = "Game_$Clean" }
    if ($Clean.Length -gt 180) { $Clean = $Clean.Substring(0, 180) }
    if ([string]::IsNullOrWhiteSpace($Clean)) { $Clean = "UnknownGame" }

    return $Clean
}

# ------------------------------------------------------------------------------
# Single recursive scan for EBOOT.BIN files
# ------------------------------------------------------------------------------
function Find-EbootFiles {
    param([string]$SourcePath)

    Write-Msg "  Scanning for EBOOT.BIN files..." -Color Yellow
    Write-Log ("Starting scan of: {0}" -f $SourcePath)

    $GameList = @()

    try {
        $EbootFiles = @(Get-ChildItem -LiteralPath $SourcePath -Filter "EBOOT.BIN" -Recurse -File -ErrorAction SilentlyContinue)
        Write-Msg ("  Found {0} EBOOT.BIN file(s)." -f $EbootFiles.Count) -Color Green
        $Config.Stats.Found = $EbootFiles.Count

        if ($EbootFiles.Count -eq 0) { return $GameList }

        $Counter = 0
        foreach ($Eboot in $EbootFiles) {
            $Counter++
            $GameName   = ""
            $CurrentDir = $Eboot.Directory

            # Walk up directory tree looking for a meaningful name
            for ($i = 0; $i -lt 4; $i++) {
                if ($CurrentDir -and $CurrentDir.FullName -ne $SourcePath) {
                    if ($CurrentDir.Name -notmatch '^(PS3_GAME|USRDIR|EBOOT)$') {
                        $GameName = $CurrentDir.Name
                        break
                    }
                    $CurrentDir = $CurrentDir.Parent
                } else { break }
            }

            if ([string]::IsNullOrWhiteSpace($GameName) -or $GameName -match '^(PS3_GAME|USRDIR|EBOOT)$') {
                $PathParts = $Eboot.FullName.Replace($SourcePath, '').Trim('\').Split('\')
                $GameName  = $PathParts[0]
            }

            if ([string]::IsNullOrWhiteSpace($GameName)) { $GameName = "UnknownGame_$Counter" }

            $GameList += [PSCustomObject]@{
                GameName     = $GameName
                CleanName    = Get-CleanFileName -FileName $GameName
                EbootPath    = $Eboot.FullName
                LastModified = $Eboot.LastWriteTime
            }

            if ($Counter % $ProgressInterval -eq 0 -or $Counter -eq $EbootFiles.Count) {
                Show-Progress -Activity "Scanning" -Status ("Processed {0} of {1}" -f $Counter, $EbootFiles.Count) `
                    -Pct ([int](($Counter / $EbootFiles.Count) * 100))
            }
        }
    } catch {
        Write-Msg ("  Error during scan: {0}" -f $_) -Color Red
        Write-Log ("Error during scan: {0}" -f $_) -Level "ERROR"
    }

    return $GameList
}

# ------------------------------------------------------------------------------
# Check if existing .m3u already points to the correct EBOOT.BIN
# ------------------------------------------------------------------------------
function Test-ExistingM3u {
    param([string]$FilePath, [string]$ExpectedContent)

    if (-not (Test-Path -LiteralPath $FilePath)) { return $false }

    try {
        $Content = Get-Content -LiteralPath $FilePath -Raw -ErrorAction SilentlyContinue
        return ($null -ne $Content) -and ($Content.Trim() -eq $ExpectedContent.Trim())
    } catch { return $false }
}

# ------------------------------------------------------------------------------
# Write .m3u files
# ------------------------------------------------------------------------------
function New-M3uFiles {
    param([array]$GameList, [string]$OutputPath)

    Write-Msg "  Writing .m3u files..." -Color Yellow
    Write-Log ("Writing .m3u files to: {0}" -f $OutputPath)

    $Results      = @{ Created = 0; Skipped = 0; Errors = 0 }
    $CreatedFiles = @()
    $Counter      = 0

    foreach ($Game in $GameList) {
        $Counter++
        $M3uFile = "{0}.m3u" -f $Game.CleanName
        $M3uPath = Join-Path $OutputPath $M3uFile

        if (-not $Force -and (Test-ExistingM3u -FilePath $M3uPath -ExpectedContent $Game.EbootPath)) {
            if ($ShowSkipped) { Write-Msg ("  Skipped: {0}" -f $M3uFile) -Color Gray }
            $Results.Skipped++
            continue
        }

        try {
            [System.IO.File]::WriteAllText($M3uPath, $Game.EbootPath, [System.Text.UTF8Encoding]::new($false))

            try { (Get-Item -LiteralPath $M3uPath).LastWriteTime = $Game.LastModified } catch {}

            if (-not $Quiet -and $Results.Created -lt 5) {
                Write-Msg ("  Created: {0}" -f $M3uFile) -Color Green
                $CreatedFiles += $Game
            }

            Write-Log ("Created: {0} -> {1}" -f $M3uFile, $Game.EbootPath)
            $Results.Created++
        } catch {
            Write-Msg ("  Error: {0} -- {1}" -f $M3uFile, $_) -Color Red
            Write-Log ("Error: {0} -- {1}" -f $M3uFile, $_) -Level "ERROR"
            $Results.Errors++
        }

        if ($Counter % $ProgressInterval -eq 0 -or $Counter -eq $GameList.Count) {
            Show-Progress -Activity "Writing M3U files" -Status ("Created {0} of {1}" -f $Results.Created, $GameList.Count) `
                -Pct ([int](($Counter / $GameList.Count) * 100))
        }
    }

    $Config.Stats.Created      = $Results.Created
    $Config.Stats.Skipped      = $Results.Skipped
    $Config.Stats.Errors       = $Results.Errors
    $Config.Stats.CreatedFiles = $CreatedFiles
    return $Results
}

# ==============================================================================
# MAIN
# ==============================================================================
Write-Host ""
Write-Host "PS3 M3U Generator" -ForegroundColor Cyan
Write-Host "=================" -ForegroundColor Cyan
Write-Host ""

if ($PSVersionTable.PSVersion.Major -lt 7) {
    Write-Host ("  Note: PowerShell {0} detected. PS7+ gives better performance." -f $PSVersionTable.PSVersion.Major) -ForegroundColor Yellow
}

# Resolve source path
if ([string]::IsNullOrWhiteSpace($SourceDir) -or -not (Test-Path -LiteralPath $SourceDir -PathType Container)) {
    if (-not [string]::IsNullOrWhiteSpace($SourceDir)) {
        Write-Msg ("  Provided source path not found: {0}" -f $SourceDir) -Color Yellow
    }
    $SourceDir = Read-FolderPath -Prompt "Enter the source folder containing PS3 game directories" -AllowCreate $false
}

if (-not $SourceDir) {
    Write-Msg "  No valid source path. Exiting." -Color Red
    Set-ExitCode 1
    return
}

# Resolve destination path
if ([string]::IsNullOrWhiteSpace($DestDir) -or -not (Test-Path -LiteralPath $DestDir -PathType Container)) {
    $DestDir = Read-FolderPath -Prompt "Enter the destination folder for .m3u files" -AllowCreate $true
}

if (-not $DestDir) {
    Write-Msg "  No valid destination path. Exiting." -Color Red
    Set-ExitCode 1
    return
}

# Setup log
$Config.LogFile = Join-Path $DestDir ("PS3-M3U-Generator_{0}.log" -f (Get-Date -Format "yyyyMMdd_HHmmss"))
Write-Log "=== PS3 M3U Generator Started ==="
Write-Log ("SourceDir : {0}" -f $SourceDir)
Write-Log ("DestDir   : {0}" -f $DestDir)
Write-Log ("Force     : {0}" -f $Force)
Write-Log ("PS Version: {0}" -f $PSVersionTable.PSVersion)

Write-Host ("  Source      : {0}" -f $SourceDir) -ForegroundColor White
Write-Host ("  Destination : {0}" -f $DestDir)   -ForegroundColor White
Write-Host ("  Force       : {0}" -f $Force)      -ForegroundColor White
Write-Host ""

# Find and process
$GameList = Find-EbootFiles -SourcePath $SourceDir

if ($GameList.Count -gt 0) {
    Write-Host ""
    $Results = New-M3uFiles -GameList $GameList -OutputPath $DestDir

    # Summary
    $Duration = (Get-Date) - $Config.StartTime
    Write-Host ""
    Write-Host ("=" * 50) -ForegroundColor Cyan
    Write-Host "Summary" -ForegroundColor Cyan
    Write-Host ("=" * 50) -ForegroundColor Cyan
    Write-Host ("  Time elapsed : {0:mm\:ss}" -f $Duration) -ForegroundColor White
    Write-Host ("  EBOOT.BIN found : {0}" -f $Config.Stats.Found) -ForegroundColor Green
    Write-Host ("  .m3u created    : {0}" -f $Results.Created) -ForegroundColor Green
    if ($Results.Skipped -gt 0) {
        Write-Host ("  .m3u skipped    : {0}" -f $Results.Skipped) -ForegroundColor Yellow
    }
    if ($Results.Errors -gt 0) {
        Write-Host ("  Errors          : {0}" -f $Results.Errors) -ForegroundColor Red
    }

    if ($Config.Stats.CreatedFiles -and $Config.Stats.CreatedFiles.Count -gt 0) {
        Write-Host ""
        Write-Host "  Sample files created:" -ForegroundColor Cyan
        foreach ($Game in $Config.Stats.CreatedFiles) {
            Write-Host ("    {0}.m3u" -f $Game.CleanName) -ForegroundColor White
            Write-Host ("      -> {0}" -f $Game.EbootPath) -ForegroundColor Gray
        }
        if ($Results.Created -gt 5) {
            Write-Host ("    ... and {0} more" -f ($Results.Created - 5)) -ForegroundColor Gray
        }
    }

    Write-Host ("=" * 50) -ForegroundColor Cyan
} else {
    Write-Msg "  No EBOOT.BIN files found. No .m3u files created." -Color Yellow
}

Write-Log "=== PS3 M3U Generator Finished ==="

if ($Config.IsISE) {
    Write-Msg "  [100%] Complete" -Color Green
}

Write-Host ""
Set-ExitCode 0
