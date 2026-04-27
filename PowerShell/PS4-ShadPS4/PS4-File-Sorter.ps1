#Requires -Version 5.1
<#
.SYNOPSIS
    Sorts PS4 PKG files in a staging folder into Game, Update, DLC, and Misc subfolders.

.DESCRIPTION
    Scans a staging folder for .pkg files, classifies each one based on filename
    patterns, moves it into a named subfolder under the appropriate category folder,
    and moves any associated companion files (.nfo, .txt, .jpg, .jpeg) alongside it.

    Classification rules (applied in order):
      Update : filename matches  A####-V####
      DLC    : filename contains DLC, ADDON, TRACK, WEAPON, PACK, SEASONPASS,
                                 COSTUME, STYLE, CHARACTER, or EXPANSION
      Game   : everything else that is not an Update
      Misc   : fallthrough (rare)

    Note: for more accurate type detection, use ShadPS4-PKG-Installer.ps1 which
    calls pkg_extractor.exe --check-type on each file.

.PARAMETER DryRun
    When $true (default), shows what would be moved without touching any files.
    Set to $false in the CONFIG block to apply changes.

.EXAMPLE
    .\PS4-File-Sorter.ps1

.VERSION
    2.0.0 - Config block, DryRun default, CSV log, MIT header, ISE compatible.
    1.0.0 - Initial release.

.LICENSE
    MIT License
    Copyright (c) Paul Mardis
#>

# ==============================================================================
# CONFIG -- Edit this block. Do not put paths anywhere else in this script.
# ==============================================================================
$SourcePath = "D:\Arcade\System roms\Sony Playstation 4\Downloading PS4"
$DryRun     = $true    # Set to $false to actually move files
# ==============================================================================

function Set-ExitCode {
    param([int]$Code)
    $global:LASTEXITCODE = $Code
}

# ------------------------------------------------------------------------------
# Pre-flight
# ------------------------------------------------------------------------------
if (-not (Test-Path -LiteralPath $SourcePath)) {
    Write-Host "ERROR: Source path not found: $SourcePath" -ForegroundColor Red
    Set-ExitCode 1
    return
}

# ------------------------------------------------------------------------------
# Destination folders (always inside SourcePath)
# ------------------------------------------------------------------------------
$Folders = @{
    Game   = Join-Path $SourcePath "Games"
    Update = Join-Path $SourcePath "Updates"
    DLC    = Join-Path $SourcePath "DLC"
    Misc   = Join-Path $SourcePath "Misc"
}

$LogDir  = Join-Path $SourcePath "_Logs"
$LogFile = Join-Path $LogDir ("PS4-File-Sorter_{0}.csv" -f (Get-Date -Format "yyyyMMdd_HHmmss"))

# ------------------------------------------------------------------------------
# Classify a PKG filename into Game / Update / DLC / Misc
# ------------------------------------------------------------------------------
function Get-PkgType {
    param([string]$FileName)

    if ($FileName -match 'A\d{4}-V\d{4}') {
        return "Update"
    }

    if ($FileName -match '(?i)(DLC|ADDON|TRACK|WEAPON|PACK|SEASONPASS|COSTUME|STYLE|CHARACTER|EXPANSION)') {
        return "DLC"
    }

    if ($FileName -notmatch 'A\d{4}-V\d{4}') {
        return "Game"
    }

    return "Misc"
}

# ------------------------------------------------------------------------------
# CSV log
# ------------------------------------------------------------------------------
$LogRows = [System.Collections.Generic.List[PSCustomObject]]::new()

function Write-LogRow {
    param(
        [string]$FileName,
        [string]$Type,
        [string]$SourceFile,
        [string]$Destination,
        [string]$Action,
        [string]$Status,
        [string]$Error
    )
    $LogRows.Add([PSCustomObject]@{
        Timestamp   = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
        FileName    = $FileName
        Type        = $Type
        SourceFile  = $SourceFile
        Destination = $Destination
        Action      = $Action
        Status      = $Status
        Error       = $Error
    })
}

# ------------------------------------------------------------------------------
# Setup folders (create only if doing a live run)
# ------------------------------------------------------------------------------
if (-not $DryRun) {
    foreach ($Folder in $Folders.Values) {
        if (-not (Test-Path -LiteralPath $Folder)) {
            New-Item -ItemType Directory -Path $Folder | Out-Null
        }
    }
    if (-not (Test-Path -LiteralPath $LogDir)) {
        New-Item -ItemType Directory -Path $LogDir | Out-Null
    }
}

# ==============================================================================
# BANNER
# ==============================================================================
Write-Host ""
Write-Host "PS4 File Sorter" -ForegroundColor Cyan
Write-Host "===============" -ForegroundColor Cyan
Write-Host "  Source : $SourcePath"
Write-Host "  DryRun : $DryRun"
Write-Host ""
if ($DryRun) {
    Write-Host "  [DRY RUN] No files will be moved." -ForegroundColor Yellow
    Write-Host ""
}

# ------------------------------------------------------------------------------
# Gather PKGs (exclude files already inside a category subfolder)
# ------------------------------------------------------------------------------
$ExcludePaths = $Folders.Values + @($LogDir)

$PkgFiles = Get-ChildItem -LiteralPath $SourcePath -Filter "*.pkg" -File |
    Where-Object {
        $f = $_.FullName
        $exclude = $false
        foreach ($ep in $ExcludePaths) {
            if ($f.StartsWith($ep + '\') -or $f.StartsWith($ep + '/')) {
                $exclude = $true
                break
            }
        }
        -not $exclude
    } |
    Sort-Object Name

if ($PkgFiles.Count -eq 0) {
    Write-Host "  No PKG files found in: $SourcePath" -ForegroundColor Yellow
    Write-Host ""
    Set-ExitCode 0
    return
}

Write-Host "  Found $($PkgFiles.Count) PKG file(s) to process." -ForegroundColor White
Write-Host ""

# ==============================================================================
# MAIN LOOP
# ==============================================================================
$CountMoved   = 0
$CountSkipped = 0
$CountFailed  = 0

foreach ($Pkg in $PkgFiles) {
    $FileName   = $Pkg.Name
    $BaseName   = [System.IO.Path]::GetFileNameWithoutExtension($FileName)
    $Type       = Get-PkgType -FileName $FileName
    $TypeFolder = $Folders[$Type]
    $SubFolder  = Join-Path $TypeFolder $BaseName
    $DestPkg    = Join-Path $SubFolder $FileName

    Write-Host ("  {0,-60} -> {1}" -f $FileName, $Type) -ForegroundColor Gray

    if ($DryRun) {
        Write-Host ("    WOULD MOVE to: {0}" -f $SubFolder) -ForegroundColor Cyan
        Write-LogRow -FileName $FileName -Type $Type -SourceFile $Pkg.FullName `
            -Destination $SubFolder -Action "Move" -Status "DryRun" -Error ""
        $CountMoved++
        continue
    }

    try {
        # Create subfolder
        if (-not (Test-Path -LiteralPath $SubFolder)) {
            New-Item -ItemType Directory -Path $SubFolder | Out-Null
        }

        # Move PKG
        Move-Item -LiteralPath $Pkg.FullName -Destination $DestPkg -Force -ErrorAction Stop
        Write-Host ("    Moved  -> {0}" -f $SubFolder) -ForegroundColor Green
        Write-LogRow -FileName $FileName -Type $Type -SourceFile $Pkg.FullName `
            -Destination $SubFolder -Action "Move" -Status "Success" -Error ""
        $CountMoved++

        # Move companion files (.nfo .txt .jpg .jpeg with matching base name)
        $CompanionPatterns = @(
            "*$BaseName*.nfo"
            "*$BaseName*.txt"
            "*$BaseName*.jpg"
            "*$BaseName*.jpeg"
        )

        foreach ($Pattern in $CompanionPatterns) {
            $Companions = Get-ChildItem -LiteralPath $SourcePath -Filter $Pattern -File -ErrorAction SilentlyContinue
            foreach ($Companion in $Companions) {
                $CompanionDest = Join-Path $SubFolder $Companion.Name
                try {
                    Move-Item -LiteralPath $Companion.FullName -Destination $CompanionDest -Force -ErrorAction Stop
                    Write-Host ("    Companion: {0}" -f $Companion.Name) -ForegroundColor DarkGreen
                    Write-LogRow -FileName $Companion.Name -Type $Type -SourceFile $Companion.FullName `
                        -Destination $SubFolder -Action "MoveCompanion" -Status "Success" -Error ""
                }
                catch {
                    Write-Host ("    Companion FAILED: {0} -- {1}" -f $Companion.Name, $_.Exception.Message) -ForegroundColor Yellow
                    Write-LogRow -FileName $Companion.Name -Type $Type -SourceFile $Companion.FullName `
                        -Destination $SubFolder -Action "MoveCompanion" -Status "Failed" -Error $_.Exception.Message
                }
            }
        }
    }
    catch {
        Write-Host ("    FAILED: {0}" -f $_.Exception.Message) -ForegroundColor Red
        Write-LogRow -FileName $FileName -Type $Type -SourceFile $Pkg.FullName `
            -Destination $SubFolder -Action "Move" -Status "Failed" -Error $_.Exception.Message
        $CountFailed++
    }
}

# ==============================================================================
# WRITE LOG
# ==============================================================================
if ($LogRows.Count -gt 0 -and -not $DryRun) {
    try {
        $LogRows | Export-Csv -LiteralPath $LogFile -NoTypeInformation -Encoding UTF8
        Write-Host ""
        Write-Host "  Log: $LogFile" -ForegroundColor Gray
    }
    catch {
        Write-Host "  WARNING: Could not write log -- $_" -ForegroundColor Yellow
    }
}

# ==============================================================================
# SUMMARY
# ==============================================================================
Write-Host ""
Write-Host "Summary" -ForegroundColor White
Write-Host ("  {0,-12} {1}" -f "Moved :", $CountMoved)   -ForegroundColor Green
Write-Host ("  {0,-12} {1}" -f "Skipped :", $CountSkipped) -ForegroundColor Yellow
Write-Host ("  {0,-12} {1}" -f "Failed :", $CountFailed)  -ForegroundColor $(if ($CountFailed -gt 0) { "Red" } else { "Gray" })

if ($DryRun) {
    Write-Host ""
    Write-Host "  DRY RUN complete. Set `$DryRun = `$false in the CONFIG block to apply." -ForegroundColor Yellow
}

Write-Host ""
Set-ExitCode 0
