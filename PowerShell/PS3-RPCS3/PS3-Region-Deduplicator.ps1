#Requires -Version 5.1
<#
.SYNOPSIS
    Identifies regional duplicate PS3 ZIP archives and moves lower-priority
    copies to a destination folder.

.DESCRIPTION
    Scans a source folder for ZIP files (or other configured extensions),
    groups files by normalized title (stripping region, version, and language
    tags), selects the preferred copy based on region priority and version
    number, and moves the rest to a destination folder.

    Region priority (index 0 = highest):
      USA > World > Europe > Japan > Asia > Korea > Hong Kong > Taiwan >
      China > Singapore > Thailand > Australia > New Zealand > Canada >
      UK > France > Germany > Italy > Spain > Netherlands > Norway >
      Sweden > Denmark > Finland > Poland > Russia > Brazil > Mexico >
      Latin America > Unknown

    Version/revision selection: prefers higher version numbers when the
    region score is tied (e.g. v2.01 beats v1.00).

    Companion files are never treated as duplicates. These include files
    tagged as: DLC, Add-On, Patch, Update, Demo, BIOS, Theme, and others.
    Files smaller than $MinFullGameSizeMB are also treated as companions.

    Run with DryRun mode (option 1) first to review the CSV log before
    moving any files.

.EXAMPLE
    .\PS3-Region-Deduplicator.ps1

.NOTES
    Source and destination paths are entered interactively at startup.
    The mode prompt (Dry Run / Live) is also interactive.

    The CSV log is always written regardless of mode, making it safe to
    review what will happen before running in Live mode.

.VERSION
    2.0.0 - MIT license, replaced Unicode dashes in output with ASCII.
             Logic, companion detection, and CSV output unchanged.
    1.0.0 - Initial release.

.LICENSE
    MIT License
    Copyright (c) Paul Mardis
#>

# ==============================================================================
# CONFIG
# ==============================================================================
# File types to scan for duplicates. Add '*.iso', '*.pkg' if needed.
$FileExtensions = @('*.zip')

# Files smaller than this (MB) are treated as companions, not full games.
# Set to 0 to disable size-based companion detection.
$MinFullGameSizeMB = 20

$CompanionTags = @(
    'Unlock Key','PS2 Classics','PC Engine','Neo Geo',
    'Alt','Sample','Demo','Kiosk','Proto','Debug',
    'Program','BIOS','Theme','DLC','Add-On','Patch',
    'Update','Expansion','Pack','Avatar',
    'Dynamic Theme','Static Theme'
)

$KnownRegions = @(
    'USA','World','Europe','Japan','Asia','Korea',
    'Hong Kong','Taiwan','China','Singapore','Thailand',
    'Australia','New Zealand','Canada','UK',
    'France','Germany','Italy','Spain','Netherlands',
    'Norway','Sweden','Denmark','Finland','Poland',
    'Russia','Brazil','Mexico','Latin America'
)

# Index 0 = highest priority
$RegionPriority = @(
    'USA','World','Europe','Japan','Asia','Korea',
    'Hong Kong','Taiwan','China','Singapore','Thailand',
    'Australia','New Zealand','Canada','UK',
    'France','Germany','Italy','Spain','Netherlands',
    'Norway','Sweden','Denmark','Finland','Poland',
    'Russia','Brazil','Mexico','Latin America','Unknown'
)
# ==============================================================================

function Set-ExitCode {
    param([int]$Code)
    $global:LASTEXITCODE = $Code
}

# ------------------------------------------------------------------------------
# Companion detection
# ------------------------------------------------------------------------------
function Test-IsCompanionFile {
    param([string]$BaseName)
    foreach ($Tag in $CompanionTags) {
        $Escaped = [regex]::Escape($Tag)
        if ($BaseName -match "\($Escaped(?:\s+[^)]+)?\)") { return $true }
    }
    return $false
}

# ------------------------------------------------------------------------------
# Region extraction
# ------------------------------------------------------------------------------
function Get-RegionFromName {
    param([string]$BaseName)
    $Matches = [regex]::Matches($BaseName, '\(([^()]*)\)')
    foreach ($Match in $Matches) {
        $Token = $Match.Groups[1].Value.Trim()
        if ($KnownRegions -contains $Token) { return $Token }
    }
    return 'Unknown'
}

# ------------------------------------------------------------------------------
# Normalized title key (strips region, version, language, revision tags)
# ------------------------------------------------------------------------------
function Get-NormalizedKey {
    param([string]$BaseName)

    $Result = $BaseName

    foreach ($Region in $KnownRegions) {
        $Escaped = [regex]::Escape($Region)
        $Result = [regex]::Replace($Result, "\s*\($Escaped\)", '', 'IgnoreCase')
    }

    $Result = [regex]::Replace($Result, '\s*\([A-Za-z]{2}(?:,[A-Za-z]{2})+\)', '')
    $Result = [regex]::Replace($Result, '\s*\(v[\d]+\.[\d]+[^)]*\)', '', 'IgnoreCase')
    $Result = [regex]::Replace($Result, '\s*\(Rev\s+[^)]+\)', '', 'IgnoreCase')
    $Result = [regex]::Replace($Result, '\s*\(Version\s+[^)]+\)', '', 'IgnoreCase')
    $Result = [regex]::Replace($Result, '\s+', ' ').Trim()

    return $Result
}

# ------------------------------------------------------------------------------
# Version score (higher = newer)
# ------------------------------------------------------------------------------
function Get-VersionScore {
    param([string]$BaseName)

    $VMatch = [regex]::Match($BaseName, '\(v(\d+)\.(\d+)', 'IgnoreCase')
    if ($VMatch.Success) {
        return ([int]$VMatch.Groups[1].Value * 10000) + [int]$VMatch.Groups[2].Value
    }

    $RMatch = [regex]::Match($BaseName, '\(Rev\s+([^)]+)\)', 'IgnoreCase')
    if ($RMatch.Success) {
        $Rev = $RMatch.Groups[1].Value.Trim()
        $Num = 0
        if ([int]::TryParse($Rev, [ref]$Num)) { return $Num }
        if ($Rev.Length -eq 1 -and $Rev -match '[A-Za-z]') {
            return [int][char]($Rev.ToUpper()) - [int][char]'A' + 1
        }
    }

    return 0
}

# ------------------------------------------------------------------------------
# Region score (lower index = higher priority)
# ------------------------------------------------------------------------------
function Get-RegionScore {
    param([string]$Region)
    $Idx = [array]::IndexOf($RegionPriority, $Region)
    if ($Idx -lt 0) { return 9999 }
    return $Idx
}

# ------------------------------------------------------------------------------
# Select the preferred file from a group
# ------------------------------------------------------------------------------
function Get-PreferredFile {
    param([array]$FilesInGroup)

    if ($FilesInGroup.Count -le 1) { return $FilesInGroup[0] }

    $Sorted = $FilesInGroup | Sort-Object -Property @(
        @{ Expression = { Get-RegionScore $_.Region };   Ascending = $true  }
        @{ Expression = { Get-VersionScore $_.BaseName }; Ascending = $false }
    )
    return $Sorted[0]
}

# ==============================================================================
# MAIN
# ==============================================================================
try {
    Write-Host ""
    Write-Host "PS3 Region Deduplicator" -ForegroundColor Cyan
    Write-Host "=======================" -ForegroundColor Cyan
    Write-Host ""

    Write-Host "Enter the full path to your SOURCE folder (containing ZIP files):" -ForegroundColor Yellow
    $SourceFolder = (Read-Host "Source").Trim().Trim('"')
    if (-not (Test-Path -LiteralPath $SourceFolder -PathType Container)) {
        throw "Source folder not found: $SourceFolder"
    }

    Write-Host ""
    Write-Host "Enter the full path to your DESTINATION folder (where duplicates will be moved):" -ForegroundColor Yellow
    $DestinationFolder = (Read-Host "Destination").Trim().Trim('"')

    Write-Host ""
    Write-Host "Select mode:" -ForegroundColor Yellow
    Write-Host "  1 = Dry Run  (no files moved, report only)"
    Write-Host "  2 = Live     (actually move duplicate files)"
    Write-Host ""
    $ModeChoice = Read-Host "Enter 1 or 2"
    if ($ModeChoice -notmatch '^[12]$') { throw "Invalid choice. Enter 1 or 2." }
    $DryRun = ($ModeChoice -eq '1')

    Write-Host ""
    Write-Host ("  Source      : {0}" -f $SourceFolder)      -ForegroundColor Yellow
    Write-Host ("  Destination : {0}" -f $DestinationFolder) -ForegroundColor Yellow
    Write-Host ("  Mode        : {0}" -f $(if ($DryRun) { "Dry Run -- no files will be moved" } else { "LIVE -- files WILL be moved" })) `
        -ForegroundColor $(if ($DryRun) { "Green" } else { "Red" })
    Write-Host ""

    if (-not (Test-Path -LiteralPath $DestinationFolder)) {
        New-Item -Path $DestinationFolder -ItemType Directory -Force | Out-Null
    }

    # Collect files
    $AllFiles = @()
    foreach ($Ext in $FileExtensions) {
        $AllFiles += Get-ChildItem -LiteralPath $SourceFolder -File -Filter $Ext
    }

    if (-not $AllFiles -or $AllFiles.Count -eq 0) {
        throw "No matching files found in: $SourceFolder"
    }

    Write-Host ("Found {0} file(s). Analyzing..." -f $AllFiles.Count) -ForegroundColor Green
    Write-Host ""

    $FileObjects    = @()
    $CompanionCount = 0

    foreach ($File in $AllFiles) {
        $BaseName      = [System.IO.Path]::GetFileNameWithoutExtension($File.Name)
        $IsCompanion   = Test-IsCompanionFile -BaseName $BaseName
        $Region        = Get-RegionFromName   -BaseName $BaseName
        $NormalizedKey = Get-NormalizedKey    -BaseName $BaseName
        $FileSizeMB    = [math]::Round($File.Length / 1MB, 2)

        if (-not $IsCompanion -and $MinFullGameSizeMB -gt 0 -and $FileSizeMB -lt $MinFullGameSizeMB) {
            $IsCompanion = $true
        }

        if ($IsCompanion) { $CompanionCount++ }

        $FileObjects += [PSCustomObject]@{
            Name          = $File.Name
            FullName      = $File.FullName
            BaseName      = $BaseName
            Region        = $Region
            NormalizedKey = $NormalizedKey
            IsCompanion   = $IsCompanion
            SizeMB        = $FileSizeMB
        }
    }

    $GameFiles       = @($FileObjects | Where-Object { -not $_.IsCompanion })
    $Grouped         = $GameFiles | Group-Object -Property NormalizedKey
    $UniqueTitles    = ($Grouped | Where-Object { $_.Count -eq 1 }).Count
    $DuplicateGroups = ($Grouped | Where-Object { $_.Count -gt 1 }).Count

    Write-Host ("Companion/special files (always kept) : {0}" -f $CompanionCount)
    Write-Host ("Unique titles (no duplicates found)   : {0}" -f $UniqueTitles)
    Write-Host ("Titles with regional duplicates       : {0}" -f $DuplicateGroups)
    Write-Host ""

    $Actions = @()

    foreach ($Group in $Grouped) {
        $GroupFiles = @($Group.Group)

        if ($GroupFiles.Count -le 1) {
            $Actions += [PSCustomObject]@{
                NormalizedKey = $Group.Name
                FileName      = $GroupFiles[0].Name
                Region        = $GroupFiles[0].Region
                SizeMB        = $GroupFiles[0].SizeMB
                Action        = 'Keep'
                Reason        = 'Only copy'
                SourcePath    = $GroupFiles[0].FullName
                Destination   = ''
            }
            continue
        }

        $Winner = Get-PreferredFile -FilesInGroup $GroupFiles

        foreach ($Item in $GroupFiles) {
            $IsWinner = ($Item.FullName -eq $Winner.FullName)
            $Actions += [PSCustomObject]@{
                NormalizedKey = $Group.Name
                FileName      = $Item.Name
                Region        = $Item.Region
                SizeMB        = $Item.SizeMB
                Action        = if ($IsWinner) { 'Keep' } elseif ($DryRun) { 'WouldMove' } else { 'Move' }
                Reason        = if ($IsWinner) { 'Best region/version' } else { 'Lower priority duplicate' }
                SourcePath    = $Item.FullName
                Destination   = if ($IsWinner) { '' } else { Join-Path $DestinationFolder $Item.Name }
            }
        }
    }

    foreach ($C in ($FileObjects | Where-Object { $_.IsCompanion })) {
        $Actions += [PSCustomObject]@{
            NormalizedKey = $C.NormalizedKey
            FileName      = $C.Name
            Region        = $C.Region
            SizeMB        = $C.SizeMB
            Action        = 'Keep'
            Reason        = 'Companion -- excluded from dedup'
            SourcePath    = $C.FullName
            Destination   = ''
        }
    }

    $ToMove   = @($Actions | Where-Object { $_.Action -in @('WouldMove','Move') })
    $ToKeep   = @($Actions | Where-Object { $_.Action -eq 'Keep' })
    $MoveSize = [math]::Round(($ToMove | Measure-Object -Property SizeMB -Sum).Sum / 1024, 2)

    Write-Host "--------------------------------------" -ForegroundColor Cyan
    Write-Host ("Files to keep : {0}" -f $ToKeep.Count)
    Write-Host ("Files to move : {0}  (~{1} GB)" -f $ToMove.Count, $MoveSize)
    Write-Host "--------------------------------------" -ForegroundColor Cyan
    Write-Host ""

    # Always write CSV log
    $Timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $LogFile   = Join-Path $SourceFolder ("PS3_DuplicateLog_{0}.csv" -f $Timestamp)
    $Actions | Sort-Object NormalizedKey, Action |
        Export-Csv -LiteralPath $LogFile -NoTypeInformation -Encoding UTF8
    Write-Host ("Log saved: {0}" -f $LogFile) -ForegroundColor Green
    Write-Host ""

    # Move files (live mode only)
    $LockedFiles = @()

    if (-not $DryRun -and $ToMove.Count -gt 0) {
        $MoveIndex    = 0
        $MovedCount   = 0
        $SkippedCount = 0

        foreach ($Item in $ToMove) {
            $MoveIndex++
            Write-Host ("  Moving [{0}/{1}] {2}" -f $MoveIndex, $ToMove.Count, $Item.FileName) -ForegroundColor DarkGray

            if (-not (Test-Path -LiteralPath $Item.SourcePath)) {
                Write-Host ("    SKIPPED (not found): {0}" -f $Item.FileName) -ForegroundColor Yellow
                $SkippedCount++
                continue
            }

            $DestPath = $Item.Destination
            if (Test-Path -LiteralPath $DestPath) {
                $Base     = [System.IO.Path]::GetFileNameWithoutExtension($Item.FileName)
                $Ext      = [System.IO.Path]::GetExtension($Item.FileName)
                $DestPath = Join-Path $DestinationFolder ("{0}__DUP_{1}{2}" -f $Base, (Get-Date -Format "yyyyMMddHHmmssfff"), $Ext)
            }

            try {
                Move-Item -LiteralPath $Item.SourcePath -Destination $DestPath -Force
                $MovedCount++
            }
            catch {
                $ErrMsg = $_.Exception.Message
                Write-Host ("    LOCKED -- skipping: {0}" -f $Item.FileName) -ForegroundColor Yellow
                $LockedFiles += [PSCustomObject]@{
                    FileName   = $Item.FileName
                    Region     = $Item.Region
                    SizeMB     = $Item.SizeMB
                    SourcePath = $Item.SourcePath
                    Error      = $ErrMsg
                }
            }
        }

        Write-Host ""
        Write-Host "--------------------------------------" -ForegroundColor Cyan
        Write-Host ("Moved successfully : {0}" -f $MovedCount)
        if ($LockedFiles.Count -gt 0) {
            Write-Host ("Locked (in use)    : {0}" -f $LockedFiles.Count) -ForegroundColor Yellow
        }
        if ($SkippedCount -gt 0) {
            Write-Host ("Skipped (missing)  : {0}" -f $SkippedCount) -ForegroundColor DarkGray
        }
        Write-Host "--------------------------------------" -ForegroundColor Cyan

        if ($LockedFiles.Count -gt 0) {
            $LockedLog = Join-Path $SourceFolder ("PS3_LockedFiles_{0}.csv" -f $Timestamp)
            $LockedFiles | Export-Csv -LiteralPath $LockedLog -NoTypeInformation -Encoding UTF8
            Write-Host ""
            Write-Host ("Locked file log: {0}" -f $LockedLog) -ForegroundColor Yellow
            Write-Host "Close LaunchBox (or whatever has these files open) and re-run in Live mode to retry." -ForegroundColor Yellow
        }
    }

    if ($ToMove.Count -gt 0) {
        Write-Host ""
        Write-Host ("Preview -- {0} (first 30):" -f $(if ($DryRun) { "would move" } else { "moved" })) -ForegroundColor Yellow
        $ToMove | Select-Object -First 30 FileName, Region, SizeMB, Reason | Format-Table -AutoSize
    } else {
        Write-Host "No duplicates found -- nothing to move." -ForegroundColor Green
    }

    Write-Host ""
    Write-Host "Done." -ForegroundColor Green
    Write-Host ""
}
catch {
    Write-Host ""
    Write-Host ("Error: {0}" -f $_.Exception.Message) -ForegroundColor Red
    Write-Host ""
    Set-ExitCode 1
}
