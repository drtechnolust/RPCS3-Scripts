#Requires -Version 5.1
<#
.SYNOPSIS
    Identifies PS3 disc games that also exist as HDD copies and moves disc
    duplicates to a holding folder.

.DESCRIPTION
    Scans both dev_hdd0\disc and dev_hdd0\game folders, reads PARAM.SFO from
    each game to extract the Title ID, and matches disc games against HDD copies
    by Title ID. When a match is found, the HDD copy is preferred and the disc
    folder is moved to dev_hdd0\_DeDuplication_Disc.

    PARAM.SFO is parsed natively -- no external tools required.
    Disc-only games (no HDD copy) are kept in place.
    HDD-only games are logged but not touched.

    Non-game folders are skipped: DATA, INSTALL, cache, lock files.
    Game data folders (CATEGORY = GD, SD, AS) are also skipped.

    Run with DryRun = $true first to review the CSV log before moving anything.

.PARAMETER DryRun
    When $true (default), shows what would be moved without touching any folders.
    Set to $false in the CONFIG block to apply changes.

.EXAMPLE
    .\PS3-Disc-HDD-Deduplicator.ps1

.VERSION
    2.0.0 - Config block with pre-filled PS3 paths, DryRun default, MIT license.
             Replaced Unicode dashes with ASCII. Logic unchanged.
    1.0.0 - Initial release.

.LICENSE
    MIT License
    Copyright (c) Paul Mardis
#>

# ==============================================================================
# CONFIG -- Edit this block. Do not put paths anywhere else in this script.
# ==============================================================================
$PS3Root  = "D:\Arcade\System roms\Sony Playstation 3"
$DryRun   = $true    # Set to $false to actually move folders
# Derived paths (do not edit unless your folder structure differs)
$DiscPath  = Join-Path $PS3Root "dev_hdd0\disc"
$GamePath  = Join-Path $PS3Root "dev_hdd0\game"
$DedupeDir = Join-Path $PS3Root "dev_hdd0\_DeDuplication_Disc"
$LogDir    = Join-Path $PS3Root "_Logs"
# ==============================================================================

function Set-ExitCode {
    param([int]$Code)
    $global:LASTEXITCODE = $Code
}

# ------------------------------------------------------------------------------
# Parse PARAM.SFO -- returns hashtable of all fields or $null
# ------------------------------------------------------------------------------
function Read-ParamSfo {
    param([string]$SfoPath)

    try {
        $Bytes = [System.IO.File]::ReadAllBytes($SfoPath)

        if ($Bytes.Length -lt 20 -or
            $Bytes[0] -ne 0x00 -or $Bytes[1] -ne 0x50 -or
            $Bytes[2] -ne 0x53 -or $Bytes[3] -ne 0x46) { return $null }

        $KeyTableOffset  = [BitConverter]::ToInt32($Bytes, 8)
        $DataTableOffset = [BitConverter]::ToInt32($Bytes, 12)
        $EntryCount      = [BitConverter]::ToInt32($Bytes, 16)

        $Result = @{}

        for ($i = 0; $i -lt $EntryCount; $i++) {
            $EntryBase  = 20 + ($i * 16)
            $KeyOffset  = [BitConverter]::ToInt16($Bytes, $EntryBase)
            $DataType   = $Bytes[$EntryBase + 3]
            $DataLen    = [BitConverter]::ToInt32($Bytes, $EntryBase + 4)
            $DataOffset = [BitConverter]::ToInt32($Bytes, $EntryBase + 12)

            $KeyStart = $KeyTableOffset + $KeyOffset
            $KeyEnd   = $KeyStart
            while ($KeyEnd -lt $Bytes.Length -and $Bytes[$KeyEnd] -ne 0) { $KeyEnd++ }
            $Key = [System.Text.Encoding]::ASCII.GetString($Bytes, $KeyStart, $KeyEnd - $KeyStart)

            $ValStart = $DataTableOffset + $DataOffset

            if ($DataType -eq 2 -or $DataType -eq 0x04) {
                $Result[$Key] = [System.Text.Encoding]::UTF8.GetString($Bytes, $ValStart, $DataLen).TrimEnd([char]0)
            } elseif ($DataType -eq 4) {
                $Result[$Key] = [BitConverter]::ToInt32($Bytes, $ValStart)
            }
        }

        return $Result
    }
    catch { return $null }
}

# ------------------------------------------------------------------------------
# Region from Title ID prefix
# ------------------------------------------------------------------------------
function Get-RegionFromTitleId {
    param([string]$TitleId)
    if ($TitleId -match '^(BLUS|BCUS|NPUB)') { return 'USA' }
    if ($TitleId -match '^(BLES|BCES|NPEB)') { return 'Europe' }
    if ($TitleId -match '^(BLJS|BCJS|BLJM|BCJM|NPJB)') { return 'Japan' }
    if ($TitleId -match '^(BLAS|BCAS)') { return 'Asia' }
    if ($TitleId -match '^NP') { return 'PSN' }
    return 'Unknown'
}

# ------------------------------------------------------------------------------
# Find PARAM.SFO for a disc or HDD game folder
# ------------------------------------------------------------------------------
function Find-ParamSfo {
    param([string]$FolderPath, [string]$Type)

    if ($Type -eq 'disc') {
        $Sfo = Join-Path $FolderPath "PS3_GAME\PARAM.SFO"
        if (Test-Path -LiteralPath $Sfo) { return $Sfo }
        $Sfo = Join-Path $FolderPath "PARAM.SFO"
        if (Test-Path -LiteralPath $Sfo) { return $Sfo }
    } elseif ($Type -eq 'game') {
        $Sfo = Join-Path $FolderPath "PARAM.SFO"
        if (Test-Path -LiteralPath $Sfo) { return $Sfo }
    }
    return $null
}

# ==============================================================================
# PRE-FLIGHT
# ==============================================================================
foreach ($Path in @($DiscPath, $GamePath)) {
    if (-not (Test-Path -LiteralPath $Path)) {
        Write-Host ("ERROR: Folder not found: {0}" -f $Path) -ForegroundColor Red
        Set-ExitCode 1
        return
    }
}

if (-not $DryRun -and -not (Test-Path -LiteralPath $DedupeDir)) {
    New-Item -Path $DedupeDir -ItemType Directory -Force | Out-Null
}

if (-not (Test-Path -LiteralPath $LogDir)) {
    New-Item -Path $LogDir -ItemType Directory -Force | Out-Null
}

$Timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$LogFile   = Join-Path $LogDir ("PS3-Disc-HDD-Deduplicator_{0}.csv" -f $Timestamp)

# ==============================================================================
# BANNER
# ==============================================================================
Write-Host ""
Write-Host "PS3 Disc vs HDD Deduplicator" -ForegroundColor Cyan
Write-Host "============================" -ForegroundColor Cyan
Write-Host ("  Disc folder : {0}" -f $DiscPath)  -ForegroundColor Yellow
Write-Host ("  Game folder : {0}" -f $GamePath)  -ForegroundColor Yellow
Write-Host ("  Dedupe dir  : {0}" -f $DedupeDir) -ForegroundColor Yellow
Write-Host ("  Mode        : {0}" -f $(if ($DryRun) { "Dry Run -- no folders will be moved" } else { "LIVE -- folders WILL be moved" })) `
    -ForegroundColor $(if ($DryRun) { "Green" } else { "Red" })
Write-Host ""
if ($DryRun) {
    Write-Host "  [DRY RUN] No folders will be moved." -ForegroundColor Yellow
    Write-Host ""
}

# ==============================================================================
# SCAN HDD GAME FOLDER
# ==============================================================================
Write-Host "Scanning HDD game folder..." -ForegroundColor Green

$HddGames   = @{}
$SkippedHdd = 0

foreach ($Folder in (Get-ChildItem -LiteralPath $GamePath -Directory)) {
    if ($Folder.Name -match '[-_](DATA|INSTALL|cache|GAMEDATA)$' -or
        $Folder.Name -match '^\.' -or
        $Folder.Name -eq '$locks') {
        $SkippedHdd++
        continue
    }

    $SfoPath = Find-ParamSfo -FolderPath $Folder.FullName -Type 'game'
    if (-not $SfoPath) { $SkippedHdd++; continue }

    $Sfo = Read-ParamSfo -SfoPath $SfoPath
    if (-not $Sfo) { $SkippedHdd++; continue }

    $Category = $Sfo['CATEGORY']
    if ($Category -eq 'GD' -or $Category -eq 'SD' -or $Category -eq 'AS') {
        $SkippedHdd++
        continue
    }

    $TitleId = $Sfo['TITLE_ID']
    if (-not $TitleId) { $TitleId = $Folder.Name }

    if (-not $HddGames.ContainsKey($TitleId)) {
        $HddGames[$TitleId] = [PSCustomObject]@{
            TitleId    = $TitleId
            Title      = $Sfo['TITLE']
            Region     = Get-RegionFromTitleId -TitleId $TitleId
            Category   = $Category
            FolderName = $Folder.Name
            FullPath   = $Folder.FullName
        }
    }
}

Write-Host ("  Found {0} HDD games  ({1} folders skipped as data/cache/DLC)" -f $HddGames.Count, $SkippedHdd) -ForegroundColor Gray
Write-Host ""

# ==============================================================================
# SCAN DISC FOLDER
# ==============================================================================
Write-Host "Scanning disc folder..." -ForegroundColor Green

$DiscGames   = @{}
$SkippedDisc = 0

foreach ($Folder in (Get-ChildItem -LiteralPath $DiscPath -Directory)) {
    $SfoPath = Find-ParamSfo -FolderPath $Folder.FullName -Type 'disc'
    if (-not $SfoPath) { $SkippedDisc++; continue }

    $Sfo = Read-ParamSfo -SfoPath $SfoPath
    if (-not $Sfo) { $SkippedDisc++; continue }

    $TitleId = $Sfo['TITLE_ID']
    if (-not $TitleId) { $TitleId = $Folder.Name }

    if (-not $DiscGames.ContainsKey($TitleId)) {
        $DiscGames[$TitleId] = [PSCustomObject]@{
            TitleId    = $TitleId
            Title      = $Sfo['TITLE']
            Region     = Get-RegionFromTitleId -TitleId $TitleId
            Category   = $Sfo['CATEGORY']
            FolderName = $Folder.Name
            FullPath   = $Folder.FullName
        }
    }
}

Write-Host ("  Found {0} disc games  ({1} folders skipped)" -f $DiscGames.Count, $SkippedDisc) -ForegroundColor Gray
Write-Host ""

# ==============================================================================
# MATCH AND EVALUATE
# ==============================================================================
$Actions       = @()
$MatchCount    = 0
$DiscOnlyCount = 0
$HddOnlyCount  = 0

foreach ($TitleId in $DiscGames.Keys) {
    $Disc = $DiscGames[$TitleId]

    if ($HddGames.ContainsKey($TitleId)) {
        $MatchCount++
        $Hdd = $HddGames[$TitleId]
        $Actions += [PSCustomObject]@{
            TitleId      = $TitleId
            Title        = $Disc.Title
            DiscFolder   = $Disc.FolderName
            HddFolder    = $Hdd.FolderName
            Action       = if ($DryRun) { 'WouldMove' } else { 'Move' }
            Reason       = 'HDD copy exists -- disc is duplicate'
            DiscFullPath = $Disc.FullPath
            Destination  = Join-Path $DedupeDir $Disc.FolderName
        }
    } else {
        $DiscOnlyCount++
        $Actions += [PSCustomObject]@{
            TitleId      = $TitleId
            Title        = $Disc.Title
            DiscFolder   = $Disc.FolderName
            HddFolder    = ''
            Action       = 'Keep'
            Reason       = 'No HDD copy -- disc is only copy'
            DiscFullPath = $Disc.FullPath
            Destination  = ''
        }
    }
}

foreach ($TitleId in $HddGames.Keys) {
    if (-not $DiscGames.ContainsKey($TitleId)) {
        $HddOnlyCount++
        $Hdd = $HddGames[$TitleId]
        $Actions += [PSCustomObject]@{
            TitleId      = $TitleId
            Title        = $Hdd.Title
            DiscFolder   = ''
            HddFolder    = $Hdd.FolderName
            Action       = 'Keep'
            Reason       = 'HDD only -- no disc copy'
            DiscFullPath = ''
            Destination  = ''
        }
    }
}

$ToMove = @($Actions | Where-Object { $_.Action -in @('Move','WouldMove') })
$ToKeep = @($Actions | Where-Object { $_.Action -eq 'Keep' })

Write-Host "------------------------------------------" -ForegroundColor Cyan
Write-Host ("Disc + HDD duplicates found : {0}  (disc copies will be moved)" -f $MatchCount)
Write-Host ("Disc only (no HDD copy)     : {0}  (kept)" -f $DiscOnlyCount)
Write-Host ("HDD only  (no disc copy)    : {0}  (kept)" -f $HddOnlyCount)
Write-Host "------------------------------------------" -ForegroundColor Cyan
Write-Host ""

# ==============================================================================
# WRITE CSV LOG
# ==============================================================================
$Actions | Sort-Object Action, Title |
    Export-Csv -LiteralPath $LogFile -NoTypeInformation -Encoding UTF8
Write-Host ("Log saved: {0}" -f $LogFile) -ForegroundColor Green
Write-Host ""

# ==============================================================================
# MOVE FOLDERS (live mode only)
# ==============================================================================
if (-not $DryRun -and $ToMove.Count -gt 0) {

    if (-not (Test-Path -LiteralPath $DedupeDir)) {
        New-Item -Path $DedupeDir -ItemType Directory -Force | Out-Null
    }

    $MoveIndex  = 0
    $MoveErrors = 0

    foreach ($Item in $ToMove) {
        $MoveIndex++
        Write-Host ("  Moving [{0}/{1}] {2}" -f $MoveIndex, $ToMove.Count, $Item.DiscFolder) -ForegroundColor DarkGray

        if (Test-Path -LiteralPath $Item.DiscFullPath) {
            $DestPath = $Item.Destination
            if (Test-Path -LiteralPath $DestPath) {
                $DestPath = Join-Path $DedupeDir ("{0}__DUP_{1}" -f $Item.DiscFolder, (Get-Date -Format "yyyyMMddHHmmssfff"))
            }

            try {
                Move-Item -LiteralPath $Item.DiscFullPath -Destination $DestPath -Force
            }
            catch {
                Write-Host ("    ERROR: {0}" -f $_.Exception.Message) -ForegroundColor Red
                $MoveErrors++
            }
        } else {
            Write-Host ("    SKIPPED (not found): {0}" -f $Item.DiscFolder) -ForegroundColor Yellow
        }
    }

    Write-Host ""
    if ($MoveErrors -gt 0) {
        Write-Host ("Completed with {0} error(s)." -f $MoveErrors) -ForegroundColor Yellow
    } else {
        Write-Host "All folders moved successfully." -ForegroundColor Green
    }
}

# ==============================================================================
# PREVIEW
# ==============================================================================
if ($ToMove.Count -gt 0) {
    Write-Host ""
    Write-Host ("Preview -- {0} disc folders (first 30):" -f $(if ($DryRun) { "would move" } else { "moved" })) -ForegroundColor Yellow
    $ToMove | Select-Object -First 30 Title, TitleId, DiscFolder, HddFolder, Reason | Format-Table -AutoSize
} else {
    Write-Host "No disc/HDD duplicates found." -ForegroundColor Green
}

if ($DryRun) {
    Write-Host ""
    Write-Host "  DRY RUN complete. Set `$DryRun = `$false in the CONFIG block to apply." -ForegroundColor Yellow
}

Write-Host ""
Write-Host "Done." -ForegroundColor Green
Write-Host ""
Set-ExitCode 0
