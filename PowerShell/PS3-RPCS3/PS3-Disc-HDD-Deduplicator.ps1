#Requires -Version 5.1

<#
.SYNOPSIS
    PS3 Disc vs HDD Duplicate Organizer v1.0

.DESCRIPTION
    Scans dev_hdd0\disc and dev_hdd0\game folders, parses PARAM.SFO
    from each game to extract Title ID and game name, matches duplicates,
    and moves disc copies to _DeDuplication when an HDD copy exists.

    - Reads PARAM.SFO binary format natively — no external tools needed
    - Skips game data folders (CATEGORY = GD), cache, install folders
    - Prefers HDD copy, moves disc copy
    - Dry run by default
#>

# ── Prompts ───────────────────────────────────────────────────────────────────
$PS3Root  = Read-Host "Enter path to Sony Playstation 3 folder (contains dev_hdd0)"
$choice   = Read-Host "Dry run? (Y/N)"
$DryRun   = $choice -match '^y'

$DiscPath  = Join-Path $PS3Root "dev_hdd0\disc"
$GamePath  = Join-Path $PS3Root "dev_hdd0\game"
$DedupeDir = Join-Path $PS3Root "dev_hdd0\_DeDuplication_Disc"

# ==============================================================================
# FUNCTIONS
# ==============================================================================

function Read-ParamSfo {
    param([string]$SfoPath)

    try {
        $bytes = [System.IO.File]::ReadAllBytes($SfoPath)

        # Validate magic bytes: 0x00 P S F
        if ($bytes.Length -lt 20 -or
            $bytes[0] -ne 0x00 -or
            $bytes[1] -ne 0x50 -or
            $bytes[2] -ne 0x53 -or
            $bytes[3] -ne 0x46) {
            return $null
        }

        $keyTableOffset  = [BitConverter]::ToInt32($bytes, 8)
        $dataTableOffset = [BitConverter]::ToInt32($bytes, 12)
        $entryCount      = [BitConverter]::ToInt32($bytes, 16)

        $result = @{}

        for ($i = 0; $i -lt $entryCount; $i++) {
            $entryBase   = 20 + ($i * 16)
            $keyOffset   = [BitConverter]::ToInt16($bytes, $entryBase)
            $dataType    = $bytes[$entryBase + 3]
            $dataLen     = [BitConverter]::ToInt32($bytes, $entryBase + 4)
            $dataOffset  = [BitConverter]::ToInt32($bytes, $entryBase + 12)

            # Read key string
            $keyStart = $keyTableOffset + $keyOffset
            $keyEnd   = $keyStart
            while ($keyEnd -lt $bytes.Length -and $bytes[$keyEnd] -ne 0) { $keyEnd++ }
            $key = [System.Text.Encoding]::ASCII.GetString($bytes, $keyStart, $keyEnd - $keyStart)

            # Read value
            $valStart = $dataTableOffset + $dataOffset
            if ($dataType -eq 2 -or $dataType -eq 0x04) {
                # String (utf-8)
                $rawVal = [System.Text.Encoding]::UTF8.GetString($bytes, $valStart, $dataLen).TrimEnd([char]0)
                $result[$key] = $rawVal
            } elseif ($dataType -eq 4) {
                # Integer
                $result[$key] = [BitConverter]::ToInt32($bytes, $valStart)
            }
        }

        return $result
    }
    catch {
        return $null
    }
}

function Get-RegionFromTitleId {
    param([string]$TitleId)
    if ($TitleId -match '^(BLUS|BCUS|NPUB)') { return 'USA' }
    if ($TitleId -match '^(BLES|BCES|NPEB)') { return 'Europe' }
    if ($TitleId -match '^(BLJS|BCJS|BLJM|BCJM|NPJB)') { return 'Japan' }
    if ($TitleId -match '^(BLAS|BCAS|NPUB)') { return 'Asia' }
    if ($TitleId -match '^NP') { return 'PSN' }
    return 'Unknown'
}

function Get-RegionScore {
    param([string]$Region)
    switch ($Region) {
        'USA'     { return 0 }
        'Europe'  { return 1 }
        'Asia'    { return 2 }
        'Japan'   { return 3 }
        'PSN'     { return 4 }
        default   { return 99 }
    }
}

function Find-ParamSfo {
    param([string]$FolderPath, [string]$Type)

    if ($Type -eq 'disc') {
        # disc\GameName\PS3_GAME\PARAM.SFO
        $sfo = Join-Path $FolderPath "PS3_GAME\PARAM.SFO"
        if (Test-Path $sfo) { return $sfo }
        # Some disc dumps have it at root
        $sfo = Join-Path $FolderPath "PARAM.SFO"
        if (Test-Path $sfo) { return $sfo }
    }
    elseif ($Type -eq 'game') {
        # game\TITLEID\PARAM.SFO
        $sfo = Join-Path $FolderPath "PARAM.SFO"
        if (Test-Path $sfo) { return $sfo }
    }
    return $null
}

# ==============================================================================
# MAIN
# ==============================================================================

try {
    Write-Host ""
    Write-Host "PS3 Disc vs HDD Duplicate Organizer v1.0" -ForegroundColor Cyan
    Write-Host "─────────────────────────────────────────" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "Disc folder : $DiscPath"  -ForegroundColor Yellow
    Write-Host "Game folder : $GamePath"  -ForegroundColor Yellow
    Write-Host "Dedupe dir  : $DedupeDir" -ForegroundColor Yellow
    Write-Host "Mode        : $(if ($DryRun) { 'Dry Run — no files will be moved' } else { 'LIVE — folders WILL be moved' })" -ForegroundColor $(if ($DryRun) { 'Green' } else { 'Red' })
    Write-Host ""

    foreach ($path in @($DiscPath, $GamePath)) {
        if (-not (Test-Path $path)) {
            throw "Folder not found: $path"
        }
    }

    if (-not (Test-Path $DedupeDir)) {
        New-Item -Path $DedupeDir -ItemType Directory -Force | Out-Null
    }

    # ── Scan HDD game folder ──────────────────────────────────────────────────
    Write-Host "Scanning HDD game folder..." -ForegroundColor Green

    $hddGames = @{}
    $hddFolders = Get-ChildItem -Path $GamePath -Directory

    $skippedHdd = 0

    foreach ($folder in $hddFolders) {
        # Skip obviously non-game folders
        if ($folder.Name -match '[-_](DATA|INSTALL|cache|GAMEDATA)$' -or
            $folder.Name -match '^\.' -or
            $folder.Name -eq '$locks' -or
            $folder.Name -eq '.locks') {
            $skippedHdd++
            continue
        }

        $sfoPath = Find-ParamSfo -FolderPath $folder.FullName -Type 'game'
        if (-not $sfoPath) {
            $skippedHdd++
            continue
        }

        $sfo = Read-ParamSfo -SfoPath $sfoPath
        if (-not $sfo) {
            $skippedHdd++
            continue
        }

        # Skip game data (CATEGORY = GD) and licence/system entries
        $category = $sfo['CATEGORY']
        if ($category -eq 'GD' -or $category -eq 'SD' -or $category -eq 'AS') {
            $skippedHdd++
            continue
        }

        $titleId = $sfo['TITLE_ID']
        if (-not $titleId) { $titleId = $folder.Name }

        $title  = $sfo['TITLE']
        $region = Get-RegionFromTitleId -TitleId $titleId

        if (-not $hddGames.ContainsKey($titleId)) {
            $hddGames[$titleId] = [PSCustomObject]@{
                TitleId    = $titleId
                Title      = $title
                Region     = $region
                Category   = $category
                FolderName = $folder.Name
                FullPath   = $folder.FullName
                Source     = 'HDD'
            }
        }
    }

    Write-Host "  Found $($hddGames.Count) HDD games  ($skippedHdd folders skipped as data/cache/DLC)" -ForegroundColor Gray
    Write-Host ""

    # ── Scan disc folder ──────────────────────────────────────────────────────
    Write-Host "Scanning disc folder..." -ForegroundColor Green

    $discGames  = @{}
    $discFolders = Get-ChildItem -Path $DiscPath -Directory
    $skippedDisc = 0

    foreach ($folder in $discFolders) {
        $sfoPath = Find-ParamSfo -FolderPath $folder.FullName -Type 'disc'
        if (-not $sfoPath) {
            $skippedDisc++
            continue
        }

        $sfo = Read-ParamSfo -SfoPath $sfoPath
        if (-not $sfo) {
            $skippedDisc++
            continue
        }

        $category = $sfo['CATEGORY']
        $titleId  = $sfo['TITLE_ID']
        if (-not $titleId) { $titleId = $folder.Name }

        $title  = $sfo['TITLE']
        $region = Get-RegionFromTitleId -TitleId $titleId

        if (-not $discGames.ContainsKey($titleId)) {
            $discGames[$titleId] = [PSCustomObject]@{
                TitleId    = $titleId
                Title      = $title
                Region     = $region
                Category   = $category
                FolderName = $folder.Name
                FullPath   = $folder.FullName
                Source     = 'Disc'
            }
        }
    }

    Write-Host "  Found $($discGames.Count) disc games  ($skippedDisc folders skipped)" -ForegroundColor Gray
    Write-Host ""

    # ── Match and evaluate ────────────────────────────────────────────────────
    $actions        = @()
    $matchCount     = 0
    $discOnlyCount  = 0
    $hddOnlyCount   = 0

    # Check every disc game
    foreach ($titleId in $discGames.Keys) {
        $disc = $discGames[$titleId]

        if ($hddGames.ContainsKey($titleId)) {
            # Duplicate found — HDD wins
            $matchCount++
            $hdd = $hddGames[$titleId]

            $actions += [PSCustomObject]@{
                TitleId      = $titleId
                Title        = $disc.Title
                DiscFolder   = $disc.FolderName
                HddFolder    = $hdd.FolderName
                Action       = if ($DryRun) { 'WouldMove' } else { 'Move' }
                Reason       = 'HDD copy exists — disc is duplicate'
                DiscFullPath = $disc.FullPath
                Destination  = Join-Path $DedupeDir $disc.FolderName
            }
        }
        else {
            $discOnlyCount++
            $actions += [PSCustomObject]@{
                TitleId      = $titleId
                Title        = $disc.Title
                DiscFolder   = $disc.FolderName
                HddFolder    = ''
                Action       = 'Keep'
                Reason       = 'No HDD copy — disc is only copy'
                DiscFullPath = $disc.FullPath
                Destination  = ''
            }
        }
    }

    # HDD-only games (no disc copy) — just log them
    foreach ($titleId in $hddGames.Keys) {
        if (-not $discGames.ContainsKey($titleId)) {
            $hddOnlyCount++
            $hdd = $hddGames[$titleId]
            $actions += [PSCustomObject]@{
                TitleId      = $titleId
                Title        = $hdd.Title
                DiscFolder   = ''
                HddFolder    = $hdd.FolderName
                Action       = 'Keep'
                Reason       = 'HDD only — no disc copy'
                DiscFullPath = ''
                Destination  = ''
            }
        }
    }

    # ── Summary ───────────────────────────────────────────────────────────────
    $toMove = @($actions | Where-Object { $_.Action -in @('Move','WouldMove') })
    $toKeep = @($actions | Where-Object { $_.Action -eq 'Keep' })

    Write-Host "──────────────────────────────────────────" -ForegroundColor Cyan
    Write-Host "Disc + HDD duplicates found : $matchCount  (disc copies will be moved)"
    Write-Host "Disc only (no HDD copy)     : $discOnlyCount  (kept)"
    Write-Host "HDD only  (no disc copy)    : $hddOnlyCount  (kept)"
    Write-Host "──────────────────────────────────────────" -ForegroundColor Cyan
    Write-Host ""

    # ── Save CSV ──────────────────────────────────────────────────────────────
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $logFile   = Join-Path $PS3Root "PS3_DiscHDD_Dedup_$timestamp.csv"
    $actions | Sort-Object Action, Title |
        Export-Csv -Path $logFile -NoTypeInformation -Encoding UTF8
    Write-Host "Log saved : $logFile" -ForegroundColor Green
    Write-Host ""

    # ── Move folders ──────────────────────────────────────────────────────────
    if (-not $DryRun -and $toMove.Count -gt 0) {
        $moveIndex  = 0
        $moveErrors = 0

        foreach ($item in $toMove) {
            $moveIndex++
            Write-Host "  Moving [$moveIndex/$($toMove.Count)] $($item.DiscFolder)" -ForegroundColor DarkGray

            if (Test-Path $item.DiscFullPath) {
                $destPath = $item.Destination

                if (Test-Path $destPath) {
                    $destPath = Join-Path $DedupeDir ("$($item.DiscFolder)__DUP_$(Get-Date -Format 'yyyyMMddHHmmssfff')")
                }

                try {
                    Move-Item -Path $item.DiscFullPath -Destination $destPath -Force
                }
                catch {
                    Write-Host "    ERROR: $($_.Exception.Message)" -ForegroundColor Red
                    $moveErrors++
                }
            }
            else {
                Write-Host "    SKIPPED (not found): $($item.DiscFolder)" -ForegroundColor Yellow
            }
        }

        Write-Host ""
        if ($moveErrors -gt 0) {
            Write-Host "Completed with $moveErrors error(s)." -ForegroundColor Yellow
        }
        else {
            Write-Host "All folders moved successfully." -ForegroundColor Green
        }
    }

    # ── Preview ───────────────────────────────────────────────────────────────
    if ($toMove.Count -gt 0) {
        Write-Host ""
        Write-Host "Preview — $(if ($DryRun) { 'would move' } else { 'moved' }) disc folders (first 30):" -ForegroundColor Yellow
        $toMove | Select-Object -First 30 Title, TitleId, DiscFolder, HddFolder, Reason |
            Format-Table -AutoSize
    }
    else {
        Write-Host "No disc/HDD duplicates found." -ForegroundColor Green
    }

    Write-Host ""
    Write-Host "Done." -ForegroundColor Green
    Write-Host ""
}
catch {
    Write-Host ""
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host ""
}