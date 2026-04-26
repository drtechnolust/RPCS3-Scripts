<#
===============================================================================
  SCRIPT   : Create-RPCS3-Shortcuts.ps1
  AUTHOR   : Paul Mardis
  CREATED  : 2025
  VERSION  : 1.0
  GITHUB   : https://github.com/drtechnolust/RPCS3-Scripts

===============================================================================
  COPYRIGHT & LICENSE
===============================================================================
  Copyright (c) 2025 Paul Mardis. All rights reserved.

  This script is the original work of Paul Mardis and is provided for
  personal, non-commercial use only.

  You MAY:
    - Use this script for your own personal PS3/RPCS3 setup
    - Share it with others provided this full header remains intact and
      credit is clearly given to the original author: Paul Mardis

  You MAY NOT:
    - Remove or alter this copyright notice or author attribution
    - Redistribute this script as your own work
    - Include this script in paid tools, packages, or products without
      explicit written permission from Paul Mardis
    - Claim authorship or creation of this script

  If you share or repost this script anywhere (GitHub, Reddit, forums,
  YouTube descriptions, Discord, etc.) you MUST credit:
    Paul Mardis — https://github.com/drtechnolust

===============================================================================
  DESCRIPTION
===============================================================================
  Scans a PS3 game library folder (dev_hdd0\game) and automatically creates
or updates Windows shortcuts (.lnk files) for each valid PS3 game, pointing
to RPCS3 with the --no-gui flag for seamless direct launch.

  Features:
    - Parses PARAM.SFO binary metadata to extract proper game titles
    - Derives region tags (US/EU/JP/AS/KR) from PS3 Title ID prefixes
    - Skips non-game folders (DATA, INSTALL, cache, hidden folders)
    - Logs all missing PARAM.SFO and EBOOT.BIN files for easy debugging
    - Dry run mode to safely preview all changes before committing
    - Automatically de-duplicates shortcut names
  Designed for use with LaunchBox + RPCS3 on Windows.

===============================================================================
#>

param(
    # Root folder with your PS3 Title ID folders (dev_hdd0\game)
    [string]$RootPath     = "D:\Arcade\System roms\Sony Playstation 3\dev_hdd0\game",

    # Folder where shortcuts will be created/updated
    [string]$ShortcutPath = "D:\Arcade\System roms\Sony Playstation 3\games\shortcuts",

    # Path to rpcs3.exe
    [string]$Rpcs3ExePath = "C:\Arcade\LaunchBox\Emulators\RPCS3\rpcs3.exe"
)

# --------- Ask if this should be a dry run ----------
$answer = Read-Host "Run as DRY RUN first so no changes are made? (Y/N, default Y)"
$WhatIf = $true
if ($answer -and $answer.Trim().ToUpper().StartsWith("N")) {
    $WhatIf = $false
}
if ($WhatIf) {
    Write-Host ">>> DRY RUN ENABLED - no shortcuts will actually be created/updated." -ForegroundColor Yellow
} else {
    Write-Host ">>> LIVE RUN - shortcuts will be created/updated." -ForegroundColor Green
}

if (-not (Test-Path $Rpcs3ExePath)) {
    Write-Host "ERROR: rpcs3.exe not found at path:`n  $Rpcs3ExePath" -ForegroundColor Red
    return
}

# --------- Helper: parse PARAM.SFO (TITLE, TITLE_ID, CONTENT_ID) ----------
function Get-SfoInfo {
    param(
        [Parameter(Mandatory)]
        [string]$SfoPath
    )

    if (-not (Test-Path $SfoPath)) { return $null }

    $bytes = [System.IO.File]::ReadAllBytes($SfoPath)
    if ($bytes.Length -lt 0x14) { return $null }

    # Magic 00 50 53 46 ("\0PSF")
    if ($bytes[0] -ne 0x00 -or $bytes[1] -ne 0x50 -or $bytes[2] -ne 0x53 -or $bytes[3] -ne 0x46) {
        return $null
    }

    $keyTableOffset  = [BitConverter]::ToInt32($bytes, 8)
    $dataTableOffset = [BitConverter]::ToInt32($bytes, 12)
    $indexEntries    = [BitConverter]::ToInt32($bytes, 16)

    $offset = 0x14
    $result = @{}

    for ($i = 0; $i -lt $indexEntries; $i++) {
        $keyOffset   = [BitConverter]::ToUInt16($bytes, $offset); $offset += 2
        $dataFmt     = [BitConverter]::ToUInt16($bytes, $offset); $offset += 2
        $dataLen     = [BitConverter]::ToInt32($bytes, $offset);  $offset += 4
        $dataMaxLen  = [BitConverter]::ToInt32($bytes, $offset);  $offset += 4
        $dataOffset  = [BitConverter]::ToInt32($bytes, $offset);  $offset += 4

        $keyStart = $keyTableOffset + $keyOffset
        $k = $keyStart
        while ($k -lt $bytes.Length -and $bytes[$k] -ne 0) { $k++ }
        $keyName = [System.Text.Encoding]::ASCII.GetString($bytes, $keyStart, $k - $keyStart)

        if ($dataFmt -eq 0x0204) {
            $dataStart = $dataTableOffset + $dataOffset
            if ($dataStart -lt 0 -or $dataStart -ge $bytes.Length) { continue }

            $len = [Math]::Min($dataLen, $bytes.Length - $dataStart)
            $valBytes = New-Object byte[] $len
            [Array]::Copy($bytes, $dataStart, $valBytes, 0, $len)

            $trimLen = $len
            for ($j = $len - 1; $j -ge 0; $j--) {
                if ($valBytes[$j] -eq 0) { $trimLen-- } else { break }
            }

            if ($trimLen -gt 0) {
                $value = [System.Text.Encoding]::UTF8.GetString($valBytes, 0, $trimLen)
                $result[$keyName] = $value
            }
        }
    }

    if (-not $result.ContainsKey("TITLE")) { return $null }

    [PSCustomObject]@{
        Title     = $result["TITLE"]
        TitleId   = if ($result.ContainsKey("TITLE_ID"))   { $result["TITLE_ID"]   } else { $null }
        ContentId = if ($result.ContainsKey("CONTENT_ID")) { $result["CONTENT_ID"] } else { $null }
    }
}

# --------- Helper: get region from Title ID prefix ----------
# PS3 Title IDs encode region in the first 4 letters:
#   BCUS / BLUS / NPUB / NPUA = US
#   BCES / BLES / NPEB / NPEA = EU
#   BCJS / BCAS / BLJM / NPJB / NPHA = JP/AS
#   NPKB = KR
function Get-RegionTag {
    param([string]$TitleId)

    if (-not $TitleId -or $TitleId.Length -lt 4) { return $null }

    $prefix = $TitleId.Substring(0, 4).ToUpper()

    switch -Wildcard ($prefix) {
        "BCUS" { return "US" }
        "BLUS" { return "US" }
        "NPUB" { return "US" }
        "NPUA" { return "US" }
        "BCES" { return "EU" }
        "BLES" { return "EU" }
        "NPEB" { return "EU" }
        "NPEA" { return "EU" }
        "BCJS" { return "JP" }
        "BCAS" { return "AS" }
        "BLJM" { return "JP" }
        "NPJB" { return "JP" }
        "NPHA" { return "AS" }
        "NPKB" { return "KR" }
        default { return $null }
    }
}

# --------- Skip folders that are clearly not games ----------
# DATA, INSTALL, cache, and lock folders are not launchable games
function Test-IsSkippableFolder {
    param([string]$FolderName)

    if ($FolderName -match "DATA$")    { return $true }
    if ($FolderName -match "INSTALL$") { return $true }
    if ($FolderName -match "_cache$")  { return $true }
    if ($FolderName.StartsWith("."))   { return $true }
    if ($FolderName.StartsWith("$"))   { return $true }
    return $false
}

# --------- Prep & tracking ----------
if (-not (Test-Path $ShortcutPath)) {
    if ($WhatIf) {
        Write-Host "[DRY RUN] Would create shortcut folder: $ShortcutPath"
    } else {
        New-Item -ItemType Directory -Path $ShortcutPath -Force | Out-Null
    }
}

$created       = @()
$updated       = @()
$skipped       = @()
$missingSfo    = @()
$missingEboot  = @()
$parseErrors   = @()

$wsh = New-Object -ComObject WScript.Shell
$usedNames = New-Object System.Collections.Generic.HashSet[string]

$extraBad = [char[]]@([char]0xAE, [char]0x2122, [char]0xA9)  # ®, ™, ©

Write-Host ""
Write-Host "Scanning PS3 games in: $RootPath"
Write-Host ""

Get-ChildItem -Path $RootPath -Directory | ForEach-Object {
    $gameDir  = $_
    $gamePath = $gameDir.FullName

    # Skip non-game folders (DATA, INSTALL, cache, hidden)
    if (Test-IsSkippableFolder -FolderName $gameDir.Name) {
        $skipped += $gamePath
        Write-Host "[$($gameDir.Name)] Skipped (non-game folder)" -ForegroundColor DarkGray
        return
    }

    # PS3: PARAM.SFO is in the game root (not in a subfolder like PS4's sce_sys)
    $sfoPath = Join-Path $gamePath "PARAM.SFO"
    if (-not (Test-Path $sfoPath)) {
        $missingSfo += $gamePath
        Write-Host "[$($gameDir.Name)] No PARAM.SFO found – SKIP" -ForegroundColor DarkYellow
        return
    }

    $info = Get-SfoInfo -SfoPath $sfoPath
    if (-not $info) {
        $parseErrors += $gamePath
        Write-Host "[$($gameDir.Name)] Could not read TITLE from PARAM.SFO – SKIP" -ForegroundColor Red
        return
    }

    # PS3: EBOOT.BIN is in USRDIR subfolder
    $ebootPath = Join-Path $gamePath "USRDIR\EBOOT.BIN"
    if (-not (Test-Path $ebootPath)) {
        $missingEboot += $gamePath
        Write-Host "[$($gameDir.Name)] No USRDIR\EBOOT.BIN found – SKIP" -ForegroundColor DarkYellow
        return
    }

    # Clean title (remove invalid filename chars and symbols)
    $title = $info.Title
    $invalid = [System.IO.Path]::GetInvalidFileNameChars() + $extraBad
    foreach ($c in $invalid) {
        $s = [string]$c
        if ($s.Length -gt 0) {
            $title = $title.Replace($s, "")
        }
    }
    $title = $title.Trim()
    if (-not $title) {
        $parseErrors += $gamePath
        Write-Host "[$($gameDir.Name)] Title empty after cleaning – SKIP" -ForegroundColor Red
        return
    }

    # Region tag derived from Title ID prefix (e.g. BLES = EU, BLUS = US)
    $regionTag = Get-RegionTag -TitleId $info.TitleId
    $baseName  = if ($regionTag) { "$title [$regionTag]" } else { $title }

    # De-duplicate shortcut names
    $shortcutFull = $null
    $shortcutName = $null
    $existing     = $false

    $candidateBase = $baseName
    $counter = 2
    while ($true) {
        $shortcutName = "$candidateBase.lnk"
        $candidatePath = Join-Path $ShortcutPath $shortcutName

        if (Test-Path $candidatePath) {
            $shortcutFull = $candidatePath
            $existing     = $true
            break
        }

        if (-not $usedNames.Contains($candidateBase.ToLower())) {
            $shortcutFull = $candidatePath
            $usedNames.Add($candidateBase.ToLower()) | Out-Null
            break
        }

        $candidateBase = "$baseName ($counter)"
        $counter++
    }

    # RPCS3 CLI: --no-gui skips the library window and launches the game directly
    $arguments = "--no-gui `"$ebootPath`""

    if ($WhatIf) {
        if ($existing) {
            Write-Host "[DRY RUN] Would UPDATE: $shortcutName" -ForegroundColor Cyan
        } else {
            Write-Host "[DRY RUN] Would CREATE: $shortcutName" -ForegroundColor Green
        }
        Write-Host "          Target   : $Rpcs3ExePath"
        Write-Host "          Arguments: $arguments"
        return
    }

    # Create or update shortcut
    $sc = $wsh.CreateShortcut($shortcutFull)
    $sc.TargetPath       = $Rpcs3ExePath
    $sc.Arguments        = $arguments
    $sc.WorkingDirectory = [System.IO.Path]::GetDirectoryName($Rpcs3ExePath)
    $sc.IconLocation     = $Rpcs3ExePath
    $sc.WindowStyle      = 1
    $sc.Save()

    if ($existing) {
        $updated += $shortcutFull
        Write-Host "Updated : $shortcutName" -ForegroundColor Cyan
    } else {
        $created += $shortcutFull
        Write-Host "Created : $shortcutName" -ForegroundColor Green
    }
}

Write-Host ""
Write-Host "---------- SUMMARY ----------"
Write-Host "Shortcuts CREATED      : $($created.Count)"
Write-Host "Shortcuts UPDATED      : $($updated.Count)"
Write-Host "Skipped (non-game)     : $($skipped.Count)"
Write-Host "Missing PARAM.SFO      : $($missingSfo.Count)"
Write-Host "Missing EBOOT.BIN      : $($missingEboot.Count)"
Write-Host "Parse errors           : $($parseErrors.Count)"
Write-Host "-----------------------------"

[PSCustomObject]@{
    Created      = $created
    Updated      = $updated
    Skipped      = $skipped
    MissingSfo   = $missingSfo
    MissingEboot = $missingEboot
    ParseErrors  = $parseErrors
}
