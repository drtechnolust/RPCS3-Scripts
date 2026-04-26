<#
===============================================================================
  SCRIPT   : Create-RPCS3-Shortcuts-HD-DISC.ps1
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
  Scans BOTH PS3 game library folders in a single run and automatically creates
  or updates Windows shortcuts (.lnk files) for each valid game, pointing to
  RPCS3 with the --no-gui flag for seamless direct launch.

  Handles two distinct PS3 library structures:
    HD/PSN Games  (dev_hdd0\game) — Title ID named folders, PARAM.SFO in root
    Disc Games    (dev_hdd0\disc) — Human-readable folders, PARAM.SFO in PS3_GAME

  Features:
    - Two-pass scan: processes HD/PSN games then Disc games in one run
    - Parses PARAM.SFO binary metadata for accurate game titles
    - Derives region tags from Title ID prefixes (HD) and folder names (Disc)
    - Shared de-duplication tracker prevents duplicate shortcuts across both sources
    - Unicode-safe shortcut writing via temp-file workaround for WScript.Shell
    - Skips non-game folders (DATA, INSTALL, cache, hidden)
    - Dry run mode to safely preview all changes before committing
    - Detailed per-library issue tracking in summary report
  Designed for use with LaunchBox + RPCS3 on Windows.

===============================================================================
#>

param(
    # HD/PSN games folder (Title ID named folders)
    [string]$HdGamePath   = "D:\Arcade\System roms\Sony Playstation 3\dev_hdd0\game",

    # Disc games folder (human-readable named folders)
    [string]$DiscGamePath = "D:\Arcade\System roms\Sony Playstation 3\dev_hdd0\disc",

    # Folder where all shortcuts will be created/updated
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
        $keyOffset  = [BitConverter]::ToUInt16($bytes, $offset); $offset += 2
        $dataFmt    = [BitConverter]::ToUInt16($bytes, $offset); $offset += 2
        $dataLen    = [BitConverter]::ToInt32($bytes, $offset);  $offset += 4
        $dataMaxLen = [BitConverter]::ToInt32($bytes, $offset);  $offset += 4
        $dataOffset = [BitConverter]::ToInt32($bytes, $offset);  $offset += 4

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

# --------- Helper: region from Title ID prefix (HD/PSN games) ----------
# BCUS/BLUS/NPUB/NPUA = US
# BCES/BLES/NPEB/NPEA = EU
# BCJS/BLJM/NPJB      = JP
# BCAS/NPHA           = AS
# NPKB                = KR
function Get-RegionFromTitleId {
    param([string]$TitleId)

    if (-not $TitleId -or $TitleId.Length -lt 4) { return $null }
    $prefix = $TitleId.Substring(0, 4).ToUpper()

    switch ($prefix) {
        "BCUS" { return "US" }
        "BLUS" { return "US" }
        "NPUB" { return "US" }
        "NPUA" { return "US" }
        "BCES" { return "EU" }
        "BLES" { return "EU" }
        "NPEB" { return "EU" }
        "NPEA" { return "EU" }
        "BCJS" { return "JP" }
        "BLJM" { return "JP" }
        "NPJB" { return "JP" }
        "BCAS" { return "AS" }
        "NPHA" { return "AS" }
        "NPKB" { return "KR" }
        default { return $null }
    }
}

# --------- Helper: region from folder name suffix (disc games) ----------
# Folder names end with _(USA), _(Japan), _(Europe), etc.
function Get-RegionFromFolderName {
    param([string]$FolderName)

    if ($FolderName -match '\(USA\)')     { return "US" }
    if ($FolderName -match '\(Europe\)')  { return "EU" }
    if ($FolderName -match '\(Japan\)')   { return "JP" }
    if ($FolderName -match '\(Asia\)')    { return "AS" }
    if ($FolderName -match '\(Korea\)')   { return "KR" }
    if ($FolderName -match '\(China\)')   { return "CN" }
    return $null
}

# --------- Helper: convert disc folder name to readable title ----------
# e.g. "LEGO_Indiana_Jones_The_Original_Adventures_(USA)" -> "LEGO Indiana Jones The Original Adventures"
function Get-TitleFromFolderName {
    param([string]$FolderName)

    # Strip trailing region tag e.g. _(USA), _(Japan)
    $title = $FolderName -replace '_\([^)]+\)$', ''

    # Replace underscores with spaces
    $title = $title -replace '_', ' '

    return $title.Trim()
}

# --------- Helper: skip non-game folders (HD/PSN library only) ----------
function Test-IsSkippableFolder {
    param([string]$FolderName)

    if ($FolderName -match "DATA$")    { return $true }
    if ($FolderName -match "INSTALL$") { return $true }
    if ($FolderName -match "_cache$")  { return $true }
    if ($FolderName.StartsWith("."))   { return $true }
    if ($FolderName.StartsWith("$"))   { return $true }
    return $false
}

# --------- Helper: strip invalid filename characters ----------
function Get-CleanTitle {
    param([string]$Title)

    $extraBad = [char[]]@([char]0xAE, [char]0x2122, [char]0xA9)  # ®, ™, ©
    $invalid  = [System.IO.Path]::GetInvalidFileNameChars() + $extraBad

    foreach ($c in $invalid) {
        $s = [string]$c
        if ($s.Length -gt 0) { $Title = $Title.Replace($s, "") }
    }
    return $Title.Trim()
}

# --------- Helper: resolve a unique shortcut path ----------
function Resolve-ShortcutPath {
    param(
        [string]$BaseName,
        [string]$ShortcutFolder,
        [System.Collections.Generic.HashSet[string]]$UsedNames
    )

    $candidate = $BaseName
    $counter   = 2

    while ($true) {
        $lnkName = "$candidate.lnk"
        $lnkPath = Join-Path $ShortcutFolder $lnkName

        if (Test-Path $lnkPath) {
            # Shortcut already exists on disk — update it
            return [PSCustomObject]@{ Path = $lnkPath; Name = $lnkName; Existing = $true }
        }

        if (-not $UsedNames.Contains($candidate.ToLower())) {
            $UsedNames.Add($candidate.ToLower()) | Out-Null
            return [PSCustomObject]@{ Path = $lnkPath; Name = $lnkName; Existing = $false }
        }

        $candidate = "$BaseName ($counter)"
        $counter++
    }
}

# --------- Helper: write shortcut or print dry-run preview ----------
# WScript.Shell.CreateShortcut() fails to SAVE when the .lnk path contains Unicode
# characters (Japanese, etc.) because it uses ANSI internally for the file path.
# Workaround: save to a safe ASCII temp path first, then use .NET to move/rename
# the file to the final Unicode path. The .lnk content itself is fine either way.
function Write-Shortcut {
    param(
        [string]$ShortcutPath,
        [string]$ShortcutName,
        [bool]$Existing,
        [string]$TargetExe,
        [string]$Arguments,
        [bool]$WhatIf,
        [System.Collections.ArrayList]$Created,
        [System.Collections.ArrayList]$Updated
    )

    if ($WhatIf) {
        $action = if ($Existing) { "Would UPDATE" } else { "Would CREATE" }
        $color  = if ($Existing) { "Cyan" } else { "Green" }
        Write-Host "[DRY RUN] $action`: $ShortcutName" -ForegroundColor $color
        Write-Host "          Target   : $TargetExe"
        Write-Host "          Arguments: $Arguments"
        return
    }

    try {
        # Save to a unique ASCII temp path to avoid WScript.Shell Unicode bug
        $tempPath = [System.IO.Path]::Combine(
            [System.IO.Path]::GetTempPath(),
            "_rpcs3_tmp_$([System.Guid]::NewGuid().ToString('N')).lnk"
        )

        $wsh = New-Object -ComObject WScript.Shell
        $sc  = $wsh.CreateShortcut($tempPath)
        $sc.TargetPath       = $TargetExe
        $sc.Arguments        = $Arguments
        $sc.WorkingDirectory = [System.IO.Path]::GetDirectoryName($TargetExe)
        $sc.IconLocation     = $TargetExe
        $sc.WindowStyle      = 1
        $sc.Save()

        # Move from temp to final Unicode path using .NET (handles any filename)
        if (Test-Path $ShortcutPath) {
            [System.IO.File]::Delete($ShortcutPath)
        }
        [System.IO.File]::Move($tempPath, $ShortcutPath)

        if ($Existing) {
            $Updated.Add($ShortcutPath) | Out-Null
            Write-Host "Updated : $ShortcutName" -ForegroundColor Cyan
        } else {
            $Created.Add($ShortcutPath) | Out-Null
            Write-Host "Created : $ShortcutName" -ForegroundColor Green
        }
    } catch {
        Write-Host "ERROR saving shortcut: $ShortcutName" -ForegroundColor Red
        Write-Host "  $_" -ForegroundColor DarkRed
    }
}

# --------- Shared state ----------
if (-not (Test-Path $ShortcutPath)) {
    if ($WhatIf) {
        Write-Host "[DRY RUN] Would create shortcut folder: $ShortcutPath"
    } else {
        New-Item -ItemType Directory -Path $ShortcutPath -Force | Out-Null
    }
}

$usedNames = New-Object System.Collections.Generic.HashSet[string]

# Shared created/updated lists (both passes write to these)
$created = [System.Collections.ArrayList]@()
$updated = [System.Collections.ArrayList]@()

# Per-library issue tracking
$hdSkipped        = [System.Collections.ArrayList]@()
$hdMissingSfo     = [System.Collections.ArrayList]@()
$hdMissingEboot   = [System.Collections.ArrayList]@()
$hdParseErrors    = [System.Collections.ArrayList]@()

$discMissingSfo   = [System.Collections.ArrayList]@()
$discMissingEboot = [System.Collections.ArrayList]@()
$discParseErrors  = [System.Collections.ArrayList]@()

# ===========================================================================
# PASS 1 — HD/PSN Games (dev_hdd0\game)
# Folder = Title ID | PARAM.SFO in root | EBOOT.BIN in USRDIR\
# ===========================================================================
Write-Host ""
Write-Host "=========================================" -ForegroundColor White
Write-Host " PASS 1: HD/PSN Games" -ForegroundColor White
Write-Host " $HdGamePath" -ForegroundColor DarkGray
Write-Host "=========================================" -ForegroundColor White
Write-Host ""

if (-not (Test-Path $HdGamePath)) {
    Write-Host "HD/PSN path not found - skipping: $HdGamePath" -ForegroundColor DarkYellow
} else {
    Get-ChildItem -Path $HdGamePath -Directory | ForEach-Object {
        $gameDir  = $_
        $gamePath = $gameDir.FullName

        if (Test-IsSkippableFolder -FolderName $gameDir.Name) {
            $hdSkipped.Add($gamePath) | Out-Null
            Write-Host "[$($gameDir.Name)] Skipped (non-game folder)" -ForegroundColor DarkGray
            return
        }

        $sfoPath = Join-Path $gamePath "PARAM.SFO"
        if (-not (Test-Path $sfoPath)) {
            $hdMissingSfo.Add($gamePath) | Out-Null
            Write-Host "[$($gameDir.Name)] No PARAM.SFO found – SKIP" -ForegroundColor DarkYellow
            return
        }

        $info = Get-SfoInfo -SfoPath $sfoPath
        if (-not $info) {
            $hdParseErrors.Add($gamePath) | Out-Null
            Write-Host "[$($gameDir.Name)] Could not read TITLE – SKIP" -ForegroundColor Red
            return
        }

        $ebootPath = Join-Path $gamePath "USRDIR\EBOOT.BIN"
        if (-not (Test-Path $ebootPath)) {
            $hdMissingEboot.Add($gamePath) | Out-Null
            Write-Host "[$($gameDir.Name)] No USRDIR\EBOOT.BIN found – SKIP" -ForegroundColor DarkYellow
            return
        }

        $title = Get-CleanTitle -Title $info.Title
        if (-not $title) {
            $hdParseErrors.Add($gamePath) | Out-Null
            Write-Host "[$($gameDir.Name)] Title empty after cleaning – SKIP" -ForegroundColor Red
            return
        }

        $regionTag = Get-RegionFromTitleId -TitleId $info.TitleId
        $baseName  = if ($regionTag) { "$title [$regionTag]" } else { $title }
        $resolved  = Resolve-ShortcutPath -BaseName $baseName -ShortcutFolder $ShortcutPath -UsedNames $usedNames

        Write-Shortcut `
            -ShortcutPath $resolved.Path `
            -ShortcutName $resolved.Name `
            -Existing     $resolved.Existing `
            -TargetExe    $Rpcs3ExePath `
            -Arguments    "--no-gui `"$ebootPath`"" `
            -WhatIf       $WhatIf `
            -Created      $created `
            -Updated      $updated
    }
}

# ===========================================================================
# PASS 2 — Disc Games (dev_hdd0\disc)
# Folder = readable title with region | PARAM.SFO in PS3_GAME\ | EBOOT.BIN in PS3_GAME\USRDIR\
# ===========================================================================
Write-Host ""
Write-Host "=========================================" -ForegroundColor White
Write-Host " PASS 2: Disc Games" -ForegroundColor White
Write-Host " $DiscGamePath" -ForegroundColor DarkGray
Write-Host "=========================================" -ForegroundColor White
Write-Host ""

if (-not (Test-Path $DiscGamePath)) {
    Write-Host "Disc path not found - skipping: $DiscGamePath" -ForegroundColor DarkYellow
} else {
    Get-ChildItem -Path $DiscGamePath -Directory | ForEach-Object {
        $gameDir  = $_
        $gamePath = $gameDir.FullName

        $sfoPath = Join-Path $gamePath "PS3_GAME\PARAM.SFO"
        if (-not (Test-Path $sfoPath)) {
            $discMissingSfo.Add($gamePath) | Out-Null
            Write-Host "[$($gameDir.Name)] No PS3_GAME\PARAM.SFO found – SKIP" -ForegroundColor DarkYellow
            return
        }

        $info = Get-SfoInfo -SfoPath $sfoPath
        if (-not $info) {
            $discParseErrors.Add($gamePath) | Out-Null
            Write-Host "[$($gameDir.Name)] Could not read TITLE – SKIP" -ForegroundColor Red
            return
        }

        $ebootPath = Join-Path $gamePath "PS3_GAME\USRDIR\EBOOT.BIN"
        if (-not (Test-Path $ebootPath)) {
            $discMissingEboot.Add($gamePath) | Out-Null
            Write-Host "[$($gameDir.Name)] No PS3_GAME\USRDIR\EBOOT.BIN found – SKIP" -ForegroundColor DarkYellow
            return
        }

        # Prefer PARAM.SFO title; fall back to cleaning the folder name
        $rawTitle = if ($info.Title) { $info.Title } else { Get-TitleFromFolderName -FolderName $gameDir.Name }
        $title    = Get-CleanTitle -Title $rawTitle
        if (-not $title) {
            $discParseErrors.Add($gamePath) | Out-Null
            Write-Host "[$($gameDir.Name)] Title empty after cleaning – SKIP" -ForegroundColor Red
            return
        }

        # Region: folder name first, fall back to Title ID from PARAM.SFO
        $regionTag = Get-RegionFromFolderName -FolderName $gameDir.Name
        if (-not $regionTag) {
            $regionTag = Get-RegionFromTitleId -TitleId $info.TitleId
        }
        $baseName = if ($regionTag) { "$title [$regionTag]" } else { $title }
        $resolved = Resolve-ShortcutPath -BaseName $baseName -ShortcutFolder $ShortcutPath -UsedNames $usedNames

        Write-Shortcut `
            -ShortcutPath $resolved.Path `
            -ShortcutName $resolved.Name `
            -Existing     $resolved.Existing `
            -TargetExe    $Rpcs3ExePath `
            -Arguments    "--no-gui `"$ebootPath`"" `
            -WhatIf       $WhatIf `
            -Created      $created `
            -Updated      $updated
    }
}

# --------- Summary ----------
Write-Host ""
Write-Host "================= SUMMARY =================" -ForegroundColor White
Write-Host ""
Write-Host "  OVERALL"
Write-Host "    Shortcuts CREATED        : $($created.Count)"
Write-Host "    Shortcuts UPDATED        : $($updated.Count)"
Write-Host ""
Write-Host "  HD/PSN GAMES (game\)"
Write-Host "    Skipped (non-game)       : $($hdSkipped.Count)"
Write-Host "    Missing PARAM.SFO        : $($hdMissingSfo.Count)"
Write-Host "    Missing EBOOT.BIN        : $($hdMissingEboot.Count)"
Write-Host "    Parse errors             : $($hdParseErrors.Count)"
Write-Host ""
Write-Host "  DISC GAMES (disc\)"
Write-Host "    Missing PS3_GAME\PARAM.SFO   : $($discMissingSfo.Count)"
Write-Host "    Missing PS3_GAME\EBOOT.BIN   : $($discMissingEboot.Count)"
Write-Host "    Parse errors                 : $($discParseErrors.Count)"
Write-Host ""
Write-Host "===========================================" -ForegroundColor White

[PSCustomObject]@{
    Created          = $created
    Updated          = $updated
    HdSkipped        = $hdSkipped
    HdMissingSfo     = $hdMissingSfo
    HdMissingEboot   = $hdMissingEboot
    HdParseErrors    = $hdParseErrors
    DiscMissingSfo   = $discMissingSfo
    DiscMissingEboot = $discMissingEboot
    DiscParseErrors  = $discParseErrors
}
