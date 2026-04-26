<#
.SYNOPSIS
    Create-ShadPS4-Shortcuts.ps1

.DESCRIPTION
    Scans a PS4 game library folder containing CUSA/title-ID-named subfolders and
    automatically creates or updates Windows shortcuts (.lnk files) for each valid
    game, pointing to shadps4.exe for direct launch.

    Reads each game's sce_sys\param.sfo binary metadata to extract the proper game
    title and derives a region tag (US/EU/JP/AS/KR/CN) from the Content ID. Launches
    games in fullscreen by default. Points to the fixed Pre-release folder maintained
    by the ShadPS4 Qt Launcher, which overwrites this folder in-place on every nightly
    update so shortcuts never need regenerating after an update. Includes a dry run
    mode to preview all changes before committing them. Unicode game titles (Japanese,
    etc.) are handled correctly via a temp-file save workaround that bypasses the
    WScript.Shell ANSI filename limitation.

.AUTHOR
    Paul Mardis (drtechnolust)

.LICENSE
    MIT License

    Copyright (c) 2026 Paul Mardis (drtechnolust)

    Permission is hereby granted, free of charge, to any person obtaining a copy
    of this software and associated documentation files (the "Software"), to deal
    in the Software without restriction, including without limitation the rights
    to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
    copies of the Software, and to permit persons to whom the Software is
    furnished to do so, subject to the following conditions:

    The above copyright notice and this permission notice shall be included in all
    copies or substantial portions of the Software.

    THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
    IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
    FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
    AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
    LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
    OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
    SOFTWARE.
#>
param(
    # Root folder with your CUSA folders
    [string]$RootPath     = "D:\Arcade\System roms\Sony Playstation 4\Official PS4 Games",

    # Folder where shortcuts will be created/updated
    [string]$ShortcutPath = "D:\Arcade\System roms\Sony Playstation 4\PS4 Shortcuts 2",

    # Path to shadps4.exe - points to the fixed Pre-release folder which the Qt Launcher
    # always overwrites in-place on update, so this path never needs changing
    [string]$ShadExePath  = "C:\Arcade\LaunchBox\Emulators\ShadPS4QT\versions\Pre-release\shadps4.exe"
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

if (-not (Test-Path $ShadExePath)) {
    Write-Host "ERROR: shadps4.exe not found at path:`n  $ShadExePath" -ForegroundColor Red
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
        TitleId   = $result["TITLE_ID"]
        ContentId = $result["CONTENT_ID"]
    }
}

# --------- Helper: get region from CONTENT_ID ----------
function Get-RegionTag {
    param([string]$ContentId)

    if (-not $ContentId) { return $null }
    $c = $ContentId[0].ToString().ToUpper()

    switch ($c) {
        "U" { return "US" } # Americas
        "E" { return "EU" }
        "J" { return "JP" }
        "H" { return "AS" } # Asia
        "K" { return "KR" }
        "C" { return "CN" }
        default { return $null }
    }
}

# --------- Prep & tracking ----------
if (-not (Test-Path $ShortcutPath)) {
    if ($WhatIf) {
        Write-Host "[DRY RUN] Would create shortcut folder: $ShortcutPath"
    } else {
        New-Item -ItemType Directory -Path $ShortcutPath -Force | Out-Null
    }
}

$created      = @()
$updated      = @()
$missingSfo   = @()
$missingEboot = @()
$parseErrors  = @()

$wsh       = New-Object -ComObject WScript.Shell
$usedNames = New-Object System.Collections.Generic.HashSet[string]

$extraBad = [char[]]@([char]0xAE, [char]0x2122, [char]0xA9)  # ®, ™, ©

Write-Host ""
Write-Host "Scanning PS4 games in: $RootPath"
Write-Host ""

Get-ChildItem -Path $RootPath -Directory | ForEach-Object {
    $gameDir  = $_
    $gamePath = $gameDir.FullName

    $sfoPath = Join-Path $gamePath "sce_sys\param.sfo"
    if (-not (Test-Path $sfoPath)) {
        $missingSfo += $gamePath
        Write-Host "[$($gameDir.Name)] No param.sfo found – SKIP"
        return
    }

    $info = Get-SfoInfo -SfoPath $sfoPath
    if (-not $info) {
        $parseErrors += $gamePath
        Write-Host "[$($gameDir.Name)] Could not read TITLE – SKIP"
        return
    }

    # Locate eboot.bin
    $eboot = Get-ChildItem -Path $gamePath -Recurse -Filter "eboot.bin" -File -ErrorAction SilentlyContinue | Select-Object -First 1
    if (-not $eboot) {
        $missingEboot += $gamePath
        Write-Host "[$($gameDir.Name)] No eboot.bin found – SKIP"
        return
    }

    # Clean title (Unicode-safe)
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
        Write-Host "[$($gameDir.Name)] Title empty after cleaning – SKIP"
        return
    }

    # Region tag from CONTENT_ID
    $regionTag = Get-RegionTag -ContentId $info.ContentId
    if ($regionTag) {
        $baseName = "$title [$regionTag]"
    } else {
        $baseName = $title
    }

    # Determine shortcut name (preserve existing if present, otherwise new with de-dup)
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

    # Original working argument syntax
    $argEboot  = $eboot.FullName
    $arguments = "-g `"$argEboot`" -f true"

    if ($WhatIf) {
        if ($existing) {
            Write-Host "[DRY RUN] Would UPDATE shortcut: $shortcutName"
        } else {
            Write-Host "[DRY RUN] Would CREATE shortcut: $shortcutName"
        }
        Write-Host "         Target   : $ShadExePath"
        Write-Host "         Arguments: $arguments"
        return
    }

    # Create or update shortcut
    # WScript.Shell fails to save .lnk files with Unicode characters in the filename
    # (Japanese titles etc.) — fix: save to a unique ASCII temp file first, then use
    # .NET File.Move() which handles Unicode paths correctly.
    try {
        $tempPath = [System.IO.Path]::Combine(
            [System.IO.Path]::GetTempPath(),
            "_shadps4_tmp_$([System.Guid]::NewGuid().ToString('N')).lnk"
        )

        $sc = $wsh.CreateShortcut($tempPath)
        $sc.TargetPath       = $ShadExePath
        $sc.Arguments        = $arguments
        $sc.WorkingDirectory = [System.IO.Path]::GetDirectoryName($ShadExePath)
        $sc.IconLocation     = $ShadExePath
        $sc.WindowStyle      = 1
        $sc.Save()

        if (Test-Path $shortcutFull) {
            [System.IO.File]::Delete($shortcutFull)
            $retries = 0
            while ((Test-Path $shortcutFull) -and $retries -lt 10) {
                Start-Sleep -Milliseconds 50
                $retries++
            }
        }
        [System.IO.File]::Move($tempPath, $shortcutFull)

        if ($existing) {
            $updated += $shortcutFull
            Write-Host "Updated shortcut: $shortcutName"
        } else {
            $created += $shortcutFull
            Write-Host "Created shortcut: $shortcutName"
        }
    } catch {
        Write-Host "ERROR saving shortcut: $shortcutName" -ForegroundColor Red
        Write-Host "  $_" -ForegroundColor DarkRed
    }
}

Write-Host ""
Write-Host "---------- SUMMARY ----------"
Write-Host "Shortcuts CREATED      : $($created.Count)"
Write-Host "Shortcuts UPDATED      : $($updated.Count)"
Write-Host "Missing param.sfo      : $($missingSfo.Count)"
Write-Host "Missing eboot.bin      : $($missingEboot.Count)"
Write-Host "Parse errors           : $($parseErrors.Count)"
Write-Host "-----------------------------"

[PSCustomObject]@{
    Created      = $created
    Updated      = $updated
    MissingSfo   = $missingSfo
    MissingEboot = $missingEboot
    ParseErrors  = $parseErrors
}