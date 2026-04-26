param(
    # Root folder with your CUSA folders
    [string]$RootPath     = "D:\Arcade\System roms\Sony Playstation 4\Official PS4 Games",

    # Folder where shortcuts will be created/updated
    [string]$ShortcutPath = "D:\Arcade\System roms\Sony Playstation 4\PS4 Shortcuts 3",

    # Path to shadps4.exe (CLI)
    [string]$ShadExePath  = "C:\Arcade\LaunchBox\Emulators\ShadPS4QT\versions\v0.12.5 - kyosan - 2025-11-07\shadps4.exe"
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
    param(
        [string]$ContentId
    )

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

$created       = @()
$updated       = @()
$missingSfo    = @()
$missingEboot  = @()
$parseErrors   = @()

$wsh = New-Object -ComObject WScript.Shell
$usedNames = New-Object System.Collections.Generic.HashSet[string]

$extraBad = "®","™","©"

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
            # Fix/update this existing shortcut
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

    # Build arguments for new CLI syntax after PR-1507
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
    $sc = $wsh.CreateShortcut($shortcutFull)
    $sc.TargetPath       = $ShadExePath
    $sc.Arguments        = $arguments
    $sc.WorkingDirectory = [System.IO.Path]::GetDirectoryName($ShadExePath)
    $sc.IconLocation     = $ShadExePath
    $sc.WindowStyle      = 1
    $sc.Save()

    if ($existing) {
        $updated += $shortcutFull
        Write-Host "Updated shortcut: $shortcutName"
    } else {
        $created += $shortcutFull
        Write-Host "Created shortcut: $shortcutName"
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
