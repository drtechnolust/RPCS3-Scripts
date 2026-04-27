#Requires -Version 5.1
<#
.SYNOPSIS
    Creates or updates Windows shortcuts for both PS3 HD/PSN and disc games in one run.

.DESCRIPTION
    Performs a two-pass scan of the PS3 library and creates .lnk shortcuts pointing
    to RPCS3 with --no-gui for direct game launch. Both passes write to a single
    shared shortcut folder with a shared de-duplication tracker.

    Pass 1 -- HD/PSN games (dev_hdd0\game):
      Title ID named folders, PARAM.SFO in root, EBOOT.BIN in USRDIR\

    Pass 2 -- Disc games (dev_hdd0\disc):
      Human-readable named folders, PARAM.SFO in PS3_GAME\, EBOOT.BIN in PS3_GAME\USRDIR\

    Features:
      - Parses PARAM.SFO binary metadata natively for accurate game titles
      - Derives region tag from Title ID prefix (HD/PSN) or folder name (Disc)
      - Shared de-duplication tracker prevents duplicate shortcuts across both passes
      - Unicode-safe shortcut creation via GUID temp-file workaround
      - Skips non-game folders (DATA, INSTALL, cache, hidden)
      - Dry run mode previews all changes without writing anything
      - CSV log records every decision made in both passes

.PARAMETER DryRun
    When $true (default), shows what would be created without writing shortcuts.
    Set to $false in the CONFIG block to apply changes.

.EXAMPLE
    .\Create-RPCS3-Shortcuts-HD-DISC.ps1

.NOTES
    Region tag reference for HD/PSN Title ID prefixes:
      BCUS/BLUS/NPUB/NPUA = US
      BCES/BLES/NPEB/NPEA = EU
      BCJS/BLJM/NPJB      = JP
      BCAS/NPHA            = AS
      NPKB                 = KR

.VERSION
    2.0.0 - Config block, DryRun default, CSV log, MIT license.
             Replaced Read-Host DryRun prompt with config flag.
    1.0.0 - Initial release. Two-pass scan, Unicode-safe shortcuts.

.LICENSE
    MIT License
    Copyright (c) Paul Mardis
#>

# ==============================================================================
# CONFIG -- Edit this block. Do not put paths anywhere else in this script.
# ==============================================================================
$HdGamePath   = "D:\Arcade\System roms\Sony Playstation 3\dev_hdd0\game"
$DiscGamePath = "D:\Arcade\System roms\Sony Playstation 3\dev_hdd0\disc"
$ShortcutPath = "D:\Arcade\System roms\Sony Playstation 3\games\shortcuts"
$Rpcs3ExePath = "C:\Arcade\LaunchBox\Emulators\RPCS3\rpcs3.exe"
$DryRun       = $true    # Set to $false to write shortcuts
$LogDir       = Join-Path $ShortcutPath "Logs"
# ==============================================================================

function Set-ExitCode {
    param([int]$Code)
    $global:LASTEXITCODE = $Code
}

# ------------------------------------------------------------------------------
# Pre-flight
# ------------------------------------------------------------------------------
if (-not (Test-Path -LiteralPath $Rpcs3ExePath)) {
    Write-Host "ERROR: rpcs3.exe not found: $Rpcs3ExePath" -ForegroundColor Red
    Set-ExitCode 1
    return
}

# ------------------------------------------------------------------------------
# Parse PARAM.SFO binary -- returns Title, TitleId, ContentId or $null
# ------------------------------------------------------------------------------
function Read-ParamSfo {
    param([string]$SfoPath)

    if (-not (Test-Path -LiteralPath $SfoPath)) { return $null }

    try { $bytes = [System.IO.File]::ReadAllBytes($SfoPath) }
    catch { return $null }

    if ($bytes.Length -lt 0x14) { return $null }
    if ($bytes[0] -ne 0x00 -or $bytes[1] -ne 0x50 -or
        $bytes[2] -ne 0x53 -or $bytes[3] -ne 0x46) { return $null }

    $keyTableOffset  = [BitConverter]::ToInt32($bytes, 8)
    $dataTableOffset = [BitConverter]::ToInt32($bytes, 12)
    $indexEntries    = [BitConverter]::ToInt32($bytes, 16)

    $offset = 0x14
    $result = @{}

    for ($i = 0; $i -lt $indexEntries; $i++) {
        $keyOffset  = [BitConverter]::ToUInt16($bytes, $offset); $offset += 2
        $dataFmt    = [BitConverter]::ToUInt16($bytes, $offset); $offset += 2
        $dataLen    = [BitConverter]::ToInt32($bytes,  $offset); $offset += 4
        $offset    += 4  # maxLen
        $dataOffset = [BitConverter]::ToInt32($bytes,  $offset); $offset += 4

        $keyStart = $keyTableOffset + $keyOffset
        $k = $keyStart
        while ($k -lt $bytes.Length -and $bytes[$k] -ne 0) { $k++ }
        $keyName = [System.Text.Encoding]::ASCII.GetString($bytes, $keyStart, $k - $keyStart)

        if ($dataFmt -eq 0x0204) {
            $dataStart = $dataTableOffset + $dataOffset
            if ($dataStart -lt 0 -or $dataStart -ge $bytes.Length) { continue }

            $len      = [Math]::Min($dataLen, $bytes.Length - $dataStart)
            $valBytes = New-Object byte[] $len
            [Array]::Copy($bytes, $dataStart, $valBytes, 0, $len)

            $trimLen = $len
            for ($j = $len - 1; $j -ge 0; $j--) {
                if ($valBytes[$j] -eq 0) { $trimLen-- } else { break }
            }

            if ($trimLen -gt 0) {
                $result[$keyName] = [System.Text.Encoding]::UTF8.GetString($valBytes, 0, $trimLen)
            }
        }
    }

    if (-not $result.ContainsKey("TITLE")) { return $null }

    return [PSCustomObject]@{
        Title     = $result["TITLE"]
        TitleId   = if ($result.ContainsKey("TITLE_ID"))   { $result["TITLE_ID"] }   else { $null }
        ContentId = if ($result.ContainsKey("CONTENT_ID")) { $result["CONTENT_ID"] } else { $null }
    }
}

# ------------------------------------------------------------------------------
# Region from Title ID prefix (HD/PSN games)
# ------------------------------------------------------------------------------
function Get-RegionFromTitleId {
    param([string]$TitleId)

    if (-not $TitleId -or $TitleId.Length -lt 4) { return $null }
    $prefix = $TitleId.Substring(0, 4).ToUpper()

    switch ($prefix) {
        { $_ -in @("BCUS","BLUS","NPUB","NPUA") } { return "US" }
        { $_ -in @("BCES","BLES","NPEB","NPEA") } { return "EU" }
        { $_ -in @("BCJS","BLJM","NPJB") }        { return "JP" }
        { $_ -in @("BCAS","NPHA") }               { return "AS" }
        "NPKB"                                     { return "KR" }
        default                                    { return $null }
    }
}

# ------------------------------------------------------------------------------
# Region from disc folder name suffix: (USA), (Europe), (Japan), etc.
# ------------------------------------------------------------------------------
function Get-RegionFromFolderName {
    param([string]$FolderName)

    if ($FolderName -match '\(USA\)')    { return "US" }
    if ($FolderName -match '\(Europe\)') { return "EU" }
    if ($FolderName -match '\(Japan\)')  { return "JP" }
    if ($FolderName -match '\(Asia\)')   { return "AS" }
    if ($FolderName -match '\(Korea\)')  { return "KR" }
    if ($FolderName -match '\(China\)')  { return "CN" }
    return $null
}

# ------------------------------------------------------------------------------
# Convert disc folder name to readable title
# e.g. "LEGO_Indiana_Jones_(USA)" -> "LEGO Indiana Jones"
# ------------------------------------------------------------------------------
function Get-TitleFromFolderName {
    param([string]$FolderName)
    $title = $FolderName -replace '_\([^)]+\)$', ''
    return ($title -replace '_', ' ').Trim()
}

# ------------------------------------------------------------------------------
# Skip non-game folders in HD/PSN library
# ------------------------------------------------------------------------------
function Test-IsNonGameFolder {
    param([string]$FolderName)
    if ($FolderName -match "DATA$")    { return $true }
    if ($FolderName -match "INSTALL$") { return $true }
    if ($FolderName -match "_cache$")  { return $true }
    if ($FolderName.StartsWith("."))   { return $true }
    if ($FolderName.StartsWith("$"))   { return $true }
    return $false
}

# ------------------------------------------------------------------------------
# Strip characters invalid in Windows filenames
# ------------------------------------------------------------------------------
function Get-CleanTitle {
    param([string]$Title)

    $extraBad = [char[]]@([char]0xAE, [char]0x2122, [char]0xA9)
    $invalid  = [System.IO.Path]::GetInvalidFileNameChars() + $extraBad

    foreach ($c in $invalid) {
        $s = [string]$c
        if ($s.Length -gt 0) { $Title = $Title.Replace($s, "") }
    }
    return $Title.Trim()
}

# ------------------------------------------------------------------------------
# Resolve unique shortcut path -- avoids collisions, detects existing shortcuts
# ------------------------------------------------------------------------------
function Resolve-ShortcutPath {
    param(
        [string]$BaseName,
        [string]$ShortcutFolder,
        [System.Collections.Generic.HashSet[string]]$UsedNames
    )

    $candidate = $BaseName
    $counter   = 2

    while ($true) {
        $lnkPath = Join-Path $ShortcutFolder "$candidate.lnk"

        if (Test-Path -LiteralPath $lnkPath) {
            return [PSCustomObject]@{ Path = $lnkPath; Name = "$candidate.lnk"; Existing = $true }
        }

        if (-not $UsedNames.Contains($candidate.ToLower())) {
            $UsedNames.Add($candidate.ToLower()) | Out-Null
            return [PSCustomObject]@{ Path = $lnkPath; Name = "$candidate.lnk"; Existing = $false }
        }

        $candidate = "$BaseName ($counter)"
        $counter++
    }
}

# ------------------------------------------------------------------------------
# Write shortcut using GUID temp-file pattern for Unicode safety.
# WScript.Shell fails to save .lnk files whose path contains Unicode characters.
# Fix: save to ASCII temp path, then use .NET File.Move() to the Unicode path.
# ------------------------------------------------------------------------------
function Write-Shortcut {
    param(
        [string]$ShortcutPath,
        [string]$ShortcutName,
        [bool]$Existing,
        [string]$TargetExe,
        [string]$Arguments
    )

    if ($DryRun) {
        $action = if ($Existing) { "Would UPDATE" } else { "Would CREATE" }
        $color  = if ($Existing) { "Cyan" } else { "Green" }
        Write-Host ("  [{0}] {1}" -f $action, $ShortcutName) -ForegroundColor $color
        Write-Host ("    Target : {0}" -f $TargetExe)
        Write-Host ("    Args   : {0}" -f $Arguments)
        return "dryrun"
    }

    try {
        $tempPath = [System.IO.Path]::Combine(
            [System.IO.Path]::GetTempPath(),
            ("_rpcs3_{0}.lnk" -f [System.Guid]::NewGuid().ToString("N"))
        )

        $wsh = New-Object -ComObject WScript.Shell
        $sc  = $wsh.CreateShortcut($tempPath)
        $sc.TargetPath       = $TargetExe
        $sc.Arguments        = $Arguments
        $sc.WorkingDirectory = [System.IO.Path]::GetDirectoryName($TargetExe)
        $sc.IconLocation     = $TargetExe
        $sc.WindowStyle      = 1
        $sc.Save()

        if (Test-Path -LiteralPath $ShortcutPath) {
            [System.IO.File]::Delete($ShortcutPath)
        }
        [System.IO.File]::Move($tempPath, $ShortcutPath)

        if ($Existing) {
            Write-Host ("  [Updated] {0}" -f $ShortcutName) -ForegroundColor Cyan
            return "updated"
        } else {
            Write-Host ("  [Created] {0}" -f $ShortcutName) -ForegroundColor Green
            return "created"
        }
    }
    catch {
        Write-Host ("  [ERROR]   {0} -- {1}" -f $ShortcutName, $_.Exception.Message) -ForegroundColor Red
        return "error"
    }
}

# ==============================================================================
# SETUP
# ==============================================================================
if (-not (Test-Path -LiteralPath $ShortcutPath)) {
    if ($DryRun) {
        Write-Host "[DRY RUN] Would create: $ShortcutPath" -ForegroundColor Yellow
    } else {
        New-Item -ItemType Directory -Path $ShortcutPath -Force | Out-Null
    }
}

if (-not $DryRun -and -not (Test-Path -LiteralPath $LogDir)) {
    New-Item -ItemType Directory -Path $LogDir -Force | Out-Null
}

$LogFile = Join-Path $LogDir ("Create-RPCS3-Shortcuts_{0}.csv" -f (Get-Date -Format "yyyyMMdd_HHmmss"))
$LogRows = [System.Collections.Generic.List[PSCustomObject]]::new()

function Write-LogRow {
    param(
        [string]$Pass,
        [string]$FolderName,
        [string]$Title,
        [string]$TitleId,
        [string]$Region,
        [string]$ShortcutName,
        [string]$Action,
        [string]$Status,
        [string]$Note
    )
    $LogRows.Add([PSCustomObject]@{
        Timestamp    = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
        Pass         = $Pass
        FolderName   = $FolderName
        Title        = $Title
        TitleId      = $TitleId
        Region       = $Region
        ShortcutName = $ShortcutName
        Action       = $Action
        Status       = $Status
        Note         = $Note
    })
}

# Shared state across both passes
$UsedNames = New-Object System.Collections.Generic.HashSet[string]
$Counters  = @{ Created = 0; Updated = 0; Skipped = 0; Errors = 0 }

# Per-pass issue tracking
$HdSkipped        = [System.Collections.Generic.List[string]]::new()
$HdMissingSfo     = [System.Collections.Generic.List[string]]::new()
$HdMissingEboot   = [System.Collections.Generic.List[string]]::new()
$HdParseErrors    = [System.Collections.Generic.List[string]]::new()
$DiscMissingSfo   = [System.Collections.Generic.List[string]]::new()
$DiscMissingEboot = [System.Collections.Generic.List[string]]::new()
$DiscParseErrors  = [System.Collections.Generic.List[string]]::new()

# ==============================================================================
# BANNER
# ==============================================================================
Write-Host ""
Write-Host "Create-RPCS3-Shortcuts-HD-DISC" -ForegroundColor Cyan
Write-Host "==============================" -ForegroundColor Cyan
Write-Host "  HD/PSN  : $HdGamePath"
Write-Host "  Disc    : $DiscGamePath"
Write-Host "  Output  : $ShortcutPath"
Write-Host "  RPCS3   : $Rpcs3ExePath"
Write-Host "  DryRun  : $DryRun"
Write-Host ""
if ($DryRun) {
    Write-Host "  [DRY RUN] No shortcuts will be written." -ForegroundColor Yellow
    Write-Host ""
}

# ==============================================================================
# PASS 1 -- HD/PSN Games (dev_hdd0\game)
# ==============================================================================
Write-Host "Pass 1: HD/PSN Games" -ForegroundColor White
Write-Host "  $HdGamePath" -ForegroundColor DarkGray
Write-Host ""

if (-not (Test-Path -LiteralPath $HdGamePath)) {
    Write-Host "  Path not found -- skipping: $HdGamePath" -ForegroundColor Yellow
} else {
    $HdDirs = Get-ChildItem -LiteralPath $HdGamePath -Directory | Sort-Object Name
    foreach ($GameDir in $HdDirs) {
        $FolderName = $GameDir.Name
        $GamePath   = $GameDir.FullName

        if (Test-IsNonGameFolder -FolderName $FolderName) {
            $HdSkipped.Add($GamePath)
            Write-Host ("  [SKIP] {0} (non-game folder)" -f $FolderName) -ForegroundColor DarkGray
            Write-LogRow -Pass "HDD" -FolderName $FolderName -Action "Skip" -Status "NonGame" -Note "Non-game folder pattern"
            continue
        }

        $SfoPath = Join-Path $GamePath "PARAM.SFO"
        if (-not (Test-Path -LiteralPath $SfoPath)) {
            $HdMissingSfo.Add($GamePath)
            Write-Host ("  [MISS] {0} -- no PARAM.SFO" -f $FolderName) -ForegroundColor DarkYellow
            Write-LogRow -Pass "HDD" -FolderName $FolderName -Action "Skip" -Status "NoSFO" -Note "PARAM.SFO not found"
            continue
        }

        $Info = Read-ParamSfo -SfoPath $SfoPath
        if (-not $Info) {
            $HdParseErrors.Add($GamePath)
            Write-Host ("  [ERR]  {0} -- PARAM.SFO parse failed" -f $FolderName) -ForegroundColor Red
            Write-LogRow -Pass "HDD" -FolderName $FolderName -Action "Skip" -Status "ParseError" -Note "Could not read TITLE"
            continue
        }

        $EbootPath = Join-Path $GamePath "USRDIR\EBOOT.BIN"
        if (-not (Test-Path -LiteralPath $EbootPath)) {
            $HdMissingEboot.Add($GamePath)
            Write-Host ("  [MISS] {0} -- no USRDIR\EBOOT.BIN" -f $FolderName) -ForegroundColor DarkYellow
            Write-LogRow -Pass "HDD" -FolderName $FolderName -TitleId $Info.TitleId -Title $Info.Title -Action "Skip" -Status "NoEBOOT" -Note "USRDIR\EBOOT.BIN not found"
            continue
        }

        $Title = Get-CleanTitle -Title $Info.Title
        if (-not $Title) {
            $HdParseErrors.Add($GamePath)
            Write-Host ("  [ERR]  {0} -- title empty after cleaning" -f $FolderName) -ForegroundColor Red
            Write-LogRow -Pass "HDD" -FolderName $FolderName -TitleId $Info.TitleId -Action "Skip" -Status "EmptyTitle"
            continue
        }

        $Region   = Get-RegionFromTitleId -TitleId $Info.TitleId
        $BaseName = if ($Region) { "$Title [$Region]" } else { $Title }
        $Resolved = Resolve-ShortcutPath -BaseName $BaseName -ShortcutFolder $ShortcutPath -UsedNames $UsedNames

        $Result = Write-Shortcut -ShortcutPath $Resolved.Path -ShortcutName $Resolved.Name `
            -Existing $Resolved.Existing -TargetExe $Rpcs3ExePath `
            -Arguments ("--no-gui `"{0}`"" -f $EbootPath)

        Write-LogRow -Pass "HDD" -FolderName $FolderName -Title $Title -TitleId $Info.TitleId `
            -Region $Region -ShortcutName $Resolved.Name -Action $(if ($Resolved.Existing) { "Update" } else { "Create" }) `
            -Status $Result

        switch ($Result) {
            "created" { $Counters.Created++ }
            "updated" { $Counters.Updated++ }
            "dryrun"  { if ($Resolved.Existing) { $Counters.Updated++ } else { $Counters.Created++ } }
            default   { $Counters.Errors++ }
        }
    }
}

# ==============================================================================
# PASS 2 -- Disc Games (dev_hdd0\disc)
# ==============================================================================
Write-Host ""
Write-Host "Pass 2: Disc Games" -ForegroundColor White
Write-Host "  $DiscGamePath" -ForegroundColor DarkGray
Write-Host ""

if (-not (Test-Path -LiteralPath $DiscGamePath)) {
    Write-Host "  Path not found -- skipping: $DiscGamePath" -ForegroundColor Yellow
} else {
    $DiscDirs = Get-ChildItem -LiteralPath $DiscGamePath -Directory | Sort-Object Name
    foreach ($GameDir in $DiscDirs) {
        $FolderName = $GameDir.Name
        $GamePath   = $GameDir.FullName

        $SfoPath = Join-Path $GamePath "PS3_GAME\PARAM.SFO"
        if (-not (Test-Path -LiteralPath $SfoPath)) {
            $DiscMissingSfo.Add($GamePath)
            Write-Host ("  [MISS] {0} -- no PS3_GAME\PARAM.SFO" -f $FolderName) -ForegroundColor DarkYellow
            Write-LogRow -Pass "Disc" -FolderName $FolderName -Action "Skip" -Status "NoSFO" -Note "PS3_GAME\PARAM.SFO not found"
            continue
        }

        $Info = Read-ParamSfo -SfoPath $SfoPath
        if (-not $Info) {
            $DiscParseErrors.Add($GamePath)
            Write-Host ("  [ERR]  {0} -- PARAM.SFO parse failed" -f $FolderName) -ForegroundColor Red
            Write-LogRow -Pass "Disc" -FolderName $FolderName -Action "Skip" -Status "ParseError" -Note "Could not read TITLE"
            continue
        }

        $EbootPath = Join-Path $GamePath "PS3_GAME\USRDIR\EBOOT.BIN"
        if (-not (Test-Path -LiteralPath $EbootPath)) {
            $DiscMissingEboot.Add($GamePath)
            Write-Host ("  [MISS] {0} -- no PS3_GAME\USRDIR\EBOOT.BIN" -f $FolderName) -ForegroundColor DarkYellow
            Write-LogRow -Pass "Disc" -FolderName $FolderName -TitleId $Info.TitleId -Title $Info.Title -Action "Skip" -Status "NoEBOOT"
            continue
        }

        # Prefer PARAM.SFO title; fall back to cleaning the folder name
        $RawTitle = if ($Info.Title) { $Info.Title } else { Get-TitleFromFolderName -FolderName $FolderName }
        $Title    = Get-CleanTitle -Title $RawTitle
        if (-not $Title) {
            $DiscParseErrors.Add($GamePath)
            Write-Host ("  [ERR]  {0} -- title empty after cleaning" -f $FolderName) -ForegroundColor Red
            Write-LogRow -Pass "Disc" -FolderName $FolderName -TitleId $Info.TitleId -Action "Skip" -Status "EmptyTitle"
            continue
        }

        # Region: folder name first, then Title ID
        $Region   = Get-RegionFromFolderName -FolderName $FolderName
        if (-not $Region) { $Region = Get-RegionFromTitleId -TitleId $Info.TitleId }
        $BaseName = if ($Region) { "$Title [$Region]" } else { $Title }
        $Resolved = Resolve-ShortcutPath -BaseName $BaseName -ShortcutFolder $ShortcutPath -UsedNames $UsedNames

        $Result = Write-Shortcut -ShortcutPath $Resolved.Path -ShortcutName $Resolved.Name `
            -Existing $Resolved.Existing -TargetExe $Rpcs3ExePath `
            -Arguments ("--no-gui `"{0}`"" -f $EbootPath)

        Write-LogRow -Pass "Disc" -FolderName $FolderName -Title $Title -TitleId $Info.TitleId `
            -Region $Region -ShortcutName $Resolved.Name -Action $(if ($Resolved.Existing) { "Update" } else { "Create" }) `
            -Status $Result

        switch ($Result) {
            "created" { $Counters.Created++ }
            "updated" { $Counters.Updated++ }
            "dryrun"  { if ($Resolved.Existing) { $Counters.Updated++ } else { $Counters.Created++ } }
            default   { $Counters.Errors++ }
        }
    }
}

# ==============================================================================
# WRITE LOG
# ==============================================================================
if ($LogRows.Count -gt 0 -and -not $DryRun) {
    try {
        $LogRows | Export-Csv -LiteralPath $LogFile -NoTypeInformation -Encoding UTF8
        Write-Host ""
        Write-Host ("  Log: {0}" -f $LogFile) -ForegroundColor Gray
    } catch {
        Write-Host ("  WARNING: Could not write log -- {0}" -f $_) -ForegroundColor Yellow
    }
}

# ==============================================================================
# SUMMARY
# ==============================================================================
Write-Host ""
Write-Host "Summary" -ForegroundColor White
Write-Host ("  {0,-30} {1}" -f "Shortcuts created :", $Counters.Created) -ForegroundColor Green
Write-Host ("  {0,-30} {1}" -f "Shortcuts updated :", $Counters.Updated) -ForegroundColor Cyan
Write-Host ("  {0,-30} {1}" -f "Errors :", $Counters.Errors) -ForegroundColor $(if ($Counters.Errors -gt 0) { "Red" } else { "Gray" })
Write-Host ""
Write-Host "  HD/PSN Games" -ForegroundColor DarkGray
Write-Host ("    {0,-28} {1}" -f "Skipped (non-game) :", $HdSkipped.Count)
Write-Host ("    {0,-28} {1}" -f "Missing PARAM.SFO :", $HdMissingSfo.Count)
Write-Host ("    {0,-28} {1}" -f "Missing EBOOT.BIN :", $HdMissingEboot.Count)
Write-Host ("    {0,-28} {1}" -f "Parse errors :", $HdParseErrors.Count)
Write-Host ""
Write-Host "  Disc Games" -ForegroundColor DarkGray
Write-Host ("    {0,-28} {1}" -f "Missing PS3_GAME\PARAM.SFO :", $DiscMissingSfo.Count)
Write-Host ("    {0,-28} {1}" -f "Missing PS3_GAME\EBOOT.BIN :", $DiscMissingEboot.Count)
Write-Host ("    {0,-28} {1}" -f "Parse errors :", $DiscParseErrors.Count)

if ($DryRun) {
    Write-Host ""
    Write-Host "  DRY RUN complete. Set `$DryRun = `$false in the CONFIG block to apply." -ForegroundColor Yellow
}

Write-Host ""
Set-ExitCode 0
