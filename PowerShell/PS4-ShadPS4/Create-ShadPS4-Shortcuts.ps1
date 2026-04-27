#Requires -Version 5.1
<#
.SYNOPSIS
    Creates or updates Windows shortcuts for PS4 games in a ShadPS4 library.

.DESCRIPTION
    Scans a folder of CUSA-named PS4 game directories, reads each game's
    sce_sys\param.sfo binary metadata to extract the game title and region,
    and creates a .lnk shortcut pointing to shadps4.exe for direct launch.

    Features:
      - Reads PARAM.SFO natively -- no external tools required
      - Derives region tag (US/EU/JP/AS/KR/CN) from Content ID
      - Handles Unicode game titles (Japanese, Korean, etc.) via a GUID
        temp-file workaround that bypasses WScript.Shell ANSI limitations
      - De-duplicates shortcut names when titles collide
      - Updates existing shortcuts in-place when re-run
      - Dry run mode previews all changes without writing anything
      - Writes a CSV log of every action taken

.PARAMETER DryRun
    When $true (default), shows what would be created or updated without
    writing any shortcuts. Set to $false in the CONFIG block to apply.

.EXAMPLE
    .\Create-ShadPS4-Shortcuts.ps1

.EXAMPLE
    # Run with DryRun disabled directly from the command line:
    # Open script, set $DryRun = $false in the CONFIG block, then run.

.NOTES
    Shortcut arguments use the ShadPS4 CLI syntax: -g "path\to\eboot.bin"
    The -f true flag enables fullscreen mode (confirmed working).

    ShadPS4 Qt Launcher overwrites the Pre-release folder in-place on every
    nightly update -- shortcuts pointing to that folder never need regenerating.

.VERSION
    2.0.0 - Full rewrite. Config block, CSV log, ISE-compatible, MIT header.
            Merged Unicode temp-file fix from v2 with -f true flag from v1.
    1.1.0 - Added Unicode temp-file workaround for Japanese titles.
    1.0.0 - Initial release.

.LICENSE
    MIT License
    Copyright (c) Paul Mardis
#>

# ==============================================================================
# CONFIG -- Edit this block. Do not put paths anywhere else in this script.
# ==============================================================================
$RootPath     = "D:\Arcade\System roms\Sony Playstation 4\Official PS4 Games"
$ShortcutPath = "D:\Arcade\System roms\Sony Playstation 4\PS4 Shortcuts 2"
$ShadExePath  = "C:\Arcade\LaunchBox\Emulators\ShadPS4QT\versions\Pre-release\shadps4.exe"
$Fullscreen   = $true    # Adds -f true to launch arguments
$DryRun       = $true    # Set to $false to create or update shortcuts
$LogDir       = Join-Path $ShortcutPath "Logs"
# ==============================================================================

function Set-ExitCode {
    param([int]$Code)
    $global:LASTEXITCODE = $Code
}

# ------------------------------------------------------------------------------
# Pre-flight checks
# ------------------------------------------------------------------------------
if (-not (Test-Path -LiteralPath $RootPath)) {
    Write-Host "ERROR: Game library not found: $RootPath" -ForegroundColor Red
    Set-ExitCode 1
    return
}

if (-not (Test-Path -LiteralPath $ShadExePath)) {
    Write-Host "ERROR: shadps4.exe not found: $ShadExePath" -ForegroundColor Red
    Set-ExitCode 1
    return
}

# ------------------------------------------------------------------------------
# Parse PARAM.SFO binary -- returns Title, TitleId, ContentId or $null
# ------------------------------------------------------------------------------
function Read-ParamSfo {
    param([string]$SfoPath)

    if (-not (Test-Path -LiteralPath $SfoPath)) { return $null }

    try {
        $bytes = [System.IO.File]::ReadAllBytes($SfoPath)
    } catch { return $null }

    if ($bytes.Length -lt 0x14) { return $null }

    # Magic: 00 50 53 46 ("\0PSF")
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
        $dataMaxLen = [BitConverter]::ToInt32($bytes,  $offset); $offset += 4
        $dataOffset = [BitConverter]::ToInt32($bytes,  $offset); $offset += 4

        $keyStart = $keyTableOffset + $keyOffset
        $k = $keyStart
        while ($k -lt $bytes.Length -and $bytes[$k] -ne 0) { $k++ }
        $keyName = [System.Text.Encoding]::ASCII.GetString($bytes, $keyStart, $k - $keyStart)

        # 0x0204 = UTF-8 string
        if ($dataFmt -eq 0x0204) {
            $dataStart = $dataTableOffset + $dataOffset
            if ($dataStart -lt 0 -or $dataStart -ge $bytes.Length) { continue }

            $len = [Math]::Min($dataLen, $bytes.Length - $dataStart)
            $valBytes = New-Object byte[] $len
            [Array]::Copy($bytes, $dataStart, $valBytes, 0, $len)

            # Trim null bytes from end
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
        TitleId   = if ($result.ContainsKey("TITLE_ID"))   { $result["TITLE_ID"] }   else { "" }
        ContentId = if ($result.ContainsKey("CONTENT_ID")) { $result["CONTENT_ID"] } else { "" }
    }
}

# ------------------------------------------------------------------------------
# Derive region tag from Content ID first character
# ------------------------------------------------------------------------------
function Get-RegionTag {
    param([string]$ContentId)

    if ([string]::IsNullOrWhiteSpace($ContentId)) { return "" }

    switch ($ContentId[0].ToString().ToUpper()) {
        "U" { return "US" }
        "E" { return "EU" }
        "J" { return "JP" }
        "H" { return "AS" }
        "K" { return "KR" }
        "C" { return "CN" }
        default { return "" }
    }
}

# ------------------------------------------------------------------------------
# Strip characters that are invalid in Windows filenames
# ------------------------------------------------------------------------------
function Get-SafeTitle {
    param([string]$Title)

    $invalidChars = [System.IO.Path]::GetInvalidFileNameChars()
    $extraChars   = [char[]]@([char]0xAE, [char]0x2122, [char]0xA9)  # (R) TM (C)

    $clean = $Title
    foreach ($c in ($invalidChars + $extraChars)) {
        $clean = $clean.Replace([string]$c, "")
    }

    return $clean.Trim()
}

# ------------------------------------------------------------------------------
# Create or overwrite a .lnk shortcut using the GUID temp-file pattern.
# WScript.Shell cannot save .lnk files whose path contains Unicode characters.
# Fix: save to an ASCII temp path first, then use .NET File.Move() which
# handles Unicode destination paths correctly.
# ------------------------------------------------------------------------------
function Save-Shortcut {
    param(
        [string]$ShortcutFullPath,
        [string]$TargetExe,
        [string]$Arguments,
        [string]$WorkingDir,
        [object]$WshShell
    )

    $tempPath = [System.IO.Path]::Combine(
        [System.IO.Path]::GetTempPath(),
        ("_shadps4_{0}.lnk" -f [System.Guid]::NewGuid().ToString("N"))
    )

    $sc = $WshShell.CreateShortcut($tempPath)
    $sc.TargetPath       = $TargetExe
    $sc.Arguments        = $Arguments
    $sc.WorkingDirectory = $WorkingDir
    $sc.IconLocation     = $TargetExe
    $sc.WindowStyle      = 1
    $sc.Save()

    # Delete existing shortcut first (handles update case)
    if (Test-Path -LiteralPath $ShortcutFullPath) {
        [System.IO.File]::Delete($ShortcutFullPath)
        $wait = 0
        while ((Test-Path -LiteralPath $ShortcutFullPath) -and $wait -lt 10) {
            Start-Sleep -Milliseconds 50
            $wait++
        }
    }

    [System.IO.File]::Move($tempPath, $ShortcutFullPath)
}

# ==============================================================================
# SETUP
# ==============================================================================
if (-not (Test-Path -LiteralPath $ShortcutPath)) {
    if ($DryRun) {
        Write-Host "[DRY RUN] Would create shortcut folder: $ShortcutPath" -ForegroundColor Yellow
    } else {
        New-Item -ItemType Directory -Path $ShortcutPath -Force | Out-Null
    }
}

if (-not $DryRun -and -not (Test-Path -LiteralPath $LogDir)) {
    New-Item -ItemType Directory -Path $LogDir -Force | Out-Null
}

$LogFile = Join-Path $LogDir ("Create-ShadPS4-Shortcuts_{0}.csv" -f (Get-Date -Format "yyyyMMdd_HHmmss"))

$LogRows  = [System.Collections.Generic.List[PSCustomObject]]::new()
$WshShell = New-Object -ComObject WScript.Shell
$UsedNames = New-Object System.Collections.Generic.HashSet[string]

$Arguments_Template = if ($Fullscreen) { '-g "{0}" -f true' } else { '-g "{0}"' }
$WorkingDir = [System.IO.Path]::GetDirectoryName($ShadExePath)

$CountCreated   = 0
$CountUpdated   = 0
$CountSkipped   = 0
$CountFailed    = 0
$MissingSfo     = [System.Collections.Generic.List[string]]::new()
$MissingEboot   = [System.Collections.Generic.List[string]]::new()
$ParseErrors    = [System.Collections.Generic.List[string]]::new()

function Write-LogRow {
    param(
        [string]$FolderName,
        [string]$TitleId,
        [string]$ContentId,
        [string]$Title,
        [string]$Region,
        [string]$ShortcutName,
        [string]$Action,
        [string]$Status,
        [string]$Error
    )
    $LogRows.Add([PSCustomObject]@{
        Timestamp    = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
        FolderName   = $FolderName
        TitleId      = $TitleId
        ContentId    = $ContentId
        Title        = $Title
        Region       = $Region
        ShortcutName = $ShortcutName
        Action       = $Action
        Status       = $Status
        Error        = $Error
    })
}

# ==============================================================================
# BANNER
# ==============================================================================
Write-Host ""
Write-Host "Create-ShadPS4-Shortcuts" -ForegroundColor Cyan
Write-Host "========================" -ForegroundColor Cyan
Write-Host "  Library : $RootPath"
Write-Host "  Output  : $ShortcutPath"
Write-Host "  Exe     : $ShadExePath"
Write-Host "  FullScr : $Fullscreen"
Write-Host "  DryRun  : $DryRun"
Write-Host ""
if ($DryRun) {
    Write-Host "  [DRY RUN] No shortcuts will be written." -ForegroundColor Yellow
    Write-Host ""
}

# ==============================================================================
# MAIN LOOP
# ==============================================================================
$GameDirs = Get-ChildItem -LiteralPath $RootPath -Directory | Sort-Object Name
$Total    = $GameDirs.Count
$Index    = 0

foreach ($GameDir in $GameDirs) {
    $Index++
    $FolderName = $GameDir.Name
    $GamePath   = $GameDir.FullName

    Write-Host ("  [{0}/{1}] {2}" -f $Index, $Total, $FolderName) -ForegroundColor Gray

    # --- param.sfo ---
    $SfoPath = Join-Path $GamePath "sce_sys\param.sfo"
    if (-not (Test-Path -LiteralPath $SfoPath)) {
        Write-Host "    SKIP  No param.sfo" -ForegroundColor Yellow
        $MissingSfo.Add($GamePath)
        Write-LogRow -FolderName $FolderName -Action "Skip" -Status "Skipped" -Error "No param.sfo"
        $CountSkipped++
        continue
    }

    # --- parse SFO ---
    $Info = Read-ParamSfo -SfoPath $SfoPath
    if (-not $Info) {
        Write-Host "    SKIP  Could not parse param.sfo" -ForegroundColor Yellow
        $ParseErrors.Add($GamePath)
        Write-LogRow -FolderName $FolderName -Action "Skip" -Status "Skipped" -Error "param.sfo parse failed"
        $CountSkipped++
        continue
    }

    # --- eboot.bin ---
    $Eboot = Get-ChildItem -LiteralPath $GamePath -Recurse -Filter "eboot.bin" -File -ErrorAction SilentlyContinue |
             Select-Object -First 1
    if (-not $Eboot) {
        Write-Host "    SKIP  No eboot.bin" -ForegroundColor Yellow
        $MissingEboot.Add($GamePath)
        Write-LogRow -FolderName $FolderName -TitleId $Info.TitleId -ContentId $Info.ContentId `
            -Title $Info.Title -Action "Skip" -Status "Skipped" -Error "No eboot.bin"
        $CountSkipped++
        continue
    }

    # --- clean title ---
    $CleanTitle = Get-SafeTitle -Title $Info.Title
    if (-not $CleanTitle) {
        Write-Host "    SKIP  Title empty after cleaning" -ForegroundColor Yellow
        $ParseErrors.Add($GamePath)
        Write-LogRow -FolderName $FolderName -TitleId $Info.TitleId -ContentId $Info.ContentId `
            -Title $Info.Title -Action "Skip" -Status "Skipped" -Error "Title empty after cleaning"
        $CountSkipped++
        continue
    }

    # --- region ---
    $Region   = Get-RegionTag -ContentId $Info.ContentId
    $BaseName = if ($Region) { "$CleanTitle [$Region]" } else { $CleanTitle }

    # --- de-duplicate shortcut name ---
    $ShortcutFull = $null
    $ShortcutName = $null
    $IsExisting   = $false
    $CandidateBase = $BaseName
    $Counter = 2

    while ($true) {
        $ShortcutName = "$CandidateBase.lnk"
        $CandidatePath = Join-Path $ShortcutPath $ShortcutName

        if (Test-Path -LiteralPath $CandidatePath) {
            $ShortcutFull = $CandidatePath
            $IsExisting   = $true
            break
        }

        if (-not $UsedNames.Contains($CandidateBase.ToLower())) {
            $ShortcutFull = $CandidatePath
            $UsedNames.Add($CandidateBase.ToLower()) | Out-Null
            break
        }

        $CandidateBase = "$BaseName ($Counter)"
        $Counter++
    }

    # --- build arguments ---
    $LaunchArgs = $Arguments_Template -f $Eboot.FullName
    $ActionLabel = if ($IsExisting) { "Update" } else { "Create" }

    # --- dry run ---
    if ($DryRun) {
        $Color = if ($IsExisting) { "Cyan" } else { "Green" }
        Write-Host ("    {0,-6} {1}" -f "WOULD $($ActionLabel.ToUpper()):", $ShortcutName) -ForegroundColor $Color
        Write-Host ("           Args : {0}" -f $LaunchArgs)
        Write-LogRow -FolderName $FolderName -TitleId $Info.TitleId -ContentId $Info.ContentId `
            -Title $Info.Title -Region $Region -ShortcutName $ShortcutName `
            -Action $ActionLabel -Status "DryRun" -Error ""
        if ($IsExisting) { $CountUpdated++ } else { $CountCreated++ }
        continue
    }

    # --- write shortcut ---
    try {
        Save-Shortcut -ShortcutFullPath $ShortcutFull `
                      -TargetExe       $ShadExePath `
                      -Arguments       $LaunchArgs `
                      -WorkingDir      $WorkingDir `
                      -WshShell        $WshShell

        $Color = if ($IsExisting) { "Cyan" } else { "Green" }
        Write-Host ("    {0,-6} {1}" -f "${ActionLabel}d:", $ShortcutName) -ForegroundColor $Color
        Write-LogRow -FolderName $FolderName -TitleId $Info.TitleId -ContentId $Info.ContentId `
            -Title $Info.Title -Region $Region -ShortcutName $ShortcutName `
            -Action $ActionLabel -Status "Success" -Error ""

        if ($IsExisting) { $CountUpdated++ } else { $CountCreated++ }
    }
    catch {
        Write-Host ("    FAILED: {0}" -f $ShortcutName) -ForegroundColor Red
        Write-Host ("      $_") -ForegroundColor DarkRed
        Write-LogRow -FolderName $FolderName -TitleId $Info.TitleId -ContentId $Info.ContentId `
            -Title $Info.Title -Region $Region -ShortcutName $ShortcutName `
            -Action $ActionLabel -Status "Failed" -Error $_.Exception.Message
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
    } catch {
        Write-Host "  WARNING: Could not write log -- $_" -ForegroundColor Yellow
    }
}

# ==============================================================================
# SUMMARY
# ==============================================================================
Write-Host ""
Write-Host "Summary" -ForegroundColor White
Write-Host ("  {0,-20} {1}" -f "Created :", $CountCreated)  -ForegroundColor Green
Write-Host ("  {0,-20} {1}" -f "Updated :", $CountUpdated)  -ForegroundColor Cyan
Write-Host ("  {0,-20} {1}" -f "Skipped :", $CountSkipped)  -ForegroundColor Yellow
Write-Host ("  {0,-20} {1}" -f "Failed  :", $CountFailed)   -ForegroundColor $(if ($CountFailed -gt 0) { "Red" } else { "Gray" })
Write-Host ("  {0,-20} {1}" -f "Missing SFO :", $MissingSfo.Count)   -ForegroundColor $(if ($MissingSfo.Count -gt 0) { "Yellow" } else { "Gray" })
Write-Host ("  {0,-20} {1}" -f "Missing eboot :", $MissingEboot.Count) -ForegroundColor $(if ($MissingEboot.Count -gt 0) { "Yellow" } else { "Gray" })
Write-Host ("  {0,-20} {1}" -f "Parse errors :", $ParseErrors.Count) -ForegroundColor $(if ($ParseErrors.Count -gt 0) { "Yellow" } else { "Gray" })

if ($DryRun) {
    Write-Host ""
    Write-Host "  DRY RUN complete. Set `$DryRun = `$false in the CONFIG block to apply." -ForegroundColor Yellow
}

Write-Host ""
Set-ExitCode 0