#Requires -Version 5.1
<#
.SYNOPSIS
    Creates Windows shortcuts for OpenBOR games using OpenBOR-aware executable detection.

.DESCRIPTION
    Scans a folder of OpenBOR game directories and creates .lnk shortcuts.
    Uses a four-pass search strategy optimized for OpenBOR's typical structure:

      Pass 1: Look for OpenBOR.exe / openbor.exe in game root (fastest, most common)
      Pass 2: Look for any exe with "openbor" or "bor" in the name at root level
      Pass 3: Check common subfolders (bin, engine, data)
      Pass 4: Fall back to scoring all .exe files found

    Executable scoring bonuses:
      OpenBOR.exe (exact)     = 1000
      Known OpenBOR variants  = 500
      Contains "openbor"      = 200
      Located in game root    = 100
      Alternative names       = 50
      Too deep (>3 levels)    = -30 penalty

    Shortcut names are converted to title case and have the "OpenBOR - " prefix
    stripped if present. Special characters handled via -LiteralPath throughout.

.PARAMETER DryRun
    When $true (default), shows what would be created without writing shortcuts.
    Set to $false in the CONFIG block to apply.

.EXAMPLE
    .\Create-OpenBOR-Shortcuts.ps1

.VERSION
    2.0.0 - Config block, DryRun, CSV log, MIT header. Removed emoji and verbose
            debug output. LiteralPath used throughout. Logic preserved.
    1.0.0 - Initial release.

.LICENSE
    MIT License
    Copyright (c) Paul Mardis
#>

# ==============================================================================
# CONFIG -- Edit this block. Do not put paths anywhere else in this script.
# ==============================================================================
$RootGameFolder   = "D:\Arcade\System roms\OpenBOR"
$ShortcutFolder   = "D:\Arcade\System roms\PC Games\Shortcuts\OpenBOR"
$MaxDepth         = 3
$FolderTimeoutSec = 30
$DryRun           = $true    # Set to $false to write shortcuts
$LogDir           = Join-Path (Split-Path $ShortcutFolder -Parent) "Logs"
# ==============================================================================

function Set-ExitCode {
    param([int]$Code)
    $global:LASTEXITCODE = $Code
}

if (-not (Test-Path -LiteralPath $RootGameFolder)) {
    Write-Host "ERROR: OpenBOR library not found: $RootGameFolder" -ForegroundColor Red
    Set-ExitCode 1
    return
}

if (-not $DryRun) {
    foreach ($D in @($ShortcutFolder, $LogDir)) {
        if (-not (Test-Path -LiteralPath $D)) { New-Item -ItemType Directory -Path $D -Force | Out-Null }
    }
}

$LogFile = Join-Path $LogDir ("Create-OpenBOR-Shortcuts_{0}.csv" -f (Get-Date -Format "yyyyMMdd_HHmmss"))
$LogRows = [System.Collections.Generic.List[PSCustomObject]]::new()

function Write-LogRow {
    param([string]$GameName,[string]$ChosenExe,[string]$ShortcutPath,[string]$Status,[string]$Note)
    $LogRows.Add([PSCustomObject]@{
        Timestamp    = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
        GameName     = $GameName
        ChosenExe    = $ChosenExe
        ShortcutPath = $ShortcutPath
        Status       = $Status
        Note         = $Note
    })
}

# ------------------------------------------------------------------------------
# OpenBOR-specific executable scoring
# ------------------------------------------------------------------------------
function Get-OpenBORScore {
    param([string]$ExePath, [string]$GameFolderName, [string]$GameFolderPath)

    $ExeName = [System.IO.Path]::GetFileNameWithoutExtension($ExePath).ToLower()
    $Score   = 0

    $BlockPatterns = @("*uninstall*","*setup*","*install*","*update*","*config*","*settings*","*helper*","*crash*","*test*","*service*","*server*","*redist*")
    foreach ($P in $BlockPatterns) {
        if ($ExeName -like $P) { return -1 }
    }

    if ($ExeName -eq "openbor")                                { return 1000 }

    $Variants = @("openbor-win","openbor64","openbor32","openbor_win","openborengine","openbor-win64")
    if ($Variants -contains $ExeName)                          { return 500  }

    if ($ExeName -like "*openbor*")                            { $Score += 200 }

    # Root folder bonus
    $ExeDir = [System.IO.Path]::GetDirectoryName($ExePath).ToLower()
    if ($ExeDir -eq $GameFolderPath.ToLower())                 { $Score += 100 }

    $AltNames = @("bor","beatsofrage","game","start","play","run","main")
    if ($AltNames -contains $ExeName)                          { $Score += 50  }

    # Depth penalty
    $Depth = ($ExePath.Split('\').Count - $RootGameFolder.Split('\').Count)
    if ($Depth -gt 3)                                          { $Score -= 30  }

    return $Score
}

# ------------------------------------------------------------------------------
# Title-case shortcut name with prefix stripping
# ------------------------------------------------------------------------------
function Get-ShortcutName {
    param([string]$FolderName)

    $Name = $FolderName.Trim()

    # Strip known prefixes
    foreach ($Prefix in @("OpenBOR - ","BOR - ")) {
        if ($Name -like "$Prefix*") { $Name = $Name.Substring($Prefix.Length).Trim() }
    }

    # Title case
    $TextInfo = (Get-Culture).TextInfo
    $Name = $TextInfo.ToTitleCase($Name.ToLower())

    # Fix Roman numerals
    $Name = [regex]::Replace($Name, '\b(I{1,3}|IV|V|VI{1,3}|IX|X)\b', { $args[0].Value.ToUpper() })

    # Sanitize for Windows filename
    $Name = $Name -replace '[\\/\:\*\?"<>\|]', ''
    $Name = $Name -replace '&', 'and'
    if ($Name.Length -gt 60) { $Name = $Name.Substring(0, 57).TrimEnd() + "..." }

    return $Name.Trim()
}

# ------------------------------------------------------------------------------
# Four-pass OpenBOR executable finder
# ------------------------------------------------------------------------------
function Find-OpenBORExe {
    param([string]$FolderPath, [string]$GameName)

    # Pass 1: exact OpenBOR.exe in root
    foreach ($Variant in @("OpenBOR.exe","openbor.exe","OPENBOR.EXE")) {
        $TestPath = [System.IO.Path]::Combine($FolderPath, $Variant)
        if (Test-Path -LiteralPath $TestPath -PathType Leaf) {
            return @(Get-Item -LiteralPath $TestPath)
        }
    }

    # Pass 2: any openbor/bor exe at root level
    $RootExes = @(Get-ChildItem -LiteralPath $FolderPath -Filter "*.exe" -File -ErrorAction SilentlyContinue)
    $BorExes  = $RootExes | Where-Object { $_.Name -like "*openbor*" -or $_.Name -like "*bor*" }
    if ($BorExes.Count -gt 0) { return $BorExes }

    # Pass 3: check common subfolders
    foreach ($Sub in @("bin","engine","OpenBOR","data")) {
        $SubPath = [System.IO.Path]::Combine($FolderPath, $Sub)
        if (Test-Path -LiteralPath $SubPath -PathType Container) {
            $SubExes = @(Get-ChildItem -LiteralPath $SubPath -Filter "*.exe" -File -ErrorAction SilentlyContinue |
                Where-Object { $_.Name -like "*openbor*" -or $_.Name -eq "OpenBOR.exe" })
            if ($SubExes.Count -gt 0) { return $SubExes }
        }
    }

    # Pass 4: all exe files for scoring
    return $RootExes
}

# ------------------------------------------------------------------------------
# Shortcut writer
# ------------------------------------------------------------------------------
$WshShell       = New-Object -ComObject WScript.Shell
$CreatedTargets = @{}

function Save-Shortcut {
    param([string]$TargetExe, [string]$ShortcutName, [string]$OutputFolder)

    if (-not (Test-Path -LiteralPath $TargetExe -PathType Leaf)) { return "error" }

    $ShortcutPath = [System.IO.Path]::Combine($OutputFolder, "$ShortcutName.lnk")

    if (Test-Path -LiteralPath $ShortcutPath) { return "exists_same" }
    if ($CreatedTargets.ContainsKey($TargetExe)) { return "duplicate" }

    try {
        $SC = $WshShell.CreateShortcut($ShortcutPath)
        $SC.TargetPath       = $TargetExe
        $SC.WorkingDirectory = [System.IO.Path]::GetDirectoryName($TargetExe)
        $SC.Save()
        $CreatedTargets[$TargetExe] = $ShortcutName
        if (Test-Path -LiteralPath $ShortcutPath) { return "created" }
        return "error"
    } catch { return "error" }
}

# ==============================================================================
# BANNER
# ==============================================================================
Write-Host ""
Write-Host "Create-OpenBOR-Shortcuts" -ForegroundColor Cyan
Write-Host "========================" -ForegroundColor Cyan
Write-Host "  Library   : $RootGameFolder"
Write-Host "  Shortcuts : $ShortcutFolder"
Write-Host "  DryRun    : $DryRun"
Write-Host ""
if ($DryRun) {
    Write-Host "  [DRY RUN] No shortcuts will be written." -ForegroundColor Yellow
    Write-Host ""
}

# ==============================================================================
# MAIN LOOP
# ==============================================================================
$GameFolders = @(Get-ChildItem -LiteralPath $RootGameFolder -Directory -ErrorAction Stop) | Sort-Object Name
$Total       = $GameFolders.Count
$Index       = 0
$CntCreated  = 0
$CntSkipped  = 0
$CntNotFound = 0
$CntErrors   = 0

foreach ($GameDir in $GameFolders) {
    $Index++
    $GameName = $GameDir.Name

    Write-Progress -Activity "Create-OpenBOR-Shortcuts" -Status "$GameName ($Index of $Total)" `
        -PercentComplete ([int](($Index / $Total) * 100))

    try {
        $ExeFiles = @(Find-OpenBORExe -FolderPath $GameDir.FullName -GameName $GameName)

        if ($ExeFiles.Count -eq 0) {
            Write-Host ("  [NOT FOUND] {0}" -f $GameName) -ForegroundColor DarkGray
            Write-LogRow -GameName $GameName -Status "NotFound" -Note "No .exe files found"
            $CntNotFound++
            continue
        }

        $Scored = $ExeFiles | ForEach-Object {
            [PSCustomObject]@{
                Path  = $_.FullName
                Score = Get-OpenBORScore -ExePath $_.FullName -GameFolderName $GameName -GameFolderPath $GameDir.FullName
            }
        } | Where-Object { $_.Score -ge 0 } | Sort-Object Score -Descending

        if ($Scored.Count -eq 0) {
            Write-Host ("  [BLOCKED]   {0}" -f $GameName) -ForegroundColor DarkGray
            Write-LogRow -GameName $GameName -Status "AllBlocked" -Note "All .exe files blocked by filters"
            $CntNotFound++
            continue
        }

        $ChosenExe    = $Scored[0].Path
        $ShortcutName = Get-ShortcutName -FolderName $GameName
        $ShortcutOut  = [System.IO.Path]::Combine($ShortcutFolder, "$ShortcutName.lnk")

        if ($DryRun) {
            Write-Host ("  [WOULD CREATE] {0}" -f $ShortcutName) -ForegroundColor Cyan
            Write-Host ("    Exe   : {0}" -f $ChosenExe)
            Write-Host ("    Score : {0}" -f $Scored[0].Score)
            Write-LogRow -GameName $GameName -ChosenExe $ChosenExe -ShortcutPath $ShortcutOut `
                -Status "DryRun" -Note "Score=$($Scored[0].Score)"
            $CntCreated++
            continue
        }

        $Result = Save-Shortcut -TargetExe $ChosenExe -ShortcutName $ShortcutName -OutputFolder $ShortcutFolder

        switch ($Result) {
            "created"     { Write-Host ("  [OK]     {0}" -f $ShortcutName) -ForegroundColor Green;    $CntCreated++  }
            "exists_same" { Write-Host ("  [EXISTS] {0}" -f $ShortcutName) -ForegroundColor DarkGray; $CntSkipped++  }
            "duplicate"   { Write-Host ("  [DUPE]   {0}" -f $ShortcutName) -ForegroundColor DarkGray; $CntSkipped++  }
            default       { Write-Host ("  [ERROR]  {0}" -f $ShortcutName) -ForegroundColor Red;      $CntErrors++   }
        }
        Write-LogRow -GameName $GameName -ChosenExe $ChosenExe -ShortcutPath $ShortcutOut `
            -Status $Result -Note "Score=$($Scored[0].Score)"
    }
    catch {
        Write-Host ("  [ERROR]  {0} -- {1}" -f $GameName, $_.Exception.Message) -ForegroundColor Red
        Write-LogRow -GameName $GameName -Status "Error" -Note $_.Exception.Message
        $CntErrors++
    }
}

Write-Progress -Activity "Create-OpenBOR-Shortcuts" -Completed

if ($LogRows.Count -gt 0 -and -not $DryRun) {
    try { $LogRows | Export-Csv -LiteralPath $LogFile -NoTypeInformation -Encoding UTF8 } catch {}
    Write-Host ""
    Write-Host ("  Log: {0}" -f $LogFile) -ForegroundColor Gray
}

Write-Host ""
Write-Host "Summary" -ForegroundColor White
Write-Host ("  {0,-12} {1}" -f "Created :", $CntCreated)  -ForegroundColor Green
Write-Host ("  {0,-12} {1}" -f "Skipped :", $CntSkipped)  -ForegroundColor Gray
Write-Host ("  {0,-12} {1}" -f "Not found:", $CntNotFound) -ForegroundColor DarkGray
Write-Host ("  {0,-12} {1}" -f "Errors :", $CntErrors)    -ForegroundColor $(if ($CntErrors -gt 0) { "Red" } else { "Gray" })
if ($DryRun) {
    Write-Host ""
    Write-Host "  DRY RUN complete. Set `$DryRun = `$false in the CONFIG block to apply." -ForegroundColor Yellow
}
Write-Host ""
Set-ExitCode 0
