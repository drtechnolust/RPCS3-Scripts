#Requires -Version 5.1
<#
.SYNOPSIS
    Creates Windows shortcuts for GOG games with GOG-aware executable detection.

.DESCRIPTION
    Scans a GOG game library folder and creates .lnk shortcuts optimized for
    GOG's DRM-free structure. Lighter filtering than the generic PC shortcut
    creator since GOG games rarely have aggressive folder nesting.

    GOG-specific scoring bonuses:
      - Exact folder name match                      = 1000
      - GOGLauncherFile.exe detected                 = 950
      - Name matches folder name                     = 900
      - DOSBox / ScummVM games handled               = high priority
      - Located in bin\ or root of game folder       = bonus

    Blocked patterns (lighter than generic):
      uninstall, setup, config, crash, update, redist, directx

.PARAMETER DryRun
    When $true (default), shows what would be created without writing shortcuts.
    Set to $false in the CONFIG block to apply.

.EXAMPLE
    .\Create-GOG-Shortcuts.ps1

.VERSION
    2.0.0 - Config block, DryRun, CSV log, MIT header, fixed author placeholder.
            Scoring and detection logic preserved.
    1.0.0 - Initial release.

.LICENSE
    MIT License
    Copyright (c) Paul Mardis
#>

# ==============================================================================
# CONFIG -- Edit this block. Do not put paths anywhere else in this script.
# ==============================================================================
$RootGameFolder   = "D:\Arcade\System roms\PC Games"    # Folder containing GOG game subfolders
$ShortcutFolder   = "D:\Arcade\System roms\PC Games\Shortcuts\GOG"
$MaxDepth         = 6
$FolderTimeoutSec = 180
$DryRun           = $true    # Set to $false to write shortcuts
$LogDir           = Join-Path (Split-Path $ShortcutFolder -Parent) "Logs"
# ==============================================================================

function Set-ExitCode {
    param([int]$Code)
    $global:LASTEXITCODE = $Code
}

if (-not (Test-Path -LiteralPath $RootGameFolder)) {
    Write-Host "ERROR: Game library not found: $RootGameFolder" -ForegroundColor Red
    Set-ExitCode 1
    return
}

if (-not $DryRun) {
    foreach ($D in @($ShortcutFolder, $LogDir)) {
        if (-not (Test-Path -LiteralPath $D)) { New-Item -ItemType Directory -Path $D -Force | Out-Null }
    }
}

$LogFile = Join-Path $LogDir ("Create-GOG-Shortcuts_{0}.csv" -f (Get-Date -Format "yyyyMMdd_HHmmss"))
$LogRows = [System.Collections.Generic.List[PSCustomObject]]::new()

function Write-LogRow {
    param([string]$GameName,[string]$ChosenExe,[string]$ShortcutPath,[string]$Status,[string]$Note)
    $LogRows.Add([PSCustomObject]@{
        Timestamp   = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
        GameName    = $GameName
        ChosenExe   = $ChosenExe
        ShortcutPath = $ShortcutPath
        Status      = $Status
        Note        = $Note
    })
}

# ------------------------------------------------------------------------------
# GOG-optimized executable scoring
# ------------------------------------------------------------------------------
function Get-GOGExecutableScore {
    param([string]$ExePath, [string]$GameFolderName, [string]$FolderPath)

    $ExeName    = [System.IO.Path]::GetFileNameWithoutExtension($ExePath).ToLower()
    $FolderName = $GameFolderName.ToLower()
    $Score      = 0

    # Block utility patterns (lighter than generic PC list)
    $BlockPatterns = @("*uninstall*","*setup*","*config*","*crash*","*update*","*redist*","*directx*")
    foreach ($P in $BlockPatterns) {
        if ($ExeName -like $P) { return -1 }
    }

    # GOGLauncher is the definitive GOG launcher
    if ($ExeName -eq "goglauncherfile" -or $ExeName -like "*goglauncher*") { return 950 }

    # Exact folder name match
    if ($ExeName -eq $FolderName) { return 1000 }

    # Single word match from folder name
    $Parts = $FolderName -split '[_ ]'
    foreach ($Part in $Parts) {
        if ($Part.Length -gt 3 -and $ExeName -eq $Part) { return 900 }
    }

    # DOSBox / ScummVM wrappers
    if ($ExeName -like "*dosbox*") { $Score += 100 }
    if ($ExeName -like "*scummvm*") { $Score += 100 }

    # Partial folder name match
    if ($ExeName -like "*$FolderName*") { $Score += 50 }
    foreach ($Part in $Parts) {
        if ($Part.Length -gt 3 -and $ExeName -like "*$Part*") { $Score += 15 }
    }

    # Prefer root level and bin
    $ExeDir = [System.IO.Path]::GetDirectoryName($ExePath).ToLower()
    $GoodPaths = @("\bin","\game","\app")
    foreach ($GP in $GoodPaths) {
        if ($ExeDir -like "*$GP*") { $Score += 20; break }
    }

    # Check if at root of game folder (GOG games are often flat)
    $GameRoot = $FolderPath.ToLower()
    if ($ExeDir -eq $GameRoot) { $Score += 30 }

    # Priority generic names
    if (@("start","play","run","main","launcher") -contains $ExeName) { $Score += 25 }

    return $Score
}

# ------------------------------------------------------------------------------
# Shortcut writer
# ------------------------------------------------------------------------------
$WshShell      = New-Object -ComObject WScript.Shell
$CreatedTargets = @{}

function Save-Shortcut {
    param([string]$TargetExe, [string]$GameName, [string]$OutputFolder)

    $SafeName = ($GameName -replace '[_]', ' ').Trim()
    $SafeName = $SafeName -replace '[\\/\:\*\?"<>\|]', ''
    if ($SafeName.Length -gt 60) { $SafeName = $SafeName.Substring(0, 60).TrimEnd() }

    $ShortcutPath = Join-Path $OutputFolder "$SafeName.lnk"

    if (Test-Path -LiteralPath $ShortcutPath) {
        try {
            $E = $WshShell.CreateShortcut($ShortcutPath)
            if ($E.TargetPath -eq $TargetExe) { return "exists_same" }
            return "exists_different"
        } catch { return "error" }
    }

    if ($CreatedTargets.ContainsKey($TargetExe)) { return "duplicate" }

    try {
        $SC = $WshShell.CreateShortcut($ShortcutPath)
        $SC.TargetPath       = $TargetExe
        $SC.WorkingDirectory = [System.IO.Path]::GetDirectoryName($TargetExe)
        $SC.Save()
        $CreatedTargets[$TargetExe] = $GameName
        return "created"
    } catch { return "error" }
}

# ==============================================================================
# BANNER
# ==============================================================================
Write-Host ""
Write-Host "Create-GOG-Shortcuts" -ForegroundColor Cyan
Write-Host "====================" -ForegroundColor Cyan
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
$GameFolders = Get-ChildItem -LiteralPath $RootGameFolder -Directory | Sort-Object Name
$Total = $GameFolders.Count
$Index = 0

$CntCreated  = 0
$CntSkipped  = 0
$CntNotFound = 0
$CntErrors   = 0

foreach ($GameDir in $GameFolders) {
    $Index++
    $GameName = $GameDir.Name

    Write-Progress -Activity "Create-GOG-Shortcuts" -Status "$GameName ($Index of $Total)" `
        -PercentComplete ([int](($Index / $Total) * 100))

    try {
        $ExeFiles = @(Get-ChildItem -LiteralPath $GameDir.FullName -Filter "*.exe" -File `
            -Depth $MaxDepth -ErrorAction SilentlyContinue)

        if ($ExeFiles.Count -eq 0) {
            Write-Host ("  [NOT FOUND] {0}" -f $GameName) -ForegroundColor DarkGray
            Write-LogRow -GameName $GameName -Status "NotFound" -Note "No .exe files found"
            $CntNotFound++
            continue
        }

        $Scored = $ExeFiles | ForEach-Object {
            [PSCustomObject]@{
                Path  = $_.FullName
                Score = Get-GOGExecutableScore -ExePath $_.FullName -GameFolderName $GameName -FolderPath $GameDir.FullName
            }
        } | Where-Object { $_.Score -ge 0 } | Sort-Object Score -Descending

        if ($Scored.Count -eq 0) {
            Write-Host ("  [BLOCKED]   {0}" -f $GameName) -ForegroundColor DarkGray
            Write-LogRow -GameName $GameName -Status "AllBlocked" -Note "All .exe files blocked by filters"
            $CntNotFound++
            continue
        }

        $ChosenExe   = $Scored[0].Path
        $ShortcutOut = Join-Path $ShortcutFolder "$GameName.lnk"

        if ($DryRun) {
            Write-Host ("  [WOULD CREATE] {0}" -f $GameName) -ForegroundColor Cyan
            Write-Host ("    Exe: {0}" -f $ChosenExe)
            Write-LogRow -GameName $GameName -ChosenExe $ChosenExe -ShortcutPath $ShortcutOut `
                -Status "DryRun" -Note "Score=$($Scored[0].Score)"
            $CntCreated++
            continue
        }

        $Result = Save-Shortcut -TargetExe $ChosenExe -GameName $GameName -OutputFolder $ShortcutFolder

        switch ($Result) {
            "created"          { Write-Host ("  [OK]     {0}" -f $GameName) -ForegroundColor Green;    $CntCreated++ }
            "exists_same"      { Write-Host ("  [EXISTS] {0}" -f $GameName) -ForegroundColor DarkGray; $CntSkipped++ }
            "exists_different" { Write-Host ("  [UPDATED]{0}" -f $GameName) -ForegroundColor Cyan;     $CntSkipped++ }
            "duplicate"        { Write-Host ("  [DUPE]   {0}" -f $GameName) -ForegroundColor DarkGray; $CntSkipped++ }
            default            { Write-Host ("  [ERROR]  {0}" -f $GameName) -ForegroundColor Red;      $CntErrors++  }
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

Write-Progress -Activity "Create-GOG-Shortcuts" -Completed

if ($LogRows.Count -gt 0 -and -not $DryRun) {
    try { $LogRows | Export-Csv -LiteralPath $LogFile -NoTypeInformation -Encoding UTF8 } catch {}
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
