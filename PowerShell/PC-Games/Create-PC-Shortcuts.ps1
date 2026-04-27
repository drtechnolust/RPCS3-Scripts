#Requires -Version 5.1
<#
.SYNOPSIS
    Creates Windows shortcuts for PC games organized by detected platform.

.DESCRIPTION
    Scans a PC game library folder, detects the platform (Steam, GOG, Epic, Retail)
    for each game by inspecting DLL and file signatures, scores all .exe files found
    to choose the most likely game launcher, and creates a .lnk shortcut in the
    appropriate platform subfolder.

    Platform subfolders created under ShortcutFolder:
      Steam\    GOG\    Epic\    Retail\    PC_Games\

    Executable scoring rules (highest score wins):
      - Exact name match with game folder          = 1000
      - Single-word folder name match              = 900
      - Unreal Engine *-Win64-Shipping pattern     = 500
      - Xbox/WinGDK shipping pattern               = 400+
      - Contains "game" in exe name                = 75
      - Partial folder name match                  = 50+
      - Located in known binary subfolder          = +20
      - Blocked names (uninstall, setup, crash...) = -1 (excluded)

    Special case: ELDEN RING is handled via a direct known path.

    Run with DryRun = $true first to preview shortcut decisions without writing files.
    A single CSV log replaces the five separate text log files from earlier versions.

.PARAMETER DryRun
    When $true (default), shows what would be created without writing any shortcuts.
    Set to $false in the CONFIG block to apply.

.EXAMPLE
    .\Create-PC-Shortcuts.ps1

.NOTES
    The recursive fallback search uses Start-Job for timeout protection.
    This works in both PS 5.1 and PS 7. If scanning a very large library
    on a slow drive, increase FolderTimeoutSec in the CONFIG block.

.VERSION
    3.0.0 - Config block, DryRun, CSV log, MIT header, fixed author placeholder,
            removed emoji from output, fixed shortcut output folder path.
            All platform detection and scoring logic preserved from v2.0.
    2.0.0 - Added platform detection (Steam/GOG/Epic/Retail), platform subfolders,
            WinGDK/Xbox Game Studio support, progressive depth search.
    1.0.0 - Initial release.

.LICENSE
    MIT License
    Copyright (c) Paul Mardis
#>

# ==============================================================================
# CONFIG -- Edit this block. Do not put paths anywhere else in this script.
# ==============================================================================
$RootGameFolder   = "D:\Arcade\System roms\PC Games"
$ShortcutFolder   = "D:\Arcade\System roms\PC Games\Shortcuts"
$MaxDepth         = 8
$FolderTimeoutSec = 300     # Seconds before giving up on a single game folder
$DryRun           = $true   # Set to $false to write shortcuts
$LogDir           = Join-Path $ShortcutFolder "Logs"
# ==============================================================================

function Set-ExitCode {
    param([int]$Code)
    $global:LASTEXITCODE = $Code
}

# ------------------------------------------------------------------------------
# Pre-flight
# ------------------------------------------------------------------------------
if (-not (Test-Path -LiteralPath $RootGameFolder)) {
    Write-Host "ERROR: Game library not found: $RootGameFolder" -ForegroundColor Red
    Set-ExitCode 1
    return
}

# ------------------------------------------------------------------------------
# Platform subfolder map
# ------------------------------------------------------------------------------
$PlatformFolders = @{
    Steam    = Join-Path $ShortcutFolder "Steam"
    GOG      = Join-Path $ShortcutFolder "GOG"
    Epic     = Join-Path $ShortcutFolder "Epic"
    Retail   = Join-Path $ShortcutFolder "Retail"
    PC_Games = Join-Path $ShortcutFolder "PC_Games"
}

# ------------------------------------------------------------------------------
# Setup output folders
# ------------------------------------------------------------------------------
if (-not $DryRun) {
    foreach ($Folder in ($PlatformFolders.Values + @($LogDir))) {
        if (-not (Test-Path -LiteralPath $Folder)) {
            New-Item -ItemType Directory -Path $Folder -Force | Out-Null
        }
    }
}

$LogFile = Join-Path $LogDir ("Create-PC-Shortcuts_{0}.csv" -f (Get-Date -Format "yyyyMMdd_HHmmss"))
$LogRows = [System.Collections.Generic.List[PSCustomObject]]::new()

function Write-LogRow {
    param(
        [string]$GameName,
        [string]$Platform,
        [string]$ChosenExe,
        [string]$ShortcutPath,
        [string]$Action,
        [string]$Status,
        [string]$Note
    )
    $LogRows.Add([PSCustomObject]@{
        Timestamp   = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
        GameName    = $GameName
        Platform    = $Platform
        ChosenExe   = $ChosenExe
        ShortcutPath = $ShortcutPath
        Action      = $Action
        Status      = $Status
        Note        = $Note
    })
}

# ------------------------------------------------------------------------------
# Platform detection -- inspects DLLs and files in the game folder
# ------------------------------------------------------------------------------
function Get-GamePlatform {
    param([string]$FolderPath)

    $scores = @{ Steam = 0; GOG = 0; Epic = 0; Retail = 0 }

    try {
        $Files = Get-ChildItem -LiteralPath $FolderPath -File -Recurse -Depth 2 -ErrorAction SilentlyContinue
        foreach ($File in $Files) {
            $n = $File.Name.ToLower()
            if ($n -eq "steam_api.dll" -or $n -eq "steam_api64.dll") { $scores.Steam += 10 }
            if ($n -eq "steam_appid.txt")                              { $scores.Steam += 15 }
            if ($n -eq "steamclient_loader.exe")                       { $scores.Steam += 8  }
            if ($n -like "*steam*")                                    { $scores.Steam += 2  }
            if ($n -eq "galaxy.dll" -or $n -eq "galaxy64.dll")        { $scores.GOG   += 8  }
            if ($n -like "*gog*" -or $n -like "*galaxy*")             { $scores.GOG   += 3  }
            if ($n -like "*gogdos*")                                   { $scores.GOG   += 5  }
            if ($n -like "*epic*" -or $n -like "*egs*")               { $scores.Epic  += 3  }
            if ($n -like "*eosoverlay*" -or $n -like "*epiconlineservices*") { $scores.Epic += 8 }
            if ($n -eq "easyanticheat_eos_setup.exe")                  { $scores.Epic  += 5  }
            if ($n -like "*cd*key*" -or $n -like "*serial*")          { $scores.Retail += 5 }
            if ($n -like "*manual*" -and $File.Extension -eq ".pdf")  { $scores.Retail += 3 }
        }

        $SubFolders = Get-ChildItem -LiteralPath $FolderPath -Directory -ErrorAction SilentlyContinue
        foreach ($Sub in $SubFolders) {
            $s = $Sub.Name.ToLower()
            if ($s -like "*crack*" -or $s -like "*nodvd*")          { $scores.Retail += 4 }
            if ($s -like "*goldberg*" -or $s -like "*steamemu*")     { $scores.Steam  += 6 }
            if ($s -like "*redist*")                                  { $scores.Retail += 2 }
        }

        $MaxScore = ($scores.Values | Measure-Object -Maximum).Maximum
        if ($MaxScore -gt 0) {
            return ($scores.GetEnumerator() | Where-Object { $_.Value -eq $MaxScore } | Select-Object -First 1).Key
        }
    } catch { }

    return "PC_Games"
}

# ------------------------------------------------------------------------------
# Executable scoring -- returns -1 to exclude, higher = better
# ------------------------------------------------------------------------------
function Get-ExecutableScore {
    param([string]$ExePath, [string]$GameFolderName)

    $ExeName    = [System.IO.Path]::GetFileNameWithoutExtension($ExePath).ToLower()
    $FolderName = $GameFolderName.ToLower()
    $Score      = 0

    # Block patterns (anywhere in exe name)
    $BlockPatterns = @(
        "*uninstall*","*setup*","*settings*","*helper*","*config*",
        "*language*","*crash*","*test*","*service*","*server*",
        "*update*","*install*"
    )
    foreach ($Pattern in $BlockPatterns) {
        if ($ExeName -like $Pattern) { return -1 }
    }

    # Block exact names
    $BlockNames = @(
        "unins000","crashreport","errorreport","crashreporter","crashpad",
        "redist","redistributable","vcredist","vc_redist","directx","dxwebsetup",
        "uploader","webhelper","crs-handler","crs-uploader","crs-video",
        "drivepool","quicksfv","handler","gamingrepair","unitycrashhandle64"
    )
    if ($BlockNames -contains $ExeName) { return -1 }

    # Exact folder name match
    if ($ExeName -eq $FolderName) { return 1000 }

    # Single word from folder name
    $FolderParts = $FolderName -split ' '
    foreach ($Part in $FolderParts) {
        if ($Part.Length -gt 3 -and $ExeName -eq $Part) { return 900 }
    }

    # Unreal Engine Win64-Shipping
    if ($ExeName -eq "$FolderName-win64-shipping") { return 500 }

    # Xbox/WinGDK shipping
    if ($ExeName -like "*wingdk*shipping*") { $Score += 400 }
    if ($ExeName -like "*$FolderName*" -and $ExeName -like "*win64*shipping*") { $Score += 300 }

    # Contains "game"
    if ($ExeName -like "*game*") { $Score += 75 }

    # Partial folder name match
    if ($ExeName -like "*$FolderName*") { $Score += 50 }

    # Individual word matches (min 4 chars)
    foreach ($Part in $FolderParts) {
        if ($Part.Length -gt 3 -and $ExeName -like "*$Part*") { $Score += 20 }
    }

    # Priority generic names
    $PriorityNames = @("start","play","run","main","bin")
    if ($PriorityNames -contains $ExeName) { $Score += 30 }

    # Located in known binary path
    $ExeDir    = [System.IO.Path]::GetDirectoryName($ExePath).ToLower()
    $GoodPaths = @("\bin","\binaries","\game","\app","\win64","\win32","\windows","\x64","\x86","\wingdk")
    foreach ($GoodPath in $GoodPaths) {
        if ($ExeDir -like "*$GoodPath*") { $Score += 20; break }
    }

    # Depth penalty
    $Depth = ($ExePath.Split('\').Count - $RootGameFolder.Split('\').Count)
    if ($Depth -gt 6) { $Score -= 10 }

    return $Score
}

# ------------------------------------------------------------------------------
# Find executables -- prioritises known paths, falls back to progressive scan
# ------------------------------------------------------------------------------
function Find-Executables {
    param([string]$FolderPath, [string]$GameName, [int]$MaxDepthParam, [int]$TimeoutSec)

    $ExeFiles = @()

    # Check common paths first (fast, avoids full recursive scan)
    $CommonPaths = @(
        "$FolderPath\*.exe"
        "$FolderPath\binaries_x64\*.exe"
        "$FolderPath\Binaries\WinGDK\*.exe"
        "$FolderPath\Game\*.exe"
        "$FolderPath\app\*.exe"
        "$FolderPath\bin\*.exe"
        "$FolderPath\binaries\*.exe"
        "$FolderPath\Windows\*.exe"
        "$FolderPath\x64\*.exe"
        "$FolderPath\Win64\*.exe"
    )

    foreach ($CPath in $CommonPaths) {
        if (Test-Path $CPath) {
            $Found = Get-ChildItem -Path $CPath -ErrorAction SilentlyContinue
            if ($Found) { $ExeFiles += $Found }
        }
    }

    # Deeper common patterns (only if nothing found yet)
    if ($ExeFiles.Count -eq 0) {
        $DeeperPaths = @(
            "$FolderPath\Binaries\Win64\*.exe"
            "$FolderPath\*\Binaries\Win64\*.exe"
            "$FolderPath\*\*\Binaries\Win64\*.exe"
            "$FolderPath\*\Binaries\WinGDK\*.exe"
            "$FolderPath\binaries_x86\*.exe"
        )
        foreach ($DPath in $DeeperPaths) {
            if ($ExeFiles.Count -eq 0) {
                $Found = Get-ChildItem -Path $DPath -ErrorAction SilentlyContinue
                if ($Found) { $ExeFiles += $Found }
            }
        }
    }

    if ($ExeFiles.Count -gt 0) { return $ExeFiles }

    # Progressive depth scan
    foreach ($D in @(0, 1, 2, 4, $MaxDepthParam)) {
        $ExeFiles = @(Get-ChildItem -LiteralPath $FolderPath -Filter "*.exe" -File -Depth $D -ErrorAction SilentlyContinue)
        if ($ExeFiles.Count -gt 0) { return $ExeFiles }
    }

    # Full recursive with timeout via background job
    $Job = Start-Job -ScriptBlock {
        param($Path)
        Get-ChildItem -Path $Path -Filter "*.exe" -File -Recurse -ErrorAction SilentlyContinue
    } -ArgumentList $FolderPath

    $null = Wait-Job -Job $Job -Timeout ($TimeoutSec - 5)

    if ($Job.State -eq "Running") {
        Stop-Job -Job $Job
        Write-Host ("    TIMEOUT scanning {0}" -f $GameName) -ForegroundColor Yellow
    } else {
        $ExeFiles = @(Receive-Job -Job $Job)
    }
    Remove-Job -Job $Job -Force
    return $ExeFiles
}

# ------------------------------------------------------------------------------
# Create shortcut (with sanitized-name fallback)
# ------------------------------------------------------------------------------
$WshShell        = New-Object -ComObject WScript.Shell
$CreatedTargets  = @{}   # tracks exe -> shortcut name to detect duplicates

function Save-GameShortcut {
    param([string]$TargetExe, [string]$GameName, [string]$OutputFolder)

    # Sanitize name
    $SafeName = $GameName -replace '[\\/\:\*\?"<>\|]', '_'
    $SafeName = $SafeName -replace '&', 'and'
    if ($SafeName.Length -gt 50) { $SafeName = $SafeName.Substring(0, 47) + "..." }

    $ShortcutPath = Join-Path $OutputFolder "$SafeName.lnk"

    # Already exists pointing to same target
    if (Test-Path -LiteralPath $ShortcutPath) {
        try {
            $Existing = $WshShell.CreateShortcut($ShortcutPath)
            if ($Existing.TargetPath -eq $TargetExe) { return "exists_same" }
            return "exists_different"
        } catch { return "error" }
    }

    # Duplicate target already used for another game
    if ($CreatedTargets.ContainsKey($TargetExe)) { return "duplicate" }

    try {
        $SC = $WshShell.CreateShortcut($ShortcutPath)
        $SC.TargetPath       = $TargetExe
        $SC.WorkingDirectory = [System.IO.Path]::GetDirectoryName($TargetExe)
        $SC.Save()
        $CreatedTargets[$TargetExe] = $GameName
        return "created"
    } catch {
        # Fallback: ultra-safe ASCII name
        try {
            $FallbackName = "Game_" + ($GameName -replace '[^a-zA-Z0-9]', '_')
            if ($FallbackName.Length -gt 30) { $FallbackName = $FallbackName.Substring(0, 30) }
            $FallbackPath = Join-Path $OutputFolder "$FallbackName.lnk"
            $SC2 = $WshShell.CreateShortcut($FallbackPath)
            $SC2.TargetPath       = $TargetExe
            $SC2.WorkingDirectory = [System.IO.Path]::GetDirectoryName($TargetExe)
            $SC2.Save()
            $CreatedTargets[$TargetExe] = $GameName
            return "created_sanitized"
        } catch { return "error" }
    }
}

# ==============================================================================
# BANNER
# ==============================================================================
Write-Host ""
Write-Host "Create-PC-Shortcuts" -ForegroundColor Cyan
Write-Host "===================" -ForegroundColor Cyan
Write-Host "  Library   : $RootGameFolder"
Write-Host "  Shortcuts : $ShortcutFolder"
Write-Host "  MaxDepth  : $MaxDepth"
Write-Host "  Timeout   : ${FolderTimeoutSec}s per game"
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
$Total       = $GameFolders.Count
$Index       = 0
$StartTime   = Get-Date

$CntCreated   = 0
$CntSanitized = 0
$CntSkipped   = 0
$CntNotFound  = 0
$CntErrors    = 0
$CntTimeouts  = 0

foreach ($GameDir in $GameFolders) {
    $Index++
    $GameName = $GameDir.Name

    $Pct = [int](($Index / $Total) * 100)
    Write-Progress -Activity "Create-PC-Shortcuts" -Status "$GameName ($Index of $Total)" -PercentComplete $Pct

    # Special-case known games with non-standard executable locations
    $ManualExe = $null
    if ($GameName -eq "ELDEN RING") {
        $SpecificPath = Join-Path $GameDir.FullName "Game\eldenring.exe"
        if (Test-Path -LiteralPath $SpecificPath) { $ManualExe = $SpecificPath }
    }

    try {
        $ExeFiles = @()
        $TimedOut = $false

        if ($ManualExe) {
            $ExeFiles = @(Get-Item -LiteralPath $ManualExe)
        } else {
            $SearchStart = Get-Date
            $ExeFiles    = @(Find-Executables -FolderPath $GameDir.FullName -GameName $GameName -MaxDepthParam $MaxDepth -TimeoutSec $FolderTimeoutSec)
            $SearchSecs  = ((Get-Date) - $SearchStart).TotalSeconds

            if ($SearchSecs -gt ($FolderTimeoutSec * 0.9)) {
                $TimedOut = $true
                $CntTimeouts++
                Write-Host ("  [TIMEOUT] {0} ({1:F0}s)" -f $GameName, $SearchSecs) -ForegroundColor Yellow
            }
        }

        if ($ExeFiles.Count -eq 0) {
            Write-Host ("  [NOT FOUND] {0}" -f $GameName) -ForegroundColor DarkGray
            Write-LogRow -GameName $GameName -Action "Skip" -Status "NotFound" -Note "No .exe files found"
            $CntNotFound++
            continue
        }

        # Score and filter
        $Scored = $ExeFiles | ForEach-Object {
            [PSCustomObject]@{
                Path  = $_.FullName
                Score = Get-ExecutableScore -ExePath $_.FullName -GameFolderName $GameName
            }
        } | Where-Object { $_.Score -ge 0 } | Sort-Object Score -Descending

        if ($Scored.Count -eq 0) {
            Write-Host ("  [BLOCKED]   {0}" -f $GameName) -ForegroundColor DarkGray
            Write-LogRow -GameName $GameName -Action "Skip" -Status "AllBlocked" -Note "All .exe files blocked by filters"
            $CntNotFound++
            continue
        }

        $ChosenExe = $Scored[0].Path
        $Platform  = Get-GamePlatform -FolderPath $GameDir.FullName
        $OutFolder = $PlatformFolders[$Platform]
        $ScInfo    = "Score=$($Scored[0].Score)"

        if ($DryRun) {
            Write-Host ("  [WOULD CREATE] {0,-50} [{1}]" -f $GameName, $Platform) -ForegroundColor Cyan
            Write-Host ("    Exe: {0}" -f $ChosenExe)
            Write-LogRow -GameName $GameName -Platform $Platform -ChosenExe $ChosenExe `
                -ShortcutPath (Join-Path $OutFolder "$GameName.lnk") `
                -Action "Create" -Status "DryRun" -Note $ScInfo
            $CntCreated++
            continue
        }

        $Result = Save-GameShortcut -TargetExe $ChosenExe -GameName $GameName -OutputFolder $OutFolder

        switch ($Result) {
            "created" {
                Write-Host ("  [OK]        {0,-50} [{1}]" -f $GameName, $Platform) -ForegroundColor Green
                Write-LogRow -GameName $GameName -Platform $Platform -ChosenExe $ChosenExe `
                    -ShortcutPath (Join-Path $OutFolder "$GameName.lnk") `
                    -Action "Create" -Status "Success" -Note $ScInfo
                $CntCreated++
            }
            "created_sanitized" {
                Write-Host ("  [SANITIZED] {0,-50} [{1}]" -f $GameName, $Platform) -ForegroundColor Yellow
                Write-LogRow -GameName $GameName -Platform $Platform -ChosenExe $ChosenExe `
                    -ShortcutPath (Join-Path $OutFolder "$GameName.lnk") `
                    -Action "Create" -Status "Sanitized" -Note "$ScInfo -- name sanitized"
                $CntSanitized++
            }
            "exists_same" {
                Write-Host ("  [EXISTS]    {0,-50} [{1}]" -f $GameName, $Platform) -ForegroundColor DarkGray
                Write-LogRow -GameName $GameName -Platform $Platform -ChosenExe $ChosenExe `
                    -Action "Skip" -Status "ExistsSame" -Note "Shortcut already correct"
                $CntSkipped++
            }
            "exists_different" {
                Write-Host ("  [UPDATED]   {0,-50} [{1}]" -f $GameName, $Platform) -ForegroundColor Cyan
                Write-LogRow -GameName $GameName -Platform $Platform -ChosenExe $ChosenExe `
                    -Action "Update" -Status "ExistsDifferent" -Note "Shortcut updated to new target"
                $CntSkipped++
            }
            "duplicate" {
                Write-Host ("  [DUPE]      {0,-50} [{1}]" -f $GameName, $Platform) -ForegroundColor DarkGray
                Write-LogRow -GameName $GameName -Platform $Platform -ChosenExe $ChosenExe `
                    -Action "Skip" -Status "Duplicate" -Note "Same exe already used by another shortcut"
                $CntSkipped++
            }
            default {
                Write-Host ("  [ERROR]     {0}" -f $GameName) -ForegroundColor Red
                Write-LogRow -GameName $GameName -Platform $Platform -ChosenExe $ChosenExe `
                    -Action "Create" -Status "Error" -Note "WScript.Shell failed"
                $CntErrors++
            }
        }
    }
    catch {
        Write-Host ("  [ERROR]     {0} -- {1}" -f $GameName, $_.Exception.Message) -ForegroundColor Red
        Write-LogRow -GameName $GameName -Action "Create" -Status "Error" -Note $_.Exception.Message
        $CntErrors++
    }
}

Write-Progress -Activity "Create-PC-Shortcuts" -Completed

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
$TotalTime = (Get-Date) - $StartTime
$TimeLabel = if ($TotalTime.TotalHours -ge 1) { "{0:h\:mm\:ss}" -f $TotalTime } else { "{0:mm\:ss}" -f $TotalTime }

Write-Host ""
Write-Host "Summary" -ForegroundColor White
Write-Host ("  {0,-14} {1}" -f "Created :", $CntCreated)   -ForegroundColor Green
Write-Host ("  {0,-14} {1}" -f "Sanitized :", $CntSanitized) -ForegroundColor Yellow
Write-Host ("  {0,-14} {1}" -f "Skipped :", $CntSkipped)   -ForegroundColor Gray
Write-Host ("  {0,-14} {1}" -f "Not found :", $CntNotFound) -ForegroundColor DarkGray
Write-Host ("  {0,-14} {1}" -f "Errors :", $CntErrors)     -ForegroundColor $(if ($CntErrors -gt 0) { "Red" } else { "Gray" })
Write-Host ("  {0,-14} {1}" -f "Timeouts :", $CntTimeouts) -ForegroundColor $(if ($CntTimeouts -gt 0) { "Yellow" } else { "Gray" })
Write-Host ("  {0,-14} {1}" -f "Total time :", $TimeLabel)
Write-Host ("  {0,-14} {1}" -f "Output :", $ShortcutFolder)

if ($DryRun) {
    Write-Host ""
    Write-Host "  DRY RUN complete. Set `$DryRun = `$false in the CONFIG block to apply." -ForegroundColor Yellow
}

Write-Host ""
Set-ExitCode 0
