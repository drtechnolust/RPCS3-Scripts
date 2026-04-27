#Requires -Version 5.1
<#
.SYNOPSIS
    Scans a PC game library and exports a CSV for import into LaunchBox.

.DESCRIPTION
    Scans a folder of PC game directories, identifies the most likely main
    executable for each game, and writes a CSV file that LaunchBox can import
    via: Import -> ROM Files or Folders -> From CSV File.

    CSV columns: Title, ApplicationPath, Platform

    Executable selection uses a priority list first (game.exe, start.exe, etc.)
    with folder-name matching as a tiebreaker. Unlike the shortcut creators this
    script intentionally keeps scoring simple -- LaunchBox will let you correct
    wrong entries after import.

    Blocked executables: setup, uninstall, redist, vcredist, directx, crash,
    config, update, install, service, handler.

.PARAMETER DryRun
    When $true (default), shows what would be written to the CSV without creating
    the file. Set to $false in the CONFIG block to write the CSV.

.EXAMPLE
    .\Export-LaunchBox-CSV.ps1

.NOTES
    After importing into LaunchBox:
      1. Review games with wrong executables using Edit Game
      2. Use "Audit" in LaunchBox to find missing media
      3. Run Create-PC-Shortcuts.ps1 to generate shortcuts from the same library

.VERSION
    2.0.0 - Config block, DryRun, proper CSV output with Export-Csv, MIT header.
            Removed setup.exe from priority list (was an installer, not a game).
            Removed hardcoded PC Games 2 path. Added blocked name list.
    1.0.0 - Initial release.

.LICENSE
    MIT License
    Copyright (c) Paul Mardis
#>

# ==============================================================================
# CONFIG -- Edit this block. Do not put paths anywhere else in this script.
# ==============================================================================
$RootGameFolder = "D:\Arcade\System roms\PC Games"
$OutputCsvFile  = Join-Path $env:USERPROFILE "Desktop\LaunchBoxGames.csv"
$Platform       = "Windows"       # LaunchBox platform name written to every row
$MaxDepth       = 3               # Depth to search for executables per game folder
$DryRun         = $true           # Set to $false to write the CSV file
$LogDir         = Join-Path $RootGameFolder "_Logs"
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

# ------------------------------------------------------------------------------
# Priority executable names (checked first -- must match game.exe not setup.exe)
# NOTE: "setup.exe" deliberately excluded -- it is an installer, not a launcher.
# ------------------------------------------------------------------------------
$PriorityNames = @(
    "game","start","run","main","play","launch","bin"
)

# ------------------------------------------------------------------------------
# Blocked executable names (never selected as main game exe)
# ------------------------------------------------------------------------------
$BlockedNames = @(
    "setup","uninstall","unins000","install","redist","vcredist","vc_redist",
    "directx","dxwebsetup","crashreport","crashreporter","crashpad",
    "config","configuration","settings","helper","service","server",
    "update","uploader","handler","webhelper"
)

$BlockedPatterns = @(
    "*uninstall*","*setup*","*install*","*config*","*crash*",
    "*update*","*redist*","*helper*","*service*","*server*"
)

# ------------------------------------------------------------------------------
# Select best exe from a list
# ------------------------------------------------------------------------------
function Select-MainExe {
    param([System.IO.FileInfo[]]$ExeFiles, [string]$GameName)

    $GameNameLower = $GameName.ToLower()

    # Filter out blocked files first
    $Candidates = $ExeFiles | Where-Object {
        $n = [System.IO.Path]::GetFileNameWithoutExtension($_.Name).ToLower()
        $blocked = $false
        if ($BlockedNames -contains $n) { $blocked = $true }
        foreach ($P in $BlockedPatterns) { if ($n -like $P) { $blocked = $true } }
        -not $blocked
    }

    if ($Candidates.Count -eq 0) { return $null }

    # Exact folder name match
    $ExactMatch = $Candidates | Where-Object {
        [System.IO.Path]::GetFileNameWithoutExtension($_.Name).ToLower() -eq $GameNameLower
    } | Select-Object -First 1
    if ($ExactMatch) { return $ExactMatch }

    # Priority list match
    foreach ($PName in $PriorityNames) {
        $PriorityMatch = $Candidates | Where-Object {
            [System.IO.Path]::GetFileNameWithoutExtension($_.Name).ToLower() -eq $PName
        } | Select-Object -First 1
        if ($PriorityMatch) { return $PriorityMatch }
    }

    # Partial folder name match
    $PartialMatch = $Candidates | Where-Object {
        [System.IO.Path]::GetFileNameWithoutExtension($_.Name).ToLower() -like "*$GameNameLower*"
    } | Select-Object -First 1
    if ($PartialMatch) { return $PartialMatch }

    # First remaining candidate
    return $Candidates | Select-Object -First 1
}

# ------------------------------------------------------------------------------
# Find exe files -- single-path descent (simple and fast for LaunchBox export)
# ------------------------------------------------------------------------------
function Find-GameExe {
    param([string]$FolderPath, [string]$GameName, [int]$MaxDepthParam)

    $ExeFiles = @()
    $CurrentFolder = $FolderPath
    $Depth = 0

    while ($Depth -le $MaxDepthParam) {
        $ExeFiles = @(Get-ChildItem -LiteralPath $CurrentFolder -Filter "*.exe" -File -ErrorAction SilentlyContinue)
        if ($ExeFiles.Count -gt 0 -and $Depth -gt 0) { break }

        $SubDirs = @(Get-ChildItem -LiteralPath $CurrentFolder -Directory -ErrorAction SilentlyContinue)

        # Stop descending if multiple subdirs (ambiguous) or none
        if ($SubDirs.Count -ne 1) { break }

        $CurrentFolder = $SubDirs[0].FullName
        $Depth++
    }

    return $ExeFiles
}

# ==============================================================================
# SETUP
# ==============================================================================
if (-not $DryRun -and -not (Test-Path -LiteralPath $LogDir)) {
    New-Item -ItemType Directory -Path $LogDir | Out-Null
}

$CsvRows     = [System.Collections.Generic.List[PSCustomObject]]::new()
$MissingExe  = [System.Collections.Generic.List[string]]::new()

# ==============================================================================
# BANNER
# ==============================================================================
Write-Host ""
Write-Host "Export-LaunchBox-CSV" -ForegroundColor Cyan
Write-Host "====================" -ForegroundColor Cyan
Write-Host "  Library  : $RootGameFolder"
Write-Host "  Output   : $OutputCsvFile"
Write-Host "  Platform : $Platform"
Write-Host "  DryRun   : $DryRun"
Write-Host ""
if ($DryRun) {
    Write-Host "  [DRY RUN] No CSV file will be written." -ForegroundColor Yellow
    Write-Host ""
}

# ==============================================================================
# MAIN LOOP
# ==============================================================================
$GameFolders = Get-ChildItem -LiteralPath $RootGameFolder -Directory | Sort-Object Name
$Total       = $GameFolders.Count
$Index       = 0
$CntFound    = 0
$CntMissing  = 0

Write-Host ("  Found {0} game folders." -f $Total) -ForegroundColor White
Write-Host ""

foreach ($GameDir in $GameFolders) {
    $Index++
    $GameName = $GameDir.Name

    Write-Progress -Activity "Export-LaunchBox-CSV" -Status "$GameName ($Index of $Total)" `
        -PercentComplete ([int](($Index / $Total) * 100))

    $ExeFiles = @(Find-GameExe -FolderPath $GameDir.FullName -GameName $GameName -MaxDepthParam $MaxDepth)

    if ($ExeFiles.Count -eq 0) {
        Write-Host ("  [NOT FOUND] {0}" -f $GameName) -ForegroundColor DarkGray
        $MissingExe.Add($GameName)
        $CntMissing++
        continue
    }

    $MainExe = Select-MainExe -ExeFiles $ExeFiles -GameName $GameName

    if ($null -eq $MainExe) {
        Write-Host ("  [BLOCKED]   {0}" -f $GameName) -ForegroundColor DarkGray
        $MissingExe.Add($GameName)
        $CntMissing++
        continue
    }

    Write-Host ("  [OK]        {0,-55} -> {1}" -f $GameName, $MainExe.Name) -ForegroundColor Green

    $CsvRows.Add([PSCustomObject]@{
        Title           = $GameName
        ApplicationPath = $MainExe.FullName
        Platform        = $Platform
    })
    $CntFound++
}

Write-Progress -Activity "Export-LaunchBox-CSV" -Completed

# ==============================================================================
# WRITE CSV
# ==============================================================================
if (-not $DryRun -and $CsvRows.Count -gt 0) {
    try {
        $CsvDir = Split-Path $OutputCsvFile -Parent
        if (-not (Test-Path -LiteralPath $CsvDir)) {
            New-Item -ItemType Directory -Path $CsvDir | Out-Null
        }
        $CsvRows | Export-Csv -LiteralPath $OutputCsvFile -NoTypeInformation -Encoding UTF8
        Write-Host ""
        Write-Host ("  CSV written: {0}" -f $OutputCsvFile) -ForegroundColor Green
        Write-Host "  Import into LaunchBox via:" -ForegroundColor Gray
        Write-Host "    Import -> ROM Files or Folders -> From CSV File" -ForegroundColor Gray
    }
    catch {
        Write-Host ("  ERROR writing CSV: {0}" -f $_.Exception.Message) -ForegroundColor Red
    }

    # Write missing list to log
    if ($MissingExe.Count -gt 0) {
        $MissingLog = Join-Path $LogDir ("Export-LaunchBox-CSV_Missing_{0}.txt" -f (Get-Date -Format "yyyyMMdd_HHmmss"))
        try {
            $MissingExe | Out-File -FilePath $MissingLog -Encoding UTF8
            Write-Host ("  Missing exe log: {0}" -f $MissingLog) -ForegroundColor Gray
        } catch {}
    }
}

# ==============================================================================
# SUMMARY
# ==============================================================================
Write-Host ""
Write-Host "Summary" -ForegroundColor White
Write-Host ("  {0,-14} {1}" -f "Games found :", $CntFound)   -ForegroundColor Green
Write-Host ("  {0,-14} {1}" -f "No exe found :", $CntMissing) -ForegroundColor $(if ($CntMissing -gt 0) { "Yellow" } else { "Gray" })
Write-Host ("  {0,-14} {1}" -f "Output :", $OutputCsvFile)

if ($DryRun) {
    Write-Host ""
    Write-Host "  DRY RUN complete. Set `$DryRun = `$false in the CONFIG block to write the CSV." -ForegroundColor Yellow
}

Write-Host ""
Set-ExitCode 0
