#Requires -Version 5.1
<#
.SYNOPSIS
    Bulk installs PS4 PKG files using pkg_extractor.exe with single-type or mixed-mode routing.

.DESCRIPTION
    Processes a folder of PS4 PKG files and extracts them to the correct destination
    using pkg_extractor.exe. Supports two modes:

      Single-type  -- You declare the type (Game / Patch / DLC). No --check-type call.
                      Fastest. Use when your source folder contains one type only.

      Mixed folder -- Runs --check-type on every PKG to detect type automatically.
                      Routes Games and Patches to GamesDir, DLC to DLCDir.
                      Use when your source folder contains mixed types.

    In both modes the extractor handles subfolder naming:
      Game  -> GamesDir\CUSA#####
      Patch -> GamesDir\CUSA#####-patch
      DLC   -> DLCDir\CUSA#####\ContentID

    Features:
      - Interactive prompts for all paths (ISE and terminal compatible)
      - Dry run option previews extraction without writing anything
      - Incremental CSV log written after each file (survives Ctrl+C)
      - Text log with full stdout/stderr from extractor
      - Successfully extracted PKGs moved to _Completed subfolder
      - Progress bar shown after each file

.NOTES
    Requires: pkg_extractor.exe (ShadPS4Plus tool)
    Pre-fill $DefaultExtractorPath in the CONFIG block to skip the extractor prompt.

.EXAMPLE
    .\ShadPS4-PKG-Installer.ps1

.VERSION
    2.0.0 - MIT header, config block for common defaults, removed box-drawing
            characters from banner, ISE-compatible exit pattern. All logic preserved.
    1.1.0 - Added mixed-mode type detection via --check-type.
    1.0.0 - Initial release (single-type mode only).

.LICENSE
    MIT License
    Copyright (c) Paul Mardis
#>

# ==============================================================================
# CONFIG -- Pre-fill these to skip repeated prompts. Leave blank to be prompted.
# ==============================================================================
$DefaultSourceFolder   = ""    # e.g. "D:\Arcade\System roms\Sony Playstation 4\Downloading PS4\Games"
$DefaultGamesDir       = "D:\Arcade\System roms\Sony Playstation 4\Official PS4 Games"
$DefaultDLCDir         = ""    # e.g. "D:\Arcade\System roms\Sony Playstation 4\DLC"
$DefaultExtractorPath  = ""    # e.g. "C:\Tools\pkg_extractor.exe"
# ==============================================================================

function Set-ExitCode {
    param([int]$Code)
    $global:LASTEXITCODE = $Code
}

# ------------------------------------------------------------------------------
# Console input helpers
# ------------------------------------------------------------------------------
function Read-FolderPath {
    param([string]$Prompt, [string]$Default = "")

    if (-not [string]::IsNullOrWhiteSpace($Default)) {
        if (Test-Path -LiteralPath $Default -PathType Container) {
            Write-Host "  Using default: $Default" -ForegroundColor DarkGreen
            return $Default
        }
        Write-Host "  Default path not found ($Default) -- please enter manually." -ForegroundColor Yellow
    }

    while ($true) {
        Write-Host "  $Prompt" -ForegroundColor DarkCyan
        $raw = Read-Host "  Folder path"
        $p   = $raw.Trim().Trim('"').Trim("'")
        if ([string]::IsNullOrWhiteSpace($p)) {
            Write-Host "  Path cannot be empty." -ForegroundColor Yellow
            continue
        }
        if (-not (Test-Path -LiteralPath $p -PathType Container)) {
            Write-Host "  Folder not found: $p" -ForegroundColor Yellow
            continue
        }
        return $p
    }
}

function Read-FilePath {
    param([string]$Prompt, [string]$Default = "")

    if (-not [string]::IsNullOrWhiteSpace($Default)) {
        if (Test-Path -LiteralPath $Default -PathType Leaf) {
            Write-Host "  Using default: $Default" -ForegroundColor DarkGreen
            return $Default
        }
        Write-Host "  Default path not found ($Default) -- please enter manually." -ForegroundColor Yellow
    }

    while ($true) {
        Write-Host "  $Prompt" -ForegroundColor DarkCyan
        $raw = Read-Host "  File path"
        $p   = $raw.Trim().Trim('"').Trim("'")
        if ([string]::IsNullOrWhiteSpace($p)) {
            Write-Host "  Path cannot be empty." -ForegroundColor Yellow
            continue
        }
        if (-not (Test-Path -LiteralPath $p -PathType Leaf)) {
            Write-Host "  File not found: $p" -ForegroundColor Yellow
            continue
        }
        return $p
    }
}

function Read-Mode {
    while ($true) {
        Write-Host "  Select mode:" -ForegroundColor DarkCyan
        Write-Host "    1 = Single-type  (all PKGs are the same type, no auto-detect)" -ForegroundColor White
        Write-Host "    2 = Mixed folder (mixed types, uses --check-type to route each PKG)" -ForegroundColor White
        $answer = Read-Host "  Enter 1 or 2"
        switch ($answer.Trim()) {
            '1' { return 'Single' }
            '2' { return 'Mixed'  }
            default { Write-Host "  Please enter 1 or 2." -ForegroundColor Yellow }
        }
    }
}

function Read-PkgType {
    while ($true) {
        Write-Host "  What type of PKGs are in this folder?" -ForegroundColor DarkCyan
        Write-Host "    1 = Game" -ForegroundColor White
        Write-Host "    2 = Patch" -ForegroundColor White
        Write-Host "    3 = DLC" -ForegroundColor White
        $answer = Read-Host "  Enter 1, 2 or 3"
        switch ($answer.Trim()) {
            '1' { return 'Game'  }
            '2' { return 'Patch' }
            '3' { return 'DLC'   }
            default { Write-Host "  Please enter 1, 2 or 3." -ForegroundColor Yellow }
        }
    }
}

function Invoke-YesNo {
    param([string]$Prompt, [bool]$Default = $true)
    $defaultText = if ($Default) { "Y/n" } else { "y/N" }
    while ($true) {
        $answer = Read-Host "$Prompt [$defaultText]"
        if ([string]::IsNullOrWhiteSpace($answer)) { return $Default }
        switch -Regex ($answer.Trim()) {
            '^(y|yes)$' { return $true  }
            '^(n|no)$'  { return $false }
            default     { Write-Host "  Please enter Y or N." -ForegroundColor Yellow }
        }
    }
}

# ------------------------------------------------------------------------------
# Filesystem helpers
# ------------------------------------------------------------------------------
function Confirm-Folder {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) {
        New-Item -ItemType Directory -Path $Path -Force | Out-Null
    }
}

function Write-LogLine {
    param([string]$Path, [string]$Text)
    $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Add-Content -LiteralPath $Path -Value "[$ts] $Text"
}

function Write-CsvRow {
    param([string]$CsvPath, [bool]$WriteHeader, [pscustomobject]$Row)
    if ($WriteHeader) {
        $Row | Export-Csv -LiteralPath $CsvPath -NoTypeInformation -Encoding UTF8
    } else {
        $Row | Export-Csv -LiteralPath $CsvPath -NoTypeInformation -Encoding UTF8 -Append
    }
}

function Move-ToCompleted {
    param([string]$FilePath, [string]$CompletedFolder)
    Confirm-Folder -Path $CompletedFolder
    $name = [IO.Path]::GetFileName($FilePath)
    $dest = Join-Path $CompletedFolder $name
    if (Test-Path -LiteralPath $dest) {
        $base = [IO.Path]::GetFileNameWithoutExtension($name)
        $ext  = [IO.Path]::GetExtension($name)
        $dest = Join-Path $CompletedFolder ("{0}_{1}{2}" -f $base, (Get-Date -Format "yyyyMMdd_HHmmss"), $ext)
    }
    Move-Item -LiteralPath $FilePath -Destination $dest -Force
    return $dest
}

# ------------------------------------------------------------------------------
# External process runner
# ------------------------------------------------------------------------------
function Invoke-External {
    param(
        [string]$ExePath,
        [string[]]$Arguments,
        [string]$WorkingDirectory = (Split-Path -Path $ExePath -Parent)
    )
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName               = $ExePath
    $psi.Arguments              = ($Arguments -join ' ')
    $psi.WorkingDirectory       = $WorkingDirectory
    $psi.UseShellExecute        = $false
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError  = $true
    $psi.CreateNoWindow         = $true

    $proc = New-Object System.Diagnostics.Process
    $proc.StartInfo = $psi
    [void]$proc.Start()
    $stdout = $proc.StandardOutput.ReadToEnd()
    $stderr = $proc.StandardError.ReadToEnd()
    $proc.WaitForExit()

    return [PSCustomObject]@{
        ExitCode = $proc.ExitCode
        StdOut   = $stdout
        StdErr   = $stderr
    }
}

function Get-PkgTypeName {
    param([int]$ExitCode)
    switch ($ExitCode) {
        101 { return 'Game'  }
        102 { return 'Patch' }
        103 { return 'DLC'   }
        default { return 'Unknown' }
    }
}

# ------------------------------------------------------------------------------
# Progress bar (ASCII only)
# ------------------------------------------------------------------------------
function Write-ProgressBar {
    param([int]$Current, [int]$Total, [int]$Success, [int]$Errors, [int]$Skipped, [int]$BarWidth = 40)
    if ($Total -eq 0) { return }
    $pct    = [math]::Round(($Current / $Total) * 100, 0)
    $filled = [int][math]::Round(($Current / $Total) * $BarWidth, 0)
    $empty  = $BarWidth - $filled
    $bar    = ('#' * $filled) + ('-' * $empty)
    Write-Host ("  [{0}] {1}% ({2}/{3})  OK:{4}  ERR:{5}  SKIP:{6}" -f $bar, $pct, $Current, $Total, $Success, $Errors, $Skipped) -ForegroundColor Cyan
}

# ==============================================================================
# BANNER
# ==============================================================================
Write-Host ""
Write-Host "================================================" -ForegroundColor Cyan
Write-Host "  ShadPS4 Bulk PKG Installer" -ForegroundColor Cyan
Write-Host "================================================" -ForegroundColor Cyan
Write-Host ""

# ==============================================================================
# INTERACTIVE SETUP
# ==============================================================================
Write-Host "--- MODE ---" -ForegroundColor White
$mode = Read-Mode
Write-Host "  Mode : $mode" -ForegroundColor DarkGreen

Write-Host ""
Write-Host "--- SOURCE FOLDER ---" -ForegroundColor White
$sourceFolder = Read-FolderPath -Prompt "Enter the folder path containing your PKG files" `
                                -Default $DefaultSourceFolder
Write-Host "  Source : $sourceFolder" -ForegroundColor DarkGreen

Write-Host ""
Write-Host "--- PKG_EXTRACTOR.EXE ---" -ForegroundColor White
$extractorPath = Read-FilePath -Prompt "Enter the full path to pkg_extractor.exe" `
                               -Default $DefaultExtractorPath
Write-Host "  Extractor : $extractorPath" -ForegroundColor DarkGreen

# Destination(s)
$pkgType    = $null
$destFolder = $null
$gamesDir   = $null
$dlcDir     = $null

if ($mode -eq 'Single') {
    Write-Host ""
    Write-Host "--- PKG TYPE ---" -ForegroundColor White
    $pkgType = Read-PkgType
    Write-Host "  Type : $pkgType" -ForegroundColor DarkGreen

    Write-Host ""
    Write-Host "--- DESTINATION FOLDER ---" -ForegroundColor White
    $destPrompt = switch ($pkgType) {
        'Game'  { "Enter your Official PS4 Games folder (extractor creates CUSA##### subfolders here)" }
        'Patch' { "Enter your Official PS4 Games folder (extractor creates CUSA#####-patch subfolders here)" }
        'DLC'   { "Enter your DLC folder (extractor creates CUSA#####\ContentID subfolders here)" }
    }
    $destFolder = Read-FolderPath -Prompt $destPrompt -Default $DefaultGamesDir
    Write-Host "  Destination : $destFolder" -ForegroundColor DarkGreen
}
else {
    Write-Host ""
    Write-Host "--- GAMES DESTINATION ---" -ForegroundColor White
    $gamesDir = Read-FolderPath -Prompt "Enter your Official PS4 Games folder (for Games and Patches)" `
                                -Default $DefaultGamesDir
    Write-Host "  Games : $gamesDir" -ForegroundColor DarkGreen

    Write-Host ""
    Write-Host "--- DLC DESTINATION ---" -ForegroundColor White
    $dlcDir = Read-FolderPath -Prompt "Enter your DLC folder" -Default $DefaultDLCDir
    Write-Host "  DLC   : $dlcDir" -ForegroundColor DarkGreen
}

Write-Host ""
Write-Host "--- OPTIONS ---" -ForegroundColor White
$recurse = Invoke-YesNo "  Search subfolders recursively?" $false
$dryRun  = Invoke-YesNo "  Dry run only (no actual extraction)?" $false

# ==============================================================================
# FOLDER SETUP
# ==============================================================================
$logFolder       = Join-Path $sourceFolder "_Logs"
$completedFolder = Join-Path $sourceFolder "_Completed"

Confirm-Folder -Path $logFolder
Confirm-Folder -Path $completedFolder

$modeLabel = if ($mode -eq 'Single') { "Single_$pkgType" } else { "Mixed" }
$stamp     = Get-Date -Format "yyyyMMdd_HHmmss"
$csvLog    = Join-Path $logFolder ("ShadPS4_${modeLabel}_${stamp}.csv")
$txtLog    = Join-Path $logFolder ("ShadPS4_${modeLabel}_${stamp}.txt")

Write-LogLine -Path $txtLog -Text "Run started"
Write-LogLine -Path $txtLog -Text "Mode      : $mode"
if ($mode -eq 'Single') {
    Write-LogLine -Path $txtLog -Text "PkgType   : $pkgType"
    Write-LogLine -Path $txtLog -Text "Dest      : $destFolder"
} else {
    Write-LogLine -Path $txtLog -Text "GamesDir  : $gamesDir"
    Write-LogLine -Path $txtLog -Text "DLCDir    : $dlcDir"
}
Write-LogLine -Path $txtLog -Text "Source    : $sourceFolder"
Write-LogLine -Path $txtLog -Text "Extractor : $extractorPath"
Write-LogLine -Path $txtLog -Text "Recurse   : $recurse"
Write-LogLine -Path $txtLog -Text "DryRun    : $dryRun"

# ==============================================================================
# COLLECT PKG FILES
# ==============================================================================
$getParams = @{
    LiteralPath = $sourceFolder
    Filter      = "*.pkg"
    File        = $true
}
if ($recurse) { $getParams.Recurse = $true }

$pkgFiles = Get-ChildItem @getParams |
    Where-Object {
        $_.FullName -notlike (Join-Path $logFolder       '*') -and
        $_.FullName -notlike (Join-Path $completedFolder '*')
    } |
    Sort-Object FullName

if (-not $pkgFiles -or $pkgFiles.Count -eq 0) {
    Write-Host ""
    Write-Host "  No PKG files found in: $sourceFolder" -ForegroundColor Yellow
    Write-LogLine -Path $txtLog -Text "No PKG files found."
    Set-ExitCode 0
    return
}

$total = $pkgFiles.Count
Write-Host ""
Write-Host ("  Found {0} PKG file(s) to process." -f $total) -ForegroundColor Green
Write-Host ""

# ==============================================================================
# COUNTERS
# ==============================================================================
$cSuccess = 0
$cError   = 0
$cSkipped = 0
$cDryRun  = 0
$firstRow = $true
$index    = 0

# ==============================================================================
# MAIN LOOP
# ==============================================================================
foreach ($pkg in $pkgFiles) {
    $index++

    $status       = "Pending"
    $notes        = ""
    $installExit  = $null
    $finalPkgPath = $pkg.FullName
    $detectedType = $pkgType
    $resolvedDest = $destFolder

    Write-Host ""
    Write-Host ("  [{0}/{1}] {2}" -f $index, $total, $pkg.Name) -ForegroundColor Cyan
    Write-LogLine -Path $txtLog -Text ("--- [{0}/{1}] {2}" -f $index, $total, $pkg.FullName)

    try {

        # --- Mixed mode: detect type ---
        if ($mode -eq 'Mixed') {
            $check        = Invoke-External -ExePath $extractorPath `
                                            -Arguments @('"' + $pkg.FullName + '"', '--check-type')
            $detectedType = Get-PkgTypeName -ExitCode $check.ExitCode

            if (-not [string]::IsNullOrWhiteSpace($check.StdOut)) {
                Write-LogLine -Path $txtLog -Text "CHECK STDOUT:"
                Add-Content -LiteralPath $txtLog -Value $check.StdOut.Trim()
            }
            if (-not [string]::IsNullOrWhiteSpace($check.StdErr)) {
                Write-LogLine -Path $txtLog -Text "CHECK STDERR:"
                Add-Content -LiteralPath $txtLog -Value $check.StdErr.Trim()
            }

            if ($detectedType -eq 'Unknown') {
                $status = "SkippedUnknown"
                $notes  = "Extractor could not identify PKG (exit code $($check.ExitCode))"
                $cSkipped++
                Write-Host ("    SKIPPED - {0}" -f $notes) -ForegroundColor Yellow
                Write-LogLine -Path $txtLog -Text "SKIPPED: $notes"

                $row = [PSCustomObject]@{
                    TimeStamp        = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
                    Mode             = $mode
                    DetectedType     = $detectedType
                    FileName         = $pkg.Name
                    OriginalFullPath = $pkg.FullName
                    FinalPkgPath     = $finalPkgPath
                    SizeMB           = [math]::Round(($pkg.Length / 1MB), 2)
                    Destination      = ''
                    DryRun           = $dryRun
                    Status           = $status
                    InstallExitCode  = $installExit
                    Notes            = $notes
                }
                Write-CsvRow -CsvPath $csvLog -WriteHeader $firstRow -Row $row
                $firstRow = $false
                Write-ProgressBar -Current $index -Total $total -Success $cSuccess -Errors $cError -Skipped $cSkipped
                continue
            }

            $resolvedDest = if ($detectedType -eq 'DLC') { $dlcDir } else { $gamesDir }
            Write-Host ("    Detected : {0}  ->  {1}" -f $detectedType, $resolvedDest) -ForegroundColor White
            Write-LogLine -Path $txtLog -Text "Detected=$detectedType | Dest=$resolvedDest"
        }

        # --- Extract or dry run ---
        if ($dryRun) {
            $status = "DryRun"
            $notes  = "Would extract to: $resolvedDest"
            $cDryRun++
            Write-Host ("    DRY RUN - {0}" -f $notes) -ForegroundColor Yellow
            Write-LogLine -Path $txtLog -Text "DRY RUN: $notes"
        }
        else {
            $install     = Invoke-External -ExePath $extractorPath `
                                           -Arguments @('"' + $pkg.FullName + '"', '"' + $resolvedDest + '"')
            $installExit = $install.ExitCode

            if (-not [string]::IsNullOrWhiteSpace($install.StdOut)) {
                Write-LogLine -Path $txtLog -Text "STDOUT:"
                Add-Content -LiteralPath $txtLog -Value $install.StdOut.Trim()
            }
            if (-not [string]::IsNullOrWhiteSpace($install.StdErr)) {
                Write-LogLine -Path $txtLog -Text "STDERR:"
                Add-Content -LiteralPath $txtLog -Value $install.StdErr.Trim()
            }

            if ($installExit -eq 0) {
                $status = "Success"
                $notes  = "Extracted to $resolvedDest"
                $cSuccess++
                Write-Host "    SUCCESS" -ForegroundColor Green
                Write-LogLine -Path $txtLog -Text "SUCCESS: $notes"

                $newPath      = Move-ToCompleted -FilePath $pkg.FullName -CompletedFolder $completedFolder
                $finalPkgPath = $newPath
                Write-Host ("    Moved to : {0}" -f $newPath) -ForegroundColor DarkGreen
                Write-LogLine -Path $txtLog -Text "Moved to: $newPath"
            }
            else {
                $status = "Failed"
                $notes  = "Extractor returned exit code $installExit"
                $cError++
                Write-Host ("    FAILED - exit code {0}" -f $installExit) -ForegroundColor Red
                Write-LogLine -Path $txtLog -Text "FAILED: $notes"
            }
        }
    }
    catch {
        $status = "Error"
        $notes  = $_.Exception.Message
        $cError++
        Write-Host ("    ERROR: {0}" -f $notes) -ForegroundColor Red
        Write-LogLine -Path $txtLog -Text "ERROR: $notes"
    }

    # Write CSV row immediately (survives Ctrl+C)
    $row = [PSCustomObject]@{
        TimeStamp        = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
        Mode             = $mode
        DetectedType     = $detectedType
        FileName         = $pkg.Name
        OriginalFullPath = $pkg.FullName
        FinalPkgPath     = $finalPkgPath
        SizeMB           = [math]::Round(($pkg.Length / 1MB), 2)
        Destination      = $resolvedDest
        DryRun           = $dryRun
        Status           = $status
        InstallExitCode  = $installExit
        Notes            = $notes
    }
    Write-CsvRow -CsvPath $csvLog -WriteHeader $firstRow -Row $row
    $firstRow = $false

    Write-ProgressBar -Current $index -Total $total -Success $cSuccess -Errors $cError -Skipped ($cSkipped + $cDryRun)
}

# ==============================================================================
# SUMMARY
# ==============================================================================
Write-Host ""
Write-Host ""
Write-Host "================================================" -ForegroundColor Cyan
Write-Host "  DONE" -ForegroundColor Cyan
Write-Host "================================================" -ForegroundColor Cyan
Write-Host ("  Mode            : {0}" -f $mode)
if ($mode -eq 'Single') {
    Write-Host ("  Type            : {0}" -f $pkgType)
}
Write-Host ("  Total           : {0}" -f $total)
Write-Host ("  Success         : {0}" -f $cSuccess)  -ForegroundColor Green
Write-Host ("  Failed / Errors : {0}" -f $cError)    -ForegroundColor $(if ($cError -gt 0) { "Red" } else { "Gray" })
Write-Host ("  Dry Run         : {0}" -f $cDryRun)   -ForegroundColor Yellow
if ($mode -eq 'Mixed') {
    Write-Host ("  Skipped Unknown : {0}" -f $cSkipped) -ForegroundColor Yellow
}
Write-Host ""
Write-Host ("  Completed folder : {0}" -f $completedFolder)
Write-Host ("  CSV Log          : {0}" -f $csvLog)
Write-Host ("  TXT Log          : {0}" -f $txtLog)
Write-Host ""

Write-LogLine -Path $txtLog -Text ("Run finished. Mode={0}, Total={1}, Success={2}, Failed={3}, DryRun={4}, Skipped={5}" -f $mode, $total, $cSuccess, $cError, $cDryRun, $cSkipped)
Set-ExitCode 0
