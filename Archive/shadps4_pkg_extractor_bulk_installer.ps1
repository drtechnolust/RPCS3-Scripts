#requires -version 5.1
<#
ShadPS4Plus / pkg_extractor.exe Bulk Installer
Option B - Folder-per-type mode

Usage:
  Run once per folder type (Game / Patch / DLC).
  You declare what type the PKGs are - no --check-type call.
  The extractor handles subfolder naming inside the destination.

Destinations:
  Game  -> GamesDir    (extractor creates GamesDir\CUSA#####)
  Patch -> GamesDir    (extractor creates GamesDir\CUSA#####-patch)
  DLC   -> DLCDir      (extractor creates DLCDir\CUSA#####\ContentID)

Notes:
  - Designed for PowerShell ISE (no Windows Forms dialogs)
  - All paths typed/pasted at console prompts
  - Strips surrounding quotes from pasted paths
  - Incremental CSV log written per-file (survives Ctrl+C)
  - Successfully extracted PKGs moved to _Completed subfolder
#>

# ---------------------------------------------------------------------------
# Console input helpers
# ---------------------------------------------------------------------------

function Read-FolderPath {
    param([string]$Prompt)
    while ($true) {
        Write-Host "  $Prompt" -ForegroundColor DarkCyan
        $raw = Read-Host "  Folder path"
        $p   = $raw.Trim().Trim('"').Trim("'")
        if ([string]::IsNullOrWhiteSpace($p)) {
            Write-Host "  Path cannot be empty. Please try again." -ForegroundColor Yellow
            continue
        }
        if (-not (Test-Path -LiteralPath $p -PathType Container)) {
            Write-Host "  Folder not found: $p" -ForegroundColor Yellow
            Write-Host "  Please check the path and try again." -ForegroundColor Yellow
            continue
        }
        return $p
    }
}

function Read-FilePath {
    param([string]$Prompt)
    while ($true) {
        Write-Host "  $Prompt" -ForegroundColor DarkCyan
        $raw = Read-Host "  File path"
        $p   = $raw.Trim().Trim('"').Trim("'")
        if ([string]::IsNullOrWhiteSpace($p)) {
            Write-Host "  Path cannot be empty. Please try again." -ForegroundColor Yellow
            continue
        }
        if (-not (Test-Path -LiteralPath $p -PathType Leaf)) {
            Write-Host "  File not found: $p" -ForegroundColor Yellow
            Write-Host "  Please check the path and try again." -ForegroundColor Yellow
            continue
        }
        return $p
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

function Ask-YesNo {
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

# ---------------------------------------------------------------------------
# Filesystem helpers
# ---------------------------------------------------------------------------

function Ensure-Folder {
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
    Ensure-Folder -Path $CompletedFolder
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

# ---------------------------------------------------------------------------
# External process runner
# ---------------------------------------------------------------------------

function Invoke-External {
    param(
        [Parameter(Mandatory)][string]$ExePath,
        [Parameter(Mandatory)][string[]]$Arguments,
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

    return [pscustomobject]@{
        ExitCode = $proc.ExitCode
        StdOut   = $stdout
        StdErr   = $stderr
    }
}

# ---------------------------------------------------------------------------
# Progress bar
# ---------------------------------------------------------------------------

function Write-ProgressBar {
    param([int]$Current, [int]$Total, [int]$Success, [int]$Errors, [int]$Skipped, [int]$BarWidth = 40)
    if ($Total -eq 0) { return }
    $pct    = [math]::Round(($Current / $Total) * 100, 0)
    $filled = [int][math]::Round(($Current / $Total) * $BarWidth, 0)
    $empty  = $BarWidth - $filled
    $bar    = ('#' * $filled) + ('-' * $empty)
    Write-Host "  [$bar] $pct% ($Current/$Total)  OK:$Success  ERR:$Errors  SKIP:$Skipped" -ForegroundColor Cyan
}

# ---------------------------------------------------------------------------
# Banner
# ---------------------------------------------------------------------------

Write-Host ""
Write-Host "================================================" -ForegroundColor Cyan
Write-Host "  ShadPS4 Bulk PKG Installer  (Option B)       " -ForegroundColor Cyan
Write-Host "================================================" -ForegroundColor Cyan
Write-Host ""

# ---------------------------------------------------------------------------
# Inputs
# ---------------------------------------------------------------------------

Write-Host "--- PKG TYPE ---" -ForegroundColor White
$pkgType = Read-PkgType
Write-Host "  Type : $pkgType" -ForegroundColor DarkGreen

Write-Host ""
Write-Host "--- SOURCE FOLDER ---" -ForegroundColor White
$sourceFolder = Read-FolderPath -Prompt "Enter the folder path containing your $pkgType PKG files"
Write-Host "  Source : $sourceFolder" -ForegroundColor DarkGreen

Write-Host ""
Write-Host "--- PKG_EXTRACTOR.EXE ---" -ForegroundColor White
$extractorPath = Read-FilePath -Prompt "Enter the full path to pkg_extractor.exe"
Write-Host "  Extractor : $extractorPath" -ForegroundColor DarkGreen

Write-Host ""
Write-Host "--- DESTINATION FOLDER ---" -ForegroundColor White
$destPrompt = switch ($pkgType) {
    'Game'  { "Enter your Official PS4 Games folder (extractor creates CUSA##### subfolders here)" }
    'Patch' { "Enter your Official PS4 Games folder (extractor creates CUSA#####-patch subfolders here)" }
    'DLC'   { "Enter your DLC folder (extractor creates CUSA#####\ContentID subfolders here)" }
}
$destFolder = Read-FolderPath -Prompt $destPrompt
Write-Host "  Destination : $destFolder" -ForegroundColor DarkGreen

# ---------------------------------------------------------------------------
# Options
# ---------------------------------------------------------------------------

Write-Host ""
Write-Host "--- OPTIONS ---" -ForegroundColor White
$recurse = Ask-YesNo "  Search subfolders recursively?" $false
$dryRun  = Ask-YesNo "  Dry run only (no actual extraction)?" $false

# ---------------------------------------------------------------------------
# Folder setup
# ---------------------------------------------------------------------------

$logFolder       = Join-Path $sourceFolder "_Logs"
$completedFolder = Join-Path $sourceFolder "_Completed"

Ensure-Folder -Path $logFolder
Ensure-Folder -Path $completedFolder

$stamp  = Get-Date -Format "yyyyMMdd_HHmmss"
$csvLog = Join-Path $logFolder "ShadPS4_${pkgType}_$stamp.csv"
$txtLog = Join-Path $logFolder "ShadPS4_${pkgType}_$stamp.txt"

Write-LogLine -Path $txtLog -Text "Run started"
Write-LogLine -Path $txtLog -Text "PkgType   : $pkgType"
Write-LogLine -Path $txtLog -Text "Source    : $sourceFolder"
Write-LogLine -Path $txtLog -Text "Extractor : $extractorPath"
Write-LogLine -Path $txtLog -Text "Dest      : $destFolder"
Write-LogLine -Path $txtLog -Text "Recurse   : $recurse"
Write-LogLine -Path $txtLog -Text "DryRun    : $dryRun"

# ---------------------------------------------------------------------------
# Collect PKG files
# ---------------------------------------------------------------------------

$params = @{
    LiteralPath = $sourceFolder
    Filter      = "*.pkg"
    File        = $true
}
if ($recurse) { $params.Recurse = $true }

$pkgFiles = Get-ChildItem @params |
    Where-Object {
        $_.FullName -notlike (Join-Path $logFolder       '*') -and
        $_.FullName -notlike (Join-Path $completedFolder '*')
    } |
    Sort-Object FullName

if (-not $pkgFiles -or $pkgFiles.Count -eq 0) {
    Write-Host ""
    Write-Host "  No PKG files found in: $sourceFolder" -ForegroundColor Yellow
    Write-LogLine -Path $txtLog -Text "No PKG files found."
    return
}

$total = $pkgFiles.Count
Write-Host ""
Write-Host "  Found $total PKG file(s) to process." -ForegroundColor Green
Write-Host ""

# ---------------------------------------------------------------------------
# Counters
# ---------------------------------------------------------------------------

$cSuccess = 0
$cError   = 0
$cDryRun  = 0
$firstRow = $true
$index    = 0

# ---------------------------------------------------------------------------
# Main loop
# ---------------------------------------------------------------------------

foreach ($pkg in $pkgFiles) {
    $index++

    $status       = "Pending"
    $notes        = ""
    $installExit  = $null
    $finalPkgPath = $pkg.FullName

    Write-Host ""
    Write-Host "  [$index/$total] $($pkg.Name)" -ForegroundColor Cyan
    Write-LogLine -Path $txtLog -Text "--- [$index/$total] $($pkg.FullName)"

    try {
        if ($dryRun) {
            $status  = "DryRun"
            $notes   = "Would extract to: $destFolder"
            $cDryRun++
            Write-Host "    DRY RUN - $notes" -ForegroundColor Yellow
            Write-LogLine -Path $txtLog -Text "DRY RUN: $notes"
        }
        else {
            $install     = Invoke-External -ExePath $extractorPath `
                                           -Arguments @('"' + $pkg.FullName + '"', '"' + $destFolder + '"')
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
                $notes  = "Extracted successfully to $destFolder"
                $cSuccess++
                Write-Host "    SUCCESS" -ForegroundColor Green
                Write-LogLine -Path $txtLog -Text "SUCCESS: $notes"

                $newPath      = Move-ToCompleted -FilePath $pkg.FullName -CompletedFolder $completedFolder
                $finalPkgPath = $newPath
                Write-Host "    Moved to : $newPath" -ForegroundColor DarkGreen
                Write-LogLine -Path $txtLog -Text "Moved to: $newPath"
            }
            else {
                $status = "Failed"
                $notes  = "Extractor returned exit code $installExit"
                $cError++
                Write-Host "    FAILED - exit code $installExit" -ForegroundColor Red
                Write-LogLine -Path $txtLog -Text "FAILED: $notes"
            }
        }
    }
    catch {
        $status = "Error"
        $notes  = $_.Exception.Message
        $cError++
        Write-Host "    ERROR: $notes" -ForegroundColor Red
        Write-LogLine -Path $txtLog -Text "ERROR: $notes"
    }

    # Write CSV row immediately (survives Ctrl+C)
    $row = [pscustomobject]@{
        TimeStamp        = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        PkgType          = $pkgType
        FileName         = $pkg.Name
        OriginalFullPath = $pkg.FullName
        FinalPkgPath     = $finalPkgPath
        SizeMB           = [math]::Round(($pkg.Length / 1MB), 2)
        Destination      = $destFolder
        DryRun           = $dryRun
        Status           = $status
        InstallExitCode  = $installExit
        Notes            = $notes
    }
    Write-CsvRow -CsvPath $csvLog -WriteHeader $firstRow -Row $row
    $firstRow = $false

    Write-ProgressBar -Current $index -Total $total -Success $cSuccess -Errors $cError -Skipped $cDryRun
}

# ---------------------------------------------------------------------------
# Summary
# ---------------------------------------------------------------------------

Write-Host ""
Write-Host ""
Write-Host "================================================" -ForegroundColor Cyan
Write-Host "  DONE" -ForegroundColor Cyan
Write-Host "================================================" -ForegroundColor Cyan
Write-Host "  Type             : $pkgType"
Write-Host "  Total            : $total"
Write-Host "  Success          : $cSuccess"  -ForegroundColor Green
Write-Host "  Failed / Errors  : $cError"    -ForegroundColor Red
Write-Host "  Dry Run          : $cDryRun"   -ForegroundColor Yellow
Write-Host ""
Write-Host "  Completed folder : $completedFolder"
Write-Host "  CSV Log          : $csvLog"
Write-Host "  TXT Log          : $txtLog"
Write-Host ""

Write-LogLine -Path $txtLog -Text "Run finished. Type=$pkgType, Total=$total, Success=$cSuccess, Failed=$cError, DryRun=$cDryRun"