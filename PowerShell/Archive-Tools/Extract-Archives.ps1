# ============================================================
# Extract-Archives.ps1
# Extracts .zip and .rar files to individual folders using 7-Zip,
# shows progress, and moves completed archives to a 'Complete' folder.
# ============================================================

# ---- CONFIGURATION -----------------------------------------
$DryRunLimit = 10   # Number of files to process in dry run
# ------------------------------------------------------------

# Prompt for source folder
Write-Host ""
Write-Host "=== Archive Extractor (7-Zip) ===" -ForegroundColor Cyan
Write-Host ""
$SourcePath = Read-Host "Enter SOURCE path (folder containing your ZIP/RAR files)"
$SourcePath = $SourcePath.Trim('"').Trim("'")
if (-not (Test-Path $SourcePath)) {
    Write-Host "Source path not found: $SourcePath" -ForegroundColor Red; return
}

# Prompt for destination folder
$DestPath = Read-Host "Enter DESTINATION path (where extracted folders will be created)"
$DestPath = $DestPath.Trim('"').Trim("'")

Write-Host ""
Write-Host "Source      : $SourcePath" -ForegroundColor Gray
Write-Host "Destination : $DestPath" -ForegroundColor Gray

# ---- LOCATE 7-ZIP ------------------------------------------
$7zipPath = $null
$7zipCandidates = @(
    "C:\Program Files\7-Zip\7z.exe",
    "C:\Program Files (x86)\7-Zip\7z.exe"
)
foreach ($candidate in $7zipCandidates) {
    if (Test-Path $candidate) { $7zipPath = $candidate; break }
}

if (-not $7zipPath) {
    Write-Host ""
    Write-Host "ERROR: 7-Zip not found. Please install it from https://www.7-zip.org/" -ForegroundColor Red
    Write-Host "       Checked:" -ForegroundColor Red
    foreach ($c in $7zipCandidates) { Write-Host "         $c" -ForegroundColor Red }
    return
}

Write-Host "7-Zip           : $7zipPath" -ForegroundColor Gray

# Build Complete folder path
$CompletePath = Join-Path $DestPath "Complete"

# Gather all ZIP and RAR files
$Archives = Get-ChildItem -Path $SourcePath -File | Where-Object { $_.Extension -match '^\.(zip|rar)$' } | Sort-Object Name

if ($Archives.Count -eq 0) {
    Write-Host ""
    Write-Host "No .zip or .rar files found in: $SourcePath" -ForegroundColor Red
    return
}

Write-Host "Archives found  : $($Archives.Count)" -ForegroundColor Cyan

# ---- DRY RUN PROMPT ----------------------------------------
Write-Host ""
$RunMode = Read-Host "Run mode? Type 'dry' for first $DryRunLimit files, or 'full' for all"

if ($RunMode -eq 'dry') {
    $Archives = $Archives | Select-Object -First $DryRunLimit
    Write-Host ""
    Write-Host "--- DRY RUN: Processing first $($Archives.Count) archive(s) ---" -ForegroundColor Magenta
} elseif ($RunMode -eq 'full') {
    Write-Host ""
    Write-Host "--- FULL RUN: Processing all $($Archives.Count) archive(s) ---" -ForegroundColor Green
} else {
    Write-Host "Invalid input. Please type 'dry' or 'full'. Exiting." -ForegroundColor Red
    return
}

# Create destination and Complete folders if needed
if (-not (Test-Path $DestPath))     { New-Item -ItemType Directory -Path $DestPath     | Out-Null }
if (-not (Test-Path $CompletePath)) { New-Item -ItemType Directory -Path $CompletePath | Out-Null }

# ---- COUNTERS ----------------------------------------------
$Total   = $Archives.Count
$Success = 0
$Failed  = 0
$Skipped = 0
$Errors  = @()

Write-Host "Complete folder : $CompletePath" -ForegroundColor Gray
Write-Host ("-" * 60)

# ---- MAIN LOOP ---------------------------------------------
$i = 0
foreach ($Archive in $Archives) {
    $i++
    $ArchiveName = $Archive.Name
    $BaseName    = $Archive.BaseName
    $ExtractTo   = Join-Path $DestPath $BaseName

    Write-Host ""
    Write-Host "[$i/$Total] $ArchiveName" -ForegroundColor Cyan

    # Skip if output folder already exists
    if (Test-Path $ExtractTo) {
        Write-Host "  [SKIP] Output folder already exists: $ExtractTo" -ForegroundColor Yellow
        $Skipped++
        continue
    }

    # Create the target extraction folder
    New-Item -ItemType Directory -Path $ExtractTo | Out-Null

    $ExtractOK = $false

    try {
        Write-Host "  Extracting via 7-Zip..." -ForegroundColor Gray

        $pinfo = New-Object System.Diagnostics.ProcessStartInfo
        $pinfo.FileName               = $7zipPath
        $pinfo.Arguments              = "x `"$($Archive.FullName)`" -o`"$ExtractTo`" -y"
        $pinfo.RedirectStandardOutput = $true
        $pinfo.RedirectStandardError  = $true
        $pinfo.UseShellExecute        = $false
        $pinfo.CreateNoWindow         = $true

        $process = New-Object System.Diagnostics.Process
        $process.StartInfo = $pinfo
        $process.Start() | Out-Null

        $fileCount = 0
        while (-not $process.StandardOutput.EndOfStream) {
            $line = $process.StandardOutput.ReadLine()
            if ($line -match '^Extracting\s+(.+)') {
                $fileCount++
                $shortName = $Matches[1].Trim()
                if ($shortName.Length -gt 50) { $shortName = "..." + $shortName.Substring($shortName.Length - 47) }
                Write-Host -NoNewline "`r  Files extracted: $fileCount  ($shortName)                        " -ForegroundColor Yellow
            }
        }

        $process.WaitForExit()

        if ($process.ExitCode -ne 0) {
            $errText = $process.StandardError.ReadToEnd().Trim()
            throw "7-Zip exit code $($process.ExitCode): $errText"
        }

        Write-Host "`r  Done - $fileCount file(s) extracted                                    " -ForegroundColor Green
        $ExtractOK = $true

    } catch {
        Write-Host ""
        Write-Host "  [ERROR] $_" -ForegroundColor Red
        $Errors += "[$ArchiveName] $_"
        $Failed++
        # Clean up empty/partial folder
        if (Test-Path $ExtractTo) {
            $itemCount = (Get-ChildItem $ExtractTo -Recurse -ErrorAction SilentlyContinue).Count
            if ($itemCount -eq 0) { Remove-Item -Path $ExtractTo -Force -Recurse -ErrorAction SilentlyContinue }
        }
        continue
    }

    if ($ExtractOK) {
        # Move original archive to Complete folder
        $CompleteDest = Join-Path $CompletePath $ArchiveName
        try {
            Move-Item -LiteralPath $Archive.FullName -Destination $CompleteDest -Force
            Write-Host "  Moved to Complete: $CompleteDest" -ForegroundColor DarkGreen
        } catch {
            Write-Host "  [WARN] Could not move archive to Complete: $_" -ForegroundColor Yellow
        }
        $Success++
    }
}

# ---- SUMMARY -----------------------------------------------
Write-Host ""
Write-Host ("=" * 60)
Write-Host "  SUMMARY" -ForegroundColor Cyan
Write-Host ("=" * 60)
Write-Host "  Total processed : $Total"
Write-Host "  Succeeded       : $Success" -ForegroundColor Green
Write-Host "  Skipped         : $Skipped" -ForegroundColor Yellow
Write-Host "  Failed          : $Failed"  -ForegroundColor $(if ($Failed -gt 0) { 'Red' } else { 'Gray' })

if ($Errors.Count -gt 0) {
    Write-Host ""
    Write-Host "  Errors:" -ForegroundColor Red
    foreach ($e in $Errors) { Write-Host "    $e" -ForegroundColor Red }
}

Write-Host ("=" * 60)
Write-Host ""