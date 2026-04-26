<#
.SYNOPSIS
    Archive Extractor with Fixed Path Handling
#>

[CmdletBinding()]
param(
    [int]$MaxParallelJobs = [Environment]::ProcessorCount
)

# Color-coded output functions
function Write-Success { param($Message) Write-Host $Message -ForegroundColor Green }
function Write-Info { param($Message) Write-Host $Message -ForegroundColor Cyan }
function Write-Error { param($Message) Write-Host $Message -ForegroundColor Red }
function Write-Warning { param($Message) Write-Host $Message -ForegroundColor Yellow }

# Statistics tracking
$script:Stats = @{
    Extracted = 0
    Failed = 0
    Deleted = 0
    Unblocked = 0
}

# Auto-detect 7-Zip installation
function Find-SevenZip {
    $commonPaths = @(
        "C:\Program Files\7-Zip\7z.exe",
        "C:\Program Files (x86)\7-Zip\7z.exe"
    )
    
    foreach ($path in $commonPaths) {
        if (Test-Path $path) { return $path }
    }
    return $null
}

# Check if file is blocked and unblock it
function Unblock-Archive {
    param($FilePath)
    
    try {
        $zoneFile = "$FilePath`:Zone.Identifier"
        if (Test-Path $zoneFile) {
            Write-Warning "🔓 Unblocking: $([System.IO.Path]::GetFileName($FilePath))"
            Unblock-File -Path $FilePath
            $script:Stats.Unblocked++
            return $true
        }
        return $false
    }
    catch {
        Write-Warning "Could not unblock $([System.IO.Path]::GetFileName($FilePath)): $($_.Exception.Message)"
        return $false
    }
}

# Get and validate directory path
function Get-ValidDirectory {
    param([string]$Prompt, [bool]$MustExist = $true)
    
    do {
        $path = Read-Host $Prompt
        if ([string]::IsNullOrWhiteSpace($path)) {
            Write-Error "Path cannot be empty."
            continue
        }
        
        $path = [System.IO.Path]::GetFullPath($path)
        
        if ($MustExist -and -not (Test-Path $path -PathType Container)) {
            Write-Error "Directory does not exist: $path"
            continue
        }
        
        if (-not $MustExist -and -not (Test-Path $path)) {
            try {
                New-Item -ItemType Directory -Path $path -Force | Out-Null
                Write-Success "Created directory: $path"
            }
            catch {
                Write-Error "Failed to create directory: $($_.Exception.Message)"
                continue
            }
        }
        return $path
    } while ($true)
}

# Extract single archive with proper path handling
function Extract-SingleArchive {
    param(
        $Archive,
        [string]$BaseDestination,
        [string]$SourceDir,
        [string]$SevenZipPath,
        [bool]$DeleteOriginal
    )
    
    try {
        # First, unblock the archive if needed
        $wasBlocked = Unblock-Archive -FilePath $Archive.FullName
        
        # Calculate destination path - simplified to avoid path issues
        $destFolder = Join-Path $BaseDestination $Archive.BaseName
        
        # Create destination directory
        if (-not (Test-Path $destFolder)) {
            New-Item -ItemType Directory -Path $destFolder -Force | Out-Null
        }
        
        Write-Info "📦 Extracting: $($Archive.Name)"
        Write-Verbose "Source: $($Archive.FullName)"
        Write-Verbose "Destination: $destFolder"
        
        # Use proper quoting for paths with spaces - this is the key fix!
        $arguments = @(
            'x',
            '-y',
            "-o`"$destFolder`"",
            "`"$($Archive.FullName)`""
        )
        
        Write-Verbose "7-Zip Command: $SevenZipPath $($arguments -join ' ')"
        
        # Extract using 7-Zip with properly quoted paths
        $process = Start-Process -FilePath $SevenZipPath -ArgumentList $arguments -Wait -NoNewWindow -PassThru -RedirectStandardOutput "extract_output.txt" -RedirectStandardError "extract_error.txt"
        
        # Read output for debugging
        $extractOutput = Get-Content "extract_output.txt" -Raw -ErrorAction SilentlyContinue
        $extractError = Get-Content "extract_error.txt" -Raw -ErrorAction SilentlyContinue
        
        # Clean up temp files
        Remove-Item "extract_output.txt", "extract_error.txt" -Force -ErrorAction SilentlyContinue
        
        if ($process.ExitCode -eq 0) {
            Write-Success "✅ $($Archive.Name)"
            
            if ($DeleteOriginal) {
                try {
                    Remove-Item $Archive.FullName -Force
                    Write-Verbose "🗑️ Deleted: $($Archive.Name)"
                    return @{ Success = $true; Deleted = $true }
                }
                catch {
                    Write-Warning "Could not delete: $($Archive.Name) - $($_.Exception.Message)"
                    return @{ Success = $true; Deleted = $false }
                }
            }
            return @{ Success = $true; Deleted = $false }
        }
        else {
            Write-Error "❌ $($Archive.Name) - Exit code: $($process.ExitCode)"
            if ($extractOutput) { Write-Host "Output: $extractOutput" -ForegroundColor Gray }
            if ($extractError) { Write-Host "Error: $extractError" -ForegroundColor Red }
            return @{ Success = $false; Deleted = $false }
        }
    }
    catch {
        Write-Error "❌ $($Archive.Name) - Exception: $($_.Exception.Message)"
        return @{ Success = $false; Deleted = $false }
    }
}

# Test single archive first
function Test-SingleArchive {
    param($SevenZipPath, $TestArchive, $TempDestination)
    
    Write-Info "🧪 Testing with first archive: $($TestArchive.Name)"
    
    try {
        $testFolder = Join-Path $TempDestination "test_$([System.Guid]::NewGuid().ToString('N')[0..7] -join '')"
        New-Item -ItemType Directory -Path $testFolder -Force | Out-Null
        
        $arguments = @(
            'x',
            '-y',
            "-o`"$testFolder`"",
            "`"$($TestArchive.FullName)`""
        )
        
        Write-Info "Test command: $SevenZipPath $($arguments -join ' ')"
        
        $process = Start-Process -FilePath $SevenZipPath -ArgumentList $arguments -Wait -NoNewWindow -PassThru
        
        $success = ($process.ExitCode -eq 0)
        
        if ($success) {
            Write-Success "✅ Test extraction successful!"
            $extractedFiles = Get-ChildItem -Path $testFolder -Recurse | Measure-Object
            Write-Info "Extracted $($extractedFiles.Count) file(s)"
        } else {
            Write-Error "❌ Test extraction failed with exit code: $($process.ExitCode)"
        }
        
        # Clean up test folder
        Remove-Item $testFolder -Recurse -Force -ErrorAction SilentlyContinue
        
        return $success
    }
    catch {
        Write-Error "Test failed: $($_.Exception.Message)"
        return $false
    }
}

# Unblock all files in directory
function Unblock-AllFiles {
    param([string]$Directory)
    
    Write-Info "🔍 Checking for blocked files in: $Directory"
    
    $blockedCount = 0
    $allFiles = Get-ChildItem -Path $Directory -File -Recurse -Include *.zip,*.rar,*.7z
    
    foreach ($file in $allFiles) {
        if (Unblock-Archive -FilePath $file.FullName) {
            $blockedCount++
        }
    }
    
    if ($blockedCount -gt 0) {
        Write-Success "🔓 Unblocked $blockedCount file(s)"
    } else {
        Write-Info "✅ No blocked files found"
    }
}

# Sequential extraction
function Start-SequentialExtraction {
    param(
        [array]$Archives,
        [string]$BaseDestination,
        [string]$SourceDir,
        [string]$SevenZipPath,
        [bool]$DeleteOriginals
    )
    
    Write-Info "🔄 Sequential processing mode"
    $processed = 0
    
    foreach ($archive in $Archives) {
        $processed++
        Write-Info "[$processed/$($Archives.Count)] Processing: $($archive.Name)"
        
        $result = Extract-SingleArchive -Archive $archive -BaseDestination $BaseDestination -SourceDir $SourceDir -SevenZipPath $SevenZipPath -DeleteOriginal $DeleteOriginals
        
        if ($result.Success) {
            $script:Stats.Extracted++
            if ($result.Deleted) { $script:Stats.Deleted++ }
        }
        else {
            $script:Stats.Failed++
        }
        
        # Show progress
        $percent = [Math]::Round(($processed / $Archives.Count) * 100, 1)
        Write-Host "Progress: $percent% ($processed/$($Archives.Count))" -ForegroundColor Yellow
    }
}

# MAIN SCRIPT
Clear-Host
Write-Host "🚀 Archive Extractor with Fixed Path Handling" -ForegroundColor Yellow
Write-Host "=============================================" -ForegroundColor Yellow

Write-Info "💻 System: $($env:NUMBER_OF_PROCESSORS) cores, PowerShell $($PSVersionTable.PSVersion)"

# Auto-detect 7-Zip
Write-Info "🔍 Detecting 7-Zip..."
$sevenZipPath = Find-SevenZip

if (-not $sevenZipPath) {
    Write-Error "❌ 7-Zip not found. Please install from: https://www.7-zip.org/"
    Read-Host "Press Enter to exit"
    exit 1
}

Write-Success "✅ Found 7-Zip: $sevenZipPath"

# Get directories
Write-Host ""
$sourceDir = Get-ValidDirectory -Prompt "📁 Source directory" -MustExist $true
$destDir = Get-ValidDirectory -Prompt "📁 Destination directory" -MustExist $false

# Auto-unblock files first
Write-Host ""
Write-Info "🔓 Auto-unblocking files..."
Unblock-AllFiles -Directory $sourceDir

# Find archives
Write-Host ""
Write-Info "🔍 Scanning..."
$allArchives = Get-ChildItem -Path $sourceDir -File -Include *.zip,*.rar,*.7z -Recurse

if ($allArchives.Count -eq 0) {
    Write-Info "❌ No archives found"
    Read-Host "Press Enter to exit"
    exit 0
}

$totalSize = ($allArchives | Measure-Object Length -Sum).Sum / 1MB
Write-Success "📊 Found: $($allArchives.Count) archives ($([Math]::Round($totalSize, 2)) MB)"

# Test with first archive
Write-Host ""
if (-not (Test-SingleArchive -SevenZipPath $sevenZipPath -TestArchive $allArchives[0] -TempDestination $destDir)) {
    Write-Error "❌ Test extraction failed. Please check your archives and paths."
    Read-Host "Press Enter to exit"
    exit 1
}

# Configuration
Write-Host ""
do {
    $deleteChoice = Read-Host "🗑️ Delete originals after extraction? (y/n)"
    $deleteOriginals = $deleteChoice -match '^[Yy]'
} while ($deleteChoice -notmatch '^[YyNn]$')

# Show summary
Write-Host ""
Write-Host "📋 Summary:" -ForegroundColor Yellow
Write-Host "  Source: $sourceDir"
Write-Host "  Destination: $destDir"
Write-Host "  Archives: $($allArchives.Count)"
Write-Host "  Delete originals: $(if($deleteOriginals){'Yes'}else{'No'})"
Write-Host "  Files unblocked: $($script:Stats.Unblocked)"

$confirm = Read-Host "`n🚀 Start extraction? (y/n)"
if ($confirm -notmatch '^[Yy]$') { exit 0 }

# Start extraction
$startTime = Get-Date
Write-Host ""

Start-SequentialExtraction -Archives $allArchives -BaseDestination $destDir -SourceDir $sourceDir -SevenZipPath $sevenZipPath -DeleteOriginals $deleteOriginals

# Process nested archives
Write-Host ""
Write-Info "🔍 Checking for nested archives..."
$roundNumber = 1
do {
    $nestedArchives = Get-ChildItem -Path $destDir -File -Include *.zip,*.rar,*.7z -Recurse
    if ($nestedArchives.Count -eq 0) { break }
    
    Write-Info "📦 Round ${roundNumber}: Found $($nestedArchives.Count) nested archives"
    
    # Unblock nested archives too
    foreach ($nested in $nestedArchives) {
        Unblock-Archive -FilePath $nested.FullName | Out-Null
    }
    
    Start-SequentialExtraction -Archives $nestedArchives -BaseDestination $destDir -SourceDir $destDir -SevenZipPath $sevenZipPath -DeleteOriginals $deleteOriginals
    
    $roundNumber++
} while ($nestedArchives.Count -gt 0 -and $roundNumber -le 5)

# Results
$endTime = Get-Date
$duration = $endTime - $startTime
$speed = if ($duration.TotalSeconds -gt 0) { [Math]::Round($totalSize / $duration.TotalSeconds, 2) } else { 0 }

Write-Host ""
Write-Host "🎉 Extraction Complete!" -ForegroundColor Green
Write-Host "========================" -ForegroundColor Green
Write-Host "  ⏱️ Time: $($duration.ToString('mm\:ss'))"
Write-Host "  🚀 Speed: $speed MB/sec"
Write-Host "  🔓 Unblocked: $($script:Stats.Unblocked)" -ForegroundColor Yellow
Write-Host "  ✅ Success: $($script:Stats.Extracted)" -ForegroundColor Green
Write-Host "  ❌ Failed: $($script:Stats.Failed)" -ForegroundColor $(if($script:Stats.Failed -gt 0){'Red'}else{'Green'})
if ($deleteOriginals) {
    Write-Host "  🗑️ Deleted: $($script:Stats.Deleted)" -ForegroundColor Green
}

Read-Host "`nPress Enter to exit"