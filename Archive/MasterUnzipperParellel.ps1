<#
.SYNOPSIS
    Extract-NestedArchives.ps1 - Fast parallel extraction of nested archives
.DESCRIPTION
    Interactively prompts for source and destination directories, then extracts
    archives in parallel for maximum speed. Handles nested archives intelligently.
    Requires: 7-Zip and PowerShell 7+ for best parallel performance
#>

[CmdletBinding()]
param(
    [string]$SevenZipExe = "7z.exe",
    [int]$MaxParallelJobs = [Environment]::ProcessorCount
)

# Thread-safe statistics using synchronized hashtable
$script:Stats = [System.Collections.Hashtable]::Synchronized(@{
    Extracted = 0
    Failed = 0
    Deleted = 0
    Processing = 0
})

# Color-coded output functions (thread-safe)
function Write-Success { param($Message) Write-Host $Message -ForegroundColor Green }
function Write-Info { param($Message) Write-Host $Message -ForegroundColor Cyan }
function Write-Error { param($Message) Write-Host $Message -ForegroundColor Red }
function Write-Progress { param($Message) Write-Host $Message -ForegroundColor Yellow }

# Validate 7-Zip installation
function Test-SevenZip {
    param([string]$Path)
    try {
        $null = & $Path 2>$null
        return $true
    }
    catch { return $false }
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

# Fast parallel archive extraction
function Start-ParallelExtraction {
    param(
        [array]$Archives,
        [string]$BaseDestination,
        [string]$SourceDir,
        [string]$SevenZipPath,
        [bool]$DeleteOriginals,
        [int]$MaxJobs
    )
    
    # Check PowerShell version for optimal parallel support
    $useModernParallel = $PSVersionTable.PSVersion.Major -ge 7
    
    if ($useModernParallel) {
        Write-Info "Using PowerShell 7+ parallel processing (fast mode)"
        Start-ModernParallelExtraction @PSBoundParameters
    } else {
        Write-Info "Using PowerShell 5.x job-based processing (compatible mode)"
        Start-JobBasedExtraction @PSBoundParameters
    }
}

# PowerShell 7+ ForEach-Object -Parallel (fastest)
function Start-ModernParallelExtraction {
    param($Archives, $BaseDestination, $SourceDir, $SevenZipPath, $DeleteOriginals, $MaxJobs)
    
    $Archives | ForEach-Object -Parallel {
        # Import variables into parallel scope
        $archive = $_
        $baseDestination = $using:BaseDestination
        $sourceDir = $using:SourceDir
        $sevenZipPath = $using:SevenZipPath
        $deleteOriginals = $using:DeleteOriginals
        $stats = $using:script:Stats
        
        # Extract single archive
        try {
            $stats.Processing++
            
            # Create destination path
            $relativePath = $archive.Directory.FullName.Substring($sourceDir.Length).TrimStart('\', '/')
            $destFolder = Join-Path $baseDestination $relativePath
            $destFolder = Join-Path $destFolder $archive.BaseName
            
            if (-not (Test-Path $destFolder)) {
                New-Item -ItemType Directory -Path $destFolder -Force | Out-Null
            }
            
            Write-Host "⚡ [Thread $([Threading.Thread]::CurrentThread.ManagedThreadId)] Extracting: $($archive.Name)" -ForegroundColor Cyan
            
            # Extract using 7-Zip
            $process = Start-Process -FilePath $sevenZipPath -ArgumentList @(
                "x", "-y", "-o$destFolder", $archive.FullName
            ) -Wait -NoNewWindow -PassThru -RedirectStandardError "error_$([Guid]::NewGuid()).txt"
            
            if ($process.ExitCode -eq 0) {
                $stats.Extracted++
                Write-Host "✅ [Thread $([Threading.Thread]::CurrentThread.ManagedThreadId)] Success: $($archive.Name)" -ForegroundColor Green
                
                if ($deleteOriginals) {
                    Remove-Item $archive.FullName -Force -ErrorAction SilentlyContinue
                    $stats.Deleted++
                }
            } else {
                $stats.Failed++
                Write-Host "❌ [Thread $([Threading.Thread]::CurrentThread.ManagedThreadId)] Failed: $($archive.Name)" -ForegroundColor Red
            }
        }
        catch {
            $stats.Failed++
            Write-Host "💥 [Thread $([Threading.Thread]::CurrentThread.ManagedThreadId)] Error: $($archive.Name) - $($_.Exception.Message)" -ForegroundColor Red
        }
        finally {
            $stats.Processing--
        }
    } -ThrottleLimit $MaxJobs
}

# PowerShell 5.x job-based processing (compatible)
function Start-JobBasedExtraction {
    param($Archives, $BaseDestination, $SourceDir, $SevenZipPath, $DeleteOriginals, $MaxJobs)
    
    $jobQueue = @()
    $runningJobs = @()
    $archiveQueue = [System.Collections.Queue]::new($Archives)
    
    while ($archiveQueue.Count -gt 0 -or $runningJobs.Count -gt 0) {
        # Start new jobs if slots available
        while ($runningJobs.Count -lt $MaxJobs -and $archiveQueue.Count -gt 0) {
            $archive = $archiveQueue.Dequeue()
            
            $job = Start-Job -ScriptBlock {
                param($Archive, $BaseDestination, $SourceDir, $SevenZipPath, $DeleteOriginals)
                
                try {
                    # Create destination path
                    $relativePath = $Archive.Directory.FullName.Substring($SourceDir.Length).TrimStart('\', '/')
                    $destFolder = Join-Path $BaseDestination $relativePath
                    $destFolder = Join-Path $destFolder $Archive.BaseName
                    
                    if (-not (Test-Path $destFolder)) {
                        New-Item -ItemType Directory -Path $destFolder -Force | Out-Null
                    }
                    
                    # Extract using 7-Zip
                    $process = Start-Process -FilePath $SevenZipPath -ArgumentList @(
                        "x", "-y", "-o$destFolder", $Archive.FullName
                    ) -Wait -NoNewWindow -PassThru
                    
                    $result = @{
                        Archive = $Archive.Name
                        Success = ($process.ExitCode -eq 0)
                        DestFolder = $destFolder
                        ArchivePath = $Archive.FullName
                    }
                    
                    return $result
                } catch {
                    return @{
                        Archive = $Archive.Name
                        Success = $false
                        Error = $_.Exception.Message
                    }
                }
            } -ArgumentList $archive, $BaseDestination, $SourceDir, $SevenZipPath, $DeleteOriginals
            
            $runningJobs += $job
            Write-Progress "🚀 Started job for: $($archive.Name) (Active jobs: $($runningJobs.Count))"
        }
        
        # Check completed jobs
        $completedJobs = $runningJobs | Where-Object { $_.State -eq 'Completed' }
        
        foreach ($job in $completedJobs) {
            $result = Receive-Job $job
            Remove-Job $job
            
            if ($result.Success) {
                $script:Stats.Extracted++
                Write-Success "✅ Completed: $($result.Archive)"
                
                if ($DeleteOriginals -and $result.ArchivePath) {
                    Remove-Item $result.ArchivePath -Force -ErrorAction SilentlyContinue
                    $script:Stats.Deleted++
                }
            } else {
                $script:Stats.Failed++
                Write-Error "❌ Failed: $($result.Archive) - $($result.Error)"
            }
        }
        
        # Remove completed jobs from running list
        $runningJobs = $runningJobs | Where-Object { $_.State -ne 'Completed' }
        
        # Brief pause to prevent CPU spinning
        Start-Sleep -Milliseconds 100
    }
}

# Handle nested archives after parallel extraction
function Process-NestedArchives {
    param($BaseDestination, $SourceDir, $SevenZipPath, $DeleteOriginals)
    
    Write-Info "🔍 Scanning for nested archives..."
    
    do {
        $nestedArchives = Get-ChildItem -Path $BaseDestination -File -Include *.zip,*.rar,*.7z -Recurse
        
        if ($nestedArchives.Count -eq 0) { break }
        
        Write-Info "📦 Found $($nestedArchives.Count) nested archive(s), processing..."
        Start-ParallelExtraction -Archives $nestedArchives -BaseDestination $BaseDestination -SourceDir $BaseDestination -SevenZipPath $SevenZipPath -DeleteOriginals $DeleteOriginals -MaxJobs $MaxParallelJobs
        
    } while ($nestedArchives.Count -gt 0)
}

# MAIN SCRIPT
Clear-Host
Write-Host "🚀 High-Speed Parallel Archive Extractor" -ForegroundColor Yellow
Write-Host "=========================================" -ForegroundColor Yellow

# Show system info
Write-Info "💻 System Info: $($env:NUMBER_OF_PROCESSORS) CPU cores, PowerShell $($PSVersionTable.PSVersion)"
Write-Info "⚡ Max parallel jobs: $MaxParallelJobs"

# Validate 7-Zip
Write-Info "🔧 Checking 7-Zip installation..."
if (-not (Test-SevenZip $SevenZipExe)) {
    Write-Error "7-Zip not found. Please ensure 7z.exe is in your PATH."
    $customPath = Read-Host "Enter full path to 7z.exe (or press Enter to exit)"
    if ([string]::IsNullOrWhiteSpace($customPath)) { exit 1 }
    
    if (-not (Test-SevenZip $customPath)) {
        Write-Error "7-Zip still not found at: $customPath"
        exit 1
    }
    $SevenZipExe = $customPath
}
Write-Success "✅ 7-Zip found: $SevenZipExe"

# Get directories
$sourceDir = Get-ValidDirectory -Prompt "📁 Enter source directory to scan for archives" -MustExist $true
$destDir = Get-ValidDirectory -Prompt "📁 Enter destination directory for extracted files" -MustExist $false

# Configuration
do {
    $deleteChoice = Read-Host "🗑️  Delete original archives after extraction? (y/n)"
    $deleteOriginals = $deleteChoice -match '^[Yy]'
} while ($deleteChoice -notmatch '^[YyNn]$')

# Find archives
Write-Info "🔍 Scanning for archives..."
$allArchives = Get-ChildItem -Path $sourceDir -File -Include *.zip,*.rar,*.7z -Recurse | Sort-Object Length -Descending

if ($allArchives.Count -eq 0) {
    Write-Info "❌ No archives found in: $sourceDir"
    exit 0
}

# Show summary
$totalSize = ($allArchives | Measure-Object Length -Sum).Sum / 1MB
Write-Success "📊 Found $($allArchives.Count) archive(s) totaling $([Math]::Round($totalSize, 2)) MB"

Write-Host "`n📋 Summary:" -ForegroundColor Yellow
Write-Host "  Source: $sourceDir"
Write-Host "  Destination: $destDir"
Write-Host "  Archives: $($allArchives.Count)"
Write-Host "  Total size: $([Math]::Round($totalSize, 2)) MB"
Write-Host "  Parallel jobs: $MaxParallelJobs"
Write-Host "  Delete originals: $(if($deleteOriginals){'Yes'}else{'No'})"

do {
    $confirm = Read-Host "`n🚀 Start high-speed extraction? (y/n)"
} while ($confirm -notmatch '^[YyNn]$')

if ($confirm -notmatch '^[Yy]$') {
    Write-Info "⏹️  Operation cancelled."
    exit 0
}

# Start parallel extraction
$startTime = Get-Date
Write-Progress "`n🚀 Starting parallel extraction..."

try {
    # Extract root level archives in parallel
    Start-ParallelExtraction -Archives $allArchives -BaseDestination $destDir -SourceDir $sourceDir -SevenZipPath $SevenZipExe -DeleteOriginals $deleteOriginals -MaxJobs $MaxParallelJobs
    
    # Process any nested archives found
    Process-NestedArchives -BaseDestination $destDir -SourceDir $sourceDir -SevenZipPath $SevenZipExe -DeleteOriginals $deleteOriginals
}
catch {
    Write-Error "💥 Unexpected error: $($_.Exception.Message)"
}

# Final summary
$endTime = Get-Date
$duration = $endTime - $startTime
$speed = if ($duration.TotalSeconds -gt 0) { [Math]::Round($totalSize / $duration.TotalSeconds, 2) } else { 0 }

Write-Host "`n🎉 Extraction Complete!" -ForegroundColor Green
Write-Host "========================" -ForegroundColor Green
Write-Host "  ⏱️  Time taken: $($duration.ToString('mm\:ss'))"
Write-Host "  📊 Speed: $speed MB/sec"
Write-Host "  ✅ Extracted: $($script:Stats.Extracted)" -ForegroundColor Green
Write-Host "  ❌ Failed: $($script:Stats.Failed)" -ForegroundColor $(if($script:Stats.Failed -gt 0){'Red'}else{'Green'})
if ($deleteOriginals) {
    Write-Host "  🗑️  Deleted: $($script:Stats.Deleted)" -ForegroundColor Green
}
Write-Host "  📁 Destination: $destDir"