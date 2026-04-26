<#
.SYNOPSIS
    Archive Extractor - Fast Parallel & Reliable Sequential Modes with Installer Extraction
#>

[CmdletBinding()]
param(
    [int]$MaxParallelJobs = [Environment]::ProcessorCount,
    [switch]$ForceSequential
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
    SetupExtracted = 0
    SetupFailed = 0
    PathTooLongErrors = 0
}

# Safe Get-ChildItem that handles long paths
function Get-ChildItemSafe {
    param(
        [string]$Path,
        [string[]]$Include = @(),
        [switch]$Recurse,
        [switch]$File,
        [switch]$Directory,
        [int]$MaxDepth = 10
    )
    
    $results = @()
    $directories = @($Path)
    $currentDepth = 0
    
    while ($directories.Count -gt 0 -and $currentDepth -lt $MaxDepth) {
        $nextDirectories = @()
        
        foreach ($dir in $directories) {
            try {
                # Skip if path is getting too long
                if ($dir.Length -gt 200) {
                    $script:Stats.PathTooLongErrors++
                    Write-Warning "⚠️ Skipping path (too long): $($dir.Substring(0, 50))..."
                    continue
                }
                
                $items = Get-ChildItem -Path $dir -ErrorAction SilentlyContinue
                
                foreach ($item in $items) {
                    # Check if item matches criteria
                    $match = $false
                    
                    if ($File -and -not $item.PSIsContainer) {
                        if ($Include.Count -eq 0) {
                            $match = $true
                        } else {
                            foreach ($pattern in $Include) {
                                if ($item.Name -like $pattern) {
                                    $match = $true
                                    break
                                }
                            }
                        }
                    } elseif ($Directory -and $item.PSIsContainer) {
                        $match = $true
                    } elseif (-not $File -and -not $Directory) {
                        if ($Include.Count -eq 0) {
                            $match = $true
                        } else {
                            foreach ($pattern in $Include) {
                                if ($item.Name -like $pattern) {
                                    $match = $true
                                    break
                                }
                            }
                        }
                    }
                    
                    if ($match) {
                        $results += $item
                    }
                    
                    # Add directories for next level if recursing
                    if ($Recurse -and $item.PSIsContainer) {
                        $nextDirectories += $item.FullName
                    }
                }
            }
            catch {
                $script:Stats.PathTooLongErrors++
                Write-Warning "⚠️ Skipping directory (error): $($dir.Substring(0, 50))..."
            }
        }
        
        $directories = $nextDirectories
        $currentDepth++
    }
    
    return $results
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

# Auto-detect InnoSetup Unpacker
function Find-InnoSetupUnpacker {
    $commonPaths = @(
        "C:\Program Files\InnoSetup\innounp.exe",
        "C:\Program Files (x86)\InnoSetup\innounp.exe",
        ".\tools\innounp.exe",
        ".\innounp.exe"
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
        # Skip if path is too long
        if ($FilePath.Length -gt 240) {
            return $false
        }
        
        $zoneFile = "$FilePath`:Zone.Identifier"
        if (Test-Path $zoneFile) {
            Unblock-File -Path $FilePath
            $script:Stats.Unblocked++
            return $true
        }
        return $false
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

# Detect installer and binary files with long path handling
function Find-ExtractableFiles {
    param([string]$Directory)
    
    Write-Info "🔍 Scanning for extractable installers and binaries..."
    
    $extractableFiles = @{
        SetupExe = @()
        BinFiles = @()
        IsoFiles = @()
        MsiFiles = @()
    }
    
    # Find setup.exe files
    try {
        $setupFiles = Get-ChildItemSafe -Path $Directory -Include "setup.exe" -Recurse -File -MaxDepth 8
        $extractableFiles.SetupExe += $setupFiles
    }
    catch {
        Write-Warning "Error scanning for setup.exe files: $($_.Exception.Message)"
    }
    
    # Find .bin files
    try {
        $binFiles = Get-ChildItemSafe -Path $Directory -Include "*.bin" -Recurse -File -MaxDepth 8
        $extractableFiles.BinFiles += $binFiles
    }
    catch {
        Write-Warning "Error scanning for .bin files: $($_.Exception.Message)"
    }
    
    # Find .iso files
    try {
        $isoFiles = Get-ChildItemSafe -Path $Directory -Include "*.iso" -Recurse -File -MaxDepth 8
        $extractableFiles.IsoFiles += $isoFiles
    }
    catch {
        Write-Warning "Error scanning for .iso files: $($_.Exception.Message)"
    }
    
    # Find .msi files
    try {
        $msiFiles = Get-ChildItemSafe -Path $Directory -Include "*.msi" -Recurse -File -MaxDepth 8
        $extractableFiles.MsiFiles += $msiFiles
    }
    catch {
        Write-Warning "Error scanning for .msi files: $($_.Exception.Message)"
    }
    
    $totalFound = $extractableFiles.SetupExe.Count + $extractableFiles.BinFiles.Count + 
                  $extractableFiles.IsoFiles.Count + $extractableFiles.MsiFiles.Count
    
    if ($totalFound -gt 0) {
        Write-Success "📦 Found extractable files:"
        if ($extractableFiles.SetupExe.Count -gt 0) { Write-Info "  🔧 Setup.exe files: $($extractableFiles.SetupExe.Count)" }
        if ($extractableFiles.BinFiles.Count -gt 0) { Write-Info "  📀 BIN files: $($extractableFiles.BinFiles.Count)" }
        if ($extractableFiles.IsoFiles.Count -gt 0) { Write-Info "  💿 ISO files: $($extractableFiles.IsoFiles.Count)" }
        if ($extractableFiles.MsiFiles.Count -gt 0) { Write-Info "  📦 MSI files: $($extractableFiles.MsiFiles.Count)" }
    } else {
        Write-Info "ℹ️ No extractable installer files found"
    }
    
    if ($script:Stats.PathTooLongErrors -gt 0) {
        Write-Warning "⚠️ Skipped $($script:Stats.PathTooLongErrors) paths due to length limitations"
    }
    
    return $extractableFiles
}

# Extract setup.exe using multiple methods
function Extract-SetupFile {
    param(
        [System.IO.FileInfo]$SetupFile,
        [string]$DestinationBase,
        [string]$SevenZipPath,
        [string]$InnoUnpackerPath
    )
    
    # Check path length before proceeding
    if ($SetupFile.FullName.Length -gt 240) {
        Write-Warning "⚠️ Skipping setup.exe (path too long): $($SetupFile.Name)"
        return $false
    }
    
    $setupName = $SetupFile.BaseName
    $setupDestination = Join-Path $DestinationBase "Setup_Extracted\$setupName"
    
    # Ensure destination path isn't too long
    if ($setupDestination.Length -gt 200) {
        $setupName = $setupName.Substring(0, [Math]::Min(50, $setupName.Length))
        $setupDestination = Join-Path $DestinationBase "Setup_Extracted\$setupName"
    }
    
    # Create destination directory
    if (-not (Test-Path $setupDestination)) {
        try {
            New-Item -ItemType Directory -Path $setupDestination -Force | Out-Null
        }
        catch {
            Write-Warning "⚠️ Could not create destination for: $($SetupFile.Name)"
            return $false
        }
    }
    
    $success = $false
    $method = ""
    
    # Method 1: Try InnoSetup Unpacker first
    if ($InnoUnpackerPath -and -not $success) {
        try {
            Write-Info "🔧 Trying InnoSetup unpacker for: $($SetupFile.Name)"
            $arguments = @(
                "-x",
                "-d`"$setupDestination`"",
                "`"$($SetupFile.FullName)`""
            )
            
            $process = Start-Process -FilePath $InnoUnpackerPath -ArgumentList $arguments -Wait -NoNewWindow -PassThru
            
            if ($process.ExitCode -eq 0 -and (Get-ChildItem $setupDestination -Force -ErrorAction SilentlyContinue).Count -gt 0) {
                $success = $true
                $method = "InnoSetup Unpacker"
            }
        }
        catch {
            Write-Warning "InnoSetup unpacker failed: $($_.Exception.Message)"
        }
    }
    
    # Method 2: Try 7-Zip extraction
    if (-not $success) {
        try {
            Write-Info "🗜️ Trying 7-Zip extraction for: $($SetupFile.Name)"
            $arguments = @(
                'x',
                '-y',
                "-o`"$setupDestination`"",
                "`"$($SetupFile.FullName)`""
            )
            
            $process = Start-Process -FilePath $SevenZipPath -ArgumentList $arguments -Wait -NoNewWindow -PassThru
            
            if ($process.ExitCode -eq 0 -and (Get-ChildItem $setupDestination -Force -ErrorAction SilentlyContinue).Count -gt 0) {
                $success = $true
                $method = "7-Zip"
            }
        }
        catch {
            Write-Warning "7-Zip extraction failed: $($_.Exception.Message)"
        }
    }
    
    if ($success) {
        Write-Success "✅ Extracted setup.exe using $method`: $($SetupFile.Name)"
        return $true
    } else {
        Write-Warning "⚠️ Could not extract: $($SetupFile.Name)"
        # Remove empty destination directory
        if (Test-Path $setupDestination -and (Get-ChildItem $setupDestination -Force -ErrorAction SilentlyContinue).Count -eq 0) {
            Remove-Item $setupDestination -Force -Recurse -ErrorAction SilentlyContinue
        }
        return $false
    }
}

# Extract binary files
function Extract-BinaryFile {
    param(
        [System.IO.FileInfo]$BinaryFile,
        [string]$DestinationBase,
        [string]$SevenZipPath
    )
    
    # Check path length before proceeding
    if ($BinaryFile.FullName.Length -gt 240) {
        Write-Warning "⚠️ Skipping binary file (path too long): $($BinaryFile.Name)"
        return $false
    }
    
    $extension = $BinaryFile.Extension.ToLower()
    $fileName = $BinaryFile.BaseName
    $binaryDestination = Join-Path $DestinationBase "Binary_Extracted\$fileName"
    
    # Ensure destination path isn't too long
    if ($binaryDestination.Length -gt 200) {
        $fileName = $fileName.Substring(0, [Math]::Min(50, $fileName.Length))
        $binaryDestination = Join-Path $DestinationBase "Binary_Extracted\$fileName"
    }
    
    # Create destination directory
    if (-not (Test-Path $binaryDestination)) {
        try {
            New-Item -ItemType Directory -Path $binaryDestination -Force | Out-Null
        }
        catch {
            Write-Warning "⚠️ Could not create destination for: $($BinaryFile.Name)"
            return $false
        }
    }
    
    try {
        Write-Info "📀 Extracting $extension file: $($BinaryFile.Name)"
        
        $arguments = @(
            'x',
            '-y',
            "-o`"$binaryDestination`"",
            "`"$($BinaryFile.FullName)`""
        )
        
        $process = Start-Process -FilePath $SevenZipPath -ArgumentList $arguments -Wait -NoNewWindow -PassThru
        
        if ($process.ExitCode -eq 0 -and (Get-ChildItem $binaryDestination -Force -ErrorAction SilentlyContinue).Count -gt 0) {
            Write-Success "✅ Extracted $extension`: $($BinaryFile.Name)"
            return $true
        } else {
            Write-Warning "⚠️ Could not extract $extension`: $($BinaryFile.Name)"
            # Remove empty destination directory
            if (Test-Path $binaryDestination -and (Get-ChildItem $binaryDestination -Force -ErrorAction SilentlyContinue).Count -eq 0) {
                Remove-Item $binaryDestination -Force -Recurse -ErrorAction SilentlyContinue
            }
            return $false
        }
    }
    catch {
        Write-Error "❌ Error extracting $($BinaryFile.Name): $($_.Exception.Message)"
        return $false
    }
}

# Process all extractable files
function Start-InstallerExtraction {
    param(
        [hashtable]$ExtractableFiles,
        [string]$DestinationBase,
        [string]$SevenZipPath,
        [string]$InnoUnpackerPath
    )
    
    $totalFiles = $ExtractableFiles.SetupExe.Count + $ExtractableFiles.BinFiles.Count + 
                  $ExtractableFiles.IsoFiles.Count + $ExtractableFiles.MsiFiles.Count
    
    if ($totalFiles -eq 0) {
        return @{ SetupSuccess = 0; SetupFailed = 0; BinarySuccess = 0; BinaryFailed = 0 }
    }
    
    Write-Info "🚀 Starting installer/binary extraction ($totalFiles files)..."
    
    $extractionStats = @{
        SetupSuccess = 0
        SetupFailed = 0
        BinarySuccess = 0
        BinaryFailed = 0
    }
    
    # Extract setup.exe files
    if ($ExtractableFiles.SetupExe.Count -gt 0) {
        Write-Info "🔧 Processing $($ExtractableFiles.SetupExe.Count) setup.exe files..."
        
        foreach ($setupFile in $ExtractableFiles.SetupExe) {
            if (Extract-SetupFile -SetupFile $setupFile -DestinationBase $DestinationBase -SevenZipPath $SevenZipPath -InnoUnpackerPath $InnoUnpackerPath) {
                $extractionStats.SetupSuccess++
            } else {
                $extractionStats.SetupFailed++
            }
        }
    }
    
    # Extract binary files
    $allBinaryFiles = @()
    $allBinaryFiles += $ExtractableFiles.BinFiles
    $allBinaryFiles += $ExtractableFiles.IsoFiles
    $allBinaryFiles += $ExtractableFiles.MsiFiles
    
    if ($allBinaryFiles.Count -gt 0) {
        Write-Info "📀 Processing $($allBinaryFiles.Count) binary files..."
        
        foreach ($binaryFile in $allBinaryFiles) {
            if (Extract-BinaryFile -BinaryFile $binaryFile -DestinationBase $DestinationBase -SevenZipPath $SevenZipPath) {
                $extractionStats.BinarySuccess++
            } else {
                $extractionStats.BinaryFailed++
            }
        }
    }
    
    # Report extraction results
    Write-Host ""
    Write-Host "📊 Installer/Binary Extraction Results:" -ForegroundColor Yellow
    if ($ExtractableFiles.SetupExe.Count -gt 0) {
        Write-Host "  🔧 Setup.exe - Success: $($extractionStats.SetupSuccess), Failed: $($extractionStats.SetupFailed)" -ForegroundColor $(if($extractionStats.SetupFailed -gt 0){'Yellow'}else{'Green'})
    }
    if ($allBinaryFiles.Count -gt 0) {
        Write-Host "  📀 Binary files - Success: $($extractionStats.BinarySuccess), Failed: $($extractionStats.BinaryFailed)" -ForegroundColor $(if($extractionStats.BinaryFailed -gt 0){'Yellow'}else{'Green'})
    }
    
    return $extractionStats
}

# Fast parallel extraction (fixed with proper paths)
function Start-ParallelExtraction {
    param(
        [array]$Archives,
        [string]$BaseDestination,
        [string]$SevenZipPath,
        [bool]$DeleteOriginals,
        [int]$MaxJobs
    )
    
    Write-Info "⚡ Parallel mode: Processing $($Archives.Count) archives with $MaxJobs jobs"
    
    $Archives | ForEach-Object -Parallel {
        # Import variables into parallel scope
        $archive = $_
        $baseDestination = $using:BaseDestination
        $sevenZipPath = $using:SevenZipPath
        $deleteOriginals = $using:DeleteOriginals
        $stats = $using:script:Stats
        
        try {
            # Skip if path is too long
            if ($archive.FullName.Length -gt 240) {
                Write-Host "⚠️ [Thread $([Threading.Thread]::CurrentThread.ManagedThreadId)] Skipping (path too long): $($archive.Name)" -ForegroundColor Yellow
                return
            }
            
            # Unblock if needed
            $zoneFile = "$($archive.FullName)`:Zone.Identifier"
            if (Test-Path $zoneFile) {
                Unblock-File -Path $archive.FullName
                $stats.Unblocked++
            }
            
            # Create destination with shortened name if needed
            $destName = $archive.BaseName
            if ($destName.Length -gt 100) {
                $destName = $destName.Substring(0, 100)
            }
            $destFolder = Join-Path $baseDestination $destName
            
            if (-not (Test-Path $destFolder)) {
                New-Item -ItemType Directory -Path $destFolder -Force | Out-Null
            }
            
            # Prepare arguments with proper quoting
            $arguments = @(
                'x',
                '-y',
                "-o`"$destFolder`"",
                "`"$($archive.FullName)`""
            )
            
            Write-Host "⚡ [Thread $([Threading.Thread]::CurrentThread.ManagedThreadId)] Extracting: $($archive.Name)" -ForegroundColor Cyan
            
            # Extract
            $process = Start-Process -FilePath $sevenZipPath -ArgumentList $arguments -Wait -NoNewWindow -PassThru
            
            if ($process.ExitCode -eq 0) {
                $stats.Extracted++
                Write-Host "✅ [Thread $([Threading.Thread]::CurrentThread.ManagedThreadId)] Success: $($archive.Name)" -ForegroundColor Green
                
                if ($deleteOriginals) {
                    Remove-Item $archive.FullName -Force -ErrorAction SilentlyContinue
                    $stats.Deleted++
                }
            } else {
                $stats.Failed++
                Write-Host "❌ [Thread $([Threading.Thread]::CurrentThread.ManagedThreadId)] Failed: $($archive.Name) - Exit code: $($process.ExitCode)" -ForegroundColor Red
            }
        }
        catch {
            $stats.Failed++
            Write-Host "💥 [Thread $([Threading.Thread]::CurrentThread.ManagedThreadId)] Error: $($archive.Name) - $($_.Exception.Message)" -ForegroundColor Red
        }
    } -ThrottleLimit $MaxJobs
}

# Reliable sequential extraction
function Start-SequentialExtraction {
    param(
        [array]$Archives,
        [string]$BaseDestination,
        [string]$SevenZipPath,
        [bool]$DeleteOriginals
    )
    
    Write-Info "🔄 Sequential mode: Processing $($Archives.Count) archives one by one"
    $processed = 0
    
    foreach ($archive in $Archives) {
        $processed++
        
        try {
            # Skip if path is too long
            if ($archive.FullName.Length -gt 240) {
                Write-Warning "⚠️ Skipping (path too long): $($archive.Name)"
                continue
            }
            
            # Unblock if needed
            Unblock-Archive -FilePath $archive.FullName | Out-Null
            
            # Create destination with shortened name if needed
            $destName = $archive.BaseName
            if ($destName.Length -gt 100) {
                $destName = $destName.Substring(0, 100)
            }
            $destFolder = Join-Path $BaseDestination $destName
            
            if (-not (Test-Path $destFolder)) {
                New-Item -ItemType Directory -Path $destFolder -Force | Out-Null
            }
            
            Write-Info "[$processed/$($Archives.Count)] Extracting: $($archive.Name)"
            
            # Prepare arguments with proper quoting
            $arguments = @(
                'x',
                '-y',
                "-o`"$destFolder`"",
                "`"$($archive.FullName)`""
            )
            
            # Extract
            $process = Start-Process -FilePath $SevenZipPath -ArgumentList $arguments -Wait -NoNewWindow -PassThru
            
            if ($process.ExitCode -eq 0) {
                $script:Stats.Extracted++
                Write-Success "✅ $($archive.Name)"
                
                if ($DeleteOriginals) {
                    try {
                        Remove-Item $archive.FullName -Force
                        $script:Stats.Deleted++
                    }
                    catch {
                        Write-Warning "Could not delete: $($archive.Name)"
                    }
                }
            }
            else {
                $script:Stats.Failed++
                Write-Error "❌ $($archive.Name) - Exit code: $($process.ExitCode)"
            }
        }
        catch {
            $script:Stats.Failed++
            Write-Error "❌ $($archive.Name) - Exception: $($_.Exception.Message)"
        }
        
        # Show progress
        $percent = [Math]::Round(($processed / $Archives.Count) * 100, 1)
        Write-Host "Progress: $percent% ($processed/$($Archives.Count))" -ForegroundColor Yellow
    }
}

# Unblock all files in directory with long path handling
function Unblock-AllFiles {
    param([string]$Directory)
    
    Write-Info "🔍 Checking for blocked files in: $Directory"
    
    $blockedCount = 0
    
    try {
        $allFiles = Get-ChildItemSafe -Path $Directory -Include @("*.zip", "*.rar", "*.7z") -Recurse -File -MaxDepth 8
        
        foreach ($file in $allFiles) {
            if (Unblock-Archive -FilePath $file.FullName) {
                $blockedCount++
            }
        }
    }
    catch {
        Write-Warning "Error during unblock scan: $($_.Exception.Message)"
    }
    
    if ($blockedCount -gt 0) {
        Write-Success "🔓 Unblocked $blockedCount file(s)"
    } else {
        Write-Info "✅ No blocked files found"
    }
    
    if ($script:Stats.PathTooLongErrors -gt 0) {
        Write-Warning "⚠️ Skipped $($script:Stats.PathTooLongErrors) paths due to length limitations"
    }
}

# MAIN SCRIPT
Clear-Host
Write-Host "🚀 High-Speed Archive Extractor v4.1 with Installer Extraction" -ForegroundColor Yellow
Write-Host "=================================================================" -ForegroundColor Yellow

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
$allArchives = Get-ChildItemSafe -Path $sourceDir -Include @("*.zip", "*.rar", "*.7z") -Recurse -File -MaxDepth 8

if ($allArchives.Count -eq 0) {
    Write-Info "❌ No archives found"
    Read-Host "Press Enter to exit"
    exit 0
}

$totalSize = ($allArchives | Measure-Object Length -Sum).Sum / 1MB
Write-Success "📊 Found: $($allArchives.Count) archives ($([Math]::Round($totalSize, 2)) MB)"

# Choose processing mode
Write-Host ""
if (-not $ForceSequential -and $PSVersionTable.PSVersion.Major -ge 7) {
    Write-Host "🎯 Processing Mode Options:" -ForegroundColor Yellow
    Write-Host "  1. ⚡ PARALLEL (Fast) - Uses multiple CPU cores simultaneously"
    Write-Host "  2. 🔄 SEQUENTIAL (Reliable) - One at a time, slower but safer"
    Write-Host ""
    
    do {
        $modeChoice = Read-Host "Choose mode (1=Parallel/Fast, 2=Sequential/Reliable)"
        $useParallel = $modeChoice -eq "1"
    } while ($modeChoice -notmatch '^[12]$')
} else {
    Write-Info "Using sequential mode (PowerShell 5.x or forced)"
    $useParallel = $false
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
Write-Host "  Mode: $(if($useParallel){"⚡ PARALLEL ($MaxParallelJobs threads)"}else{"🔄 SEQUENTIAL"})"
Write-Host "  Archives: $($allArchives.Count)"
Write-Host "  Total size: $([Math]::Round($totalSize, 2)) MB"
Write-Host "  Delete originals: $(if($deleteOriginals){'Yes'}else{'No'})"

$confirm = Read-Host "`n🚀 Start extraction? (y/n)"
if ($confirm -notmatch '^[Yy]$') { exit 0 }

# Start extraction
$startTime = Get-Date
Write-Host ""
Write-Host "🔄 Phase 1: Archive Extraction" -ForegroundColor Yellow
Write-Host "==============================" -ForegroundColor Yellow

if ($useParallel) {
    Start-ParallelExtraction -Archives $allArchives -BaseDestination $destDir -SevenZipPath $sevenZipPath -DeleteOriginals $deleteOriginals -MaxJobs $MaxParallelJobs
} else {
    Start-SequentialExtraction -Archives $allArchives -BaseDestination $destDir -SevenZipPath $sevenZipPath -DeleteOriginals $deleteOriginals
}

# Process nested archives
Write-Host ""
Write-Host "🔄 Phase 2: Nested Archive Processing" -ForegroundColor Yellow
Write-Host "=====================================" -ForegroundColor Yellow
Write-Info "🔍 Checking for nested archives..."
$roundNumber = 1
do {
    $nestedArchives = Get-ChildItemSafe -Path $destDir -Include @("*.zip", "*.rar", "*.7z") -Recurse -File -MaxDepth 6
    if ($nestedArchives.Count -eq 0) { break }
    
    Write-Info "📦 Round $roundNumber`: Found $($nestedArchives.Count) nested archives"
    
    if ($useParallel) {
        Start-ParallelExtraction -Archives $nestedArchives -BaseDestination $destDir -SevenZipPath $sevenZipPath -DeleteOriginals $deleteOriginals -MaxJobs $MaxParallelJobs
    } else {
        Start-SequentialExtraction -Archives $nestedArchives -BaseDestination $destDir -SevenZipPath $sevenZipPath -DeleteOriginals $deleteOriginals
    }
    
    $roundNumber++
} while ($nestedArchives.Count -gt 0 -and $roundNumber -le 5)

# INSTALLER AND BINARY EXTRACTION PHASE
Write-Host ""
Write-Host "🔧 Phase 3: Installer & Binary Extraction" -ForegroundColor Yellow
Write-Host "========================================" -ForegroundColor Yellow

# Auto-detect InnoSetup Unpacker
$innoUnpackerPath = Find-InnoSetupUnpacker
if ($innoUnpackerPath) {
    Write-Success "✅ Found InnoSetup Unpacker: $innoUnpackerPath"
} else {
    Write-Warning "⚠️ InnoSetup Unpacker not found. Download from: https://innounp.sourceforge.io/"
    Write-Info "   Setup.exe extraction will use 7-Zip only (less reliable)"
}

# Find extractable files
$extractableFiles = Find-ExtractableFiles -Directory $destDir

# Ask user if they want to extract installers/binaries
if (($extractableFiles.SetupExe.Count + $extractableFiles.BinFiles.Count + $extractableFiles.IsoFiles.Count + $extractableFiles.MsiFiles.Count) -gt 0) {
    Write-Host ""
    $extractChoice = Read-Host "🔧 Extract setup.exe and binary files? (y/n)"
    
    if ($extractChoice -match '^[Yy]$') {
        $installerStats = Start-InstallerExtraction -ExtractableFiles $extractableFiles -DestinationBase $destDir -SevenZipPath $sevenZipPath -InnoUnpackerPath $innoUnpackerPath
        
        # Update global stats
        $script:Stats.SetupExtracted = $installerStats.SetupSuccess + $installerStats.BinarySuccess
        $script:Stats.SetupFailed = $installerStats.SetupFailed + $installerStats.BinaryFailed
    }
}

# Results
$endTime = Get-Date
$duration = $endTime - $startTime
$speed = if ($duration.TotalSeconds -gt 0) { [Math]::Round($totalSize / $duration.TotalSeconds, 2) } else { 0 }

Write-Host ""
Write-Host "🎉 Complete Extraction Results!" -ForegroundColor Green
Write-Host "===============================" -ForegroundColor Green
Write-Host "  ⏱️ Total Time: $($duration.ToString('mm\:ss'))"
Write-Host "  🚀 Speed: $speed MB/sec"
Write-Host "  🔧 Mode: $(if($useParallel){'Parallel'}else{'Sequential'})"
Write-Host ""
Write-Host "📊 Archive Extraction:" -ForegroundColor Cyan
Write-Host "  🔓 Unblocked: $($script:Stats.Unblocked)" -ForegroundColor Yellow
Write-Host "  ✅ Success: $($script:Stats.Extracted)" -ForegroundColor Green
Write-Host "  ❌ Failed: $($script:Stats.Failed)" -ForegroundColor $(if($script:Stats.Failed -gt 0){'Red'}else{'Green'})
if ($deleteOriginals) {
    Write-Host "  🗑️ Deleted: $($script:Stats.Deleted)" -ForegroundColor Green
}

if ($script:Stats.SetupExtracted -or $script:Stats.SetupFailed) {
    Write-Host ""
    Write-Host "📊 Installer/Binary Extraction:" -ForegroundColor Cyan
    Write-Host "  ✅ Extracted: $($script:Stats.SetupExtracted)" -ForegroundColor Green
    Write-Host "  ❌ Failed: $($script:Stats.SetupFailed)" -ForegroundColor $(if($script:Stats.SetupFailed -gt 0){'Red'}else{'Green'})
}

if ($script:Stats.PathTooLongErrors -gt 0) {
    Write-Host ""
    Write-Host "⚠️ Path Length Issues:" -ForegroundColor Yellow
    Write-Host "  Skipped paths: $($script:Stats.PathTooLongErrors)" -ForegroundColor Yellow
    Write-Host "  (Windows 260-character limit reached)" -ForegroundColor Yellow
}

Write-Host ""
Write-Host "📁 Organized Structure:" -ForegroundColor Cyan
Write-Host "  📦 Archives → Individual folders"
Write-Host "  🔧 Setup.exe → Setup_Extracted\"
Write-Host "  📀 BIN/ISO/MSI → Binary_Extracted\"

Read-Host "`nPress Enter to exit"