# LaunchBox Batch File Generator - Enhanced Version with User Input

# Get user input for the root game folder
do {
    $rootGameFolder = Read-Host "Enter the path to your games folder"
    
    if (-not $rootGameFolder) {
        Write-Host "Please enter a valid path." -ForegroundColor Red
        continue
    }
    
    # Remove quotes if user included them
    $rootGameFolder = $rootGameFolder.Trim('"')
    
    if (-not (Test-Path $rootGameFolder)) {
        Write-Host "Path does not exist: $rootGameFolder" -ForegroundColor Red
        Write-Host "Please enter a valid folder path." -ForegroundColor Yellow
        continue
    }
    
    # Check if it's actually a directory
    if (-not (Get-Item $rootGameFolder).PSIsContainer) {
        Write-Host "Path is not a directory: $rootGameFolder" -ForegroundColor Red
        continue
    }
    
    break
} while ($true)

Write-Host "Scanning folder: $rootGameFolder" -ForegroundColor Green

$batchOutputFolder = "$rootGameFolder\_BatchFiles"
$maxDepth = 8  # Increased from 5 to 8 to handle deeper folder structures
$folderTimeout = 300  # Increased from 60 to 300 seconds per game folder
$logFound = "$batchOutputFolder\FoundGames.log"
$logNotFound = "$batchOutputFolder\NotFoundGames.log"
$logSkipped = "$batchOutputFolder\SkippedGames.log"  
$logErrors = "$batchOutputFolder\ErrorGames.log"
$logBlocked = "$batchOutputFolder\BlockedFiles.log"  # New log for blocked files debug

# Ensure batch output folder exists
if (!(Test-Path $batchOutputFolder)) {
    New-Item -ItemType Directory -Path $batchOutputFolder | Out-Null
}

# Clear logs if they exist
Remove-Item -Path $logFound, $logNotFound, $logSkipped, $logErrors, $logBlocked -ErrorAction SilentlyContinue

# Store a hashtable of created batch files to check for duplicates
$createdBatchFiles = @{}

function Get-ExecutableScore {
    param($exePath, $gameFolderName)
    $exeName = [System.IO.Path]::GetFileNameWithoutExtension($exePath).ToLower()
    $folderName = $gameFolderName.ToLower()
    $score = 0
    
    # Phase 1: Block files with these problematic words anywhere in the name (High-Value additions)
    $blockPatterns = @(
        "*uninstall*", "*setup*", "*settings*", "*helper*",
        "*config*", "*launcher*", "*language*", "*crash*", 
        "*test*", "*service*", "*server*", "*update*", "*install*"
    )
    foreach ($pattern in $blockPatterns) {
        if ($exeName -like $pattern) {
            # Log blocked files for debugging
            "$gameFolderName - BLOCKED by pattern '$pattern': $exePath" | Out-File -FilePath $logBlocked -Append
            return -1
        }
    }
    
    # Enhanced blacklist with additional exclusions
    $badNames = @(
        "unins000", 
        "crashreport", "errorreport", "crashreporter", "crashpad",
        "redist", "redistributable", "vcredist", "vc_redist",
        "directx", "dxwebsetup",
        "uploader", "webhelper", 
        "crs-handler", "crs-uploader", "crs-video",
        "drivepool", "quicksfv", "handler",
        "gamingrepair", "unitycrashhandle64"
    )
    
    if ($badNames -contains $exeName) {
        # Log blocked files for debugging
        "$gameFolderName - BLOCKED by blacklist: $exePath" | Out-File -FilePath $logBlocked -Append
        return -1
    }
    
    # HIGHEST PRIORITY: Exact match with folder name gets maximum score
    if ($exeName -eq $folderName) { 
        return 1000  # Guaranteed highest score
    }
    
    # HIGH PRIORITY: Check for shipping executables (Unreal Engine pattern)
    # Look for pattern like "gamename-win64-shipping"
    $shippingPattern = "$folderName-win64-shipping"
    if ($exeName -eq $shippingPattern) {
        return 500  # Very high score, but less than exact match
    }
    
    # Also check for variations of shipping executables
    if ($exeName -like "*$folderName*" -and $exeName -like "*win64*shipping*") {
        $score += 100
    }
    
    # ENHANCED: Check for "game" anywhere in executable name (fixes gsgameexe.exe, gameexe.exe, etc.)
    if ($exeName -like "*game*") { 
        $score += 75  # High bonus for any game-related executable
    }
    
    # Partial match is good
    if ($exeName -like "*$folderName*") { $score += 50 }
    
    # Game name as part of the executable name is promising
    $gameNameParts = $folderName -split ' '
    foreach ($part in $gameNameParts) {
        if ($part.Length -gt 3 -and $exeName -like "*$part*") {
            $score += 20
        }
    }
    
    # Priority executable names (removed "launcher", enhanced "game" detection above)
    $priorityNames = @("start", "play", "run", "main", "bin")
    if ($priorityNames -contains $exeName) { $score += 30 }
    
    # Priority for executables in typical game folders
    $exeLocation = [System.IO.Path]::GetDirectoryName($exePath).ToLower()
    $goodPaths = @("\bin", "\binaries", "\game", "\app", "\win64", "\win32", "\windows", "\x64", "\x86")
    foreach ($goodPath in $goodPaths) {
        if ($exeLocation -like "*$goodPath*") {
            $score += 20
            break
        }
    }
    
    # If it's extremely deep in subfolders, slightly lower priority
    $folderDepth = ($exePath.Split('\').Count - $rootGameFolder.Split('\').Count)
    if ($folderDepth -gt 6) {
        $score -= 10
    }
    
    return $score
}

function Create-BatchFile {
    param (
        [string]$targetExe,
        [string]$batchName,
        [string]$outputFolder
    )
    
    # Sanitize the batch file name by removing/replacing problematic characters
    $safeBatchName = $batchName -replace '[\\\/\:\*\?"<>\|]', '_'  # Replace illegal characters
    $safeBatchName = $safeBatchName -replace '&', 'and'  # Replace & with 'and'
    
    # Limit filename length to prevent path too long errors
    if ($safeBatchName.Length -gt 50) {
        $safeBatchName = $safeBatchName.Substring(0, 47) + "..."
    }
    
    # Check if this batch file already exists based on target executable
    $batchPath = "$outputFolder\$safeBatchName.bat"
    
    # Check if the batch file exists and contains the same target
    if (Test-Path $batchPath) {
        try {
            $existingContent = Get-Content $batchPath -Raw
            # Check if the existing batch file contains the same executable path
            if ($existingContent -match [regex]::Escape($targetExe)) {
                return "exists_same"
            } else {
                return "exists_different"
            }
        } catch {
            Write-Host "Error checking existing batch file: $_" -ForegroundColor Yellow
            return "error"
        }
    }
    
    # Check if we've already created a batch file to this executable
    if ($createdBatchFiles.ContainsKey($targetExe)) {
        return "duplicate"
    }
    
    # Create the batch file content
    try {
        # Get the directory of the executable for setting working directory
        $exeDirectory = [System.IO.Path]::GetDirectoryName($targetExe)
        $exeFileName = [System.IO.Path]::GetFileName($targetExe)
        
        # Create batch file content with proper formatting
        $batchContent = @"
@echo off
REM Auto-generated batch file for $batchName
REM Target: $targetExe

cd /d "$exeDirectory"
start "" "$exeFileName"
"@
        
        # Write the batch file
        $batchContent | Out-File -FilePath $batchPath -Encoding ASCII
        
        # Add to our tracking hashtable with original name for logging
        $createdBatchFiles[$targetExe] = $batchName
        
        return "created"
    } catch {
        # Try a more aggressive filename sanitization and shorter name
        try {
            $ultraSafeBatchName = "Game_" + ($batchName -replace '[^a-zA-Z0-9]', '_')
            if ($ultraSafeBatchName.Length -gt 30) {
                $ultraSafeBatchName = $ultraSafeBatchName.Substring(0, 30)
            }
            
            $batchPath = "$outputFolder\$ultraSafeBatchName.bat"
            
            # Get the directory of the executable for setting working directory
            $exeDirectory = [System.IO.Path]::GetDirectoryName($targetExe)
            $exeFileName = [System.IO.Path]::GetFileName($targetExe)
            
            # Create batch file content with proper formatting
            $batchContent = @"
@echo off
REM Auto-generated batch file for $batchName
REM Target: $targetExe

cd /d "$exeDirectory"
start "" "$exeFileName"
"@
            
            # Write the batch file
            $batchContent | Out-File -FilePath $batchPath -Encoding ASCII
            
            # Add to our tracking hashtable
            $createdBatchFiles[$targetExe] = $batchName
            
            # Return special status for sanitized name
            return "created_sanitized"
        } catch {
            Write-Host "Error creating batch file for $batchName`: $_" -ForegroundColor Yellow
            return "error"
        }
    }
}

function Format-TimeSpan {
    param (
        [TimeSpan]$TimeSpan
    )
    
    if ($TimeSpan.TotalHours -ge 1) {
        return "{0:h\:mm\:ss}" -f $TimeSpan
    } else {
        return "{0:mm\:ss}" -f $TimeSpan
    }
}

function Find-Executables {
    param (
        [string]$folderPath,
        [int]$maxDepth,
        [string]$gameName,
        [int]$timeout
    )
    
    $timeoutTime = (Get-Date).AddSeconds($timeout)
    Write-Progress -Id 2 -Activity "Finding executables" -Status "Scanning $gameName..."
    
    $exeFiles = @()
    
    try {
        # IMPORTANT: First check specifically for common game folder structures
        $commonGamePaths = @(
            "$folderPath\Game\*.exe",
            "$folderPath\app\*.exe",
            "$folderPath\bin\*.exe",
            "$folderPath\binaries\*.exe",
            "$folderPath\Windows\*.exe",
            "$folderPath\x64\*.exe",
            "$folderPath\Win64\*.exe",
            "$folderPath\executable\*.exe",
            "$folderPath\program\*.exe",
            "$folderPath\launcher\*.exe",
            "$folderPath\main\*.exe"
        )
        
        foreach ($path in $commonGamePaths) {
            if (Test-Path $path) {
                $foundExes = Get-ChildItem -Path $path -ErrorAction SilentlyContinue
                if ($foundExes -and $foundExes.Count -gt 0) {
                    Write-Host "Found executables in common path: $path" -ForegroundColor Green
                    $exeFiles += $foundExes
                }
            }
        }
        
        # Also check for deeper common structures (like Unreal Engine games)
        $deeperCommonPaths = @(
            "$folderPath\Engine\Binaries\Win64\*.exe",
            "$folderPath\Binaries\Win64\*.exe",
            "$folderPath\*\Binaries\Win64\*.exe",
            "$folderPath\*\*\Binaries\Win64\*.exe",
            "$folderPath\*\*\*\Binaries\Win64\*.exe"
        )
        
        foreach ($path in $deeperCommonPaths) {
            if ($exeFiles.Count -eq 0) {
                try {
                    $foundExes = Get-ChildItem -Path $path -ErrorAction SilentlyContinue
                    if ($foundExes -and $foundExes.Count -gt 0) {
                        Write-Host "Found executables in deeper path: $path" -ForegroundColor Green
                        $exeFiles += $foundExes
                    }
                } catch {
                    # Just continue if path is invalid
                }
            }
        }
        
        # If we already found executables in common paths, don't do the expensive searches
        if ($exeFiles.Count -gt 0) {
            Write-Progress -Id 2 -Activity "Finding executables" -Status "Found executables in common paths" -Completed
            return $exeFiles
        }
        
        # Continue with the regular search strategy
        # Try a simple search at the root level first
        $exeFiles = Get-ChildItem -Path $folderPath -Filter "*.exe" -File -ErrorAction SilentlyContinue
        
        # If no files found, search one level deeper
        if ($exeFiles.Count -eq 0) {
            $exeFiles = Get-ChildItem -Path $folderPath -Filter "*.exe" -File -Depth 1 -ErrorAction SilentlyContinue
        }
        
        # If still no files, try progressively deeper searches
        if ($exeFiles.Count -eq 0) {
            # Try depth 2
            $exeFiles = Get-ChildItem -Path $folderPath -Filter "*.exe" -File -Depth 2 -ErrorAction SilentlyContinue
        }
        
        if ($exeFiles.Count -eq 0) {
            # Try depth 4
            $exeFiles = Get-ChildItem -Path $folderPath -Filter "*.exe" -File -Depth 4 -ErrorAction SilentlyContinue
        }
        
        # If still no files, try the full max depth search
        if ($exeFiles.Count -eq 0) {
            $exeFiles = Get-ChildItem -Path $folderPath -Filter "*.exe" -File -Depth $maxDepth -ErrorAction SilentlyContinue
        }
        
        # Final full recursive search if needed, with timeout protection
        if ($exeFiles.Count -eq 0) {
            $scriptBlock = {
                param($path)
                Get-ChildItem -Path $path -Filter "*.exe" -File -Recurse -ErrorAction SilentlyContinue
            }
            
            $job = Start-Job -ScriptBlock $scriptBlock -ArgumentList $folderPath
            
            # Wait for job to complete or timeout
            $null = Wait-Job -Job $job -Timeout ($timeout - 5)
            
            if ($job.State -eq "Running") {
                Stop-Job -Job $job
                Write-Host "Timeout reached while scanning $gameName recursively." -ForegroundColor Yellow
            } else {
                $exeFiles = Receive-Job -Job $job
            }
            
            Remove-Job -Job $job -Force
        }
    }
    catch {
        Write-Host "Error scanning $folderPath for executables: $_" -ForegroundColor Red
    }
    
    Write-Progress -Id 2 -Activity "Finding executables" -Completed
    return $exeFiles
}

# Get list of game folders to process
Write-Host "Checking for game folders..." -ForegroundColor Yellow
$gameFolders = Get-ChildItem -Path $rootGameFolder -Directory

if ($gameFolders.Count -eq 0) {
    Write-Host "No subdirectories found in $rootGameFolder" -ForegroundColor Red
    Write-Host "Make sure the path contains game folders to scan." -ForegroundColor Yellow
    Read-Host "Press Enter to exit"
    exit
}

$total = $gameFolders.Count
$index = 0
$created = 0
$skipped = 0
$notFound = 0
$sanitized = 0
$errors = 0
$timeouts = 0

# Start timing
$startTime = Get-Date
$lastUpdateTime = $startTime

Write-Host "Starting to process $total game folders..." -ForegroundColor Green

foreach ($folder in $gameFolders) {
    $index++
    $gameName = $folder.Name
    
    # Calculate timing statistics
    $currentTime = Get-Date
    $elapsedTime = $currentTime - $startTime
    $itemsRemaining = $total - $index
    
    # Only recalculate estimated time every 5 items or every 10 seconds to avoid fluctuations
    if (($index % 5 -eq 0) -or (($currentTime - $lastUpdateTime).TotalSeconds -ge 10)) {
        $lastUpdateTime = $currentTime
        
        if ($index -gt 1) {  # Need at least 2 items to calculate average time
            $averageTimePerItem = $elapsedTime.TotalSeconds / ($index - 1)
            $estimatedTimeRemaining = [TimeSpan]::FromSeconds($averageTimePerItem * $itemsRemaining)
            $formattedTimeRemaining = Format-TimeSpan -TimeSpan $estimatedTimeRemaining
            $formattedElapsedTime = Format-TimeSpan -TimeSpan $elapsedTime
            
            $statusMessage = "$gameName ($index of $total) - $created created, $sanitized sanitized, $skipped skipped, $notFound not found, $errors errors"
            $progressStatus = "$statusMessage | Elapsed: $formattedElapsedTime | Remaining: $formattedTimeRemaining"
        } else {
            $progressStatus = "$gameName ($index of $total)"
        }
    } else {
        $progressStatus = "$gameName ($index of $total) - $created created, $sanitized sanitized, $skipped skipped, $notFound not found, $errors errors"
    }
    
    Write-Progress -Id 1 -Activity "Scanning Games" -Status $progressStatus -PercentComplete (($index / $total) * 100)
    
    # Special handling for known problematic games
    $manualExePath = $null
    if ($gameName -eq "ELDEN RING") {
        $specificPath = "$($folder.FullName)\Game\eldenring.exe"
        if (Test-Path $specificPath) {
            Write-Host "Found Elden Ring executable via direct path!" -ForegroundColor Green
            $manualExePath = $specificPath
        }
    }
    
    # Use our new function to find executables with improved recursive search
    try {
        $exeFiles = @()
        
        # Use manual path if we have one
        if ($manualExePath) {
            $exeFiles = @(Get-Item -Path $manualExePath)
        } else {
            $searchStartTime = Get-Date
            $exeFiles = Find-Executables -folderPath $folder.FullName -maxDepth $maxDepth -gameName $gameName -timeout $folderTimeout
            $searchTime = (Get-Date) - $searchStartTime
            
            # Check if we likely hit a timeout (over 90% of the timeout time used)
            if ($searchTime.TotalSeconds -gt ($folderTimeout * 0.9)) {
                Write-Host "Warning: $gameName search took $($searchTime.TotalSeconds) seconds" -ForegroundColor Yellow
                $timeouts++
                # Log the timeout to errors log
                "$gameName - Timeout reached while scanning for executables" | Out-File -FilePath $logErrors -Append
            }
        }
        
        if ($exeFiles.Count -eq 0) {
            "$gameName - No executable found" | Out-File -FilePath $logNotFound -Append
            $notFound++
            continue
        }
        
        $scored = $exeFiles | ForEach-Object {
            [PSCustomObject]@{
                Path  = $_.FullName
                Score = Get-ExecutableScore -exePath $_.FullName -gameFolderName $gameName
            }
        } | Where-Object { $_.Score -ge 0 } | Sort-Object Score -Descending
        
        if ($scored.Count -eq 0) {
            "$gameName - No suitable exe (all files blocked by filters)" | Out-File -FilePath $logNotFound -Append
            # Also log what files were found but blocked
            foreach ($exe in $exeFiles) {
                $exeName = [System.IO.Path]::GetFileNameWithoutExtension($exe.FullName).ToLower()
                "$gameName - Found but filtered: $($exe.FullName) (score would be negative)" | Out-File -FilePath $logBlocked -Append
            }
            $notFound++
            continue
        }
        
        $chosenExe = $scored[0].Path
        $batchStatus = Create-BatchFile -targetExe $chosenExe -batchName $gameName -outputFolder $batchOutputFolder
        
        # Log based on the result
        switch ($batchStatus) {
            "created" {
                "$gameName - $chosenExe" | Out-File -FilePath $logFound -Append
                $created++
            }
            "created_sanitized" {
                "$gameName - $chosenExe (with sanitized name)" | Out-File -FilePath $logFound -Append
                $sanitized++
            }
            "exists_same" {
                "$gameName - Batch file exists and targets same executable: $chosenExe" | Out-File -FilePath $logSkipped -Append
                $skipped++
            }
            "exists_different" {
                "$gameName - Batch file exists but targets different executable. New: $chosenExe" | Out-File -FilePath $logSkipped -Append
                $skipped++
            }
            "duplicate" {
                "$gameName - Duplicate executable already used for: $($createdBatchFiles[$chosenExe])" | Out-File -FilePath $logSkipped -Append
                $skipped++
            }
            "error" {
                "$gameName - Error creating batch file for: $chosenExe" | Out-File -FilePath $logErrors -Append
                $errors++
            }
        }
    } catch {
        # Log any errors and continue with the next folder
        Write-Host "Error processing $gameName`: $_" -ForegroundColor Red
        "$gameName - Error: $_" | Out-File -FilePath $logErrors -Append
        $errors++
    }
}

$totalTime = (Get-Date) - $startTime
$formattedTotalTime = Format-TimeSpan -TimeSpan $totalTime

Write-Progress -Id 1 -Activity "Scanning Games" -Completed

Write-Host ""
Write-Host "==================== SCAN COMPLETE ====================" -ForegroundColor Green
Write-Host "Done! Batch files saved to: $batchOutputFolder" -ForegroundColor Green
Write-Host "Total time: $formattedTotalTime" -ForegroundColor Cyan
Write-Host "Results: $created created, $sanitized sanitized, $skipped skipped, $notFound not found, $errors errors, $timeouts timeouts" -ForegroundColor White
Write-Host "See FoundGames.log, NotFoundGames.log, SkippedGames.log, ErrorGames.log, and BlockedFiles.log for details" -ForegroundColor Yellow
Write-Host ""

# Pause so user can see results
Read-Host "Press Enter to exit"