# Enhanced Archive Extraction Script
param(
    [int]$MaxParallel = 6,
    [int]$MaxRetries = 3,
    [int]$RetryDelaySeconds = 5
)

# Configuration
$sourcePath = "\\10.10.1.99\Arcade\System roms\TeknoParrot JollyGim"
$sevenZipPath = "C:\Program Files\7-Zip\7z.exe"
$logFile = Join-Path $sourcePath "extraction_log.txt"
$errorLogFile = Join-Path $sourcePath "extraction_errors.txt"
$checkpointFile = Join-Path $sourcePath "extraction_checkpoint.json"
$mutexName = "Global\ArchiveExtractionCheckpointLock"

function Test-Prerequisites {
    Write-Host "Validating prerequisites..." -ForegroundColor Yellow
    
    if (-not (Test-Path $sevenZipPath)) {
        throw "7-Zip not found at: $sevenZipPath"
    }
    
    try {
        $testPath = Split-Path $sourcePath -Parent
        if (-not (Test-Path $testPath -ErrorAction Stop)) {
            throw "Cannot access network path: $sourcePath"
        }
    } catch {
        throw "Network connectivity issue: $($_.Exception.Message)"
    }
    
    if (-not (Test-Path $sourcePath)) {
        throw "Source path not accessible: $sourcePath"
    }
    
    Write-Host "✓ All prerequisites validated successfully" -ForegroundColor Green
}

function Test-PathLength {
    param([string]$Path)
    return $Path.Length -lt 250
}

function Show-FolderCreation {
    param([string]$FolderPath, [string]$ArchiveName)
    $folderName = Split-Path $FolderPath -Leaf
    Write-Host "📁 Created folder: $folderName (for $ArchiveName)" -ForegroundColor DarkGreen
}

function Show-ProgressBar {
    param(
        [int]$Current,
        [int]$Total,
        [int]$Completed,
        [int]$Failed,
        [int]$Active,
        [string]$CurrentFile = ""
    )
    
    if ($Total -eq 0) { return }
    
    $percent = [math]::Round(($Current / $Total) * 100, 1)
    $progressWidth = 40
    $filledWidth = [math]::Round(($percent / 100) * $progressWidth)
    $emptyWidth = $progressWidth - $filledWidth
    
    $filled = "█" * $filledWidth
    $empty = "░" * $emptyWidth
    $progressBar = "[$filled$empty]"
    
    Write-Host "`r$progressBar $percent% ($Current/$Total) ✅$Completed ❌$Failed 🔄$Active" -NoNewline
    
    if ($CurrentFile) {
        $displayFile = if ($CurrentFile.Length -gt 30) { 
            "..." + $CurrentFile.Substring($CurrentFile.Length - 27)
        } else { 
            $CurrentFile 
        }
        Write-Host " | $displayFile" -NoNewline
    }
}

function Initialize-Checkpoint {
    if (Test-Path $checkpointFile) {
        try {
            $content = Get-Content $checkpointFile -Raw -ErrorAction Stop
            if ($content) {
                $checkpoint = $content | ConvertFrom-Json -AsHashtable
                Write-Host "✓ Loaded checkpoint with $($checkpoint.Count) entries" -ForegroundColor Green
                return $checkpoint
            }
        } catch {
            Write-Warning "Checkpoint file corrupted, creating new one"
        }
    }
    Write-Host "✓ Creating new checkpoint file" -ForegroundColor Green
    return @{}
}

# Job script block
$scriptBlock = {
    param (
        $archivePath, $outputFolder, $sevenZipPath,
        $logFile, $errorLogFile, $checkpointFile, $mutexName,
        $maxRetries, $retryDelay
    )
    
    $fileKey = $archivePath.ToLowerInvariant()
    $success = $false
    $lastError = ""
    $fileName = Split-Path $archivePath -Leaf
    $folderName = Split-Path $outputFolder -Leaf
    
    for ($attempt = 1; $attempt -le $maxRetries; $attempt++) {
        try {
            Write-Output "🔄 [$attempt/$maxRetries] Extracting: $fileName"
            
            $result = & "$sevenZipPath" x "`"$archivePath`"" -o"`"$outputFolder`"" -y 2>&1
            
            if ($LASTEXITCODE -eq 0) {
                $success = $true
                Write-Output "✅ SUCCESS: $fileName → $folderName"
                Add-Content -Path $logFile -Value "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - SUCCESS: $archivePath"
                break
            } else {
                $lastError = "7-Zip exit code: $LASTEXITCODE"
                Write-Output "❌ FAILED [$attempt/$maxRetries]: $fileName - $lastError"
            }
        } catch {
            $lastError = $_.Exception.Message
            Write-Output "❌ EXCEPTION [$attempt/$maxRetries]: $fileName - $lastError"
        }
        
        if ($attempt -lt $maxRetries) {
            $delay = $retryDelay * [Math]::Pow(2, $attempt - 1)
            Write-Output "⏳ Waiting $delay seconds..."
            Start-Sleep -Seconds $delay
        }
    }
    
    # Update checkpoint
    try {
        $mutex = [System.Threading.Mutex]::OpenExisting($mutexName)
        $acquired = $mutex.WaitOne(30000)
        
        if ($acquired) {
            try {
                $checkpoint = @{}
                if (Test-Path $checkpointFile) {
                    $content = Get-Content $checkpointFile -Raw -ErrorAction SilentlyContinue
                    if ($content) {
                        $checkpoint = $content | ConvertFrom-Json -AsHashtable
                    }
                }
                
                if ($success) {
                    $checkpoint[$fileKey] = @{
                        status = "completed"
                        timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
                        attempts = $attempt
                    }
                } else {
                    $checkpoint[$fileKey] = @{
                        status = "failed"
                        timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
                        attempts = $maxRetries
                        error = $lastError
                    }
                    Add-Content -Path $errorLogFile -Value "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - FAILED: $archivePath - $lastError"
                }
                
                $checkpoint | ConvertTo-Json -Depth 10 | Set-Content -Path $checkpointFile -Encoding UTF8
            } finally {
                $mutex.ReleaseMutex()
            }
        }
    } catch {
        Write-Output "Checkpoint update error: $($_.Exception.Message)"
    }
    
    return @{
        Success = $success
        Error = $lastError
        Attempts = $attempt
        FileName = $fileName
        FolderName = $folderName
    }
}

# Main execution
try {
    Write-Host "=== Enhanced Archive Extraction Script ===" -ForegroundColor Cyan
    Write-Host "Source: $sourcePath"
    Write-Host "Max Parallel Jobs: $MaxParallel"
    Write-Host "Max Retries: $MaxRetries"
    Write-Host ""
    
    Test-Prerequisites
    
    $checkpoint = Initialize-Checkpoint
    
    $mutex = New-Object System.Threading.Mutex($false, $mutexName)
    Write-Host "✓ Created global mutex" -ForegroundColor Green
    
    Write-Host "Scanning for archives..." -ForegroundColor Yellow
    $allArchives = Get-ChildItem -Path $sourcePath -File -Include *.zip, *.7z, *.rar -Recurse -ErrorAction SilentlyContinue
    
    $archives = $allArchives | Where-Object {
        $key = $_.FullName.ToLowerInvariant()
        $alreadyProcessed = $checkpoint.ContainsKey($key) -and $checkpoint[$key].status -eq "completed"
        $pathTooLong = -not (Test-PathLength $_.FullName)
        
        if ($pathTooLong) {
            Write-Warning "Skipping file with path too long: $($_.Name)"
            return $false
        }
        
        return -not $alreadyProcessed
    }
    
    $total = $archives.Count
    $skipped = $allArchives.Count - $total
    
    Write-Host "✓ Found $($allArchives.Count) total archives" -ForegroundColor Green
    Write-Host "✓ Skipping $skipped already processed archives" -ForegroundColor Green
    Write-Host "✓ Processing $total new archives" -ForegroundColor Green
    
    if ($total -eq 0) {
        Write-Host "No new archives to process!" -ForegroundColor Yellow
        return
    }
    
    "=== Extraction Log Started: $(Get-Date) ===" | Out-File $logFile -Encoding UTF8
    "=== Error Log Started: $(Get-Date) ===" | Out-File $errorLogFile -Encoding UTF8
    
    $counter = 0
    $jobs = @()
    $completed = 0
    $failed = 0
    
    foreach ($archive in $archives) {
        $counter++
        $baseName = [System.IO.Path]::GetFileNameWithoutExtension($archive.Name)
        $outputFolder = Join-Path $sourcePath $baseName
        
        if (-not (Test-Path $outputFolder)) {
            try {
                New-Item -ItemType Directory -Path $outputFolder -Force | Out-Null
                Show-FolderCreation -FolderPath $outputFolder -ArchiveName $archive.Name
            } catch {
                Write-Warning "Failed to create output folder: $outputFolder"
                continue
            }
        }
        
        # Wait for available slot
        while ($jobs.Count -ge $MaxParallel) {
            $completedJobs = $jobs | Where-Object { $_.State -eq 'Completed' }
            foreach ($job in $completedJobs) {
                $result = Receive-Job $job
                if ($result.Success) {
                    $completed++
                    Write-Host "`n✅ COMPLETED: $($result.FileName)" -ForegroundColor Green
                } else {
                    $failed++
                    Write-Host "`n❌ FAILED: $($result.FileName)" -ForegroundColor Red
                }
                Remove-Job $job
            }
            $jobs = $jobs | Where-Object { $_.State -eq 'Running' }
            
            if ($jobs.Count -ge $MaxParallel) {
                Show-ProgressBar -Current $counter -Total $total -Completed $completed -Failed $failed -Active $jobs.Count -CurrentFile $archive.Name
                Start-Sleep -Seconds 2
            }
        }
        
        # Launch new job
        $job = Start-Job -ScriptBlock $scriptBlock -ArgumentList @(
            $archive.FullName, $outputFolder, $sevenZipPath,
            $logFile, $errorLogFile, $checkpointFile, $mutexName,
            $MaxRetries, $RetryDelaySeconds
        )
        $jobs += $job
    }
    
    # Wait for remaining jobs
    Write-Host "`n🔄 Waiting for all jobs to complete..." -ForegroundColor Yellow
    
    while ($jobs.Count -gt 0) {
        $runningJobs = $jobs | Where-Object { $_.State -eq 'Running' }
        $completedJobs = $jobs | Where-Object { $_.State -eq 'Completed' }
        
        foreach ($job in $completedJobs) {
            $result = Receive-Job $job
            if ($result.Success) {
                $completed++
                Write-Host "`n✅ COMPLETED: $($result.FileName)" -ForegroundColor Green
            } else {
                $failed++
                Write-Host "`n❌ FAILED: $($result.FileName)" -ForegroundColor Red
            }
            Remove-Job $job
        }
        
        $jobs = $runningJobs
        
        if ($jobs.Count -gt 0) {
            Show-ProgressBar -Current ($completed + $failed) -Total $total -Completed $completed -Failed $failed -Active $jobs.Count
            Start-Sleep -Seconds 3
        }
    }
    
    Write-Host ""
    Show-ProgressBar -Current $total -Total $total -Completed $completed -Failed $failed -Active 0
    Write-Host "`n"
    
    Write-Host "`n=== EXTRACTION COMPLETE ===" -ForegroundColor Green
    Write-Host "✓ Successfully extracted: $completed archives" -ForegroundColor Green
    if ($failed -gt 0) {
        Write-Host "✗ Failed to extract: $failed archives" -ForegroundColor Red
        Write-Host "Check error log: $errorLogFile"
    }
    Write-Host "Full log: $logFile"

} catch {
    Write-Error "Script execution failed: $($_.Exception.Message)"
    exit 1
} finally {
    if ($mutex) {
        $mutex.Dispose()
    }
    Get-Job | Remove-Job -Force -ErrorAction SilentlyContinue
}