# Function to test if path exists and is accessible
function Test-PathAccess {
    param($Path)
    try {
        $null = Get-ChildItem -Path $Path -ErrorAction Stop
        return $true
    }
    catch {
        return $false
    }
}

# Clear any previous errors
$Error.Clear()

Write-Host "=== Game Folder Mover Script ===" -ForegroundColor Cyan
Write-Host ""

# Prompt for source path with better validation
do {
    $sourcePath = Read-Host "Enter the source path (parent folder containing numbered TEKNO folders)"
    
    if ([string]::IsNullOrWhiteSpace($sourcePath)) {
        Write-Host "Path cannot be empty. Please try again." -ForegroundColor Red
        continue
    }
    
    # Remove quotes if user added them
    $sourcePath = $sourcePath.Trim('"')
    
    Write-Host "Testing source path: $sourcePath" -ForegroundColor Yellow
    
    if (-not (Test-Path $sourcePath)) {
        Write-Host "ERROR: Source path does not exist!" -ForegroundColor Red
        Write-Host "Please check the path and try again." -ForegroundColor Yellow
        continue
    }
    
    if (-not (Test-PathAccess $sourcePath)) {
        Write-Host "ERROR: Cannot access source path (permission denied)!" -ForegroundColor Red
        Write-Host "Please check permissions and try again." -ForegroundColor Yellow
        continue
    }
    
    Write-Host "✓ Source path is valid and accessible" -ForegroundColor Green
    break
    
} while ($true)

# Show what's in the source directory
Write-Host "`nScanning source directory..." -ForegroundColor Yellow
try {
    $allFolders = Get-ChildItem -Path $sourcePath -Directory -ErrorAction Stop
    Write-Host "Found $($allFolders.Count) total folders in source directory:" -ForegroundColor Cyan
    foreach ($folder in $allFolders) {
        Write-Host "  - $($folder.Name)" -ForegroundColor Gray
    }
} catch {
    Write-Host "ERROR: Failed to scan source directory: $($_.Exception.Message)" -ForegroundColor Red
    exit
}

# Find TEKNO folders with more flexible pattern matching
Write-Host "`nLooking for TEKNO folders..." -ForegroundColor Yellow
$teknoFolders = $allFolders | Where-Object { $_.Name -match 'TEKNO' }

if ($teknoFolders.Count -eq 0) {
    Write-Host "No TEKNO folders found!" -ForegroundColor Red
    Write-Host "Looking for folders with these patterns:" -ForegroundColor Yellow
    Write-Host "  - *TEKNO* (contains TEKNO anywhere in name)" -ForegroundColor Gray
    Write-Host ""
    Write-Host "Available folders:" -ForegroundColor Yellow
    $allFolders | ForEach-Object { Write-Host "  - $($_.Name)" -ForegroundColor Gray }
    
    $continue = Read-Host "`nWould you like to manually specify which folders to process? (y/n)"
    if ($continue -ne 'y' -and $continue -ne 'Y') {
        exit
    }
    
    # Manual folder selection
    Write-Host "`nEnter folder names to process (one per line, empty line to finish):" -ForegroundColor Yellow
    $selectedFolders = @()
    do {
        $folderName = Read-Host "Folder name"
        if ([string]::IsNullOrWhiteSpace($folderName)) { break }
        
        $folder = $allFolders | Where-Object { $_.Name -eq $folderName }
        if ($folder) {
            $selectedFolders += $folder
            Write-Host "✓ Added: $folderName" -ForegroundColor Green
        } else {
            Write-Host "✗ Folder not found: $folderName" -ForegroundColor Red
        }
    } while ($true)
    
    $teknoFolders = $selectedFolders
}

Write-Host "Found $($teknoFolders.Count) folders to process:" -ForegroundColor Green
foreach ($folder in $teknoFolders) {
    Write-Host "  ✓ $($folder.Name)" -ForegroundColor Green
}

if ($teknoFolders.Count -eq 0) {
    Write-Host "No folders selected. Exiting." -ForegroundColor Yellow
    exit
}

# Prompt for destination path
do {
    $destinationPath = Read-Host "`nEnter the destination path (where games should be moved)"
    
    if ([string]::IsNullOrWhiteSpace($destinationPath)) {
        Write-Host "Path cannot be empty. Please try again." -ForegroundColor Red
        continue
    }
    
    # Remove quotes if user added them
    $destinationPath = $destinationPath.Trim('"')
    
    if (-not (Test-Path $destinationPath)) {
        Write-Host "Destination path does not exist: $destinationPath" -ForegroundColor Yellow
        $create = Read-Host "Would you like to create it? (y/n)"
        if ($create -eq 'y' -or $create -eq 'Y') {
            try {
                New-Item -ItemType Directory -Path $destinationPath -Force -ErrorAction Stop
                Write-Host "✓ Created destination folder: $destinationPath" -ForegroundColor Green
                break
            } catch {
                Write-Host "ERROR: Failed to create destination folder: $($_.Exception.Message)" -ForegroundColor Red
                continue
            }
        } else {
            continue
        }
    } else {
        Write-Host "✓ Destination path exists" -ForegroundColor Green
        break
    }
} while ($true)

# Scan for games in TEKNO folders
Write-Host "`n=== SCANNING FOR GAMES ===" -ForegroundColor Yellow
$moveOperations = @()
$totalGames = 0

foreach ($teknoFolder in $teknoFolders) {
    Write-Host "`nProcessing folder: $($teknoFolder.Name)" -ForegroundColor Magenta
    
    try {
        $games = Get-ChildItem -Path $teknoFolder.FullName -Directory -ErrorAction Stop
        Write-Host "  Found $($games.Count) games" -ForegroundColor Cyan
        
        foreach ($game in $games) {
            $currentPath = $game.FullName
            $newPath = Join-Path $destinationPath $game.Name
            
            Write-Host "    - $($game.Name)" -ForegroundColor Gray
            
            $moveOperations += [PSCustomObject]@{
                Source = $currentPath
                Destination = $newPath
                GameName = $game.Name
                OriginalTeknoFolder = $teknoFolder.Name
            }
            $totalGames++
        }
    } catch {
        Write-Host "  ERROR: Cannot access folder $($teknoFolder.Name): $($_.Exception.Message)" -ForegroundColor Red
    }
}

if ($totalGames -eq 0) {
    Write-Host "`nNo games found to move!" -ForegroundColor Yellow
    exit
}

# Show dry run
Write-Host "`n=== DRY RUN - PREVIEW OF MOVES ===" -ForegroundColor Yellow
Write-Host "Total games to move: $totalGames" -ForegroundColor Cyan
Write-Host ""

foreach ($operation in $moveOperations) {
    Write-Host "GAME: $($operation.GameName)" -ForegroundColor White
    Write-Host "  FROM: $($operation.Source)" -ForegroundColor Gray
    Write-Host "  TO:   $($operation.Destination)" -ForegroundColor Gray
    
    # Check for conflicts
    if (Test-Path $operation.Destination) {
        Write-Host "  ⚠️  WARNING: Destination already exists!" -ForegroundColor Yellow
    }
    Write-Host ""
}

# Confirm execution
Write-Host "=== SUMMARY ===" -ForegroundColor Yellow
Write-Host "Source: $sourcePath" -ForegroundColor Gray
Write-Host "Destination: $destinationPath" -ForegroundColor Gray
Write-Host "Games to move: $totalGames" -ForegroundColor Cyan

Write-Host "`nDo you want to proceed with moving these games? (y/n): " -ForegroundColor Yellow -NoNewline
$proceed = Read-Host

if ($proceed -ne 'y' -and $proceed -ne 'Y') {
    Write-Host "Operation cancelled." -ForegroundColor Yellow
    exit
}

# Execute moves
Write-Host "`n=== EXECUTING MOVES ===" -ForegroundColor Green
$successCount = 0
$errorCount = 0
$skippedCount = 0

foreach ($operation in $moveOperations) {
    Write-Host "Moving: $($operation.GameName)..." -NoNewline
    
    try {
        if (Test-Path $operation.Destination) {
            Write-Host " SKIPPED (already exists)" -ForegroundColor Yellow
            $skippedCount++
            continue
        }
        
        Move-Item -Path $operation.Source -Destination $operation.Destination -Force -ErrorAction Stop
        Write-Host " ✓ SUCCESS" -ForegroundColor Green
        $successCount++
    }
    catch {
        Write-Host " ✗ ERROR: $($_.Exception.Message)" -ForegroundColor Red
        $errorCount++
    }
}

# Final results
Write-Host "`n=== FINAL RESULTS ===" -ForegroundColor Yellow
Write-Host "✓ Successfully moved: $successCount games" -ForegroundColor Green
Write-Host "⚠️  Skipped (already existed): $skippedCount games" -ForegroundColor Yellow
Write-Host "✗ Errors: $errorCount games" -ForegroundColor Red

Write-Host "`nOperation completed!" -ForegroundColor Cyan