# Interactive LaunchBox to CoinOps Image Copier - Enhanced with Platform Matching
param(
    [string]$LaunchBoxRoot = "",
    [string]$CoinOpsRoot = ""
)

function Get-CoinOpsPlatforms {
    param($coinOpsRoot)
    
    $collectionsPath = Join-Path $coinOpsRoot "collections"
    $availablePlatforms = @{}
    
    if (Test-Path $collectionsPath) {
        $platformDirs = Get-ChildItem -Path $collectionsPath -Directory
        foreach ($dir in $platformDirs) {
            $mediumArtworkPath = Join-Path $dir.FullName "medium_artwork\poster"
            if (Test-Path (Split-Path $mediumArtworkPath -Parent)) {
                $availablePlatforms[$dir.Name] = $mediumArtworkPath
            }
        }
    }
    
    # Also check direct subfolders under CoinOps root
    $rootDirs = Get-ChildItem -Path $coinOpsRoot -Directory | Where-Object { $_.Name -ne "collections" }
    foreach ($dir in $rootDirs) {
        $mediumArtworkPath = Join-Path $dir.FullName "medium_artwork\poster"
        if (Test-Path (Split-Path $mediumArtworkPath -Parent)) {
            $availablePlatforms[$dir.Name] = $mediumArtworkPath
        }
    }
    
    return $availablePlatforms
}

function Get-CoinOpsDestination {
    param($platformName, $coinOpsRoot, $availablePlatforms)
    
    # Try exact matches first (case insensitive)
    $exactMatch = $availablePlatforms.Keys | Where-Object { $_.ToLower() -eq $platformName.ToLower() }
    if ($exactMatch) {
        Write-Host "  ✓ Exact match found: $exactMatch" -ForegroundColor Green
        return $availablePlatforms[$exactMatch]
    }
    
    # Try common variations
    $variations = @(
        ($platformName -replace 'Sega ', ''),
        ($platformName -replace 'Nintendo ', ''),
        ($platformName -replace 'Sony ', ''),
        ($platformName -replace 'Microsoft ', ''),
        ($platformName -replace ' ', ''),
        ($platformName -replace ' ', '_'),
        ($platformName -replace ' ', '-')
    )
    
    foreach ($variation in $variations) {
        $match = $availablePlatforms.Keys | Where-Object { $_.ToLower() -eq $variation.ToLower() }
        if ($match) {
            Write-Host "  ✓ Variation match found: $match" -ForegroundColor Green
            return $availablePlatforms[$match]
        }
    }
    
    # Try partial/contains matching
    foreach ($variation in $variations) {
        $match = $availablePlatforms.Keys | Where-Object { 
            $_.ToLower().Contains($variation.ToLower()) -or $variation.ToLower().Contains($_.ToLower())
        }
        if ($match) {
            Write-Host "  ✓ Partial match found: $match" -ForegroundColor Green
            return $availablePlatforms[$match]
        }
    }
    
    # Fallback to original logic if no match found
    Write-Host "  ! No match found, using default path structure" -ForegroundColor Yellow
    return "$coinOpsRoot\collections\$platformName\medium_artwork\poster"
}

function Show-Menu {
    param($platforms, $availableCoinOpsPlatforms)
    
    Clear-Host
    Write-Host "=== LaunchBox Platform Image Copier (Box - Front) ===" -ForegroundColor Green
    Write-Host ""
    Write-Host "CoinOps Status: Found $($availableCoinOpsPlatforms.Count) platforms with artwork folders" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "Available LaunchBox Platforms:" -ForegroundColor Yellow
    Write-Host ""
    
    for ($i = 0; $i -lt $platforms.Count; $i++) {
        $platform = $platforms[$i]
        $boxFrontPath = Join-Path $platform.FullName "Box - Front"
        
        Write-Host "$($i + 1). $($platform.Name)" -ForegroundColor Cyan
        Write-Host "   Path: $boxFrontPath" -ForegroundColor Gray
        
        # Count image files in Box - Front folder
        if (Test-Path $boxFrontPath) {
            $imageCount = (Get-ChildItem -Path $boxFrontPath -Include "*.png", "*.jpg", "*.jpeg" -Recurse -ErrorAction SilentlyContinue).Count
            Write-Host "   Box - Front Images: $imageCount files" -ForegroundColor Green
        } else {
            Write-Host "   Box - Front Images: Folder not found" -ForegroundColor Red
        }
        
        # Show CoinOps matching status
        $destPath = Get-CoinOpsDestination -platformName $platform.Name -coinOpsRoot $script:CoinOpsRoot -availablePlatforms $availableCoinOpsPlatforms
        $destExists = Test-Path (Split-Path $destPath -Parent)
        if ($destExists) {
            Write-Host "   CoinOps Match: ✓ Found" -ForegroundColor Green
        } else {
            Write-Host "   CoinOps Match: ! Will create new" -ForegroundColor Yellow
        }
        
        Write-Host ""
    }
    
    Write-Host "Options:" -ForegroundColor Yellow
    Write-Host "- Enter numbers separated by commas (e.g., 1,3,5)"
    Write-Host "- Enter 'all' to select all platforms"
    Write-Host "- Enter 'list' to see available CoinOps platforms"
    Write-Host "- Enter 'q' to quit"
    Write-Host ""
}

function Show-CoinOpsPlatforms {
    param($availablePlatforms)
    
    Clear-Host
    Write-Host "=== Available CoinOps Platforms ===" -ForegroundColor Green
    Write-Host ""
    
    $sortedPlatforms = $availablePlatforms.Keys | Sort-Object
    foreach ($platform in $sortedPlatforms) {
        Write-Host "• $platform" -ForegroundColor Cyan
        Write-Host "  $($availablePlatforms[$platform])" -ForegroundColor Gray
        Write-Host ""
    }
    
    Write-Host "Press any key to return to menu..."
    Read-Host
}

function Get-LaunchBoxPlatforms {
    param($launchBoxPath)
    
    $imagesPath = Join-Path $launchBoxPath "Images"
    
    if (!(Test-Path $imagesPath)) {
        Write-Error "LaunchBox Images path not found: $imagesPath"
        return $null
    }
    
    # Get all directories that contain a "Box - Front" subfolder
    $platforms = @()
    $allDirs = Get-ChildItem -Path $imagesPath -Directory | Sort-Object Name
    
    foreach ($dir in $allDirs) {
        $boxFrontPath = Join-Path $dir.FullName "Box - Front"
        if (Test-Path $boxFrontPath) {
            $platforms += $dir
        }
    }
    
    return $platforms
}

function Copy-PlatformImages {
    param($sourcePath, $destPath, $platformName)
    
    # The source path already points to the platform folder, just add Box - Front
    $boxFrontPath = Join-Path $sourcePath "Box - Front"
    
    Write-Host "Processing: $platformName" -ForegroundColor Green
    Write-Host "Source: $boxFrontPath" -ForegroundColor Gray
    Write-Host "Destination: $destPath" -ForegroundColor Gray
    
    # Check if Box - Front folder exists
    if (!(Test-Path $boxFrontPath)) {
        Write-Host "Box - Front folder not found for this platform" -ForegroundColor Red
        Write-Host ("=" * 50)
        Write-Host ""
        return
    }
    
    # Create destination directory if it doesn't exist
    if (!(Test-Path $destPath)) {
        New-Item -ItemType Directory -Path $destPath -Force | Out-Null
        Write-Host "Created destination directory" -ForegroundColor Yellow
    }
    
    # Get all image files from Box - Front folder
    $imageFiles = Get-ChildItem -Path $boxFrontPath -Include "*.png", "*.jpg", "*.jpeg" -Recurse -ErrorAction SilentlyContinue
    
    if ($imageFiles.Count -eq 0) {
        Write-Host "No image files found in Box - Front folder" -ForegroundColor Red
        Write-Host ("=" * 50)
        Write-Host ""
        return
    }
    
    Write-Host "Copying $($imageFiles.Count) box art images..." -ForegroundColor Cyan
    
    $copiedCount = 0
    $errorCount = 0
    
    foreach ($file in $imageFiles) {
        $destinationFile = Join-Path $destPath $file.Name
        
        try {
            Copy-Item $file.FullName $destinationFile -Force
            $copiedCount++
            Write-Host "  ✓ $($file.Name)" -ForegroundColor Green
        } catch {
            $errorCount++
            Write-Host "  ✗ Failed: $($file.Name) - $($_.Exception.Message)" -ForegroundColor Red
        }
    }
    
    Write-Host "Completed: $copiedCount copied, $errorCount errors" -ForegroundColor Yellow
    Write-Host ("=" * 50)
    Write-Host ""
}

# Main script execution
Clear-Host

# Get LaunchBox root path
if ($LaunchBoxRoot -eq "") {
    Write-Host "Enter your LaunchBox root directory path:" -ForegroundColor Yellow
    Write-Host "(e.g., C:\LaunchBox)" -ForegroundColor Gray
    $LaunchBoxRoot = Read-Host "LaunchBox Path"
}

if (!(Test-Path $LaunchBoxRoot)) {
    Write-Error "LaunchBox path does not exist: $LaunchBoxRoot"
    exit 1
}

# Get CoinOps root path
if ($CoinOpsRoot -eq "") {
    Write-Host "Enter your CoinOps root directory path:" -ForegroundColor Yellow
    Write-Host "(e.g., C:\CoinOps)" -ForegroundColor Gray
    $CoinOpsRoot = Read-Host "CoinOps Path"
}

if (!(Test-Path $CoinOpsRoot)) {
    Write-Error "CoinOps path does not exist: $CoinOpsRoot"
    exit 1
}

# Scan CoinOps platforms
Write-Host "Scanning CoinOps platforms..." -ForegroundColor Cyan
$availableCoinOpsPlatforms = Get-CoinOpsPlatforms -coinOpsRoot $CoinOpsRoot
Write-Host "Found $($availableCoinOpsPlatforms.Count) CoinOps platforms with artwork folders" -ForegroundColor Green

# Get available LaunchBox platforms
Write-Host "Scanning LaunchBox platforms with Box - Front folders..." -ForegroundColor Cyan
$platforms = Get-LaunchBoxPlatforms -launchBoxPath $LaunchBoxRoot

if (!$platforms -or $platforms.Count -eq 0) {
    Write-Error "No platforms with Box - Front folders found in LaunchBox"
    exit 1
}

# Main menu loop
do {
    Show-Menu -platforms $platforms -availableCoinOpsPlatforms $availableCoinOpsPlatforms
    $selection = Read-Host "Select platforms"
    
    if ($selection -eq 'q') {
        Write-Host "Goodbye!" -ForegroundColor Green
        exit 0
    }
    
    if ($selection -eq 'list') {
        Show-CoinOpsPlatforms -availablePlatforms $availableCoinOpsPlatforms
        continue
    }
    
    $selectedPlatforms = @()
    
    if ($selection -eq 'all') {
        $selectedPlatforms = $platforms
    } else {
        $numbers = $selection.Split(',') | ForEach-Object { $_.Trim() }
        foreach ($num in $numbers) {
            if ($num -match '^\d+$' -and [int]$num -ge 1 -and [int]$num -le $platforms.Count) {
                $selectedPlatforms += $platforms[[int]$num - 1]
            }
        }
    }
    
    if ($selectedPlatforms.Count -eq 0) {
        Write-Host "Invalid selection. Press any key to continue..." -ForegroundColor Red
        Read-Host
        continue
    }
    
    # Show selected platforms and destinations
    Clear-Host
    Write-Host "=== Copy Preview (Box - Front Images) ===" -ForegroundColor Green
    Write-Host ""
    
    foreach ($platform in $selectedPlatforms) {
        $destPath = Get-CoinOpsDestination -platformName $platform.Name -coinOpsRoot $CoinOpsRoot -availablePlatforms $availableCoinOpsPlatforms
        $boxFrontPath = Join-Path $platform.FullName "Box - Front"
        
        Write-Host "Platform: $($platform.Name)" -ForegroundColor Cyan
        Write-Host "  From: $boxFrontPath" -ForegroundColor Gray
        Write-Host "  To:   $destPath" -ForegroundColor Gray
        
        # Show file count
        if (Test-Path $boxFrontPath) {
            $imageCount = (Get-ChildItem -Path $boxFrontPath -Include "*.png", "*.jpg", "*.jpeg" -Recurse -ErrorAction SilentlyContinue).Count
            Write-Host "  Files: $imageCount box art images" -ForegroundColor Green
        } else {
            Write-Host "  Files: Box - Front folder not found" -ForegroundColor Red
        }
        Write-Host ""
    }
    
    $confirm = Read-Host "Proceed with copy? (y/n)"
    
    if ($confirm -eq 'y' -or $confirm -eq 'yes') {
        Clear-Host
        Write-Host "=== Starting Copy Process (Box - Front Images) ===" -ForegroundColor Green
        Write-Host ""
        
        foreach ($platform in $selectedPlatforms) {
            $destPath = Get-CoinOpsDestination -platformName $platform.Name -coinOpsRoot $CoinOpsRoot -availablePlatforms $availableCoinOpsPlatforms
            Copy-PlatformImages -sourcePath $platform.FullName -destPath $destPath -platformName $platform.Name
        }
        
        Write-Host "All copy operations completed!" -ForegroundColor Green
        Write-Host "Press any key to return to menu..."
        Read-Host
    }
    
} while ($true)