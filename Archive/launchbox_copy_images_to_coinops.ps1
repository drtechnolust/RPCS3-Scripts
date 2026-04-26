# Interactive LaunchBox to CoinOps Image Copier - Fixed CoinOps Path Structure
param(
    [string]$LaunchBoxRoot = "",
    [string]$CoinOpsRoot = ""
)

function Show-Menu {
    param($platforms)
    
    Clear-Host
    Write-Host "=== LaunchBox Platform Image Copier (Box - Front) ===" -ForegroundColor Green
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
        Write-Host ""
    }
    
    Write-Host "Options:" -ForegroundColor Yellow
    Write-Host "- Enter numbers separated by commas (e.g., 1,3,5)"
    Write-Host "- Enter 'all' to select all platforms"
    Write-Host "- Enter 'q' to quit"
    Write-Host ""
}

function Get-LaunchBoxPlatforms {
    param($launchBoxPath)
    
    # Changed from Images\Platforms to just Images
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

function Get-CoinOpsDestination {
    param($platformName, $coinOpsRoot)
    
    # Common CoinOps folder structure patterns - FIXED to use medium_artwork
    $possiblePaths = @(
        "$coinOpsRoot\collections\$platformName\medium_artwork\poster",
        "$coinOpsRoot\collections\$($platformName.ToLower())\medium_artwork\poster",
        "$coinOpsRoot\$platformName\medium_artwork\poster",
        "$coinOpsRoot\$($platformName.ToLower())\medium_artwork\poster",
        "$coinOpsRoot\collections\$($platformName -replace ' ', '')\medium_artwork\poster",
        "$coinOpsRoot\collections\$($platformName -replace ' ', '_')\medium_artwork\poster",
        "$coinOpsRoot\collections\$($platformName -replace ' ', '-')\medium_artwork\poster"
    )
    
    # Check if any collections folder exists, otherwise use the first pattern
    foreach ($path in $possiblePaths) {
        $collectionsParent = Split-Path (Split-Path $path -Parent) -Parent
        if (Test-Path $collectionsParent) {
            return $path
        }
    }
    
    return $possiblePaths[0]  # Default to first pattern
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

# Get available platforms
Write-Host "Scanning LaunchBox platforms with Box - Front folders..." -ForegroundColor Cyan
$platforms = Get-LaunchBoxPlatforms -launchBoxPath $LaunchBoxRoot

if (!$platforms -or $platforms.Count -eq 0) {
    Write-Error "No platforms with Box - Front folders found in LaunchBox"
    exit 1
}

# Main menu loop
do {
    Show-Menu -platforms $platforms
    $selection = Read-Host "Select platforms"
    
    if ($selection -eq 'q') {
        Write-Host "Goodbye!" -ForegroundColor Green
        exit 0
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
        $destPath = Get-CoinOpsDestination -platformName $platform.Name -coinOpsRoot $CoinOpsRoot
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
            $destPath = Get-CoinOpsDestination -platformName $platform.Name -coinOpsRoot $CoinOpsRoot
            Copy-PlatformImages -sourcePath $platform.FullName -destPath $destPath -platformName $platform.Name
        }
        
        Write-Host "All copy operations completed!" -ForegroundColor Green
        Write-Host "Press any key to return to menu..."
        Read-Host
    }
    
} while ($true)