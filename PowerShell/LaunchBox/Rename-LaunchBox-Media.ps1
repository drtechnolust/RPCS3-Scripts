# LaunchBox Media Renamer - Images & Videos
# Removes LaunchBox's "-01", "-02" etc. suffixes from images and videos

param(
    [switch]$DryRun
)

function Rename-LaunchBoxMedia {
    param(
        [string]$SourcePath,
        [bool]$IsDryRun
    )
    
    Write-Host "Scanning for LaunchBox '-01' suffixed media files 3 levels deep..." -ForegroundColor Yellow
    Write-Host "Source: $SourcePath" -ForegroundColor Cyan
    
    if ($IsDryRun) {
        Write-Host "=== DRY RUN MODE - NO FILES WILL BE CHANGED ===" -ForegroundColor Red
    }
    
    # Include both image AND video formats
    $mediaFiles = Get-ChildItem -Path $SourcePath -Recurse -Depth 3 -Include "*.png", "*.jpg", "*.jpeg", "*.gif", "*.bmp", "*.mp4", "*.avi", "*.wmv", "*.mov", "*.mkv" | 
                  Where-Object { $_.BaseName -match "-\d{2}$" }  # Only dash-number at END of filename
    
    if ($mediaFiles.Count -eq 0) {
        Write-Host "No LaunchBox '-01' style media files found." -ForegroundColor Green
        return
    }
    
    # Group by file type for better reporting
    $imageFiles = $mediaFiles | Where-Object { $_.Extension -match '\.(png|jpg|jpeg|gif|bmp)$' }
    $videoFiles = $mediaFiles | Where-Object { $_.Extension -match '\.(mp4|avi|wmv|mov|mkv)$' }
    
    Write-Host "Found $($mediaFiles.Count) LaunchBox numbered media files:" -ForegroundColor Green
    if ($imageFiles.Count -gt 0) {
        Write-Host "  $($imageFiles.Count) image files" -ForegroundColor Cyan
    }
    if ($videoFiles.Count -gt 0) {
        Write-Host "  $($videoFiles.Count) video files" -ForegroundColor Magenta
    }
    Write-Host ""
    
    $renamedCount = 0
    
    foreach ($file in $mediaFiles) {
        # Remove only the -## pattern at the end
        $newBaseName = $file.BaseName -replace "-\d{2}$", ""
        $newName = $newBaseName + $file.Extension
        $newPath = Join-Path $file.Directory $newName
        
        # Check if target file already exists
        if (Test-Path $newPath) {
            Write-Host "SKIP: $($file.Name) → Target exists: $newName" -ForegroundColor Yellow
            continue
        }
        
        if ($IsDryRun) {
            $fileType = if ($file.Extension -match '\.(mp4|avi|wmv|mov|mkv)$') { "VIDEO" } else { "IMAGE" }
            Write-Host "WOULD RENAME ($fileType): $($file.Name) → $newName" -ForegroundColor Cyan
        } else {
            try {
                Rename-Item -Path $file.FullName -NewName $newName -ErrorAction Stop
                $fileType = if ($file.Extension -match '\.(mp4|avi|wmv|mov|mkv)$') { "VIDEO" } else { "IMAGE" }
                Write-Host "RENAMED ($fileType): $($file.Name) → $newName" -ForegroundColor Green
                $renamedCount++
            }
            catch {
                Write-Host "ERROR: Failed to rename $($file.Name) - $($_.Exception.Message)" -ForegroundColor Red
            }
        }
    }
    
    Write-Host ""
    if ($IsDryRun) {
        Write-Host "Dry run complete. $($mediaFiles.Count) media files would be processed." -ForegroundColor Yellow
    } else {
        Write-Host "Renaming complete! $renamedCount media files renamed successfully." -ForegroundColor Green
    }
}

# Main script execution
Clear-Host
Write-Host "LaunchBox Media Renamer - Images & Videos" -ForegroundColor Magenta
Write-Host "Removes LaunchBox's '-01', '-02' etc. suffixes" -ForegroundColor White
Write-Host "Supports: Images (PNG, JPG, GIF, BMP) & Videos (MP4, AVI, WMV, MOV, MKV)" -ForegroundColor Gray
Write-Host "=" * 70

# Prompt for source directory
do {
    $sourcePath = Read-Host "Enter the path to your LaunchBox media folder"
    $sourcePath = $sourcePath.Trim('"')
    
    if (-not (Test-Path $sourcePath)) {
        Write-Host "Path does not exist. Please try again." -ForegroundColor Red
    }
} while (-not (Test-Path $sourcePath))

# Ask for dry run if not specified as parameter
if (-not $PSBoundParameters.ContainsKey('DryRun')) {
    $dryRunChoice = Read-Host "Run in dry-run mode first? (y/n)"
    $DryRun = $dryRunChoice -eq 'y' -or $dryRunChoice -eq 'yes'
}

# Execute the renaming
Rename-LaunchBoxMedia -SourcePath $sourcePath -IsDryRun $DryRun

# If it was a dry run, ask if they want to run for real
if ($DryRun) {
    Write-Host ""
    $realRun = Read-Host "Run the actual renaming now? (y/n)"
    if ($realRun -eq 'y' -or $realRun -eq 'yes') {
        Write-Host ""
        Rename-LaunchBoxMedia -SourcePath $sourcePath -IsDryRun $false
    }
}

Write-Host ""
Write-Host "Script completed. Press any key to exit..." -ForegroundColor White

# Fixed ReadKey - no more errors!
try {
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
} catch {
    $null = Read-Host "Press Enter to exit"
}