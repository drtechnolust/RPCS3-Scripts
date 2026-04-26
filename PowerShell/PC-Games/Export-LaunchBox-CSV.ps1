# LaunchBox Game Scanner Script
# This script scans a folder structure with games, identifies executable files,
# and generates a CSV file that can be imported into LaunchBox

# Configuration
$rootGameFolder = "D:\Arcade\System roms\PC Games 2"
$outputCsvFile = "$env:USERPROFILE\Desktop\LaunchBoxGames.csv"
$maxDepth = 3  # Maximum subfolder depth to search for executables

# Common game executables to prioritize
$priorityExecutables = @(
    "game.exe", "start.exe", "launcher.exe", "run.exe", 
    "main.exe", "play.exe", "setup.exe"
)

# Function to determine if an executable is likely to be the main game executable
function Is-LikelyMainExecutable {
    param($exePath, $gameName)
    
    $fileName = [System.IO.Path]::GetFileNameWithoutExtension($exePath).ToLower()
    $folderName = [System.IO.Path]::GetFileName([System.IO.Path]::GetDirectoryName($exePath)).ToLower()
    $gameNameLower = $gameName.ToLower()
    
    # Check if the executable name matches the game name
    if ($fileName -eq $gameNameLower) {
        return $true
    }
    
    # Check if it's a common game executable
    foreach ($priorityExe in $priorityExecutables) {
        if ($fileName -eq [System.IO.Path]::GetFileNameWithoutExtension($priorityExe)) {
            return $true
        }
    }
    
    # Check if the executable contains the game name
    if ($fileName -match $gameNameLower) {
        return $true
    }
    
    return $false
}

# Create CSV header
"Title,ApplicationPath,Platform" | Out-File -FilePath $outputCsvFile -Encoding utf8

# Get all game folders in the root directory
$gameFolders = Get-ChildItem -Path $rootGameFolder -Directory

$totalGames = $gameFolders.Count
$currentGame = 0

Write-Host "Found $totalGames game folders to process"

foreach ($gameFolder in $gameFolders) {
    $currentGame++
    $progress = [math]::Round(($currentGame / $totalGames) * 100)
    Write-Progress -Activity "Scanning Games" -Status "$($gameFolder.Name) ($currentGame of $totalGames)" -PercentComplete $progress
    
    $gameName = $gameFolder.Name
    $exeFiles = @()
    
    # Search for executable files with limited depth
    $depth = 0
    $currentFolder = $gameFolder.FullName
    
    while ($depth -lt $maxDepth) {
        $exeFiles += Get-ChildItem -Path $currentFolder -Filter "*.exe" -File -ErrorAction SilentlyContinue
        
        # If we found executables, no need to go deeper
        if ($exeFiles.Count -gt 0 -and $depth -gt 0) {
            break
        }
        
        # Look for subdirectories that might contain the game
        $subDirs = Get-ChildItem -Path $currentFolder -Directory -ErrorAction SilentlyContinue
        
        # If no subdirectories or multiple subdirectories, stop searching deeper
        if ($subDirs.Count -ne 1) {
            break
        }
        
        # Move to the single subdirectory
        $currentFolder = $subDirs[0].FullName
        $depth++
    }
    
    # If no executables found, log and continue
    if ($exeFiles.Count -eq 0) {
        Write-Host "No executables found for $gameName" -ForegroundColor Yellow
        continue
    }
    
    # Try to find the main executable
    $mainExe = $null
    
    # First, check for executables that match priority criteria
    foreach ($exe in $exeFiles) {
        if (Is-LikelyMainExecutable -exePath $exe.FullName -gameName $gameName) {
            $mainExe = $exe
            break
        }
    }
    
    # If no priority match, take the first executable
    if ($null -eq $mainExe -and $exeFiles.Count -gt 0) {
        $mainExe = $exeFiles[0]
    }
    
    # Write to CSV
    if ($null -ne $mainExe) {
        "$gameName,$($mainExe.FullName),PC" | Out-File -FilePath $outputCsvFile -Encoding utf8 -Append
    }
}

Write-Host "Processing complete! CSV file saved to: $outputCsvFile" -ForegroundColor Green
Write-Host "You can now import this file into LaunchBox using the 'Import → ROM Files or Folders → From Comma Separated Values (CSV) File' option."