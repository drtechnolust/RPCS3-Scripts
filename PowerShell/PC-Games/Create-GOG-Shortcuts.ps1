<#
.SYNOPSIS
    GOG Games Shortcut Generator
.DESCRIPTION
    Automatically creates shortcuts for GOG games with folder-name-based naming.
    Optimized for GOG's DRM-free structure and naming conventions.
    Shortcut names match folder names (e.g., Kona Day One.lnk from "Kona_Day_One" folder).
.AUTHOR
    [Your Name]
.VERSION
    1.0
.DATE
    June 5, 2025
.COPYRIGHT
    Copyright (c) 2025 [Your Name]. All rights reserved.
.FEATURES
    - GOG-optimized executable detection
    - Sentence case shortcut naming (Balls of Steel.lnk)
    - Prioritizes [GameName].exe pattern
    - Handles GOG launchers (goglauncherfile.exe)
    - DOSBox and ScummVM game support
    - Enhanced root-level and bin folder searching
#>

# GOG Games Shortcut Generator
$rootGameFolder = "D:\Arcade\System roms\GOG Games"
$shortcutOutputFolder = "$rootGameFolder\_GOG_Shortcuts"
$maxDepth = 6  # GOG games are typically simpler structure
$folderTimeout = 180  # Shorter timeout for simpler GOG structure
$logFound = "$shortcutOutputFolder\FoundGames.log"
$logNotFound = "$shortcutOutputFolder\NotFoundGames.log"
$logSkipped = "$shortcutOutputFolder\SkippedGames.log"
$logErrors = "$shortcutOutputFolder\ErrorGames.log"
$logBlocked = "$shortcutOutputFolder\BlockedFiles.log"

# Ensure shortcut output folder exists
if (!(Test-Path $shortcutOutputFolder)) {
    New-Item -ItemType Directory -Path $shortcutOutputFolder | Out-Null
}

# Clear logs if they exist
Remove-Item -Path $logFound, $logNotFound, $logSkipped, $logErrors, $logBlocked -ErrorAction SilentlyContinue

# Store a hashtable of created shortcuts to check for duplicates
$createdShortcuts = @{}

function Get-GOGExecutableScore {
    param($exePath, $gameFolderName, $folderPath)
    $exeName = [System.IO.Path]::GetFileNameWithoutExtension($exePath).ToLower()
    $folderName = $gameFolderName.ToLower()
    $score = 0
    
    # Block obvious utility files (lighter filtering for GOG)
    $blockPatterns = @(
        "*uninstall*", "*setup*", "*config*", "*crash*", 
        "*update*", "*redist*", "*directx*"
    )
    foreach ($pattern in $blockPatterns) {
        if ($exeName -like $pattern) {
            "$gameFolderName - BLOCKED by pattern '$pattern': $exePath" | Out-File -FilePath $logBlocked -Append
            return -1
        }
    }
    
    # GOG-specific blacklist
    $gogBadNames = @(
        "unins000", "dosbox", "scummvm", "galaxy", "gogdosconfig",
        "vcredist", "dxsetup", "oalinst", "physx", "dotnetfx", "gameuninstallhelper"
    )
    
    if ($gogBadNames -contains $exeName) {
        "$gameFolderName - BLOCKED by blacklist: $exePath" | Out-File -FilePath $logBlocked -Append
        return -1
    }
    
    # HIGHEST PRIORITY: Exact folder name match (most common GOG pattern)
    if ($exeName -eq $folderName) { 
        return 1000  # Perfect match - almost always correct for GOG
    }
    
    # VERY HIGH PRIORITY: Single word from folder name (e.g., "kona.exe" from "Kona_Day_One")
    $folderParts = $folderName -split '[\s_-]+'  # Split on spaces, underscores, hyphens
    foreach ($part in $folderParts) {
        if ($part.Length -gt 2 -and $exeName -eq $part.ToLower()) {
            return 950  # Very high score for single-word folder matches
        }
    }
    
    # Remove spaces, underscores, and hyphens for compound matching
    $cleanFolderName = $folderName -replace '[^a-zA-Z0-9]', ''
    if ($exeName -eq $cleanFolderName) {
        return 900  # High score for compound name matches
    }
    
    # Check if exe name is contained in folder name parts
    foreach ($part in $folderParts) {
        if ($part.Length -gt 2 -and $part.ToLower() -eq $exeName) {
            return 850  # Good match for word parts
        }
    }
    
    # HIGH PRIORITY: GOG launcher files (modern GOG pattern)
    if ($exeName -eq "goglauncherfile" -or $exeName -like "*gog*launcher*") {
        return 800  # GOG's preferred launcher
    }
    
    # GOG launcher pattern with game name
    if (($exeName -like "*$folderName*launcher*") -or ($exeName -like "*launcher*" -and $exeName -like "*$folderName*")) {
        return 750  # Game-specific GOG launcher
    }
    
    # MEDIUM-HIGH: Generic "game.exe" (very common in GOG)
    if ($exeName -eq "game") {
        return 600  # Very common in GOG but less specific
    }
    
    # DOSBox wrapped games (exclude plain dosbox.exe)
    if ($exeName -like "*dosbox*" -and $exeName -ne "dosbox" -and $exeName -like "*$folderName*") {
        return 550  # DOSBox-wrapped game
    }
    
    # ScummVM wrapped games
    if ($exeName -like "*scummvm*" -and $exeName -ne "scummvm" -and $exeName -like "*$folderName*") {
        return 550  # ScummVM-wrapped game
    }
    
    # Partial folder name match
    if ($exeName -like "*$folderName*") { 
        $score += 100 
    }
    
    # Game name parts in executable
    foreach ($part in $folderParts) {
        if ($part.Length -gt 3 -and $exeName -like "*$part*") {
            $score += 25
        }
    }
    
    # Priority for executables in typical GOG folders
    $exeLocation = [System.IO.Path]::GetDirectoryName($exePath).ToLower()
    $gogGoodPaths = @("\bin", "\game", "\app", "\dosbox", "\scummvm", "\executable")
    foreach ($goodPath in $gogGoodPaths) {
        if ($exeLocation -like "*$goodPath*") {
            $score += 15
            break
        }
    }
    
    # Bonus for being in root directory (very common for GOG)
    if ($exeLocation -eq $folderPath.ToLower()) {
        $score += 10
    }
    
    # Slight penalty for very deep nesting (GOG games are usually simpler)
    $folderDepth = ($exePath.Split('\').Count - $rootGameFolder.Split('\').Count)
    if ($folderDepth -gt 4) {
        $score -= 5
    }
    
    return $score
}

function Create-GOGShortcut {
    param (
        [string]$targetExe,
        [string]$gameFolderName,
        [string]$outputFolder
    )
    
    # Create sentence case shortcut name (e.g., "Kona Day One.lnk", "Balls of Steel.lnk")
    $shortcutName = $gameFolderName -replace '[^\w\s-]', '' -replace '\s+', ' ' -replace '^\s+|\s+$', ''
    
    # Ensure proper sentence case
    $shortcutName = (Get-Culture).TextInfo.ToTitleCase($shortcutName.ToLower())
    
    # Limit length to prevent issues
    if ($shortcutName.Length -gt 45) {
        $shortcutName = $shortcutName.Substring(0, 45).TrimEnd()
    }
    
    $shortcutPath = "$outputFolder\$shortcutName.lnk"
    
    # Check if shortcut already exists and points to same target
    if (Test-Path $shortcutPath) {
        try {
            $WScriptShell = New-Object -ComObject WScript.Shell
            $existingShortcut = $WScriptShell.CreateShortcut($shortcutPath)
            $existingTarget = $existingShortcut.TargetPath
            
            if ($existingTarget -eq $targetExe) {
                return "exists_same"
            } else {
                return "exists_different"
            }
        } catch {
            Write-Host "Error checking existing shortcut: $_" -ForegroundColor Yellow
            return "error"
        }
    }
    
    # Check if we've already created a shortcut to this executable
    if ($createdShortcuts.ContainsKey($targetExe)) {
        return "duplicate"
    }
    
    # Create the shortcut
    try {
        $WScriptShell = New-Object -ComObject WScript.Shell
        $shortcut = $WScriptShell.CreateShortcut($shortcutPath)
        $shortcut.TargetPath = $targetExe
        $shortcut.WorkingDirectory = [System.IO.Path]::GetDirectoryName($targetExe)
        
        # Try to set description
        $shortcut.Description = "GOG Game: $gameFolderName"
        
        $shortcut.Save()
        
        # Add to tracking hashtable
        $createdShortcuts[$targetExe] = $shortcutName
        
        return "created"
    } catch {
        Write-Host "Error creating shortcut for $gameFolderName`: $_" -ForegroundColor Yellow
        return "error"
    }
}

function Format-TimeSpan {
    param ([TimeSpan]$TimeSpan)
    
    if ($TimeSpan.TotalHours -ge 1) {
        return "{0:h\:mm\:ss}" -f $TimeSpan
    } else {
        return "{0:mm\:ss}" -f $TimeSpan
    }
}

function Find-GOGExecutables {
    param (
        [string]$folderPath,
        [int]$maxDepth,
        [string]$gameName,
        [int]$timeout
    )
    
    Write-Progress -Id 2 -Activity "Finding GOG executables" -Status "Scanning $gameName..."
    
    $exeFiles = @()
    
    try {
        # GOG-optimized search paths (simpler structure than Steam games)
        $gogCommonPaths = @(
            "$folderPath\*.exe",                    # ROOT LEVEL - most common for GOG
            "$folderPath\bin\*.exe",                # Binary folder
            "$folderPath\game\*.exe",               # Game folder
            "$folderPath\app\*.exe",                # Application folder
            "$folderPath\executable\*.exe",         # Executable folder
            "$folderPath\DOSBOX\*.exe",             # DOSBox games
            "$folderPath\ScummVM\*.exe",            # ScummVM games
            "$folderPath\support\*.exe"             # GOG support folder
        )
        
        foreach ($path in $gogCommonPaths) {
            if (Test-Path $path) {
                $foundExes = Get-ChildItem -Path $path -ErrorAction SilentlyContinue
                if ($foundExes -and $foundExes.Count -gt 0) {
                    Write-Host "Found GOG executables in: $path" -ForegroundColor Green
                    $exeFiles += $foundExes
                }
            }
        }
        
        # If we found executables in common paths, don't do expensive searches
        if ($exeFiles.Count -gt 0) {
            Write-Progress -Id 2 -Activity "Finding GOG executables" -Status "Found executables in common paths" -Completed
            return $exeFiles
        }
        
        # Progressive depth search for GOG games (simpler than Steam)
        for ($depth = 1; $depth -le $maxDepth; $depth++) {
            $exeFiles = Get-ChildItem -Path $folderPath -Filter "*.exe" -File -Depth $depth -ErrorAction SilentlyContinue
            if ($exeFiles.Count -gt 0) {
                break
            }
        }
        
        # Final recursive search with timeout if needed
        if ($exeFiles.Count -eq 0) {
            $scriptBlock = {
                param($path)
                Get-ChildItem -Path $path -Filter "*.exe" -File -Recurse -ErrorAction SilentlyContinue
            }
            
            $job = Start-Job -ScriptBlock $scriptBlock -ArgumentList $folderPath
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
        Write-Host "Error scanning $folderPath for GOG executables: $_" -ForegroundColor Red
    }
    
    Write-Progress -Id 2 -Activity "Finding GOG executables" -Completed
    return $exeFiles
}

# Get list of game folders to process
$gameFolders = Get-ChildItem -Path $rootGameFolder -Directory
$total = $gameFolders.Count
$index = 0
$created = 0
$skipped = 0
$notFound = 0
$errors = 0
$timeouts = 0

# Start timing
$startTime = Get-Date
$lastUpdateTime = $startTime

Write-Host "=====================================================================" -ForegroundColor Cyan
Write-Host "GOG Games Shortcut Generator v1.0" -ForegroundColor Cyan
Write-Host "Created by: [Your Name] | $(Get-Date -Format 'yyyy-MM-dd')" -ForegroundColor Green
Write-Host "=====================================================================" -ForegroundColor Cyan
Write-Host "Processing $total GOG game folders..."
Write-Host "Shortcuts will be named: [FolderName].lnk (e.g., Balls of Steel.lnk, Kona Day One.lnk)" -ForegroundColor Yellow

foreach ($folder in $gameFolders) {
    $index++
    $gameName = $folder.Name
    
    # Calculate timing statistics
    $currentTime = Get-Date
    $elapsedTime = $currentTime - $startTime
    $itemsRemaining = $total - $index
    
    if (($index % 5 -eq 0) -or (($currentTime - $lastUpdateTime).TotalSeconds -ge 10)) {
        $lastUpdateTime = $currentTime
        
        if ($index -gt 1) {
            $averageTimePerItem = $elapsedTime.TotalSeconds / ($index - 1)
            $estimatedTimeRemaining = [TimeSpan]::FromSeconds($averageTimePerItem * $itemsRemaining)
            $formattedTimeRemaining = Format-TimeSpan -TimeSpan $estimatedTimeRemaining
            $formattedElapsedTime = Format-TimeSpan -TimeSpan $elapsedTime
            
            $statusMessage = "$gameName ($index of $total) - $created created, $skipped skipped, $notFound not found, $errors errors"
            $progressStatus = "$statusMessage | Elapsed: $formattedElapsedTime | Remaining: $formattedTimeRemaining"
        } else {
            $progressStatus = "$gameName ($index of $total)"
        }
    } else {
        $progressStatus = "$gameName ($index of $total) - $created created, $skipped skipped, $notFound not found, $errors errors"
    }
    
    Write-Progress -Id 1 -Activity "Processing GOG Games" -Status $progressStatus -PercentComplete (($index / $total) * 100)
    
    try {
        $searchStartTime = Get-Date
        $exeFiles = Find-GOGExecutables -folderPath $folder.FullName -maxDepth $maxDepth -gameName $gameName -timeout $folderTimeout
        $searchTime = (Get-Date) - $searchStartTime
        
        # Check for timeout
        if ($searchTime.TotalSeconds -gt ($folderTimeout * 0.9)) {
            Write-Host "Warning: $gameName search took $($searchTime.TotalSeconds) seconds" -ForegroundColor Yellow
            $timeouts++
            "$gameName - Timeout reached while scanning for executables" | Out-File -FilePath $logErrors -Append
        }
        
        if ($exeFiles.Count -eq 0) {
            "$gameName - No executable found" | Out-File -FilePath $logNotFound -Append
            $notFound++
            continue
        }
        
        # Score and sort executables
        $scored = $exeFiles | ForEach-Object {
            [PSCustomObject]@{
                Path  = $_.FullName
                Score = Get-GOGExecutableScore -exePath $_.FullName -gameFolderName $gameName -folderPath $folder.FullName
            }
        } | Where-Object { $_.Score -ge 0 } | Sort-Object Score -Descending
        
        if ($scored.Count -eq 0) {
            "$gameName - No suitable exe (all files blocked by filters)" | Out-File -FilePath $logNotFound -Append
            foreach ($exe in $exeFiles) {
                $exeName = [System.IO.Path]::GetFileNameWithoutExtension($exe.FullName).ToLower()
                "$gameName - Found but filtered: $($exe.FullName) (score would be negative)" | Out-File -FilePath $logBlocked -Append
            }
            $notFound++
            continue
        }
        
        $chosenExe = $scored[0].Path
        $shortcutStatus = Create-GOGShortcut -targetExe $chosenExe -gameFolderName $gameName -outputFolder $shortcutOutputFolder
        
        # Log based on the result
        switch ($shortcutStatus) {
            "created" {
                $shortcutName = $gameName -replace '[^\w\s-]', '' -replace '\s+', ' ' -replace '^\s+|\s+$', ''
                $shortcutName = (Get-Culture).TextInfo.ToTitleCase($shortcutName.ToLower())
                "$gameName -> $shortcutName.lnk - $chosenExe" | Out-File -FilePath $logFound -Append
                $created++
            }
            "exists_same" {
                "$gameName - Shortcut exists and points to same executable: $chosenExe" | Out-File -FilePath $logSkipped -Append
                $skipped++
            }
            "exists_different" {
                "$gameName - Shortcut exists but points to different executable. New: $chosenExe" | Out-File -FilePath $logSkipped -Append
                $skipped++
            }
            "duplicate" {
                "$gameName - Duplicate executable already used for: $($createdShortcuts[$chosenExe])" | Out-File -FilePath $logSkipped -Append
                $skipped++
            }
            "error" {
                "$gameName - Error creating shortcut for: $chosenExe" | Out-File -FilePath $logErrors -Append
                $errors++
            }
        }
    } catch {
        Write-Host "Error processing $gameName`: $_" -ForegroundColor Red
        "$gameName - Error: $_" | Out-File -FilePath $logErrors -Append
        $errors++
    }
}

$totalTime = (Get-Date) - $startTime
$formattedTotalTime = Format-TimeSpan -TimeSpan $totalTime

Write-Host "=====================================================================" -ForegroundColor Cyan
Write-Host "✅ GOG Shortcut Generation Complete!" -ForegroundColor Green
Write-Host "⏱️ Total time: $formattedTotalTime" -ForegroundColor White
Write-Host "📊 Results: $created created, $skipped skipped, $notFound not found, $errors errors, $timeouts timeouts" -ForegroundColor White
Write-Host "📁 Shortcuts saved to: $shortcutOutputFolder" -ForegroundColor Yellow
Write-Host "📄 See FoundGames.log, NotFoundGames.log, SkippedGames.log, ErrorGames.log, and BlockedFiles.log for details" -ForegroundColor White
Write-Host "=====================================================================" -ForegroundColor Cyan