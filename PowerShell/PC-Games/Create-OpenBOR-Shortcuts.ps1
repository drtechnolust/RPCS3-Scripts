# OpenBOR Shortcut Generator - Enhanced for Special Characters in Folder Names
# This version properly handles folders with brackets, ampersands, and other special characters

$rootGameFolder = "D:\Arcade\System roms\OpenBOR Games"
$shortcutOutputFolder = "$rootGameFolder\Shortcuts"
$maxDepth = 3
$folderTimeout = 30
$logFound = "$shortcutOutputFolder\FoundGames.log"
$logNotFound = "$shortcutOutputFolder\NotFoundGames.log"
$logSkipped = "$shortcutOutputFolder\SkippedGames.log"
$logBlocked = "$shortcutOutputFolder\BlockedFiles.log"
$logErrors = "$shortcutOutputFolder\ErrorGames.log"

# Ensure shortcut output folder exists
if (!(Test-Path -LiteralPath $shortcutOutputFolder)) {
    New-Item -ItemType Directory -Path $shortcutOutputFolder -Force | Out-Null
}

# Clear logs if they exist
@($logFound, $logNotFound, $logSkipped, $logErrors, $logBlocked) | ForEach-Object {
    if (Test-Path -LiteralPath $_) {
        Remove-Item -LiteralPath $_ -Force -ErrorAction SilentlyContinue
    }
}

# Store a hashtable of created shortcuts to check for duplicates
$createdShortcuts = @{}

function Get-SafePath {
    param([string]$Path)
    
    # Always use -LiteralPath for paths with special characters
    # This function helps ensure we're using literal paths consistently
    return $Path
}

function Test-SafePath {
    param(
        [string]$Path,
        [string]$PathType = "Any"
    )
    
    try {
        if ($PathType -eq "Container") {
            return Test-Path -LiteralPath $Path -PathType Container -ErrorAction Stop
        } elseif ($PathType -eq "Leaf") {
            return Test-Path -LiteralPath $Path -PathType Leaf -ErrorAction Stop
        } else {
            return Test-Path -LiteralPath $Path -ErrorAction Stop
        }
    } catch {
        return $false
    }
}

function Get-SafeChildItem {
    param(
        [string]$Path,
        [string]$Filter = $null,
        [switch]$File,
        [switch]$Directory,
        [switch]$Recurse
    )
    
    $params = @{
        LiteralPath = $Path
        ErrorAction = 'Stop'
    }
    
    if ($Filter) { $params['Filter'] = $Filter }
    if ($File) { $params['File'] = $true }
    if ($Directory) { $params['Directory'] = $true }
    if ($Recurse) { $params['Recurse'] = $true }
    
    try {
        return Get-ChildItem @params
    } catch {
        Write-Host "    Error accessing path: $_" -ForegroundColor Red
        return @()
    }
}

function Join-SafePath {
    param(
        [string]$Path,
        [string]$ChildPath
    )
    
    # Use [System.IO.Path]::Combine for safer path joining
    return [System.IO.Path]::Combine($Path, $ChildPath)
}

function Get-OpenBORExecutableScore {
    param($exePath, $gameFolderName)
    
    try {
        $exeName = [System.IO.Path]::GetFileNameWithoutExtension($exePath).ToLower()
        $score = 0
        
        Write-Host "      Analyzing: $exeName" -ForegroundColor DarkYellow
        
        # Block obviously non-game executables
        $blockPatterns = @(
            "*uninstall*", "*setup*", "*install*", "*update*", 
            "*config*", "*settings*", "*helper*", "*crash*", 
            "*test*", "*service*", "*server*", "*redist*"
        )
        
        foreach ($pattern in $blockPatterns) {
            if ($exeName -like $pattern) {
                Write-Host "        BLOCKED by pattern: $pattern" -ForegroundColor Red
                return -1
            }
        }
        
        # HIGHEST PRIORITY: OpenBOR.exe is the standard executable name
        if ($exeName -eq "openbor") { 
            Write-Host "        MATCH: Standard OpenBOR.exe (1000 points)" -ForegroundColor Green
            return 1000
        }
        
        # HIGH PRIORITY: Variations of OpenBOR executable names
        $openborVariations = @("openbor-win", "openbor64", "openbor32", "openbor_win", "openborengine", "openbor-win64")
        if ($openborVariations -contains $exeName) {
            Write-Host "        MATCH: OpenBOR variation (500 points)" -ForegroundColor Green
            return 500
        }
        
        # MEDIUM-HIGH PRIORITY: Any executable with "openbor" in the name
        if ($exeName -like "*openbor*") {
            $score += 200
            Write-Host "        PARTIAL: Contains 'openbor' (+200 points)" -ForegroundColor Yellow
        }
        
        # Priority for executables in root directory
        try {
            $exeLocation = [System.IO.Path]::GetDirectoryName($exePath)
            $currentFolder = Get-Item -LiteralPath $exeLocation -ErrorAction SilentlyContinue
            
            if ($currentFolder) {
                # Check if we're in the game's root folder
                $gameRootFolder = Get-Item -LiteralPath (Join-SafePath $rootGameFolder $gameFolderName) -ErrorAction SilentlyContinue
                if ($gameRootFolder -and $currentFolder.FullName -eq $gameRootFolder.FullName) {
                    $score += 100
                    Write-Host "        BONUS: In game root folder (+100 points)" -ForegroundColor Yellow
                }
            }
        } catch {
            # Simple depth check as fallback
            $relativePath = $exePath.Substring($rootGameFolder.Length)
            $depth = ($relativePath -split '\\').Count - 2
            if ($depth -eq 1) {
                $score += 50
                Write-Host "        BONUS: Appears to be in root (+50 points)" -ForegroundColor Yellow
            }
        }
        
        # LOWER PRIORITY: Alternative names
        $alternativeNames = @("bor", "beatsofrage", "game", "start", "play", "run", "main")
        if ($alternativeNames -contains $exeName) {
            $score += 50
            Write-Host "        MATCH: Alternative name (+50 points)" -ForegroundColor Yellow
        }
        
        # Penalize if too deep in folder structure
        $folderDepth = ($exePath.Split('\').Count - $rootGameFolder.Split('\').Count)
        if ($folderDepth -gt 3) {
            $score -= 30
            Write-Host "        PENALTY: Too deep in folders (-30 points)" -ForegroundColor Yellow
        }
        
        Write-Host "        FINAL SCORE: $score points" -ForegroundColor $(if ($score -ge 100) { "Green" } elseif ($score -gt 0) { "Yellow" } else { "Red" })
        return $score
        
    } catch {
        Write-Host "        ERROR in scoring: $_" -ForegroundColor Red
        return 0
    }
}

function Convert-ToSentenceCase {
    param([string]$text)
    
    if ([string]::IsNullOrWhiteSpace($text)) {
        return $text
    }
    
    try {
        $cleaned = $text.Trim()
        
        # Remove common prefixes
        if ($cleaned.StartsWith("OpenBOR - ", [StringComparison]::InvariantCultureIgnoreCase)) {
            $cleaned = $cleaned.Substring(10).Trim()
        } elseif ($cleaned.StartsWith("BOR - ", [StringComparison]::InvariantCultureIgnoreCase)) {
            $cleaned = $cleaned.Substring(6).Trim()
        }
        
        # Convert to sentence case
        $textInfo = (Get-Culture).TextInfo
        $result = $textInfo.ToTitleCase($cleaned.ToLower())
        
        # Fix Roman numerals
        $result = $result -replace '\b(I{1,3}|IV|V|VI{1,3}|IX|X)\b', { $_.Value.ToUpper() }
        
        return $result.Trim()
        
    } catch {
        return $text.Trim()
    }
}

function Create-Shortcut {
    param (
        [string]$targetExe,
        [string]$shortcutName,
        [string]$outputFolder
    )
    
    Write-Host "    Creating shortcut for: '$shortcutName'" -ForegroundColor Cyan
    
    try {
        # Validate inputs
        if ([string]::IsNullOrWhiteSpace($targetExe) -or 
            [string]::IsNullOrWhiteSpace($shortcutName) -or 
            [string]::IsNullOrWhiteSpace($outputFolder)) {
            Write-Host "    ✗ Invalid parameters" -ForegroundColor Red
            return "error"
        }
        
        # Validate target exe exists
        if (!(Test-SafePath $targetExe -PathType "Leaf")) {
            Write-Host "    ✗ Target executable does not exist: '$targetExe'" -ForegroundColor Red
            return "error"
        }
        
        # Ensure output folder exists
        if (!(Test-SafePath $outputFolder -PathType "Container")) {
            New-Item -ItemType Directory -Path $outputFolder -Force | Out-Null
        }
        
        # Convert to sentence case and sanitize
        $sentenceCaseShortcutName = Convert-ToSentenceCase -text $shortcutName
        $safeShortcutName = $sentenceCaseShortcutName -replace '[\\\/\:\*\?"<>\|]', '_'
        $safeShortcutName = $safeShortcutName -replace '&', 'and'
        
        # Limit filename length
        if ($safeShortcutName.Length -gt 50) {
            $safeShortcutName = $safeShortcutName.Substring(0, 47) + "..."
        }
        
        # Check for duplicates
        $shortcutPath = Join-SafePath $outputFolder "$safeShortcutName.lnk"
        
        if (Test-SafePath $shortcutPath) {
            Write-Host "    ↻ Shortcut already exists" -ForegroundColor Blue
            return "exists_same"
        }
        
        if ($createdShortcuts.ContainsKey($targetExe)) {
            Write-Host "    ↻ Duplicate executable" -ForegroundColor Blue
            return "duplicate"
        }
        
        # Create the shortcut
        $WScriptShell = New-Object -ComObject WScript.Shell
        $shortcut = $WScriptShell.CreateShortcut($shortcutPath)
        $shortcut.TargetPath = $targetExe
        $shortcut.WorkingDirectory = [System.IO.Path]::GetDirectoryName($targetExe)
        $shortcut.Save()
        
        # Verify creation
        if (Test-SafePath $shortcutPath) {
            $createdShortcuts[$targetExe] = $sentenceCaseShortcutName
            Write-Host "    ✅ Shortcut created successfully" -ForegroundColor Green
            return "created"
        } else {
            Write-Host "    ✗ Shortcut creation failed" -ForegroundColor Red
            return "error"
        }
        
    } catch {
        Write-Host "    ✗ Error in Create-Shortcut: $_" -ForegroundColor Red
        return "error"
    }
}

function Find-OpenBORExecutables {
    param (
        [string]$folderPath,
        [int]$maxDepth,
        [string]$gameName,
        [int]$timeout
    )
    
    $displayName = if ($gameName.Length -gt 50) { $gameName.Substring(0, 47) + "..." } else { $gameName }
    Write-Host "Scanning: $displayName..." -ForegroundColor Cyan
    
    $exeFiles = @()
    
    try {
        # Validate folder access
        if (!(Test-SafePath $folderPath -PathType "Container")) {
            Write-Host "  ✗ Folder does not exist or is not accessible: $folderPath" -ForegroundColor Red
            "$gameName - Folder not accessible: $folderPath" | Out-File -FilePath $logErrors -Append -ErrorAction SilentlyContinue
            return @()
        }
        
        Write-Host "  ✓ Folder accessible" -ForegroundColor Green
        
        # FIRST: Check for OpenBOR.exe in root directory
        Write-Host "  Checking root for OpenBOR.exe..." -ForegroundColor Gray
        
        $openBORVariants = @("OpenBOR.exe", "openbor.exe", "OPENBOR.EXE", "OpenBor.exe")
        foreach ($variant in $openBORVariants) {
            $testPath = Join-SafePath $folderPath $variant
            if (Test-SafePath $testPath -PathType "Leaf") {
                Write-Host "  ✓ Found $variant in root!" -ForegroundColor Green
                try {
                    $foundFile = Get-Item -LiteralPath $testPath -ErrorAction Stop
                    $exeFiles += $foundFile
                    return $exeFiles
                } catch {
                    Write-Host "    Error accessing file: $_" -ForegroundColor Yellow
                }
            }
        }
        
        # SECOND: Search for any .exe files with OpenBOR patterns
        Write-Host "  Searching for OpenBOR executables..." -ForegroundColor Gray
        $allExeFiles = Get-SafeChildItem -Path $folderPath -Filter "*.exe" -File
        
        foreach ($exe in $allExeFiles) {
            $exeName = $exe.Name.ToLower()
            if ($exeName -like "*openbor*" -or $exeName -like "*bor*") {
                Write-Host "    Found potential OpenBOR file: $($exe.Name)" -ForegroundColor Yellow
                $exeFiles += $exe
            }
        }
        
        if ($exeFiles.Count -gt 0) {
            return $exeFiles
        }
        
        # THIRD: Check common subfolders
        Write-Host "  Checking subfolders..." -ForegroundColor Gray
        $subfolders = @("bin", "engine", "OpenBOR", "data")
        foreach ($subfolder in $subfolders) {
            $subPath = Join-SafePath $folderPath $subfolder
            if (Test-SafePath $subPath -PathType "Container") {
                $subExeFiles = Get-SafeChildItem -Path $subPath -Filter "*.exe" -File
                foreach ($exe in $subExeFiles) {
                    if ($exe.Name -like "*openbor*" -or $exe.Name -eq "OpenBOR.exe") {
                        Write-Host "  ✓ Found $($exe.Name) in $subfolder!" -ForegroundColor Green
                        $exeFiles += $exe
                    }
                }
            }
        }
        
        # FOURTH: If still nothing, get all exe files for scoring
        if ($exeFiles.Count -eq 0) {
            Write-Host "  Getting all .exe files for analysis..." -ForegroundColor Gray
            $exeFiles = Get-SafeChildItem -Path $folderPath -Filter "*.exe" -File
        }
        
        Write-Host "  Final result: Found $($exeFiles.Count) executable(s)" -ForegroundColor $(if ($exeFiles.Count -gt 0) { "Green" } else { "Red" })
        
    } catch {
        Write-Host "  ✗ Error during search: $_" -ForegroundColor Red
        "$gameName - Search error: $folderPath - $_" | Out-File -FilePath $logErrors -Append -ErrorAction SilentlyContinue
    }
    
    return $exeFiles
}

function Format-TimeSpan {
    param ([TimeSpan]$TimeSpan)
    
    if ($TimeSpan.TotalHours -ge 1) {
        return "{0:h\:mm\:ss}" -f $TimeSpan
    } else {
        return "{0:mm\:ss}" -f $TimeSpan
    }
}

# Main script execution
Write-Host "Starting OpenBOR Shortcut Generator..." -ForegroundColor Green
Write-Host "Source: $rootGameFolder" -ForegroundColor Cyan
Write-Host "Output: $shortcutOutputFolder" -ForegroundColor Cyan
Write-Host ""

Write-Host "Enumerating game folders..." -ForegroundColor Yellow
try {
    $gameFolders = Get-SafeChildItem -Path $rootGameFolder -Directory
    Write-Host "Found $($gameFolders.Count) game folders" -ForegroundColor Green
    
    # Show sample folder names
    Write-Host "Sample folder names:" -ForegroundColor Gray
    for ($i = 0; $i -lt [Math]::Min(5, $gameFolders.Count); $i++) {
        Write-Host "  [$($i+1)] '$($gameFolders[$i].Name)'" -ForegroundColor Gray
    }
    Write-Host ""
    
} catch {
    Write-Host "Error enumerating folders: $_" -ForegroundColor Red
    Write-Host "Please check the root game folder path: $rootGameFolder" -ForegroundColor Yellow
    exit 1
}

$total = $gameFolders.Count
$index = 0
$created = 0
$skipped = 0
$notFound = 0
$errors = 0

$startTime = Get-Date

foreach ($folder in $gameFolders) {
    $index++
    $gameName = $folder.Name
    
    # Update progress
    $percentComplete = ($index / $total) * 100
    Write-Progress -Id 1 -Activity "Processing OpenBOR Games" -Status "$gameName ($index of $total)" -PercentComplete $percentComplete
    
    try {
        Write-Host "Processing [$index/$total]: $gameName" -ForegroundColor White
        
        # Use folder's FullName directly
        $actualPath = $folder.FullName
        
        # Find executables
        $exeFiles = Find-OpenBORExecutables -folderPath $actualPath -maxDepth $maxDepth -gameName $gameName -timeout $folderTimeout
        
        if ($exeFiles.Count -eq 0) {
            "$gameName - No OpenBOR executable found" | Out-File -FilePath $logNotFound -Append
            Write-Host "  ✗ No executables found" -ForegroundColor Red
            $notFound++
            continue
        }
        
        # Score and filter executables
        $scored = @()
        foreach ($exe in $exeFiles) {
            $score = Get-OpenBORExecutableScore -exePath $exe.FullName -gameFolderName $gameName
            
            if ($score -ge 0) {
                $scored += [PSCustomObject]@{
                    Path  = $exe.FullName
                    Score = $score
                }
            }
        }
        
        if ($scored.Count -eq 0) {
            "$gameName - No suitable executable found (all blocked)" | Out-File -FilePath $logNotFound -Append
            Write-Host "  ✗ All executables blocked by filters" -ForegroundColor Red
            $notFound++
            continue
        }
        
        # Select best executable
        $scored = $scored | Sort-Object Score -Descending
        $chosenExe = $scored[0].Path
        
        Write-Host "  Selected: $(Split-Path $chosenExe -Leaf) (score: $($scored[0].Score))" -ForegroundColor Green
        
        # Create shortcut
        $shortcutStatus = Create-Shortcut -targetExe $chosenExe -shortcutName $gameName -outputFolder $shortcutOutputFolder
        
        switch ($shortcutStatus) {
            "created" {
                $shortcutName = $createdShortcuts[$chosenExe]
                "$gameName -> '$shortcutName' - $chosenExe" | Out-File -FilePath $logFound -Append
                Write-Host "  ✓ Created: '$shortcutName'" -ForegroundColor Green
                $created++
            }
            "exists_same" {
                "$gameName - Shortcut already exists" | Out-File -FilePath $logSkipped -Append
                $skipped++
            }
            "duplicate" {
                "$gameName - Duplicate executable" | Out-File -FilePath $logSkipped -Append
                $skipped++
            }
            "error" {
                "$gameName - Error creating shortcut" | Out-File -FilePath $logErrors -Append
                Write-Host "  ✗ Error creating shortcut" -ForegroundColor Red
                $errors++
            }
        }
        
    } catch {
        Write-Host "  ✗ Error processing $gameName`: $_" -ForegroundColor Red
        "$gameName - Processing error: $_" | Out-File -FilePath $logErrors -Append
        $errors++
    }
}

$totalTime = (Get-Date) - $startTime
$formattedTotalTime = Format-TimeSpan -TimeSpan $totalTime

Write-Host ""
Write-Host "✅ OpenBOR Shortcut Generation Complete!" -ForegroundColor Green
Write-Host "🎮 Shortcuts saved to: $shortcutOutputFolder" -ForegroundColor Cyan
Write-Host "⏱️ Total time: $formattedTotalTime" -ForegroundColor Yellow
Write-Host "📊 Results:" -ForegroundColor White
Write-Host "   • $created shortcuts created" -ForegroundColor Green
Write-Host "   • $skipped shortcuts skipped (already exist)" -ForegroundColor Blue
Write-Host "   • $notFound games with no suitable executable found" -ForegroundColor Red
Write-Host "   • $errors errors encountered" -ForegroundColor Red
Write-Host ""
Write-Host "📄 Check the log files for detailed information:" -ForegroundColor White
Write-Host "   • FoundGames.log - Successfully created shortcuts" -ForegroundColor Green
Write-Host "   • NotFoundGames.log - Games where no executable was found" -ForegroundColor Red
Write-Host "   • SkippedGames.log - Games that were skipped" -ForegroundColor Blue
Write-Host "   • ErrorGames.log - Games that caused errors" -ForegroundColor Red
Write-Host "   • BlockedFiles.log - Executables blocked by filters" -ForegroundColor Yellow