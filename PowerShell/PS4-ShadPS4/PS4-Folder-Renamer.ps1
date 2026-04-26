# PS4 Game Folder Formatter - FIXED VERSION
param(
    [string]$Path = "",
    [switch]$WhatIf,
    [switch]$Verbose
)

function Select-SourceFolder {
    Write-Host "PS4 Game Folder Formatter - FIXED VERSION" -ForegroundColor Cyan
    Write-Host "=========================================" -ForegroundColor Cyan
    Write-Host ""
    
    do {
        Write-Host "Enter folder path (or 'quit' to exit): " -NoNewline -ForegroundColor White
        $inputPath = Read-Host
        
        if ($inputPath.ToLower() -in @('quit', 'q', 'exit')) {
            return $null
        }
        
        $inputPath = $inputPath.Trim().Trim('"').Trim("'").Trim()
        
        if ([string]::IsNullOrWhiteSpace($inputPath)) {
            Write-Host "Please enter a valid path." -ForegroundColor Red
            continue
        }
        
        if (Test-Path $inputPath -PathType Container) {
            Write-Host "✅ Valid folder found!" -ForegroundColor Green
            return $inputPath
        } else {
            Write-Host "❌ Folder '$inputPath' does not exist." -ForegroundColor Red
        }
        
    } while ($true)
}

# PROPER formatting rules
$gameRules = @{
    # Words that should always be UPPERCASE
    'UppercaseWords' = @('NBA', 'NFL', 'WWE', 'UFC', 'HD', 'VR', 'DLC', 'GOTY', 'PS4', 'PS5', 'PC', 'US', 'EU', 'UK', 'JP', 'USA', 'RPG', 'FPS', 'RTS', 'DX')
    
    # Roman numerals
    'RomanNumerals' = @('II', 'III', 'IV', 'V', 'VI', 'VII', 'VIII', 'IX', 'X', 'XI', 'XII', 'XIII', 'XIV', 'XV', 'XVI', 'XVII', 'XVIII', 'XIX', 'XX')
    
    # Words that should stay lowercase (except at start)
    'LowercaseWords' = @('a', 'an', 'and', 'as', 'at', 'but', 'by', 'for', 'if', 'in', 'nor', 'of', 'on', 'or', 'so', 'the', 'to', 'up', 'yet', 'vs', 'with')
    
    # Common corrections
    'Corrections' = @{
        'spiderman' = 'Spider-Man'
        'xmen' = 'X-Men'
        'pacman' = 'Pac-Man'
        'cyberpunk' = 'Cyberpunk'
        'bloodborne' = 'Bloodborne'
    }
    
    # Series patterns
    'SeriesPatterns' = @{
        '^call of duty (.+)$' = 'Call of Duty $1'
        '^metal gear solid (.+)$' = 'Metal Gear Solid $1'
        '^grand theft auto(.*)$' = 'Grand Theft Auto$1'
        '^detroit become human(.*)$' = 'Detroit Become Human$1'
        '^god of war (.+)$' = 'God of War $1'
        '^dragon ball (.+)$' = 'Dragon Ball $1'
        '^dead rising(.*)$' = 'Dead Rising$1'
    }
}

function Format-GameTitle {
    param([string]$title)
    
    if ([string]::IsNullOrWhiteSpace($title)) { return $title }
    
    Write-Host "    DEBUG: Processing '$title'" -ForegroundColor Gray
    
    $result = $title.Trim().ToLower()
    
    # Handle Japanese/Chinese characters - preserve them
    if ($result -match '[\p{IsHiragana}\p{IsKatakana}\p{IsCJKUnifiedIdeographs}]') {
        Write-Host "    DEBUG: Contains Asian characters, preserving as-is" -ForegroundColor Gray
        return $title.Trim()
    }
    
    # Apply series patterns
    foreach ($pattern in $gameRules.SeriesPatterns.Keys) {
        if ($result -match $pattern) {
            $replacement = $gameRules.SeriesPatterns[$pattern]
            $result = $result -ireplace $pattern, $replacement
            Write-Host "    DEBUG: Applied series pattern: $pattern" -ForegroundColor Yellow
            break
        }
    }
    
    # Apply common corrections
    foreach ($wrong in $gameRules.Corrections.Keys) {
        $correct = $gameRules.Corrections[$wrong]
        if ($result -match [regex]::Escape($wrong)) {
            $result = $result -ireplace [regex]::Escape($wrong), $correct
            Write-Host "    DEBUG: Applied correction: $wrong -> $correct" -ForegroundColor Yellow
        }
    }
    
    # Apply title case with rules
    $words = $result -split '\s+'
    $formatted = @()
    
    for ($i = 0; $i -lt $words.Length; $i++) {
        $word = $words[$i].Trim()
        if ([string]::IsNullOrEmpty($word)) { continue }
        
        # Remove punctuation for comparison
        $cleanWord = $word -replace '[^\w]', ''
        $lowerClean = $cleanWord.ToLower()
        $upperClean = $cleanWord.ToUpper()
        
        # Check if it should be uppercase
        if ($gameRules.UppercaseWords -contains $upperClean) {
            $formatted += $word -ireplace [regex]::Escape($cleanWord), $upperClean
            Write-Host "    DEBUG: Made uppercase: $cleanWord -> $upperClean" -ForegroundColor Green
        }
        # Check if it's a roman numeral
        elseif ($gameRules.RomanNumerals -contains $upperClean) {
            $formatted += $word -ireplace [regex]::Escape($cleanWord), $upperClean
            Write-Host "    DEBUG: Roman numeral: $cleanWord -> $upperClean" -ForegroundColor Green
        }
        # Check if it should stay lowercase (not at start)
        elseif ($i -gt 0 -and $gameRules.LowercaseWords -contains $lowerClean) {
            $formatted += $word -ireplace [regex]::Escape($cleanWord), $lowerClean
            Write-Host "    DEBUG: Kept lowercase: $cleanWord -> $lowerClean" -ForegroundColor Green
        }
        # Numbers stay as-is
        elseif ($cleanWord -match '^\d+$') {
            $formatted += $word
        }
        # Default title case
        else {
            $titleCased = $cleanWord.Substring(0,1).ToUpper() + $cleanWord.Substring(1).ToLower()
            $formatted += $word -ireplace [regex]::Escape($cleanWord), $titleCased
            Write-Host "    DEBUG: Title cased: $cleanWord -> $titleCased" -ForegroundColor Green
        }
    }
    
    $final = ($formatted -join ' ')
    
    # Clean up
    $final = $final -replace '\s+', ' '
    $final = $final.Trim()
    
    # Remove Windows-forbidden characters
    $final = $final -replace '[<>:"|?*/\\]', ''
    
    Write-Host "    DEBUG: Final result: '$final'" -ForegroundColor Cyan
    return $final
}

# Main execution
if ([string]::IsNullOrWhiteSpace($Path)) {
    $Path = Select-SourceFolder
    if ($null -eq $Path) {
        exit
    }
}

try {
    $folders = Get-ChildItem -Path $Path -Directory -ErrorAction Stop
    Write-Host "Found $($folders.Count) folders" -ForegroundColor Green
    
    if ($folders.Count -eq 0) {
        Write-Host "No folders found!" -ForegroundColor Red
        Read-Host "Press Enter to exit"
        exit
    }
    
    Write-Host "`nTesting first 10 folders..." -ForegroundColor Yellow
    
    $testFolders = $folders | Select-Object -First 10
    $changesNeeded = 0
    
    foreach ($folder in $testFolders) {
        $originalName = $folder.Name
        $newName = Format-GameTitle $originalName
        
        Write-Host "`nFolder: $originalName" -ForegroundColor White
        if ($originalName -ne $newName) {
            Write-Host "  BEFORE: $originalName" -ForegroundColor Red
            Write-Host "  AFTER:  $newName" -ForegroundColor Green
            $changesNeeded++
        } else {
            Write-Host "  No change needed" -ForegroundColor Gray
        }
    }
    
    Write-Host "`nSummary: $changesNeeded of 10 folders need changes" -ForegroundColor Cyan
    
    if ($changesNeeded -gt 0) {
        Write-Host "`nDo you want to rename ALL folders? (Y/N): " -NoNewline -ForegroundColor Yellow
        $response = Read-Host
        
        if ($response.ToLower() -eq 'y') {
            Write-Host "`nProcessing all folders..." -ForegroundColor Cyan
            $renamed = 0
            $errors = 0
            
            foreach ($folder in $folders) {
                $originalName = $folder.Name
                $newName = Format-GameTitle $originalName
                
                if ($originalName -ne $newName) {
                    try {
                        Write-Host "Renaming: $originalName -> $newName" -ForegroundColor Yellow
                        Rename-Item -Path $folder.FullName -NewName $newName -ErrorAction Stop
                        $renamed++
                        Write-Host "  ✅ Success" -ForegroundColor Green
                    } catch {
                        Write-Host "  ❌ Failed: $($_.Exception.Message)" -ForegroundColor Red
                        $errors++
                    }
                }
            }
            
            Write-Host "`n🎮 COMPLETE! 🎮" -ForegroundColor Cyan
            Write-Host "Successfully renamed: $renamed folders" -ForegroundColor Green
            Write-Host "Errors: $errors" -ForegroundColor $(if ($errors -gt 0) { 'Red' } else { 'Green' })
        }
    } else {
        Write-Host "`nNo changes needed!" -ForegroundColor Green
    }

} catch {
    Write-Host "ERROR: $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host "`nPress Enter to exit..." -ForegroundColor Gray
Read-Host