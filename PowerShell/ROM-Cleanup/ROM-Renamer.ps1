# ROM File Cleanup Script (Improved Version)
param(
    [string]$SourcePath = $null,
    [switch]$Recurse,
    [switch]$Backup,
    [switch]$DryRun,
    [string[]]$ExcludePatterns = @(),
    [string]$ConfigFile = $null,
    [string]$LogFile = $null
)

# Function to log messages
function Write-Log {
    param(
        [string]$Message,
        [string]$Level = "INFO",
        [string]$LogPath
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] [$Level] $Message"
    Add-Content -Path $LogPath -Value $logEntry
    # Also output to console based on level
    switch ($Level) {
        "INFO" { Write-Host $logEntry -ForegroundColor White }
        "WARNING" { Write-Host $logEntry -ForegroundColor Yellow }
        "ERROR" { Write-Host $logEntry -ForegroundColor Red }
    }
}

# Function to convert to title case with gaming-specific rules
function Convert-ToTitleCase {
    param([string]$text, [hashtable]$customReplacements)
    
    # Convert to title case
    $titleCase = (Get-Culture).TextInfo.ToTitleCase($text.ToLower())
    
    # Apply default replacements
    $titleCase = $titleCase -replace '\bGbc\b', 'GBC'
    $titleCase = $titleCase -replace '\bGba\b', 'GBA'
    $titleCase = $titleCase -replace '\bNes\b', 'NES'
    $titleCase = $titleCase -replace '\bSnes\b', 'SNES'
    $titleCase = $titleCase -replace '\bN64\b', 'N64'
    $titleCase = $titleCase -replace '\bPs1\b', 'PS1'
    $titleCase = $titleCase -replace '\bPs2\b', 'PS2'
    $titleCase = $titleCase -replace '\bUsa\b', 'USA'
    $titleCase = $titleCase -replace '\bUk\b', 'UK'
    $titleCase = $titleCase -replace '\bJap\b', 'JAP'
    $titleCase = $titleCase -replace '\bEur\b', 'EUR'
    $titleCase = $titleCase -replace '\bNtsc\b', 'NTSC'
    $titleCase = $titleCase -replace '\bPal\b', 'PAL'
    $titleCase = $titleCase -replace '\bRpg\b', 'RPG'
    $titleCase = $titleCase -replace '\bFps\b', 'FPS'
    $titleCase = $titleCase -replace '\bMmo\b', 'MMO'
    $titleCase = $titleCase -replace '\bDx\b', 'DX'
    $titleCase = $titleCase -replace '\bEx\b', 'EX'
    $titleCase = $titleCase -replace '\bGt\b', 'GT'
    $titleCase = $titleCase -replace '\bHd\b', 'HD'
    $titleCase = $titleCase -replace '\b3d\b', '3D'
    $titleCase = $titleCase -replace '\b2d\b', '2D'
    $titleCase = $titleCase -replace '\bIi\b', 'II'
    $titleCase = $titleCase -replace '\bIii\b', 'III'
    $titleCase = $titleCase -replace '\bIv\b', 'IV'
    $titleCase = $titleCase -replace '\bVi\b', 'VI'
    $titleCase = $titleCase -replace '\bVii\b', 'VII'
    $titleCase = $titleCase -replace '\bViii\b', 'VIII'
    # Expanded abbreviations
    $titleCase = $titleCase -replace '\bDlc\b', 'DLC'
    $titleCase = $titleCase -replace '\bVr\b', 'VR'
    $titleCase = $titleCase -replace '\bAmi\b', 'AMI'
    $titleCase = $titleCase -replace '\bNesica\b', 'NESiCA'
    $titleCase = $titleCase -replace '\bVs\b', 'vs.'
    $titleCase = $titleCase -replace '\bOf\b', 'of'  # Keep lowercase if needed
    
    # Apply custom replacements from config
    foreach ($key in $customReplacements.Keys) {
        $titleCase = $titleCase -replace $key, $customReplacements[$key]
    }
    
    # Fix specific region codes and other common abbreviations
    $titleCase = $titleCase -replace '\(Usa\)', '(USA)'
    $titleCase = $titleCase -replace '\(Eur\)', '(EUR)'
    $titleCase = $titleCase -replace '\(Jap\)', '(JAP)'
    $titleCase = $titleCase -replace '\(Uk\)', '(UK)'
    $titleCase = $titleCase -replace '\(Je\)', '(JE)'
    $titleCase = $titleCase -replace '\(Ue\)', '(UE)'
    $titleCase = $titleCase -replace '\(Ju\)', '(JU)'
    $titleCase = $titleCase -replace '\(E\)', '(E)'
    $titleCase = $titleCase -replace '\(J\)', '(J)'
    $titleCase = $titleCase -replace '\(U\)', '(U)'
    $titleCase = $titleCase -replace '\(Ger\)', '(GER)'
    $titleCase = $titleCase -replace '\(Fra\)', '(FRA)'
    $titleCase = $titleCase -replace '\(Spa\)', '(SPA)'
    $titleCase = $titleCase -replace '\(Ita\)', '(ITA)'
    
    return $titleCase
}

# Function to clean filename
function Clean-Filename {
    param(
        [string]$filename,
        [string]$source,
        [string]$enhancements,
        [bool]$isDirectory = $false,
        [hashtable]$customReplacements = @{}
    )
    
    # Check exclude patterns
    foreach ($pattern in $ExcludePatterns) {
        if ($filename -match $pattern) {
            return $filename  # Skip cleaning if matches exclude
        }
    }
    
    # Remove organizational numbering patterns:
    # - 3-4 digits followed by space: "001 Game Name" or "1133 Game Name"
    # - 3-4 digits followed by underscore-dash-underscore: "992_-_spiderman"
    # - 3-4 digits followed by underscore: "1100_hugo_black"
    $cleaned = $filename -replace '^\d{3,4}[a-z]?(\s+|_-_|_)', ''
    
    # Fix possessive without apostrophe (convert -s to s)
    $cleaned = $cleaned -replace '-s\b', 's'
    
    # Replace underscores with spaces
    $cleaned = $cleaned -replace '_', ' '
    
    # Clean up multiple spaces first
    $cleaned = $cleaned -replace '\s+', ' '
    
    # Only add spaces around dashes that are NOT part of hyphenated words
    # Keep hyphens in words like "Harley-Davidson", "F-1", "3-D" but space standalone dashes
    $cleaned = $cleaned -replace '\s+-\s+', ' - '  # Already spaced dashes stay the same
    $cleaned = $cleaned -replace '(?<=\s)-(?=\s)', '-'  # Remove extra spaces from spaced dashes
    
    # Trim whitespace
    $cleaned = $cleaned.Trim()
    
    # Remove invalid filename characters (Windows-safe)
    $invalidChars = [System.IO.Path]::GetInvalidFileNameChars()
    foreach ($char in $invalidChars) {
        $cleaned = $cleaned -replace [regex]::Escape($char), ''
    }
    
    # For directories, we don't need to handle extensions
    if ($isDirectory) {
        # Apply title case to the entire folder name
        $cleaned = Convert-ToTitleCase $cleaned $customReplacements
    } else {
        # Handle multi-dot extensions (e.g., .tar.gz)
        $extension = $cleaned -replace '^.*(\.[^.]+(\.[^.]+)?)$', '$1'
        $nameWithoutExt = $cleaned.Substring(0, $cleaned.Length - $extension.Length)
        
        # Apply title case to the filename (but not the extension)
        $nameWithoutExt = Convert-ToTitleCase $nameWithoutExt $customReplacements
        
        # Rebuild filename with extension
        $cleaned = "$nameWithoutExt$extension"
    }
    
    # Add source and enhancements if provided
    if ($source -and $source -ne '') {
        if ($isDirectory) {
            $cleaned = "$cleaned [$source]"
        } else {
            $extension = $cleaned -replace '^.*(\.[^.]+(\.[^.]+)?)$', '$1'
            $nameWithoutExt = $cleaned.Substring(0, $cleaned.Length - $extension.Length)
            $cleaned = "$nameWithoutExt [$source]$extension"
        }
    }
    
    if ($enhancements -and $enhancements -ne '') {
        if ($isDirectory) {
            $cleaned = "$cleaned {$enhancements}"
        } else {
            $extension = $cleaned -replace '^.*(\.[^.]+(\.[^.]+)?)$', '$1'
            $nameWithoutExt = $cleaned.Substring(0, $cleaned.Length - $extension.Length)
            $cleaned = "$nameWithoutExt {$enhancements}$extension"
        }
    }
    
    return $cleaned
}

# Function to test the regex pattern (Expanded)
function Test-CleaningPattern {
    param([hashtable]$customReplacements = @{})
    
    Write-Host "`nTesting cleaning pattern on sample names:" -ForegroundColor Cyan
    Write-Host "==========================================" -ForegroundColor Cyan
    
    $testItems = @(
        @{Name="001 Dragon Quest Monsters (J).zip"; IsDir=$false},
        @{Name="010 rampage world tour (E).zip"; IsDir=$false}, 
        @{Name="992_-_spiderman_(j).zip"; IsDir=$false},
        @{Name="1111_-_tony_hawk-s_pro_skater_3.zip"; IsDir=$false},
        @{Name="488 F-1 World Grand Prix 2 (JE).zip"; IsDir=$false},
        @{Name="445b Wings of Fury (E)(M3)[T-Polish_Arpi].zip"; IsDir=$false},
        @{Name="Harley-Davidson Motor Cycles - Race Across America (USA).zip"; IsDir=$false},
        @{Name="001_nintendo_games"; IsDir=$true},
        @{Name="992_-_action_games"; IsDir=$true},
        @{Name="mario-s_adventure"; IsDir=$true},
        # New samples
        @{Name="Pokémon_Edition_–_Special.tar.gz"; IsDir=$false},  # Unicode and multi-dot
        @{Name="Game <With> Invalid: Chars?.zip"; IsDir=$false},  # Invalid chars
        @{Name="2024EditionNoSeparator.zip"; IsDir=$false},       # No prefix separator
        @{Name="DLC_Content_VR_Mode (AMI).zip"; IsDir=$false},    # New abbreviations
        @{Name="ネスゲーム_日本語.zip"; IsDir=$false}               # Unicode Japanese
    )
    
    foreach ($testItem in $testItems) {
        $itemType = if ($testItem.IsDir) { "[FOLDER]" } else { "[FILE]  " }
        $cleaned = Clean-Filename -filename $testItem.Name -source "" -enhancements "" -isDirectory $testItem.IsDir -customReplacements $customReplacements
        $willChange = $testItem.Name -ne $cleaned
        
        if ($willChange) {
            Write-Host "$itemType WILL CHANGE: " -NoNewline -ForegroundColor Yellow
        } else {
            Write-Host "$itemType UNCHANGED:   " -NoNewline -ForegroundColor Green
        }
        Write-Host $testItem.Name
        if ($willChange) {
            Write-Host "                    → " -NoNewline -ForegroundColor Gray
            Write-Host $cleaned -ForegroundColor White
        }
        Write-Host ""
    }
}

# Load custom replacements from config if provided
$customReplacements = @{}
if ($ConfigFile -and (Test-Path $ConfigFile)) {
    $configContent = Get-Content $ConfigFile -Raw | ConvertFrom-Json
    $customReplacements = $configContent.Replacements | ConvertTo-Hashtable  # Assume JSON like {"Replacements": {"\\bFoo\\b": "FOO"}}
}

# Get source path if not provided
if (-not $SourcePath) {
    # Show pattern test first
    Test-CleaningPattern -customReplacements $customReplacements
    
    Write-Host "The script will:" -ForegroundColor Cyan
    Write-Host "• Rename both FILES and FOLDERS (recursively if -Recurse is used)" -ForegroundColor Yellow
    Write-Host "• Remove 3-4 digit organizational numbers (001, 992, 1111, etc.)" -ForegroundColor White
    Write-Host "• Handle patterns like '1111_-_tony_hawk-s_pro_skater_3'" -ForegroundColor White
    Write-Host "• Fix possessives without apostrophes (hawk-s → hawks)" -ForegroundColor White
    Write-Host "• Convert to proper Title Case with expanded gaming terms" -ForegroundColor White
    Write-Host "• Keep legitimate game numbers (10-Pin, 4x4, 3-D, Alpha 3, etc.)" -ForegroundColor White
    Write-Host "• Preserve hyphenated names (Harley-Davidson stays hyphenated)" -ForegroundColor White
    Write-Host "• Keep region codes uppercase (JE, USA, EUR, etc.)" -ForegroundColor White
    Write-Host "• Replace underscores with spaces" -ForegroundColor White
    Write-Host "• Add optional source/enhancement tags" -ForegroundColor White
    Write-Host "• Remove invalid filename characters" -ForegroundColor White
    Write-Host "• Optional: Backup originals, dry run, exclude patterns, custom config, logging" -ForegroundColor White
    Write-Host ""
    
    $SourcePath = Read-Host "Enter the source directory path"
}

# Verify path exists
if (-not (Test-Path $SourcePath)) {
    Write-Error "Path does not exist: $SourcePath"
    exit 1
}

# Set up logging
if (-not $LogFile) {
    $LogFile = Join-Path $SourcePath "rename_log.txt"
}
Write-Log -Message "Script started. SourcePath: $SourcePath, Recurse: $Recurse, Backup: $Backup, DryRun: $DryRun" -LogPath $LogFile

# Get optional metadata
$source = Read-Host "Enter source tag (optional, press Enter to skip)"
$enhancements = Read-Host "Enter enhancements tag (optional, press Enter to skip)"

# Get all files AND folders in the directory
$items = Get-ChildItem -Path $SourcePath -Recurse:$Recurse

if ($items.Count -eq 0) {
    Write-Warning "No files or folders found in the specified directory."
    Write-Log -Message "No items found." -Level "WARNING" -LogPath $LogFile
    exit 0
}

Write-Host "`nItems to be renamed:" -ForegroundColor Yellow
Write-Host "====================" -ForegroundColor Yellow

# Preview changes
$renameOperations = @()
foreach ($item in $items) {
    $isDirectory = $item.PSIsContainer
    $newName = Clean-Filename -filename $item.Name -source $source -enhancements $enhancements -isDirectory $isDirectory -customReplacements $customReplacements
    
    if ($newName -ne $item.Name) {
        $itemType = if ($isDirectory) { "FOLDER" } else { "FILE" }
        
        $renameOperations += @{
            OldName = $item.Name
            NewName = $newName
            OldPath = $item.FullName
            NewPath = Join-Path $item.Parent.FullName $newName
            ItemType = $itemType
        }
        
        Write-Host "[$itemType]" -ForegroundColor Cyan
        Write-Host "  OLD: " -NoNewline -ForegroundColor Red
        Write-Host $item.Name
        Write-Host "  NEW: " -NoNewline -ForegroundColor Green
        Write-Host $newName
        Write-Host ""
    }
}

if ($renameOperations.Count -eq 0) {
    Write-Host "No files or folders need renaming." -ForegroundColor Green
    Write-Log -Message "No renames needed." -LogPath $LogFile
    exit 0
}

# Count files and folders separately
$fileCount = ($renameOperations | Where-Object { $_.ItemType -eq "FILE" }).Count
$folderCount = ($renameOperations | Where-Object { $_.ItemType -eq "FOLDER" }).Count

Write-Host "Found items to rename:" -ForegroundColor Cyan
if ($fileCount -gt 0) {
    Write-Host "  • $fileCount file(s)" -ForegroundColor White
}
if ($folderCount -gt 0) {
    Write-Host "  • $folderCount folder(s)" -ForegroundColor White
}

if ($DryRun) {
    Write-Host "Dry run mode: No actual renames performed." -ForegroundColor Yellow
    Write-Log -Message "Dry run completed. No changes made." -LogPath $LogFile
    exit 0
}

$confirm = Read-Host "`nDo you want to proceed with renaming? (y/N)"

if ($confirm -match '^[Yy]') {
    Write-Host "`nRenaming items..." -ForegroundColor Yellow
    Write-Log -Message "Renaming started." -LogPath $LogFile
    
    $successCount = 0
    $errorCount = 0
    
    # Process folders first (in case there are nested items)
    $sortedOperations = $renameOperations | Sort-Object -Property @{Expression = {$_.ItemType}; Descending = $true}, OldPath -Descending  # Deeper paths first for recursion
    
    foreach ($operation in $sortedOperations) {
        $retryCount = 0
        $maxRetries = 3
        $success = $false
        
        while ($retryCount -lt $maxRetries -and -not $success) {
            try {
                # Refresh and check if source item still exists
                if (-not (Test-Path $operation.OldPath)) {
                    Write-Log -Message "Source $($operation.ItemType.ToLower()) no longer exists: $($operation.OldName)" -Level "WARNING" -LogPath $LogFile
                    $errorCount++
                    break
                }
                
                # Check if target name already exists
                if (Test-Path $operation.NewPath) {
                    Write-Log -Message "Target $($operation.ItemType.ToLower()) already exists: $($operation.NewName) (Conflict)" -Level "ERROR" -LogPath $LogFile
                    $errorCount++
                    break
                }
                
                # Backup if enabled
                if ($Backup) {
                    $backupPath = $operation.OldPath + ".bak"
                    Copy-Item -Path $operation.OldPath -Destination $backupPath -Recurse:$operation.ItemType -eq "FOLDER"
                    Write-Log -Message "Backed up: $($operation.OldPath) to $backupPath" -LogPath $LogFile
                }
                
                Rename-Item -Path $operation.OldPath -NewName $operation.NewName -ErrorAction Stop
                Write-Host "✓ " -NoNewline -ForegroundColor Green
                Write-Host "[$($operation.ItemType)] Renamed: $($operation.OldName) → $($operation.NewName)"
                Write-Log -Message "Renamed: $($operation.OldName) → $($operation.NewName)" -LogPath $LogFile
                $successCount++
                $success = $true
            }
            catch {
                $errorType = if ($_.Exception.Message -match "access|permission") { "Permissions" } 
                             elseif ($_.Exception.Message -match "exists") { "Conflict" } 
                             else { "Unknown" }
                Write-Log -Message "Failed to rename $($operation.ItemType.ToLower()) $($operation.OldName): $($_.Exception.Message) (Type: $errorType, Retry: $retryCount)" -Level "ERROR" -LogPath $LogFile
                $retryCount++
                if ($retryCount -ge $maxRetries) {
                    $errorCount++
                }
                Start-Sleep -Milliseconds 500  # Brief delay before retry
            }
        }
    }
    
    Write-Host "`nRenaming complete!" -ForegroundColor Green
    Write-Host "Successfully renamed: $successCount item(s)" -ForegroundColor Green
    if ($errorCount -gt 0) {
        Write-Host "Errors encountered: $errorCount item(s)" -ForegroundColor Red
    }
    Write-Log -Message "Renaming complete. Success: $successCount, Errors: $errorCount" -LogPath $LogFile
} else {
    Write-Host "Operation cancelled." -ForegroundColor Yellow
    Write-Log -Message "Operation cancelled by user." -LogPath $LogFile
}

# Pause to see results
Read-Host "`nPress Enter to exit"