# clean_game_names.ps1
# Nintendo Switch ROM Filename Cleaner

function Clean-GameName {
    param([string]$Name)
    
    # Replace common separators with spaces
    $Name = $Name -replace '[_\-:]+', ' '
    
    # Handle numbers at start
    if ($Name -match '^\d+\s*') {
        $Name = $Name -replace '^(\d+)\s*(.*)$', '$1 $2'
    }
    
    # Remove version info that leaked through
    $Name = $Name -replace '\s+v\d+.*$', '' 
    $Name = $Name -replace '\s+(base|app|update).*$', ''
    
    # Clean up spacing
    $Name = ($Name -split '\s+' | Where-Object { $_ }) -join ' '
    
    # Title case while preserving all-caps words
    $words = $Name -split ' '
    $cleanedWords = foreach ($word in $words) {
        if ($word -cmatch '^[A-Z]+$' -and $word.Length -gt 2) {
            # Keep all-caps words
            $word
        } elseif ($word -cmatch '^[a-z]' -or ($word -cmatch '^[a-z]' -and $word.Length -gt 1)) {
            # Title case for lowercase words
            (Get-Culture).TextInfo.ToTitleCase($word.ToLower())
        } else {
            # Keep mixed case
            $word
        }
    }
    
    return $cleanedWords -join ' '
}

function Extract-RegionAndVersion {
    param([string]$Filename)
    
    # Remove file extensions
    $nameWithoutExt = $Filename -replace '\.(xci|nsp)$', ''
    
    # Extract region
    $regionMatch = [regex]::Match($nameWithoutExt, '__([A-Z]{2,3})_')
    $region = if ($regionMatch.Success) { $regionMatch.Groups[1].Value } else { $null }
    
    # Extract version
    $versionMatch = [regex]::Match($nameWithoutExt, '_v(\d+(?:\.\d+)?)_?')
    $version = if ($versionMatch.Success) { $versionMatch.Groups[1].Value } else { $null }
    
    return @($region, $version)
}

function Clean-Filename {
    param([string]$Filename)
    
    # Get file extension
    $extension = [System.IO.Path]::GetExtension($Filename)
    $nameWithoutExt = $Filename -replace '\.(xci|nsp)$', ''
    
    $region, $version = Extract-RegionAndVersion -Filename $Filename
    
    # Extract game name (everything before hex code)
    $match = [regex]::Match($nameWithoutExt, '^(.+?)_[0-9A-Fa-f]{12,}.*$')
    if ($match.Success) {
        $gameName = $match.Groups[1].Value
    } else {
        # Fallback
        $parts = $nameWithoutExt -split '_'
        $gameName = if ($parts.Count -gt 1) { $parts[0] } else { $nameWithoutExt }
    }
    
    # Clean the game name
    $cleanedName = Clean-GameName -Name $gameName
    
    # Build final name with region and version
    $finalName = $cleanedName
    if ($region) { $finalName += " [$region]" }
    if ($version) { $finalName += " [v$version]" }
    
    return @($finalName, $region, $version, $extension)
}

function Get-DirectoryPath {
    Write-Host "Nintendo Switch ROM Filename Cleaner" -ForegroundColor Green
    Write-Host "====================================" -ForegroundColor Green
    Write-Host ""
    
    while ($true) {
        Write-Host "Please enter the directory path containing your ROM files (.xci/.nsp):" -ForegroundColor Cyan
        Write-Host "Examples:" -ForegroundColor Gray
        Write-Host "  C:\Games\Nintendo Switch" -ForegroundColor Gray
        Write-Host "  D:\ROMs\Switch" -ForegroundColor Gray
        Write-Host "  . (for current directory)" -ForegroundColor Gray
        Write-Host ""
        
        $directoryPath = Read-Host "Directory path"
        
        # Handle empty input
        if ([string]::IsNullOrWhiteSpace($directoryPath)) {
            Write-Host "Please enter a valid directory path." -ForegroundColor Red
            continue
        }
        
        # Expand relative paths
        if ($directoryPath -eq ".") {
            $directoryPath = Get-Location
        }
        
        # Check if directory exists
        if (Test-Path -Path $directoryPath -PathType Container) {
            # Check for ROM files
            $xciFiles = Get-ChildItem -Path $directoryPath -Filter "*.xci"
            $nspFiles = Get-ChildItem -Path $directoryPath -Filter "*.nsp"
            $totalFiles = $xciFiles.Count + $nspFiles.Count
            
            if ($totalFiles -eq 0) {
                Write-Host "No .xci or .nsp files found in this directory. Please check the path." -ForegroundColor Red
                $allFiles = Get-ChildItem -Path $directoryPath | Select-Object -First 5
                if ($allFiles) {
                    Write-Host "Files found: $($allFiles.Name -join ', ')..." -ForegroundColor Gray
                }
                continue
            }
            
            Write-Host "✓ Found $($xciFiles.Count) .xci files and $($nspFiles.Count) .nsp files" -ForegroundColor Green
            return $directoryPath
        } else {
            Write-Host "Directory not found. Please check the path and try again." -ForegroundColor Red
        }
    }
}

function Get-RunMode {
    Write-Host ""
    Write-Host "Choose run mode:" -ForegroundColor Cyan
    Write-Host "1. Dry run (preview only - no files will be renamed)" -ForegroundColor Yellow
    Write-Host "2. Actual rename (files will be renamed)" -ForegroundColor Red
    Write-Host ""
    
    while ($true) {
        $choice = Read-Host "Enter your choice (1 or 2)"
        
        switch ($choice) {
            "1" { 
                Write-Host "✓ Dry run mode selected - no files will be changed" -ForegroundColor Green
                return $false 
            }
            "2" { 
                Write-Host "⚠️  Actual rename mode selected - files WILL be renamed!" -ForegroundColor Red
                Write-Host "Are you sure? (y/n)" -ForegroundColor Yellow
                $confirm = Read-Host
                if ($confirm -match "^[yY]") {
                    return $true
                } else {
                    Write-Host "Switching to dry run mode for safety" -ForegroundColor Green
                    return $false
                }
            }
            default { 
                Write-Host "Please enter 1 or 2" -ForegroundColor Red 
            }
        }
    }
}

function Process-GameFiles {
    param(
        [string]$DirectoryPath,
        [string]$OutputCsv,
        [bool]$RenameFiles
    )
    
    # Get both .xci and .nsp files
    $xciFiles = Get-ChildItem -Path $DirectoryPath -Filter "*.xci"
    $nspFiles = Get-ChildItem -Path $DirectoryPath -Filter "*.nsp"
    $files = $xciFiles + $nspFiles | Sort-Object Name
    
    $mappings = @()
    
    Write-Host ""
    Write-Host ("=" * 80) -ForegroundColor Gray
    if ($RenameFiles) {
        Write-Host "🔧 RENAMING FILES - Changes will be made!" -ForegroundColor Red
    } else {
        Write-Host "👁️  DRY RUN - Preview only, no changes will be made" -ForegroundColor Green
    }
    Write-Host ("=" * 80) -ForegroundColor Gray
    Write-Host "Processing $($files.Count) ROM files ($($xciFiles.Count) .xci, $($nspFiles.Count) .nsp)..." -ForegroundColor Cyan
    Write-Host ""
    
    foreach ($file in $files) {
        $originalName = $file.Name
        $cleanedName, $region, $version, $extension = Clean-Filename -Filename $originalName
        $newFilename = "$cleanedName$extension"
        
        $mapping = [PSCustomObject]@{
            OriginalFilename = $originalName
            CleanedGameName = $cleanedName
            NewFilename = $newFilename
            Region = if ($region) { $region } else { "Unknown" }
            Version = if ($version) { $version } else { "Unknown" }
            FileType = $extension.TrimStart('.')
        }
        
        $mappings += $mapping
        
        # Show the mapping
        if ($originalName -eq $newFilename) {
            Write-Host "✓ No change: $originalName" -ForegroundColor Gray
        } else {
            $fileTypeIcon = if ($extension -eq ".xci") { "🎮" } else { "📦" }
            Write-Host "$fileTypeIcon $originalName" -ForegroundColor Yellow
            Write-Host "   -> $newFilename" -ForegroundColor Cyan
            Write-Host "   Region: $($mapping.Region) | Version: $($mapping.Version) | Type: $($mapping.FileType.ToUpper())" -ForegroundColor Magenta
        }
        
        # Optionally rename files
        if ($RenameFiles) {
            $newPath = Join-Path $file.DirectoryName $newFilename
            if (-not (Test-Path $newPath)) {
                try {
                    Rename-Item -Path $file.FullName -NewName $newFilename -ErrorAction Stop
                    Write-Host "   ✅ Renamed successfully" -ForegroundColor Green
                } catch {
                    Write-Host "   ❌ Error renaming: $($_.Exception.Message)" -ForegroundColor Red
                }
            } else {
                Write-Host "   ⚠️  Target file already exists, skipping" -ForegroundColor Red
            }
        }
        
        Write-Host ""
    }
    
    # Export to CSV
    try {
        $mappings | Export-Csv -Path $OutputCsv -NoTypeInformation -Encoding UTF8
        Write-Host "📄 Mapping saved to: $OutputCsv" -ForegroundColor Green
    } catch {
        Write-Host "❌ Error saving CSV: $($_.Exception.Message)" -ForegroundColor Red
    }
    
    # Show statistics
    Write-Host ""
    Write-Host ("=" * 80) -ForegroundColor Gray
    Write-Host "📊 SUMMARY" -ForegroundColor Green
    Write-Host ("=" * 80) -ForegroundColor Gray
    Write-Host "Total files processed: $($mappings.Count)" -ForegroundColor White
    
    $changedFiles = $mappings | Where-Object { $_.OriginalFilename -ne $_.NewFilename }
    Write-Host "Files that would change: $($changedFiles.Count)" -ForegroundColor Yellow
    Write-Host "Files staying the same: $($mappings.Count - $changedFiles.Count)" -ForegroundColor Gray
    
    # Show file type distribution
    $fileTypeCounts = $mappings | Group-Object -Property FileType | Sort-Object Name
    Write-Host ""
    Write-Host "File type distribution:" -ForegroundColor Cyan
    foreach ($group in $fileTypeCounts) {
        $icon = if ($group.Name -eq "xci") { "🎮" } else { "📦" }
        Write-Host "  $icon $($group.Name.ToUpper()): $($group.Count) files" -ForegroundColor White
    }
    
    # Show region distribution
    $regionCounts = $mappings | Group-Object -Property Region | Sort-Object Name
    Write-Host ""
    Write-Host "Region distribution:" -ForegroundColor Cyan
    foreach ($group in $regionCounts) {
        Write-Host "  $($group.Name): $($group.Count) files" -ForegroundColor White
    }
    
    return $mappings
}

function Show-FilePatterns {
    param([string]$DirectoryPath)
    
    $xciFiles = Get-ChildItem -Path $DirectoryPath -Filter "*.xci" | Select-Object -First 10
    $nspFiles = Get-ChildItem -Path $DirectoryPath -Filter "*.nsp" | Select-Object -First 10
    $files = $xciFiles + $nspFiles
    
    $regions = @()
    $versions = @()
    
    Write-Host ""
    Write-Host "🔍 Analyzing filename patterns..." -ForegroundColor Green
    Write-Host ""
    
    foreach ($file in $files) {
        $region, $version = Extract-RegionAndVersion -Filename $file.Name
        if ($region) { $regions += $region }
        if ($version) { $versions += $version }
    }
    
    $uniqueRegions = $regions | Sort-Object -Unique
    $uniqueVersions = $versions | Sort-Object -Unique
    
    Write-Host "Found regions: $($uniqueRegions -join ', ')" -ForegroundColor Cyan
    Write-Host "Found versions: $($uniqueVersions -join ', ')" -ForegroundColor Cyan
    
    Write-Host ""
    Write-Host "Sample files:" -ForegroundColor Gray
    $files | Select-Object -First 5 | ForEach-Object {
        $icon = if ($_.Extension -eq ".xci") { "🎮" } else { "📦" }
        Write-Host "  $icon $($_.Name)" -ForegroundColor Gray
    }
}

# Main execution
Clear-Host

# Get directory path
$directoryPath = Get-DirectoryPath

# Show file patterns
Show-FilePatterns -DirectoryPath $directoryPath

# Get run mode
$renameFiles = Get-RunMode

# Generate output CSV filename
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$outputCsv = Join-Path $directoryPath "switch_rom_mapping_$timestamp.csv"

# Process files
$mappings = Process-GameFiles -DirectoryPath $directoryPath -OutputCsv $outputCsv -RenameFiles $renameFiles

# Final message with fixed pause
Write-Host ""
Write-Host ("=" * 80) -ForegroundColor Gray
if ($renameFiles) {
    Write-Host "✅ ROM file renaming completed!" -ForegroundColor Green
    Write-Host "Check the results above for any errors." -ForegroundColor Yellow
} else {
    Write-Host "👁️  Dry run completed!" -ForegroundColor Green
    Write-Host "To actually rename the files, run the script again and choose option 2." -ForegroundColor Yellow
}
Write-Host "📄 Full mapping saved to: $outputCsv" -ForegroundColor Cyan
Write-Host ""

# Fixed pause method - compatible with all PowerShell environments
try {
    if ($Host.UI.RawUI -and $Host.UI.RawUI.KeyAvailable -ne $null) {
        Write-Host "Press any key to exit..."
        $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    } else {
        throw "ReadKey not supported"
    }
} catch {
    # Fallback for environments that don't support ReadKey (ISE, VS Code, etc.)
    Write-Host "Press Enter to exit..."
    Read-Host | Out-Null
}