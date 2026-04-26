<#  Adventure Academia DLC Organizer - Handles Long Filenames  #>
##########
# CONFIG
##########
$SourceFolder = "D:\Arcade\System roms\Nintendo Switch\Organized\Archives-RAR\Unzipped\Adventure Academia The Fractured Continent _17DLC__US_ NSP"  # Update this path
$DestRoot = "D:\Arcade\System roms\Nintendo Switch\Organized"
$LogFile = Join-Path $DestRoot 'adventure_academia_dlc_organization.txt'

##########
# Setup Destination Folders
##########
$DlcDir = Join-Path $DestRoot 'Updates-DLC\Adventure Academia DLC'
$BackupDir = Join-Path $DestRoot 'Updates-DLC\Adventure Academia DLC\Original Names'

# Create folders
$AllDirs = @($DlcDir, $BackupDir)
foreach ($dir in $AllDirs) {
    if (-not (Test-Path -LiteralPath $dir)) { 
        New-Item -ItemType Directory -Path $dir -Force | Out-Null 
        Write-Host "Created folder: $dir" -ForegroundColor Green
    }
}

# Initialize log
"Adventure Academia DLC Organization - $(Get-Date)" | Out-File -LiteralPath $LogFile
"Source: $SourceFolder" | Out-File -LiteralPath $LogFile -Append
"Destination: $DlcDir" | Out-File -LiteralPath $LogFile -Append
"=" * 60 | Out-File -LiteralPath $LogFile -Append

##########
# FUNCTIONS
##########
function Get-DlcIdFromFilename {
    param([string]$filename)
    # Extract the DLC ID from the brackets
    if ($filename -match '\[DLC\]\[([A-F0-9]+)\]') {
        return $matches[1]
    }
    return "Unknown"
}

function Get-ShortDlcName {
    param([string]$filename, [int]$counter)
    
    $dlcId = Get-DlcIdFromFilename $filename
    
    # Create a short, manageable name
    $shortName = "Adventure_Academia_DLC_$($dlcId.Substring(-4)).nsp"  # Last 4 chars of ID
    
    return $shortName
}

##########
# MAIN PROCESSING
##########
Write-Host "`nProcessing Adventure Academia DLC files..." -ForegroundColor Yellow

try {
    # Get all NSP files in the folder
    $DlcFiles = Get-ChildItem -LiteralPath $SourceFolder -Filter "*.nsp" -ErrorAction SilentlyContinue
    
    Write-Host "Found $($DlcFiles.Count) DLC files to process" -ForegroundColor Cyan
    "Found $($DlcFiles.Count) DLC files to process" | Out-File -LiteralPath $LogFile -Append
    
    $ProcessedCount = 0
    $ErrorCount = 0
    
    foreach ($file in $DlcFiles) {
        try {
            # Extract DLC ID for unique naming
            $dlcId = Get-DlcIdFromFilename $file.Name
            
            # Create short filename
            $shortName = "Adventure_Academia_DLC_$($dlcId.Substring(-8)).nsp"  # Last 8 chars for uniqueness
            $destPath = Join-Path $DlcDir $shortName
            
            # Handle duplicates
            $counter = 1
            while (Test-Path -LiteralPath $destPath) {
                $nameWithoutExt = [System.IO.Path]::GetFileNameWithoutExtension($shortName)
                $destPath = Join-Path $DlcDir "$nameWithoutExt-$counter.nsp"
                $counter++
            }
            
            # Create a mapping file for reference
            $mappingEntry = "Short: $([System.IO.Path]::GetFileName($destPath)) | Original: $($file.Name) | DLC ID: $dlcId"
            $mappingEntry | Out-File -LiteralPath $LogFile -Append
            
            # Move the file
            Move-Item -LiteralPath $file.FullName -Destination $destPath -Force
            
            Write-Host "MOVED: $($file.Name.Substring(0, [Math]::Min(50, $file.Name.Length)))... -> $([System.IO.Path]::GetFileName($destPath))" -ForegroundColor Green
            $ProcessedCount++
        }
        catch {
            $ErrorCount++
            $errorMsg = "ERROR processing $($file.Name): $($_.Exception.Message)"
            Write-Host $errorMsg -ForegroundColor Red
            $errorMsg | Out-File -LiteralPath $LogFile -Append
        }
    }
    
    # Summary
    Write-Host "`n" + "="*50 -ForegroundColor Magenta
    Write-Host "ADVENTURE ACADEMIA DLC ORGANIZATION COMPLETE!" -ForegroundColor Magenta
    Write-Host "="*50 -ForegroundColor Magenta
    
    $summary = @"

SUMMARY:
- Files Processed: $ProcessedCount
- Errors: $ErrorCount
- Destination: $DlcDir

All DLC files have been renamed with shorter, manageable names.
Check the log file for original name mappings: $LogFile

"@
    
    Write-Host $summary -ForegroundColor Cyan
    $summary | Out-File -LiteralPath $LogFile -Append
    
}
catch {
    Write-Host "Error accessing source folder: $($_.Exception.Message)" -ForegroundColor Red
    "Error accessing source folder: $($_.Exception.Message)" | Out-File -LiteralPath $LogFile -Append
}

Write-Host "`nDone! Check your organized DLC folder." -ForegroundColor Green