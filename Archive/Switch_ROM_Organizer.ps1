<#  Switch ROM Organizer -- v2.0 for Nested Folder Structure  #>
##########
# CONFIG – EDIT ME
##########
$SourceRoot = 'D:\Arcade\System roms\Nintendo Switch\New folder'  # Your current messy folder
$DestRoot   = 'D:\Arcade\System roms\Nintendo Switch\New folder\Organized'                               # Will be created if missing
$LogFile    = Join-Path $DestRoot 'organization_log.txt'

##########
# Derived destinations – do NOT edit unless you want different names
##########
$XciDir      = Join-Path $DestRoot 'XCI'
$NspDir      = Join-Path $DestRoot 'NSP'
$EshopDir    = Join-Path $DestRoot 'NSP-eShop'
$RarDir      = Join-Path $DestRoot 'Archives-RAR'
$UpdatesDir  = Join-Path $DestRoot 'Updates-DLC'
$EmptyDir    = Join-Path $DestRoot '_Empty-Folders-Log'

# Create destination folders if they do not yet exist
$AllDirs = @($DestRoot, $XciDir, $NspDir, $EshopDir, $RarDir, $UpdatesDir, $EmptyDir)
foreach ($dir in $AllDirs) {
    if (-not (Test-Path $dir)) { 
        New-Item -ItemType Directory -Path $dir | Out-Null 
        Write-Host "Created folder: $dir" -ForegroundColor Green
    }
}

# Initialize log file
"Switch ROM Organization Log - $(Get-Date)" | Out-File $LogFile
"Source: $SourceRoot" | Out-File $LogFile -Append
"Destination: $DestRoot" | Out-File $LogFile -Append
"=" * 50 | Out-File $LogFile -Append

##########
# Functions
##########
function Get-SafeFileName {
    param([string]$fileName)
    # Remove invalid characters and limit length
    $safe = $fileName -replace '[<>:"/\\|?*]', '_'
    if ($safe.Length -gt 100) { $safe = $safe.Substring(0, 100) }
    return $safe
}

function Test-IsUpdate {
    param([string]$fileName)
    # Check if file appears to be an update or DLC
    return ($fileName -match '\[v\d+\.\d+\.\d+\]' -or 
            $fileName -match '\(Update\)' -or 
            $fileName -match '\(DLC\)' -or
            $fileName -match 'Update' -or
            $fileName -match 'DLC')
}

##########
# Phase 1: Scan for empty folders
##########
Write-Host "`nPhase 1: Scanning for empty folders..." -ForegroundColor Yellow
$EmptyFolders = @()
$AllFolders = Get-ChildItem -Path $SourceRoot -Recurse -Directory

foreach ($folder in $AllFolders) {
    $contents = Get-ChildItem -Path $folder.FullName -Force
    if ($contents.Count -eq 0) {
        $EmptyFolders += $folder.FullName
        "$($folder.FullName)" | Out-File $LogFile -Append
    }
}

Write-Host "Found $($EmptyFolders.Count) empty folders" -ForegroundColor Cyan
"Found $($EmptyFolders.Count) empty folders:" | Out-File $LogFile -Append
$EmptyFolders | Out-File $LogFile -Append

##########
# Phase 2: Scan and organize ROM files
##########
Write-Host "`nPhase 2: Scanning for ROM files..." -ForegroundColor Yellow

$FileCount = @{
    XCI = 0
    NSP = 0
    EshopNSP = 0
    Updates = 0
    RAR = 0
    Errors = 0
}

# Get all ROM files recursively
$AllFiles = Get-ChildItem -Path $SourceRoot -Recurse -File -Include *.xci,*.nsp,*.rar

Write-Host "Found $($AllFiles.Count) total files to process" -ForegroundColor Cyan
"Found $($AllFiles.Count) total files to process" | Out-File $LogFile -Append

foreach ($file in $AllFiles) {
    try {
        $nameLower = $file.Name.ToLower()
        $parentLower = $file.Directory.Name.ToLower()
        $targetDir = $null
        $category = ""
        
        # Determine destination based on file type and naming patterns
        switch ($file.Extension.ToLower()) {
            '.xci' { 
                if (Test-IsUpdate $file.Name) {
                    $targetDir = $UpdatesDir
                    $category = "Update/DLC XCI"
                    $FileCount.Updates++
                } else {
                    $targetDir = $XciDir
                    $category = "XCI"
                    $FileCount.XCI++
                }
                break 
            }
            '.nsp' { 
                if (Test-IsUpdate $file.Name) {
                    $targetDir = $UpdatesDir
                    $category = "Update/DLC NSP"
                    $FileCount.Updates++
                } elseif (($nameLower + $parentLower) -match 'eshop') {
                    $targetDir = $EshopDir
                    $category = "eShop NSP"
                    $FileCount.EshopNSP++
                } else {
                    $targetDir = $NspDir
                    $category = "NSP"
                    $FileCount.NSP++
                }
                break 
            }
            '.rar' { 
                $targetDir = $RarDir
                $category = "RAR Archive"
                $FileCount.RAR++
                break 
            }
        }
        
        if ($null -ne $targetDir) {
            $safeName = Get-SafeFileName $file.Name
            $destPath = Join-Path $targetDir $safeName
            
            # Handle duplicates by adding number suffix
            $counter = 1
            $originalDestPath = $destPath
            while (Test-Path $destPath) {
                $nameWithoutExt = [System.IO.Path]::GetFileNameWithoutExtension($safeName)
                $ext = [System.IO.Path]::GetExtension($safeName)
                $destPath = Join-Path $targetDir "$nameWithoutExt-$counter$ext"
                $counter++
            }
            
            # Log the move
            $logEntry = "[$category] $($file.FullName) -> $destPath"
            $logEntry | Out-File $LogFile -Append
            
            # Move the file (remove -WhatIf when ready)
            Move-Item -LiteralPath $file.FullName -Destination $destPath -Force -WhatIf
            
            Write-Host "[$category] $($file.Name)" -ForegroundColor Green
        }
    }
    catch {
        $FileCount.Errors++
        $errorEntry = "ERROR processing $($file.FullName): $($_.Exception.Message)"
        Write-Host $errorEntry -ForegroundColor Red
        $errorEntry | Out-File $LogFile -Append
    }
}

##########
# Phase 3: Summary Report
##########
Write-Host "`n" + "="*60 -ForegroundColor Magenta
Write-Host "ORGANIZATION COMPLETE!" -ForegroundColor Magenta
Write-Host "="*60 -ForegroundColor Magenta

$summary = @"

SUMMARY REPORT:
- XCI Files: $($FileCount.XCI)
- NSP Files: $($FileCount.NSP)
- eShop NSP Files: $($FileCount.EshopNSP)
- Updates/DLC: $($FileCount.Updates)
- RAR Archives: $($FileCount.RAR)
- Empty Folders Found: $($EmptyFolders.Count)
- Errors: $($FileCount.Errors)

NEXT STEPS:
1. Remove -WhatIf from script to actually move files
2. Review log file: $LogFile
3. Consider deleting empty source folders after verification
4. Extract RAR files if needed

"@

Write-Host $summary -ForegroundColor Cyan
$summary | Out-File $LogFile -Append

##########
# Phase 4: Optional - Clean up empty folders (COMMENTED OUT FOR SAFETY)
##########
<#
Write-Host "`nPhase 4: Cleaning up empty folders..." -ForegroundColor Yellow
foreach ($emptyFolder in $EmptyFolders) {
    if (Test-Path $emptyFolder) {
        Remove-Item -Path $emptyFolder -Force -WhatIf
        Write-Host "Removed empty folder: $emptyFolder" -ForegroundColor DarkGray
    }
}
#>

Write-Host "`nLog file created at: $LogFile" -ForegroundColor Yellow
Write-Host "Remove -WhatIf parameters when ready to execute moves!" -ForegroundColor Red