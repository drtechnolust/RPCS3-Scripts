# CoinOPS Symlink Script - FIXED COMPLETE VERSION - ROMs and Videos Only
param(
    [string]$CoinOPSPath = "C:\Arcade\CoinOPS DeLuxe UNIVERSE 2025\collections",
    [string]$SystemRomsPath = "D:\Arcade\System roms",
    [string]$LaunchBoxVideosPath = "C:\Arcade\LaunchBox\Videos",
    [switch]$WhatIf = $false
)

# Check if running as Administrator
function Test-Administrator {
    $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($currentUser)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

# Function to check if a path is already a symlink
function Test-IsSymlink {
    param([string]$Path)
    
    if (-not (Test-Path $Path)) {
        return $false
    }
    
    $item = Get-Item $Path -Force
    return $item.LinkType -eq "SymbolicLink"
}

# Function to get symlink target
function Get-SymlinkTarget {
    param([string]$Path)
    
    if (Test-IsSymlink $Path) {
        $item = Get-Item $Path -Force
        return $item.Target
    }
    return $null
}

# Function to get user confirmation
function Get-UserConfirmation {
    param(
        [string]$Message,
        [string]$DefaultChoice = "n"
    )
    
    if ($DefaultChoice -eq "y") {
        $choices = "[Y/n]"
    } else {
        $choices = "[y/N]"
    }
    
    $response = Read-Host "$Message $choices"
    
    if ([string]::IsNullOrWhiteSpace($response)) {
        return $DefaultChoice -eq "y"
    }
    
    return $response.ToLower() -in @("y", "yes", "1", "true")
}

# FIXED: Platform selection function
function Get-SelectedPlatforms {
    param($Platforms)
    
    Write-Host ""
    Write-Host "=== PLATFORM SELECTION MENU ===" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "Available platforms:" -ForegroundColor Yellow
    
    # Show numbered list
    for ($i = 0; $i -lt $Platforms.Count; $i++) {
        $num = $i + 1
        Write-Host ("{0,2}. {1}" -f $num, $Platforms[$i].Name) -ForegroundColor Gray
    }
    
    Write-Host ""
    Write-Host "Selection options:" -ForegroundColor Yellow
    Write-Host "  - Single numbers: 1,3,5" -ForegroundColor Gray
    Write-Host "  - Ranges: 1-5,10-15" -ForegroundColor Gray
    Write-Host "  - Mixed: 1,3,7-12,20" -ForegroundColor Gray
    Write-Host "  - All: all" -ForegroundColor Gray
    Write-Host "  - Cancel: none or just press Enter" -ForegroundColor Gray
    
    do {
        $selection = Read-Host "`nEnter your selection"
        
        if ([string]::IsNullOrWhiteSpace($selection) -or $selection.ToLower() -eq "none") {
            return @()
        }
        
        if ($selection.ToLower() -eq "all") {
            Write-Host ""
            Write-Host "Selected: ALL $($Platforms.Count) platforms" -ForegroundColor Green
            $confirm = Get-UserConfirmation "Proceed with all platforms?" "n"
            if ($confirm) {
                return $Platforms
            } else {
                Write-Host "Please make a new selection." -ForegroundColor Yellow
                continue
            }
        }
        
        # Parse the selection
        $selectedPlatforms = @()
        $invalidSelection = $false
        $parts = $selection -split ","
        
        foreach ($part in $parts) {
            $part = $part.Trim()
            
            if ($part -match "^(\d+)-(\d+)$") {
                # Range like "5-10"
                $start = [int]$matches[1]
                $end = [int]$matches[2]
                
                for ($i = $start; $i -le $end; $i++) {
                    if ($i -ge 1 -and $i -le $Platforms.Count) {
                        $selectedPlatforms += $Platforms[$i - 1]
                    } else {
                        Write-Host "Invalid range: $i is not a valid platform number" -ForegroundColor Red
                        $invalidSelection = $true
                    }
                }
            }
            elseif ($part -match "^\d+$") {
                # Single number like "3"
                $num = [int]$part
                if ($num -ge 1 -and $num -le $Platforms.Count) {
                    $selectedPlatforms += $Platforms[$num - 1]
                } else {
                    Write-Host "Invalid selection: $num is not a valid platform number (1-$($Platforms.Count))" -ForegroundColor Red
                    $invalidSelection = $true
                }
            } else {
                Write-Host "Invalid format: '$part' - use numbers, ranges, or 'all'" -ForegroundColor Red
                $invalidSelection = $true
            }
        }
        
        if ($invalidSelection) {
            Write-Host "Please try again with valid selections." -ForegroundColor Yellow
            continue
        }
        
        # Remove duplicates
        $selectedPlatforms = $selectedPlatforms | Sort-Object Name -Unique
        
        if ($selectedPlatforms.Count -gt 0) {
            Write-Host ""
            Write-Host "Selected platforms:" -ForegroundColor Green
            foreach ($platform in $selectedPlatforms) {
                Write-Host "  - $($platform.Name)" -ForegroundColor Gray
            }
            
            $confirm = Get-UserConfirmation "`nProceed with these $($selectedPlatforms.Count) platforms?" "y"
            if ($confirm) {
                return $selectedPlatforms
            } else {
                Write-Host "Please make a new selection." -ForegroundColor Yellow
                continue
            }
        } else {
            Write-Host "No valid platforms selected. Try again." -ForegroundColor Red
        }
        
    } while ($true)
}

# Enhanced verification function
function Test-SymlinkSetup {
    param(
        [string]$PlatformName,
        [string]$PlatformPath
    )
    
    Write-Host ""
    Write-Host "=== ENHANCED VERIFICATION for $PlatformName ===" -ForegroundColor Yellow
    
    # Define paths
    $coinopsRoms = Join-Path $PlatformPath "roms"
    $coinopsVideo = Join-Path $PlatformPath "medium_artwork\video"
    
    $systemRoms = Join-Path $SystemRomsPath ($PlatformName + "\roms")
    $launchboxVideo = Join-Path $LaunchBoxVideosPath $PlatformName
    
    Write-Host ""
    Write-Host "PLATFORM MAPPING CHECK:" -ForegroundColor Cyan
    Write-Host "  CoinOPS Platform Name: $PlatformName" -ForegroundColor White
    Write-Host "  ROM Source: $systemRoms" -ForegroundColor Gray
    Write-Host "  Video Source: $launchboxVideo" -ForegroundColor Gray
    
    Write-Host ""
    Write-Host "CURRENT STATE ANALYSIS:" -ForegroundColor Cyan
    
    # Check ROMs
    Write-Host ""
    Write-Host "  ROMs:" -ForegroundColor White
    if (Test-IsSymlink $coinopsRoms) {
        $target = Get-SymlinkTarget $coinopsRoms
        Write-Host "    CoinOPS roms: ALREADY SYMLINKED -> $target" -ForegroundColor Blue
    } elseif (Test-Path $coinopsRoms) {
        Write-Host "    CoinOPS roms: EXISTS (real folder)" -ForegroundColor Green
    } else {
        Write-Host "    CoinOPS roms: MISSING" -ForegroundColor Red
    }
    
    Write-Host "    ROM source: $(if (Test-Path $systemRoms) {'EXISTS'} else {'MISSING - CHECK PATH!'})" -ForegroundColor $(if (Test-Path $systemRoms) {'Green'} else {'Red'})
    
    # Check Videos
    Write-Host ""
    Write-Host "  Videos:" -ForegroundColor White
    if (Test-IsSymlink $coinopsVideo) {
        $target = Get-SymlinkTarget $coinopsVideo
        Write-Host "    CoinOPS video: ALREADY SYMLINKED -> $target" -ForegroundColor Blue
    } elseif (Test-Path $coinopsVideo) {
        Write-Host "    CoinOPS video: EXISTS (real folder)" -ForegroundColor Green
    } else {
        Write-Host "    CoinOPS video: MISSING" -ForegroundColor Red
    }
    
    Write-Host "    Video source: $(if (Test-Path $launchboxVideo) {'EXISTS'} else {'MISSING - CHECK PATH!'})" -ForegroundColor $(if (Test-Path $launchboxVideo) {'Green'} else {'Red'})
    
    Write-Host ""
    Write-Host "PLANNED ACTIONS:" -ForegroundColor Cyan
    
    # ROM actions
    if (Test-IsSymlink $coinopsRoms) {
        Write-Host "  ROMs: SKIP - Already symlinked" -ForegroundColor Blue
    } elseif (Test-Path $systemRoms) {
        if (Test-Path $coinopsRoms) {
            Write-Host "  ROMs: Backup existing folder -> roms-backup-[timestamp]" -ForegroundColor Yellow
        }
        Write-Host "  ROMs: Create symlink $coinopsRoms -> $systemRoms" -ForegroundColor Green
    } else {
        Write-Host "  ROMs: SKIP - Source not found" -ForegroundColor Red
    }
    
    # Video actions  
    if (Test-IsSymlink $coinopsVideo) {
        Write-Host "  Videos: SKIP - Already symlinked" -ForegroundColor Blue
    } elseif (Test-Path $launchboxVideo) {
        if (Test-Path $coinopsVideo) {
            Write-Host "  Videos: Backup existing folder -> video-backup-[timestamp]" -ForegroundColor Yellow
        }
        Write-Host "  Videos: Create symlink $coinopsVideo -> $launchboxVideo" -ForegroundColor Green
    } else {
        Write-Host "  Videos: SKIP - Source not found" -ForegroundColor Red
    }
    
    return @{
        CoinOPSRoms = $coinopsRoms
        CoinOPSVideo = $coinopsVideo
        SystemRoms = $systemRoms
        LaunchBoxVideo = $launchboxVideo
        RomsIsSymlink = (Test-IsSymlink $coinopsRoms)
        VideoIsSymlink = (Test-IsSymlink $coinopsVideo)
        RomsSourceExists = (Test-Path $systemRoms)
        VideoSourceExists = (Test-Path $launchboxVideo)
    }
}

# Function to backup existing folder
function Backup-ExistingFolder {
    param([string]$FolderPath)
    
    if (Test-Path $FolderPath) {
        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $folderName = Split-Path $FolderPath -Leaf
        $parentPath = Split-Path $FolderPath -Parent
        $backupPath = Join-Path $parentPath ($folderName + "-backup-" + $timestamp)
        
        if ($WhatIf) {
            Write-Host "  WHATIF: Would backup '$folderName' to '$folderName-backup-$timestamp'" -ForegroundColor Yellow
            return $true
        }
        
        Write-Host "  Backing up existing folder: $folderName -> $folderName-backup-$timestamp" -ForegroundColor Yellow
        try {
            Rename-Item -Path $FolderPath -NewName $backupPath -ErrorAction Stop
            return $true
        } catch {
            Write-Host "  ERROR: Could not backup folder: $($_.Exception.Message)" -ForegroundColor Red
            return $false
        }
    }
    return $true
}

# Function to create symlink
function New-SafeSymlink {
    param(
        [string]$SymlinkLocation,    # WHERE to create the symlink (CoinOPS folder)
        [string]$PointsTo,          # WHAT it points to (external source)
        [string]$Description,
        [string]$Type,
        [bool]$IsAlreadySymlink,
        [bool]$SourceExists
    )
    
    Write-Host ""
    Write-Host "[$Type] $Description" -ForegroundColor Cyan
    
    if ($IsAlreadySymlink) {
        Write-Host "  SKIPPED: Already a symlink" -ForegroundColor Blue
        return "already_symlink"
    }
    
    if (-not $SourceExists) {
        Write-Host "  SKIPPED: Source '$PointsTo' not found" -ForegroundColor Yellow
        return "source_missing"
    }
    
    if ($WhatIf) {
        Write-Host "  WHATIF: mklink /d `"$SymlinkLocation`" `"$PointsTo`"" -ForegroundColor Green
        return "whatif"
    }
    
    # Backup existing CoinOPS folder (only if it's a real folder, not a symlink)
    if (-not (Backup-ExistingFolder -FolderPath $SymlinkLocation)) {
        return "backup_failed"
    }
    
    # Create symlink: CoinOPS location -> external source
    Write-Host "  Creating symlink: $SymlinkLocation -> $PointsTo" -ForegroundColor Green
    
    try {
        cmd /c "mklink /d `"$SymlinkLocation`" `"$PointsTo`""
        if ($LASTEXITCODE -eq 0) {
            Write-Host "  SUCCESS!" -ForegroundColor Green
            return "success"
        } else {
            Write-Host "  FAILED: Error code $LASTEXITCODE" -ForegroundColor Red
            return "failed"
        }
    } catch {
        Write-Host "  ERROR: $($_.Exception.Message)" -ForegroundColor Red
        return "error"
    }
}

# Main script
Clear-Host
Write-Host "=== CoinOPS Symlink Script - FIXED VERSION ===" -ForegroundColor Cyan
Write-Host "ROMs and Videos Only - Enhanced Verification with Platform Selection" -ForegroundColor Gray

# Check admin rights
if (-not (Test-Administrator)) {
    Write-Host ""
    Write-Host "ERROR: This script requires Administrator privileges!" -ForegroundColor Red
    Write-Host "Please right-click PowerShell and 'Run as Administrator'" -ForegroundColor Yellow
    Read-Host "Press Enter to exit"
    exit 1
}

Write-Host ""
Write-Host "Configuration:" -ForegroundColor Yellow
Write-Host "  CoinOPS Path: $CoinOPSPath" -ForegroundColor Gray
Write-Host "  System ROMs Path: $SystemRomsPath" -ForegroundColor Gray
Write-Host "  LaunchBox Videos: $LaunchBoxVideosPath" -ForegroundColor Gray

if ($WhatIf) {
    Write-Host ""
    Write-Host "WHATIF MODE - No changes will be made" -ForegroundColor Yellow
}

# Verify paths exist
if (-not (Test-Path $CoinOPSPath)) {
    Write-Host ""
    Write-Host "ERROR: CoinOPS path not found: $CoinOPSPath" -ForegroundColor Red
    Read-Host "Press Enter to exit"
    exit 1
}

# Get platforms and sort them
$platforms = Get-ChildItem -Path $CoinOPSPath -Directory | Sort-Object Name

if ($platforms.Count -eq 0) {
    Write-Host ""
    Write-Host "ERROR: No platform collections found in: $CoinOPSPath" -ForegroundColor Red
    Read-Host "Press Enter to exit"
    exit 1
}

Write-Host ""
Write-Host "Found $($platforms.Count) platform collections" -ForegroundColor Green

# FIXED: Force platform selection menu to appear
$selectedPlatforms = @()
do {
    $selectedPlatforms = Get-SelectedPlatforms -Platforms $platforms
    
    if ($selectedPlatforms.Count -eq 0) {
        Write-Host ""
        Write-Host "No platforms selected." -ForegroundColor Yellow
        $tryAgain = Get-UserConfirmation "Try platform selection again?" "y"
        if (-not $tryAgain) {
            Write-Host "Exiting script." -ForegroundColor Yellow
            exit 0
        }
    }
} while ($selectedPlatforms.Count -eq 0)

Write-Host ""
Write-Host "Starting verification and processing..." -ForegroundColor Green

# Process each selected platform with continue prompts
$totalSelected = $selectedPlatforms.Count
$currentIndex = 0

foreach ($platform in $selectedPlatforms) {
    $currentIndex++
    
    Write-Host ""
    Write-Host "=" * 80 -ForegroundColor Cyan
    Write-Host "PLATFORM $currentIndex of $totalSelected" -ForegroundColor Cyan
    Write-Host "=" * 80 -ForegroundColor Cyan
    
    $paths = Test-SymlinkSetup -PlatformName $platform.Name -PlatformPath $platform.FullName
    
    if (-not (Get-UserConfirmation "`nProceed with $($platform.Name)?")) {
        Write-Host "Skipping $($platform.Name)" -ForegroundColor Yellow
        
        # Ask if they want to continue to next platform after skipping
        if ($currentIndex -lt $totalSelected) {
            if (-not (Get-UserConfirmation "`nContinue to next platform? ($currentIndex/$totalSelected completed)")) {
                Write-Host "Script terminated by user." -ForegroundColor Yellow
                exit 0
            }
        }
        continue
    }
    
    Write-Host ""
    Write-Host "Processing $($platform.Name)... [$currentIndex of $totalSelected]" -ForegroundColor Green
    
    # Create symlinks - ROMs and Videos only
    $results = @()
    $results += New-SafeSymlink -SymlinkLocation $paths.CoinOPSRoms -PointsTo $paths.SystemRoms -Description "ROMs" -Type "ROM" -IsAlreadySymlink $paths.RomsIsSymlink -SourceExists $paths.RomsSourceExists
    $results += New-SafeSymlink -SymlinkLocation $paths.CoinOPSVideo -PointsTo $paths.LaunchBoxVideo -Description "Videos" -Type "VIDEO" -IsAlreadySymlink $paths.VideoIsSymlink -SourceExists $paths.VideoSourceExists
    
    $successCount = ($results | Where-Object { $_ -eq "success" }).Count
    $skippedCount = ($results | Where-Object { $_ -in @("already_symlink", "source_missing") }).Count
    $totalAttempts = ($results | Where-Object { $_ -ne $null }).Count
    
    Write-Host ""
    Write-Host "Results: $successCount created, $skippedCount skipped, $totalAttempts total" -ForegroundColor $(if ($successCount -gt 0) { "Green" } else { "Yellow" })
    
    # Ask to continue to next platform (except for the last one)
    if ($currentIndex -lt $totalSelected) {
        Write-Host ""
        Write-Host "Platform $currentIndex of $totalSelected completed: $($platform.Name)" -ForegroundColor Cyan
        $nextPlatform = $selectedPlatforms[$currentIndex].Name
        Write-Host "Next platform: $nextPlatform" -ForegroundColor Gray
        
        if (-not (Get-UserConfirmation "`nContinue to next platform?")) {
            Write-Host "Script terminated by user. Processed $currentIndex of $totalSelected platforms." -ForegroundColor Yellow
            exit 0
        }
        
        Write-Host ""
        Write-Host "Continuing to next platform..." -ForegroundColor Green
    }
}

Write-Host ""
Write-Host "================================================================================" -ForegroundColor Green
Write-Host "ALL PLATFORMS COMPLETE!" -ForegroundColor Green
Write-Host "Successfully processed $totalSelected platforms" -ForegroundColor Gray
Write-Host "================================================================================" -ForegroundColor Green

Read-Host "Press Enter to exit"