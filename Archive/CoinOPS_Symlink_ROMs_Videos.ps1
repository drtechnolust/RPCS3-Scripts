# CoinOPS Symlink Script - Enhanced Verification - ROMs and Videos Only
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

# Function to parse platform selection
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
    Write-Host "  - None/Cancel: none or just press Enter" -ForegroundColor Gray
    
    $selection = Read-Host "`nEnter your selection"
    
    if ([string]::IsNullOrWhiteSpace($selection) -or $selection.ToLower() -eq "none") {
        return @()
    }
    
    if ($selection.ToLower() -eq "all") {
        return $Platforms
    }
    
    # Parse the selection
    $selectedPlatforms = @()
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
                }
            }
        }
        elseif ($part -match "^\d+$") {
            # Single number like "3"
            $num = [int]$part
            if ($num -ge 1 -and $num -le $Platforms.Count) {
                $selectedPlatforms += $Platforms[$num - 1]
            }
        }
    }
    
    # Remove duplicates
    $selectedPlatforms = $selectedPlatforms | Sort-Object Name -Unique
    
    if ($selectedPlatforms.Count -gt 0) {
        Write-Host ""
        Write-Host "Selected platforms:" -ForegroundColor Green
        foreach ($platform in $selectedPlatforms) {
            Write-Host "  - $($platform.Name)" -ForegroundColor Gray
        }
        
        $confirm = Get-UserConfirmation "`nProceed with these platforms?" "y"
        if ($confirm) {
            return $selectedPlatforms
        }
    }
    
    return @()
}

# Main script
Clear-Host
Write-Host "=== CoinOPS Symlink Script - Enhanced Verification ===" -ForegroundColor Cyan
Write-Host "ROMs and Videos Only - with Smart Detection" -ForegroundColor Gray

# Check admin rights
if (-not (Test-Administrator)) {
    Write-Host ""
    Write-Host "ERROR: This script requires Administrator privileges!" -ForegroundColor Red
    Write-Host "Please right-click PowerShell and 'Run as Administrator'" -ForegroundColor Yellow
    exit 1
}

if ($WhatIf) {
    Write-Host ""
    Write-Host "WHATIF MODE - No changes will be made" -ForegroundColor Yellow
}

# Verify paths exist
if (-not (Test-Path $CoinOPSPath)) {
    Write-Host ""
    Write-Host "ERROR: CoinOPS path not found: $CoinOPSPath" -ForegroundColor Red
    exit 1
}

# Get platforms
$platforms = Get-ChildItem -Path $CoinOPSPath -Directory

Write-Host ""
Write-Host "Found $($platforms.Count) platform collections" -ForegroundColor Green

# Get selected platforms
$selectedPlatforms = Get-SelectedPlatforms -Platforms $platforms

if ($selectedPlatforms.Count -eq 0) {
    Write-Host ""
    Write-Host "No platforms selected. Exiting." -ForegroundColor Yellow
    exit 0
}

# Process each selected platform
foreach ($platform in $selectedPlatforms) {
    $paths = Test-SymlinkSetup -PlatformName $platform.Name -PlatformPath $platform.FullName
    
    if (-not (Get-UserConfirmation "`nProceed with $($platform.Name)?")) {
        Write-Host "Skipping $($platform.Name)" -ForegroundColor Yellow
        continue
    }
    
    Write-Host ""
    Write-Host "Processing $($platform.Name)..." -ForegroundColor Green
    
    # Create symlinks - ROMs and Videos only
    $results = @()
    $results += New-SafeSymlink -SymlinkLocation $paths.CoinOPSRoms -PointsTo $paths.SystemRoms -Description "ROMs" -Type "ROM" -IsAlreadySymlink $paths.RomsIsSymlink -SourceExists $paths.RomsSourceExists
    $results += New-SafeSymlink -SymlinkLocation $paths.CoinOPSVideo -PointsTo $paths.LaunchBoxVideo -Description "Videos" -Type "VIDEO" -IsAlreadySymlink $paths.VideoIsSymlink -SourceExists $paths.VideoSourceExists
    
    $successCount = ($results | Where-Object { $_ -eq "success" }).Count
    $skippedCount = ($results | Where-Object { $_ -in @("already_symlink", "source_missing") }).Count
    $totalAttempts = ($results | Where-Object { $_ -ne $null }).Count
    
    Write-Host ""
    Write-Host "Results: $successCount created, $skippedCount skipped, $totalAttempts total" -ForegroundColor $(if ($successCount -gt 0) { "Green" } else { "Yellow" })
}

Write-Host ""
Write-Host "COMPLETE!" -ForegroundColor Green