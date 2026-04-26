<#
.SYNOPSIS
    CoinOps Launch Config Updater - RetroArchXiso Edition (Fixed Paths)

.DESCRIPTION
    Updates launcher.windows.conf files for your CoinOps systems in the collections folder

.PARAMETER CoinOpsPath
    Path to the CoinOps root directory

.PARAMETER CreateBackup
    Create backup files before making changes

.PARAMETER WhatIf
    Preview changes without actually making them

.EXAMPLE
    .\Update-YourCoinOpsConfigs.ps1 -CoinOpsPath "C:\Arcade\CoinOPS Deluxe UNIVERSE 2025" -CreateBackup
#>

param(
    [Parameter(Mandatory=$false)]
    [string]$CoinOpsPath,
    
    [Parameter(Mandatory=$false)]
    [switch]$CreateBackup,
    
    [Parameter(Mandatory=$false)]
    [switch]$WhatIf,
    
    [Parameter(Mandatory=$false)]
    [switch]$ShowMappings,
    
    [Parameter(Mandatory=$false)]
    [string]$RetroArchPath = "emulators\RetroArchXiso\retroarch.exe",
    
    [Parameter(Mandatory=$false)]
    [string]$RetroArchWorkingDir = "emulators\RetroArchXiso",
    
    [Parameter(Mandatory=$false)]
    [string]$CoresPath = "cores"
)

# Error handling
$ErrorActionPreference = "Stop"

# Add CTRL+C handler
$null = Register-EngineEvent -SourceIdentifier PowerShell.Exiting -Action {
    Write-Host "`n🛑 Script interrupted by user" -ForegroundColor Yellow
    Write-Host "Process was stopped safely." -ForegroundColor Green
}

# Systems to EXCLUDE (you handle these with different emulators)
$ExcludedSystems = @(
    # Digital Distribution
    'gog.com',
    'pc gamer is pc games',
    'collector',
    
    # Sony PlayStation (ALL)
    'playstation 1',
    'sony playstation',
    'sony pocketstation',
    'playstation 2',
    'sony playstation 2',
    'playstation 3',
    'sony playstation 3',
    'sony playstation 3 ware',
    'playstation 4',
    'playstation portable',
    'sony psp',
    'sony psp japan',
    'sony playstation minis',
    
    # Microsoft Xbox (ALL)
    'microsoft xbox',
    'xbox',
    'microsoft xbox 360',
    'xbox 360',
    'xbox 360 - copy',
    'xbox live arcade',
    
    # Dreamcast (ALL)
    'dreamcast',
    'sega dreamcast',
    'sega dreamcast import',
    'dreamcast aw ports',
    'dreamcast redream sub',
    
    # Nintendo Wii/WiiU (ALL)
    'nintendo wii',
    'wii',
    'nintendo wiiware',
    'nintendo wii u',
    'wii u',
    'nintendo wii u ware'
)

# Your ALLOWED System-to-Core Mappings (excluding the above systems)
$YourSystemCoreMap = @{
    # 3D Systems
    '3do' = 'opera_libretro.dll'
    'panasonic 3do' = 'opera_libretro.dll'
    
    # Action/Arcade Systems
    'aae' = 'mame_libretro.dll'
    'american laser games' = 'mame_libretro.dll'
    'gaelco' = 'mame_libretro.dll'
    'nesicaxlive' = 'mame_libretro.dll'
    'sega ringedge' = 'flycast_libretro.dll'
    'sega triforce' = 'flycast_libretro.dll'
    'triforce' = 'flycast_libretro.dll'
    'taito type x' = 'mame_libretro.dll'
    'taito type x3' = 'mame_libretro.dll'
    'zinc' = 'mame_libretro.dll'
    
    # Computer Systems - Amiga
    'amiga official' = 'puae_libretro.dll'
    'commodore amiga' = 'puae_libretro.dll'
    'commodore amiga cd32' = 'puae_libretro.dll'
    'commodore cdtv' = 'puae_libretro.dll'
    
    # Computer Systems - Other
    'amstrad cpc' = 'cap32_libretro.dll'
    'commodore 64' = 'vice_x64_libretro.dll'
    'microsoft msx' = 'bluemsx_libretro.dll'
    'microsoft msx2' = 'bluemsx_libretro.dll'
    'microsoft msx2+' = 'bluemsx_libretro.dll'
    'msx1' = 'bluemsx_libretro.dll'
    'msx2' = 'bluemsx_libretro.dll'
    'sinclair zx spectrum' = 'fuse_libretro.dll'
    'zx spectrum' = 'fuse_libretro.dll'
    'atari 800xl' = 'atari800_libretro.dll'
    
    # Atari Systems
    'atari 2600' = 'stella_libretro.dll'
    'atari 5200' = 'atari800_libretro.dll'
    'atari 7800' = 'prosystem_libretro.dll'
    'atari jaguar' = 'virtualjaguar_libretro.dll'
    'atari lynx' = 'handy_libretro.dll'
    
    # Bandai Systems
    'bandai sufami turbo' = 'snes9x_libretro.dll'
    'bandai wonderswan' = 'mednafen_wswan_libretro.dll'
    'bandai wonderswan color' = 'mednafen_wswan_libretro.dll'
    
    # Classic Systems
    'colecovision' = 'bluemsx_libretro.dll'
    'intellivision' = 'freeintv_libretro.dll'
    
    # Game & Watch
    'game and watch' = 'gw_libretro.dll'
    'nintendo game and watch' = 'gw_libretro.dll'
    
    # Nintendo Handheld
    'game boy advance' = 'mgba_libretro.dll'
    'nintendo game boy advance' = 'mgba_libretro.dll'
    'nintendo game boy' = 'gambatte_libretro.dll'
    'nintendo game boy color' = 'gambatte_libretro.dll'
    'nintendo pokemon mini' = 'pokemini_libretro.dll'
    'nintendo virtual boy' = 'mednafen_vb_libretro.dll'
    
    # Nintendo Home Consoles
    'nintendo entertainment system' = 'nestopia_libretro.dll'
    'nintendo entertainment system hacks' = 'nestopia_libretro.dll'
    'nintendo famicom' = 'nestopia_libretro.dll'
    'nintendo famicom disk system' = 'nestopia_libretro.dll'
    'nintendo famicom translated' = 'nestopia_libretro.dll'
    'nes hd' = 'nestopia_libretro.dll'
    
    'super nintendo' = 'snes9x_libretro.dll'
    'super nintendo entertainment system' = 'snes9x_libretro.dll'
    'super nintendo entertainment system hacks' = 'snes9x_libretro.dll'
    'super nintendo entertainment system translated' = 'snes9x_libretro.dll'
    'nintendo super famicom' = 'snes9x_libretro.dll'
    'nintendo super game boy' = 'snes9x_libretro.dll'
    'nintendo satellaview' = 'snes9x_libretro.dll'
    'nintendo sufami turbo' = 'snes9x_libretro.dll'
    'snes cd' = 'snes9x_libretro.dll'
    
    'nintendo 64' = 'mupen64plus_next_libretro.dll'
    'nintendo 64 dd' = 'mupen64plus_next_libretro.dll'
    
    # Nintendo Modern (GameCube ONLY - Wii/WiiU excluded)
    'gamecube' = 'dolphin_libretro.dll'
    'nintendo gamecube' = 'dolphin_libretro.dll'
    'gamecube dolphin sub' = 'dolphin_libretro.dll'
    
    # Nintendo Portables (DS/3DS/Switch)
    'nintendo ds' = 'desmume_libretro.dll'
    'nintendo ds japan' = 'desmume_libretro.dll'
    'nintendo 3ds' = 'citra_libretro.dll'
    'nintendo switch' = 'yuzu_libretro.dll'
    
    # NEC Systems
    'nec pc engine' = 'mednafen_pce_fast_libretro.dll'
    'nec pc engine-cd' = 'mednafen_pce_fast_libretro.dll'
    'nec pc-fx' = 'mednafen_pcfx_libretro.dll'
    'nec supergrafx' = 'mednafen_supergrafx_libretro.dll'
    'nec turbografx-16' = 'mednafen_pce_fast_libretro.dll'
    'nec turbografx-cd' = 'mednafen_pce_fast_libretro.dll'
    'pc engine' = 'mednafen_pce_fast_libretro.dll'
    
    # Sega Systems (Dreamcast excluded)
    'genesis' = 'genesis_plus_gx_libretro.dll'
    'sega genesis' = 'genesis_plus_gx_libretro.dll'
    'sega mega drive' = 'genesis_plus_gx_libretro.dll'
    
    'master system' = 'genesis_plus_gx_libretro.dll'
    'sega master system' = 'genesis_plus_gx_libretro.dll'
    'sega master system imports' = 'genesis_plus_gx_libretro.dll'
    
    'sega game gear' = 'genesis_plus_gx_libretro.dll'
    
    'sega 32x' = 'picodrive_libretro.dll'
    
    'sega cd' = 'genesis_plus_gx_libretro.dll'
    'sega mega-cd' = 'genesis_plus_gx_libretro.dll'
    
    'saturn' = 'mednafen_saturn_libretro.dll'
    'sega saturn' = 'mednafen_saturn_libretro.dll'
    'sega saturn japan' = 'mednafen_saturn_libretro.dll'
    'sega ages' = 'mednafen_saturn_libretro.dll'
    
    'sega vmu' = 'vemulator_libretro.dll'
    
    # SNK Systems
    'snk neo geo cd' = 'fbneo_libretro.dll'
    'snk neo geo pocket' = 'mednafen_ngp_libretro.dll'
    'snk neo geo pocket color' = 'mednafen_ngp_libretro.dll'
    'ngage' = 'eka2l1_libretro.dll'
    
    # Specialty Systems
    'scummvm' = 'scummvm_libretro.dll'
    'openbor' = 'openbor_libretro.dll'
    'mugen' = 'mugen_libretro.dll'
    'doom' = 'prboom_libretro.dll'
    'wow action max' = 'mame_libretro.dll'
    'dice' = 'mame_libretro.dll'
}

# Statistics tracking
$Global:Stats = @{
    TotalFound = 0
    Updated = 0
    Skipped = 0
    ExcludedByUser = 0
    Errors = 0
    BackupsCreated = 0
}

$Global:ProcessedSystems = @{}
$Global:ExcludedSystems = @{}
$Global:ErrorLog = @()

function Write-ColorOutput {
    param(
        [string]$Message,
        [ConsoleColor]$ForegroundColor = [ConsoleColor]::White,
        [switch]$NoNewline
    )
    
    $currentFG = $Host.UI.RawUI.ForegroundColor
    $Host.UI.RawUI.ForegroundColor = $ForegroundColor
    
    if ($NoNewline) {
        Write-Host $Message -NoNewline
    } else {
        Write-Host $Message
    }
    
    $Host.UI.RawUI.ForegroundColor = $currentFG
}

function Show-Header {
    Clear-Host
    Write-ColorOutput "╔══════════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
    Write-ColorOutput "║     CoinOps Config Updater - RetroArchXiso Edition          ║" -ForegroundColor Cyan
    Write-ColorOutput "║          (Fixed for Collections/launcher.windows.conf)      ║" -ForegroundColor Cyan
    Write-ColorOutput "╚══════════════════════════════════════════════════════════════╝" -ForegroundColor Cyan
    Write-Host ""
    Write-ColorOutput "🎯 RetroArch Path: $RetroArchPath" -ForegroundColor Green
    Write-ColorOutput "📁 Working Directory: $RetroArchWorkingDir" -ForegroundColor Green
    Write-ColorOutput "📄 Config File: launcher.windows.conf" -ForegroundColor Green
    Write-Host ""
}

function Test-SystemExcluded {
    param([string]$SystemName)
    
    $normalizedSystemName = $SystemName.ToLower().Trim()
    return $ExcludedSystems -contains $normalizedSystemName
}

function Get-SystemConfigFiles {
    param([string]$Path)
    
    Write-ColorOutput "🔍 Scanning collections folder for launcher.windows.conf files..." -ForegroundColor Yellow
    
    # The collections folder path
    $collectionsPath = Join-Path $Path "collections"
    
    if (-not (Test-Path $collectionsPath)) {
        Write-ColorOutput "❌ Collections folder not found at: $collectionsPath" -ForegroundColor Red
        return @()
    }
    
    # Folders to exclude from scanning
    $ExcludeFolders = @(
        '_common', 
        '_My Collections', 
        'Collector', 
        'zzzShutdown'
    )
    
    try {
        $scanStart = Get-Date
        $configFiles = @()
        
        # Get all folders in collections directory
        $systemFolders = Get-ChildItem -Path $collectionsPath -Directory -ErrorAction Stop
        $totalFolders = $systemFolders.Count
        
        Write-ColorOutput "📁 Found $totalFolders potential system folders in collections" -ForegroundColor Cyan
        
        $processedFolders = 0
        foreach ($folder in $systemFolders) {
            $processedFolders++
            
            # Skip excluded folders
            if ($ExcludeFolders -contains $folder.Name) {
                Write-ColorOutput "   ⏭️  [$processedFolders/$totalFolders] Skipping: $($folder.Name)" -ForegroundColor Yellow
                continue
            }
            
            # Look for launcher.windows.conf in this system folder
            $configPath = Join-Path $folder.FullName "launcher.windows.conf"
            if (Test-Path $configPath) {
                $configFiles += Get-Item $configPath
                Write-ColorOutput "   ✅ [$processedFolders/$totalFolders] Found config: $($folder.Name)" -ForegroundColor Green
            } else {
                Write-ColorOutput "   ⚠️  [$processedFolders/$totalFolders] No config: $($folder.Name)" -ForegroundColor DarkGray
            }
            
            # Show progress every 25 folders
            if ($totalFolders -gt 50 -and $processedFolders % 25 -eq 0) {
                $percent = [math]::Round(($processedFolders / $totalFolders) * 100, 1)
                Write-ColorOutput "   📊 Folder scan progress: $percent%" -ForegroundColor Cyan
            }
        }
        
        $scanTime = (Get-Date) - $scanStart
        Write-ColorOutput "🎯 Scan completed in $([math]::Round($scanTime.TotalSeconds, 2)) seconds" -ForegroundColor Green
        Write-ColorOutput "📄 Found $($configFiles.Count) launcher.windows.conf files" -ForegroundColor Green
        
        return $configFiles
        
    }
    catch {
        Write-ColorOutput "❌ ERROR during folder scan: $($_.Exception.Message)" -ForegroundColor Red
        throw
    }
}

function Backup-ConfigFiles {
    param([string]$Path)
    
    $backupStartTime = Get-Date
    Write-ColorOutput "💾 Getting configuration files for backup..." -ForegroundColor Yellow
    
    try {
        # Use the collections scanning method
        $configFiles = Get-SystemConfigFiles -Path $Path
        
        if ($configFiles.Count -eq 0) {
            Write-ColorOutput "⚠️  No launcher.windows.conf files found to backup!" -ForegroundColor Yellow
            return
        }
        
        Write-ColorOutput "💾 Starting backup of $($configFiles.Count) files..." -ForegroundColor Yellow
        Write-Host ""
        
        $backupCounter = 0
        $skippedCounter = 0
        
        foreach ($file in $configFiles) {
            $systemName = $file.Directory.Name
            $backupPath = $file.FullName + ".backup"
            
            try {
                if (-not (Test-Path $backupPath)) {
                    Copy-Item -Path $file.FullName -Destination $backupPath -ErrorAction Stop
                    $backupCounter++
                    $Global:Stats.BackupsCreated++
                    Write-ColorOutput "   ✓ [$($backupCounter + $skippedCounter)/$($configFiles.Count)] $systemName" -ForegroundColor Green
                } else {
                    $skippedCounter++
                    Write-ColorOutput "   ⚠ [$($backupCounter + $skippedCounter)/$($configFiles.Count)] $systemName (backup exists)" -ForegroundColor Yellow
                }
                
            }
            catch {
                Write-ColorOutput "   ✗ [$($backupCounter + $skippedCounter + 1)/$($configFiles.Count)] FAILED: $systemName" -ForegroundColor Red
                Write-ColorOutput "     Error: $($_.Exception.Message)" -ForegroundColor Red
                $Global:ErrorLog += "Backup failed: $($file.FullName) - $($_.Exception.Message)"
            }
        }
        
        $totalBackupTime = (Get-Date) - $backupStartTime
        Write-Host ""
        Write-ColorOutput "💾 Backup Summary (completed in $([math]::Round($totalBackupTime.TotalSeconds, 1)) seconds):" -ForegroundColor Cyan
        Write-ColorOutput "   ✓ New backups created: $backupCounter" -ForegroundColor Green
        Write-ColorOutput "   ⚠ Skipped (already existed): $skippedCounter" -ForegroundColor Yellow
        Write-ColorOutput "   📁 Total files: $($configFiles.Count)" -ForegroundColor White
        Write-Host ""
        
    }
    catch {
        Write-ColorOutput "❌ ERROR during backup: $($_.Exception.Message)" -ForegroundColor Red
        throw
    }
}

function Update-LaunchConfig {
    param(
        [string]$ConfigPath,
        [string]$SystemName,
        [switch]$WhatIf
    )
    
    try {
        $normalizedSystemName = $SystemName.ToLower().Trim()
        
        # Check if system is excluded
        if (Test-SystemExcluded -SystemName $SystemName) {
            Write-ColorOutput "   🚫 EXCLUDED (using your custom emulator)" -ForegroundColor Red
            $Global:Stats.ExcludedByUser++
            
            # Track excluded systems
            if (-not $Global:ExcludedSystems.ContainsKey($SystemName)) {
                $Global:ExcludedSystems[$SystemName] = 0
            }
            $Global:ExcludedSystems[$SystemName]++
            
            return $false
        }
        
        # Get RetroArch core
        $coreDll = $YourSystemCoreMap[$normalizedSystemName]
        
        if (-not $coreDll) {
            Write-ColorOutput "   ⚠️  No RetroArch mapping found" -ForegroundColor Yellow
            $Global:Stats.Skipped++
            return $false
        }
        
        # Track processed systems
        if (-not $Global:ProcessedSystems.ContainsKey($SystemName)) {
            $Global:ProcessedSystems[$SystemName] = @{
                Core = $coreDll
                Count = 0
            }
        }
        $Global:ProcessedSystems[$SystemName].Count++
        
        # Create configuration content with YOUR RetroArch path
        $configContent = @"
[launch]
executable = $RetroArchPath
working_directory = $RetroArchWorkingDir
arguments = -L $CoresPath\$coreDll "%ITEM_FILEPATH%"
wait_for_exit = true
block_input_during_launch = true
"@
        
        if ($WhatIf) {
            Write-ColorOutput "   🔮 PREVIEW: Would use RetroArchXiso core: $coreDll" -ForegroundColor Magenta
            return $true
        }
        
        # Write the configuration
        $configContent | Out-File -FilePath $ConfigPath -Encoding UTF8 -ErrorAction Stop
        Write-ColorOutput "   ✅ Updated -> $coreDll" -ForegroundColor Green
        $Global:Stats.Updated++
        return $true
        
    }
    catch {
        Write-ColorOutput "   ❌ ERROR: $($_.Exception.Message)" -ForegroundColor Red
        $Global:ErrorLog += "Update failed: $ConfigPath - $($_.Exception.Message)"
        $Global:Stats.Errors++
        return $false
    }
}

function Show-Summary {
    Write-Host ""
    Write-ColorOutput "🏁 PROCESSING COMPLETE" -ForegroundColor Green
    Write-ColorOutput "════════════════════════════════════════" -ForegroundColor Green
    
    if ($WhatIf) {
        Write-ColorOutput "🔮 PREVIEW MODE - No actual changes made" -ForegroundColor Magenta
        Write-ColorOutput "   📁 Files found: $($Global:Stats.TotalFound)" -ForegroundColor White
        Write-ColorOutput "   ✅ Would update: $($Global:Stats.Updated)" -ForegroundColor Green
        Write-ColorOutput "   🚫 Would exclude: $($Global:Stats.ExcludedByUser)" -ForegroundColor Red
        Write-ColorOutput "   ⚠️  Would skip: $($Global:Stats.Skipped)" -ForegroundColor Yellow
    } else {
        Write-ColorOutput "📁 Files processed: $($Global:Stats.TotalFound)" -ForegroundColor White
        Write-ColorOutput "✅ Updated with RetroArchXiso: $($Global:Stats.Updated)" -ForegroundColor Green
        Write-ColorOutput "🚫 Excluded (your custom emulators): $($Global:Stats.ExcludedByUser)" -ForegroundColor Red
        Write-ColorOutput "⚠️  Skipped (unknown): $($Global:Stats.Skipped)" -ForegroundColor Yellow
        Write-ColorOutput "💾 Backups created: $($Global:Stats.BackupsCreated)" -ForegroundColor Cyan
        
        if ($Global:Stats.Errors -gt 0) {
            Write-ColorOutput "❌ Errors: $($Global:Stats.Errors)" -ForegroundColor Red
        }
    }
    
    # Show RetroArch systems processed
    if ($Global:ProcessedSystems.Count -gt 0) {
        Write-ColorOutput "`n✅ RetroArchXiso Systems Updated:" -ForegroundColor Green
        $Global:ProcessedSystems.GetEnumerator() | Sort-Object Name | ForEach-Object {
            Write-ColorOutput "   🎮 $($_.Key) ($($_.Value.Count) files) -> $($_.Value.Core)" -ForegroundColor White
        }
    }
    
    # Show excluded systems
    if ($Global:ExcludedSystems.Count -gt 0) {
        Write-ColorOutput "`n🚫 Excluded Systems (Using Your Custom Emulators):" -ForegroundColor Red
        $Global:ExcludedSystems.GetEnumerator() | Sort-Object Name | ForEach-Object {
            Write-ColorOutput "   ❌ $($_.Key) ($($_.Value) files skipped)" -ForegroundColor Red
        }
    }
    
    if ($Global:ErrorLog.Count -gt 0) {
        Write-ColorOutput "`n❌ Error Details:" -ForegroundColor Red
        $Global:ErrorLog | ForEach-Object {
            Write-ColorOutput "   • $_" -ForegroundColor Red
        }
    }
}

# Main execution
function Main {
    try {
        Show-Header
        
        # Get path
        if (-not $CoinOpsPath) {
            $CoinOpsPath = Read-Host "📁 Enter your CoinOps directory path"
            if ([string]::IsNullOrWhiteSpace($CoinOpsPath)) {
                Write-ColorOutput "❌ No path provided!" -ForegroundColor Red
                return
            }
            $CoinOpsPath = $CoinOpsPath.Trim().Trim('"')
        }
        
        if (-not (Test-Path $CoinOpsPath)) {
            Write-ColorOutput "❌ Path not found: $CoinOpsPath" -ForegroundColor Red
            return
        }
        
        # Check if collections folder exists
        $collectionsPath = Join-Path $CoinOpsPath "collections"
        if (-not (Test-Path $collectionsPath)) {
            Write-ColorOutput "❌ Collections folder not found at: $collectionsPath" -ForegroundColor Red
            Write-ColorOutput "   Make sure you're pointing to the root CoinOps directory" -ForegroundColor Yellow
            return
        }
        
        # Get backup preference if not specified
        if (-not $WhatIf -and -not $PSBoundParameters.ContainsKey('CreateBackup')) {
            $backupResponse = Read-Host "💾 Create backups before updating? (y/n)"
            $CreateBackup = $backupResponse -eq 'y' -or $backupResponse -eq 'Y'
        }
        
        Write-ColorOutput "🔍 Target Directory: $CoinOpsPath" -ForegroundColor Cyan
        Write-ColorOutput "📁 Collections Path: $collectionsPath" -ForegroundColor Cyan
        Write-ColorOutput "🚫 Will SKIP your custom emulator systems (PlayStation, Xbox, Dreamcast, Wii/WiiU, PC)" -ForegroundColor Red
        Write-Host ""
        
        # Create backups with collections scanning
        if ($CreateBackup -and -not $WhatIf) {
            Backup-ConfigFiles -Path $CoinOpsPath
        }
        
        # Find config files using collections scanning
        Write-ColorOutput "🔎 Getting launcher.windows.conf files to update..." -ForegroundColor Yellow
        $configFiles = Get-SystemConfigFiles -Path $CoinOpsPath
        $Global:Stats.TotalFound = $configFiles.Count
        
        if ($configFiles.Count -eq 0) {
            Write-ColorOutput "⚠️  No launcher.windows.conf files found!" -ForegroundColor Yellow
            return
        }
        
        Write-ColorOutput "✅ Ready to process $($configFiles.Count) configuration files" -ForegroundColor Green
        Write-Host ""
        
        # Process files
        Write-ColorOutput "⚙️  Processing files with RetroArchXiso..." -ForegroundColor Yellow
        
        $updateStartTime = Get-Date
        $counter = 0
        foreach ($configFile in $configFiles) {
            $counter++
            $systemName = $configFile.Directory.Name
            
            Write-ColorOutput "[$counter/$($configFiles.Count)] 🎮 $systemName" -ForegroundColor White -NoNewline
            Update-LaunchConfig -ConfigPath $configFile.FullName -SystemName $systemName -WhatIf:$WhatIf | Out-Null
            
            # Show progress every 25 files
            if ($configFiles.Count -gt 50 -and $counter % 25 -eq 0) {
                $percent = [math]::Round(($counter / $configFiles.Count) * 100, 1)
                $elapsed = (Get-Date) - $updateStartTime
                $rate = $counter / $elapsed.TotalSeconds
                $eta = ($configFiles.Count - $counter) / $rate
                Write-ColorOutput "`n   📊 Progress: $percent% | ETA: $([math]::Round($eta, 0)) seconds" -ForegroundColor Cyan
            }
        }
        
        Show-Summary
        
    }
    catch {
        Write-ColorOutput "💥 SCRIPT ERROR: $($_.Exception.Message)" -ForegroundColor Red
        Write-ColorOutput "Line: $($_.InvocationInfo.ScriptLineNumber)" -ForegroundColor Red
    }
    finally {
        Write-Host ""
        if ($Host.Name -eq 'ConsoleHost') {
            Read-Host "Press Enter to exit"
        }
    }
}

# Execute the script
Main