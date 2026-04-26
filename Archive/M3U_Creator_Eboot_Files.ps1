<#
.SYNOPSIS
    PS3 Library .m3u Generator - PowerShell ISE Compatible (Fixed)
.DESCRIPTION
    Scans PS3 game directories up to 4 levels deep for EBOOT.BIN files and creates .m3u playlists.
.AUTHOR
    PS3 M3U Generator v2.1 - ISE Fixed Edition
#>

param(
    [string]$SourceDir = "",
    [string]$DestDir = "",
    [switch]$Force,
    [switch]$Quiet,
    [switch]$ShowSkipped
)

# Global configuration
$Config = @{
    MaxDepth = 4
    LogFile = ""
    StartTime = Get-Date
    Stats = @{
        Found = 0
        Created = 0
        Skipped = 0
        Errors = 0
    }
    IsISE = $host.Name -eq "Windows PowerShell ISE Host"
}

function Write-ColorMessage {
    param(
        [string]$Message,
        [string]$Color = "White"
    )
    
    if (-not $Quiet) {
        Write-Host $Message -ForegroundColor $Color
    }
}

function Write-LogEntry {
    param(
        [string]$Message,
        [string]$Level = "INFO"
    )
    
    if ($Config.LogFile -and (Test-Path (Split-Path $Config.LogFile -Parent))) {
        try {
            $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            $logEntry = "[$timestamp] [$Level] $Message"
            Add-Content -Path $Config.LogFile -Value $logEntry -Encoding UTF8 -ErrorAction SilentlyContinue
        } catch {
            # Ignore logging errors
        }
    }
}

function Show-ProgressMessage {
    param(
        [string]$Activity,
        [string]$Status,
        [int]$PercentComplete = 0
    )
    
    if ($Config.IsISE) {
        Write-ColorMessage "[$PercentComplete%] $Activity - $Status" -Color Yellow
    } else {
        Write-Progress -Activity $Activity -Status $Status -PercentComplete $PercentComplete
    }
}

function Get-DirectoryPath {
    param(
        [string]$Prompt,
        [string]$DefaultPath = ""
    )
    
    do {
        Write-Host ""
        if ($DefaultPath) {
            Write-Host "$Prompt" -ForegroundColor Cyan
            Write-Host "Default: $DefaultPath" -ForegroundColor Gray
            Write-Host "Press Enter for default or type new path: " -ForegroundColor White -NoNewline
            $userInput = Read-Host
            
            if ([string]::IsNullOrWhiteSpace($userInput)) {
                $path = $DefaultPath
            } else {
                $path = $userInput
            }
        } else {
            Write-Host "$Prompt" -ForegroundColor Cyan
            Write-Host "Enter path: " -ForegroundColor White -NoNewline
            $path = Read-Host
        }
        
        if ([string]::IsNullOrWhiteSpace($path)) {
            Write-ColorMessage "Please enter a valid directory path." -Color Red
            continue
        }
        
        # Clean up the path
        $path = $path.Trim().Trim('"').Trim("'")
        
        # Expand environment variables and relative paths
        try {
            $path = [System.Environment]::ExpandEnvironmentVariables($path)
            $path = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($path)
        } catch {
            Write-ColorMessage "Invalid path format: $path" -Color Red
            continue
        }
        
        # Check if directory exists, create if needed for destination
        if (-not (Test-Path -Path $path -PathType Container)) {
            if ($Prompt -like "*destination*" -or $Prompt -like "*output*") {
                $create = Read-Host "Directory doesn't exist. Create it? (y/N)"
                if ($create -eq 'y' -or $create -eq 'Y') {
                    try {
                        New-Item -Path $path -ItemType Directory -Force | Out-Null
                        Write-ColorMessage "Created directory: $path" -Color Green
                        Write-LogEntry "Created directory: $path"
                        return $path
                    } catch {
                        Write-ColorMessage "Failed to create directory: $_" -Color Red
                        continue
                    }
                }
            } else {
                Write-ColorMessage "Directory does not exist: $path" -Color Red
                continue
            }
        } else {
            return $path
        }
    } while ($true)
}

function Get-CleanFileName {
    param(
        [string]$FileName
    )
    
    if ([string]::IsNullOrWhiteSpace($FileName)) {
        return "UnknownGame"
    }
    
    # Replace invalid characters
    $cleanName = $FileName -replace '[<>:"/\\|?*]', '_'
    $cleanName = $cleanName -replace '\s+', '_'
    $cleanName = $cleanName -replace '_+', '_'
    $cleanName = $cleanName.Trim('._')
    
    # Handle Windows reserved names
    $reserved = @('CON', 'PRN', 'AUX', 'NUL', 'COM1', 'COM2', 'COM3', 'COM4', 'COM5', 'COM6', 'COM7', 'COM8', 'COM9', 'LPT1', 'LPT2', 'LPT3', 'LPT4', 'LPT5', 'LPT6', 'LPT7', 'LPT8', 'LPT9')
    if ($reserved -contains $cleanName.ToUpper()) {
        $cleanName = "Game_$cleanName"
    }
    
    # Limit length
    if ($cleanName.Length -gt 180) {
        $cleanName = $cleanName.Substring(0, 180)
    }
    
    if ([string]::IsNullOrWhiteSpace($cleanName)) {
        $cleanName = "UnknownGame"
    }
    
    return $cleanName
}

function Find-EbootFiles {
    param(
        [string]$SourcePath
    )
    
    Write-ColorMessage "Scanning for EBOOT.BIN files..." -Color Yellow
    Write-LogEntry "Starting scan of: $SourcePath"
    
    $gameList = @()
    $ebootFiles = @()
    
    try {
        # Search each depth level
        for ($depth = 1; $depth -le $Config.MaxDepth; $depth++) {
            $searchPattern = $SourcePath + ('\*' * $depth) + '\EBOOT.BIN'
            $foundFiles = @(Get-ChildItem -Path $searchPattern -File -ErrorAction SilentlyContinue)
            
            if ($foundFiles.Count -gt 0) {
                Write-ColorMessage "  Depth $depth`: Found $($foundFiles.Count) files" -Color Gray
                $ebootFiles += $foundFiles
            }
            
            Show-ProgressMessage -Activity "Scanning directories" -Status "Depth $depth of $($Config.MaxDepth)" -PercentComplete (($depth / $Config.MaxDepth) * 100)
        }
        
        Write-ColorMessage "Total EBOOT.BIN files found: $($ebootFiles.Count)" -Color Green
        $Config.Stats.Found = $ebootFiles.Count
        
        if ($ebootFiles.Count -eq 0) {
            return $gameList
        }
        
        # Process each EBOOT.BIN file
        $counter = 0
        foreach ($eboot in $ebootFiles) {
            $counter++
            
            # Find game name from directory structure
            $gameName = ""
            $currentDir = $eboot.Directory
            
            # Look up the directory tree for a meaningful name
            for ($i = 0; $i -lt 4; $i++) {
                if ($currentDir -and $currentDir.FullName -ne $SourcePath) {
                    if ($currentDir.Name -notmatch '^(PS3_GAME|USRDIR|EBOOT)$') {
                        $gameName = $currentDir.Name
                        break
                    }
                    $currentDir = $currentDir.Parent
                } else {
                    break
                }
            }
            
            # Fallback naming
            if ([string]::IsNullOrWhiteSpace($gameName) -or $gameName -match '^(PS3_GAME|USRDIR|EBOOT)$') {
                $pathParts = $eboot.FullName.Replace($SourcePath, '').Trim('\').Split('\')
                $gameName = $pathParts[0]
            }
            
            if ([string]::IsNullOrWhiteSpace($gameName)) {
                $gameName = "UnknownGame_$counter"
            }
            
            $gameInfo = [PSCustomObject]@{
                GameName = $gameName
                CleanName = Get-CleanFileName -FileName $gameName
                EbootPath = $eboot.FullName
                LastModified = $eboot.LastWriteTime
            }
            
            $gameList += $gameInfo
            
            if ($counter % 10 -eq 0 -or $counter -eq $ebootFiles.Count) {
                Show-ProgressMessage -Activity "Processing files" -Status "Processed $counter of $($ebootFiles.Count)" -PercentComplete (($counter / $ebootFiles.Count) * 100)
            }
        }
        
    } catch {
        Write-ColorMessage "Error during scan: $_" -Color Red
        Write-LogEntry "Error during scan: $_" -Level "ERROR"
    }
    
    return $gameList
}

function Test-ExistingM3u {
    param(
        [string]$FilePath,
        [string]$ExpectedContent
    )
    
    if (-not (Test-Path $FilePath)) {
        return $false
    }
    
    try {
        $content = Get-Content $FilePath -Raw -ErrorAction SilentlyContinue
        return $content.Trim() -eq $ExpectedContent.Trim()
    } catch {
        return $false
    }
}

function New-M3uFiles {
    param(
        [array]$GameList,
        [string]$OutputPath
    )
    
    Write-ColorMessage "Creating M3U files..." -Color Yellow
    Write-LogEntry "Creating M3U files in: $OutputPath"
    
    $results = @{
        Created = 0
        Skipped = 0
        Errors = 0
    }
    
    $counter = 0
    foreach ($game in $GameList) {
        $counter++
        $m3uFile = "$($game.CleanName).m3u"
        $m3uPath = Join-Path -Path $OutputPath -ChildPath $m3uFile
        
        # Check if we should skip existing files
        if (-not $Force -and (Test-ExistingM3u -FilePath $m3uPath -ExpectedContent $game.EbootPath)) {
            if ($ShowSkipped) {
                Write-ColorMessage "  Skipped: $m3uFile" -Color Gray
            }
            $results.Skipped++
            continue
        }
        
        try {
            # Write the M3U file
            $game.EbootPath | Out-File -FilePath $m3uPath -Encoding UTF8
            
            # Set timestamp
            try {
                (Get-Item $m3uPath).LastWriteTime = $game.LastModified
            } catch {
                # Ignore timestamp errors
            }
            
            Write-ColorMessage "  Created: $m3uFile" -Color Green
            Write-LogEntry "Created: $m3uFile -> $($game.EbootPath)"
            $results.Created++
            
        } catch {
            Write-ColorMessage "  Error: $m3uFile - $_" -Color Red
            Write-LogEntry "Error creating $m3uFile`: $_" -Level "ERROR"
            $results.Errors++
        }
        
        if ($counter % 10 -eq 0 -or $counter -eq $GameList.Count) {
            Show-ProgressMessage -Activity "Creating M3U files" -Status "Created $counter of $($GameList.Count)" -PercentComplete (($counter / $GameList.Count) * 100)
        }
    }
    
    $Config.Stats.Created = $results.Created
    $Config.Stats.Skipped = $results.Skipped
    $Config.Stats.Errors = $results.Errors
    
    return $results
}

function Show-Results {
    param(
        [array]$GameList,
        [hashtable]$Results
    )
    
    $duration = (Get-Date) - $Config.StartTime
    
    Write-Host ""
    Write-ColorMessage "=" * 60 -Color Cyan
    Write-ColorMessage "RESULTS SUMMARY" -Color Cyan
    Write-ColorMessage "=" * 60 -Color Cyan
    
    Write-ColorMessage "Execution Time: $($duration.ToString('mm\:ss'))" -Color White
    Write-ColorMessage "Games Found: $($Config.Stats.Found)" -Color Green
    Write-ColorMessage "M3U Created: $($Results.Created)" -Color Green
    
    if ($Results.Skipped -gt 0) {
        Write-ColorMessage "M3U Skipped: $($Results.Skipped)" -Color Yellow
    }
    
    if ($Results.Errors -gt 0) {
        Write-ColorMessage "Errors: $($Results.Errors)" -Color Red
    }
    
    # Show examples
    if ($GameList.Count -gt 0) {
        Write-Host ""
        Write-ColorMessage "Sample files created:" -Color Cyan
        
        $sampleCount = [Math]::Min(5, $GameList.Count)
        for ($i = 0; $i -lt $sampleCount; $i++) {
            $game = $GameList[$i]
            Write-ColorMessage "  $($game.CleanName).m3u" -Color White
            Write-ColorMessage "    -> $($game.EbootPath)" -Color Gray
        }
        
        if ($GameList.Count -gt 5) {
            Write-ColorMessage "  ... and $($GameList.Count - 5) more" -Color Gray
        }
    }
    
    Write-ColorMessage "=" * 60 -Color Cyan
}

function Initialize-Session {
    param(
        [string]$OutputPath
    )
    
    # Setup logging
    $logName = "PS3_M3U_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
    $Config.LogFile = Join-Path -Path $OutputPath -ChildPath $logName
    
    Write-LogEntry "=== PS3 M3U Generator Started ===" -Level "INFO"
    Write-LogEntry "Version: 2.1 ISE Fixed" -Level "INFO"
    Write-LogEntry "Host: $($host.Name)" -Level "INFO"
}

# Main execution starts here
Write-Host ""
Write-ColorMessage "=" * 60 -Color Cyan
Write-ColorMessage "PS3 Library M3U Generator v2.1" -Color Cyan
Write-ColorMessage "PowerShell ISE Compatible" -Color Cyan

# Prompt for source directory if not provided
$sourcePath = Get-DirectoryPath -Prompt "Enter the source directory containing PS3 game folders" -DefaultPath $SourceDir
if (-not $sourcePath) {
    Write-ColorMessage "No valid source directory provided. Exiting." -Color Red
    exit
}

# Prompt for destination directory if not provided
$destPath = Get-DirectoryPath -Prompt "Enter the destination directory for M3U files" -DefaultPath $DestDir
if (-not $destPath) {
    Write-ColorMessage "No valid destination directory provided. Exiting." -Color Red
    exit
}

# Initialize session
Initialize-Session -OutputPath $destPath

# Find EBOOT.BIN files
$gameList = Find-EbootFiles -SourcePath $sourcePath

# Create M3U files if any games were found
if ($gameList.Count -gt 0) {
    $results = New-M3uFiles -GameList $gameList -OutputPath $destPath
    Show-Results -GameList $gameList -Results $results
} else {
    Write-ColorMessage "No EBOOT.BIN files found. No M3U files created." -Color Yellow
}

Write-LogEntry "=== PS3 M3U Generator Finished ===" -Level "INFO"