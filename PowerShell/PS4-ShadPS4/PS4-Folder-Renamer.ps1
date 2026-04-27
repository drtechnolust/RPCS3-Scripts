#Requires -Version 5.1
<#
.SYNOPSIS
    Renames PS4 game folders to proper title case with gaming-aware formatting rules.

.DESCRIPTION
    Scans a folder of PS4 game directories and renames each one using title case
    with a comprehensive set of gaming-specific rules:

      - Series patterns applied first (Call of Duty, God of War, etc.)
      - Common corrections applied (Spider-Man, Pac-Man, X-Men, etc.)
      - Uppercase preserved for known acronyms (NBA, NFL, RPG, HD, VR, etc.)
      - Roman numerals preserved (II, III, IV ... XX)
      - Articles and prepositions kept lowercase when not at start of title
      - Asian characters (Japanese, Korean, Chinese) preserved as-is
      - Windows-forbidden characters stripped

    Run with DryRun = $true first to preview all renames before applying.
    A CSV log records every rename decision.

.PARAMETER DryRun
    When $true (default), shows what would be renamed without touching any folders.
    Set to $false in the CONFIG block to apply.

.EXAMPLE
    .\PS4-Folder-Renamer.ps1

.VERSION
    2.0.0 - Config block, DryRun default, CSV log, removed debug output and emoji,
            removed interactive prompts, ISE-compatible exits. MIT header.
    1.0.0 - Initial release.

.LICENSE
    MIT License
    Copyright (c) Paul Mardis
#>

# ==============================================================================
# CONFIG -- Edit this block. Do not put paths anywhere else in this script.
# ==============================================================================
$GameLibraryPath = "D:\Arcade\System roms\Sony Playstation 4\Official PS4 Games"
$DryRun          = $true    # Set to $false to actually rename folders
$LogDir          = Join-Path $GameLibraryPath "_Logs"
# ==============================================================================

function Set-ExitCode {
    param([int]$Code)
    $global:LASTEXITCODE = $Code
}

# ------------------------------------------------------------------------------
# Pre-flight
# ------------------------------------------------------------------------------
if (-not (Test-Path -LiteralPath $GameLibraryPath)) {
    Write-Host "ERROR: Game library path not found: $GameLibraryPath" -ForegroundColor Red
    Set-ExitCode 1
    return
}

# ==============================================================================
# FORMATTING RULES
# Extend these lists as needed -- they are the only place you should edit
# game-specific naming logic.
# ==============================================================================

# Words always rendered in full uppercase
$UppercaseWords = @(
    'NBA','NFL','WWE','UFC','NHL','MLB','FIFA','MMA',
    'HD','VR','DLC','GOTY','PS4','PS5','PC','RPG','FPS','RTS','DX',
    'US','EU','UK','JP','USA','TV','AI','MMO','PvP','PvE'
)

# Roman numerals always rendered uppercase
$RomanNumerals = @(
    'II','III','IV','V','VI','VII','VIII','IX','X',
    'XI','XII','XIII','XIV','XV','XVI','XVII','XVIII','XIX','XX'
)

# Articles and prepositions kept lowercase unless at start of title
$LowercaseWords = @(
    'a','an','and','as','at','but','by','for','if',
    'in','nor','of','on','or','so','the','to','up','yet','vs','with'
)

# Exact string corrections (applied case-insensitively by key)
$Corrections = @{
    'spiderman'  = 'Spider-Man'
    'xmen'       = 'X-Men'
    'pacman'     = 'Pac-Man'
    'cyberpunk'  = 'Cyberpunk'
    'bloodborne' = 'Bloodborne'
    'ratchet and clank' = 'Ratchet and Clank'
    'nioh'       = 'Nioh'
    'sekiro'     = 'Sekiro'
    'persona'    = 'Persona'
}

# Regex patterns matched against the lowercased title; first match wins
# $1 captures the rest of the title after the series name
$SeriesPatterns = [ordered]@{
    '^call of duty (.+)$'         = 'Call of Duty $1'
    '^metal gear solid (.+)$'     = 'Metal Gear Solid $1'
    '^grand theft auto(.*)$'      = 'Grand Theft Auto$1'
    '^detroit become human(.*)$'  = 'Detroit Become Human$1'
    '^god of war(.*)$'            = 'God of War$1'
    '^dragon ball (.+)$'          = 'Dragon Ball $1'
    '^dead rising(.*)$'           = 'Dead Rising$1'
    '^final fantasy (.+)$'        = 'Final Fantasy $1'
    '^assassins creed(.*)$'       = "Assassin's Creed$1"
    '^resident evil (.+)$'        = 'Resident Evil $1'
    '^the last of us(.*)$'        = 'The Last of Us$1'
    '^horizon (.+)$'              = 'Horizon $1'
    '^marvel(.*)$'                = 'Marvel$1'
    '^star wars (.+)$'            = 'Star Wars $1'
}

# ==============================================================================
# FORMAT FUNCTION
# ==============================================================================
function Format-GameTitle {
    param([string]$Title)

    if ([string]::IsNullOrWhiteSpace($Title)) { return $Title }

    $working = $Title.Trim()

    # Preserve Asian-script titles untouched
    if ($working -match '[\u3040-\u30FF\u4E00-\u9FFF\uAC00-\uD7AF]') {
        return $working
    }

    $lower = $working.ToLower()

    # Apply series patterns (first match wins)
    foreach ($Pattern in $SeriesPatterns.Keys) {
        if ($lower -match $Pattern) {
            $lower = $lower -ireplace $Pattern, $SeriesPatterns[$Pattern]
            break
        }
    }

    # Apply exact corrections
    foreach ($Key in $Corrections.Keys) {
        $Escaped = [regex]::Escape($Key)
        $lower = $lower -ireplace $Escaped, $Corrections[$Key]
    }

    # Word-by-word title case
    $Words     = $lower -split '\s+'
    $Formatted = @()

    for ($i = 0; $i -lt $Words.Count; $i++) {
        $Word = $Words[$i]
        if ([string]::IsNullOrEmpty($Word)) { continue }

        $Clean      = $Word -replace '[^\w]', ''
        $UpperClean = $Clean.ToUpper()
        $LowerClean = $Clean.ToLower()

        if ($UppercaseWords -contains $UpperClean) {
            $Formatted += $Word -ireplace [regex]::Escape($Clean), $UpperClean
        }
        elseif ($RomanNumerals -contains $UpperClean) {
            $Formatted += $Word -ireplace [regex]::Escape($Clean), $UpperClean
        }
        elseif ($i -gt 0 -and ($LowercaseWords -contains $LowerClean)) {
            $Formatted += $Word -ireplace [regex]::Escape($Clean), $LowerClean
        }
        elseif ($Clean -match '^\d+$') {
            $Formatted += $Word
        }
        else {
            if ($Clean.Length -gt 0) {
                $TitleCased = $Clean.Substring(0,1).ToUpper() + $Clean.Substring(1).ToLower()
                $Formatted += $Word -ireplace [regex]::Escape($Clean), $TitleCased
            }
        }
    }

    $Result = ($Formatted -join ' ') -replace '\s+', ' '
    $Result = $Result.Trim()

    # Strip Windows-forbidden characters
    $Result = $Result -replace '[<>:"|?*/\\]', ''

    return $Result
}

# ==============================================================================
# CSV LOG
# ==============================================================================
$LogRows = [System.Collections.Generic.List[PSCustomObject]]::new()

function Write-LogRow {
    param(
        [string]$OriginalName,
        [string]$NewName,
        [string]$FolderPath,
        [string]$Action,
        [string]$Status,
        [string]$Error
    )
    $LogRows.Add([PSCustomObject]@{
        Timestamp    = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
        OriginalName = $OriginalName
        NewName      = $NewName
        FolderPath   = $FolderPath
        Action       = $Action
        Status       = $Status
        Error        = $Error
    })
}

# ==============================================================================
# SETUP
# ==============================================================================
$LogFile = Join-Path $LogDir ("PS4-Folder-Renamer_{0}.csv" -f (Get-Date -Format "yyyyMMdd_HHmmss"))

if (-not $DryRun -and -not (Test-Path -LiteralPath $LogDir)) {
    New-Item -ItemType Directory -Path $LogDir | Out-Null
}

# ==============================================================================
# BANNER
# ==============================================================================
Write-Host ""
Write-Host "PS4 Folder Renamer" -ForegroundColor Cyan
Write-Host "==================" -ForegroundColor Cyan
Write-Host "  Library : $GameLibraryPath"
Write-Host "  DryRun  : $DryRun"
Write-Host ""
if ($DryRun) {
    Write-Host "  [DRY RUN] No folders will be renamed." -ForegroundColor Yellow
    Write-Host ""
}

# ==============================================================================
# SCAN
# ==============================================================================
$Folders = Get-ChildItem -LiteralPath $GameLibraryPath -Directory |
           Where-Object { $_.Name -ne "_Logs" } |
           Sort-Object Name

if ($Folders.Count -eq 0) {
    Write-Host "  No folders found in: $GameLibraryPath" -ForegroundColor Yellow
    Set-ExitCode 0
    return
}

Write-Host ("  Found {0} folder(s) to evaluate." -f $Folders.Count) -ForegroundColor White
Write-Host ""

$CountRenamed  = 0
$CountNoChange = 0
$CountFailed   = 0

# ==============================================================================
# MAIN LOOP
# ==============================================================================
foreach ($Folder in $Folders) {
    $OriginalName = $Folder.Name
    $NewName      = Format-GameTitle -Title $OriginalName

    if ($OriginalName -eq $NewName) {
        Write-Host ("  [OK]     {0}" -f $OriginalName) -ForegroundColor DarkGray
        Write-LogRow -OriginalName $OriginalName -NewName $NewName -FolderPath $Folder.FullName `
            -Action "Rename" -Status "NoChange" -Error ""
        $CountNoChange++
        continue
    }

    Write-Host ("  [RENAME] {0}" -f $OriginalName) -ForegroundColor Yellow
    Write-Host ("           -> {0}" -f $NewName)    -ForegroundColor Cyan

    if ($DryRun) {
        Write-LogRow -OriginalName $OriginalName -NewName $NewName -FolderPath $Folder.FullName `
            -Action "Rename" -Status "DryRun" -Error ""
        $CountRenamed++
        continue
    }

    # Check collision
    $NewPath = Join-Path (Split-Path $Folder.FullName -Parent) $NewName
    if (Test-Path -LiteralPath $NewPath) {
        Write-Host ("    SKIP   Target already exists: {0}" -f $NewName) -ForegroundColor Yellow
        Write-LogRow -OriginalName $OriginalName -NewName $NewName -FolderPath $Folder.FullName `
            -Action "Rename" -Status "Skipped" -Error "Target path already exists"
        $CountNoChange++
        continue
    }

    try {
        Rename-Item -LiteralPath $Folder.FullName -NewName $NewName -ErrorAction Stop
        Write-Host "    OK" -ForegroundColor Green
        Write-LogRow -OriginalName $OriginalName -NewName $NewName -FolderPath $Folder.FullName `
            -Action "Rename" -Status "Success" -Error ""
        $CountRenamed++
    }
    catch {
        Write-Host ("    FAILED: {0}" -f $_.Exception.Message) -ForegroundColor Red
        Write-LogRow -OriginalName $OriginalName -NewName $NewName -FolderPath $Folder.FullName `
            -Action "Rename" -Status "Failed" -Error $_.Exception.Message
        $CountFailed++
    }
}

# ==============================================================================
# WRITE LOG
# ==============================================================================
if ($LogRows.Count -gt 0 -and -not $DryRun) {
    try {
        $LogRows | Export-Csv -LiteralPath $LogFile -NoTypeInformation -Encoding UTF8
        Write-Host ""
        Write-Host "  Log: $LogFile" -ForegroundColor Gray
    }
    catch {
        Write-Host "  WARNING: Could not write log -- $_" -ForegroundColor Yellow
    }
}

# ==============================================================================
# SUMMARY
# ==============================================================================
Write-Host ""
Write-Host "Summary" -ForegroundColor White
Write-Host ("  {0,-14} {1}" -f "Renamed :", $CountRenamed)  -ForegroundColor $(if ($CountRenamed -gt 0) { "Green" } else { "Gray" })
Write-Host ("  {0,-14} {1}" -f "No change :", $CountNoChange) -ForegroundColor Gray
Write-Host ("  {0,-14} {1}" -f "Failed :", $CountFailed)    -ForegroundColor $(if ($CountFailed -gt 0) { "Red" } else { "Gray" })

if ($DryRun) {
    Write-Host ""
    Write-Host "  DRY RUN complete. Set `$DryRun = `$false in the CONFIG block to apply." -ForegroundColor Yellow
}

Write-Host ""
Set-ExitCode 0
