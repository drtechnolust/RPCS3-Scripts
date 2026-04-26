<# 
.SYNOPSIS
  Copies CoinOPS "Nostalgic Room" image assets from each system into three top-level
  folders (Nostalgic Room / Nostalgic Room Night / TV Promox), renaming each file
  to the LaunchBox platform naming scheme.

.PARAMETER CollectionsRoot
  The path to ...\CoinOPS Deluxe UNIVERSE 2025\layouts\Arcades\collections

.PARAMETER ImagesRelativePath
  Relative path from each system folder to its images folder (default: 'layout\images')

.PARAMETER DryRun
  If set, only prints what would happen without making changes.

.PARAMETER Force
  Overwrite existing files in destination folders.

.EXAMPLE
  Press F5 to run with default path (edit $DefaultCollectionsRoot below)

.EXAMPLE
  .\Export-NostalgicRoomAssets.ps1 -CollectionsRoot 'C:\Arcade\CoinOPS Deluxe UNIVERSE 2025\layouts\Arcades\collections' -DryRun -Verbose

.EXAMPLE
  .\Export-NostalgicRoomAssets.ps1 -CollectionsRoot 'C:\Arcade\CoinOPS Deluxe UNIVERSE 2025\layouts\Arcades\collections' -Force

#>

[CmdletBinding(SupportsShouldProcess=$true)]
param(
  [Parameter(Mandatory=$false)]
  [string]$CollectionsRoot,

  [string]$ImagesRelativePath = 'layout\images',

  [switch]$DryRun,

  [switch]$Force
)

# ============================================================================
# ISE-FRIENDLY CONFIGURATION
# Edit this path to your CoinOPS collections folder, then press F5 to run!
# ============================================================================
$DefaultCollectionsRoot = 'C:\Arcade\CoinOPS Deluxe UNIVERSE 2025\layouts\Arcades\collections'
$DefaultDryRun = $false  # Set to $false to actually copy files when pressing F5
# ============================================================================

# If no CollectionsRoot provided (e.g., running with F5), use defaults
if ([string]::IsNullOrWhiteSpace($CollectionsRoot)) {
  $CollectionsRoot = $DefaultCollectionsRoot
  if (-not $PSBoundParameters.ContainsKey('DryRun')) {
    $DryRun = $DefaultDryRun
  }
  Write-Host "Using default settings (press F5 mode):" -ForegroundColor Magenta
  Write-Host "  CollectionsRoot: $CollectionsRoot" -ForegroundColor Gray
  Write-Host "  DryRun: $DryRun`n" -ForegroundColor Gray
}

# Validate path exists
if (-not (Test-Path $CollectionsRoot -PathType Container)) {
  Write-Error "CollectionsRoot path does not exist: $CollectionsRoot"
  Write-Host "`nPlease edit the `$DefaultCollectionsRoot variable at the top of this script." -ForegroundColor Yellow
  return
}

# --- Config: target folders and source file names we expect ---
$TargetFolders = @(
  'Nostalgic Room',
  'Nostalgic Room Night',
  'TV Promox'
)

$ExpectedFiles = @{
  'Nostalgic Room'       = 'Nostalgic Room.png'
  'Nostalgic Room Night' = 'Nostalgic Room Night.png'
  'TV Promox'            = 'TV Promox.png'
}

# --- Mapping from CoinOPS folder names -> LaunchBox platform names ---
$PlatformMap = @{
  'Amstrad CPC'                       = 'Amstrad CPC'
  'Arcades'                           = 'Arcade'
  'Atari 2600'                        = 'Atari 2600'
  'Atari 5200'                        = 'Atari 5200'
  'Atari 7800'                        = 'Atari 7800'
  'Atari Jaguar'                      = 'Atari Jaguar'
  'Atari Lynx'                        = 'Atari Lynx'
  'Colecovision'                      = 'ColecoVision'
  'Nintendo Entertainment System'     = 'Nintendo Entertainment System'
  'Super Nintendo'                    = 'Super Nintendo Entertainment System'
  'Nintendo 64'                       = 'Nintendo 64'
  'Nintendo Switch'                   = 'Nintendo Switch'
  'Game Boy'                          = 'Nintendo Game Boy'
  'Game Boy Color'                    = 'Nintendo Game Boy Color'
  'Game Boy Advance'                  = 'Nintendo Game Boy Advance'
  'GameCube'                          = 'Nintendo GameCube'
  'Genesis'                           = 'Sega Genesis'
  'Megadrive'                         = 'Sega Mega Drive'
  'Master System'                     = 'Sega Master System'
  'Game Gear'                         = 'Sega Game Gear'
  'Sega 32X'                          = 'Sega 32X'
  'Sega CD'                           = 'Sega CD'
  'Sega SG-1000'                      = 'Sega SG-1000'
  'Dreamcast'                         = 'Sega Dreamcast'
  'Saturn'                            = 'Sega Saturn'
  'Intellivision'                     = 'Mattel Intellivision'
  'MSX'                               = 'MSX'
  'MSX2'                              = 'MSX2'
  'Commodore 64'                      = 'Commodore 64'
  'Commodore Amiga'                   = 'Commodore Amiga'
  'PC Engine'                         = 'NEC PC Engine'
  'TurboGrafx-16'                     = 'NEC TurboGrafx-16'
  'Panasonic 3DO'                     = '3DO Interactive Multiplayer'
  'Odyssey 2'                         = 'Magnavox Odyssey 2'
  'Vectrex'                           = 'Vectrex'
  'Playstation 1'                     = 'Sony Playstation'
  'Playstation 2'                     = 'Sony Playstation 2'
  'Playstation 3'                     = 'Sony Playstation 3'
  'Playstation Portable'              = 'Sony PSP'
  'Wii'                               = 'Nintendo Wii'
  'Wii U'                             = 'Nintendo Wii U'
  'Xbox'                              = 'Microsoft Xbox'
  'Xbox 360'                          = 'Microsoft Xbox 360'
  'Game and Watch'                    = 'Nintendo Game and Watch'
  'Neo Geo'                           = 'SNK Neo Geo AES'
  'Neo Geo Pocket'                    = 'SNK Neo Geo Pocket'
  'Neo Geo Pocket Color'              = 'SNK Neo Geo Pocket Color'
  'WonderSwan'                        = 'Bandai WonderSwan'
  'WonderSwan Color'                  = 'Bandai WonderSwan Color'
  'Pinball'                           = 'Pinball'
  'PC Gamer'                          = 'Windows'
  'Racer'                             = 'Racing'
}

$ExcludeSystems = @('_common', 'zzzShutdown')

# --- Statistics tracking ---
$stats = @{
  Copied = 0
  Skipped_Missing = 0
  Skipped_NoImages = 0
  Skipped_Exists = 0
  SystemsProcessed = 0
}

# --- Prep: ensure target folders exist ---
Write-Host "`n=== Nostalgic Room Asset Export ===" -ForegroundColor Cyan
Write-Host "Collections Root: $CollectionsRoot" -ForegroundColor Gray
Write-Host "Mode: $(if ($DryRun) { 'DRY RUN (no changes will be made)' } else { 'LIVE' })`n" -ForegroundColor $(if ($DryRun) { 'Yellow' } else { 'Green' })

foreach ($tf in $TargetFolders) {
  $dest = Join-Path $CollectionsRoot $tf
  if (-not (Test-Path $dest)) {
    if ($DryRun) {
      Write-Host "[DRY RUN] Would create folder: $dest" -ForegroundColor Yellow
    } else {
      try {
        New-Item -ItemType Directory -Path $dest -ErrorAction Stop | Out-Null
        Write-Verbose "Created folder: $dest"
      }
      catch {
        Write-Error "Failed to create folder '$dest': $_"
        return
      }
    }
  }
}

# --- Enumerate system folders ---
try {
  $systemDirs = Get-ChildItem -Path $CollectionsRoot -Directory -ErrorAction Stop |
    Where-Object { $_.Name -notin $TargetFolders -and $_.Name -notin $ExcludeSystems }
}
catch {
  Write-Error "Failed to enumerate system folders: $_"
  return
}

if (-not $systemDirs) {
  Write-Warning "No system folders found under: $CollectionsRoot"
  return
}

Write-Host "Found $($systemDirs.Count) system folder(s) to process`n" -ForegroundColor Cyan

# --- Process each system ---
foreach ($sys in $systemDirs) {
  $systemName = $sys.Name
  $stats.SystemsProcessed++
  
  $lbName = $PlatformMap[$systemName]
  if ([string]::IsNullOrWhiteSpace($lbName)) { 
    $lbName = $systemName
    Write-Verbose "No mapping for '$systemName', using folder name as-is"
  }

  $imagesDir = Join-Path $sys.FullName $ImagesRelativePath
  if (-not (Test-Path $imagesDir)) {
    Write-Verbose "No images folder for '$systemName' at $imagesDir"
    $stats.Skipped_NoImages++
    continue
  }

  foreach ($tf in $TargetFolders) {
    $expectedFile = $ExpectedFiles[$tf]
    $src = Join-Path $imagesDir $expectedFile
    
    if (-not (Test-Path $src)) {
      Write-Verbose "Missing '$expectedFile' for '$systemName'"
      $stats.Skipped_Missing++
      continue
    }

    $destFolder = Join-Path $CollectionsRoot $tf
    $destFile = Join-Path $destFolder ("{0}.png" -f $lbName)

    # Check if destination exists and Force is not set
    if ((Test-Path $destFile) -and -not $Force -and -not $DryRun) {
      Write-Verbose "Skipping existing file (use -Force to overwrite): $destFile"
      $stats.Skipped_Exists++
      continue
    }

    if ($DryRun) {
      Write-Host "[DRY RUN] Would copy:" -ForegroundColor Yellow
      Write-Host "  From: $src" -ForegroundColor Gray
      Write-Host "  To:   $destFile" -ForegroundColor Gray
    } else {
      try {
        Copy-Item -LiteralPath $src -Destination $destFile -Force:$Force -ErrorAction Stop
        $stats.Copied++
        Write-Verbose "Copied: $systemName -> $tf/$lbName.png"
      }
      catch {
        Write-Warning "Failed to copy '$src' to '$destFile': $_"
      }
    }
  }
}

# --- Summary Report ---
Write-Host "`n=== Summary ===" -ForegroundColor Cyan
Write-Host "Systems processed:        $($stats.SystemsProcessed)" -ForegroundColor White
Write-Host "Files copied:             $($stats.Copied)" -ForegroundColor Green
Write-Host "Missing images:           $($stats.Skipped_Missing)" -ForegroundColor Yellow
Write-Host "No images folder:         $($stats.Skipped_NoImages)" -ForegroundColor Yellow
Write-Host "Already exists (skipped): $($stats.Skipped_Exists)" -ForegroundColor Gray

if ($DryRun) {
  Write-Host "`nThis was a DRY RUN. No files were actually copied." -ForegroundColor Yellow
  Write-Host "Run without -DryRun to perform the actual copy operation." -ForegroundColor Yellow
}

Write-Host ""