<# 
Populate platform artwork using a fixed list (from screenshots) + smart aliases/regex
DrTechnolust build — no LaunchBox XML needed

Usage (dry-run first):
.\Populate-Art-FromList.ps1 `
 -ArtworkSourceDir "C:\Arcade\...\Media\Nostalgic Room" `
 -OutputDir "C:\Arcade\...\Media\Nostalgic Room\Extras" `
 -DryRun
#>

[CmdletBinding()]
param(
  [Parameter(Mandatory=$true)]
  [string]$ArtworkSourceDir,      # Folder with _Default.png, Windows.png, system PNGs
  [Parameter(Mandatory=$true)]
  [string]$OutputDir,             # Where the per-platform PNGs go (can be same as source)
  [switch]$DryRun
)

function Sanitize-FileName {
  param([string]$Name)
  $invalid = [IO.Path]::GetInvalidFileNameChars() -join ''
  $regex = "[{0}]" -f [RegEx]::Escape($invalid)
  return ($Name -replace $regex, ' ').Trim()
}

# --- prereqs ---
foreach ($p in @($ArtworkSourceDir,$OutputDir)) {
  if (-not (Test-Path $p)) { throw "Path not found: $p" }
}
$DefaultPng = Join-Path $ArtworkSourceDir '_Default.png'
$WindowsPng = Join-Path $ArtworkSourceDir 'Windows.png'
if (-not (Test-Path $DefaultPng)) { Write-Warning "_Default.png missing in $ArtworkSourceDir" }
if (-not (Test-Path $WindowsPng)) { Write-Warning "Windows.png missing in $ArtworkSourceDir (computer fallback)" }

# ---------------------------------------------------------------------
# Platform names — consolidated from your screenshots (consoles, micros, customs)
# ---------------------------------------------------------------------
$platforms = @(
  # --- Computers / Micros / Collections shown ---
  'Ramber Pegasus','Acorn Archimedes','Acorn Atom','Acorn Electron','Apogey BK-01',
  'Apple II','Apple IIGS','Atari ST','BBC Microcomputer System','Boomer Shooters',
  'Camputers Lynx','Commodore 64','Commodore CDTV','Commodore MAX Machine',
  'Commodore PET','Commodore VIC-20','Daphne','DICE','Dragon 32/64','Elektronika BK',
  'Exidy Sorcerer','Fujitsu FM-7','Jupiter Ace','Mattel Aquarius','Memotech MTX512',
  'Microsoft MS-DOS','Microsoft MSX','Microsoft MSX2','Microsoft MSX2+',
  'NEC PC-88','NEC PC-9801','NESiCAxLive','OpenBOR','Oric Atmos','PC Games','PopCap',
  'SAM Coupé','ScummVM','Sega SC-3000','Sharp X68000','Sinclair ZX Spectrum',
  'Tandy TRS-80','Tomy Tutor','Vector-06C','Windows',

  # --- Consoles (full list) ---
  '3DO Interactive Multiplayer','Amstrad GX4000','APF Imagination Machine',
  'Atari 2600','Atari 5200','Atari 7800','Atari Jaguar','Atari Jaguar CD',
  'Bally Astrocade','Casio Loopy','Casio PV-1000','Casio PV-2000','ColecoVision',
  'Commodore Amiga CD32','Emerson Arcadia 2001','Entex Adventure Vision',
  'Epoch Super Cassette Vision','Fairchild Channel F','Funtech Super Acan',
  'GCE Vectrex','Interton VC 4000','Magnavox Odyssey','Magnavox Odyssey 2',
  'Mattel Intellivision',
  'Microsoft Xbox','Microsoft Xbox 360','Microsoft Xbox 360 Live Arcade',
  'NEC PC Engine','NEC PC Engine-CD','NEC PC-FX','NEC SuperGrafx','NEC TurboGrafx-16','NEC TurboGrafx-CD',
  'Nintendo 64','Nintendo 64DD','Nintendo Entertainment System',
  'Super Nintendo Entertainment System','Super Nintendo Entertainment System (Hacks)',
  'Nintendo Famicom','Nintendo Famicom Disk System','Nintendo GameCube',
  'Nintendo Satellaview','Nintendo Super Famicom',
  'Nintendo Switch','Nintendo Switch E-Shop','Nintendo Wii','Nintendo Wii U',
  'Nintendo Wii U Virtual Console GBA','Nintendo Wii U Virtual Console MSX',
  'Nintendo Wii U Virtual Console N64','Nintendo Wii U Virtual Console NES',
  'Nintendo Wii U Virtual Console Nintendo DS','Nintendo Wii U Virtual Console SFC',
  'Nintendo Wii U Virtual Console TG16-PCE','Nintendo WiiWare',
  'Nuon','Othello Multivision','Philips CD-i','Sega 32X','Sega CD','Sega Dreamcast',
  'Sega Dreamcast Indies','Sega Genesis','Sega Mark III','Sega Master System',
  'Sega Master System Imports','Sega Mega Drive','Sega Mega-CD','Sega Saturn',
  'SNK Neo Geo CD','Sony PlayStation','Sony PlayStation 2','Sony PlayStation 2 Asia',
  'Sony PlayStation 3','Sony PlayStation 4','Sony PlayStation Asia','Sony PlayStation Japan',
  'Super Nintendo Entertainment System Translated','Super Nintendo MSU-1',
  'VTech CreatiVision','VTech Socrates','VTech V.Smile','WoW Action Max',

  # --- Handhelds & extras ---
  'Atari Lynx','Bandai Sufami Turbo','Bandai WonderSwan','Bandai WonderSwan Color',
  'Epoch Game Pocket Computer','GamePark GP32','Hartung Game Master','Mega Duck',
  'Nintendo 3DS','Nintendo 3DS Europe','Nintendo 3DS Japan','Nintendo 64 Japan',
  'Nintendo DS','Nintendo DS Japan','Nintendo Game Boy','Nintendo Game Boy Advance',
  'Nintendo Game Boy Color','Nintendo Pokemon Mini','Nintendo Virtual Boy',
  'Sega Dreamcast VMU','Sega Game Gear','Sega Pico','SNK Neo Geo Pocket Color',
  'Sony PlayStation Minis','Sony PlayStation Vita','Sony PocketStation',
  'Sony PSP','Sony PSP JP','Tiger Game.com',

  # --- SMW collections (map to SNES art) ---
  'Super Mario World','SMW Adventures','SMW Collabs','SMW Contests','SMW Enhanced',
  'SMW General','SMW Imitations','SMW Japanese','SMW Jokes','SMW Kaizo',
  'SMW Puzzle','SMW Thrash','SMW Unknown',

  # misc playlists
  'Favorites'
) | Sort-Object -Unique

# ---------------------------------------------------------------------
# Aliases (map to existing backdrops you actually have)
# ---------------------------------------------------------------------
$alias = [ordered]@{}

# Build with assignments so duplicates never throw
$alias['Sega Mega Drive']   = 'Sega Genesis'
$alias['Sega Mega-CD']      = 'Sega Mega Drive'
$alias['Sega CD']           = 'Sega Mega Drive'
$alias['Sega Mark III']     = 'Sega Master System'
$alias['Sega SC-3000']      = 'Sega Master System'
$alias['Sega Dreamcast Indies'] = 'Sega Dreamcast'
$alias['Sega Master System Imports'] = 'Sega Master System'

$alias['Nintendo Super Famicom'] = 'Super Nintendo Entertainment System'
$alias['Super Famicom']          = 'Super Nintendo Entertainment System'
$alias['Nintendo Famicom']       = 'Nintendo Entertainment System'
$alias['Famicom']                = 'Nintendo Entertainment System'
$alias['Nintendo Satellaview']   = 'Super Nintendo Entertainment System'
$alias['Nintendo 64DD']          = 'Nintendo 64'
$alias['Nintendo 64 Japan']      = 'Nintendo 64'
$alias['Nintendo Switch E-Shop'] = 'Nintendo Switch'    # hyphen spelling
$alias['Nintendo WiiWare']       = 'Nintendo Wii'

# Wii U Virtual Console → use Wii U backdrop
$alias['Nintendo Wii U Virtual Console GBA']        = 'Nintendo Wii U'
$alias['Nintendo Wii U Virtual Console MSX']        = 'Nintendo Wii U'
$alias['Nintendo Wii U Virtual Console N64']        = 'Nintendo Wii U'
$alias['Nintendo Wii U Virtual Console NES']        = 'Nintendo Wii U'
$alias['Nintendo Wii U Virtual Console Nintendo DS'] = 'Nintendo Wii U'
$alias['Nintendo Wii U Virtual Console SFC']        = 'Nintendo Wii U'
$alias['Nintendo Wii U Virtual Console TG16-PCE']   = 'Nintendo Wii U'

# NEC family → consolidate to PC Engine art you have
$alias['NEC TurboGrafx-16'] = 'NEC PC Engine'
$alias['NEC TurboGrafx-CD'] = 'NEC PC Engine'
$alias['NEC SuperGrafx']    = 'NEC PC Engine'
$alias['NEC PC Engine-CD']  = 'NEC PC Engine'
$alias['NEC PC-FX']         = '_Default'   # no dedicated art in your set

# Sony variants
$alias['Sony PlayStation Asia']  = 'Sony PlayStation'
$alias['Sony PlayStation Japan'] = 'Sony PlayStation'
$alias['Sony PlayStation 2 Asia']= 'Sony PlayStation 2'
$alias['Sony PSP JP']            = 'Sony PSP'
$alias['Sony PlayStation 4']     = '_Default'   # no PS4 art in your set

# Odd consoles w/o art
$alias['Nuon']                 = '_Default'
$alias['Othello Multivision']  = '_Default'
$alias['Philips CD-i']         = '_Default'
$alias['SNK Neo Geo CD']       = '_Default'
$alias['VTech CreatiVision']   = '_Default'
$alias['VTech Socrates']       = '_Default'
$alias['VTech V.Smile']        = '_Default'
$alias['WoW Action Max']       = '_Default'
$alias['Favorites']            = '_Default'
$alias['Ramber Pegasus']       = '_Default'  # unknown clone brand; use default

# SMW collections → map to SNES art
$alias['Super Mario World'] = 'Super Nintendo Entertainment System'
$alias['SMW Adventures']    = 'Super Nintendo Entertainment System'
$alias['SMW Collabs']       = 'Super Nintendo Entertainment System'
$alias['SMW Contests']      = 'Super Nintendo Entertainment System'
$alias['SMW Enhanced']      = 'Super Nintendo Entertainment System'
$alias['SMW General']       = 'Super Nintendo Entertainment System'
$alias['SMW Imitations']    = 'Super Nintendo Entertainment System'
$alias['SMW Japanese']      = 'Super Nintendo Entertainment System'
$alias['SMW Jokes']         = 'Super Nintendo Entertainment System'
$alias['SMW Kaizo']         = 'Super Nintendo Entertainment System'
$alias['SMW Puzzle']        = 'Super Nintendo Entertainment System'
$alias['SMW Thrash']        = 'Super Nintendo Entertainment System'
$alias['SMW Unknown']       = 'Super Nintendo Entertainment System'

# Arcade-alikes
$alias['NESiCAxLive'] = 'Arcade'
$alias['OpenBOR']     = 'Arcade'
$alias['Daphne']      = 'Arcade'
$alias['DICE']        = 'Arcade'
$alias['Boomer Shooters'] = 'Arcade'

# Computers → often prefer Windows fallback unless you have bespoke art
$alias['Microsoft MS-DOS'] = 'Windows'
$alias['PC Games']         = 'Windows'
$alias['PopCap']           = 'Windows'
$alias['ScummVM']          = 'Windows'
$alias['Apple II']         = 'Windows'
$alias['Apple IIGS']       = 'Windows'
$alias['BBC Microcomputer System'] = 'Windows'

# Handheld odds
$alias['Nintendo Dreamcast VMU'] = 'Sega Dreamcast'   # typo safety
$alias['Sega Dreamcast VMU']     = 'Sega Dreamcast'

# ---------------------------------------------------------------------
# Regex rules to normalize noisy names (run before alias lookup)
# ---------------------------------------------------------------------
$regexRules = @(
  @{ Pattern='(?i)\(Hacks?\)';                 Target='' },  # strip (Hacks)
  @{ Pattern='(?i)\b(Imports?|Indies?)\b';     Target='' },  # strip extra labels
  @{ Pattern='(?i)\b(Japan|Asia|Europe|USA)\b';Target='' },  # strip region words
  @{ Pattern='(?i)^Nintendo Switch (E-?Shop)'; Target='Nintendo Switch' },
  @{ Pattern='(?i)^Nintendo Wii U Virtual Console'; Target='Nintendo Wii U' },
  @{ Pattern='(?i)^Nintendo Wii Virtual Console';  Target='Nintendo Wii' }
)

function Try-RegexCollapse {
  param([string]$name)
  $collapsed = $name
  foreach ($rule in $regexRules) {
    if ($collapsed -match $rule.Pattern) {
      if ($rule.Target) {
        return $rule.Target
      } else {
        $collapsed = ($collapsed -replace $rule.Pattern,'').Trim()
        $collapsed = $collapsed -replace '\s{2,}',' '
        $collapsed = $collapsed -replace '\(\s*\)',''
      }
    }
  }
  return $collapsed
}

function Is-Computerish {
  param([string]$name)
  return ($name -match '(?i)Windows|MS-?DOS|Apple|Acorn|Amstrad|BBC|Commodore|VIC-20|PET|CDTV|MAX Machine|Atari ST|Electron|Archimedes|Apog(e|ey)y BK|Camputers Lynx|Dragon|Elektronika BK|Exidy Sorcerer|FM-7|Jupiter Ace|Aquarius|MTX512|MSX|PC-?88|PC-?9801|Sharp X68000|TRS-80|Oric|Sinclair|Spectrum|SAM Coup|ZX|Vector-06C|CreatiVision|Socrates|V\.Smile')
}

# ---------------------------------------------------------------------
# Resolve a source PNG for a platform
# ---------------------------------------------------------------------
function Find-Artwork {
  param([string]$Platform)

  # try collapsed name then alias
  $c = Try-RegexCollapse $Platform
  $tryNames = @($Platform)
  if ($c -ne $Platform) { $tryNames += $c }
  foreach ($n in @($Platform,$c)) {
    if ($alias.ContainsKey($n)) { $tryNames += $alias[$n] }
  }
  $tryNames = $tryNames | Where-Object { $_ } | Select-Object -Unique

  foreach ($n in $tryNames) {
    switch ($n) {
      '_Default' { if (Test-Path $DefaultPng) { return $DefaultPng } }
      'Windows'  { if (Test-Path $WindowsPng) { return $WindowsPng } }
      default {
        $p = Join-Path $ArtworkSourceDir ("{0}.png" -f (Sanitize-FileName $n))
        if (Test-Path $p) { return $p }
      }
    }
  }

  if (Is-Computerish $Platform) {
    if (Test-Path $WindowsPng) { return $WindowsPng }
  }
  if (Test-Path $DefaultPng) { return $DefaultPng }
  return $null
}

# ---------------------------------------------------------------------
# Do the work
# ---------------------------------------------------------------------
$results = @()
foreach ($plat in $platforms) {
  $dest = Join-Path $OutputDir ("{0}.png" -f (Sanitize-FileName $plat))
  if (Test-Path $dest) {
    $results += [pscustomobject]@{Platform=$plat;Action='Skip';Reason='Already exists';Source='';Target=$dest}
    continue
  }
  $src = Find-Artwork $plat
  if ($src) {
    $results += [pscustomobject]@{Platform=$plat;Action='Copy';Reason='Resolved';Source=$src;Target=$dest}
    if (-not $DryRun) { Copy-Item -LiteralPath $src -Destination $dest -Force }
  } else {
    $results += [pscustomobject]@{Platform=$plat;Action='Skip';Reason='No art or fallback';Source='';Target=$dest}
  }
}

# Summary + log
$copied = $results | Where-Object {$_.Action -eq 'Copy'}
$skipped = $results | Where-Object {$_.Action -ne 'Copy'}
Write-Host "Total: $($results.Count)  Copied: $($copied.Count)  Skipped: $($skipped.Count)"
$log = Join-Path $OutputDir ("populate-art_{0:yyyyMMdd_HHmmss}.csv" -f (Get-Date))
$results | Export-Csv -NoTypeInformation -LiteralPath $log
Write-Host "Log saved to $log"
if ($DryRun) { Write-Host "Dry-run only. Re-run without -DryRun to apply." }
