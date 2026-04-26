#Requires -Version 5.1
<#
.SYNOPSIS
    Restructures the drtechnolust Scripts repo into a clean language-first,
    platform-second folder layout.

.DESCRIPTION
    Creates the full recommended folder structure, moves every known script to
    its correct home (renaming to the canonical filename where appropriate),
    archives remaining superseded scripts, pulls scripts out of .vscode, and
    generates a README stub in every new folder.

    Handles filenames with spaces OR underscores automatically -- no manual
    renaming needed before running.

    Safe defaults:
      - DryRun = $true (no files are moved until you change this)
      - Nothing is deleted -- archive moves only
      - CSV log written for every action

    Run order:
      1. .\Invoke-RepoRestructure.ps1          (dry run, review output)
      2. Open script, set DryRun = $false
      3. .\Invoke-RepoRestructure.ps1          (live run)

.VERSION
    1.0.0 - Initial release.

.LICENSE
    MIT License
    Copyright (c) Paul Mardis
#>

# ==============================================================================
# CONFIG
# ==============================================================================
$ScriptRoot = $PSScriptRoot
$DryRun     = $false     # Set to $false to apply changes

$LogDir     = Join-Path $ScriptRoot "Logs"
$LogFile    = Join-Path $LogDir ("RepoRestructure_{0}.csv" -f (Get-Date -Format "yyyyMMdd_HHmmss"))
# ==============================================================================

function Set-ExitCode { param([int]$Code) $global:LASTEXITCODE = $Code }

# ------------------------------------------------------------------------------
# Fuzzy file resolver: tries exact name, underscores-as-spaces, spaces-as-underscores
# ------------------------------------------------------------------------------
function Resolve-SourceFile {
    param(
        [string]$Name,
        [string]$Root
    )

    $candidates = @(
        (Join-Path $Root $Name),
        (Join-Path $Root ($Name -replace '_', ' ')),
        (Join-Path $Root ($Name -replace ' ', '_'))
    )

    foreach ($c in $candidates) {
        if (Test-Path -LiteralPath $c) { return $c }
    }

    return $null
}

# ------------------------------------------------------------------------------
# Create a folder if it does not exist
# ------------------------------------------------------------------------------
function New-SafeDir {
    param([string]$Path)

    if (Test-Path -LiteralPath $Path) { return }

    if ($DryRun) {
        Write-Host "  [DIR] Would create : $Path" -ForegroundColor DarkCyan
    }
    else {
        New-Item -ItemType Directory -Path $Path -Force | Out-Null
        Write-Host "  [DIR] Created      : $Path" -ForegroundColor DarkCyan
    }
}

# ------------------------------------------------------------------------------
# CSV log
# ------------------------------------------------------------------------------
$LogRows = [System.Collections.Generic.List[PSCustomObject]]::new()

function Write-LogRow {
    param(
        [string]$Source,
        [string]$Dest,
        [string]$Action,
        [string]$Status,
        [string]$Note
    )
    $LogRows.Add([PSCustomObject]@{
        Timestamp  = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
        Source     = $Source
        Dest       = $Dest
        Action     = $Action
        Status     = $Status
        Note       = $Note
    })
}

# ------------------------------------------------------------------------------
# Move one file (or dry-run report it)
# ------------------------------------------------------------------------------
$Counters = @{ Moved = 0; Skipped = 0; Missing = 0; Failed = 0 }

function Invoke-FileMove {
    param(
        [string]$SourceRelative,       # relative to ScriptRoot, fuzzy-matched
        [string]$DestRelative,         # relative to ScriptRoot, exact target path
        [string]$SearchRoot = $ScriptRoot,
        [string]$Note = ""
    )

    $ResolvedSource = Resolve-SourceFile -Name $SourceRelative -Root $SearchRoot
    $DestFull       = Join-Path $ScriptRoot $DestRelative

    if ($null -eq $ResolvedSource) {
        Write-Host "  NOT FOUND : $SourceRelative" -ForegroundColor DarkGray
        Write-LogRow -Source (Join-Path $SearchRoot $SourceRelative) -Dest $DestFull `
            -Action "Move" -Status "Missing" -Note "Source not found (may already be moved)"
        $Counters.Missing++
        return
    }

    if (Test-Path -LiteralPath $DestFull) {
        Write-Host "  SKIP      : $(Split-Path $DestFull -Leaf) (destination already exists)" -ForegroundColor Yellow
        Write-LogRow -Source $ResolvedSource -Dest $DestFull `
            -Action "Move" -Status "Skipped" -Note "Destination already exists"
        $Counters.Skipped++
        return
    }

    # Ensure destination folder exists
    New-SafeDir -Path (Split-Path $DestFull -Parent)

    if ($DryRun) {
        Write-Host ("  WOULD MOVE: {0,-60} -> {1}" -f (Split-Path $ResolvedSource -Leaf), $DestRelative) -ForegroundColor Cyan
        Write-LogRow -Source $ResolvedSource -Dest $DestFull -Action "Move" -Status "DryRun" -Note $Note
        $Counters.Moved++
        return
    }

    try {
        Move-Item -LiteralPath $ResolvedSource -Destination $DestFull -ErrorAction Stop
        Write-Host ("  Moved     : {0,-60} -> {1}" -f (Split-Path $ResolvedSource -Leaf), $DestRelative) -ForegroundColor Green
        Write-LogRow -Source $ResolvedSource -Dest $DestFull -Action "Move" -Status "Success" -Note $Note
        $Counters.Moved++
    }
    catch {
        Write-Host "  FAILED    : $(Split-Path $ResolvedSource -Leaf) -- $($_.Exception.Message)" -ForegroundColor Red
        Write-LogRow -Source $ResolvedSource -Dest $DestFull -Action "Move" -Status "Failed" -Note $_.Exception.Message
        $Counters.Failed++
    }
}

# ------------------------------------------------------------------------------
# Write a README stub into a new folder
# ------------------------------------------------------------------------------
function Write-ReadmeStub {
    param([string]$FolderPath, [string]$Title, [string]$Body)

    $ReadmePath = Join-Path $FolderPath "README.md"
    if (Test-Path -LiteralPath $ReadmePath) { return }

    $Content = "# $Title`r`n`r`n$Body`r`n"

    if ($DryRun) { return }

    try {
        $Utf8NoBom = New-Object System.Text.UTF8Encoding($false)
        [System.IO.File]::WriteAllText($ReadmePath, $Content, $Utf8NoBom)
    }
    catch { }
}

# ==============================================================================
# BANNER
# ==============================================================================
Write-Host ""
Write-Host "Repo Restructure" -ForegroundColor Cyan
Write-Host "================" -ForegroundColor Cyan
Write-Host "  Root   : $ScriptRoot"
Write-Host "  DryRun : $DryRun"
Write-Host ""
if ($DryRun) {
    Write-Host "  [DRY RUN] No files will be moved." -ForegroundColor Yellow
    Write-Host ""
}

# ==============================================================================
# STEP 1 -- Create folder structure
# ==============================================================================
Write-Host "Creating folder structure..." -ForegroundColor White

$NewFolders = @(
    "PowerShell\PS3-RPCS3"
    "PowerShell\PS4-ShadPS4"
    "PowerShell\Xbox"
    "PowerShell\Nintendo-Switch"
    "PowerShell\PC-Games"
    "PowerShell\CoinOPS"
    "PowerShell\LaunchBox"
    "PowerShell\ROM-Cleanup"
    "PowerShell\Archive-Tools"
    "PowerShell\TeknoParrot"
    "PowerShell\Paperless-NGX"
    "PowerShell\Enterprise-IT"
    "PowerShell\Modules"
    "Python\PS4"
    "Python\Switch"
    "Python\TeknoParrot"
    "Python\PDF"
    "Python\LaunchBox"
    "Python\Utilities"
    "LaunchBox-Plugins"
    "AutoHotKey\ShadPS4"
    "VSCode"
    "Templates"
    "Docs"
    "Archive"
    "Logs"
)

foreach ($Folder in $NewFolders) {
    New-SafeDir -Path (Join-Path $ScriptRoot $Folder)
}

# ==============================================================================
# STEP 2 -- Archive remaining superseded scripts (v1 NOT FOUND items)
# ==============================================================================
Write-Host ""
Write-Host "Archiving remaining superseded scripts..." -ForegroundColor White

$ArchiveMoves = @(
    # PC Shortcut family
    @{ Src = "pcgames_2_chatgpt_good.ps1";                                        Dest = "Archive\pcgames_2_chatgpt_good.ps1" }
    @{ Src = "Shortcut_Creator_V12_with_debugging.ps1";                           Dest = "Archive\Shortcut_Creator_V12_with_debugging.ps1" }
    @{ Src = "Shortcut_Creator_V12_with_debugging_and_platformnames.ps1";         Dest = "Archive\Shortcut_Creator_V12_with_debugging_and_platformnames.ps1" }
    @{ Src = "Shortcut_Creator_V12_with_debugging_and_platformnames_v2.ps1";      Dest = "Archive\Shortcut_Creator_V12_with_debugging_and_platformnames_v2.ps1" }

    # CoinOPS symlink family
    @{ Src = "CoinOPS_Symlink_ROMs_Videos.ps1";         Dest = "Archive\CoinOPS_Symlink_ROMs_Videos.ps1" }
    @{ Src = "CoinOPS_Symlink_ROMs_Videos_Final.ps1";   Dest = "Archive\CoinOPS_Symlink_ROMs_Videos_Final.ps1" }

    # LaunchBox image copier v1
    @{ Src = "launchbox_copy_images_to_coinops.ps1";    Dest = "Archive\launchbox_copy_images_to_coinops.ps1" }

    # PS3 M3U v1
    @{ Src = "M3U_Creator_Eboot_Files.ps1";             Dest = "Archive\M3U_Creator_Eboot_Files.ps1" }

    # RPCS3 installer v1
    @{ Src = "Unzip_and_Install_PS3_Games.ps1";         Dest = "Archive\Unzip_and_Install_PS3_Games.ps1" }

    # PS4 shortcut creator v1
    @{ Src = "Sony_PS4_CUSA_Shortcut_Creater.ps1";      Dest = "Archive\Sony_PS4_CUSA_Shortcut_Creater.ps1" }

    # PS3 deduplicator v1 (superseded by file-size-aware version)
    @{ Src = "PS3_Region_DeDuplicate_Organizer_v2.ps1"; Dest = "Archive\PS3_Region_DeDuplicate_Organizer_v2.ps1" }

    # ShadPS4 single-type installer
    @{ Src = "shadps4_pkg_extractor_bulk_installer.ps1"; Dest = "Archive\shadps4_pkg_extractor_bulk_installer.ps1" }

    # Switch organizer family
    @{ Src = "Switch_ROM_Organizer.ps1";                Dest = "Archive\Switch_ROM_Organizer.ps1" }
    @{ Src = "Switch_ROM_Organizer_Mover.ps1";          Dest = "Archive\Switch_ROM_Organizer_Mover.ps1" }

    # Switch renamer family
    @{ Src = "Switch_Renamer.ps1";                      Dest = "Archive\Switch_Renamer.ps1" }
    @{ Src = "Switch_Renamer_to_clean_file_names.ps1";  Dest = "Archive\Switch_Renamer_to_clean_file_names.ps1" }
    @{ Src = "Switch_Duplicate_Mover.ps1";              Dest = "Archive\Switch_Duplicate_Mover.ps1" }

    # One-time task
    @{ Src = "Switch_Organizer_-_Long_Names_Only.ps1";  Dest = "Archive\Switch_Organizer_-_Long_Names_Only.ps1" }

    # Paperless Tags v1
    @{ Src = "PaperlessNGX_-_Tags.ps1";                 Dest = "Archive\PaperlessNGX_-_Tags.ps1" }

    # One-liner
    @{ Src = "Folder_List_On_Screen.ps1";               Dest = "Archive\Folder_List_On_Screen.ps1" }
)

foreach ($M in $ArchiveMoves) {
    Invoke-FileMove -SourceRelative $M.Src -DestRelative $M.Dest -Note "Superseded -- archive"
}

# ==============================================================================
# STEP 3 -- Move PowerShell scripts to platform folders (with rename)
# ==============================================================================
Write-Host ""
Write-Host "Moving PowerShell scripts..." -ForegroundColor White

$PSMoves = @(

    # --- PS3 / RPCS3 ---
    @{ Src = "RPCS3-Batch-Installer.ps1";
       Dest = "PowerShell\PS3-RPCS3\RPCS3-Batch-Installer.ps1" }

    @{ Src = "Create-RPCS3-Shortcuts-HD-DISC.ps1";
       Dest = "PowerShell\PS3-RPCS3\Create-RPCS3-Shortcuts-HD-DISC.ps1" }

    @{ Src = "PS3_Region_DeDuplicate_Organizer_w_file_size_v2.ps1";
       Dest = "PowerShell\PS3-RPCS3\PS3-Region-Deduplicator.ps1" }

    @{ Src = "ps3_disc_hdd_deduplicator.ps1";
       Dest = "PowerShell\PS3-RPCS3\PS3-Disc-HDD-Deduplicator.ps1" }

    @{ Src = "M3U_Creator_Eboot_Files_V2.ps1";
       Dest = "PowerShell\PS3-RPCS3\PS3-M3U-Generator.ps1" }

    @{ Src = "param_ps3_tester_ps3.ps1";
       Dest = "PowerShell\PS3-RPCS3\Debug-ParamSFO.ps1" }

    # --- PS4 / ShadPS4 ---
    @{ Src = "Create-ShadPS4-Shortcuts.ps1";
       Dest = "PowerShell\PS4-ShadPS4\Create-ShadPS4-Shortcuts.ps1" }

    @{ Src = "Sony_PS4_CUSA_Shortcut_Creater_v2.ps1";
       Dest = "PowerShell\PS4-ShadPS4\Sony_PS4_CUSA_Shortcut_Creater_v2.ps1" }

    @{ Src = "shadps4_pkg_extractor_bulk_installer_mixedmode.ps1";
       Dest = "PowerShell\PS4-ShadPS4\ShadPS4-PKG-Installer.ps1" }

    @{ Src = "PS4_-_Folder_Renamer_Sentence_Case.ps1";
       Dest = "PowerShell\PS4-ShadPS4\PS4-Folder-Renamer.ps1" }

    @{ Src = "Sony_PS4_File_Sorter.ps1";
       Dest = "PowerShell\PS4-ShadPS4\PS4-File-Sorter.ps1" }

    # --- Xbox ---
    @{ Src = "Microsoft_Xbox_Region_Dedupe.ps1";
       Dest = "PowerShell\Xbox\Xbox-Region-Deduplicator.ps1" }

    @{ Src = "Audits_the_Xbox_duplicate_log_CSV.ps1";
       Dest = "PowerShell\Xbox\Audit-Xbox-Dedupe-CSV.ps1" }

    # --- Nintendo Switch ---
    @{ Src = "SwitchOrganizerGood.ps1";
       Dest = "PowerShell\Nintendo-Switch\Switch-ROM-Organizer.ps1" }

    @{ Src = "Cleans_up_nintendo_switch_folder_file_names.ps1";
       Dest = "PowerShell\Nintendo-Switch\Switch-Name-Cleaner.ps1" }

    # --- PC Games ---
    @{ Src = "Shortcut_Creator_V12_with_debugging_and_platformnames_v3_ownership.ps1";
       Dest = "PowerShell\PC-Games\Create-PC-Shortcuts.ps1" }

    @{ Src = "ShortcutCreater_Batch_Files.ps1";
       Dest = "PowerShell\PC-Games\Create-Batch-Launchers.ps1" }

    @{ Src = "GOGshortcutcreator.ps1";
       Dest = "PowerShell\PC-Games\Create-GOG-Shortcuts.ps1" }

    @{ Src = "OpenBor_Shortcut_creator.ps1";
       Dest = "PowerShell\PC-Games\Create-OpenBOR-Shortcuts.ps1" }

    @{ Src = "pcshortcutschangepaths.ps1";
       Dest = "PowerShell\PC-Games\Update-Shortcut-Paths.ps1" }

    @{ Src = "PC_Games_Script.ps1";
       Dest = "PowerShell\PC-Games\Export-LaunchBox-CSV.ps1" }

    # --- CoinOPS ---
    @{ Src = "CoinOPS_Symlink_ROMs_Videos_Final_v2.ps1";
       Dest = "PowerShell\CoinOPS\Create-CoinOPS-Symlinks.ps1" }

    @{ Src = "CoinOps_Launch_Config_Updater_retroarch.ps1";
       Dest = "PowerShell\CoinOPS\Update-CoinOPS-LaunchConfigs.ps1" }

    @{ Src = "Copies_CoinOPS_Nostalgic_Roomassets.ps1";
       Dest = "PowerShell\CoinOPS\Copy-CoinOPS-NostalgicRoom-Assets.ps1" }

    @{ Src = "launchbox_copy_images_to_coinops_V2.ps1";
       Dest = "PowerShell\CoinOPS\Copy-LaunchBox-Images-to-CoinOPS.ps1" }

    # --- LaunchBox ---
    @{ Src = "LaunchBox_Media_Renamer.ps1";
       Dest = "PowerShell\LaunchBox\Rename-LaunchBox-Media.ps1" }

    @{ Src = "Platform_artwork_name_to_Launchbox_scheme.ps1";
       Dest = "PowerShell\LaunchBox\Map-Platform-Artwork-Names.ps1" }

    # --- ROM Cleanup ---
    @{ Src = "GameBoy_Color_Renamer.ps1";
       Dest = "PowerShell\ROM-Cleanup\ROM-Renamer.ps1" }

    @{ Src = "Region_Renamer.ps1";
       Dest = "PowerShell\ROM-Cleanup\Region-Code-Renamer.ps1" }

    # --- Archive Tools ---
    @{ Src = "1_7zip_Extractor_and_Mover.ps1";
       Dest = "PowerShell\Archive-Tools\Extract-Archives.ps1" }

    @{ Src = "RARandINNOExtracter.ps1";
       Dest = "PowerShell\Archive-Tools\Extract-Archives-Parallel.ps1" }

    @{ Src = "Zip_Verification.ps1";
       Dest = "PowerShell\Archive-Tools\Check-Archive-Extraction.ps1" }

    # --- TeknoParrot ---
    @{ Src = "Tekno_Mover.ps1";
       Dest = "PowerShell\TeknoParrot\Move-TeknoParrot-Games.ps1" }

    # --- Paperless-NGX ---
    # NOTE: Use the FIXED versions (token removed) from the Scripts\Scripts subfolder
    # if they are present there -- otherwise moves the root version.
    # After this run, paste your new API token into each file's CONFIG block.
    @{ Src = "PaperlessNGX_-_Categories.ps1";
       Dest = "PowerShell\Paperless-NGX\Sync-Paperless-Categories.ps1" }

    @{ Src = "PaperlessNGX_-_Document_Type.ps1";
       Dest = "PowerShell\Paperless-NGX\Sync-Paperless-DocumentTypes.ps1" }

    @{ Src = "PaperlessNGX_-_Tags_2.ps1";
       Dest = "PowerShell\Paperless-NGX\Sync-Paperless-Tags.ps1" }

    # --- Enterprise IT ---
    @{ Src = "SecureBootProd.ps1";
       Dest = "PowerShell\Enterprise-IT\Detect-SecureBootCertUpdate.ps1" }

    @{ Src = "SecureBootProdRemediation.ps1";
       Dest = "PowerShell\Enterprise-IT\Remediate-SecureBootCertUpdate.ps1" }
)

foreach ($M in $PSMoves) {
    Invoke-FileMove -SourceRelative $M.Src -DestRelative $M.Dest
}

# ==============================================================================
# STEP 4 -- Fixed Paperless scripts from Scripts\Scripts subfolder
# Move the token-safe versions if they exist there, then clean up the subfolder
# ==============================================================================
Write-Host ""
Write-Host "Checking Scripts\Scripts subfolder for fixed Paperless scripts..." -ForegroundColor White

$SubFolder = Join-Path $ScriptRoot "Scripts"
if (Test-Path -LiteralPath $SubFolder) {
    $PaperlessFixes = @(
        @{ Src = "Scripts\PaperlessNGX_-_Categories.ps1";
           Dest = "PowerShell\Paperless-NGX\Sync-Paperless-Categories.ps1" }
        @{ Src = "Scripts\PaperlessNGX_-_Document_Type.ps1";
           Dest = "PowerShell\Paperless-NGX\Sync-Paperless-DocumentTypes.ps1" }
        @{ Src = "Scripts\PaperlessNGX_-_Tags_2.ps1";
           Dest = "PowerShell\Paperless-NGX\Sync-Paperless-Tags.ps1" }
    )
    foreach ($M in $PaperlessFixes) {
        Invoke-FileMove -SourceRelative $M.Src -DestRelative $M.Dest `
            -Note "Token-safe version from Scripts subfolder -- preferred over root version"
    }
}
else {
    Write-Host "  Scripts subfolder not found -- skipping" -ForegroundColor DarkGray
}

# ==============================================================================
# STEP 5 -- Move Python scripts
# ==============================================================================
Write-Host ""
Write-Host "Moving Python scripts..." -ForegroundColor White

$PyMoves = @(
    @{ Src = "PS4_Organizer.py";                        Dest = "Python\PS4\PS4_Organizer.py" }
    @{ Src = "Switch_ROM_Renamer_to_Clean_Names.py";    Dest = "Python\Switch\Switch_ROM_Renamer.py" }
    @{ Src = "Switch_ROM_Renamer_to_Clean_Names_v2.py"; Dest = "Python\Switch\Switch_ROM_Renamer_v2.py" }
    @{ Src = "TeknoparrotNameCleanerLB.py";             Dest = "Python\TeknoParrot\TeknoparrotNameCleanerLB.py" }
    @{ Src = "TeknoparrotNameCleanerLBUserinput.py";    Dest = "Python\TeknoParrot\TeknoparrotNameCleanerLBUserinput.py" }
    @{ Src = "PDFSecurityScanner.py";                   Dest = "Python\PDF\PDFSecurityScanner.py" }
    @{ Src = "PDFSecurityScannerV2.py";                 Dest = "Python\PDF\PDFSecurityScannerV2.py" }
    @{ Src = "LBPathExtraction.py";                     Dest = "Python\LaunchBox\LBPathExtraction.py" }
    @{ Src = "extract_rom_path.py";                     Dest = "Python\Utilities\extract_rom_path.py" }
    @{ Src = "1file_webscraper.py";                     Dest = "Python\Utilities\1file_webscraper.py" }
)

foreach ($M in $PyMoves) {
    Invoke-FileMove -SourceRelative $M.Src -DestRelative $M.Dest
}

# ==============================================================================
# STEP 6 -- Move AutoHotKey scripts
# ==============================================================================
Write-Host ""
Write-Host "Moving AutoHotKey scripts..." -ForegroundColor White

Invoke-FileMove -SourceRelative "shadps4_auto_ok_logger.ahk" `
                -DestRelative   "AutoHotKey\ShadPS4\shadps4_auto_ok_logger.ahk"

# ==============================================================================
# STEP 7 -- Pull scripts out of .vscode (they do not belong there)
# ==============================================================================
Write-Host ""
Write-Host "Cleaning up .vscode folder..." -ForegroundColor White

$VsCodeRoot = Join-Path $ScriptRoot ".vscode"

if (Test-Path -LiteralPath $VsCodeRoot) {

    # ImageCleanup -- LaunchBox image processing, needs proper assessment
    Invoke-FileMove -SourceRelative "ImageCleanup_v4.1.ps1" `
                    -DestRelative   "PowerShell\LaunchBox\ImageCleanup_v4.1.ps1" `
                    -SearchRoot     $VsCodeRoot `
                    -Note           "Pulled from .vscode -- needs review and rename"

    # RPCS3-Batch-Installer-Final-version -- check if newer than root version; keeping for comparison
    Invoke-FileMove -SourceRelative "RPCS3-Batch-Installer-Final-version.ps1" `
                    -DestRelative   "PowerShell\PS3-RPCS3\RPCS3-Batch-Installer-Final-version.ps1" `
                    -SearchRoot     $VsCodeRoot `
                    -Note           "Pulled from .vscode -- compare with RPCS3-Batch-Installer.ps1 and archive whichever is older"

    # All Install-RPCS3Pkgs variants -- superseded by RPCS3-Batch-Installer, archive them
    $InstallVariants = @(
        "Install-RPCS3Pkgs.ps1"
        "Install-RPCS3Pkgs(1).ps1"
        "Install-RPCS3Pkgs-working.ps1"
        "Install-RPCS3Pkgs-working2.ps1"
    )

    foreach ($Variant in $InstallVariants) {
        $VsCodeSource = Join-Path $VsCodeRoot $Variant
        if (Test-Path -LiteralPath $VsCodeSource) {
            Invoke-FileMove -SourceRelative $Variant `
                            -DestRelative   "Archive\$Variant" `
                            -SearchRoot     $VsCodeRoot `
                            -Note           "Superseded by RPCS3-Batch-Installer.ps1 -- archived from .vscode"
        }
    }

    # Install-RPCS3Pkgs-working3 has a truncated name in the screenshot -- match by wildcard
    $Working3 = Get-ChildItem -LiteralPath $VsCodeRoot -Filter "Install-RPCS3Pkgs-working3*" -File -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($Working3) {
        $ArchiveDest = Join-Path $ScriptRoot "Archive" $Working3.Name
        if (-not (Test-Path -LiteralPath $ArchiveDest)) {
            if ($DryRun) {
                Write-Host ("  WOULD MOVE: {0} -> Archive\" -f $Working3.Name) -ForegroundColor Cyan
                Write-LogRow -Source $Working3.FullName -Dest $ArchiveDest -Action "Move" -Status "DryRun" `
                    -Note "Superseded by RPCS3-Batch-Installer.ps1 -- archived from .vscode"
                $Counters.Moved++
            }
            else {
                try {
                    Move-Item -LiteralPath $Working3.FullName -Destination $ArchiveDest -ErrorAction Stop
                    Write-Host ("  Moved     : {0} -> Archive\" -f $Working3.Name) -ForegroundColor Green
                    Write-LogRow -Source $Working3.FullName -Dest $ArchiveDest -Action "Move" -Status "Success" `
                        -Note "Superseded -- archived from .vscode"
                    $Counters.Moved++
                }
                catch {
                    Write-Host "  FAILED    : $($Working3.Name) -- $($_.Exception.Message)" -ForegroundColor Red
                    $Counters.Failed++
                }
            }
        }
    }

    # Duplicate RPCS3-Batch-Installer in .vscode -- archive it (root version is canonical)
    $VsCodeBatchInstaller = Join-Path $VsCodeRoot "RPCS3-Batch-Installer.ps1"
    if (Test-Path -LiteralPath $VsCodeBatchInstaller) {
        Invoke-FileMove -SourceRelative "RPCS3-Batch-Installer.ps1" `
                        -DestRelative   "Archive\RPCS3-Batch-Installer-vscode-copy.ps1" `
                        -SearchRoot     $VsCodeRoot `
                        -Note           "Duplicate of root version -- archived from .vscode"
    }

    # Move launch.json to VSCode\ (that is the correct home)
    Invoke-FileMove -SourceRelative "launch.json" `
                    -DestRelative   "VSCode\launch.json" `
                    -SearchRoot     $VsCodeRoot `
                    -Note           "VS Code workspace config"
}
else {
    Write-Host "  .vscode folder not found -- skipping" -ForegroundColor DarkGray
}

# ==============================================================================
# STEP 8 -- Move the Docs reports to Docs\
# ==============================================================================
Write-Host ""
Write-Host "Moving audit reports to Docs..." -ForegroundColor White

$DocMoves = @(
    "SCRIPT_LIBRARY_AUDIT.md"
    "SCRIPT_INDEX.md"
    "SCRIPT_BUG_REPORT.md"
    "SCRIPT_CONSOLIDATION_PLAN.md"
    "SCRIPT_REPO_RESTRUCTURE_PLAN.md"
)

foreach ($Doc in $DocMoves) {
    Invoke-FileMove -SourceRelative $Doc -DestRelative "Docs\$Doc"
}

# ==============================================================================
# STEP 9 -- Generate README stubs in new folders
# ==============================================================================
if (-not $DryRun) {
    Write-Host ""
    Write-Host "Writing README stubs..." -ForegroundColor White

    $ReadmeMap = @(
        @{ Path = "PowerShell\PS3-RPCS3";        Title = "PS3 / RPCS3";           Body = "PowerShell scripts for PS3 library management, RPCS3 shortcut creation, PKG installation, deduplication, and M3U generation." }
        @{ Path = "PowerShell\PS4-ShadPS4";      Title = "PS4 / ShadPS4";         Body = "PowerShell scripts for PS4 game management, ShadPS4 shortcut creation, PKG extraction, and folder organization." }
        @{ Path = "PowerShell\Xbox";             Title = "Xbox";                   Body = "PowerShell scripts for Xbox ISO regional deduplication and audit reporting." }
        @{ Path = "PowerShell\Nintendo-Switch";  Title = "Nintendo Switch";        Body = "PowerShell scripts for Switch ROM organization and filename normalization." }
        @{ Path = "PowerShell\PC-Games";         Title = "PC Games";               Body = "PowerShell scripts for PC game shortcut creation, executable scoring, and LaunchBox CSV export." }
        @{ Path = "PowerShell\CoinOPS";          Title = "CoinOPS";                Body = "PowerShell scripts for CoinOPS symlink creation, config updates, and asset copying." }
        @{ Path = "PowerShell\LaunchBox";        Title = "LaunchBox";              Body = "PowerShell scripts for LaunchBox media renaming and platform artwork mapping." }
        @{ Path = "PowerShell\ROM-Cleanup";      Title = "ROM Cleanup";            Body = "PowerShell scripts for general ROM file renaming and region code normalization." }
        @{ Path = "PowerShell\Archive-Tools";    Title = "Archive Tools";          Body = "PowerShell scripts for 7-Zip extraction, parallel extraction, and archive verification." }
        @{ Path = "PowerShell\TeknoParrot";      Title = "TeknoParrot";            Body = "PowerShell scripts for TeknoParrot game folder management." }
        @{ Path = "PowerShell\Paperless-NGX";    Title = "Paperless-NGX";         Body = "PowerShell scripts for syncing tags, document types, and custom fields to Paperless-NGX via its REST API.`n`n> **Security:** Add your API token to the `$Token` field in each script's CONFIG block. Never commit a live token." }
        @{ Path = "PowerShell\Enterprise-IT";    Title = "Enterprise IT";          Body = "PowerShell scripts for enterprise IT management (Secure Boot certificate updates, Intune remediation)." }
        @{ Path = "PowerShell\Modules";          Title = "PowerShell Modules";     Body = "Shared .psm1 helper modules imported by scripts across the library." }
        @{ Path = "Python\PS4";                  Title = "Python -- PS4";          Body = "Python scripts for PS4 library organization and metadata processing." }
        @{ Path = "Python\Switch";               Title = "Python -- Switch";       Body = "Python scripts for Nintendo Switch ROM filename normalization." }
        @{ Path = "Python\TeknoParrot";          Title = "Python -- TeknoParrot";  Body = "Python scripts for TeknoParrot game name cleaning and LaunchBox integration." }
        @{ Path = "Python\PDF";                  Title = "Python -- PDF";          Body = "Python scripts for PDF security scanning and processing." }
        @{ Path = "Python\LaunchBox";            Title = "Python -- LaunchBox";    Body = "Python scripts for LaunchBox data processing and path extraction." }
        @{ Path = "Python\Utilities";            Title = "Python -- Utilities";    Body = "General-purpose Python utility scripts." }
        @{ Path = "LaunchBox-Plugins";           Title = "LaunchBox Plugins";      Body = "C# plugin projects for LaunchBox and BigBox.`n`nLanguage: C# / .NET Framework 4.7.2 (or .NET 6+)`nSDK: LaunchBox.Common.dll + Unbroken.LaunchBox.Windows.dll`nOutput: compiled .dll copied to C:\Arcade\LaunchBox\Plugins\" }
        @{ Path = "AutoHotKey\ShadPS4";          Title = "AutoHotKey -- ShadPS4";  Body = "AutoHotKey scripts for ShadPS4 automation." }
        @{ Path = "Templates";                   Title = "Templates";              Body = "Script templates for PowerShell, Python, and C# LaunchBox plugin projects." }
        @{ Path = "Docs";                        Title = "Documentation";          Body = "Library audit reports, script index, bug report, and restructure plan." }
        @{ Path = "Archive";                     Title = "Archive";                Body = "Superseded, broken, and one-time-use scripts preserved for reference. Nothing here is maintained." }
    )

    foreach ($R in $ReadmeMap) {
        Write-ReadmeStub -FolderPath (Join-Path $ScriptRoot $R.Path) -Title $R.Title -Body $R.Body
    }
}

# ==============================================================================
# STEP 10 -- Write CSV log
# ==============================================================================
if (-not (Test-Path -LiteralPath $LogDir)) {
    New-Item -ItemType Directory -Path $LogDir | Out-Null
}

if ($LogRows.Count -gt 0) {
    try {
        $LogRows | Export-Csv -Path $LogFile -NoTypeInformation -Encoding UTF8
        Write-Host ""
        Write-Host "  Log written: $LogFile" -ForegroundColor Gray
    }
    catch {
        Write-Host "  WARNING: Could not write log -- $($_.Exception.Message)" -ForegroundColor Yellow
    }
}

# ==============================================================================
# SUMMARY
# ==============================================================================
Write-Host ""
Write-Host "Summary" -ForegroundColor White
Write-Host ("  {0,-12}: {1}" -f $(if ($DryRun) { "Would move" } else { "Moved" }), $Counters.Moved)
Write-Host ("  {0,-12}: {1}" -f "Skipped", $Counters.Skipped)     -ForegroundColor Yellow
Write-Host ("  {0,-12}: {1}" -f "Not found", $Counters.Missing)   -ForegroundColor DarkGray
Write-Host ("  {0,-12}: {1}" -f "Failed", $Counters.Failed)       -ForegroundColor $(if ($Counters.Failed -gt 0) { "Red" } else { "Gray" })

if ($DryRun) {
    Write-Host ""
    Write-Host "  DRY RUN complete." -ForegroundColor Yellow
    Write-Host "  Review the output above, then set `$DryRun = `$false and run again." -ForegroundColor Yellow
}
else {
    Write-Host ""
    Write-Host "  Restructure complete." -ForegroundColor Green
    Write-Host "  Review any NOT FOUND items -- they may already be in Archive or moved." -ForegroundColor Gray
    Write-Host ""
    Write-Host "  NEXT STEPS:" -ForegroundColor White
    Write-Host "  1. Paste your Paperless-NGX API token into PowerShell\Paperless-NGX\ scripts" -ForegroundColor Gray
    Write-Host "  2. Compare .vscode\RPCS3-Batch-Installer-Final-version.ps1 with PS3-RPCS3 version" -ForegroundColor Gray
    Write-Host "  3. Review PowerShell\LaunchBox\ImageCleanup_v4.1.ps1 and rename/move as needed" -ForegroundColor Gray
    Write-Host "  4. Review PowerShell\PS4-ShadPS4\ -- two shortcut creators landed there, keep one" -ForegroundColor Gray
    Write-Host "  5. Commit to GitHub with the suggested commit message below" -ForegroundColor Gray
    Write-Host ""
    Write-Host "  Suggested commit message:" -ForegroundColor White
    Write-Host "  chore: restructure repo into language/platform folder layout" -ForegroundColor Cyan
    Write-Host "  Move all PS, Python, AHK scripts to canonical paths." -ForegroundColor Cyan
    Write-Host "  Archive remaining superseded scripts. Pull scripts from .vscode." -ForegroundColor Cyan
    Write-Host "  Add README stubs to all new folders. Per April 2026 audit." -ForegroundColor Cyan
}

Write-Host ""
Set-ExitCode 0