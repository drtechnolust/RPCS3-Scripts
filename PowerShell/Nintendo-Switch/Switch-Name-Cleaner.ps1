#requires -version 5.1
<#
.SYNOPSIS
    Cleans up Nintendo Switch folder and file names.

.DESCRIPTION
    Folders: game title only, no tags.
    Files: game title + version tags preserved, release prefixes stripped.
    Files with meaningless short names fall back to parent folder title.
    Outputs to a separate destination folder.

.NOTES
    Run in PowerShell ISE.
    Prompts for source path, destination path, and dry run mode.
#>

function Get-SafeFileName {
    param(
        [Parameter(Mandatory)]
        [string]$Name
    )

    $invalid = [System.IO.Path]::GetInvalidFileNameChars()
    foreach ($char in $invalid) {
        $Name = $Name.Replace([string]$char, '')
    }

    $Name = $Name.Trim()
    $Name = $Name -replace '\s{2,}', ' '
    return $Name
}

function Resolve-UniquePath {
    param(
        [Parameter(Mandatory)] [string]$DesiredPath
    )

    if (-not (Test-Path -LiteralPath $DesiredPath)) {
        return $DesiredPath
    }

    $dir  = Split-Path -LiteralPath $DesiredPath -Parent
    $leaf = Split-Path -LiteralPath $DesiredPath -Leaf
    $base = [System.IO.Path]::GetFileNameWithoutExtension($leaf)
    $ext  = [System.IO.Path]::GetExtension($leaf)

    $i = 1
    do {
        $candidate = Join-Path $dir ("{0} ({1}){2}" -f $base, $i, $ext)
        $i++
    } until (-not (Test-Path -LiteralPath $candidate))

    return $candidate
}

function Normalize-TitleCase {
    param(
        [Parameter(Mandatory)]
        [string]$Text
    )

    $culture = [System.Globalization.CultureInfo]::CurrentCulture
    return $culture.TextInfo.ToTitleCase($Text.ToLower())
}

function Should-SkipFolder {
    param(
        [Parameter(Mandatory)]
        [string]$FolderName
    )

    if ($FolderName -match '^_') { return $true }
    if ($FolderName -match '(?i)SWITCH\s*\(BACKUP\)') { return $true }
    if ($FolderName -match '(?i)^Nintendo\s+Switch$') { return $true }

    return $false
}

function Clean-SwitchName {
    param(
        [Parameter(Mandatory)]
        [string]$OriginalName,

        [switch]$IsFolder,
        [string]$FallbackTitle = "",
        [switch]$KeepRegion,
        [switch]$KeepLanguageTags,
        [switch]$KeepSceneGroup
    )

    $ext      = [System.IO.Path]::GetExtension($OriginalName)
    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($OriginalName)

    if ([string]::IsNullOrWhiteSpace($ext)) {
        $baseName = $OriginalName
        $ext = ""
    }

    # Strip release group filename prefixes e.g. v-, sxs-, hr-, nsw-
    if (-not $IsFolder) {
        $baseName = $baseName -replace '^(?i)(v|sxs|hr|nsw)-', ''
    }

    # Normalize separators
    $work = $baseName
    $work = $work -replace '[\._]+', ' '
    $work = $work -replace '\s{2,}', ' '
    $work = $work.Trim()

    # Extract metadata before stripping (files only)
    $updateMatch  = [regex]::Match($work, '(?i)\bUpdate\s+v?([0-9]+(?:\.[0-9A-Za-z]+)*)\b')
    $versionMatch = [regex]::Match($work, '(?i)\bv([0-9]+(?:\.[0-9A-Za-z]+)*)\b')
    $dlcUnlocker  = $work -match '(?i)\bDLC\s+Unlocker\b'
    $dlc          = $work -match '(?i)\bDLC\b'
    $updateWord   = $work -match '(?i)\bUpdate\b'

    # Also capture bare numeric version at end e.g. "gb_v1310720" -> v1310720
    $bareVersionMatch = [regex]::Match($work, '(?i)\bv([0-9]{4,})\b')

    $region = $null
    if ($KeepRegion -and -not $IsFolder) {
        $regionMatch = [regex]::Match($work, '(?i)\b(USA|EUR|JPN|JAP|ASIA|KOR|CHN|ENG)\b')
        if ($regionMatch.Success) {
            $region = $regionMatch.Groups[1].Value.ToUpper()
            if ($region -eq 'JAP') { $region = 'JPN' }
        }
    }

    $langTags = @()
    if ($KeepLanguageTags -and -not $IsFolder) {
        foreach ($m in [regex]::Matches($work, '(?i)\b(MULTI|MULTi)\b')) {
            $langTags += 'MULTI'
        }
        $langTags = $langTags | Select-Object -Unique
    }

    $sceneGroup = $null
    if ($KeepSceneGroup -and -not $IsFolder) {
        $sceneMatch = [regex]::Match($work, '(?i)\b(NSW-[A-Z0-9]+)\b')
        if ($sceneMatch.Success) {
            $sceneGroup = $sceneMatch.Groups[1].Value.ToUpper()
        }
    }

    # Strip all noisy tokens
    $patternsToRemove = @(
        '(?i)\bNSW-[A-Z0-9]+\b',
        '(?i)\bNSW\b',
        '(?i)\bSUXXORS\b',
        '(?i)\bVENOM\b',
        '(?i)\bLIGHTFORCE\b',
        '(?i)\bHR\b',
        '(?i)\bINTERNAL\b',
        '(?i)\bPROPER\b',
        '(?i)\bREADNFO\b',
        '(?i)\bREPACK\b',
        '(?i)\bDUMPED\b',
        '(?i)\bEBOOT\b',
        '(?i)\bNFOFIX\b',
        '(?i)\bONELOAD\b',
        '(?i)\bDIRFIX\b',
        '(?i)\bUNLOCKED\b',
        '(?i)\bMULTI\b',
        '(?i)\bMULTi\b',
        '(?i)\bUSA\b',
        '(?i)\bEUR\b',
        '(?i)\bJPN\b',
        '(?i)\bJAP\b',
        '(?i)\bASIA\b',
        '(?i)\bKOR\b',
        '(?i)\bCHN\b',
        '(?i)\bENG\b'
    )

    foreach ($pattern in $patternsToRemove) {
        $work = $work -replace $pattern, ' '
    }

    # Strip version/update/dlc from title area (re-added for files below)
    $work = $work -replace '(?i)\bUpdate\s+v?[0-9]+(?:\.[0-9A-Za-z]+)*\b', ' '
    $work = $work -replace '(?i)\bDLC\s+Unlocker\b', ' '
    $work = $work -replace '(?i)\bDLC\b', ' '
    $work = $work -replace '(?i)\bUpdate\b', ' '
    $work = $work -replace '(?i)\bv[0-9]+(?:\.[0-9A-Za-z]+)*\b', ' '

    # Remove orphaned standalone numbers left by partial version stripping
    $work = $work -replace '(?<!\w)\d+(?!\w)', ' '

    # Clean up
    $work = $work -replace '\s{2,}', ' '
    $work = $work.Trim(' ', '-', '_', '.', '[', ']')

    # FIX: If remaining title is too short to be meaningful, use fallback (parent folder title)
    $titleCandidate = $work -replace '\s', ''
    if ($titleCandidate.Length -le 4 -and $FallbackTitle) {
        $title = $FallbackTitle
    }
    else {
        $title = Normalize-TitleCase -Text $work
        $title = $title -replace '\s{2,}', ' '
        $title = $title.Trim()
    }

    # Folders: title only, no suffixes
    if ($IsFolder) {
        $newName = Get-SafeFileName -Name $title
        return $newName
    }

    # Files: rebuild suffixes
    $suffixes = New-Object System.Collections.Generic.List[string]

    if ($updateMatch.Success) {
        $suffixes.Add("Update v$($updateMatch.Groups[1].Value)")
    }
    elseif ($updateWord -and $versionMatch.Success) {
        $suffixes.Add("Update v$($versionMatch.Groups[1].Value)")
    }
    elseif ($dlcUnlocker) {
        $suffixes.Add("DLC Unlocker")
    }
    elseif ($dlc) {
        $suffixes.Add("DLC")
    }
    elseif ($versionMatch.Success -and -not $updateWord) {
        $suffixes.Add("v$($versionMatch.Groups[1].Value)")
    }
    elseif ($bareVersionMatch.Success) {
        # Catch bare numeric versions like v1310720 that survive stripping
        $suffixes.Add("v$($bareVersionMatch.Groups[1].Value)")
    }

    if ($KeepRegion -and $region) {
        $suffixes.Add($region)
    }
    if ($KeepLanguageTags -and $langTags.Count -gt 0) {
        foreach ($tag in $langTags) { $suffixes.Add($tag) }
    }
    if ($KeepSceneGroup -and $sceneGroup) {
        $suffixes.Add($sceneGroup)
    }

    $newName = $title
    if ($suffixes.Count -gt 0) {
        $newName += " " + ($suffixes -join ' ')
    }

    $newName = Get-SafeFileName -Name $newName
    return $newName + $ext
}

# -----------------------------
# Options
# -----------------------------
$RenameFolders    = $true
$RenameFiles      = $true
$RecurseFiles     = $true
$KeepRegion       = $false
$KeepLanguageTags = $false
$KeepSceneGroup   = $false

# -----------------------------
# Prompts
# -----------------------------
$sourcePath = (Read-Host "Enter source folder path").Trim('"')

if (-not (Test-Path -LiteralPath $sourcePath)) {
    Write-Warning "Source path not found: $sourcePath"
    return
}

$destPath = (Read-Host "Enter destination folder path").Trim('"')

if (-not $destPath) {
    Write-Warning "No destination path provided. Exiting."
    return
}

$dryRunInput = Read-Host "Dry run? (Y/N)"
$DryRun = $dryRunInput.Trim().ToUpper() -eq 'Y'

# -----------------------------
# Setup
# -----------------------------
if (-not $DryRun) {
    if (-not (Test-Path -LiteralPath $destPath)) {
        New-Item -ItemType Directory -Path $destPath -Force | Out-Null
    }
}

$logPath = Join-Path $destPath ("SwitchCleanup_{0}.log" -f (Get-Date -Format "yyyyMMdd_HHmmss"))

if (-not $DryRun) {
    "Switch Cleanup Log - $(Get-Date)" | Out-File -FilePath $logPath -Encoding UTF8
}

Write-Host "`nSource : $sourcePath" -ForegroundColor Cyan
Write-Host "Dest   : $destPath"    -ForegroundColor Cyan
Write-Host "DryRun : $DryRun"      -ForegroundColor Yellow
if ($DryRun) {
    Write-Host "Log    : (dry run - no log written)`n" -ForegroundColor Yellow
}
else {
    Write-Host "Log    : $logPath`n" -ForegroundColor Cyan
}

function Write-Log {
    param([string]$Message)
    if (-not $DryRun) {
        Add-Content -LiteralPath $logPath -Value $Message
    }
}

# -----------------------------
# Copy files with cleaned names
# -----------------------------
if ($RenameFiles) {
    $files = if ($RecurseFiles) {
        Get-ChildItem -LiteralPath $sourcePath -File -Recurse -ErrorAction SilentlyContinue
    }
    else {
        Get-ChildItem -LiteralPath $sourcePath -File -ErrorAction SilentlyContinue
    }

    foreach ($file in $files) {
        # Get clean parent folder title for fallback
        $parentFolderName = Split-Path $file.DirectoryName -Leaf
        $fallbackTitle = if (Should-SkipFolder -FolderName $parentFolderName) {
            ""
        }
        else {
            Clean-SwitchName -OriginalName $parentFolderName -IsFolder `
                -KeepRegion:$KeepRegion `
                -KeepLanguageTags:$KeepLanguageTags `
                -KeepSceneGroup:$KeepSceneGroup
        }

        $newLeaf = Clean-SwitchName -OriginalName $file.Name `
            -FallbackTitle $fallbackTitle `
            -KeepRegion:$KeepRegion `
            -KeepLanguageTags:$KeepLanguageTags `
            -KeepSceneGroup:$KeepSceneGroup

        # Mirror cleaned folder structure under destination
        $relativeDir = $file.DirectoryName.Substring($sourcePath.Length).TrimStart('\', '/')
        $cleanRelDir = if ($relativeDir) {
            ($relativeDir -split '[\\/]' | ForEach-Object {
                $seg = $_
                if (Should-SkipFolder -FolderName $seg) {
                    $seg
                }
                else {
                    Clean-SwitchName -OriginalName $seg -IsFolder `
                        -KeepRegion:$KeepRegion `
                        -KeepLanguageTags:$KeepLanguageTags `
                        -KeepSceneGroup:$KeepSceneGroup
                }
            }) -join [System.IO.Path]::DirectorySeparatorChar
        }
        else { "" }

        $targetDir  = if ($cleanRelDir) { Join-Path $destPath $cleanRelDir } else { $destPath }
        $targetPath = Resolve-UniquePath -DesiredPath (Join-Path $targetDir $newLeaf)

        $msg = "FILE  : `"$($file.FullName)`" -> `"$targetPath`""
        Write-Log $msg

        if ($DryRun) {
            Write-Host "[DRYRUN] $msg" -ForegroundColor DarkYellow
        }
        else {
            try {
                if (-not (Test-Path -LiteralPath $targetDir)) {
                    New-Item -ItemType Directory -Path $targetDir -Force | Out-Null
                }
                Copy-Item -LiteralPath $file.FullName -Destination $targetPath -ErrorAction Stop
                Write-Host "[COPIED] $msg" -ForegroundColor Green
            }
            catch {
                $err = "[ERROR] $($file.FullName) -- $($_.Exception.Message)"
                Write-Log $err
                Write-Warning $err
            }
        }
    }
}

# -----------------------------
# Log folder renames
# -----------------------------
if ($RenameFolders) {
    $folders = Get-ChildItem -LiteralPath $sourcePath -Directory -Recurse -ErrorAction SilentlyContinue |
        Sort-Object { $_.FullName.Length } -Descending

    foreach ($folder in $folders) {
        if (Should-SkipFolder -FolderName $folder.Name) {
            Write-Host "[SKIPPED] `"$($folder.FullName)`"" -ForegroundColor DarkGray
            continue
        }

        $newLeaf = Clean-SwitchName -OriginalName $folder.Name -IsFolder `
            -KeepRegion:$KeepRegion `
            -KeepLanguageTags:$KeepLanguageTags `
            -KeepSceneGroup:$KeepSceneGroup

        $relPath    = $folder.FullName.Substring($sourcePath.Length).TrimStart('\', '/')
        $targetPath = Join-Path $destPath $relPath

        $msg = "FOLDER: `"$($folder.FullName)`" -> `"$(Join-Path (Split-Path $targetPath -Parent) $newLeaf)`""
        Write-Log $msg

        if ($DryRun) {
            Write-Host "[DRYRUN] $msg" -ForegroundColor DarkCyan
        }
        else {
            Write-Host "[CREATED] $msg" -ForegroundColor Green
        }
    }
}

Write-Host "`nDone." -ForegroundColor Green
if ($DryRun) {
    Write-Host "Dry run complete. No files were copied or renamed." -ForegroundColor Yellow
}
else {
    Write-Host "Log saved to: $logPath" -ForegroundColor Cyan
}