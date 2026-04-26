#Requires -Version 5.1

<#
.SYNOPSIS
    Xbox Region Duplicate Organizer v1.0

.DESCRIPTION
    Scans a folder of original Xbox ISO files, identifies regional duplicates,
    and moves lower-priority copies to a destination folder.

    Region priority (1G1R, English-first):
        USA > World > Europe > Japan > Asia > Korea > ...

    - Companion/variant files (Demo, Beta, Online Enabled, Kiosk, etc.)
      are NEVER touched — always kept.
    - When multiple versions exist (v1.00 vs v2.01), highest version wins.

.PARAMETER Source
    Path to the folder containing your Xbox ISO files.

.PARAMETER Destination
    Path to the folder where duplicate ISOs will be moved.
    Created automatically if it does not exist.

.PARAMETER DryRun
    Switch. When set, no files are moved — only a report is shown and the CSV
    log is written. Omit this switch to perform an actual move.

.EXAMPLE
    .\Xbox_RegionDuplicateOrganizer.ps1 -Source "D:\Xbox\Games" -Destination "D:\Xbox\Dupes" -DryRun

.EXAMPLE
    .\Xbox_RegionDuplicateOrganizer.ps1 -Source "D:\Xbox\Games" -Destination "D:\Xbox\Dupes"
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory, Position = 0)]
    [ValidateScript({ Test-Path $_ -PathType Container })]
    [string]$Source,

    [Parameter(Mandatory, Position = 1)]
    [string]$Destination,

    [Parameter()]
    [switch]$DryRun
)

# ── CONFIG ────────────────────────────────────────────────────────────────────
$FileExtensions = @('*.iso')   # Xbox uses ISO; add '*.xiso' if needed
# ─────────────────────────────────────────────────────────────────────────────

# Tags that mark a file as a special/companion — never moved
$CompanionTags = @(
    'Online Enabled','Demo','Kiosk','Beta','Proto',
    'Sample','Debug','Program','BIOS','Alt',
    'Bonus Disc','Bonus','Soundtrack','Theme'
)

# Full region names AND short codes both supported
$KnownRegions = @(
    # Short codes (No-Intro style)
    'US','EU','JP','AS','KO','GE','FR','IT','ES','AU',
    'NZ','CA','UK','TW','HK','CH','NL','NO','SE','DK',
    'FI','PL','RU','BR','MX',
    # Long-form fallbacks
    'USA','World','Europe','Japan','Asia','Korea',
    'Hong Kong','Taiwan','China','Australia',
    'France','Germany','Italy','Spain','Netherlands',
    'Norway','Sweden','Denmark','Finland','Poland',
    'Russia','Brazil','Mexico','Latin America',
    'Canada','New Zealand','Singapore','Thailand'
)

# Priority order — index 0 = best
# Short codes first; long-form entries at the end as fallback
$RegionPriority = @(
    'US','USA',
    'World',
    'EU','Europe',
    'JP','Japan',
    'AS','Asia',
    'KO','Korea',
    'AU','Australia',
    'UK',
    'FR','France',
    'GE','Germany',
    'IT','Italy',
    'ES','Spain',
    'NL','Netherlands',
    'NO','Norway',
    'SE','Sweden',
    'DK','Denmark',
    'FI','Finland',
    'PL','Poland',
    'RU','Russia',
    'BR','Brazil',
    'MX','Mexico',
    'Latin America',
    'HK','Hong Kong',
    'TW','Taiwan',
    'CH','China',
    'CA','Canada',
    'NZ','New Zealand',
    'SG','Singapore',
    'TH','Thailand',
    'Unknown'
)

# ==============================================================================
# FUNCTIONS
# ==============================================================================

function Test-IsCompanionFile {
    param([string]$BaseName)
    foreach ($tag in $CompanionTags) {
        $escaped = [regex]::Escape($tag)
        if ($BaseName -match "\($escaped\)") { return $true }
    }
    return $false
}

function Get-RegionFromName {
    param([string]$BaseName)
    $m = [regex]::Matches($BaseName, '\(([^()]*)\)')
    foreach ($match in $m) {
        $token = $match.Groups[1].Value.Trim()
        if ($KnownRegions -contains $token) { return $token }
    }
    return 'Unknown'
}

function Get-NormalizedKey {
    param([string]$BaseName)
    $result = $BaseName

    # Strip region tags — both short and long
    foreach ($region in $KnownRegions) {
        $escaped = [regex]::Escape($region)
        $result = [regex]::Replace($result, "\s*\($escaped\)", '', 'IgnoreCase')
    }

    # Strip language tags: (En,Fr,De) / (En,Ja) etc.
    $result = [regex]::Replace($result, '\s*\([A-Za-z]{2}(?:,[A-Za-z]{2})+\)', '')

    # Strip version tags: (v1.00), (v2.01)
    $result = [regex]::Replace($result, '\s*\(v[\d]+\.[\d]+[^)]*\)', '', 'IgnoreCase')

    # Strip revision tags: (Rev 1), (Rev A)
    $result = [regex]::Replace($result, '\s*\(Rev\s+[^)]+\)', '', 'IgnoreCase')

    # Strip (Version X) tags
    $result = [regex]::Replace($result, '\s*\(Version\s+[^)]+\)', '', 'IgnoreCase')

    # Strip companion tags so e.g. "Game (Online Enabled)" doesn't form its own group
    foreach ($tag in $CompanionTags) {
        $escaped = [regex]::Escape($tag)
        $result = [regex]::Replace($result, "\s*\($escaped\)", '', 'IgnoreCase')
    }

    # Normalize whitespace
    $result = [regex]::Replace($result, '\s+', ' ').Trim()
    return $result
}

function Get-VersionScore {
    param([string]$BaseName)

    $vMatch = [regex]::Match($BaseName, '\(v(\d+)\.(\d+)', 'IgnoreCase')
    if ($vMatch.Success) {
        $major = [int]$vMatch.Groups[1].Value
        $minor = [int]$vMatch.Groups[2].Value
        return ($major * 10000) + $minor
    }

    $rMatch = [regex]::Match($BaseName, '\(Rev\s+([^)]+)\)', 'IgnoreCase')
    if ($rMatch.Success) {
        $rev = $rMatch.Groups[1].Value.Trim()
        $num = 0
        if ([int]::TryParse($rev, [ref]$num)) { return $num }
        if ($rev.Length -eq 1 -and $rev -match '[A-Za-z]') {
            return [int][char]($rev.ToUpper()) - [int][char]'A' + 1
        }
    }

    return 0
}

function Get-RegionScore {
    param([string]$Region)
    $idx = [array]::IndexOf($RegionPriority, $Region)
    if ($idx -lt 0) { return 9999 }
    return $idx
}

function Get-PreferredFile {
    param([array]$FilesInGroup)
    if ($FilesInGroup.Count -le 1) { return $FilesInGroup[0] }

    $sorted = $FilesInGroup | Sort-Object -Property @(
        @{ Expression = { Get-RegionScore  $_.Region   }; Ascending = $true  },
        @{ Expression = { Get-VersionScore $_.BaseName }; Ascending = $false }
    )
    return $sorted[0]
}

# ==============================================================================
# MAIN
# ==============================================================================

try {
    $sourceFolder      = $Source
    $destinationFolder = $Destination
    $dryRun            = $DryRun.IsPresent

    Write-Host ""
    Write-Host "Xbox Region Duplicate Organizer v1.0" -ForegroundColor Cyan
    Write-Host "--------------------------------------" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "Source      : $sourceFolder"      -ForegroundColor Yellow
    Write-Host "Destination : $destinationFolder" -ForegroundColor Yellow
    Write-Host "Mode        : $(if ($dryRun) { 'Dry Run — no files will be moved' } else { 'LIVE — files WILL be moved' })" -ForegroundColor $(if ($dryRun) { 'Green' } else { 'Red' })
    Write-Host ""

    if (-not (Test-Path $destinationFolder)) {
        New-Item -Path $destinationFolder -ItemType Directory -Force | Out-Null
    }

    # ── Collect files — skip .part files ─────────────────────────────────────
    $allFiles = @()
    foreach ($ext in $FileExtensions) {
        $allFiles += Get-ChildItem -Path $sourceFolder -File -Filter $ext |
                     Where-Object { $_.Name -notlike '*.part' }
    }

    if (-not $allFiles -or $allFiles.Count -eq 0) {
        throw "No matching files found in: $sourceFolder"
    }

    Write-Host "Found $($allFiles.Count) file(s). Analyzing..." -ForegroundColor Green
    Write-Host ""

    $fileObjects    = @()
    $companionCount = 0

    foreach ($file in $allFiles) {
        $baseName      = [System.IO.Path]::GetFileNameWithoutExtension($file.Name)
        $isCompanion   = Test-IsCompanionFile -BaseName $baseName
        $region        = Get-RegionFromName   -BaseName $baseName
        $normalizedKey = Get-NormalizedKey    -BaseName $baseName

        if ($isCompanion) { $companionCount++ }

        $fileObjects += [PSCustomObject]@{
            Name          = $file.Name
            FullName      = $file.FullName
            BaseName      = $baseName
            Region        = $region
            NormalizedKey = $normalizedKey
            IsCompanion   = $isCompanion
            SizeMB        = [math]::Round($file.Length / 1MB, 2)
        }
    }

    # ── Stats ─────────────────────────────────────────────────────────────────
    $gameFiles       = @($fileObjects | Where-Object { -not $_.IsCompanion })
    $grouped         = $gameFiles | Group-Object -Property NormalizedKey
    $uniqueTitles    = ($grouped | Where-Object { $_.Count -eq 1 }).Count
    $duplicateGroups = ($grouped | Where-Object { $_.Count -gt 1 }).Count

    Write-Host "Companion/special files (always kept) : $companionCount"
    Write-Host "Unique titles (no duplicates found)   : $uniqueTitles"
    Write-Host "Titles with regional duplicates       : $duplicateGroups"
    Write-Host ""

    # ── Evaluate groups ───────────────────────────────────────────────────────
    $actions = @()

    foreach ($group in $grouped) {
        $groupFiles = @($group.Group)

        if ($groupFiles.Count -le 1) {
            $actions += [PSCustomObject]@{
                NormalizedKey = $group.Name
                FileName      = $groupFiles[0].Name
                Region        = $groupFiles[0].Region
                SizeMB        = $groupFiles[0].SizeMB
                Action        = 'Keep'
                Reason        = 'Only copy'
                SourcePath    = $groupFiles[0].FullName
                Destination   = ''
            }
            continue
        }

        $winner = Get-PreferredFile -FilesInGroup $groupFiles

        foreach ($item in $groupFiles) {
            $isWinner = ($item.FullName -eq $winner.FullName)
            $actions += [PSCustomObject]@{
                NormalizedKey = $group.Name
                FileName      = $item.Name
                Region        = $item.Region
                SizeMB        = $item.SizeMB
                Action        = if ($isWinner) { 'Keep' } elseif ($dryRun) { 'WouldMove' } else { 'Move' }
                Reason        = if ($isWinner) { 'Best region/version' } else { 'Lower priority duplicate' }
                SourcePath    = $item.FullName
                Destination   = if ($isWinner) { '' } else { Join-Path $destinationFolder $item.Name }
            }
        }
    }

    # Log companion files
    foreach ($c in ($fileObjects | Where-Object { $_.IsCompanion })) {
        $actions += [PSCustomObject]@{
            NormalizedKey = $c.NormalizedKey
            FileName      = $c.Name
            Region        = $c.Region
            SizeMB        = $c.SizeMB
            Action        = 'Keep'
            Reason        = 'Companion — excluded from dedup'
            SourcePath    = $c.FullName
            Destination   = ''
        }
    }

    # ── Summary ───────────────────────────────────────────────────────────────
    $toMove   = @($actions | Where-Object { $_.Action -in @('WouldMove','Move') })
    $toKeep   = @($actions | Where-Object { $_.Action -eq 'Keep' })
    $moveSize = [math]::Round(($toMove | Measure-Object -Property SizeMB -Sum).Sum / 1024, 2)

    Write-Host "──────────────────────────────────────" -ForegroundColor Cyan
    Write-Host "Files to keep : $($toKeep.Count)"
    Write-Host "Files to move : $($toMove.Count)  (~$moveSize GB)"
    Write-Host "──────────────────────────────────────" -ForegroundColor Cyan
    Write-Host ""

    # ── Save CSV log ──────────────────────────────────────────────────────────
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $logFile   = Join-Path $sourceFolder "Xbox_DuplicateLog_$timestamp.csv"
    $actions | Sort-Object NormalizedKey, Action |
        Export-Csv -Path $logFile -NoTypeInformation -Encoding UTF8
    Write-Host "Log saved : $logFile" -ForegroundColor Green
    Write-Host ""

    # ── Move files ────────────────────────────────────────────────────────────
    if (-not $dryRun -and $toMove.Count -gt 0) {
        $moveIndex  = 0
        $moveErrors = 0

        foreach ($item in $toMove) {
            $moveIndex++
            Write-Host "  Moving [$moveIndex/$($toMove.Count)] $($item.FileName)" -ForegroundColor DarkGray

            if (Test-Path $item.SourcePath) {
                $destPath = $item.Destination

                if (Test-Path $destPath) {
                    $base     = [System.IO.Path]::GetFileNameWithoutExtension($item.FileName)
                    $ext      = [System.IO.Path]::GetExtension($item.FileName)
                    $destPath = Join-Path $destinationFolder ("{0}__DUP_{1}{2}" -f $base, (Get-Date -Format "yyyyMMddHHmmssfff"), $ext)
                }

                try {
                    Move-Item -Path $item.SourcePath -Destination $destPath -Force
                }
                catch {
                    Write-Host "    ERROR: $($_.Exception.Message)" -ForegroundColor Red
                    $moveErrors++
                }
            }
            else {
                Write-Host "    SKIPPED (file not found): $($item.FileName)" -ForegroundColor Yellow
            }
        }

        Write-Host ""
        if ($moveErrors -gt 0) {
            Write-Host "Completed with $moveErrors error(s)." -ForegroundColor Yellow
        } else {
            Write-Host "All files moved successfully." -ForegroundColor Green
        }
    }

    # ── Preview ───────────────────────────────────────────────────────────────
    if ($toMove.Count -gt 0) {
        Write-Host ""
        Write-Host "Preview — $(if ($dryRun) { 'would move' } else { 'moved' }) (first 30):" -ForegroundColor Yellow
        $toMove | Select-Object -First 30 FileName, Region, SizeMB, Reason | Format-Table -AutoSize
    } else {
        Write-Host "No duplicates found — nothing to move." -ForegroundColor Green
    }

    Write-Host ""
    Write-Host "Done." -ForegroundColor Green
    Write-Host ""
}
catch {
    Write-Host ""
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host ""
}