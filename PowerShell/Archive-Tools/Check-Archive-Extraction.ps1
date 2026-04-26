<#
.SYNOPSIS
    Checks which archives have been extracted in a directory containing mixed archives and folders.

.DESCRIPTION
    This script scans a directory for archive files (.rar, .zip, .7z, etc.) and checks if a 
    corresponding folder exists with the same name, indicating the archive has been extracted.
    Useful for cleaning up extracted archives to save disk space.

.PARAMETER Path
    The directory path to scan for archives. This directory should contain both archive files
    and their extracted folders.

.PARAMETER DeleteExtracted
    If specified, prompts to delete archive files that have already been extracted.
    Requires confirmation before deletion.

.PARAMETER ShowOnlyMatches
    If specified, only shows archives that have been extracted (hides unextracted list).

.PARAMETER ShowOnlyUnextracted
    If specified, only shows archives that have NOT been extracted (hides extracted list).

.EXAMPLE
    .\TeknoParrot-Check.ps1 -Path "D:\Arcade\System roms\TeknoParrot"
    
    Basic check - Shows all archives and their extraction status.

.EXAMPLE
    .\TeknoParrot-Check.ps1 -Path "D:\Arcade\System roms\TeknoParrot" -ShowOnlyUnextracted
    
    Shows only archives that haven't been extracted yet.

.EXAMPLE
    .\TeknoParrot-Check.ps1 -Path "D:\Arcade\System roms\TeknoParrot" -ShowOnlyMatches
    
    Shows only archives that have been extracted (useful to see what can be deleted).

.EXAMPLE
    .\TeknoParrot-Check.ps1 -Path "D:\Arcade\System roms\TeknoParrot" -DeleteExtracted
    
    Checks archives and offers to delete those that have been extracted.
    Requires typing 'YES' to confirm deletion.

.NOTES
    Author: Archive Verification Script
    Version: 1.0
    
    The script assumes that archives are extracted to folders with the same name as the archive
    file (without the extension). For example:
    - "Game Name.rar" extracts to "Game Name" folder
    - "Another Game.zip" extracts to "Another Game" folder

.LINK
    https://docs.microsoft.com/en-us/powershell/
#>

# TeknoParrot Archive Verification Script
# Checks for archives that have been extracted in the same directory

param(
    [Parameter(Mandatory=$true)]
    [string]$Path = "D:\Arcade\System roms\TeknoParrot",
    
    [Parameter()]
    [switch]$DeleteExtracted,
    
    [Parameter()]
    [switch]$ShowOnlyMatches,
    
    [Parameter()]
    [switch]$ShowOnlyUnextracted
)

# Load required assemblies
Add-Type -AssemblyName System.IO.Compression.FileSystem

Write-Host "`nScanning directory: $Path" -ForegroundColor Cyan
Write-Host "=" * 80

# Get all archives (RAR, ZIP, 7Z, and other common formats)
$archives = Get-ChildItem -Path $Path -File | Where-Object { 
    $_.Extension -in @('.rar', '.zip', '.7z', '.7zip', '.001', '.tar', '.gz', '.bz2', '.xz', '.cab', '.iso') 
}

$results = @{
    Extracted = @()
    NotExtracted = @()
    Errors = @()
}

$total = $archives.Count
$count = 0

Write-Host "Found $total archive files to check`n" -ForegroundColor Yellow

foreach ($archive in $archives) {
    $count++
    $archiveName = [System.IO.Path]::GetFileNameWithoutExtension($archive.Name)
    $expectedFolder = Join-Path $Path $archiveName
    
    Write-Progress -Activity "Checking archives..." -Status "$count of $total - $($archive.Name)" -PercentComplete (($count / $total) * 100)
    
    # Check if corresponding folder exists
    $folderExists = Test-Path $expectedFolder
    
    $result = [PSCustomObject]@{
        Archive = $archive.Name
        Size = "{0:N2} MB" -f ($archive.Length / 1MB)
        FolderExists = $folderExists
        FolderPath = if ($folderExists) { $expectedFolder } else { "Not found" }
        Status = if ($folderExists) { "Extracted" } else { "Not Extracted" }
    }
    
    if ($folderExists) {
        $results.Extracted += $result
    } else {
        $results.NotExtracted += $result
    }
}

Write-Progress -Activity "Checking archives..." -Completed

# Display results
if (-not $ShowOnlyUnextracted) {
    Write-Host "`nEXTRACTED ARCHIVES ($($results.Extracted.Count)):" -ForegroundColor Green
    Write-Host "-" * 80
    if ($results.Extracted.Count -gt 0) {
        $results.Extracted | Format-Table Archive, Size, Status -AutoSize
    } else {
        Write-Host "No extracted archives found." -ForegroundColor Gray
    }
}

if (-not $ShowOnlyMatches) {
    Write-Host "`nNOT EXTRACTED ARCHIVES ($($results.NotExtracted.Count)):" -ForegroundColor Red
    Write-Host "-" * 80
    if ($results.NotExtracted.Count -gt 0) {
        $results.NotExtracted | Format-Table Archive, Size, Status -AutoSize
    } else {
        Write-Host "All archives have been extracted!" -ForegroundColor Gray
    }
}

# Summary
Write-Host "`nSUMMARY:" -ForegroundColor Cyan
Write-Host "-" * 80
Write-Host "Total archives: $total" -ForegroundColor White
Write-Host "Extracted: $($results.Extracted.Count) ($([math]::Round($results.Extracted.Count/$total*100, 2))%)" -ForegroundColor Green
Write-Host "Not extracted: $($results.NotExtracted.Count) ($([math]::Round($results.NotExtracted.Count/$total*100, 2))%)" -ForegroundColor Red

# Calculate space that could be freed
if ($results.Extracted.Count -gt 0) {
    $extractedSize = ($archives | Where-Object { 
        [System.IO.Path]::GetFileNameWithoutExtension($_.Name) -in $results.Extracted.Archive.Replace([System.IO.Path]::GetExtension($_), '')
    } | Measure-Object -Property Length -Sum).Sum / 1GB
    
    Write-Host "`nSpace that could be freed by deleting extracted archives: $([math]::Round($extractedSize, 2)) GB" -ForegroundColor Yellow
}

# Option to delete extracted archives
if ($DeleteExtracted -and $results.Extracted.Count -gt 0) {
    Write-Host "`nDELETE MODE ENABLED!" -ForegroundColor Red
    $confirm = Read-Host "Delete $($results.Extracted.Count) extracted archive files? (YES/N)"
    
    if ($confirm -eq 'YES') {
        foreach ($item in $results.Extracted) {
            $archivePath = Join-Path $Path $item.Archive
            Remove-Item -Path $archivePath -Force
            Write-Host "Deleted: $($item.Archive)" -ForegroundColor Red
        }
        Write-Host "`nDeleted $($results.Extracted.Count) archive files." -ForegroundColor Green
    } else {
        Write-Host "Deletion cancelled." -ForegroundColor Yellow
    }
}

# Export results to CSV
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$logFile = Join-Path $Path "archive_check_$timestamp.csv"
$allResults = $results.Extracted + $results.NotExtracted
$allResults | Export-Csv -Path $logFile -NoTypeInformation

Write-Host "`nResults saved to: $logFile" -ForegroundColor Cyan