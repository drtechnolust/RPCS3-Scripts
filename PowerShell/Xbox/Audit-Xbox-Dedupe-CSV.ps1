#Requires -Version 5.1

<#
.SYNOPSIS
    Audits the Xbox duplicate log CSV for JP-region files that were moved,
    and shows what beat them (the keeper in the same group).
#>

$csvPath = Read-Host "Paste the full path to your Xbox_DuplicateLog CSV"

if (-not (Test-Path $csvPath)) {
    Write-Host "File not found: $csvPath" -ForegroundColor Red
    exit
}

$all = Import-Csv -Path $csvPath

# Find all JP entries that were moved
$jpMoved = $all | Where-Object {
    $_.Region -in @('JP','Japan') -and $_.Action -in @('Move','WouldMove')
}

if ($jpMoved.Count -eq 0) {
    Write-Host "No JP-region files were moved." -ForegroundColor Green
    exit
}

Write-Host ""
Write-Host "JP files that were moved ($($jpMoved.Count) total):" -ForegroundColor Cyan
Write-Host "─────────────────────────────────────────────────────────────────" -ForegroundColor Cyan
Write-Host ""

foreach ($item in $jpMoved | Sort-Object NormalizedKey) {
    # Find what was kept in the same group
    $keeper = $all | Where-Object {
        $_.NormalizedKey -eq $item.NormalizedKey -and $_.Action -eq 'Keep'
    }

    Write-Host "MOVED : $($item.FileName)" -ForegroundColor Yellow
    if ($keeper) {
        foreach ($k in $keeper) {
            Write-Host "KEPT  : $($k.FileName)  [$($k.Region)] — $($k.Reason)" -ForegroundColor Green
        }
    } else {
        Write-Host "KEPT  : (no keeper found — may be a data issue)" -ForegroundColor Red
    }
    Write-Host ""
}

# Export a focused report
$reportPath = Join-Path ([System.IO.Path]::GetDirectoryName($csvPath)) "JP_Audit_Report.csv"

$report = foreach ($item in $jpMoved | Sort-Object NormalizedKey) {
    $keeper = $all | Where-Object {
        $_.NormalizedKey -eq $item.NormalizedKey -and $_.Action -eq 'Keep'
    } | Select-Object -First 1

    [PSCustomObject]@{
        NormalizedKey   = $item.NormalizedKey
        JP_File_Moved   = $item.FileName
        Kept_File       = if ($keeper) { $keeper.FileName } else { '(none found)' }
        Kept_Region     = if ($keeper) { $keeper.Region   } else { '' }
        Kept_Reason     = if ($keeper) { $keeper.Reason   } else { '' }
        JP_SizeMB       = $item.SizeMB
    }
}

$report | Export-Csv -Path $reportPath -NoTypeInformation -Encoding UTF8
Write-Host "─────────────────────────────────────────────────────────────────" -ForegroundColor Cyan
Write-Host "Audit report saved: $reportPath" -ForegroundColor Green
Write-Host ""