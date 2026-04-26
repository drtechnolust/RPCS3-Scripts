# ================================
# Image Cleanup and Sorting Script
# v4.1 - PS7 Parallel, Final
# ================================

#Requires -Version 7.0

using namespace System.Collections.Concurrent

# -------- CONFIG --------
$SourceRoot      = "S:\Organize 2025\Images"
$DestinationRoot = "S:\Organize 2025\Image_Cleanup"

# $true  = safe preview, nothing moves
# $false = files WILL be moved
$DryRun = $false

# Threads: 8 for SSD, 4 for HDD
$ThrottleLimit = 8

# Thresholds
$MinWidth      = 300
$MinHeight     = 300
$MinFileSizeKB = 30

# Web-asset filename keywords
$WebKeywords = @(
    'icon','icons','arrow','arrows','button','buttons','bullet','bullets',
    'spacer','thumb','thumbnail','logo','logos','pixel','pixels','banner',
    'background','bg','divider','sprite','emoticon','smiley','avatar','nav'
)

$PhotoExtensions  = @('.jpg','.jpeg','.tif','.tiff','.bmp','.raw','.dng','.cr2','.nef','.arw')
$DesignExtensions = @('.psd','.ai','.eps','.xcf','.cap')
$NonImageExtensions = @('.htm','.html')

# -------- FOLDERS --------
$Folders = @{
    LikelyPhotos    = Join-Path $DestinationRoot "Likely_Photos"
    LargePNGReview  = Join-Path $DestinationRoot "Large_PNG_Review"
    TransparentPNGs = Join-Path $DestinationRoot "Transparent_PNGs"
    AnimatedGIFs    = Join-Path $DestinationRoot "Animated_GIFs"
    WebAssets       = Join-Path $DestinationRoot "Web_Assets"
    TooSmall        = Join-Path $DestinationRoot "Too_Small"
    DesignSource    = Join-Path $DestinationRoot "Design_Source"
    NonImageFiles   = Join-Path $DestinationRoot "Non_Image_Files"
    Unknown         = Join-Path $DestinationRoot "Unknown"
    Errors          = Join-Path $DestinationRoot "Errors"
    Logs            = Join-Path $DestinationRoot "_Logs"
}

foreach ($f in $Folders.Values) {
    if (-not (Test-Path $f)) { New-Item -ItemType Directory -Path $f -Force | Out-Null }
}

$TimeStamp = Get-Date -Format "yyyyMMdd_HHmmss"
$LogFile   = Join-Path $Folders.Logs "ImageCleanup_$TimeStamp.csv"

# -------- RESUME INDEX --------
$ProcessedPaths = [System.Collections.Generic.HashSet[string]]::new(
    [System.StringComparer]::OrdinalIgnoreCase
)
$priorLogs = Get-ChildItem -Path $Folders.Logs -Filter "ImageCleanup_*.csv" -File -ErrorAction SilentlyContinue
if ($priorLogs) {
    Write-Host "Loading resume index from $($priorLogs.Count) prior log(s)..." -ForegroundColor DarkGray
    foreach ($log in $priorLogs) {
        Import-Csv -Path $log.FullName | ForEach-Object { [void]$ProcessedPaths.Add($_.SourcePath) }
    }
    Write-Host "  $($ProcessedPaths.Count) previously processed paths indexed." -ForegroundColor DarkGray
    Write-Host ""
}

# -------- SHARED STATE --------
# long[] = reference type, correctly shared across parallel threads via $using:
$SharedCounter = [long[]]::new(1)
$ResultsBag    = [ConcurrentBag[object]]::new()
$DestCounters  = [ConcurrentDictionary[string,int]]::new()

# -------- BANNER --------
$stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

Write-Host "============================================" -ForegroundColor Cyan
Write-Host "  Image Cleanup and Sorter  v4.1 (Fast)"    -ForegroundColor Cyan
Write-Host "  PowerShell $($PSVersionTable.PSVersion)"
Write-Host "============================================" -ForegroundColor Cyan
Write-Host "  Source      : $SourceRoot"
Write-Host "  Destination : $DestinationRoot"
Write-Host "  Threads     : $ThrottleLimit"
if ($DryRun) {
    Write-Host "  Mode        : DRY RUN - no files will be moved" -ForegroundColor Yellow
} else {
    Write-Host "  Mode        : LIVE - files WILL be moved" -ForegroundColor Red
}
Write-Host ""

# -------- SCAN --------
Write-Host "Scanning directory tree..." -ForegroundColor Cyan
$allFiles = Get-ChildItem -LiteralPath $SourceRoot -Recurse -File -ErrorAction SilentlyContinue

$files = if ($ProcessedPaths.Count -gt 0) {
    @($allFiles | Where-Object { -not $ProcessedPaths.Contains($_.FullName) })
} else {
    @($allFiles)
}

$totalAll  = $allFiles.Count
$totalTodo = $files.Count
$skipped   = $totalAll - $totalTodo

Write-Host "  Total files : $("{0:N0}" -f $totalAll)" -ForegroundColor Cyan
if ($skipped -gt 0) {
    Write-Host "  Skipping    : $("{0:N0}" -f $skipped) already processed" -ForegroundColor DarkGray
}
Write-Host "  Processing  : $("{0:N0}" -f $totalTodo)" -ForegroundColor Cyan
Write-Host ""

if ($totalTodo -eq 0) {
    Write-Host "Nothing left to process. Exiting." -ForegroundColor Yellow
    exit
}

Write-Host "Running... progress bar will update every 5,000 files." -ForegroundColor Cyan
Write-Host ""

# -------- PARALLEL PROCESSING --------
$files | ForEach-Object -Parallel {

    # Bring all config into parallel scope
    $Folders            = $using:Folders
    $DryRun             = $using:DryRun
    $MinWidth           = $using:MinWidth
    $MinHeight          = $using:MinHeight
    $MinFileSizeKB      = $using:MinFileSizeKB
    $WebKeywords        = $using:WebKeywords
    $PhotoExtensions    = $using:PhotoExtensions
    $DesignExtensions   = $using:DesignExtensions
    $NonImageExtensions = $using:NonImageExtensions
    $ResultsBag         = $using:ResultsBag
    $DestCounters       = $using:DestCounters
    $SharedCounter      = $using:SharedCounter   # long[] reference type
    $TotalTodo          = $using:totalTodo

    $file      = $_
    $ext       = $file.Extension.ToLowerInvariant()
    $nameLower = $file.Name.ToLowerInvariant()
    $sizeKB    = $file.Length / 1KB

    # Increment shared counter atomically
    $n = [System.Threading.Interlocked]::Increment([ref]$SharedCounter[0])

    # Update progress bar every 5,000 files
    if ($n % 5000 -eq 0) {
        $pct = [math]::Round(($n / $TotalTodo) * 100, 1)
        Write-Progress -Activity "Image Cleanup & Sorter v4.1" `
            -Status "$("{0:N0}" -f $n) of $("{0:N0}" -f $TotalTodo) files  ($pct%)" `
            -PercentComplete $pct
    }

    # ---- Helpers ----

    function Get-DestPath($folder, $fname) {
        $key = "$folder|$fname"
        $idx = $using:DestCounters
        $i   = $idx.AddOrUpdate($key, 0, [System.Func[string,int,int]]{ param($k,$v) $v + 1 })
        if ($i -eq 0) { return [System.IO.Path]::Combine($folder, $fname) }
        $b = [System.IO.Path]::GetFileNameWithoutExtension($fname)
        $x = [System.IO.Path]::GetExtension($fname)
        return [System.IO.Path]::Combine($folder, "${b}_${i}${x}")
    }

    function Write-Result($f, $cat, $reason, $dest, $w=0, $h=0) {
        $dp = Get-DestPath $dest $f.Name
        if (-not $using:DryRun) {
            Move-Item -LiteralPath $f.FullName -Destination $dp -Force -ErrorAction SilentlyContinue
        }
        ($using:ResultsBag).Add([PSCustomObject]@{
            DateProcessed   = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
            DryRun          = $using:DryRun
            Category        = $cat
            Reason          = $reason
            SourcePath      = $f.FullName
            DestinationPath = $dp
            Extension       = $f.Extension
            FileSizeKB      = [math]::Round($f.Length / 1KB, 2)
            Width           = $w
            Height          = $h
        })
    }

    function Test-WebKeyword($name) {
        foreach ($kw in $using:WebKeywords) {
            if ($name -match [regex]::Escape($kw)) { return $true }
        }
        return $false
    }

    function Get-JpegDimensions($path) {
        try {
            $b = [System.IO.File]::ReadAllBytes($path)
            $i = 2
            while ($i -lt ($b.Length - 8)) {
                if ($b[$i] -ne 0xFF) { break }
                $m = $b[$i+1]
                if ($m -in @(0xC0,0xC1,0xC2,0xC3,0xC5,0xC6,0xC7,0xC9,0xCA,0xCB,0xCD,0xCE,0xCF)) {
                    return [PSCustomObject]@{
                        Width  = ($b[$i+7] -shl 8) -bor $b[$i+8]
                        Height = ($b[$i+5] -shl 8) -bor $b[$i+6]
                    }
                }
                $seg = ($b[$i+2] -shl 8) -bor $b[$i+3]
                $i  += 2 + $seg
            }
        } catch {}
        return $null
    }

    function Get-PngInfo($path) {
        try {
            $buf = New-Object byte[] 28
            $fs  = [System.IO.File]::OpenRead($path)
            [void]$fs.Read($buf, 0, 28)
            $fs.Close()
            if ($buf[0] -eq 0x89 -and $buf[1] -eq 0x50) {
                return [PSCustomObject]@{
                    Width    = ($buf[16] -shl 24) -bor ($buf[17] -shl 16) -bor ($buf[18] -shl 8) -bor $buf[19]
                    Height   = ($buf[20] -shl 24) -bor ($buf[21] -shl 16) -bor ($buf[22] -shl 8) -bor $buf[23]
                    HasAlpha = ($buf[25] -band 0x04) -ne 0
                }
            }
        } catch {}
        return $null
    }

    function Get-GifInfo($path) {
        try {
            $b      = [System.IO.File]::ReadAllBytes($path)
            $width  = ($b[7] -shl 8) -bor $b[6]
            $height = ($b[9] -shl 8) -bor $b[8]
            $fc     = 0
            $i      = 13
            if ($b[10] -band 0x80) { $i += 3 * [math]::Pow(2, ($b[10] -band 0x07) + 1) }
            while ($i -lt ($b.Length - 1)) {
                if ($b[$i] -eq 0x3B) { break }
                if ($b[$i] -eq 0x21 -and ($i+1) -lt $b.Length -and $b[$i+1] -eq 0xF9) {
                    $fc++; $i += 8; continue
                }
                $i++
                if ($i -ge $b.Length) { break }
                $bs = $b[$i]
                while ($bs -ne 0 -and $i -lt $b.Length) {
                    $i += $bs + 1
                    if ($i -lt $b.Length) { $bs = $b[$i] }
                }
                $i++
            }
            return [PSCustomObject]@{ Width=$width; Height=$height; IsAnimated=($fc -gt 1) }
        } catch {}
        return $null
    }

    function Get-GdiDimensions($path) {
        Add-Type -AssemblyName System.Drawing
        $img = $null
        try {
            $img = [System.Drawing.Image]::FromFile($path)
            return [PSCustomObject]@{ Width=$img.Width; Height=$img.Height }
        } catch { return $null }
        finally { if ($null -ne $img) { $img.Dispose() } }
    }

    # ---- Classification ----
    try {
        if ($NonImageExtensions -contains $ext) {
            Write-Result $file "Non_Image_Files" "HTML/HTM file" $Folders.NonImageFiles
            return
        }

        if ($DesignExtensions -contains $ext) {
            Write-Result $file "Design_Source" "Design/source file" $Folders.DesignSource
            return
        }

        $hasKW  = Test-WebKeyword $nameLower
        $tinyKB = ($sizeKB -lt $MinFileSizeKB)

        if ($ext -eq '.gif') {
            if ($tinyKB) { Write-Result $file "Web_Assets" "GIF below min size" $Folders.WebAssets; return }
            $gi = Get-GifInfo $file.FullName
            if (-not $gi) { Write-Result $file "Errors" "Could not read GIF" $Folders.Errors; return }
            if ($gi.IsAnimated) {
                Write-Result $file "Animated_GIFs" "Animated GIF" $Folders.AnimatedGIFs $gi.Width $gi.Height
            } elseif (($gi.Width -lt $MinWidth) -or ($gi.Height -lt $MinHeight) -or $hasKW) {
                Write-Result $file "Web_Assets" "Small/static GIF or web keyword" $Folders.WebAssets $gi.Width $gi.Height
            } else {
                Write-Result $file "Unknown" "GIF needs manual review" $Folders.Unknown $gi.Width $gi.Height
            }
            return
        }

        if ($ext -eq '.png') {
            if ($tinyKB) { Write-Result $file "Web_Assets" "PNG below min size" $Folders.WebAssets; return }
            $pi = Get-PngInfo $file.FullName
            if (-not $pi) {
                $gdi = Get-GdiDimensions $file.FullName
                if (-not $gdi) { Write-Result $file "Errors" "Could not read PNG" $Folders.Errors; return }
                $pi = [PSCustomObject]@{ Width=$gdi.Width; Height=$gdi.Height; HasAlpha=$false }
            }
            $sm = ($pi.Width -lt $MinWidth) -or ($pi.Height -lt $MinHeight)
            if ($sm -and $pi.HasAlpha) {
                Write-Result $file "Transparent_PNGs" "Small transparent PNG" $Folders.TransparentPNGs $pi.Width $pi.Height
            } elseif ($sm -or $hasKW) {
                Write-Result $file "Web_Assets" "Small PNG or web keyword" $Folders.WebAssets $pi.Width $pi.Height
            } else {
                Write-Result $file "Large_PNG_Review" "Large PNG for review" $Folders.LargePNGReview $pi.Width $pi.Height
            }
            return
        }

        if ($PhotoExtensions -contains $ext) {
            if ($tinyKB) { Write-Result $file "Too_Small" "Below min file size" $Folders.TooSmall; return }
            if ($hasKW)  { Write-Result $file "Too_Small" "Web-style filename"  $Folders.TooSmall; return }
            $dim = $null
            if ($ext -in @('.jpg','.jpeg')) { $dim = Get-JpegDimensions $file.FullName }
            if (-not $dim) { $dim = Get-GdiDimensions $file.FullName }
            if (-not $dim) { Write-Result $file "Errors" "Could not read dimensions: $ext" $Folders.Errors; return }
            if (($dim.Width -lt $MinWidth) -or ($dim.Height -lt $MinHeight)) {
                Write-Result $file "Too_Small" "Dimensions below threshold" $Folders.TooSmall $dim.Width $dim.Height
            } else {
                Write-Result $file "Likely_Photos" "Likely photo" $Folders.LikelyPhotos $dim.Width $dim.Height
            }
            return
        }

        Write-Result $file "Unknown" "Unhandled extension: $ext" $Folders.Unknown

    } catch {
        ($using:ResultsBag).Add([PSCustomObject]@{
            DateProcessed   = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
            DryRun          = $using:DryRun
            Category        = 'Errors'
            Reason          = $_.Exception.Message
            SourcePath      = $file.FullName
            DestinationPath = ''
            Extension       = $file.Extension
            FileSizeKB      = [math]::Round($file.Length / 1KB, 2)
            Width           = 0
            Height          = 0
        })
    }

} -ThrottleLimit $ThrottleLimit

Write-Progress -Activity "Image Cleanup & Sorter v4.1" -Completed

# -------- EXPORT LOG --------
Write-Host "Writing log..." -ForegroundColor Cyan
$allResults = $ResultsBag.ToArray()
$allResults | Export-Csv -NoTypeInformation -Encoding UTF8 -Path $LogFile

$stopwatch.Stop()

# -------- SUMMARY --------
Write-Host ""
Write-Host "============================================" -ForegroundColor Green
Write-Host "  Summary"                                   -ForegroundColor Green
Write-Host "============================================" -ForegroundColor Green
Write-Host ""

$allResults |
    Group-Object Category |
    Sort-Object Name |
    ForEach-Object {
        $mb = [math]::Round(($_.Group | Measure-Object -Property FileSizeKB -Sum).Sum / 1KB, 1)
        "  {0,-22} {1,8:N0} files   {2,8} MB" -f $_.Name, $_.Count, $mb
    } | Write-Host

$grandMB = [math]::Round(($allResults | Measure-Object -Property FileSizeKB -Sum).Sum / 1KB, 1)

Write-Host ""
Write-Host ("  {0,-22} {1,8:N0} files   {2,8} MB" -f "TOTAL", $allResults.Count, $grandMB) -ForegroundColor Cyan
Write-Host ""
Write-Host ("  Elapsed  : {0:hh\:mm\:ss\.ff}" -f $stopwatch.Elapsed) -ForegroundColor DarkGray
Write-Host  "  Log      : $LogFile" -ForegroundColor DarkGray
Write-Host ""

if ($DryRun) {
    Write-Host "  DRY RUN complete. Set `$DryRun = `$false to move files for real." -ForegroundColor Yellow
    Write-Host ""
}