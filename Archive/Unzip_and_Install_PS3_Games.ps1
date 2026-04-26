param(
    [string]$SourceRoot = "D:\Arcade\System roms\Sony Playstation 3\Sony - PlayStation 3 (PSN) (Content)",
    [string]$Rpcs3Exe   = "C:\Arcade\LaunchBox\Emulators\RPCS3\rpcs3.exe",
    [string]$Rpcs3Root  = "C:\Arcade\LaunchBox\Emulators\RPCS3",
    [string]$TempRoot   = "D:\Arcade\RPCS3_Temp",
    [switch]$ForceReinstall
)

# ── Unicode-safe output ──────────────────────────────────────────────────────
$OutputEncoding = New-Object System.Text.UTF8Encoding($false)
try { [Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false) } catch {}

# ── Derived paths ────────────────────────────────────────────────────────────
$GameRoot      = Join-Path $Rpcs3Root "dev_hdd0\game"
$RapTargetRoot = Join-Path $Rpcs3Root "dev_hdd0\home\00000001\exdata"
$PS3Root       = "D:\Arcade\System roms\Sony Playstation 3"
$StateRoot     = Join-Path $PS3Root "_Automation"
$InstalledRoot = Join-Path $PS3Root "_Installed"
$FailedRoot    = Join-Path $PS3Root "_Failed"
$LogFile       = Join-Path $StateRoot "install_log.csv"
$StateFile     = Join-Path $StateRoot "installed_state.json"

# ════════════════════════════════════════════════════════════════════════════
#  INTERACTIVE RUN-TYPE PROMPT
# ════════════════════════════════════════════════════════════════════════════
function Show-Banner {
    Write-Host ""
    Write-Host "╔══════════════════════════════════════════════════════╗" -ForegroundColor Cyan
    Write-Host "║         RPCS3 Batch PKG Installer                   ║" -ForegroundColor Cyan
    Write-Host "╚══════════════════════════════════════════════════════╝" -ForegroundColor Cyan
    Write-Host ""
}

function Prompt-RunType {
    Write-Host "  Select run type:" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "  [1]  Normal     - Install new ZIPs, skip already-installed" -ForegroundColor White
    Write-Host "  [2]  Dry Run    - Preview only, no changes made" -ForegroundColor White
    Write-Host "  [3]  Force      - Reinstall ALL ZIPs (ignore state file)" -ForegroundColor White
    Write-Host "  [4]  Failed     - Retry ZIPs currently in _failed folder" -ForegroundColor White
    Write-Host "  [Q]  Quit" -ForegroundColor White
    Write-Host ""

    do {
        $key = Read-Host "  Your choice"
        $key = $key.Trim().ToUpper()
    } while ($key -notin @("1","2","3","4","Q"))

    return $key
}

Show-Banner
$runChoice = Prompt-RunType

if ($runChoice -eq "Q") {
    Write-Host "`n  Aborted." -ForegroundColor Red
    return
}

$DryRun        = $runChoice -eq "2"
$ForceReinstall= $runChoice -eq "3"
$RetryFailed   = $runChoice -eq "4"

$modeLabel = switch ($runChoice) {
    "1" { "Normal"    }
    "2" { "Dry Run"   }
    "3" { "Force Reinstall" }
    "4" { "Retry Failed"    }
}

Write-Host ""
Write-Host "  Mode     : " -NoNewline -ForegroundColor Gray
Write-Host $modeLabel -ForegroundColor Cyan
Write-Host "  Source   : " -NoNewline -ForegroundColor Gray
Write-Host $SourceRoot -ForegroundColor Cyan
Write-Host "  RPCS3    : " -NoNewline -ForegroundColor Gray
Write-Host $Rpcs3Exe -ForegroundColor Cyan
Write-Host ""

# ════════════════════════════════════════════════════════════════════════════
#  HELPER FUNCTIONS
# ════════════════════════════════════════════════════════════════════════════
function Ensure-Folder {
    param([string]$LiteralPath)
    if (-not (Test-Path -LiteralPath $LiteralPath)) {
        New-Item -ItemType Directory -Path $LiteralPath -Force | Out-Null
    }
}

function Write-TextUtf8NoBom {
    param([string]$LiteralPath, [string]$Text)
    $utf8 = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllText($LiteralPath, $Text, $utf8)
}

function Append-LineUtf8NoBom {
    param([string]$LiteralPath, [string]$Line)
    $utf8 = New-Object System.Text.UTF8Encoding($false)
    $sw   = New-Object System.IO.StreamWriter($LiteralPath, $true, $utf8)
    try   { $sw.WriteLine($Line) }
    finally { $sw.Dispose() }
}

function Write-LogRow {
    param(
        [string]$ArchivePath, [string]$PkgPath, [string]$RapPath,
        [string]$Status, [string]$Details, [string]$DetectedGameFolder,
        [int]$ExitCode
    )
    $row = [PSCustomObject]@{
        Timestamp          = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
        ArchivePath        = $ArchivePath
        PkgPath            = $PkgPath
        RapPath            = $RapPath
        Status             = $Status
        Details            = $Details
        DetectedGameFolder = $DetectedGameFolder
        ExitCode           = $ExitCode
    }
    if (-not (Test-Path -LiteralPath $LogFile)) {
        $header = ($row | ConvertTo-Csv -NoTypeInformation)[0]
        Append-LineUtf8NoBom -LiteralPath $LogFile -Line $header
    }
    $line = ($row | ConvertTo-Csv -NoTypeInformation)[1]
    Append-LineUtf8NoBom -LiteralPath $LogFile -Line $line
}

function Load-State {
    if (Test-Path -LiteralPath $StateFile) {
        try {
            $raw = [System.IO.File]::ReadAllText($StateFile, [System.Text.Encoding]::UTF8)
            if ([string]::IsNullOrWhiteSpace($raw)) { return @{} }
            # PS 5.1 safe: ConvertFrom-Json returns PSCustomObject, convert manually
            $obj = $raw | ConvertFrom-Json
            $ht  = @{}
            $obj.PSObject.Properties | ForEach-Object { $ht[$_.Name] = $_.Value }
            return $ht
        }
        catch { return @{} }
    }
    return @{}
}

function Save-State {
    param([hashtable]$State)
    $json = $State | ConvertTo-Json -Depth 10
    Write-TextUtf8NoBom -LiteralPath $StateFile -Text $json
}

function Get-FileHashSafe {
    param([string]$LiteralPath)
    try { return (Get-FileHash -LiteralPath $LiteralPath -Algorithm SHA256).Hash }
    catch { return $null }
}

function Get-GameSnapshot {
    param([string]$LiteralPath)
    $snapshot = @{}
    if (Test-Path -LiteralPath $LiteralPath) {
        Get-ChildItem -LiteralPath $LiteralPath -Directory | ForEach-Object {
            $paramSfo = Join-Path $_.FullName "PARAM.SFO"
            $lastWrite = $_.LastWriteTimeUtc
            if (Test-Path -LiteralPath $paramSfo) {
                $lastWrite = (Get-Item -LiteralPath $paramSfo).LastWriteTimeUtc
            }
            $snapshot[$_.FullName] = $lastWrite
        }
    }
    return $snapshot
}

function Find-NewOrChangedGameFolder {
    param([hashtable]$Before, [hashtable]$After, [datetime]$InstallStartUtc)
    $candidates = @()
    foreach ($path in $After.Keys) {
        if (-not $Before.ContainsKey($path)) {
            $candidates += [PSCustomObject]@{ Path = $path; LastWrite = $After[$path]; Reason = "New" }
        }
        elseif ($After[$path] -gt $Before[$path] -and $After[$path] -ge $InstallStartUtc.AddMinutes(-1)) {
            $candidates += [PSCustomObject]@{ Path = $path; LastWrite = $After[$path]; Reason = "Updated" }
        }
    }
    return ($candidates | Sort-Object LastWrite -Descending)
}

function Move-ToFolder {
    param([string]$LiteralPath, [string]$DestinationFolder)
    if (-not (Test-Path -LiteralPath $LiteralPath)) { return $null }
    Ensure-Folder -LiteralPath $DestinationFolder
    $leaf = [System.IO.Path]::GetFileName($LiteralPath)
    $dest = Join-Path $DestinationFolder $leaf
    if (Test-Path -LiteralPath $dest) {
        $base = [System.IO.Path]::GetFileNameWithoutExtension($LiteralPath)
        $ext  = [System.IO.Path]::GetExtension($LiteralPath)
        $dest = Join-Path $DestinationFolder ("{0}_{1}{2}" -f $base, (Get-Date -Format "yyyyMMdd_HHmmss"), $ext)
    }
    Move-Item -LiteralPath $LiteralPath -Destination $dest -Force
    return $dest
}

function Remove-FolderSafe {
    param([string]$LiteralPath)
    if (Test-Path -LiteralPath $LiteralPath) {
        Remove-Item -LiteralPath $LiteralPath -Recurse -Force -ErrorAction SilentlyContinue
    }
}

function Get-MatchingRapForPkg {
    param([System.IO.FileInfo]$PkgFile, [System.IO.FileInfo[]]$RapFiles)
    $pkgBase = [System.IO.Path]::GetFileNameWithoutExtension($PkgFile.Name)
    return ($RapFiles | Where-Object {
        [System.IO.Path]::GetFileNameWithoutExtension($_.Name) -eq $pkgBase
    } | Select-Object -First 1)
}

# ── Progress display helpers ─────────────────────────────────────────────────
function Write-ProgressLine {
    param([int]$Current, [int]$Total, [string]$ZipName)
    $pct  = [math]::Floor(($Current / $Total) * 100)
    $done = [math]::Floor($pct / 5)   # 20-char bar
    $bar  = ("█" * $done).PadRight(20, "░")
    Write-Host ("`r  [{0}] {1,3}%  ({2}/{3})  {4}" -f $bar, $pct, $Current, $Total, $ZipName) -NoNewline -ForegroundColor Cyan
}

function Write-StatusLine {
    param([string]$Icon, [string]$Color, [string]$Label, [string]$Detail)
    Write-Host ""
    Write-Host ("  {0} {1,-10} {2}" -f $Icon, $Label, $Detail) -ForegroundColor $Color
}

# ════════════════════════════════════════════════════════════════════════════
#  PRE-FLIGHT CHECKS
# ════════════════════════════════════════════════════════════════════════════
Ensure-Folder -LiteralPath $StateRoot
Ensure-Folder -LiteralPath $InstalledRoot
Ensure-Folder -LiteralPath $FailedRoot
Ensure-Folder -LiteralPath $TempRoot
Ensure-Folder -LiteralPath $RapTargetRoot

if (-not (Test-Path -LiteralPath $Rpcs3Exe)) {
    Write-Host "  ERROR: RPCS3 executable not found: $Rpcs3Exe" -ForegroundColor Red
    return
}
if (-not (Test-Path -LiteralPath $GameRoot)) {
    Write-Host "  ERROR: RPCS3 game folder not found: $GameRoot" -ForegroundColor Red
    return
}

# ════════════════════════════════════════════════════════════════════════════
#  COLLECT ZIP FILES
# ════════════════════════════════════════════════════════════════════════════
$scanRoot = if ($RetryFailed) { $FailedRoot } else { $SourceRoot }

$zipFiles = Get-ChildItem -LiteralPath $scanRoot -Recurse -File -Filter *.zip |
    Where-Object {
        $_.FullName -notmatch '\\_installed\\' -and
        $_.FullName -notmatch '\\_automation\\'
    } |
    Sort-Object FullName

if (-not $zipFiles) {
    $label = if ($RetryFailed) { "_failed folder" } else { "source folder" }
    Write-Host "  No ZIP files found in $label." -ForegroundColor Yellow
    return
}

$total = $zipFiles.Count
Write-Host "  Found $total ZIP file(s) to process." -ForegroundColor White
Write-Host ""

$state = Load-State

# ── Session counters ─────────────────────────────────────────────────────────
$counters = @{ Success = 0; Skipped = 0; Failed = 0; Error = 0 }
$failedItems = New-Object System.Collections.Generic.List[PSCustomObject]

# ════════════════════════════════════════════════════════════════════════════
#  MAIN LOOP
# ════════════════════════════════════════════════════════════════════════════
$idx = 0
foreach ($zip in $zipFiles) {
    $idx++
    $zipPath = $zip.FullName
    $zipName = $zip.Name

    Write-ProgressLine -Current $idx -Total $total -ZipName $zipName

    $zipHash = Get-FileHashSafe -LiteralPath $zipPath

    # ── Skip check ────────────────────────────────────────────────────────
    if (-not $ForceReinstall -and -not $RetryFailed -and $zipHash -and $state.ContainsKey($zipHash)) {
        Write-StatusLine -Icon "⏭" -Color DarkGray -Label "SKIPPED" -Detail $zipName
        $counters.Skipped++
        Write-LogRow -ArchivePath $zipPath -PkgPath "" -RapPath "" -Status "Skipped" `
            -Details "Already in state file" `
            -DetectedGameFolder ($state[$zipHash].DetectedGameFolder) -ExitCode 0
        continue
    }

    $extractFolder = Join-Path $TempRoot ([System.IO.Path]::GetFileNameWithoutExtension($zip.Name))

    try {
        Remove-FolderSafe -LiteralPath $extractFolder
        Ensure-Folder -LiteralPath $extractFolder
        Expand-Archive -LiteralPath $zipPath -DestinationPath $extractFolder -Force

        $pkgFiles = Get-ChildItem -LiteralPath $extractFolder -Recurse -File -Filter *.pkg
        $rapFiles = Get-ChildItem -LiteralPath $extractFolder -Recurse -File -Filter *.rap

        # ── No PKG found ──────────────────────────────────────────────────
        if (-not $pkgFiles) {
            Write-StatusLine -Icon "✖" -Color Red -Label "NO PKG" -Detail $zipName
            $counters.Failed++
            $failedItems.Add([PSCustomObject]@{
                Zip    = $zipName
                Reason = "No PKG file found inside archive"
                Exit   = "N/A"
            })
            Write-LogRow -ArchivePath $zipPath -PkgPath "" -RapPath "" -Status "Failed" `
                -Details "No PKG found inside ZIP" -DetectedGameFolder "" -ExitCode -1
            Move-ToFolder -LiteralPath $zipPath -DestinationFolder $FailedRoot | Out-Null
            Remove-FolderSafe -LiteralPath $extractFolder
            continue
        }

        # ── Dry Run ───────────────────────────────────────────────────────
        if ($DryRun) {
            foreach ($pkg in $pkgFiles) {
                $rap = Get-MatchingRapForPkg -PkgFile $pkg -RapFiles $rapFiles
                $rapNote = if ($rap) { "RAP matched" } else { "No RAP" }
                Write-StatusLine -Icon "🔍" -Color DarkCyan -Label "DRY RUN" -Detail ("$($pkg.Name)  [$rapNote]")
                Write-LogRow -ArchivePath $zipPath -PkgPath $pkg.FullName `
                    -RapPath $(if($rap){$rap.FullName}else{""}) `
                    -Status "DryRun" -Details "No changes made" -DetectedGameFolder "" -ExitCode 0
            }
            Remove-FolderSafe -LiteralPath $extractFolder
            continue
        }

        # ── Install ───────────────────────────────────────────────────────
        $archiveSuccess  = $true
        $detectedFolders = New-Object System.Collections.Generic.List[string]

        foreach ($pkg in $pkgFiles) {
            $rap = Get-MatchingRapForPkg -PkgFile $pkg -RapFiles $rapFiles

            $before          = Get-GameSnapshot -LiteralPath $GameRoot
            $installStartUtc = (Get-Date).ToUniversalTime()

            if ($rap) {
                $rapDest = Join-Path $RapTargetRoot $rap.Name
                Copy-Item -LiteralPath $rap.FullName -Destination $rapDest -Force
            }

            Write-Host ""
            Write-Host ("  ▶ Installing: {0}" -f $pkg.Name) -ForegroundColor White

            $proc = Start-Process -FilePath $Rpcs3Exe `
                                  -ArgumentList @("--installpkg", $pkg.FullName) `
                                  -PassThru -Wait -NoNewWindow
            $exitCode = $proc.ExitCode
            Start-Sleep -Seconds 3

            $after   = Get-GameSnapshot -LiteralPath $GameRoot
            $changes = Find-NewOrChangedGameFolder -Before $before -After $after -InstallStartUtc $installStartUtc
            $detectedFolder = if ($changes) { $changes[0].Path } else { "" }

            if ($exitCode -eq 0 -and $detectedFolder) {
                $detectedFolders.Add($detectedFolder) | Out-Null
                $folderName = [System.IO.Path]::GetFileName($detectedFolder)
                Write-StatusLine -Icon "✔" -Color Green -Label "OK" -Detail ("{0}  →  {1}" -f $pkg.Name, $folderName)
                Write-LogRow -ArchivePath $zipPath -PkgPath $pkg.FullName `
                    -RapPath $(if($rap){$rap.FullName}else{""}) `
                    -Status "Success" -Details "RPCS3 exit 0, folder detected" `
                    -DetectedGameFolder $detectedFolder -ExitCode $exitCode
            }
            else {
                $archiveSuccess = $false
                $reason = if ($exitCode -ne 0) {
                    "RPCS3 exit code $exitCode"
                } else {
                    "No game folder detected after install"
                }
                Write-StatusLine -Icon "✖" -Color Red -Label "FAILED" -Detail ("{0}  —  {1}" -f $pkg.Name, $reason)
                $failedItems.Add([PSCustomObject]@{
                    Zip    = $zipName
                    Reason = $reason
                    Exit   = $exitCode
                })
                Write-LogRow -ArchivePath $zipPath -PkgPath $pkg.FullName `
                    -RapPath $(if($rap){$rap.FullName}else{""}) `
                    -Status "Failed" -Details $reason `
                    -DetectedGameFolder $detectedFolder -ExitCode $exitCode
            }
        }

        if ($archiveSuccess) {
            $counters.Success++
            $movedTo = Move-ToFolder -LiteralPath $zipPath -DestinationFolder $InstalledRoot
            if ($zipHash) {
                $state[$zipHash] = @{
                    Timestamp           = (Get-Date).ToString("s")
                    ArchiveOriginalPath = $zipPath
                    ArchiveMovedTo      = $movedTo
                    DetectedGameFolder  = ($detectedFolders -join "; ")
                }
                Save-State -State $state
            }
        }
        else {
            $counters.Failed++
            Move-ToFolder -LiteralPath $zipPath -DestinationFolder $FailedRoot | Out-Null
        }

        Remove-FolderSafe -LiteralPath $extractFolder
    }
    catch {
        Write-StatusLine -Icon "⚠" -Color Red -Label "ERROR" -Detail ("{0}  —  {1}" -f $zipName, $_.Exception.Message)
        $counters.Error++
        $failedItems.Add([PSCustomObject]@{
            Zip    = $zipName
            Reason = $_.Exception.Message
            Exit   = "exception"
        })
        Write-LogRow -ArchivePath $zipPath -PkgPath "" -RapPath "" -Status "Error" `
            -Details $_.Exception.Message -DetectedGameFolder "" -ExitCode -1

        if (Test-Path -LiteralPath $zipPath) {
            Move-ToFolder -LiteralPath $zipPath -DestinationFolder $FailedRoot | Out-Null
        }
        Remove-FolderSafe -LiteralPath $extractFolder
    }
}

# ════════════════════════════════════════════════════════════════════════════
#  SESSION SUMMARY
# ════════════════════════════════════════════════════════════════════════════
Write-Host ""
Write-Host ""
Write-Host "╔══════════════════════════════════════════════════════╗" -ForegroundColor Cyan
Write-Host "║                   SESSION SUMMARY                   ║" -ForegroundColor Cyan
Write-Host "╚══════════════════════════════════════════════════════╝" -ForegroundColor Cyan
Write-Host ""
Write-Host ("  Total processed : {0}" -f $total)           -ForegroundColor White
Write-Host ("  ✔  Succeeded    : {0}" -f $counters.Success) -ForegroundColor Green
Write-Host ("  ⏭  Skipped      : {0}" -f $counters.Skipped) -ForegroundColor DarkGray
Write-Host ("  ✖  Failed       : {0}" -f ($counters.Failed + $counters.Error)) -ForegroundColor $(if(($counters.Failed + $counters.Error) -gt 0){"Red"}else{"Green"})

if ($failedItems.Count -gt 0) {
    Write-Host ""
    Write-Host "  ── Failed Items ──────────────────────────────────────" -ForegroundColor Red
    foreach ($item in $failedItems) {
        Write-Host ("  • {0}" -f $item.Zip)     -ForegroundColor Yellow
        Write-Host ("    Reason : {0}" -f $item.Reason) -ForegroundColor Gray
        Write-Host ("    Exit   : {0}" -f $item.Exit)   -ForegroundColor Gray
    }
}

Write-Host ""
Write-Host "  Log   : $LogFile"   -ForegroundColor DarkGray
Write-Host "  State : $StateFile" -ForegroundColor DarkGray
Write-Host ""