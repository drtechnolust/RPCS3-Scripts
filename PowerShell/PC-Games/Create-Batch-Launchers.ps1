#Requires -Version 5.1
<#
.SYNOPSIS
    Creates .bat launcher files for PC games using the same executable scoring
    logic as Create-PC-Shortcuts.ps1, but outputs .bat files instead of .lnk.

.DESCRIPTION
    Scans a PC game library folder and creates a .bat launcher for each game.
    Each .bat file sets the working directory to the game's exe folder and
    launches the exe directly, which is more reliable than .lnk shortcuts for
    games that require their working directory to be set correctly.

    Uses the same executable detection and scoring engine as Create-PC-Shortcuts.ps1:
      - Progressive depth search (common paths first, then recursive fallback)
      - Blocks setup, uninstall, crash, config, redist executables
      - Scores based on folder name match, binary path, and game naming patterns
      - Timeout protection via background job for deep folder structures

    Output: one .bat file per game in BatchFiles\ subfolder.

.PARAMETER DryRun
    When $true (default), shows what would be created without writing any files.
    Set to $false in the CONFIG block to apply.

.EXAMPLE
    .\Create-Batch-Launchers.ps1

.VERSION
    2.0.0 - Config block, DryRun, CSV log, MIT header. Replaced interactive
            Read-Host prompt with config block. All detection logic preserved.
    1.0.0 - Initial release.

.LICENSE
    MIT License
    Copyright (c) Paul Mardis
#>

# ==============================================================================
# CONFIG -- Edit this block. Do not put paths anywhere else in this script.
# ==============================================================================
$RootGameFolder   = "D:\Arcade\System roms\PC Games"
$BatchOutputFolder = "D:\Arcade\System roms\PC Games\Shortcuts\BatchFiles"
$MaxDepth         = 8
$FolderTimeoutSec = 300
$DryRun           = $true    # Set to $false to write .bat files
$LogDir           = Join-Path (Split-Path $BatchOutputFolder -Parent) "Logs"
# ==============================================================================

function Set-ExitCode {
    param([int]$Code)
    $global:LASTEXITCODE = $Code
}

if (-not (Test-Path -LiteralPath $RootGameFolder)) {
    Write-Host "ERROR: Game library not found: $RootGameFolder" -ForegroundColor Red
    Set-ExitCode 1
    return
}

if (-not $DryRun) {
    foreach ($D in @($BatchOutputFolder, $LogDir)) {
        if (-not (Test-Path -LiteralPath $D)) { New-Item -ItemType Directory -Path $D -Force | Out-Null }
    }
}

$LogFile = Join-Path $LogDir ("Create-Batch-Launchers_{0}.csv" -f (Get-Date -Format "yyyyMMdd_HHmmss"))
$LogRows = [System.Collections.Generic.List[PSCustomObject]]::new()

function Write-LogRow {
    param([string]$GameName,[string]$ChosenExe,[string]$BatchPath,[string]$Status,[string]$Note)
    $LogRows.Add([PSCustomObject]@{
        Timestamp = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
        GameName  = $GameName
        ChosenExe = $ChosenExe
        BatchPath = $BatchPath
        Status    = $Status
        Note      = $Note
    })
}

# ------------------------------------------------------------------------------
# Executable scoring (same ruleset as Create-PC-Shortcuts.ps1)
# ------------------------------------------------------------------------------
function Get-ExecutableScore {
    param([string]$ExePath, [string]$GameFolderName)

    $ExeName    = [System.IO.Path]::GetFileNameWithoutExtension($ExePath).ToLower()
    $FolderName = $GameFolderName.ToLower()
    $Score      = 0

    $BlockPatterns = @("*uninstall*","*setup*","*settings*","*helper*","*config*","*language*","*crash*","*test*","*service*","*server*","*update*","*install*")
    foreach ($P in $BlockPatterns) { if ($ExeName -like $P) { return -1 } }

    $BlockNames = @("unins000","crashreport","errorreport","crashreporter","crashpad","redist","redistributable","vcredist","vc_redist","directx","dxwebsetup","uploader","webhelper","crs-handler","crs-uploader","crs-video","drivepool","quicksfv","handler","gamingrepair","unitycrashhandle64")
    if ($BlockNames -contains $ExeName) { return -1 }

    if ($ExeName -eq $FolderName)              { return 1000 }

    $Parts = $FolderName -split ' '
    foreach ($Part in $Parts) {
        if ($Part.Length -gt 3 -and $ExeName -eq $Part) { return 900 }
    }

    if ($ExeName -eq "$FolderName-win64-shipping") { return 500 }
    if ($ExeName -like "*wingdk*shipping*")         { $Score += 400 }
    if ($ExeName -like "*$FolderName*" -and $ExeName -like "*win64*shipping*") { $Score += 300 }
    if ($ExeName -like "*game*")                    { $Score += 75  }
    if ($ExeName -like "*$FolderName*")             { $Score += 50  }

    foreach ($Part in $Parts) {
        if ($Part.Length -gt 3 -and $ExeName -like "*$Part*") { $Score += 20 }
    }

    if (@("start","play","run","main","bin") -contains $ExeName) { $Score += 30 }

    $ExeDir    = [System.IO.Path]::GetDirectoryName($ExePath).ToLower()
    $GoodPaths = @("\bin","\binaries","\game","\app","\win64","\win32","\windows","\x64","\x86","\wingdk")
    foreach ($GP in $GoodPaths) {
        if ($ExeDir -like "*$GP*") { $Score += 20; break }
    }

    $Depth = ($ExePath.Split('\').Count - $RootGameFolder.Split('\').Count)
    if ($Depth -gt 6) { $Score -= 10 }

    return $Score
}

# ------------------------------------------------------------------------------
# Progressive exe finder with timeout (same as Create-PC-Shortcuts.ps1)
# ------------------------------------------------------------------------------
function Find-Executables {
    param([string]$FolderPath, [string]$GameName, [int]$MaxDepthParam, [int]$TimeoutSec)

    $ExeFiles = @()

    $CommonPaths = @(
        "$FolderPath\*.exe"
        "$FolderPath\binaries_x64\*.exe"
        "$FolderPath\Binaries\WinGDK\*.exe"
        "$FolderPath\Game\*.exe"
        "$FolderPath\app\*.exe"
        "$FolderPath\bin\*.exe"
        "$FolderPath\binaries\*.exe"
        "$FolderPath\Windows\*.exe"
        "$FolderPath\x64\*.exe"
        "$FolderPath\Win64\*.exe"
    )
    foreach ($CP in $CommonPaths) {
        if (Test-Path $CP) {
            $F = Get-ChildItem -Path $CP -ErrorAction SilentlyContinue
            if ($F) { $ExeFiles += $F }
        }
    }

    if ($ExeFiles.Count -eq 0) {
        $DeeperPaths = @(
            "$FolderPath\Binaries\Win64\*.exe"
            "$FolderPath\*\Binaries\Win64\*.exe"
            "$FolderPath\*\*\Binaries\Win64\*.exe"
            "$FolderPath\*\Binaries\WinGDK\*.exe"
            "$FolderPath\binaries_x86\*.exe"
        )
        foreach ($DP in $DeeperPaths) {
            if ($ExeFiles.Count -eq 0) {
                $F = Get-ChildItem -Path $DP -ErrorAction SilentlyContinue
                if ($F) { $ExeFiles += $F }
            }
        }
    }

    if ($ExeFiles.Count -gt 0) { return $ExeFiles }

    foreach ($D in @(0, 1, 2, 4, $MaxDepthParam)) {
        $ExeFiles = @(Get-ChildItem -LiteralPath $FolderPath -Filter "*.exe" -File -Depth $D -ErrorAction SilentlyContinue)
        if ($ExeFiles.Count -gt 0) { return $ExeFiles }
    }

    $Job = Start-Job -ScriptBlock { param($P) Get-ChildItem -Path $P -Filter "*.exe" -File -Recurse -ErrorAction SilentlyContinue } -ArgumentList $FolderPath
    $null = Wait-Job -Job $Job -Timeout ($TimeoutSec - 5)
    if ($Job.State -eq "Running") { Stop-Job -Job $Job } else { $ExeFiles = @(Receive-Job -Job $Job) }
    Remove-Job -Job $Job -Force
    return $ExeFiles
}

# ------------------------------------------------------------------------------
# Write .bat file
# ------------------------------------------------------------------------------
$CreatedTargets = @{}
$Utf8NoBom      = New-Object System.Text.UTF8Encoding($false)

function Save-BatchFile {
    param([string]$TargetExe, [string]$GameName, [string]$OutputFolder)

    $SafeName  = $GameName -replace '[\\/\:\*\?"<>\|]', '_'
    $SafeName  = $SafeName -replace '&', 'and'
    if ($SafeName.Length -gt 50) { $SafeName = $SafeName.Substring(0, 47) + "..." }

    $BatchPath = Join-Path $OutputFolder "$SafeName.bat"
    $WorkDir   = [System.IO.Path]::GetDirectoryName($TargetExe)

    if (Test-Path -LiteralPath $BatchPath)        { return "exists" }
    if ($CreatedTargets.ContainsKey($TargetExe))  { return "duplicate" }

    try {
        $Content = "@echo off`r`ncd /d `"$WorkDir`"`r`nstart `"`" `"$TargetExe`"`r`n"
        [System.IO.File]::WriteAllText($BatchPath, $Content, $Utf8NoBom)
        $CreatedTargets[$TargetExe] = $GameName
        return "created"
    } catch { return "error" }
}

# ==============================================================================
# BANNER
# ==============================================================================
Write-Host ""
Write-Host "Create-Batch-Launchers" -ForegroundColor Cyan
Write-Host "======================" -ForegroundColor Cyan
Write-Host "  Library   : $RootGameFolder"
Write-Host "  Output    : $BatchOutputFolder"
Write-Host "  DryRun    : $DryRun"
Write-Host ""
if ($DryRun) {
    Write-Host "  [DRY RUN] No .bat files will be written." -ForegroundColor Yellow
    Write-Host ""
}

# ==============================================================================
# MAIN LOOP
# ==============================================================================
$GameFolders = Get-ChildItem -LiteralPath $RootGameFolder -Directory | Sort-Object Name
$Total       = $GameFolders.Count
$Index       = 0
$CntCreated  = 0
$CntSkipped  = 0
$CntNotFound = 0
$CntErrors   = 0
$CntTimeouts = 0

foreach ($GameDir in $GameFolders) {
    $Index++
    $GameName = $GameDir.Name

    Write-Progress -Activity "Create-Batch-Launchers" -Status "$GameName ($Index of $Total)" `
        -PercentComplete ([int](($Index / $Total) * 100))

    try {
        $SearchStart = Get-Date
        $ExeFiles    = @(Find-Executables -FolderPath $GameDir.FullName -GameName $GameName -MaxDepthParam $MaxDepth -TimeoutSec $FolderTimeoutSec)
        $SearchSecs  = ((Get-Date) - $SearchStart).TotalSeconds

        if ($SearchSecs -gt ($FolderTimeoutSec * 0.9)) {
            $CntTimeouts++
            Write-Host ("  [TIMEOUT]   {0}" -f $GameName) -ForegroundColor Yellow
        }

        if ($ExeFiles.Count -eq 0) {
            Write-Host ("  [NOT FOUND] {0}" -f $GameName) -ForegroundColor DarkGray
            Write-LogRow -GameName $GameName -Status "NotFound" -Note "No .exe files found"
            $CntNotFound++
            continue
        }

        $Scored = $ExeFiles | ForEach-Object {
            [PSCustomObject]@{
                Path  = $_.FullName
                Score = Get-ExecutableScore -ExePath $_.FullName -GameFolderName $GameName
            }
        } | Where-Object { $_.Score -ge 0 } | Sort-Object Score -Descending

        if ($Scored.Count -eq 0) {
            Write-Host ("  [BLOCKED]   {0}" -f $GameName) -ForegroundColor DarkGray
            Write-LogRow -GameName $GameName -Status "AllBlocked" -Note "All .exe files blocked by filters"
            $CntNotFound++
            continue
        }

        $ChosenExe = $Scored[0].Path
        $BatchOut  = Join-Path $BatchOutputFolder (($GameName -replace '[\\/\:\*\?"<>\|]', '_') + ".bat")

        if ($DryRun) {
            Write-Host ("  [WOULD CREATE] {0}.bat" -f $GameName) -ForegroundColor Cyan
            Write-Host ("    Exe: {0}" -f $ChosenExe)
            Write-LogRow -GameName $GameName -ChosenExe $ChosenExe -BatchPath $BatchOut `
                -Status "DryRun" -Note "Score=$($Scored[0].Score)"
            $CntCreated++
            continue
        }

        $Result = Save-BatchFile -TargetExe $ChosenExe -GameName $GameName -OutputFolder $BatchOutputFolder

        switch ($Result) {
            "created"   { Write-Host ("  [OK]     {0}" -f $GameName) -ForegroundColor Green;    $CntCreated++ }
            "exists"    { Write-Host ("  [EXISTS] {0}" -f $GameName) -ForegroundColor DarkGray; $CntSkipped++ }
            "duplicate" { Write-Host ("  [DUPE]   {0}" -f $GameName) -ForegroundColor DarkGray; $CntSkipped++ }
            default     { Write-Host ("  [ERROR]  {0}" -f $GameName) -ForegroundColor Red;      $CntErrors++  }
        }
        Write-LogRow -GameName $GameName -ChosenExe $ChosenExe -BatchPath $BatchOut `
            -Status $Result -Note "Score=$($Scored[0].Score)"
    }
    catch {
        Write-Host ("  [ERROR]  {0} -- {1}" -f $GameName, $_.Exception.Message) -ForegroundColor Red
        Write-LogRow -GameName $GameName -Status "Error" -Note $_.Exception.Message
        $CntErrors++
    }
}

Write-Progress -Activity "Create-Batch-Launchers" -Completed

if ($LogRows.Count -gt 0 -and -not $DryRun) {
    try { $LogRows | Export-Csv -LiteralPath $LogFile -NoTypeInformation -Encoding UTF8 } catch {}
    Write-Host ""
    Write-Host ("  Log: {0}" -f $LogFile) -ForegroundColor Gray
}

Write-Host ""
Write-Host "Summary" -ForegroundColor White
Write-Host ("  {0,-12} {1}" -f "Created :", $CntCreated)  -ForegroundColor Green
Write-Host ("  {0,-12} {1}" -f "Skipped :", $CntSkipped)  -ForegroundColor Gray
Write-Host ("  {0,-12} {1}" -f "Not found:", $CntNotFound) -ForegroundColor DarkGray
Write-Host ("  {0,-12} {1}" -f "Errors :", $CntErrors)    -ForegroundColor $(if ($CntErrors -gt 0) { "Red" } else { "Gray" })
Write-Host ("  {0,-12} {1}" -f "Timeouts :", $CntTimeouts) -ForegroundColor $(if ($CntTimeouts -gt 0) { "Yellow" } else { "Gray" })
if ($DryRun) {
    Write-Host ""
    Write-Host "  DRY RUN complete. Set `$DryRun = `$false in the CONFIG block to apply." -ForegroundColor Yellow
}
Write-Host ""
Set-ExitCode 0
