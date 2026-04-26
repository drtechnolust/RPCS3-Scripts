# LaunchBox Shortcut Generator
$rootGameFolder = "D:\Arcade\System roms\PC Games 2"
$shortcutOutputFolder = "$rootGameFolder\Shortcuts2"
$maxDepth = 3

$logFound = "$shortcutOutputFolder\FoundGames.log"
$logNotFound = "$shortcutOutputFolder\NotFoundGames.log"

# Ensure shortcut output folder exists
if (!(Test-Path $shortcutOutputFolder)) {
    New-Item -ItemType Directory -Path $shortcutOutputFolder | Out-Null
}

# Clear logs if they exist
Remove-Item -Path $logFound, $logNotFound -ErrorAction SilentlyContinue

function Get-ExecutableScore {
    param($exePath, $gameFolderName)

    $exeName = [System.IO.Path]::GetFileNameWithoutExtension($exePath).ToLower()
    $folderName = $gameFolderName.ToLower()

    $score = 0

    $badNames = @("uninstall", "unins000", "setup", "crashreport", "errorreport")
    if ($badNames -contains $exeName) {
        return -1
    }

    if ($exeName -eq $folderName) { $score += 10 }
    if ($exeName -like "*$folderName*") { $score += 5 }

    $priorityNames = @("game", "launcher", "start", "play", "run", "main")
    if ($priorityNames -contains $exeName) { $score += 3 }

    return $score
}

function Create-Shortcut {
    param (
        [string]$targetExe,
        [string]$shortcutName,
        [string]$outputFolder
    )

    $WScriptShell = New-Object -ComObject WScript.Shell
    $shortcut = $WScriptShell.CreateShortcut("$outputFolder\$shortcutName.lnk")
    $shortcut.TargetPath = $targetExe
    $shortcut.WorkingDirectory = [System.IO.Path]::GetDirectoryName($targetExe)
    $shortcut.Save()
}

$gameFolders = Get-ChildItem -Path $rootGameFolder -Directory
$total = $gameFolders.Count
$index = 0

foreach ($folder in $gameFolders) {
    $index++
    $gameName = $folder.Name
    Write-Progress -Activity "Scanning Games" -Status "$gameName ($index of $total)" -PercentComplete (($index / $total) * 100)

    $exeFiles = @()
    $currentFolder = $folder.FullName
    $depth = 0

    while ($depth -lt $maxDepth) {
        $exeFiles += Get-ChildItem -Path $currentFolder -Filter "*.exe" -File -Recurse:$false -ErrorAction SilentlyContinue

        if ($exeFiles.Count -gt 0 -and $depth -gt 0) {
            break
        }

        $subDirs = Get-ChildItem -Path $currentFolder -Directory -ErrorAction SilentlyContinue
        if ($subDirs.Count -ne 1) { break }

        $currentFolder = $subDirs[0].FullName
        $depth++
    }

    if ($exeFiles.Count -eq 0) {
        "$gameName - No executable found" | Out-File -FilePath $logNotFound -Append
        continue
    }

    $scored = $exeFiles | ForEach-Object {
        [PSCustomObject]@{
            Path  = $_.FullName
            Score = Get-ExecutableScore -exePath $_.FullName -gameFolderName $gameName
        }
    } | Where-Object { $_.Score -ge 0 } | Sort-Object Score -Descending

    if ($scored.Count -eq 0) {
        "$gameName - No suitable exe" | Out-File -FilePath $logNotFound -Append
        continue
    }

    $chosenExe = $scored[0].Path
    Create-Shortcut -targetExe $chosenExe -shortcutName $gameName -outputFolder $shortcutOutputFolder
    "$gameName - $chosenExe" | Out-File -FilePath $logFound -Append
}

Write-Host "✅ Done! Shortcuts saved to: $shortcutOutputFolder"
Write-Host "📄 See FoundGames.log and NotFoundGames.log for details"
