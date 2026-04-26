$shortcutsFolder = "\\10.10.1.99\retro games\Arcade\System roms\PC Games 2\shortcuts"  # Original folder containing your shortcuts
$newShortcutsFolder = "\\10.10.1.99\retro games\Arcade\System roms\PC Games 2\newshortcuts"  # New folder to save updated shortcuts
$oldPath = "D:\arcade\system roms\pc games 2"
$newPath = "D:\arcade\system roms\pc games"

# Variables for monitoring
$updatedCount = 0
$skippedCount = 0
$failedCount = 0
$startTime = Get-Date

# Create the new shortcuts folder if it doesn't exist
if (-not (Test-Path $newShortcutsFolder)) {
    New-Item -ItemType Directory -Path $newShortcutsFolder
    Write-Host "Created new folder: $newShortcutsFolder"
} else {
    Write-Host "Folder already exists: $newShortcutsFolder"
}

# Loop through all .lnk files in the folder
Get-ChildItem -Path $shortcutsFolder -Filter *.lnk | ForEach-Object {
    $shortcut = $_.FullName
    Write-Host "Processing shortcut: $($_.Name)" -ForegroundColor Cyan
    
    try {
        # Create a Shell COM Object to work with shortcuts
        $shell = New-Object -ComObject WScript.Shell
        $link = $shell.CreateShortcut($shortcut)

        # Replace the old target and start in paths with the new paths
        if ($link.TargetPath -like "$oldPath*") {
            Write-Host "Original TargetPath: $($link.TargetPath)" -ForegroundColor Yellow
            Write-Host "Original StartIn Path: $($link.WorkingDirectory)" -ForegroundColor Yellow

            # Updating target and start-in paths
            $link.TargetPath = $link.TargetPath.Replace($oldPath, $newPath)
            $link.WorkingDirectory = $link.WorkingDirectory.Replace($oldPath, $newPath)

            # Save the updated shortcut to the new folder
            $newShortcutPath = Join-Path $newShortcutsFolder $_.Name
            $newLink = $shell.CreateShortcut($newShortcutPath)
            $newLink.TargetPath = $link.TargetPath
            $newLink.WorkingDirectory = $link.WorkingDirectory
            $newLink.IconLocation = $link.IconLocation
            $newLink.Save()

            Write-Host "Updated TargetPath: $($newLink.TargetPath)" -ForegroundColor Green
            Write-Host "Updated StartIn Path: $($newLink.WorkingDirectory)" -ForegroundColor Green

            Write-Host "New shortcut saved to: $newShortcutPath" -ForegroundColor Green
            $updatedCount++
        } else {
            Write-Host "No changes for: $($_.Name)" -ForegroundColor Gray
            $skippedCount++
        }
    } catch {
        Write-Host "Error processing $($_.Name): $_" -ForegroundColor Red
        $failedCount++
    }
}

# Monitoring output
$endTime = Get-Date
$duration = $endTime - $startTime

Write-Host ""
Write-Host "------------------------------------" -ForegroundColor Magenta
Write-Host "Process completed!" -ForegroundColor Magenta
Write-Host "Start time: $startTime" -ForegroundColor Magenta
Write-Host "End time: $endTime" -ForegroundColor Magenta
Write-Host "Duration: $duration" -ForegroundColor Magenta
Write-Host "------------------------------------" -ForegroundColor Magenta
Write-Host "Updated shortcuts: $updatedCount" -ForegroundColor Green
Write-Host "Skipped shortcuts: $skippedCount" -ForegroundColor Yellow
Write-Host "Failed shortcuts: $failedCount" -ForegroundColor Red
