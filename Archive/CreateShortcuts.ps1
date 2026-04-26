# Define the path to the games folder
$gamesFolderPath = "\\10.10.1.99\retro games\Arcade\System roms\PC Games 2"
# Define the directory where you want to save the shortcuts
$shortcutDestination = "\\10.10.1.99\retro games\Arcade\System roms\PC Games"

Write-Host "Starting the script..." -ForegroundColor Green
Write-Host "Searching for all folders in: $gamesFolderPath" -ForegroundColor Yellow

# Ensure the destination folder exists
if (!(Test-Path -Path $shortcutDestination)) {
    Write-Host "Shortcut destination folder does not exist. Creating folder..." -ForegroundColor Cyan
    New-Item -ItemType Directory -Path $shortcutDestination
}

# Function to create a shortcut
function Create-Shortcut($targetPath, $shortcutName) {
    Write-Host "Creating shortcut for: $shortcutName" -ForegroundColor Cyan
    $shell = New-Object -ComObject WScript.Shell
    $shortcut = $shell.CreateShortcut("$shortcutDestination\$shortcutName.lnk")
    $shortcut.TargetPath = $targetPath
    $shortcut.Save()
    Write-Host "Shortcut created: $shortcutName" -ForegroundColor Green
}

# Get all directories from the games folder
try {
    $folders = Get-ChildItem -Path $gamesFolderPath -Directory
    Write-Host "$($folders.Count) folders found." -ForegroundColor Green
} catch {
    Write-Host "Error: Unable to access or retrieve folders. Check permissions or network access." -ForegroundColor Red
    exit
}

# Initialize a counter for tracking progress
$counter = 0
$skipped = 0

# Loop through each folder and process .exe files inside
foreach ($folder in $folders) {
    Write-Host "Processing folder: $($folder.Name)" -ForegroundColor Yellow

    # Recursively get all .exe files from the folder
    $exeFiles = Get-ChildItem -Path $folder.FullName -Recurse -Filter *.exe

    foreach ($file in $exeFiles) {
        # Filter out common executables that are not the game's main launcher
        if ($file.Name -notmatch "uninstall|setup|readme|support|helper|config") {
            # Use the folder name as the shortcut name
            $gameFolderName = Split-Path -Parent $file.FullName | Split-Path -Leaf
            Write-Host "Processing: $gameFolderName ($counter processed so far)" -ForegroundColor Yellow
            Create-Shortcut $file.FullName $gameFolderName

            # Increment the counter and display the current progress
            $counter++
        } else {
            Write-Host "Skipping: $($file.Name) (Not a valid launcher)" -ForegroundColor DarkGray
            $skipped++
        }
    }
}

Write-Host "Shortcuts created successfully for all folders! Total shortcuts: $counter" -ForegroundColor Green
Write-Host "Total files skipped: $skipped" -ForegroundColor Yellow
