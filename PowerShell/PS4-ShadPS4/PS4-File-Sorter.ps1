# Set your main directory
$sourcePath = "D:\Arcade\System roms\Sony Playstation 4\Downloading PS4"

# Define destination folders
$folders = @{
    Game   = Join-Path $sourcePath "Games"
    Update = Join-Path $sourcePath "Updates"
    DLC    = Join-Path $sourcePath "DLC"
    Misc   = Join-Path $sourcePath "Misc"
}

# Create destination folders if missing
foreach ($folder in $folders.Values) {
    if (-not (Test-Path $folder)) {
        New-Item -ItemType Directory -Path $folder | Out-Null
    }
}

# Define log file
$logFile = Join-Path $sourcePath "OrganizeLog.txt"
"[$(Get-Date)] --- PKG Organizer Started ---`n" | Out-File -FilePath $logFile -Encoding UTF8

# Function to classify .pkg file
function Get-PkgType {
    param ($fileName)
    if ($fileName -match "A\d{4}-V\d{4}") {
        return "Update"
    }
    elseif ($fileName -match "(?i)(DLC|ADDON|TRACK|WEAPON|PACK|SEASONPASS|COSTUME|STYLE|CHARACTER|EXPANSION)") {
        return "DLC"
    }
    elseif ($fileName -match "(?i)v1\.00" -or ($fileName -notmatch "A\d{4}-V\d{4}" -and $fileName -notmatch "(?i)(DLC|ADDON)")) {
        return "Game"
    }
    else {
        return "Misc"
    }
}

# Process .pkg files
Get-ChildItem -Path $sourcePath -Filter *.pkg -File | ForEach-Object {
    try {
        $fileName = $_.Name
        $filePath = $_.FullName
        $type     = Get-PkgType $fileName

        # Use base name without extension as subfolder name
        $baseName = [System.IO.Path]::GetFileNameWithoutExtension($fileName)
        $targetSubFolder = Join-Path $folders[$type] $baseName

        if (-not (Test-Path $targetSubFolder)) {
            New-Item -ItemType Directory -Path $targetSubFolder | Out-Null
        }

        # Move the main .pkg
        Move-Item -Path $filePath -Destination $targetSubFolder -Force
        Add-Content -Path $logFile -Value "Moved: $fileName -> $type\$baseName"

        # Move associated companion files
        $companionPatterns = "*$baseName*.nfo","*$baseName*.txt","*$baseName*.jpg","*$baseName*.jpeg"
        foreach ($pattern in $companionPatterns) {
            Get-ChildItem -Path $sourcePath -Filter $pattern -File | ForEach-Object {
                try {
                    Move-Item -Path $_.FullName -Destination $targetSubFolder -Force
                    Add-Content -Path $logFile -Value "Associated: $($_.Name) -> $type\$baseName"
                } catch {
                    Add-Content -Path $logFile -Value "ERROR: Could not move companion $($_.Name): $($_.Exception.Message)"
                }
            }
        }
    } catch {
        Add-Content -Path $logFile -Value "ERROR: Failed to process $($_.Name): $($_.Exception.Message)"
    }
}

Add-Content -Path $logFile -Value "`n[$(Get-Date)] --- PKG Organizer Complete ---"
