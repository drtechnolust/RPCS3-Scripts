# File Country Code Renamer - PowerShell Version

function Find-CountryCodes {
    param([string]$Directory)
    
    $pattern = '\(([^)]+)\)'
    $codes = @()
    
    Get-ChildItem -Path $Directory -File | ForEach-Object {
        $matches = [regex]::Matches($_.Name, $pattern)
        foreach ($match in $matches) {
            $codes += $match.Groups[1].Value
        }
    }
    
    return ($codes | Sort-Object -Unique)
}

function Get-ConversionMapping {
    return @{
        'USA' = 'US'
        'Europe' = 'EU'
        'Japan' = 'JP'
        'United Kingdom' = 'UK'
        'Australia' = 'AU'
        'Canada' = 'CA'
        'Germany' = 'DE'
        'France' = 'FR'
        'Spain' = 'ES'
        'Italy' = 'IT'
        'Netherlands' = 'NL'
        'Sweden' = 'SE'
        'Korea' = 'KR'
        'China' = 'CN'
        'Brazil' = 'BR'
    }
}

function Select-Conversions {
    param(
        [array]$AvailableCodes,
        [hashtable]$ConversionMapping
    )
    
    Write-Host "`nAvailable country/region codes found in files:" -ForegroundColor Cyan
    for ($i = 0; $i -lt $AvailableCodes.Length; $i++) {
        $code = $AvailableCodes[$i]
        $suggested = if ($ConversionMapping.ContainsKey($code)) { $ConversionMapping[$code] } else { "No suggestion" }
        Write-Host "$($i + 1). $code -> $suggested"
    }
    
    Write-Host "`nConversion options:" -ForegroundColor Yellow
    Write-Host "1. Apply all suggested conversions"
    Write-Host "2. Select specific conversions"
    Write-Host "3. Add custom conversions"
    Write-Host "4. Cancel"
    
    $choice = Read-Host "`nEnter your choice (1-4)"
    $selectedConversions = @{}
    
    switch ($choice) {
        '1' {
            # Apply all available conversions
            foreach ($code in $AvailableCodes) {
                if ($ConversionMapping.ContainsKey($code)) {
                    $selectedConversions[$code] = $ConversionMapping[$code]
                }
            }
        }
        '2' {
            # Select specific conversions
            Write-Host "`nEnter the numbers of conversions to apply (comma-separated, e.g., 1,3,5):"
            $selections = Read-Host "Selection"
            $selectionArray = $selections -split ',' | ForEach-Object { $_.Trim() }
            
            foreach ($sel in $selectionArray) {
                try {
                    $idx = [int]$sel - 1
                    if ($idx -ge 0 -and $idx -lt $AvailableCodes.Length) {
                        $code = $AvailableCodes[$idx]
                        if ($ConversionMapping.ContainsKey($code)) {
                            $selectedConversions[$code] = $ConversionMapping[$code]
                        }
                    }
                } catch {
                    Write-Host "Invalid selection: $sel" -ForegroundColor Red
                }
            }
        }
        '3' {
            # Add custom conversions
            Write-Host "`nEnter custom conversions (format: 'old=new', use 'old=' to remove, empty line to finish):"
            do {
                $conversion = Read-Host "Conversion (old=new)"
                if ($conversion -ne '') {
                    if ($conversion -match '^(.+)=(.*)$') {  # Changed to allow empty values
                        $old = $matches[1].Trim()
                        $new = $matches[2].Trim()
                        if ($AvailableCodes -contains $old) {
                            $selectedConversions[$old] = $new
                            if ($new -eq '') {
                                Write-Host "Will remove: ($old)" -ForegroundColor Green
                            } else {
                                Write-Host "Will convert: ($old) -> ($new)" -ForegroundColor Green
                            }
                        } else {
                            Write-Host "Warning: '$old' not found in available codes" -ForegroundColor Yellow
                        }
                    } else {
                        Write-Host "Invalid format. Use 'old=new' or 'old=' to remove" -ForegroundColor Red
                    }
                }
            } while ($conversion -ne '')
        }
        '4' {
            Write-Host "Operation cancelled." -ForegroundColor Yellow
            return $null
        }
        default {
            Write-Host "Invalid choice." -ForegroundColor Red
            return $null
        }
    }
    
    return $selectedConversions
}

function Get-PreviewChanges {
    param(
        [string]$Directory,
        [hashtable]$Conversions
    )
    
    $changes = @()
    
    Get-ChildItem -Path $Directory -File | ForEach-Object {
        $oldName = $_.Name
        $newName = $oldName
        
        foreach ($conversion in $Conversions.GetEnumerator()) {
            $oldCode = $conversion.Key
            $newCode = $conversion.Value
            
            if ($newCode -eq '' -or $newCode -eq $null) {
                # Remove the entire parentheses group
                $pattern = [regex]::Escape("($oldCode)")
                $newName = $newName -replace $pattern, ''
            } else {
                # Replace with new code
                $pattern = [regex]::Escape("($oldCode)")
                $replacement = "($newCode)"
                $newName = $newName -replace $pattern, $replacement
            }
        }
        
        # Clean up any double spaces that might result from removals
        $newName = $newName -replace '\s+', ' '
        $newName = $newName.Trim()
        
        if ($newName -ne $oldName) {
            $changes += [PSCustomObject]@{
                OldName = $oldName
                NewName = $newName
                FullOldPath = $_.FullName
                FullNewPath = Join-Path $Directory $newName
            }
        }
    }
    
    return $changes
}

function Invoke-FileRename {
    param(
        [array]$Changes,
        [bool]$DryRun = $true
    )
    
    $successCount = 0
    $errorCount = 0
    
    foreach ($change in $Changes) {
        try {
            if (-not $DryRun) {
                # Check if target file already exists
                if (Test-Path $change.FullNewPath) {
                    Write-Host "Error: Target file already exists: $($change.NewName)" -ForegroundColor Red
                    $errorCount++
                    continue
                }
                
                Rename-Item -Path $change.FullOldPath -NewName $change.NewName
            }
            
            $prefix = if ($DryRun) { "[DRY RUN] " } else { "" }
            Write-Host "$prefix$($change.OldName) -> $($change.NewName)" -ForegroundColor Green
            $successCount++
            
        } catch {
            Write-Host "Error renaming $($change.OldName): $($_.Exception.Message)" -ForegroundColor Red
            $errorCount++
        }
    }
    
    return @{
        Success = $successCount
        Errors = $errorCount
    }
}

# Main Script
function Start-CountryCodeRenamer {
    Write-Host "=== File Country Code Renamer ===" -ForegroundColor Magenta
    Write-Host ""
    
    # Get source directory
    do {
        $directory = Read-Host "Enter the source directory path"
        $directory = $directory.Trim('"')
        
        if (-not (Test-Path $directory -PathType Container)) {
            Write-Host "Invalid directory. Please try again." -ForegroundColor Red
        }
    } while (-not (Test-Path $directory -PathType Container))
    
    Write-Host "`nScanning directory: $directory" -ForegroundColor Cyan
    
    # Find available country codes
    $availableCodes = Find-CountryCodes -Directory $directory
    
    if ($availableCodes.Count -eq 0) {
        Write-Host "No country/region codes found in parentheses." -ForegroundColor Yellow
        return
    }
    
    # Get conversion mapping
    $conversionMapping = Get-ConversionMapping
    
    # Let user select conversions
    $selectedConversions = Select-Conversions -AvailableCodes $availableCodes -ConversionMapping $conversionMapping
    
    if ($null -eq $selectedConversions -or $selectedConversions.Count -eq 0) {
        Write-Host "No conversions selected. Exiting." -ForegroundColor Yellow
        return
    }
    
    Write-Host "`nSelected conversions:" -ForegroundColor Cyan
    foreach ($conversion in $selectedConversions.GetEnumerator()) {
        if ($conversion.Value -eq '') {
            Write-Host "  Remove: ($($conversion.Key))" -ForegroundColor Red
        } else {
            Write-Host "  Convert: ($($conversion.Key)) -> ($($conversion.Value))" -ForegroundColor Green
        }
    }
    
    # Preview changes
    $changes = Get-PreviewChanges -Directory $directory -Conversions $selectedConversions
    
    if ($changes.Count -eq 0) {
        Write-Host "`nNo files need to be renamed." -ForegroundColor Yellow
        return
    }
    
    Write-Host "`nPreview of changes ($($changes.Count) files):" -ForegroundColor Cyan
    Write-Host ("-" * 50)
    
    # Show preview (dry run)
    $null = Invoke-FileRename -Changes $changes -DryRun $true
    
    # Confirm execution
    Write-Host "`nReady to rename $($changes.Count) files." -ForegroundColor Yellow
    $confirm = Read-Host "Proceed with renaming? (y/N)"
    
    if ($confirm.ToLower() -eq 'y') {
        Write-Host "`nRenaming files..." -ForegroundColor Cyan
        $result = Invoke-FileRename -Changes $changes -DryRun $false
        Write-Host "`nCompleted: $($result.Success) successful, $($result.Errors) errors" -ForegroundColor Green
    } else {
        Write-Host "Operation cancelled." -ForegroundColor Yellow
    }
}

# Run the script
Start-CountryCodeRenamer