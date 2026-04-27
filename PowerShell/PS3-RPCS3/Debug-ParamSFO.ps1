#Requires -Version 5.1
<#
.SYNOPSIS
    Diagnostic tool: reads a single PARAM.SFO file and dumps all fields to console.

.DESCRIPTION
    Prompts for the path to a PARAM.SFO file, validates the magic bytes,
    and prints every indexed field with its format type, data length,
    raw string value, and raw integer value.

    Use this to verify that Read-ParamSfo parser logic in other scripts
    is reading the correct field offsets for a given game.

.EXAMPLE
    .\Debug-ParamSFO.ps1

.NOTES
    This is a diagnostic/development tool. It is not part of the automated
    library pipeline. Run it against a specific PARAM.SFO when investigating
    why a game is not being detected correctly.

.VERSION
    1.1.0 - MIT license header, replaced bare exit with Set-ExitCode, added help block.
    1.0.0 - Initial release.

.LICENSE
    MIT License
    Copyright (c) Paul Mardis
#>

function Set-ExitCode {
    param([int]$Code)
    $global:LASTEXITCODE = $Code
}

Write-Host ""
Write-Host "Debug-ParamSFO" -ForegroundColor Cyan
Write-Host "==============" -ForegroundColor Cyan
Write-Host ""

$SfoPath = (Read-Host "Paste full path to a PARAM.SFO file").Trim().Trim('"').Trim("'")

if (-not (Test-Path -LiteralPath $SfoPath -PathType Leaf)) {
    Write-Host ("ERROR: File not found: {0}" -f $SfoPath) -ForegroundColor Red
    Set-ExitCode 1
    return
}

try {
    $Bytes = [System.IO.File]::ReadAllBytes($SfoPath)
}
catch {
    Write-Host ("ERROR: Could not read file: {0}" -f $_.Exception.Message) -ForegroundColor Red
    Set-ExitCode 1
    return
}

Write-Host ""
Write-Host ("Magic bytes: {0} {1} {2} {3}" -f $Bytes[0], $Bytes[1], $Bytes[2], $Bytes[3]) -ForegroundColor Cyan
Write-Host ("Expected   :  0  80  83  70  (0x00 P S F)") -ForegroundColor Gray

if ($Bytes[0] -ne 0x00 -or $Bytes[1] -ne 0x50 -or $Bytes[2] -ne 0x53 -or $Bytes[3] -ne 0x46) {
    Write-Host ""
    Write-Host "WARNING: Magic bytes do not match. This may not be a valid PARAM.SFO." -ForegroundColor Yellow
}

$KeyTableOffset  = [BitConverter]::ToInt32($Bytes, 8)
$DataTableOffset = [BitConverter]::ToInt32($Bytes, 12)
$EntryCount      = [BitConverter]::ToInt32($Bytes, 16)

Write-Host ""
Write-Host ("Key table offset  : {0} (0x{0:X})" -f $KeyTableOffset)
Write-Host ("Data table offset : {0} (0x{0:X})" -f $DataTableOffset)
Write-Host ("Entry count       : {0}" -f $EntryCount)
Write-Host ""
Write-Host "Fields:" -ForegroundColor White
Write-Host ("-" * 80) -ForegroundColor DarkGray

for ($i = 0; $i -lt $EntryCount; $i++) {
    $EntryBase  = 20 + ($i * 16)
    $KeyOffset  = [BitConverter]::ToInt16($Bytes, $EntryBase)
    $FmtLow     = $Bytes[$EntryBase + 2]
    $FmtHigh    = $Bytes[$EntryBase + 3]
    $DataLen    = [BitConverter]::ToInt32($Bytes, $EntryBase + 4)
    $DataMaxLen = [BitConverter]::ToInt32($Bytes, $EntryBase + 8)
    $DataOffset = [BitConverter]::ToInt32($Bytes, $EntryBase + 12)

    $KeyStart = $KeyTableOffset + $KeyOffset
    $KeyEnd   = $KeyStart
    while ($KeyEnd -lt $Bytes.Length -and $Bytes[$KeyEnd] -ne 0) { $KeyEnd++ }
    $Key = [System.Text.Encoding]::ASCII.GetString($Bytes, $KeyStart, $KeyEnd - $KeyStart)

    $ValStart = $DataTableOffset + $DataOffset
    $StrVal   = [System.Text.Encoding]::UTF8.GetString($Bytes, $ValStart, $DataLen).TrimEnd([char]0)
    $IntVal   = if ($DataLen -ge 4) { [BitConverter]::ToInt32($Bytes, $ValStart) } else { 0 }

    $FmtLabel = switch ($FmtHigh) {
        0x02 { "UTF8-string" }
        0x04 { "Int32" }
        0x00 { "Binary" }
        default { "Unknown(0x{0:X2})" -f $FmtHigh }
    }

    Write-Host ("[{0:D2}] {1,-18} | {2,-12} | Len={3,4} | StrVal='{4}'" -f `
        $i, $Key, $FmtLabel, $DataLen, $StrVal) -ForegroundColor Yellow

    if ($FmtHigh -eq 0x04) {
        Write-Host ("      IntVal = {0}" -f $IntVal) -ForegroundColor Gray
    }
}

Write-Host ("-" * 80) -ForegroundColor DarkGray
Write-Host ""
Write-Host "Done." -ForegroundColor Green
Write-Host ""
Set-ExitCode 0
