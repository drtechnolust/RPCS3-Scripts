# PARAM.SFO Debug Tester
# Paste path to a single PARAM.SFO to see what the parser reads from it

$sfoPath = Read-Host "Paste full path to a PARAM.SFO file"

if (-not (Test-Path $sfoPath)) {
    Write-Host "File not found" -ForegroundColor Red
    exit
}

$bytes = [System.IO.File]::ReadAllBytes($sfoPath)

Write-Host ""
Write-Host "Magic bytes: $($bytes[0]) $($bytes[1]) $($bytes[2]) $($bytes[3])" -ForegroundColor Cyan
Write-Host "Expected:    0  80 83 70" -ForegroundColor Gray
Write-Host ""

$keyTableOffset  = [BitConverter]::ToInt32($bytes, 8)
$dataTableOffset = [BitConverter]::ToInt32($bytes, 12)
$entryCount      = [BitConverter]::ToInt32($bytes, 16)

Write-Host "Key table offset  : $keyTableOffset"
Write-Host "Data table offset : $dataTableOffset"
Write-Host "Entry count       : $entryCount"
Write-Host ""

for ($i = 0; $i -lt $entryCount; $i++) {
    $entryBase  = 20 + ($i * 16)
    $keyOffset  = [BitConverter]::ToInt16($bytes, $entryBase)
    $fmtLow     = $bytes[$entryBase + 2]
    $fmtHigh    = $bytes[$entryBase + 3]
    $dataLen    = [BitConverter]::ToInt32($bytes, $entryBase + 4)
    $dataMaxLen = [BitConverter]::ToInt32($bytes, $entryBase + 8)
    $dataOffset = [BitConverter]::ToInt32($bytes, $entryBase + 12)

    # Read key
    $keyStart = $keyTableOffset + $keyOffset
    $keyEnd   = $keyStart
    while ($keyEnd -lt $bytes.Length -and $bytes[$keyEnd] -ne 0) { $keyEnd++ }
    $key = [System.Text.Encoding]::ASCII.GetString($bytes, $keyStart, $keyEnd - $keyStart)

    # Read value as string regardless
    $valStart = $dataTableOffset + $dataOffset
    $valStr   = [System.Text.Encoding]::UTF8.GetString($bytes, $valStart, $dataLen).TrimEnd([char]0)
    $valInt   = if ($dataLen -ge 4) { [BitConverter]::ToInt32($bytes, $valStart) } else { 0 }

    Write-Host "[$i] Key='$key'  FmtLow=0x$($fmtLow.ToString('X2'))  FmtHigh=0x$($fmtHigh.ToString('X2'))  Len=$dataLen  StrVal='$valStr'  IntVal=$valInt" -ForegroundColor Yellow
}