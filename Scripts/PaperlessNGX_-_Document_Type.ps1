#Requires -Version 5.1
<#
.SYNOPSIS
    Syncs document type definitions to a Paperless-NGX instance via the REST API.

.DESCRIPTION
    Creates missing document types and updates existing ones by name.
    Skips entries that are already correct to avoid unnecessary PATCH calls.
    Safe to rerun. Does not delete existing document types.

.PARAMETER DryRun
    When set to $true, shows what would be created or updated without making any API calls.
    Defaults to $true. Set to $false in the config block to apply changes.

.EXAMPLE
    .\Sync-Paperless-DocumentTypes.ps1

.NOTES
    Before running:
    1. Paste your Paperless-NGX API token into the $Token field in the CONFIG block below.
    2. Confirm your Paperless-NGX URL in the $PaperlessUrl field.
    3. Run with $DryRun = $true first to preview changes.
    4. Set $DryRun = $false to apply.

    To generate a token: Paperless-NGX -> Your Profile -> API Token -> Create / Regenerate.
    Never commit a real token to source control.

.VERSION
    1.1.0 - Moved credentials to config block. DryRun now defaults to $true. Added MIT header.
    1.0.0 - Initial release.

.LICENSE
    MIT License
    Copyright (c) Paul Mardis
#>

# ==============================================================================
# CONFIG -- Edit this block. Do not put credentials anywhere else in this file.
# ==============================================================================
$PaperlessUrl = "http://10.10.1.99:8000"   # Update if your URL or port changes
$Token        = ""                          # Paste your API token here. Never commit a real token.
$DryRun       = $true                       # Set to $false to apply changes
# ==============================================================================

if ([string]::IsNullOrWhiteSpace($Token)) {
    Write-Host "ERROR: No API token configured." -ForegroundColor Red
    Write-Host "       Open this script and paste your token into the `$Token field in the CONFIG block." -ForegroundColor Yellow
    exit 1
}

$Headers = @{
    Authorization  = "Token $Token"
    "Content-Type" = "application/json"
}

$Summary = [ordered]@{
    Created = 0
    Updated = 0
    Skipped = 0
    Failed  = 0
}

# ------------------------------------------------------------------------------
# Helper: extract readable error detail from a Paperless API error response
# ------------------------------------------------------------------------------
function Get-ApiErrorDetail {
    param($ErrorRecord)
    try {
        $Response = $ErrorRecord.Exception.Response
        if ($Response -and $Response.GetResponseStream) {
            $Reader = New-Object System.IO.StreamReader($Response.GetResponseStream())
            $Reader.BaseStream.Position = 0
            $Reader.DiscardBufferedData()
            return $Reader.ReadToEnd()
        }
    } catch {}
    return $ErrorRecord.Exception.Message
}

# ------------------------------------------------------------------------------
# Helper: page through a Paperless-NGX list endpoint
# ------------------------------------------------------------------------------
function Invoke-PaperlessGetAll {
    param([string]$Endpoint)

    $Results = @()
    $NextUrl = "$PaperlessUrl$Endpoint"

    while ($NextUrl) {
        $Response = Invoke-RestMethod -Uri $NextUrl -Method Get -Headers $Headers -ErrorAction Stop

        if ($null -ne $Response.results) {
            foreach ($Item in $Response.results) {
                if ($null -ne $Item) { $Results += $Item }
            }
            $NextUrl = $Response.next
        }
        else {
            return $Response
        }
    }

    return $Results
}

# ------------------------------------------------------------------------------
# Helper: build a name-keyed lookup of all existing document types
# ------------------------------------------------------------------------------
function Get-ExistingDocumentTypesByName {
    $AllTypes = Invoke-PaperlessGetAll -Endpoint "/api/document_types/?page_size=1000"
    $Lookup = @{}
    foreach ($Type in $AllTypes) {
        if ($null -eq $Type) { continue }
        if (-not ($Type.PSObject.Properties.Name -contains "name")) { continue }
        if ([string]::IsNullOrWhiteSpace($Type.name)) { continue }
        $Lookup[$Type.name] = $Type
    }
    return $Lookup
}

# ------------------------------------------------------------------------------
# Helper: normalize any value to a trimmed string for comparison
# ------------------------------------------------------------------------------
function Normalize-String {
    param([object]$Value)
    if ($null -eq $Value) { return "" }
    return ([string]$Value).Trim()
}

# ------------------------------------------------------------------------------
# Core: create or update a single document type
# ------------------------------------------------------------------------------
function Invoke-SyncDocumentType {
    param(
        [hashtable]$Definition,
        [hashtable]$ExistingTypes,
        [int]$Index,
        [int]$Total
    )

    $Name   = $Definition.name
    $Prefix = "[$Index/$Total]"

    $Body = @{ name = $Definition.name }

    if ($Definition.ContainsKey("matching_algorithm") -and $null -ne $Definition.matching_algorithm) {
        $Body["matching_algorithm"] = $Definition.matching_algorithm
    }
    if ($Definition.ContainsKey("matching") -and $null -ne $Definition.matching) {
        $Body["matching"] = $Definition.matching
    }

    if ($ExistingTypes.ContainsKey($Name)) {
        $Existing    = $ExistingTypes[$Name]
        $NeedsPatch  = $false

        if ((Normalize-String $Existing.name) -ne (Normalize-String $Definition.name)) {
            $NeedsPatch = $true
        }
        if ($Definition.ContainsKey("matching_algorithm") -and $null -ne $Definition.matching_algorithm) {
            if ($Existing.matching_algorithm -ne $Definition.matching_algorithm) { $NeedsPatch = $true }
        }
        if ($Definition.ContainsKey("matching")) {
            if ((Normalize-String $Existing.matching) -ne (Normalize-String $Definition.matching)) { $NeedsPatch = $true }
        }

        if (-not $NeedsPatch) {
            Write-Host "$Prefix Skipped (unchanged): $Name" -ForegroundColor DarkGray
            $Summary.Skipped++
            return
        }

        if ($DryRun) {
            Write-Host "$Prefix WOULD UPDATE: $Name" -ForegroundColor Yellow
            return
        }

        try {
            $Json = $Body | ConvertTo-Json -Depth 10
            Invoke-RestMethod -Uri "$PaperlessUrl/api/document_types/$($Existing.id)/" -Method Patch -Headers $Headers -Body $Json -ErrorAction Stop | Out-Null
            Write-Host "$Prefix Updated: $Name" -ForegroundColor Cyan
            $Summary.Updated++
        }
        catch {
            Write-Host "$Prefix FAILED updating: $Name" -ForegroundColor Red
            Write-Host (Get-ApiErrorDetail $_)
            $Summary.Failed++
        }
    }
    else {
        if ($DryRun) {
            Write-Host "$Prefix WOULD CREATE: $Name" -ForegroundColor Green
            return
        }

        try {
            $Json = $Body | ConvertTo-Json -Depth 10
            Invoke-RestMethod -Uri "$PaperlessUrl/api/document_types/" -Method Post -Headers $Headers -Body $Json -ErrorAction Stop | Out-Null
            Write-Host "$Prefix Created: $Name" -ForegroundColor Green
            $Summary.Created++
        }
        catch {
            Write-Host "$Prefix FAILED creating: $Name" -ForegroundColor Red
            Write-Host (Get-ApiErrorDetail $_)
            $Summary.Failed++
        }
    }
}

# ==============================================================================
# DOCUMENT TYPE DEFINITIONS
# Add, remove, or edit entries here. Safe to rerun.
# ==============================================================================
$DocumentTypeDefinitions = @(
    @{ name = "Account Statements";        matching_algorithm = $null; matching = $null }
    @{ name = "Bills";                     matching_algorithm = $null; matching = $null }
    @{ name = "Contracts";                 matching_algorithm = $null; matching = $null }
    @{ name = "Correspondence";            matching_algorithm = $null; matching = $null }
    @{ name = "Delivery Note";             matching_algorithm = $null; matching = $null }
    @{ name = "Forms";                     matching_algorithm = $null; matching = $null }
    @{ name = "Government Documents";      matching_algorithm = $null; matching = $null }
    @{ name = "Identity Documents";        matching_algorithm = $null; matching = $null }
    @{ name = "Insurance Documents";       matching_algorithm = $null; matching = $null }
    @{ name = "Invoices";                  matching_algorithm = $null; matching = $null }
    @{ name = "Legal";                     matching_algorithm = $null; matching = $null }
    @{ name = "Manuals / Instructions";    matching_algorithm = $null; matching = $null }
    @{ name = "Medical Records";           matching_algorithm = $null; matching = $null }
    @{ name = "Orders / Purchase Records"; matching_algorithm = $null; matching = $null }
    @{ name = "Other";                     matching_algorithm = $null; matching = $null }
    @{ name = "Photos / Scans";            matching_algorithm = $null; matching = $null }
    @{ name = "Receipts";                  matching_algorithm = $null; matching = $null }
    @{ name = "Reports";                   matching_algorithm = $null; matching = $null }
    @{ name = "Tax Documents";             matching_algorithm = $null; matching = $null }
    @{ name = "Technical Documentation";   matching_algorithm = $null; matching = $null }
)

# ==============================================================================
# MAIN
# ==============================================================================
Write-Host ""
Write-Host "Paperless-NGX Document Type Sync" -ForegroundColor Cyan
Write-Host "=================================" -ForegroundColor Cyan
Write-Host "  URL     : $PaperlessUrl"
Write-Host "  DryRun  : $DryRun"
Write-Host "  Types   : $($DocumentTypeDefinitions.Count)"
Write-Host ""

if ($DryRun) {
    Write-Host "  [DRY RUN] No changes will be made." -ForegroundColor Yellow
    Write-Host ""
}

Write-Host "Reading existing document types..." -ForegroundColor White
$ExistingTypes = Get-ExistingDocumentTypesByName
Write-Host "  Found $($ExistingTypes.Count) existing type(s)." -ForegroundColor Gray
Write-Host ""

Write-Host "Syncing document types..." -ForegroundColor White
$Total = $DocumentTypeDefinitions.Count
for ($i = 0; $i -lt $Total; $i++) {
    Invoke-SyncDocumentType -Definition $DocumentTypeDefinitions[$i] -ExistingTypes $ExistingTypes -Index ($i + 1) -Total $Total
}

Write-Host ""
Write-Host "Summary" -ForegroundColor White
Write-Host "  Created : $($Summary.Created)"
Write-Host "  Updated : $($Summary.Updated)"
Write-Host "  Skipped : $($Summary.Skipped)"
Write-Host "  Failed  : $($Summary.Failed)"

if ($DryRun) {
    Write-Host ""
    Write-Host "  DRY RUN complete. Set `$DryRun = `$false in the CONFIG block to apply changes." -ForegroundColor Yellow
}
else {
    Write-Host ""
    Write-Host "  Done." -ForegroundColor Green
}
Write-Host ""
