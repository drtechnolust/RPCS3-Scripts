#Requires -Version 5.1
<#
.SYNOPSIS
    Syncs custom field definitions to a Paperless-NGX instance via the REST API.

.DESCRIPTION
    Creates missing custom fields and updates existing ones by name.
    Supports select fields with predefined options.
    Safe to rerun -- skips fields that already match.
    Does not delete existing fields.

.PARAMETER DryRun
    When set to $true, shows what would be created or updated without making any API calls.
    Defaults to $true. Set to $false in the config block to apply changes.

.EXAMPLE
    .\Sync-Paperless-Categories.ps1

.NOTES
    Before running:
    1. Paste your Paperless-NGX API token into the $Token field in the CONFIG block below.
    2. Confirm your Paperless-NGX URL in the $PaperlessUrl field.
    3. Run with $DryRun = $true first to preview changes.
    4. Set $DryRun = $false to apply.

    To generate a token: Paperless-NGX -> Your Profile -> API Token -> Create / Regenerate.
    Never commit a real token to source control.

.VERSION
    1.1.0 - Moved credentials to config block. Added DryRun. Added MIT header.
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
# Helper: page through a Paperless-NGX list endpoint
# ------------------------------------------------------------------------------
function Invoke-PaperlessGetAll {
    param([string]$Endpoint)

    $Results = @()
    $NextUrl = "$PaperlessUrl$Endpoint"

    while ($NextUrl) {
        $Response = Invoke-RestMethod -Uri $NextUrl -Method Get -Headers $Headers -ErrorAction Stop
        if ($null -ne $Response.results) {
            $Results += $Response.results
            $NextUrl = $Response.next
        }
        else {
            return $Response
        }
    }

    return $Results
}

# ------------------------------------------------------------------------------
# Helper: build a name-keyed lookup of all existing custom fields
# ------------------------------------------------------------------------------
function Get-CustomFieldLookup {
    $AllFields = Invoke-PaperlessGetAll -Endpoint "/api/custom_fields/?page_size=1000"
    $Lookup = @{}
    foreach ($Field in $AllFields) {
        if ($null -ne $Field -and -not [string]::IsNullOrWhiteSpace($Field.name)) {
            $Lookup[$Field.name] = $Field
        }
    }
    return $Lookup
}

# ------------------------------------------------------------------------------
# Helper: normalize select options to @{ label = "..." } objects
# ------------------------------------------------------------------------------
function Normalize-SelectOptions {
    param([array]$Options)

    if ($null -eq $Options) { return @() }

    $Normalized = @()
    foreach ($Opt in $Options) {
        if ($Opt -is [string]) {
            $Normalized += @{ label = $Opt }
        }
        elseif ($Opt -is [hashtable] -and $Opt.ContainsKey("label")) {
            $Normalized += @{ label = [string]$Opt.label }
        }
        elseif ($Opt.PSObject.Properties.Name -contains "label") {
            $Normalized += @{ label = [string]$Opt.label }
        }
    }
    return $Normalized
}

# ------------------------------------------------------------------------------
# Core: create or update a single custom field
# ------------------------------------------------------------------------------
function Invoke-SyncCustomField {
    param(
        [hashtable]$Definition,
        [hashtable]$ExistingLookup
    )

    $Name = $Definition.name

    $Body = @{
        name      = $Definition.name
        data_type = $Definition.data_type
    }

    if ($Definition.ContainsKey("extra_data") -and $null -ne $Definition.extra_data) {
        $Body["extra_data"] = $Definition.extra_data
    }

    if ($ExistingLookup.ContainsKey($Name)) {
        $Existing = $ExistingLookup[$Name]

        # Paperless-NGX does not allow changing data_type after creation
        if ($Existing.data_type -ne $Definition.data_type) {
            Write-Host "  SKIP  Type change not allowed: $Name (existing=$($Existing.data_type), desired=$($Definition.data_type))" -ForegroundColor Yellow
            $Summary.Skipped++
            return
        }

        if ($DryRun) {
            Write-Host "  WOULD UPDATE  $Name" -ForegroundColor Yellow
            return
        }

        try {
            $Json = $Body | ConvertTo-Json -Depth 10
            Invoke-RestMethod -Uri "$PaperlessUrl/api/custom_fields/$($Existing.id)/" -Method Patch -Headers $Headers -Body $Json -ErrorAction Stop | Out-Null
            Write-Host "  Updated  $Name" -ForegroundColor Cyan
            $Summary.Updated++
        }
        catch {
            Write-Host "  FAILED updating  $Name -- $($_.Exception.Message)" -ForegroundColor Red
            $Summary.Failed++
        }
    }
    else {
        if ($DryRun) {
            Write-Host "  WOULD CREATE  $Name ($($Definition.data_type))" -ForegroundColor Green
            return
        }

        try {
            $Json = $Body | ConvertTo-Json -Depth 10
            Invoke-RestMethod -Uri "$PaperlessUrl/api/custom_fields/" -Method Post -Headers $Headers -Body $Json -ErrorAction Stop | Out-Null
            Write-Host "  Created  $Name" -ForegroundColor Green
            $Summary.Created++
        }
        catch {
            Write-Host "  FAILED creating  $Name -- $($_.Exception.Message)" -ForegroundColor Red
            $Summary.Failed++
        }
    }
}

# ==============================================================================
# CUSTOM FIELD DEFINITIONS
# Add, remove, or edit entries here. Safe to rerun -- existing fields are patched,
# not duplicated. data_type cannot be changed after a field is created.
# ==============================================================================
$CustomFieldDefinitions = @(

    @{
        name      = "Review Status"
        data_type = "select"
        extra_data = @{
            select_options = @(
                @{ label = "New" }
                @{ label = "Reviewed" }
                @{ label = "Needs Review" }
                @{ label = "Needs OCR Fix" }
                @{ label = "Needs Metadata Fix" }
                @{ label = "Approved" }
            )
        }
    }

    @{
        name      = "Retention Class"
        data_type = "select"
        extra_data = @{
            select_options = @(
                @{ label = "Permanent" }
                @{ label = "1 Year" }
                @{ label = "3 Years" }
                @{ label = "7 Years" }
                @{ label = "10 Years" }
                @{ label = "Archive Only" }
                @{ label = "Review for Deletion" }
            )
        }
    }

    @{
        name      = "Document Quality"
        data_type = "select"
        extra_data = @{
            select_options = @(
                @{ label = "Good" }
                @{ label = "OCR Issues" }
                @{ label = "Poor Scan" }
                @{ label = "Incomplete" }
                @{ label = "Duplicate Suspected" }
            )
        }
    }

    @{
        name      = "Received Method"
        data_type = "select"
        extra_data = @{
            select_options = @(
                @{ label = "Scan" }
                @{ label = "Email" }
                @{ label = "Upload" }
                @{ label = "Bulk Import" }
                @{ label = "Mobile App" }
                @{ label = "Filesystem Import" }
            )
        }
    }

    @{
        name      = "Confidentiality"
        data_type = "select"
        extra_data = @{
            select_options = @(
                @{ label = "Normal" }
                @{ label = "Sensitive" }
                @{ label = "Restricted" }
            )
        }
    }

    @{ name = "Source Folder";     data_type = "string";   extra_data = $null }
    @{ name = "Legacy Filename";   data_type = "string";   extra_data = $null }
    @{ name = "Account Number";    data_type = "string";   extra_data = $null }
    @{ name = "Policy Number";     data_type = "string";   extra_data = $null }
    @{ name = "Claim Number";      data_type = "string";   extra_data = $null }
    @{ name = "Case or Matter ID"; data_type = "string";   extra_data = $null }
    @{ name = "Project Code";      data_type = "string";   extra_data = $null }
    @{ name = "Owner Department";  data_type = "string";   extra_data = $null }
    @{ name = "Invoice Number";    data_type = "string";   extra_data = $null }
    @{ name = "Check Number";      data_type = "string";   extra_data = $null }

    @{ name = "Expiration Date";   data_type = "date";     extra_data = $null }
    @{ name = "Renewal Date";      data_type = "date";     extra_data = $null }
    @{ name = "Effective Date";    data_type = "date";     extra_data = $null }

    @{ name = "Amount";            data_type = "monetary"; extra_data = $null }
)

# ==============================================================================
# MAIN
# ==============================================================================
Write-Host ""
Write-Host "Paperless-NGX Custom Field Sync" -ForegroundColor Cyan
Write-Host "================================" -ForegroundColor Cyan
Write-Host "  URL     : $PaperlessUrl"
Write-Host "  DryRun  : $DryRun"
Write-Host "  Fields  : $($CustomFieldDefinitions.Count)"
Write-Host ""

if ($DryRun) {
    Write-Host "  [DRY RUN] No changes will be made." -ForegroundColor Yellow
    Write-Host ""
}

Write-Host "Reading existing custom fields..." -ForegroundColor White
$ExistingFields = Get-CustomFieldLookup
Write-Host "  Found $($ExistingFields.Count) existing field(s)." -ForegroundColor Gray
Write-Host ""

Write-Host "Syncing fields..." -ForegroundColor White
foreach ($Definition in $CustomFieldDefinitions) {
    Invoke-SyncCustomField -Definition $Definition -ExistingLookup $ExistingFields
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
    Write-Host "  Review the Custom Fields screen in Paperless to confirm results." -ForegroundColor Gray
}
Write-Host ""
