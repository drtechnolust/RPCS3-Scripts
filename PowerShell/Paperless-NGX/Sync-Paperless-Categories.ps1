# ============================================
# Paperless-ngx Custom Field Sync Script
# - Creates missing custom fields
# - Updates existing custom fields by name
# - Supports select fields with options
# ============================================

$PaperlessUrl = "http://10.10.1.99:8000"
$Token        = "f9c252393faa9925f8998d8427657bdd3c73aa46"

$Headers = @{
    Authorization = "Token $Token"
    "Content-Type" = "application/json"
}

function Invoke-PaperlessGetAll {
    param(
        [string]$Endpoint
    )

    $results = @()
    $nextUrl = "$PaperlessUrl$Endpoint"

    while ($nextUrl) {
        $response = Invoke-RestMethod -Uri $nextUrl -Method Get -Headers $Headers -ErrorAction Stop
        if ($null -ne $response.results) {
            $results += $response.results
            $nextUrl = $response.next
        }
        else {
            return $response
        }
    }

    return $results
}

function Get-CustomFieldLookup {
    $allFields = Invoke-PaperlessGetAll -Endpoint "/api/custom_fields/?page_size=1000"
    $lookup = @{}

    foreach ($field in $allFields) {
        $lookup[$field.name] = $field
    }

    return $lookup
}

function Normalize-SelectOptions {
    param(
        [array]$Options
    )

    if ($null -eq $Options) { return @() }

    $normalized = @()

    foreach ($opt in $Options) {
        if ($opt -is [string]) {
            $normalized += @{ label = $opt }
        }
        elseif ($opt -is [hashtable]) {
            if ($opt.ContainsKey("label")) {
                $normalized += @{ label = [string]$opt.label }
            }
        }
        elseif ($opt.PSObject.Properties.Name -contains "label") {
            $normalized += @{ label = [string]$opt.label }
        }
    }

    return $normalized
}

function New-OrUpdate-CustomField {
    param(
        [hashtable]$Definition,
        [hashtable]$ExistingLookup
    )

    $name = $Definition.name

    $body = @{
        name      = $Definition.name
        data_type = $Definition.data_type
    }

    if ($Definition.ContainsKey("extra_data") -and $null -ne $Definition.extra_data) {
        $body["extra_data"] = $Definition.extra_data
    }

    $json = $body | ConvertTo-Json -Depth 10

    if ($ExistingLookup.ContainsKey($name)) {
        $existing = $ExistingLookup[$name]
        $fieldId = $existing.id

        # Important: Paperless docs note custom field data type cannot be changed after creation.
        # So we only PATCH if the data_type matches or if you want to update metadata/options only.
        $existingType = $existing.data_type
        if ($existingType -ne $Definition.data_type) {
            Write-Host "Skipping type change for '$name' (existing: $existingType, desired: $($Definition.data_type))" -ForegroundColor Yellow
            return
        }

        try {
            Invoke-RestMethod -Uri "$PaperlessUrl/api/custom_fields/$fieldId/" -Method Patch -Headers $Headers -Body $json -ErrorAction Stop | Out-Null
            Write-Host "Updated custom field: $name" -ForegroundColor Cyan
        }
        catch {
            Write-Host "Failed updating custom field: $name" -ForegroundColor Red
            Write-Host $_.Exception.Message
        }
    }
    else {
        try {
            Invoke-RestMethod -Uri "$PaperlessUrl/api/custom_fields/" -Method Post -Headers $Headers -Body $json -ErrorAction Stop | Out-Null
            Write-Host "Created custom field: $name" -ForegroundColor Green
        }
        catch {
            Write-Host "Failed creating custom field: $name" -ForegroundColor Red
            Write-Host $_.Exception.Message
        }
    }
}

# --------------------------------------------
# Custom field definitions
# Common safe set for your environment
# --------------------------------------------
# Notes:
# - Keep custom fields minimal and structured
# - "Category" is intentionally omitted from the target list
# - Select fields use extra_data.select_options
# --------------------------------------------

$CustomFieldDefinitions = @(
    @{
        name = "Review Status"
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
    },
    @{
        name = "Retention Class"
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
    },
    @{
        name = "Document Quality"
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
    },
    @{
        name = "Received Method"
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
    },
    @{
        name = "Confidentiality"
        data_type = "select"
        extra_data = @{
            select_options = @(
                @{ label = "Normal" }
                @{ label = "Sensitive" }
                @{ label = "Restricted" }
            )
        }
    },

    @{ name = "Source Folder";     data_type = "string";  extra_data = $null }
    @{ name = "Legacy Filename";   data_type = "string";  extra_data = $null }
    @{ name = "Account Number";    data_type = "string";  extra_data = $null }
    @{ name = "Policy Number";     data_type = "string";  extra_data = $null }
    @{ name = "Claim Number";      data_type = "string";  extra_data = $null }
    @{ name = "Case or Matter ID"; data_type = "string";  extra_data = $null }
    @{ name = "Project Code";      data_type = "string";  extra_data = $null }
    @{ name = "Owner Department";  data_type = "string";  extra_data = $null }

    @{ name = "Expiration Date";   data_type = "date";    extra_data = $null }
    @{ name = "Renewal Date";      data_type = "date";    extra_data = $null }
    @{ name = "Effective Date";    data_type = "date";    extra_data = $null }

    @{ name = "Invoice Number";    data_type = "string";  extra_data = $null }
    @{ name = "Check Number";      data_type = "string";  extra_data = $null }
    @{ name = "Amount";            data_type = "monetary"; extra_data = $null }
)

Write-Host "Reading existing custom fields..." -ForegroundColor White
$ExistingFields = Get-CustomFieldLookup

Write-Host "Creating or updating custom fields..." -ForegroundColor White
foreach ($definition in $CustomFieldDefinitions) {
    New-OrUpdate-CustomField -Definition $definition -ExistingLookup $ExistingFields
}

Write-Host "Done." -ForegroundColor Green
Write-Host ""
Write-Host "Next step: review the Custom Fields screen in Paperless." -ForegroundColor White
Write-Host "If 'Category' is still present and you don't want it, delete it manually after confirming the new fields exist." -ForegroundColor Yellow