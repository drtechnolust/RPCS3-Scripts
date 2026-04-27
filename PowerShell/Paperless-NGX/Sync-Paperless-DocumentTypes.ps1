# ============================================
# Paperless-ngx Document Type Sync Script
# - Creates missing document types
# - Updates existing ones
# - Safe to rerun
# - Handles null/empty API items safely
# ============================================

$PaperlessUrl = "http://10.10.1.99:8000"
$Token = ""

# Preview only first
$DryRun = $false

$Headers = @{
    Authorization = "Token $Token"
    "Content-Type" = "application/json"
}

$Summary = [ordered]@{
    Created = 0
    Updated = 0
    Skipped = 0
    Failed  = 0
}

function Get-ApiErrorDetail {
    param($ErrorRecord)
    try {
        $response = $ErrorRecord.Exception.Response
        if ($response -and $response.GetResponseStream) {
            $reader = New-Object System.IO.StreamReader($response.GetResponseStream())
            $reader.BaseStream.Position = 0
            $reader.DiscardBufferedData()
            return $reader.ReadToEnd()
        }
    } catch {}
    return $ErrorRecord.Exception.Message
}

function Invoke-PaperlessGetAll {
    param([string]$Endpoint)

    $Results = @()
    $NextUrl = "$PaperlessUrl$Endpoint"

    while ($NextUrl) {
        $Response = Invoke-RestMethod -Uri $NextUrl -Method Get -Headers $Headers -ErrorAction Stop

        if ($null -ne $Response.results) {
            foreach ($Item in $Response.results) {
                if ($null -ne $Item) {
                    $Results += $Item
                }
            }
            $NextUrl = $Response.next
        }
        else {
            return $Response
        }
    }

    return $Results
}

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

function Normalize-String {
    param([object]$Value)
    if ($null -eq $Value) { return "" }
    return ([string]$Value).Trim()
}

function New-OrUpdate-DocumentType {
    param(
        [hashtable]$Definition,
        [hashtable]$ExistingTypes,
        [int]$Index,
        [int]$Total
    )

    $Name = $Definition.name
    $Prefix = "[$Index/$Total]"

    $Body = @{
        name = $Definition.name
    }

    if ($Definition.ContainsKey("matching_algorithm") -and $null -ne $Definition.matching_algorithm) {
        $Body["matching_algorithm"] = $Definition.matching_algorithm
    }

    if ($Definition.ContainsKey("matching") -and $null -ne $Definition.matching) {
        $Body["matching"] = $Definition.matching
    }

    if ($ExistingTypes.ContainsKey($Name)) {
        $Existing = $ExistingTypes[$Name]
        $NeedsPatch = $false

        if ((Normalize-String $Existing.name) -ne (Normalize-String $Definition.name)) {
            $NeedsPatch = $true
        }

        if ($Definition.ContainsKey("matching_algorithm") -and $null -ne $Definition.matching_algorithm) {
            if ($Existing.matching_algorithm -ne $Definition.matching_algorithm) {
                $NeedsPatch = $true
            }
        }

        if ($Definition.ContainsKey("matching")) {
            $ExistingMatching = Normalize-String $Existing.matching
            $DesiredMatching  = Normalize-String $Definition.matching
            if ($ExistingMatching -ne $DesiredMatching) {
                $NeedsPatch = $true
            }
        }

        if (-not $NeedsPatch) {
            Write-Host "$Prefix Skipped unchanged: $Name" -ForegroundColor DarkGray
            $Summary.Skipped++
            return
        }

        if ($DryRun) {
            Write-Host "$Prefix Would update: $Name" -ForegroundColor Yellow
            return
        }

        try {
            $JsonBody = $Body | ConvertTo-Json -Depth 10
            Invoke-RestMethod -Uri "$PaperlessUrl/api/document_types/$($Existing.id)/" -Method Patch -Headers $Headers -Body $JsonBody -ErrorAction Stop | Out-Null
            Write-Host "$Prefix Updated: $Name" -ForegroundColor Cyan
            $Summary.Updated++
        }
        catch {
            Write-Host "$Prefix Failed updating: $Name" -ForegroundColor Red
            Write-Host (Get-ApiErrorDetail $_)
            $Summary.Failed++
        }
    }
    else {
        if ($DryRun) {
            Write-Host "$Prefix Would create: $Name" -ForegroundColor Green
            return
        }

        try {
            $JsonBody = $Body | ConvertTo-Json -Depth 10
            Invoke-RestMethod -Uri "$PaperlessUrl/api/document_types/" -Method Post -Headers $Headers -Body $JsonBody -ErrorAction Stop | Out-Null
            Write-Host "$Prefix Created: $Name" -ForegroundColor Green
            $Summary.Created++
        }
        catch {
            Write-Host "$Prefix Failed creating: $Name" -ForegroundColor Red
            Write-Host (Get-ApiErrorDetail $_)
            $Summary.Failed++
        }
    }
}

# =========================================================
# Comprehensive final document type list
# =========================================================

$DocumentTypeDefinitions = @(
    @{ name = "Account Statements";        matching_algorithm = $null; matching = $null }
    @{ name = "Bills";                     matching_algorithm = $null; matching = $null }
    @{ name = "Contracts";                 matching_algorithm = $null; matching = $null }
    @{ name = "Correspondence";            matching_algorithm = $null; matching = $null }
    @{ name = "Forms";                     matching_algorithm = $null; matching = $null }
    @{ name = "Government Documents";      matching_algorithm = $null; matching = $null }
    @{ name = "Identity Documents";        matching_algorithm = $null; matching = $null }
    @{ name = "Insurance Documents";       matching_algorithm = $null; matching = $null }
    @{ name = "Invoices";                  matching_algorithm = $null; matching = $null }
    @{ name = "Legal";                     matching_algorithm = $null; matching = $null }
    @{ name = "Manuals / Instructions";    matching_algorithm = $null; matching = $null }
    @{ name = "Medical Records";           matching_algorithm = $null; matching = $null }
    @{ name = "Orders / Purchase Records"; matching_algorithm = $null; matching = $null }
    @{ name = "Receipts";                  matching_algorithm = $null; matching = $null }
    @{ name = "Reports";                   matching_algorithm = $null; matching = $null }
    @{ name = "Tax Documents";             matching_algorithm = $null; matching = $null }
    @{ name = "Technical Documentation";   matching_algorithm = $null; matching = $null }
    @{ name = "Photos / Scans";            matching_algorithm = $null; matching = $null }
    @{ name = "Other";                     matching_algorithm = $null; matching = $null }
    @{ name = "Delivery Note";             matching_algorithm = $null; matching = $null }
)

Write-Host "Reading existing document types..." -ForegroundColor White
$ExistingTypes = Get-ExistingDocumentTypesByName

$total = $DocumentTypeDefinitions.Count
for ($i = 0; $i -lt $total; $i++) {
    New-OrUpdate-DocumentType -Definition $DocumentTypeDefinitions[$i] -ExistingTypes $ExistingTypes -Index ($i + 1) -Total $total
}

Write-Host ""
Write-Host "Summary" -ForegroundColor White
Write-Host "Created: $($Summary.Created)"
Write-Host "Updated: $($Summary.Updated)"
Write-Host "Skipped: $($Summary.Skipped)"
Write-Host "Failed:  $($Summary.Failed)"

if ($DryRun) {
    Write-Host "DRY RUN only - no changes were made." -ForegroundColor Yellow
}
else {
    Write-Host "Done." -ForegroundColor Green
}