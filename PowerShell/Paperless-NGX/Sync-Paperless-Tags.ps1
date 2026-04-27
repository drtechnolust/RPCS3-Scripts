# ============================================
# Paperless-ngx Tag Sync Script (Updated)
# - Creates missing tags
# - Updates existing tags only when needed
# - Sets colors
# - Sets parent/child relationships
# - Skips unchanged PATCH calls
# - Skips parent PATCH if already correct
# - Shows raw API error body
# - Includes receipt/mobile manual tagging tags
# ============================================

$PaperlessUrl = "http://10.10.1.99:8000"
$Token = ""

# Set to $true to preview only
$DryRun = $false

$Headers = @{
    Authorization = "Token $Token"
    "Content-Type" = "application/json"
}

$Summary = [ordered]@{
    Created = 0
    Updated = 0
    Skipped = 0
    Failed = 0
    ParentsSet = 0
    ParentsAlreadyCorrect = 0
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
        if ($Response.results) {
            $Results += $Response.results
            $NextUrl = $Response.next
        } else {
            return $Response
        }
    }

    return $Results
}

function Get-ExistingTagsByName {
    $AllTags = Invoke-PaperlessGetAll -Endpoint "/api/tags/?page_size=1000"
    $Lookup = @{}
    foreach ($Tag in $AllTags) {
        $Lookup[$Tag.name] = $Tag
    }
    return $Lookup
}

function Normalize-String {
    param([object]$Value)
    if ($null -eq $Value) { return "" }
    return ([string]$Value).Trim()
}

function New-OrUpdate-Tag {
    param(
        [hashtable]$TagDefinition,
        [hashtable]$ExistingTags,
        [int]$Index,
        [int]$Total
    )

    $Name = $TagDefinition.name
    $Body = @{
        name = $TagDefinition.name
        color = $TagDefinition.color
        is_inbox_tag = [bool]$TagDefinition.is_inbox_tag
    }

    $Prefix = "[$Index/$Total]"

    if ($ExistingTags.ContainsKey($Name)) {
        $Existing = $ExistingTags[$Name]

        $NeedsPatch = $false
        if ((Normalize-String $Existing.color) -ne (Normalize-String $Body.color)) { $NeedsPatch = $true }
        if ([bool]$Existing.is_inbox_tag -ne [bool]$Body.is_inbox_tag) { $NeedsPatch = $true }

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
            Invoke-RestMethod -Uri "$PaperlessUrl/api/tags/$($Existing.id)/" -Method Patch -Headers $Headers -Body $JsonBody -ErrorAction Stop | Out-Null
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
            Invoke-RestMethod -Uri "$PaperlessUrl/api/tags/" -Method Post -Headers $Headers -Body $JsonBody -ErrorAction Stop | Out-Null
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

function Set-TagParent {
    param(
        [string]$ChildName,
        [string]$ParentName,
        [hashtable]$ExistingTags
    )

    if (-not $ExistingTags.ContainsKey($ChildName)) {
        Write-Host "Parent skip, child missing: $ChildName" -ForegroundColor Yellow
        return
    }
    if (-not $ExistingTags.ContainsKey($ParentName)) {
        Write-Host "Parent skip, parent missing: $ParentName" -ForegroundColor Yellow
        return
    }

    $Child = $ExistingTags[$ChildName]
    $Parent = $ExistingTags[$ParentName]
    $CurrentParentId = $null

    if ($null -ne $Child.parent) {
        if ($Child.parent -is [int]) {
            $CurrentParentId = $Child.parent
        } elseif ($Child.parent.PSObject.Properties.Name -contains "id") {
            $CurrentParentId = $Child.parent.id
        }
    }

    if ($CurrentParentId -eq $Parent.id) {
        Write-Host "Parent already correct: $ChildName -> $ParentName" -ForegroundColor DarkGray
        $Summary.ParentsAlreadyCorrect++
        return
    }

    if ($DryRun) {
        Write-Host "Would set parent: $ChildName -> $ParentName" -ForegroundColor Magenta
        return
    }

    try {
        $Body = @{ parent = $Parent.id } | ConvertTo-Json
        Invoke-RestMethod -Uri "$PaperlessUrl/api/tags/$($Child.id)/" -Method Patch -Headers $Headers -Body $Body -ErrorAction Stop | Out-Null
        Write-Host "Set parent: $ChildName -> $ParentName" -ForegroundColor Magenta
        $Summary.ParentsSet++
    }
    catch {
        Write-Host "Failed parent set: $ChildName -> $ParentName" -ForegroundColor Red
        Write-Host (Get-ApiErrorDetail $_)
        $Summary.Failed++
    }
}

$Colors = @{
    Lifecycle  = "#f39c12"
    Personal   = "#27ae60"
    Finance    = "#3498db"
    Medical    = "#e74c3c"
    Insurance  = "#6c8ebf"
    Legal      = "#8e44ad"
    Work       = "#16a085"
    Tech       = "#7f8c8d"
    Homelab    = "#58d68d"
    Purchase   = "#f1c40f"
    Archive    = "#8d6e63"
}

$TagDefinitions = @(
    @{ name="Unsorted";              color=$Colors.Lifecycle; is_inbox_tag=$true }
    @{ name="Imported Legacy";       color=$Colors.Lifecycle; is_inbox_tag=$false }
    @{ name="Review Needed";         color=$Colors.Lifecycle; is_inbox_tag=$false }
    @{ name="Archive";               color=$Colors.Archive;   is_inbox_tag=$false }
    @{ name="Duplicate Suspected";   color=$Colors.Lifecycle; is_inbox_tag=$false }
    @{ name="Junk Review";           color=$Colors.Lifecycle; is_inbox_tag=$false }
    @{ name="Needs OCR Review";      color=$Colors.Lifecycle; is_inbox_tag=$false }
    @{ name="Needs Metadata Review"; color=$Colors.Lifecycle; is_inbox_tag=$false }
    @{ name="Low Confidence";        color=$Colors.Lifecycle; is_inbox_tag=$false }

    @{ name="Personal";      color=$Colors.Personal;  is_inbox_tag=$false }
    @{ name="Home";          color=$Colors.Personal;  is_inbox_tag=$false }
    @{ name="Finance";       color=$Colors.Finance;   is_inbox_tag=$false }
    @{ name="Legal";         color=$Colors.Legal;     is_inbox_tag=$false }
    @{ name="Medical";       color=$Colors.Medical;   is_inbox_tag=$false }
    @{ name="Insurance";     color=$Colors.Insurance; is_inbox_tag=$false }
    @{ name="Government";    color=$Colors.Legal;     is_inbox_tag=$false }
    @{ name="Identity";      color=$Colors.Legal;     is_inbox_tag=$false }
    @{ name="Estate";        color=$Colors.Legal;     is_inbox_tag=$false }
    @{ name="Banking";       color=$Colors.Finance;   is_inbox_tag=$false }
    @{ name="Tax";           color=$Colors.Finance;   is_inbox_tag=$false }
    @{ name="Utilities";     color=$Colors.Personal;  is_inbox_tag=$false }
    @{ name="Travel";        color=$Colors.Personal;  is_inbox_tag=$false }
    @{ name="Education";     color=$Colors.Personal;  is_inbox_tag=$false }

    @{ name="Receipt";           color=$Colors.Finance;   is_inbox_tag=$false }
    @{ name="Receipt Review";    color=$Colors.Lifecycle; is_inbox_tag=$false }
    @{ name="Reimbursable";      color=$Colors.Finance;   is_inbox_tag=$false }
    @{ name="Business Expense";  color=$Colors.Finance;   is_inbox_tag=$false }
    @{ name="Tax Deduction";     color=$Colors.Finance;   is_inbox_tag=$false }

    @{ name="Personal Meal";     color=$Colors.Personal;  is_inbox_tag=$false }
    @{ name="Restaurant";        color=$Colors.Personal;  is_inbox_tag=$false }
    @{ name="Groceries";         color=$Colors.Personal;  is_inbox_tag=$false }
    @{ name="Gas";               color=$Colors.Personal;  is_inbox_tag=$false }
    @{ name="Parking";           color=$Colors.Personal;  is_inbox_tag=$false }
    @{ name="Subscription";      color=$Colors.Personal;  is_inbox_tag=$false }
    @{ name="Warranty";          color=$Colors.Personal;  is_inbox_tag=$false }

    @{ name="Work";              color=$Colors.Work;      is_inbox_tag=$false }
    @{ name="Technology";        color=$Colors.Work;      is_inbox_tag=$false }
    @{ name="Homelab";           color=$Colors.Homelab;   is_inbox_tag=$false }

    @{ name="Work Meal";         color=$Colors.Work;      is_inbox_tag=$false }
    @{ name="Travel Expense";    color=$Colors.Work;      is_inbox_tag=$false }
    @{ name="Office Purchase";   color=$Colors.Work;      is_inbox_tag=$false }

    @{ name="Projects";       color=$Colors.Work;      is_inbox_tag=$false }
    @{ name="Vendors";        color=$Colors.Work;      is_inbox_tag=$false }
    @{ name="Change Control"; color=$Colors.Work;      is_inbox_tag=$false }
    @{ name="Support";        color=$Colors.Work;      is_inbox_tag=$false }
    @{ name="Operations";     color=$Colors.Work;      is_inbox_tag=$false }
    @{ name="Compliance";     color=$Colors.Work;      is_inbox_tag=$false }
    @{ name="Procurement";    color=$Colors.Work;      is_inbox_tag=$false }
    @{ name="Licensing";      color=$Colors.Legal;     is_inbox_tag=$false }
    @{ name="KB";             color=$Colors.Work;      is_inbox_tag=$false }
    @{ name="Architecture";   color=$Colors.Work;      is_inbox_tag=$false }
    @{ name="Security";       color=$Colors.Work;      is_inbox_tag=$false }
    @{ name="Networking";     color=$Colors.Work;      is_inbox_tag=$false }
    @{ name="Cloud";          color=$Colors.Work;      is_inbox_tag=$false }
    @{ name="Microsoft 365";  color=$Colors.Work;      is_inbox_tag=$false }
    @{ name="Azure";          color=$Colors.Work;      is_inbox_tag=$false }
    @{ name="Intune";         color=$Colors.Work;      is_inbox_tag=$false }
    @{ name="Exchange";       color=$Colors.Work;      is_inbox_tag=$false }
    @{ name="Teams";          color=$Colors.Work;      is_inbox_tag=$false }
    @{ name="Documentation";  color=$Colors.Work;      is_inbox_tag=$false }
    @{ name="Incident";       color=$Colors.Work;      is_inbox_tag=$false }
    @{ name="Migration";      color=$Colors.Work;      is_inbox_tag=$false }
    @{ name="Backup";         color=$Colors.Work;      is_inbox_tag=$false }
    @{ name="Monitoring";     color=$Colors.Work;      is_inbox_tag=$false }
    @{ name="Reporting";      color=$Colors.Work;      is_inbox_tag=$false }

    @{ name="Unraid";         color=$Colors.Homelab;   is_inbox_tag=$false }
    @{ name="Docker";         color=$Colors.Homelab;   is_inbox_tag=$false }
    @{ name="Home Assistant"; color=$Colors.Homelab;   is_inbox_tag=$false }
    @{ name="Nextcloud";      color=$Colors.Homelab;   is_inbox_tag=$false }
    @{ name="Authentik";      color=$Colors.Homelab;   is_inbox_tag=$false }
    @{ name="Uptime Kuma";    color=$Colors.Homelab;   is_inbox_tag=$false }
    @{ name="Tailscale";      color=$Colors.Homelab;   is_inbox_tag=$false }
    @{ name="Cloudflare";     color=$Colors.Homelab;   is_inbox_tag=$false }
    @{ name="DNS";            color=$Colors.Homelab;   is_inbox_tag=$false }
    @{ name="Firewall";       color=$Colors.Homelab;   is_inbox_tag=$false }
    @{ name="Virtualization"; color=$Colors.Homelab;   is_inbox_tag=$false }

    @{ name="Windows";        color=$Colors.Tech;      is_inbox_tag=$false }
    @{ name="Linux";          color=$Colors.Tech;      is_inbox_tag=$false }
    @{ name="Android";        color=$Colors.Tech;      is_inbox_tag=$false }
    @{ name="Hardware";       color=$Colors.Tech;      is_inbox_tag=$false }
    @{ name="Software";       color=$Colors.Tech;      is_inbox_tag=$false }
    @{ name="Configuration";  color=$Colors.Tech;      is_inbox_tag=$false }
    @{ name="Logs";           color=$Colors.Tech;      is_inbox_tag=$false }
    @{ name="Imaging";        color=$Colors.Work;      is_inbox_tag=$false }
    @{ name="Medical Imaging"; color=$Colors.Medical;  is_inbox_tag=$false }

    @{ name="Claims";               color=$Colors.Insurance; is_inbox_tag=$false }
    @{ name="Benefits";             color=$Colors.Insurance; is_inbox_tag=$false }
    @{ name="Policy";               color=$Colors.Insurance; is_inbox_tag=$false }
    @{ name="Coverage";             color=$Colors.Insurance; is_inbox_tag=$false }
    @{ name="Health Insurance";     color=$Colors.Insurance; is_inbox_tag=$false }
    @{ name="Auto Insurance";       color=$Colors.Insurance; is_inbox_tag=$false }
    @{ name="Home Insurance";       color=$Colors.Insurance; is_inbox_tag=$false }
    @{ name="Renters Insurance";    color=$Colors.Insurance; is_inbox_tag=$false }
    @{ name="Life Insurance";       color=$Colors.Insurance; is_inbox_tag=$false }
    @{ name="Disability Insurance"; color=$Colors.Insurance; is_inbox_tag=$false }
)

$ParentMap = @(
    @{ child="Banking"; parent="Finance" }
    @{ child="Tax"; parent="Finance" }
    @{ child="Receipt"; parent="Finance" }
    @{ child="Receipt Review"; parent="Receipt" }
    @{ child="Reimbursable"; parent="Finance" }
    @{ child="Business Expense"; parent="Finance" }
    @{ child="Tax Deduction"; parent="Finance" }

    @{ child="Utilities"; parent="Home" }

    @{ child="Insurance"; parent="Personal" }
    @{ child="Medical"; parent="Personal" }
    @{ child="Government"; parent="Personal" }
    @{ child="Identity"; parent="Personal" }
    @{ child="Estate"; parent="Personal" }
    @{ child="Travel"; parent="Personal" }
    @{ child="Education"; parent="Personal" }
    @{ child="Personal Meal"; parent="Personal" }
    @{ child="Restaurant"; parent="Personal" }
    @{ child="Groceries"; parent="Personal" }
    @{ child="Gas"; parent="Personal" }
    @{ child="Parking"; parent="Personal" }
    @{ child="Subscription"; parent="Personal" }
    @{ child="Warranty"; parent="Personal" }

    @{ child="Claims"; parent="Insurance" }
    @{ child="Benefits"; parent="Insurance" }
    @{ child="Policy"; parent="Insurance" }
    @{ child="Coverage"; parent="Insurance" }
    @{ child="Health Insurance"; parent="Insurance" }
    @{ child="Auto Insurance"; parent="Insurance" }
    @{ child="Home Insurance"; parent="Insurance" }
    @{ child="Renters Insurance"; parent="Insurance" }
    @{ child="Life Insurance"; parent="Insurance" }
    @{ child="Disability Insurance"; parent="Insurance" }
    @{ child="Medical Imaging"; parent="Medical" }

    @{ child="Technology"; parent="Work" }
    @{ child="Work Meal"; parent="Work" }
    @{ child="Travel Expense"; parent="Work" }
    @{ child="Office Purchase"; parent="Work" }
    @{ child="Projects"; parent="Work" }
    @{ child="Vendors"; parent="Work" }
    @{ child="Change Control"; parent="Work" }
    @{ child="Support"; parent="Work" }
    @{ child="Operations"; parent="Work" }
    @{ child="Compliance"; parent="Work" }
    @{ child="Procurement"; parent="Work" }
    @{ child="KB"; parent="Work" }
    @{ child="Architecture"; parent="Work" }
    @{ child="Security"; parent="Work" }
    @{ child="Networking"; parent="Work" }
    @{ child="Cloud"; parent="Technology" }
    @{ child="Microsoft 365"; parent="Technology" }
    @{ child="Azure"; parent="Cloud" }
    @{ child="Intune"; parent="Microsoft 365" }
    @{ child="Exchange"; parent="Microsoft 365" }
    @{ child="Teams"; parent="Microsoft 365" }
    @{ child="Documentation"; parent="Work" }
    @{ child="Incident"; parent="Work" }
    @{ child="Migration"; parent="Work" }
    @{ child="Backup"; parent="Work" }
    @{ child="Monitoring"; parent="Work" }
    @{ child="Reporting"; parent="Work" }
    @{ child="Imaging"; parent="Technology" }

    @{ child="Homelab"; parent="Personal" }
    @{ child="Unraid"; parent="Homelab" }
    @{ child="Docker"; parent="Homelab" }
    @{ child="Home Assistant"; parent="Homelab" }
    @{ child="Nextcloud"; parent="Homelab" }
    @{ child="Authentik"; parent="Homelab" }
    @{ child="Uptime Kuma"; parent="Homelab" }
    @{ child="Tailscale"; parent="Homelab" }
    @{ child="Cloudflare"; parent="Homelab" }
    @{ child="DNS"; parent="Homelab" }
    @{ child="Firewall"; parent="Homelab" }
    @{ child="Virtualization"; parent="Homelab" }
)

Write-Host "Reading existing tags..."
$ExistingTags = Get-ExistingTagsByName

$total = $TagDefinitions.Count
for ($i = 0; $i -lt $total; $i++) {
    New-OrUpdate-Tag -TagDefinition $TagDefinitions[$i] -ExistingTags $ExistingTags -Index ($i + 1) -Total $total
}

Write-Host "Refreshing tag cache..."
$ExistingTags = Get-ExistingTagsByName

Write-Host "Setting parents..."
foreach ($Map in $ParentMap) {
    Set-TagParent -ChildName $Map.child -ParentName $Map.parent -ExistingTags $ExistingTags
}

Write-Host ""
Write-Host "Summary"
Write-Host "Created: $($Summary.Created)"
Write-Host "Updated: $($Summary.Updated)"
Write-Host "Skipped: $($Summary.Skipped)"
Write-Host "Failed: $($Summary.Failed)"
Write-Host "Parents Set: $($Summary.ParentsSet)"
Write-Host "Parents Already Correct: $($Summary.ParentsAlreadyCorrect)"
if ($DryRun) {
    Write-Host "DRY RUN only - no changes were made." -ForegroundColor Yellow
} else {
    Write-Host "Done." -ForegroundColor Green
}