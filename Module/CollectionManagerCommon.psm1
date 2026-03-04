<#
.SYNOPSIS
    Core module for MECM Collection Manager with Offline WQL Editor.

.DESCRIPTION
    Import this module to get:
      - Structured logging (Initialize-Logging, Write-Log)
      - CM site connection management (Connect-CMSite, Disconnect-CMSite, Test-CMConnection)
      - Device collection queries and CRUD operations
      - Membership management (direct, query, include, exclude rules)
      - WQL query validation and preview
      - Template loading and parameter expansion
      - Export to CSV and HTML

.EXAMPLE
    Import-Module "$PSScriptRoot\Module\CollectionManagerCommon.psd1" -Force
    Initialize-Logging -LogPath "C:\temp\collmgr.log"
    Connect-CMSite -SiteCode 'MCM' -SMSProvider 'sccm01.contoso.com'
#>

# ---------------------------------------------------------------------------
# Module-scoped state
# ---------------------------------------------------------------------------

$script:__CMLogPath            = $null
$script:OriginalLocation       = $null
$script:ConnectedSiteCode      = $null
$script:ConnectedSMSProvider   = $null

# ---------------------------------------------------------------------------
# Logging
# ---------------------------------------------------------------------------

function Initialize-Logging {
    param([string]$LogPath)

    $script:__CMLogPath = $LogPath

    if ($LogPath) {
        $parentDir = Split-Path -Path $LogPath -Parent
        if ($parentDir -and -not (Test-Path -LiteralPath $parentDir)) {
            New-Item -ItemType Directory -Path $parentDir -Force | Out-Null
        }

        $header = "[{0}] [INFO ] === Log initialized ===" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
        Set-Content -LiteralPath $LogPath -Value $header -Encoding UTF8
    }
}

function Write-Log {
    <#
    .SYNOPSIS
        Writes a timestamped, severity-tagged log message.
    #>
    param(
        [AllowEmptyString()]
        [Parameter(Mandatory, Position = 0)]
        [string]$Message,

        [ValidateSet('INFO', 'WARN', 'ERROR')]
        [string]$Level = 'INFO',

        [switch]$Quiet
    )

    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $formatted = "[{0}] [{1,-5}] {2}" -f $timestamp, $Level, $Message

    if (-not $Quiet) {
        Write-Host $formatted
        if ($Level -eq 'ERROR') {
            $host.UI.WriteErrorLine($formatted)
        }
    }

    if ($script:__CMLogPath) {
        Add-Content -LiteralPath $script:__CMLogPath -Value $formatted -Encoding UTF8 -ErrorAction SilentlyContinue
    }
}

# ---------------------------------------------------------------------------
# CM Connection
# ---------------------------------------------------------------------------

function Connect-CMSite {
    <#
    .SYNOPSIS
        Imports the ConfigurationManager module, creates a PSDrive, and sets location.
    .DESCRIPTION
        Returns $true on success, $false on failure.
    #>
    param(
        [Parameter(Mandatory)][string]$SiteCode,
        [Parameter(Mandatory)][string]$SMSProvider
    )

    $script:OriginalLocation = Get-Location

    if (-not (Get-Module ConfigurationManager -ErrorAction SilentlyContinue)) {
        $cmModulePath = $null
        if ($env:SMS_ADMIN_UI_PATH) {
            $cmModulePath = Join-Path $env:SMS_ADMIN_UI_PATH '..\ConfigurationManager.psd1'
        }

        if (-not $cmModulePath -or -not (Test-Path -LiteralPath $cmModulePath)) {
            Write-Log "ConfigurationManager module not found. Ensure the CM console is installed." -Level ERROR
            return $false
        }

        try {
            Import-Module $cmModulePath -ErrorAction Stop
            Write-Log "Imported ConfigurationManager module"
        }
        catch {
            Write-Log "Failed to import ConfigurationManager module: $_" -Level ERROR
            return $false
        }
    }

    if (-not (Get-PSDrive -Name $SiteCode -PSProvider CMSite -ErrorAction SilentlyContinue)) {
        try {
            New-PSDrive -Name $SiteCode -PSProvider CMSite -Root $SMSProvider -ErrorAction Stop | Out-Null
            Write-Log "Created PSDrive for site $SiteCode"
        }
        catch {
            Write-Log "Failed to create PSDrive for site $SiteCode : $_" -Level ERROR
            return $false
        }
    }

    try {
        Set-Location "${SiteCode}:" -ErrorAction Stop
        $site = Get-CMSite -SiteCode $SiteCode -ErrorAction Stop
        Write-Log "Connected to site $SiteCode ($($site.SiteName))"
        $script:ConnectedSiteCode    = $SiteCode
        $script:ConnectedSMSProvider = $SMSProvider
        return $true
    }
    catch {
        Write-Log "Failed to connect to site $SiteCode : $_" -Level ERROR
        return $false
    }
}

function Disconnect-CMSite {
    <#
    .SYNOPSIS
        Restores the original location before CM connection.
    #>
    if ($script:OriginalLocation) {
        try { Set-Location $script:OriginalLocation -ErrorAction SilentlyContinue } catch { }
    }
    $script:ConnectedSiteCode    = $null
    $script:ConnectedSMSProvider = $null
    Write-Log "Disconnected from CM site"
}

function Test-CMConnection {
    <#
    .SYNOPSIS
        Returns $true if currently connected to a CM site.
    #>
    if (-not $script:ConnectedSiteCode) { return $false }

    try {
        $drive = Get-PSDrive -Name $script:ConnectedSiteCode -PSProvider CMSite -ErrorAction Stop
        return ($null -ne $drive)
    }
    catch {
        return $false
    }
}

# ---------------------------------------------------------------------------
# Collection Queries
# ---------------------------------------------------------------------------

function Get-AllDeviceCollections {
    <#
    .SYNOPSIS
        Returns all device collections with summary properties.
    #>
    Write-Log "Querying all device collections..."

    $collections = Get-CMDeviceCollection -ErrorAction Stop

    $results = foreach ($c in $collections) {
        [PSCustomObject]@{
            CollectionID       = $c.CollectionID
            Name               = $c.Name
            MemberCount        = [int]$c.MemberCount
            LimitToCollectionName = $c.LimitToCollectionName
            LimitToCollectionID   = $c.LimitToCollectionID
            Comment            = $c.Comment
            RefreshType        = switch ([int]$c.RefreshType) {
                1 { 'Manual' }
                2 { 'Periodic' }
                4 { 'Continuous' }
                6 { 'Both' }
                default { "Unknown ($($c.RefreshType))" }
            }
            CollectionRules    = $c.CollectionRules
            IsBuiltIn          = $c.CollectionID -like 'SMS*'
        }
    }

    Write-Log "Found $(@($results).Count) device collections"
    return $results
}

function Get-CollectionDetail {
    <#
    .SYNOPSIS
        Returns detailed info for a single collection including rule counts.
    #>
    param(
        [Parameter(Mandatory)][string]$CollectionId
    )

    $c = Get-CMDeviceCollection -CollectionId $CollectionId -ErrorAction Stop
    if (-not $c) { return $null }

    $queryRules = @(Get-CMDeviceCollectionQueryMembershipRule -CollectionId $CollectionId -ErrorAction SilentlyContinue)

    return [PSCustomObject]@{
        CollectionID          = $c.CollectionID
        Name                  = $c.Name
        MemberCount           = [int]$c.MemberCount
        LimitToCollectionName = $c.LimitToCollectionName
        Comment               = $c.Comment
        RefreshType           = [int]$c.RefreshType
        QueryRuleCount        = $queryRules.Count
        QueryRules            = $queryRules
        IsBuiltIn             = $c.CollectionID -like 'SMS*'
    }
}

function Get-CollectionQueryRules {
    <#
    .SYNOPSIS
        Returns all query membership rules for a collection.
    #>
    param(
        [Parameter(Mandatory)][string]$CollectionId
    )

    Write-Log "Getting query rules for collection $CollectionId..."

    $rules = Get-CMDeviceCollectionQueryMembershipRule -CollectionId $CollectionId -ErrorAction Stop

    $results = foreach ($r in $rules) {
        [PSCustomObject]@{
            RuleName        = $r.RuleName
            QueryExpression = $r.QueryExpression
        }
    }

    Write-Log "Found $(@($results).Count) query rules"
    return $results
}

function Get-CollectionMembers {
    <#
    .SYNOPSIS
        Returns current members of a collection.
    #>
    param(
        [Parameter(Mandatory)][string]$CollectionId
    )

    Write-Log "Getting members of collection $CollectionId..."

    $members = Get-CMCollectionMember -CollectionId $CollectionId -ErrorAction Stop

    $results = foreach ($m in $members) {
        [PSCustomObject]@{
            Name       = $m.Name
            ResourceID = $m.ResourceID
            Domain     = $m.Domain
            IsClient   = $m.IsClient
            IsActive   = $m.IsActive
        }
    }

    Write-Log "Found $(@($results).Count) members"
    return $results
}

# ---------------------------------------------------------------------------
# Collection CRUD
# ---------------------------------------------------------------------------

function New-ManagedCollection {
    <#
    .SYNOPSIS
        Creates a new device collection.
    #>
    param(
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][string]$LimitingCollectionName,
        [string]$Comment = '',
        [ValidateSet('Manual', 'Periodic', 'Continuous', 'Both')]
        [string]$RefreshType = 'Both'
    )

    $rtMap = @{ 'Manual' = 1; 'Periodic' = 2; 'Continuous' = 4; 'Both' = 6 }
    $rtValue = $rtMap[$RefreshType]

    Write-Log "Creating collection '$Name' (limiting: '$LimitingCollectionName', refresh: $RefreshType)..."

    $params = @{
        Name                   = $Name
        LimitingCollectionName = $LimitingCollectionName
        RefreshType            = $rtValue
        ErrorAction            = 'Stop'
    }
    if ($Comment) { $params['Comment'] = $Comment }

    # For Periodic or Both, add a 7-day refresh schedule
    if ($rtValue -in 2, 6) {
        $schedule = New-CMSchedule -RecurInterval Days -RecurCount 7
        $params['RefreshSchedule'] = $schedule
    }

    $collection = New-CMDeviceCollection @params
    Write-Log "Created collection '$Name' (ID: $($collection.CollectionID))"
    return $collection
}

function Copy-ManagedCollection {
    <#
    .SYNOPSIS
        Clones an existing collection.
    #>
    param(
        [Parameter(Mandatory)][string]$SourceCollectionId,
        [Parameter(Mandatory)][string]$NewName
    )

    Write-Log "Cloning collection $SourceCollectionId as '$NewName'..."

    $clone = Copy-CMCollection -Id $SourceCollectionId -NewName $NewName -PassThru -ErrorAction Stop
    Write-Log "Cloned to '$NewName' (ID: $($clone.CollectionID))"
    return $clone
}

function Remove-ManagedCollection {
    <#
    .SYNOPSIS
        Deletes a device collection with safety check.
    .DESCRIPTION
        Blocks deletion of built-in collections (IDs starting with SMS).
    #>
    param(
        [Parameter(Mandatory)][string]$CollectionId
    )

    if ($CollectionId -like 'SMS*') {
        Write-Log "BLOCKED: Cannot delete built-in collection $CollectionId" -Level WARN
        return $false
    }

    Write-Log "Removing collection $CollectionId..."

    Remove-CMDeviceCollection -CollectionId $CollectionId -Force -ErrorAction Stop
    Write-Log "Removed collection $CollectionId"
    return $true
}

function Set-CollectionProperties {
    <#
    .SYNOPSIS
        Updates collection name, comment, or refresh type.
    #>
    param(
        [Parameter(Mandatory)][string]$CollectionId,
        [string]$NewName,
        [string]$Comment,
        [int]$RefreshType
    )

    Write-Log "Updating properties for collection $CollectionId..."

    $params = @{
        CollectionId = $CollectionId
        ErrorAction  = 'Stop'
    }
    if ($NewName)     { $params['NewName'] = $NewName }
    if ($PSBoundParameters.ContainsKey('Comment')) { $params['Comment'] = $Comment }

    Set-CMCollection @params
    Write-Log "Updated collection $CollectionId"
}

# ---------------------------------------------------------------------------
# Membership Management
# ---------------------------------------------------------------------------

function Add-DirectMember {
    <#
    .SYNOPSIS
        Adds a device to a collection by name (resolves ResourceId automatically).
    #>
    param(
        [Parameter(Mandatory)][string]$CollectionId,
        [Parameter(Mandatory)][string]$DeviceName
    )

    Write-Log "Resolving ResourceId for '$DeviceName'..."

    $device = Get-CMDevice -Name $DeviceName -ErrorAction SilentlyContinue
    if (-not $device) {
        Write-Log "Device '$DeviceName' not found in MECM" -Level ERROR
        return $false
    }

    $resourceId = $device.ResourceID
    Write-Log "Adding '$DeviceName' (ResourceId: $resourceId) to collection $CollectionId..."

    Add-CMDeviceCollectionDirectMembershipRule -CollectionId $CollectionId -ResourceId $resourceId -ErrorAction Stop
    Write-Log "Added '$DeviceName' to collection $CollectionId"
    return $true
}

function Remove-DirectMember {
    <#
    .SYNOPSIS
        Removes a direct member from a collection.
    #>
    param(
        [Parameter(Mandatory)][string]$CollectionId,
        [Parameter(Mandatory)][int]$ResourceId
    )

    Write-Log "Removing ResourceId $ResourceId from collection $CollectionId..."

    Remove-CMDeviceCollectionDirectMembershipRule -CollectionId $CollectionId -ResourceId $ResourceId -Force -ErrorAction Stop
    Write-Log "Removed ResourceId $ResourceId from collection $CollectionId"
}

function Add-QueryRule {
    <#
    .SYNOPSIS
        Adds a WQL query membership rule to a collection.
    #>
    param(
        [Parameter(Mandatory)][string]$CollectionId,
        [Parameter(Mandatory)][string]$RuleName,
        [Parameter(Mandatory)][string]$QueryExpression
    )

    Write-Log "Adding query rule '$RuleName' to collection $CollectionId..."

    Add-CMDeviceCollectionQueryMembershipRule -CollectionId $CollectionId -RuleName $RuleName -QueryExpression $QueryExpression -ErrorAction Stop
    Write-Log "Added query rule '$RuleName'"
}

function Remove-QueryRule {
    <#
    .SYNOPSIS
        Removes a query membership rule from a collection.
    #>
    param(
        [Parameter(Mandatory)][string]$CollectionId,
        [Parameter(Mandatory)][string]$RuleName
    )

    Write-Log "Removing query rule '$RuleName' from collection $CollectionId..."

    Remove-CMDeviceCollectionQueryMembershipRule -CollectionId $CollectionId -RuleName $RuleName -Force -ErrorAction Stop
    Write-Log "Removed query rule '$RuleName'"
}

function Update-QueryRule {
    <#
    .SYNOPSIS
        Updates a query rule by removing the old one and adding a new one.
    .DESCRIPTION
        There is no native update cmdlet for query rules. This removes the existing
        rule by name and adds a replacement with the same or new name.
    #>
    param(
        [Parameter(Mandatory)][string]$CollectionId,
        [Parameter(Mandatory)][string]$OldRuleName,
        [Parameter(Mandatory)][string]$NewRuleName,
        [Parameter(Mandatory)][string]$NewQueryExpression
    )

    Write-Log "Updating query rule '$OldRuleName' on collection $CollectionId..."

    Remove-CMDeviceCollectionQueryMembershipRule -CollectionId $CollectionId -RuleName $OldRuleName -Force -ErrorAction Stop
    Add-CMDeviceCollectionQueryMembershipRule -CollectionId $CollectionId -RuleName $NewRuleName -QueryExpression $NewQueryExpression -ErrorAction Stop

    Write-Log "Updated query rule '$OldRuleName' -> '$NewRuleName'"
}

function Add-IncludeRule {
    <#
    .SYNOPSIS
        Adds an include collection membership rule.
    #>
    param(
        [Parameter(Mandatory)][string]$CollectionId,
        [Parameter(Mandatory)][string]$IncludeCollectionId
    )

    Write-Log "Adding include rule for $IncludeCollectionId to collection $CollectionId..."

    Add-CMDeviceCollectionIncludeMembershipRule -CollectionId $CollectionId -IncludeCollectionId $IncludeCollectionId -ErrorAction Stop
    Write-Log "Added include rule"
}

function Add-ExcludeRule {
    <#
    .SYNOPSIS
        Adds an exclude collection membership rule.
    #>
    param(
        [Parameter(Mandatory)][string]$CollectionId,
        [Parameter(Mandatory)][string]$ExcludeCollectionId
    )

    Write-Log "Adding exclude rule for $ExcludeCollectionId to collection $CollectionId..."

    Add-CMDeviceCollectionExcludeMembershipRule -CollectionId $CollectionId -ExcludeCollectionId $ExcludeCollectionId -ErrorAction Stop
    Write-Log "Added exclude rule"
}

function Invoke-CollectionEvaluation {
    <#
    .SYNOPSIS
        Forces immediate membership evaluation for a collection.
    #>
    param(
        [Parameter(Mandatory)][string]$CollectionId
    )

    Write-Log "Forcing evaluation for collection $CollectionId..."

    Invoke-CMCollectionUpdate -CollectionId $CollectionId -ErrorAction Stop
    Write-Log "Evaluation triggered for $CollectionId"
}

# ---------------------------------------------------------------------------
# WQL Validation & Preview
# ---------------------------------------------------------------------------

function Test-WqlQuery {
    <#
    .SYNOPSIS
        Validates a WQL query expression via Invoke-CMWmiQuery.
    .DESCRIPTION
        Returns a PSCustomObject with IsValid, ResultCount, and ErrorMessage.
    #>
    param(
        [Parameter(Mandatory)][string]$QueryExpression
    )

    Write-Log "Validating WQL query..."

    try {
        $results = @(Invoke-CMWmiQuery -Query $QueryExpression -Option Lazy -ErrorAction Stop)
        Write-Log "Query valid, $($results.Count) results"
        return [PSCustomObject]@{
            IsValid      = $true
            ResultCount  = $results.Count
            ErrorMessage = ''
        }
    }
    catch {
        Write-Log "Query validation failed: $_" -Level WARN
        return [PSCustomObject]@{
            IsValid      = $false
            ResultCount  = 0
            ErrorMessage = $_.Exception.Message
        }
    }
}

function Invoke-WqlPreview {
    <#
    .SYNOPSIS
        Runs a WQL query and returns the first N results for preview.
    #>
    param(
        [Parameter(Mandatory)][string]$QueryExpression,
        [int]$MaxResults = 50
    )

    Write-Log "Running WQL preview (max $MaxResults results)..."

    try {
        $results = @(Invoke-CMWmiQuery -Query $QueryExpression -ErrorAction Stop)

        $preview = if ($results.Count -gt $MaxResults) {
            $results | Select-Object -First $MaxResults
        } else {
            $results
        }

        $previewObjects = foreach ($r in $preview) {
            [PSCustomObject]@{
                Name       = $r.Name
                ResourceID = $r.ResourceID
                Domain     = $r.ResourceDomainORWorkgroup
                Client     = $r.Client
            }
        }

        Write-Log "Preview returned $($previewObjects.Count) of $($results.Count) total results"
        return [PSCustomObject]@{
            TotalCount = $results.Count
            Results    = $previewObjects
        }
    }
    catch {
        Write-Log "WQL preview failed: $_" -Level ERROR
        return [PSCustomObject]@{
            TotalCount = 0
            Results    = @()
        }
    }
}

# ---------------------------------------------------------------------------
# Templates
# ---------------------------------------------------------------------------

function Get-OperationalTemplates {
    <#
    .SYNOPSIS
        Loads the operational collection templates from JSON.
    #>
    param(
        [string]$TemplatesPath
    )

    if (-not $TemplatesPath) {
        $TemplatesPath = Join-Path (Split-Path $PSScriptRoot -Parent) 'Templates\operational-collections.json'
    }

    if (-not (Test-Path -LiteralPath $TemplatesPath)) {
        Write-Log "Operational templates not found: $TemplatesPath" -Level WARN
        return @()
    }

    try {
        $templates = Get-Content -LiteralPath $TemplatesPath -Raw | ConvertFrom-Json
        Write-Log "Loaded $(@($templates).Count) operational templates"
        return $templates
    }
    catch {
        Write-Log "Failed to load operational templates: $_" -Level ERROR
        return @()
    }
}

function Get-ParameterizedTemplates {
    <#
    .SYNOPSIS
        Loads the parameterized WQL templates from JSON.
    #>
    param(
        [string]$TemplatesPath
    )

    if (-not $TemplatesPath) {
        $TemplatesPath = Join-Path (Split-Path $PSScriptRoot -Parent) 'Templates\parameterized-templates.json'
    }

    if (-not (Test-Path -LiteralPath $TemplatesPath)) {
        Write-Log "Parameterized templates not found: $TemplatesPath" -Level WARN
        return @()
    }

    try {
        $templates = Get-Content -LiteralPath $TemplatesPath -Raw | ConvertFrom-Json
        Write-Log "Loaded $(@($templates).Count) parameterized templates"
        return $templates
    }
    catch {
        Write-Log "Failed to load parameterized templates: $_" -Level ERROR
        return @()
    }
}

function Expand-TemplateParameters {
    <#
    .SYNOPSIS
        Substitutes placeholder values in a template query string.
    .DESCRIPTION
        Takes a query string with {Placeholder} tokens and a hashtable of
        Placeholder -> Value mappings. Returns the expanded query.
    #>
    param(
        [Parameter(Mandatory)][string]$QueryTemplate,
        [Parameter(Mandatory)][hashtable]$ParameterValues
    )

    $expanded = $QueryTemplate
    foreach ($key in $ParameterValues.Keys) {
        $expanded = $expanded.Replace($key, $ParameterValues[$key])
    }

    return $expanded
}

# ---------------------------------------------------------------------------
# Export
# ---------------------------------------------------------------------------

function Export-CollectionCsv {
    <#
    .SYNOPSIS
        Exports a DataTable to CSV.
    #>
    param(
        [Parameter(Mandatory)][System.Data.DataTable]$DataTable,
        [Parameter(Mandatory)][string]$OutputPath
    )

    $parentDir = Split-Path -Path $OutputPath -Parent
    if ($parentDir -and -not (Test-Path -LiteralPath $parentDir)) {
        New-Item -ItemType Directory -Path $parentDir -Force | Out-Null
    }

    $rows = @()
    foreach ($row in $DataTable.Rows) {
        $obj = [ordered]@{}
        foreach ($col in $DataTable.Columns) {
            $obj[$col.ColumnName] = $row[$col.ColumnName]
        }
        $rows += [PSCustomObject]$obj
    }

    $rows | Export-Csv -LiteralPath $OutputPath -NoTypeInformation -Encoding UTF8
    Write-Log "Exported CSV to $OutputPath"
}

function Export-CollectionHtml {
    <#
    .SYNOPSIS
        Exports a DataTable to a self-contained HTML report.
    #>
    param(
        [Parameter(Mandatory)][System.Data.DataTable]$DataTable,
        [Parameter(Mandatory)][string]$OutputPath,
        [string]$ReportTitle = 'Collection Manager Report'
    )

    $parentDir = Split-Path -Path $OutputPath -Parent
    if ($parentDir -and -not (Test-Path -LiteralPath $parentDir)) {
        New-Item -ItemType Directory -Path $parentDir -Force | Out-Null
    }

    $css = @(
        '<style>',
        'body { font-family: "Segoe UI", Arial, sans-serif; margin: 20px; background: #fafafa; }',
        'h1 { color: #0078D4; margin-bottom: 4px; }',
        '.summary { color: #666; margin-bottom: 12px; font-size: 0.9em; }',
        'table { border-collapse: collapse; width: 100%; margin-top: 12px; }',
        'th { background: #0078D4; color: #fff; padding: 8px 12px; text-align: left; }',
        'td { padding: 6px 12px; border-bottom: 1px solid #e0e0e0; }',
        'tr:nth-child(even) { background: #f5f5f5; }',
        '</style>'
    ) -join "`r`n"

    $headerRow = ($DataTable.Columns | ForEach-Object { "<th>$($_.ColumnName)</th>" }) -join ''
    $bodyRows = foreach ($row in $DataTable.Rows) {
        $cells = foreach ($col in $DataTable.Columns) {
            $val = [string]$row[$col.ColumnName]
            "<td>$val</td>"
        }
        "<tr>$($cells -join '')</tr>"
    }

    $html = @(
        '<!DOCTYPE html>',
        '<html><head><meta charset="utf-8"><title>' + $ReportTitle + '</title>',
        $css,
        '</head><body>',
        "<h1>$ReportTitle</h1>",
        "<div class='summary'>Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') | Rows: $($DataTable.Rows.Count)</div>",
        "<table><thead><tr>$headerRow</tr></thead>",
        "<tbody>$($bodyRows -join "`r`n")</tbody></table>",
        '</body></html>'
    ) -join "`r`n"

    Set-Content -LiteralPath $OutputPath -Value $html -Encoding UTF8
    Write-Log "Exported HTML to $OutputPath"
}
