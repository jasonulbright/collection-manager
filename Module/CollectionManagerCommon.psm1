<#
.SYNOPSIS
    Core module for Collection Manager with Offline WQL Editor.

.DESCRIPTION
    Import this module to get:
      - Structured logging and CM site connection management via the
        vendored SuiteCommon module (Lib\SuiteCommon)
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
# Shared core (vendored SuiteCommon)
# ---------------------------------------------------------------------------
# Logging (Initialize-Logging, Write-Log), CM connection (Connect-CMSite,
# Disconnect-CMSite, Test-CMConnection, Get-CMConnectionInfo), and settings
# persistence come from the vendored copy at Lib\SuiteCommon\. -Global makes
# the functions resolvable from the shell script and from this module alike;
# the guard keeps a -Force reimport of this module (the background runspace
# does one) from resetting SuiteCommon state mid-session.
if (-not (Get-Module SuiteCommon)) {
    Import-Module (Join-Path $PSScriptRoot '..\Lib\SuiteCommon\SuiteCommon.psd1') -Global -DisableNameChecking
}

# ---------------------------------------------------------------------------
# Folder hierarchy (SMS_ObjectContainerNode + SMS_ObjectContainerItem)
# ---------------------------------------------------------------------------
# The CM console's left-tree folders for device collections live in
# SMS_ObjectContainerNode rows where ObjectType = 5000. Each collection's
# folder placement lives in SMS_ObjectContainerItem (InstanceKey =
# CollectionID, ObjectType = 5000, ContainerNodeID = folder). Collections
# absent from SMS_ObjectContainerItem live at the tree root.

function Get-CMCollectionFolderTree {
    <#
    .SYNOPSIS
        Returns the device-collection folder hierarchy as flat PSCustomObjects
        (FolderID, Name, ParentID). ParentID = 0 means the folder lives at
        the tree root.
    #>
    param(
        [Parameter(Mandatory)][string]$SMSProvider,
        [Parameter(Mandatory)][string]$SiteCode
    )

    $namespace = "root\sms\site_$SiteCode"
    Write-Log "Querying folder tree from $SMSProvider\$namespace..."

    $folders = @(
        Get-CimInstance -ComputerName $SMSProvider -Namespace $namespace `
            -ClassName SMS_ObjectContainerNode `
            -Filter 'ObjectType = 5000' -ErrorAction Stop |
        ForEach-Object {
            [PSCustomObject]@{
                FolderID = [int]$_.ContainerNodeID
                Name     = [string]$_.Name
                ParentID = [int]$_.ParentContainerNodeID
            }
        }
    )

    Write-Log "Found $(@($folders).Count) collection folders"
    return ,$folders
}

function Get-CMCollectionFolderMap {
    <#
    .SYNOPSIS
        Returns @{ CollectionID -> ContainerNodeID } for every device collection
        that has been moved into a folder. Collections absent from this map
        live at the tree root (FolderID = 0).
    #>
    param(
        [Parameter(Mandatory)][string]$SMSProvider,
        [Parameter(Mandatory)][string]$SiteCode
    )

    $namespace = "root\sms\site_$SiteCode"
    Write-Log "Querying folder assignments from $SMSProvider\$namespace..."

    $items = @(
        Get-CimInstance -ComputerName $SMSProvider -Namespace $namespace `
            -ClassName SMS_ObjectContainerItem `
            -Filter 'ObjectType = 5000' -ErrorAction Stop
    )

    $map = @{}
    foreach ($i in $items) {
        $map[[string]$i.InstanceKey] = [int]$i.ContainerNodeID
    }

    Write-Log "Mapped $(@($map.Keys).Count) collections to folders"
    return $map
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
    $directRuleCount = 0
    $collectionRules = @()
    try { $collectionRules = @($c.CollectionRules) } catch { $null = $_ }
    foreach ($rule in $collectionRules) {
        if ($null -eq $rule) { continue }
        $typeName = $null
        try { $typeName = [string]$rule.SmsProviderObjectPath } catch { $typeName = $null }
        if (-not $typeName) {
            try { $typeName = $rule.GetType().Name } catch { continue }
        }
        if ($typeName -like '*RuleDirect*') { $directRuleCount++ }
    }

    return [PSCustomObject]@{
        CollectionID          = $c.CollectionID
        Name                  = $c.Name
        MemberCount           = [int]$c.MemberCount
        LimitToCollectionName = $c.LimitToCollectionName
        Comment               = $c.Comment
        RefreshType           = [int]$c.RefreshType
        QueryRuleCount        = $queryRules.Count
        QueryRules            = $queryRules
        DirectRuleCount       = $directRuleCount
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

function ConvertTo-SiteSavedQueryInfo {
    param([Parameter(Mandatory)]$Query)

    return [PSCustomObject]@{
        QueryID              = [string]$Query.QueryID
        Name                 = [string]$Query.Name
        Expression           = [string]$Query.Expression
        Comment              = [string]$Query.Comments
        TargetClassName      = [string]$Query.TargetClassName
        LimitToCollectionID  = [string]$Query.LimitToCollectionID
    }
}

function Get-SiteSavedQueries {
    <#
    .SYNOPSIS
        Returns the reusable WQL queries saved in the Configuration Manager site.
    #>
    Write-Log 'Getting saved queries from the site...'
    $queries = @(Get-CMQuery -ErrorAction Stop)
    $results = foreach ($query in $queries) {
        ConvertTo-SiteSavedQueryInfo -Query $query
    }
    $results = @($results | Sort-Object Name, QueryID)
    Write-Log "Found $($results.Count) saved queries"
    return $results
}

function New-SiteSavedQuery {
    <#
    .SYNOPSIS
        Saves a new reusable WQL query in the Configuration Manager site.
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][string]$Expression,
        [string]$Comment = '',
        [string]$TargetClassName = '',
        [string]$LimitToCollectionId = ''
    )

    $params = @{
        Name       = $Name
        Expression = $Expression
        ErrorAction = 'Stop'
    }
    if (-not [string]::IsNullOrWhiteSpace($Comment)) { $params['Comment'] = $Comment }
    if (-not [string]::IsNullOrWhiteSpace($TargetClassName)) { $params['TargetClassName'] = $TargetClassName }
    if (-not [string]::IsNullOrWhiteSpace($LimitToCollectionId)) { $params['LimitToCollectionId'] = $LimitToCollectionId }

    if (-not $PSCmdlet.ShouldProcess($Name, 'Save site query')) { return }
    Write-Log "Saving site query '$Name'..."
    $query = New-CMQuery @params
    if (-not $query) {
        $query = Get-CMQuery -Name $Name -DisableWildcardHandling -ErrorAction Stop |
            Where-Object { [string]$_.Name -eq $Name } |
            Select-Object -First 1
    }
    if (-not $query) { throw "Configuration Manager did not return the saved query '$Name'." }
    return (ConvertTo-SiteSavedQueryInfo -Query $query)
}

function Set-SiteSavedQuery {
    <#
    .SYNOPSIS
        Updates a saved site query's name, WQL, and metadata.
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)][string]$QueryId,
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][string]$Expression,
        [string]$Comment = '',
        [string]$TargetClassName = '',
        [string]$LimitToCollectionId = ''
    )

    if (-not $PSCmdlet.ShouldProcess($QueryId, 'Update site query')) { return }
    Write-Log "Updating site query $QueryId..."
    $query = Set-CMQuery -Id $QueryId -NewName $Name -Expression $Expression `
        -Comment $Comment -TargetClassName $TargetClassName `
        -LimitToCollectionId $LimitToCollectionId -PassThru -ErrorAction Stop
    if (-not $query) { $query = Get-CMQuery -Id $QueryId -ErrorAction Stop }
    if (-not $query) { throw "Configuration Manager did not return the updated query '$QueryId'." }
    return (ConvertTo-SiteSavedQueryInfo -Query $query)
}

function Remove-SiteSavedQuery {
    <#
    .SYNOPSIS
        Deletes a reusable WQL query from the Configuration Manager site.
        Returns $true when it deleted the query, $false under -WhatIf.
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param([Parameter(Mandatory)][string]$QueryId)

    if (-not $PSCmdlet.ShouldProcess($QueryId, 'Remove site query')) { return $false }
    Write-Log "Removing site query $QueryId..."
    Remove-CMQuery -Id $QueryId -Force -ErrorAction Stop
    Write-Log "Removed site query $QueryId"
    return $true
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
# Every function below supports -WhatIf. Under -WhatIf it changes nothing and
# returns $null, or $false where the function returns a success flag.

function New-ManagedCollection {
    <#
    .SYNOPSIS
        Creates a new device collection.
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][string]$LimitingCollectionName,
        [string]$Comment = '',
        [ValidateSet('Manual', 'Periodic', 'Continuous', 'Both')]
        [string]$RefreshType = 'Both'
    )

    $rtMap = @{ 'Manual' = 1; 'Periodic' = 2; 'Continuous' = 4; 'Both' = 6 }
    $rtValue = $rtMap[$RefreshType]

    if (-not $PSCmdlet.ShouldProcess($Name, "Create device collection limited to '$LimitingCollectionName'")) { return }
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
        Clones an existing collection and returns the clone.
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)][string]$SourceCollectionId,
        [Parameter(Mandatory)][string]$NewName
    )

    if (-not $PSCmdlet.ShouldProcess($SourceCollectionId, "Clone collection as '$NewName'")) { return }
    Write-Log "Cloning collection $SourceCollectionId as '$NewName'..."

    [void](Copy-CMCollection -Id $SourceCollectionId -NewName $NewName -PassThru -ErrorAction Stop)
    # Copy-CMCollection -PassThru returns the source collection on the lab site (checked on
    # CM01), not the clone. Read the clone back by its exact name.
    $clone = @(Get-CMDeviceCollection -Name $NewName -ErrorAction Stop | Where-Object { [string]$_.Name -ceq $NewName }) | Select-Object -First 1
    if (-not $clone) { throw "Configuration Manager did not return the cloned collection '$NewName'." }
    Write-Log "Cloned to '$NewName' (ID: $($clone.CollectionID))"
    return $clone
}

function Remove-ManagedCollection {
    <#
    .SYNOPSIS
        Deletes a device collection with safety check.
    .DESCRIPTION
        Blocks deletion of built-in collections (IDs starting with SMS). The
        check ignores leading and trailing white space and letter case.
        Returns $true when it deleted the collection, $false when it blocked
        the deletion or ran under -WhatIf.
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)][string]$CollectionId
    )

    $CollectionId = $CollectionId.Trim()
    if ($CollectionId -like 'SMS*') {
        Write-Log "BLOCKED: Cannot delete built-in collection $CollectionId" -Level WARN
        return $false
    }

    if (-not $PSCmdlet.ShouldProcess($CollectionId, 'Remove device collection')) { return $false }
    Write-Log "Removing collection $CollectionId..."

    Remove-CMDeviceCollection -CollectionId $CollectionId -Force -ErrorAction Stop
    Write-Log "Removed collection $CollectionId"
    return $true
}

function Set-CollectionProperties {
    <#
    .SYNOPSIS
        Updates collection name, comment, or refresh type.
    .DESCRIPTION
        RefreshType is 1 (Manual), 2 (Periodic), 4 (Continuous) or 6 (Both).
        Any other value throws before the site is called. An empty Comment
        clears the comment; leaving Comment out keeps it.
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)][string]$CollectionId,
        [string]$NewName,
        [string]$Comment,
        [int]$RefreshType
    )

    $params = @{
        CollectionId = $CollectionId
        ErrorAction  = 'Stop'
    }
    if ($NewName)     { $params['NewName'] = $NewName }
    if ($PSBoundParameters.ContainsKey('Comment')) { $params['Comment'] = $Comment }
    if ($PSBoundParameters.ContainsKey('RefreshType')) {
        if ($RefreshType -notin 1, 2, 4, 6) { throw "Invalid refresh type $RefreshType. Use 1 (Manual), 2 (Periodic), 4 (Continuous) or 6 (Both)." }
        $params['RefreshType'] = $RefreshType
        # Periodic (2) and Both (6) need a refresh schedule, as in New-ManagedCollection.
        if ($RefreshType -in 2, 6) { $params['RefreshSchedule'] = New-CMSchedule -RecurInterval Days -RecurCount 7 }
    }

    if (-not $PSCmdlet.ShouldProcess($CollectionId, 'Update collection properties')) { return }
    Write-Log "Updating properties for collection $CollectionId..."
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
    .DESCRIPTION
        The name matches one device exactly (ignoring letter case). A wildcard
        such as * is plain text. Returns $false, and adds nothing, when no
        device or more than one device has that name, and under -WhatIf.
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)][string]$CollectionId,
        [Parameter(Mandatory)][string]$DeviceName
    )

    Write-Log "Resolving ResourceId for '$DeviceName'..."

    $devices = @(Get-CMDevice -Name $DeviceName -DisableWildcardHandling -ErrorAction SilentlyContinue |
            Where-Object { [string]$_.Name -ieq $DeviceName })
    if ($devices.Count -eq 0) {
        Write-Log "Device '$DeviceName' not found in Configuration Manager" -Level ERROR
        return $false
    }
    if ($devices.Count -gt 1) {
        Write-Log "More than one device is named '$DeviceName' ($($devices.Count) records); nothing was added" -Level ERROR
        return $false
    }

    $resourceId = $devices[0].ResourceID
    if (-not $PSCmdlet.ShouldProcess("$DeviceName (ResourceId $resourceId)", "Add to collection $CollectionId")) { return $false }
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
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)][string]$CollectionId,
        [Parameter(Mandatory)][int]$ResourceId
    )

    if (-not $PSCmdlet.ShouldProcess("ResourceId $ResourceId", "Remove from collection $CollectionId")) { return }
    Write-Log "Removing ResourceId $ResourceId from collection $CollectionId..."

    Remove-CMDeviceCollectionDirectMembershipRule -CollectionId $CollectionId -ResourceId $ResourceId -Force -ErrorAction Stop
    Write-Log "Removed ResourceId $ResourceId from collection $CollectionId"
}

function Add-QueryRule {
    <#
    .SYNOPSIS
        Adds a WQL query membership rule to a collection.
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)][string]$CollectionId,
        [Parameter(Mandatory)][string]$RuleName,
        [Parameter(Mandatory)][string]$QueryExpression
    )

    if (-not $PSCmdlet.ShouldProcess($CollectionId, "Add query rule '$RuleName'")) { return }
    Write-Log "Adding query rule '$RuleName' to collection $CollectionId..."

    Add-CMDeviceCollectionQueryMembershipRule -CollectionId $CollectionId -RuleName $RuleName -QueryExpression $QueryExpression -ErrorAction Stop
    Write-Log "Added query rule '$RuleName'"
}

function Remove-QueryRule {
    <#
    .SYNOPSIS
        Removes a query membership rule from a collection.
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)][string]$CollectionId,
        [Parameter(Mandatory)][string]$RuleName
    )

    if (-not $PSCmdlet.ShouldProcess($CollectionId, "Remove query rule '$RuleName'")) { return }
    Write-Log "Removing query rule '$RuleName' from collection $CollectionId..."

    Remove-CMDeviceCollectionQueryMembershipRule -CollectionId $CollectionId -RuleName $RuleName -Force -ErrorAction Stop
    Write-Log "Removed query rule '$RuleName'"
}

function Update-QueryRule {
    <#
    .SYNOPSIS
        Updates a query rule by removing the old one and adding a new one.
    .DESCRIPTION
        There is no native update cmdlet for query rules. This reads the old
        rule, removes it, and adds the replacement. When the site refuses the
        replacement, the old rule is added back and the error is thrown with a
        note on whether the old rule was restored. When the old rule does not
        exist, nothing changes and the function throws.
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)][string]$CollectionId,
        [Parameter(Mandatory)][string]$OldRuleName,
        [Parameter(Mandatory)][string]$NewRuleName,
        [Parameter(Mandatory)][string]$NewQueryExpression
    )

    $old = @(Get-CMDeviceCollectionQueryMembershipRule -CollectionId $CollectionId -RuleName $OldRuleName -ErrorAction Stop |
            Where-Object { [string]$_.RuleName -ieq $OldRuleName }) | Select-Object -First 1
    if (-not $old) { throw "Query rule '$OldRuleName' not found on collection $CollectionId." }

    if (-not $PSCmdlet.ShouldProcess($CollectionId, "Replace query rule '$OldRuleName' with '$NewRuleName'")) { return }
    Write-Log "Updating query rule '$OldRuleName' on collection $CollectionId..."

    Remove-CMDeviceCollectionQueryMembershipRule -CollectionId $CollectionId -RuleName $OldRuleName -Force -ErrorAction Stop
    try {
        Add-CMDeviceCollectionQueryMembershipRule -CollectionId $CollectionId -RuleName $NewRuleName -QueryExpression $NewQueryExpression -ErrorAction Stop
    }
    catch {
        $reason = $_.Exception.Message
        $restored = $false
        try {
            Add-CMDeviceCollectionQueryMembershipRule -CollectionId $CollectionId -RuleName $OldRuleName -QueryExpression ([string]$old.QueryExpression) -ErrorAction Stop
            $restored = $true
        } catch { Write-Log "Could not restore query rule '$OldRuleName': $($_.Exception.Message)" -Level ERROR }
        if ($restored) { throw "$reason (the old rule '$OldRuleName' was restored)" }
        throw "$reason (the old rule '$OldRuleName' was removed and not restored)"
    }

    Write-Log "Updated query rule '$OldRuleName' -> '$NewRuleName'"
}

function Add-IncludeRule {
    <#
    .SYNOPSIS
        Adds an include collection membership rule.
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)][string]$CollectionId,
        [Parameter(Mandatory)][string]$IncludeCollectionId
    )

    if (-not $PSCmdlet.ShouldProcess($CollectionId, "Include collection $IncludeCollectionId")) { return }
    Write-Log "Adding include rule for $IncludeCollectionId to collection $CollectionId..."

    Add-CMDeviceCollectionIncludeMembershipRule -CollectionId $CollectionId -IncludeCollectionId $IncludeCollectionId -ErrorAction Stop
    Write-Log "Added include rule"
}

function Add-ExcludeRule {
    <#
    .SYNOPSIS
        Adds an exclude collection membership rule.
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)][string]$CollectionId,
        [Parameter(Mandatory)][string]$ExcludeCollectionId
    )

    if (-not $PSCmdlet.ShouldProcess($CollectionId, "Exclude collection $ExcludeCollectionId")) { return }
    Write-Log "Adding exclude rule for $ExcludeCollectionId to collection $CollectionId..."

    Add-CMDeviceCollectionExcludeMembershipRule -CollectionId $CollectionId -ExcludeCollectionId $ExcludeCollectionId -ErrorAction Stop
    Write-Log "Added exclude rule"
}

function Invoke-CollectionEvaluation {
    <#
    .SYNOPSIS
        Forces immediate membership evaluation for a collection.
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)][string]$CollectionId
    )

    if (-not $PSCmdlet.ShouldProcess($CollectionId, 'Evaluate collection membership')) { return }
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
        # Read as UTF-8 and unroll the array in a second step: Windows PowerShell 5.1 decodes a
        # file without a byte order mark as ANSI and returns a JSON array as one object.
        $parsed = Get-Content -LiteralPath $TemplatesPath -Raw -Encoding UTF8 | ConvertFrom-Json
        $templates = @(@($parsed) | Where-Object { $null -ne $_ })
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
        # Read as UTF-8 and unroll the array in a second step: Windows PowerShell 5.1 decodes a
        # file without a byte order mark as ANSI and returns a JSON array as one object.
        $parsed = Get-Content -LiteralPath $TemplatesPath -Raw -Encoding UTF8 | ConvertFrom-Json
        $templates = @(@($parsed) | Where-Object { $null -ne $_ })
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

    # One pass over the template: a value that contains another token is not expanded again,
    # so a parameter value cannot pull in another parameter. The longest key wins at a position.
    if ($ParameterValues.Count -eq 0) { return $QueryTemplate }
    $keys = @($ParameterValues.Keys | Where-Object { $_ } | Sort-Object { $_.Length } -Descending | ForEach-Object { [regex]::Escape([string]$_) })
    $pattern = $keys -join '|'
    $expanded = [regex]::Replace($QueryTemplate, $pattern, {
            param($m)
            $v = $ParameterValues[$m.Value]
            if ($null -eq $v) { '' } else { [string]$v }
        })
    return $expanded
}

# ---------------------------------------------------------------------------
# Export
# ---------------------------------------------------------------------------

function ConvertTo-CollectionDataTable {
    <#
    .SYNOPSIS
        Builds a DataTable of text columns from objects, for the export functions.
    .DESCRIPTION
        Columns names the columns and their order. Without Columns the table
        takes every property of the first object. A missing property or a null
        value is an empty cell.
    #>
    param(
        [AllowNull()][AllowEmptyCollection()][object[]]$Objects,
        [string[]]$Columns
    )
    $items = @($Objects | Where-Object { $null -ne $_ })
    if (-not $Columns -or $Columns.Count -eq 0) {
        $Columns = if ($items.Count -gt 0) { @($items[0].PSObject.Properties.Name) } else { @() }
    }
    $table = New-Object System.Data.DataTable
    foreach ($c in $Columns) { [void]$table.Columns.Add([string]$c, [string]) }
    foreach ($item in $items) {
        $row = $table.NewRow()
        foreach ($c in $Columns) {
            $p = $item.PSObject.Properties[$c]
            $row[[string]$c] = if ($p -and $null -ne $p.Value) { [string]$p.Value } else { [DBNull]::Value }
        }
        [void]$table.Rows.Add($row)
    }
    return , $table
}
function Export-CollectionCsv {
    <#
    .SYNOPSIS
        Exports a DataTable, or collection objects, to CSV.
    .DESCRIPTION
        Pass -DataTable, or pass -Collections with an optional -Columns list.
    #>
    [CmdletBinding(DefaultParameterSetName = 'Table')]
    param(
        [Parameter(Mandatory, ParameterSetName = 'Table', Position = 0)][System.Data.DataTable]$DataTable,
        [Parameter(Mandatory, ParameterSetName = 'Objects')][AllowNull()][AllowEmptyCollection()][object[]]$Collections,
        [Parameter(ParameterSetName = 'Objects')][string[]]$Columns,
        [Parameter(Mandatory, Position = 1)][string]$OutputPath
    )

    if ($PSCmdlet.ParameterSetName -eq 'Objects') { $DataTable = ConvertTo-CollectionDataTable -Objects $Collections -Columns $Columns }

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
        Exports a DataTable, or collection objects, to a self-contained HTML report.
    .DESCRIPTION
        Pass -DataTable, or pass -Collections with an optional -Columns list.
    #>
    [CmdletBinding(DefaultParameterSetName = 'Table')]
    param(
        [Parameter(Mandatory, ParameterSetName = 'Table', Position = 0)][System.Data.DataTable]$DataTable,
        [Parameter(Mandatory, ParameterSetName = 'Objects')][AllowNull()][AllowEmptyCollection()][object[]]$Collections,
        [Parameter(ParameterSetName = 'Objects')][string[]]$Columns,
        [Parameter(Mandatory, Position = 1)][string]$OutputPath,
        [Parameter(Position = 2)][string]$ReportTitle = 'Collection Manager Report'
    )

    if ($PSCmdlet.ParameterSetName -eq 'Objects') { $DataTable = ConvertTo-CollectionDataTable -Objects $Collections -Columns $Columns }

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

    # HtmlEncode every site-derived string: names with angle brackets or
    # ampersands otherwise corrupt the report markup.
    $headerRow = ($DataTable.Columns | ForEach-Object { "<th>$([System.Net.WebUtility]::HtmlEncode($_.ColumnName))</th>" }) -join ''
    $bodyRows = foreach ($row in $DataTable.Rows) {
        $cells = foreach ($col in $DataTable.Columns) {
            $val = [string]$row[$col.ColumnName]
            "<td>$([System.Net.WebUtility]::HtmlEncode($val))</td>"
        }
        "<tr>$($cells -join '')</tr>"
    }

    $html = @(
        '<!DOCTYPE html>',
        '<html><head><meta charset="utf-8"><title>' + [System.Net.WebUtility]::HtmlEncode($ReportTitle) + '</title>',
        $css,
        '</head><body>',
        "<h1>$([System.Net.WebUtility]::HtmlEncode($ReportTitle))</h1>",
        "<div class='summary'>Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') | Rows: $($DataTable.Rows.Count)</div>",
        "<table><thead><tr>$headerRow</tr></thead>",
        "<tbody>$($bodyRows -join "`r`n")</tbody></table>",
        '</body></html>'
    ) -join "`r`n"

    Set-Content -LiteralPath $OutputPath -Value $html -Encoding UTF8
    Write-Log "Exported HTML to $OutputPath"
}
