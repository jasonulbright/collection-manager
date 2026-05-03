@{
    RootModule        = 'CollectionManagerCommon.psm1'
    ModuleVersion     = '1.0.0'
    GUID              = 'b2c3d4e5-f6a7-8901-bcde-f23456789012'
    Author            = 'Jason Ulbright'
    Description       = 'Collection management and offline WQL editor for MECM device collections.'
    PowerShellVersion = '5.1'

    FunctionsToExport = @(
        # Logging
        'Initialize-Logging'
        'Write-Log'

        # CM Connection
        'Connect-CMSite'
        'Disconnect-CMSite'
        'Test-CMConnection'

        # Collection Queries
        'Get-AllDeviceCollections'
        'Get-CollectionDetail'
        'Get-CollectionQueryRules'
        'Get-CollectionMembers'

        # Folder hierarchy
        'Get-CMCollectionFolderTree'
        'Get-CMCollectionFolderMap'

        # Collection CRUD
        'New-ManagedCollection'
        'Copy-ManagedCollection'
        'Remove-ManagedCollection'
        'Set-CollectionProperties'

        # Membership Management
        'Add-DirectMember'
        'Remove-DirectMember'
        'Add-QueryRule'
        'Remove-QueryRule'
        'Update-QueryRule'
        'Add-IncludeRule'
        'Add-ExcludeRule'
        'Invoke-CollectionEvaluation'

        # WQL
        'Test-WqlQuery'
        'Invoke-WqlPreview'

        # Templates
        'Get-OperationalTemplates'
        'Get-ParameterizedTemplates'
        'Expand-TemplateParameters'

        # Export
        'Export-CollectionCsv'
        'Export-CollectionHtml'
    )

    CmdletsToExport   = @()
    VariablesToExport  = @()
    AliasesToExport    = @()
}
