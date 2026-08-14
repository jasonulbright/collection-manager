@{
    RootModule        = 'CollectionManagerCommon.psm1'
    ModuleVersion     = '1.1.0'
    GUID              = '3f8e6a52-9d14-4c7b-b0e9-6a1d2c8f5b73'
    Author            = 'Jason Ulbright'
    Description       = 'Collection management and offline WQL editor for MECM device collections.'
    PowerShellVersion = '5.1'

    FunctionsToExport = @(
        # Logging and CM connection come from the vendored SuiteCommon
        # module (Lib\SuiteCommon), imported globally by the root module.

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
