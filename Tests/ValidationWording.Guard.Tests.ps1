#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.1.0' }
# Guard test for the validation wording. It reads text files only. It creates no WPF control
# and opens no window, so it runs on Windows PowerShell 5.1 and on PowerShell 7.
#
# Rules:
#   1. The WQL editor is local, but validation runs the query on the site (Test-WqlQuery calls
#      Invoke-CMWmiQuery). No text of the tool calls the validation offline.
#   2. The About text says that validation runs on the site.
#   3. The About version line is read from the header of the script, not typed by hand.

BeforeAll {
    $script:Root = Split-Path -Path $PSScriptRoot -Parent
    $script:strScript = Get-Content -LiteralPath (Join-Path -Path $script:Root -ChildPath 'start-collectionmanager.ps1') -Raw
}

Describe 'Validation wording' {
    It 'does not call the editor or the validation offline in <_>' -ForEach @(
        'start-collectionmanager.ps1', 'README.md', 'Module\CollectionManagerCommon.psm1', 'Module\CollectionManagerCommon.psd1'
    ) {
        $strText = Get-Content -LiteralPath (Join-Path -Path $script:Root -ChildPath $_) -Raw
        $strText | Should -Not -Match '(?i)\boffline\b'
    }

    It 'has a validation function that runs the query on the site' {
        $strModule = Get-Content -LiteralPath (Join-Path -Path $script:Root -ChildPath 'Module\CollectionManagerCommon.psm1') -Raw
        $strModule | Should -Match '(?s)function Test-WqlQuery.*?\$results = @\(Invoke-CMWmiQuery -Query \$QueryExpression'
    }

    It 'says in the About text that validation runs on the site' {
        $strAbout = [regex]::Match($script:strScript, '(?s)<StackPanel x:Name="paneAbout".*?</StackPanel>').Value
        $strAbout | Should -Not -BeNullOrEmpty
        $strAbout | Should -Match 'Validation runs the query on the site\. Preview Results shows the returned rows\.'
    }

    It 'says in every description that validation runs on the site: <_>' -ForEach @(
        'README.md', 'Module\CollectionManagerCommon.psm1', 'Module\CollectionManagerCommon.psd1'
    ) {
        $strText = Get-Content -LiteralPath (Join-Path -Path $script:Root -ChildPath $_) -Raw
        $strText | Should -Match '(?i)validation runs the query on the site'
    }

    It 'reads the About version from the script header' {
        $script:strScript | Should -Match '\$txtAboutVersion\.Text = .*\$script:AppVersion'
        $script:strScript | Should -Match "(?m)^\s*Version\s*:\s*[0-9][0-9\.]*[0-9]\s*$"
    }

    It 'keeps the module version equal to the script header version' {
        $strHeader = [regex]::Match($script:strScript, '(?m)^\s*Version\s*:\s*([0-9][0-9\.]*[0-9])\s*$').Groups[1].Value
        $strManifest = Get-Content -LiteralPath (Join-Path -Path $script:Root -ChildPath 'Module\CollectionManagerCommon.psd1') -Raw
        $strManifest | Should -Match ("ModuleVersion\s*=\s*'" + [regex]::Escape($strHeader) + "'")
    }
}
