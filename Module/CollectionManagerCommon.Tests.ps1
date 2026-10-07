#Requires -Modules Pester

<#
.SYNOPSIS
    Pester 5.x tests for CollectionManagerCommon shared module.

.DESCRIPTION
    Tests pure-logic functions: logging, template loading, parameter expansion,
    export. Does NOT require ConfigMgr, WMI, or administrator elevation.

.EXAMPLE
    Invoke-Pester .\CollectionManagerCommon.Tests.ps1
#>

BeforeAll {
    Import-Module "$PSScriptRoot\CollectionManagerCommon.psd1" -Force -DisableNameChecking
}

# ============================================================================
# Write-Log / Initialize-Logging
# ============================================================================

Describe 'Write-Log' {
    It 'writes formatted message to log file' {
        $logFile = Join-Path $TestDrive 'test.log'
        Initialize-Logging -LogPath $logFile
        Write-Log 'Hello world' -Quiet
        $content = Get-Content -LiteralPath $logFile -Raw
        $content | Should -Match '\[INFO \] Hello world'
    }

    It 'tags WARN messages correctly' {
        $logFile = Join-Path $TestDrive 'warn.log'
        Initialize-Logging -LogPath $logFile
        Write-Log 'Something odd' -Level WARN -Quiet
        $content = Get-Content -LiteralPath $logFile -Raw
        $content | Should -Match '\[WARN \] Something odd'
    }

    It 'tags ERROR messages correctly' {
        $logFile = Join-Path $TestDrive 'error.log'
        Initialize-Logging -LogPath $logFile
        Write-Log 'Failure' -Level ERROR -Quiet
        $content = Get-Content -LiteralPath $logFile -Raw
        $content | Should -Match '\[ERROR\] Failure'
    }

    It 'accepts empty string message' {
        $logFile = Join-Path $TestDrive 'empty.log'
        Initialize-Logging -LogPath $logFile
        { Write-Log '' -Quiet } | Should -Not -Throw
        $lines = Get-Content -LiteralPath $logFile
        $lines.Count | Should -BeGreaterOrEqual 2
    }
}

Describe 'Initialize-Logging' {
    It 'creates log file with header line' {
        $logFile = Join-Path $TestDrive 'init.log'
        Initialize-Logging -LogPath $logFile
        Test-Path -LiteralPath $logFile | Should -BeTrue
        $content = Get-Content -LiteralPath $logFile -Raw
        $content | Should -Match '\[INFO \] === Log initialized ==='
    }

    It 'creates parent directories if missing' {
        $logFile = Join-Path $TestDrive 'sub\dir\deep.log'
        Initialize-Logging -LogPath $logFile
        Test-Path -LiteralPath $logFile | Should -BeTrue
    }

    It '-Attach preserves an externally-created log file' {
        $logFile = Join-Path $TestDrive 'attach.log'
        $sentinel = "[2026-05-02 00:00:00] [INFO ] Shell-managed header"
        Set-Content -LiteralPath $logFile -Value $sentinel -Encoding UTF8
        Initialize-Logging -LogPath $logFile -Attach
        Write-Log 'Module appended line' -Quiet
        $content = Get-Content -LiteralPath $logFile -Raw
        $content | Should -Match 'Shell-managed header'
        $content | Should -Match 'Module appended line'
    }
}

# ============================================================================
# Get-OperationalTemplates
# ============================================================================

Describe 'Get-OperationalTemplates' {
    BeforeAll {
        $logFile = Join-Path $TestDrive 'tpl.log'
        Initialize-Logging -LogPath $logFile
        $tplPath = Join-Path (Split-Path $PSScriptRoot -Parent) 'Templates\operational-collections.json'
        $script:opTemplates = @(Get-OperationalTemplates -TemplatesPath $tplPath)
    }

    It 'loads templates from JSON file' {
        $script:opTemplates.Count | Should -BeGreaterThan 100
    }

    It 'each template has required properties' {
        foreach ($t in $script:opTemplates | Select-Object -First 10) {
            $t.Name | Should -Not -BeNullOrEmpty
            $t.Category | Should -Not -BeNullOrEmpty
            $t.Query | Should -Not -BeNullOrEmpty
            $t.LimitingCollection | Should -Not -BeNullOrEmpty
        }
    }

    It 'contains expected categories' {
        $categories = $script:opTemplates | ForEach-Object { $_.Category } | Sort-Object -Unique
        $categories | Should -Contain 'Clients'
        $categories | Should -Contain 'Workstations'
        $categories | Should -Contain 'Servers'
    }

    It 'Clients | All query matches expected pattern' {
        $clientsAll = $script:opTemplates | Where-Object { $_.Name -eq 'Clients | All' }
        $clientsAll | Should -Not -BeNullOrEmpty
        $clientsAll.Query | Should -Match 'SMS_R_System.Client = 1'
    }

    It 'returns empty array for missing file' {
        $result = @(Get-OperationalTemplates -TemplatesPath 'C:\nonexistent\file.json')
        $result.Count | Should -Be 0
    }
}

# ============================================================================
# Get-ParameterizedTemplates
# ============================================================================

Describe 'Get-ParameterizedTemplates' {
    BeforeAll {
        $logFile = Join-Path $TestDrive 'ptpl.log'
        Initialize-Logging -LogPath $logFile
        $tplPath = Join-Path (Split-Path $PSScriptRoot -Parent) 'Templates\parameterized-templates.json'
        $script:paramTemplates = @(Get-ParameterizedTemplates -TemplatesPath $tplPath)
    }

    It 'loads 20 parameterized templates' {
        $script:paramTemplates.Count | Should -Be 20
    }

    It 'each template has required properties' {
        foreach ($t in $script:paramTemplates) {
            $t.Name | Should -Not -BeNullOrEmpty
            $t.Category | Should -Not -BeNullOrEmpty
            $t.Query | Should -Not -BeNullOrEmpty
        }
    }

    It 'parameterized templates have Parameters array' {
        $withParams = @($script:paramTemplates | Where-Object { $_.Parameters -and $_.Parameters.Count -gt 0 })
        $withParams.Count | Should -BeGreaterOrEqual 15
    }

    It 'Software Installed by Name has SoftwareName parameter' {
        $swTpl = $script:paramTemplates | Where-Object { $_.Name -eq 'Software Installed by Name' }
        $swTpl | Should -Not -BeNullOrEmpty
        $swTpl.Parameters[0].Placeholder | Should -Be '{SoftwareName}'
    }

    It 'returns empty array for missing file' {
        $result = @(Get-ParameterizedTemplates -TemplatesPath 'C:\nonexistent\file.json')
        $result.Count | Should -Be 0
    }
}

# ============================================================================
# Expand-TemplateParameters
# ============================================================================

Describe 'Expand-TemplateParameters' {
    It 'substitutes single placeholder' {
        $query = "select * from SMS_R_System where Name like '%{DeviceName}%'"
        $result = Expand-TemplateParameters -QueryTemplate $query -ParameterValues @{ '{DeviceName}' = 'WORKSTATION01' }
        $result | Should -Be "select * from SMS_R_System where Name like '%WORKSTATION01%'"
    }

    It 'substitutes multiple placeholders' {
        $query = "where DisplayName like '%{SoftwareName}%' AND Version < '{MaxVersion}'"
        $result = Expand-TemplateParameters -QueryTemplate $query -ParameterValues @{
            '{SoftwareName}' = 'Java'
            '{MaxVersion}'   = '8.0'
        }
        $result | Should -Match 'Java'
        $result | Should -Match '8\.0'
        $result | Should -Not -Match '\{SoftwareName\}'
        $result | Should -Not -Match '\{MaxVersion\}'
    }

    It 'handles empty parameter values' {
        $query = "where Name = '{Name}'"
        $result = Expand-TemplateParameters -QueryTemplate $query -ParameterValues @{ '{Name}' = '' }
        $result | Should -Be "where Name = ''"
    }

    It 'leaves unmatched placeholders intact' {
        $query = "where Name = '{Name}' and Domain = '{Domain}'"
        $result = Expand-TemplateParameters -QueryTemplate $query -ParameterValues @{ '{Name}' = 'PC01' }
        $result | Should -Match 'PC01'
        $result | Should -Match '\{Domain\}'
    }
}

# ============================================================================
# Export-CollectionCsv
# ============================================================================

Describe 'Export-CollectionCsv' {
    It 'writes CSV file with correct columns and rows' {
        $dt = New-Object System.Data.DataTable
        [void]$dt.Columns.Add("Name", [string]); [void]$dt.Columns.Add("CollectionID", [string])
        [void]$dt.Rows.Add("Test Collection", "MCM00001")
        [void]$dt.Rows.Add("Another", "MCM00002")

        $csvPath = Join-Path $TestDrive 'export.csv'
        $logFile = Join-Path $TestDrive 'csv.log'
        Initialize-Logging -LogPath $logFile

        Export-CollectionCsv -DataTable $dt -OutputPath $csvPath
        Test-Path -LiteralPath $csvPath | Should -BeTrue
        $rows = Import-Csv -LiteralPath $csvPath
        $rows.Count | Should -Be 2
        $rows[0].Name | Should -Be 'Test Collection'
        $rows[1].CollectionID | Should -Be 'MCM00002'
    }
}

# ============================================================================
# Export-CollectionHtml
# ============================================================================

Describe 'Export-CollectionHtml' {
    It 'writes valid HTML file' {
        $dt = New-Object System.Data.DataTable
        [void]$dt.Columns.Add("Name", [string]); [void]$dt.Columns.Add("Members", [int])
        [void]$dt.Rows.Add("Collection A", 50)

        $htmlPath = Join-Path $TestDrive 'report.html'
        $logFile = Join-Path $TestDrive 'html.log'
        Initialize-Logging -LogPath $logFile

        Export-CollectionHtml -DataTable $dt -OutputPath $htmlPath -ReportTitle 'Test Report'
        Test-Path -LiteralPath $htmlPath | Should -BeTrue
        $content = Get-Content -LiteralPath $htmlPath -Raw
        $content | Should -Match 'Test Report'
        $content | Should -Match '<th>Name</th>'
        $content | Should -Match 'Collection A'
    }
}

# ============================================================================
# Saved site queries
# ============================================================================

Describe 'Saved site queries' {
    BeforeAll {
        InModuleScope CollectionManagerCommon {
            function Get-CMQuery {
                [CmdletBinding()]
                param([string]$Id, [string]$Name, [switch]$DisableWildcardHandling)
                throw 'Get-CMQuery must be mocked by the test.'
            }
            function New-CMQuery {
                [CmdletBinding()]
                param([string]$Name, [string]$Expression, [string]$Comment, [string]$TargetClassName, [string]$LimitToCollectionId)
                throw 'New-CMQuery must be mocked by the test.'
            }
            function Set-CMQuery {
                [CmdletBinding()]
                param([string]$Id, [string]$NewName, [string]$Expression, [string]$Comment, [string]$TargetClassName, [string]$LimitToCollectionId, [switch]$PassThru)
                throw 'Set-CMQuery must be mocked by the test.'
            }
            function Remove-CMQuery {
                [CmdletBinding()]
                param([string]$Id, [switch]$Force)
                throw 'Remove-CMQuery must be mocked by the test.'
            }
        }
    }

    BeforeEach {
        Mock Get-CMQuery -ModuleName CollectionManagerCommon {
            @([PSCustomObject]@{
                QueryID = 'ABC00001'
                Name = 'Windows Workstations'
                Expression = 'select * from SMS_R_System'
                Comments = 'Reusable device query'
                TargetClassName = 'SMS_R_System'
                LimitToCollectionID = '<Prompt>'
            })
        }
        Mock New-CMQuery -ModuleName CollectionManagerCommon {
            [PSCustomObject]@{
                QueryID = 'ABC00002'
                Name = $Name
                Expression = $Expression
                Comments = $Comment
                TargetClassName = $TargetClassName
                LimitToCollectionID = $LimitToCollectionId
            }
        }
        Mock Set-CMQuery -ModuleName CollectionManagerCommon {
            [PSCustomObject]@{
                QueryID = $Id
                Name = $NewName
                Expression = $Expression
                Comments = $Comment
                TargetClassName = $TargetClassName
                LimitToCollectionID = $LimitToCollectionId
            }
        }
        Mock Remove-CMQuery -ModuleName CollectionManagerCommon {}
    }

    It 'lists site queries with their reusable WQL and metadata' {
        $queries = @(Get-SiteSavedQueries)
        $queries.Count | Should -Be 1
        $queries[0].QueryID | Should -Be 'ABC00001'
        $queries[0].Expression | Should -Be 'select * from SMS_R_System'
        $queries[0].Comment | Should -Be 'Reusable device query'
        $queries[0].LimitToCollectionID | Should -Be '<Prompt>'
    }

    It 'creates a site query with its metadata' {
        $query = New-SiteSavedQuery -Name 'Windows Workstations' -Expression 'select * from SMS_R_System' `
            -Comment 'Reusable device query' -TargetClassName 'SMS_R_System' -LimitToCollectionId '<Prompt>'
        $query.QueryID | Should -Be 'ABC00002'
        $query.Name | Should -Be 'Windows Workstations'
        Should -Invoke New-CMQuery -ModuleName CollectionManagerCommon -Times 1 -ParameterFilter {
            $Name -eq 'Windows Workstations' -and $Expression -eq 'select * from SMS_R_System' -and
            $Comment -eq 'Reusable device query' -and $LimitToCollectionId -eq '<Prompt>'
        }
    }

    It 'updates the selected site query including metadata fields' {
        $query = Set-SiteSavedQuery -QueryId 'ABC00001' -Name 'Windows Workstations v2' `
            -Expression 'select * from SMS_R_System where Client = 1' -Comment '' `
            -TargetClassName 'SMS_R_System' -LimitToCollectionId ''
        $query.Name | Should -Be 'Windows Workstations v2'
        $query.LimitToCollectionID | Should -Be ''
        Should -Invoke Set-CMQuery -ModuleName CollectionManagerCommon -Times 1 -ParameterFilter {
            $Id -eq 'ABC00001' -and $NewName -eq 'Windows Workstations v2' -and
            $Expression -eq 'select * from SMS_R_System where Client = 1' -and $LimitToCollectionId -eq ''
        }
    }

    It 'deletes a site query by ID' {
        Remove-SiteSavedQuery -QueryId 'ABC00001' | Should -BeTrue
        Should -Invoke Remove-CMQuery -ModuleName CollectionManagerCommon -Times 1 -ParameterFilter {
            $Id -eq 'ABC00001' -and $Force
        }
    }
}
