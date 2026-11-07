#Requires -Modules Pester
# WQL validation and preview, template loading and expansion, CSV and HTML export.

BeforeAll {
    . "$PSScriptRoot\TestHelpers.ps1"
    Import-AppModule
    $script:Stubs = New-CmStub -Name Invoke-CMWmiQuery
    $global:TestLogs = [System.Collections.ArrayList]::new()
    Mock Write-Log { [void]$global:TestLogs.Add(@{ Message = $Message; Level = $Level }) } -ModuleName CollectionManagerCommon
}

AfterAll {
    Remove-CmStub -Name $script:Stubs
    Remove-Variable -Name Fx, TestLogs -Scope Global -ErrorAction SilentlyContinue
}

Describe 'Test-WqlQuery' {
    BeforeEach { $global:TestLogs.Clear() }

    It 'reports a valid query with its result count' {
        Mock Invoke-CMWmiQuery { 1..3 | ForEach-Object { [PSCustomObject]@{ Name = "PC$_" } } } -ModuleName CollectionManagerCommon
        $r = Test-WqlQuery -QueryExpression 'select * from SMS_R_System'
        $r.IsValid | Should -BeTrue
        $r.ResultCount | Should -Be 3
        $r.ErrorMessage | Should -Be ''
        Should -Invoke Invoke-CMWmiQuery -ModuleName CollectionManagerCommon -ParameterFilter { $Option -eq 'Lazy' -and $Query -eq 'select * from SMS_R_System' }
    }

    It 'reports a valid query with no result' {
        Mock Invoke-CMWmiQuery { } -ModuleName CollectionManagerCommon
        $r = Test-WqlQuery -QueryExpression 'select * from SMS_R_System where Name = "x"'
        $r.IsValid | Should -BeTrue
        $r.ResultCount | Should -Be 0
    }

    It 'reports an invalid query with the message and logs a warning' {
        Mock Invoke-CMWmiQuery { throw 'Invalid query "WQL"' } -ModuleName CollectionManagerCommon
        $r = Test-WqlQuery -QueryExpression 'selec nonsense'
        $r.IsValid | Should -BeFalse
        $r.ResultCount | Should -Be 0
        $r.ErrorMessage | Should -Be 'Invalid query "WQL"'
        @($global:TestLogs | Where-Object Level -eq 'WARN').Count | Should -Be 1
    }

    It 'rejects an empty expression' {
        { Test-WqlQuery -QueryExpression '' } | Should -Throw
    }

    It 'sends a long expression of <N> characters unchanged' -ForEach @(@{ N = 5000 }, @{ N = 100000 }) {
        $global:Fx = 'select ' + ('x' * $N)
        Mock Invoke-CMWmiQuery { } -ModuleName CollectionManagerCommon
        $null = Test-WqlQuery -QueryExpression $global:Fx
        Should -Invoke Invoke-CMWmiQuery -ModuleName CollectionManagerCommon -ParameterFilter { $Query -ceq $global:Fx }
    }
}

Describe 'Invoke-WqlPreview' {
    BeforeEach {
        $global:TestLogs.Clear()
        Mock Invoke-CMWmiQuery { 1..120 | ForEach-Object { [PSCustomObject]@{ Name = "PC$_"; ResourceID = $_; ResourceDomainORWorkgroup = 'CONTOSO'; Client = 1 } } } -ModuleName CollectionManagerCommon
    }

    It 'limits the preview to <Max> and reports the total' -ForEach @(
        @{ Max = 1;   Shown = 1 }
        @{ Max = 50;  Shown = 50 }
        @{ Max = 119; Shown = 119 }
        @{ Max = 120; Shown = 120 }
        @{ Max = 121; Shown = 120 }
        @{ Max = 5000; Shown = 120 }
    ) {
        $r = Invoke-WqlPreview -QueryExpression 'select 1' -MaxResults $Max
        $r.TotalCount | Should -Be 120
        @($r.Results).Count | Should -Be $Shown
    }

    It 'shows 50 rows by default' {
        @((Invoke-WqlPreview -QueryExpression 'select 1').Results).Count | Should -Be 50
    }

    It 'maps the result columns' {
        $row = (Invoke-WqlPreview -QueryExpression 'select 1' -MaxResults 1).Results | Select-Object -First 1
        $row.Name | Should -Be 'PC1'
        $row.ResourceID | Should -Be 1
        $row.Domain | Should -Be 'CONTOSO'
        $row.Client | Should -Be 1
    }

    It 'returns an empty preview for a query without result' {
        Mock Invoke-CMWmiQuery { } -ModuleName CollectionManagerCommon
        $r = Invoke-WqlPreview -QueryExpression 'select 1'
        $r.TotalCount | Should -Be 0
        @($r.Results).Count | Should -Be 0
    }

    It 'returns an empty preview and logs an error when the query fails' {
        Mock Invoke-CMWmiQuery { throw 'bad query' } -ModuleName CollectionManagerCommon
        $r = Invoke-WqlPreview -QueryExpression 'x'
        $r.TotalCount | Should -Be 0
        @($r.Results).Count | Should -Be 0
        @($global:TestLogs | Where-Object Level -eq 'ERROR').Count | Should -Be 1
    }
}

Describe 'Template files' {
    It '<Fn> returns <Case> for <Content>' -ForEach @(
        @{ Fn = 'Get-OperationalTemplates';    Case = 'the templates'; Content = '[{"Name":"a","Query":"q"},{"Name":"b","Query":"r"}]'; Expected = 2 }
        @{ Fn = 'Get-ParameterizedTemplates';  Case = 'the templates'; Content = '[{"Name":"a","Query":"q"}]'; Expected = 1 }
        @{ Fn = 'Get-OperationalTemplates';    Case = 'an empty list'; Content = '[]'; Expected = 0 }
        @{ Fn = 'Get-OperationalTemplates';    Case = 'an empty list'; Content = '{ not json'; Expected = 0 }
        @{ Fn = 'Get-ParameterizedTemplates';  Case = 'an empty list'; Content = '{ not json'; Expected = 0 }
        @{ Fn = 'Get-OperationalTemplates';    Case = 'an empty list'; Content = ''; Expected = 0 }
    ) {
        $path = Join-Path $TestDrive ('t-{0}.json' -f [guid]::NewGuid().ToString('N'))
        Set-Content -LiteralPath $path -Value $Content
        @(& $Fn -TemplatesPath $path).Count | Should -Be $Expected
    }

    It 'logs an error for a corrupt file and a warning for a missing file' {
        $global:TestLogs.Clear()
        $bad = Join-Path $TestDrive 'bad.json'; Set-Content -LiteralPath $bad -Value '{ nope'
        $null = Get-OperationalTemplates -TemplatesPath $bad
        $null = Get-OperationalTemplates -TemplatesPath (Join-Path $TestDrive 'missing.json')
        @($global:TestLogs | Where-Object Level -eq 'ERROR').Count | Should -Be 1
        @($global:TestLogs | Where-Object Level -eq 'WARN').Count | Should -Be 1
    }

    It 'reads unicode text and a path with spaces and brackets' {
        $path = Join-Path $TestDrive 'dir [1] x\t.json'
        New-Item -ItemType Directory -Path (Split-Path $path) -Force | Out-Null
        [IO.File]::WriteAllText($path, '[{"Name":"Ünï 日本","Query":"q"}]', (New-Object Text.UTF8Encoding $false))
        (@(Get-OperationalTemplates -TemplatesPath $path))[0].Name | Should -Be 'Ünï 日本'
    }

    It 'loads the shipped <Fn> file' -ForEach @(
        @{ Fn = 'Get-OperationalTemplates' }, @{ Fn = 'Get-ParameterizedTemplates' }
    ) {
        @(& $Fn).Count | Should -BeGreaterThan 10
    }

    It 'has a name and a query on every shipped template, and no name twice' {
        foreach ($fn in 'Get-OperationalTemplates', 'Get-ParameterizedTemplates') {
            $t = @(& $fn)
            ($t | Where-Object { -not $_.Name }) | Should -BeNullOrEmpty
            ($t | Where-Object { -not ($_.Query -or $_.QueryTemplate -or $_.WQL -or $_.Expression) }) | Should -BeNullOrEmpty
            ($t | Group-Object Name | Where-Object Count -gt 1 | ForEach-Object Name) | Should -BeNullOrEmpty -Because "$fn has a name twice"
        }
    }
}

Describe 'Expand-TemplateParameters' {
    It 'replaces <Name>' -ForEach @(
        @{ Name = 'one token';        Q = "Name = '{A}'";        V = @{ '{A}' = 'PC1' };                  Expected = "Name = 'PC1'" }
        @{ Name = 'a repeated token'; Q = '{A} and {A}';         V = @{ '{A}' = 'x' };                    Expected = 'x and x' }
        @{ Name = 'two tokens';       Q = '{A}-{B}';             V = @{ '{A}' = '1'; '{B}' = '2' };       Expected = '1-2' }
        @{ Name = 'a regex key';      Q = 'a.b {a.b}';           V = @{ '{a.b}' = 'z' };                  Expected = 'a.b z' }
        @{ Name = 'a $1 value';       Q = 'x {A} y';             V = @{ '{A}' = '$1$&' };                 Expected = 'x $1$& y' }
        @{ Name = 'an empty value';   Q = 'a{A}b';               V = @{ '{A}' = '' };                     Expected = 'ab' }
        @{ Name = 'no parameters';    Q = 'select 1';            V = @{};                                 Expected = 'select 1' }
        @{ Name = 'an unknown token'; Q = '{A} {B}';             V = @{ '{A}' = 'x' };                    Expected = 'x {B}' }
        @{ Name = 'a quote value';    Q = "'{A}'";               V = @{ '{A}' = "O'Brien" };              Expected = "'O'Brien'" }
        @{ Name = 'unicode';          Q = '{A}';                 V = @{ '{A}' = 'Ünï 日本' };             Expected = 'Ünï 日本' }
    ) {
        Expand-TemplateParameters -QueryTemplate $Q -ParameterValues $V | Should -BeExactly $Expected
    }

    It 'does not expand a token that a value introduces, whatever the key order' -ForEach @(
        @{ Order = 'AB'; V = ([ordered]@{ '{A}' = '{B}'; '{B}' = 'x' }) }
        @{ Order = 'BA'; V = ([ordered]@{ '{B}' = 'x'; '{A}' = '{B}' }) }
    ) {
        Expand-TemplateParameters -QueryTemplate '{A} {B}' -ParameterValues $V | Should -BeExactly '{B} x'
    }

    It 'takes the longest key at a position' {
        Expand-TemplateParameters -QueryTemplate '{AB}{A}' -ParameterValues @{ '{A}' = '1'; '{AB}' = '2' } | Should -BeExactly '21'
    }

    It 'expands a null value to nothing' {
        Expand-TemplateParameters -QueryTemplate 'a{A}b' -ParameterValues @{ '{A}' = $null } | Should -BeExactly 'ab'
    }

    It 'rejects an empty template' {
        { Expand-TemplateParameters -QueryTemplate '' -ParameterValues @{} } | Should -Throw
    }
}

Describe 'Export-CollectionCsv' {
    It 'writes the columns in table order and one line per row' {
        $t = New-RollupTable -Column Name, CollectionID, MemberCount -Row @(@('A', 'MCM00001', 3), @('B', 'MCM00002', 0))
        $path = Join-Path $TestDrive 'a.csv'
        Export-CollectionCsv -DataTable $t -OutputPath $path
        $rows = @(Import-Csv -LiteralPath $path)
        $rows.Count | Should -Be 2
        ($rows[0].PSObject.Properties.Name -join ',') | Should -Be 'Name,CollectionID,MemberCount'
        $rows[0].MemberCount | Should -Be '3'
    }

    It 'creates the missing parent folders' {
        $path = Join-Path $TestDrive 'x\y\z\a.csv'
        Export-CollectionCsv -DataTable (New-RollupTable -Column A -Row @(, @('x'))) -OutputPath $path
        Test-Path -LiteralPath $path | Should -BeTrue
    }

    It 'round-trips the hostile value <Name>' -ForEach @(
        @{ Name = 'comma'; Value = 'a,b' }
        @{ Name = 'quote'; Value = 'say "hi"' }
        @{ Name = 'newline'; Value = "l1`r`nl2" }
        @{ Name = 'unicode'; Value = 'Ünï 日本語' }
        @{ Name = 'long'; Value = ('x' * 20000) }
    ) {
        $path = Join-Path $TestDrive "h-$Name.csv"
        Export-CollectionCsv -DataTable (New-RollupTable -Column V -Row @(, @($Value))) -OutputPath $path
        (Import-Csv -LiteralPath $path)[0].V | Should -Be $Value
    }

    It 'throws on an invalid path character' {
        { Export-CollectionCsv -DataTable (New-RollupTable -Column A -Row @(, @('x'))) -OutputPath (Join-Path $TestDrive 'bad|name?.csv') } | Should -Throw
    }
}

Describe 'Export-CollectionHtml' {
    It 'writes one header cell per column and one row per data row' {
        $t = New-RollupTable -Column Name, MemberCount -Row @(@('A', 1), @('B', 2))
        $path = Join-Path $TestDrive 'a.html'
        Export-CollectionHtml -DataTable $t -OutputPath $path
        $html = Get-Content -LiteralPath $path -Raw
        ([regex]::Matches($html, '<th>')).Count | Should -Be 2
        ([regex]::Matches($html, '<tr>')).Count | Should -Be 3
        $html | Should -Match 'Rows: 2'
        $html | Should -Match '<title>\s*Collection Manager Report\s*</title>'
    }

    It 'encodes the value <Value>' -ForEach @(
        @{ Value = '<script>alert(1)</script>'; Expected = '&lt;script&gt;alert(1)&lt;/script&gt;' }
        @{ Value = 'A & B'; Expected = 'A &amp; B' }
        @{ Value = '"q"'; Expected = '&quot;q&quot;' }
        @{ Value = '</td><td>x'; Expected = '&lt;/td&gt;&lt;td&gt;x' }
    ) {
        $path = Join-Path $TestDrive 'enc.html'
        Export-CollectionHtml -DataTable (New-RollupTable -Column Name -Row @(, @($Value))) -OutputPath $path
        (Get-Content -LiteralPath $path -Raw) | Should -Match ([regex]::Escape("<td>$Expected</td>"))
    }

    It 'encodes a column name and the report title' {
        $path = Join-Path $TestDrive 'enc2.html'
        Export-CollectionHtml -DataTable (New-RollupTable -Column '<b>C</b>' -Row @(, @('x'))) -OutputPath $path -ReportTitle 'R&D <x>'
        $html = Get-Content -LiteralPath $path -Raw
        $html | Should -Match '<th>&lt;b&gt;C&lt;/b&gt;</th>'
        $html | Should -Match '<title>\s*R&amp;D &lt;x&gt;\s*</title>'
        $html | Should -Match '<h1>R&amp;D &lt;x&gt;</h1>'
    }

    It 'writes a valid page for an empty table and for a null cell' {
        $path = Join-Path $TestDrive 'e.html'
        Export-CollectionHtml -DataTable (New-RollupTable -Column A, B -Row @(, @('x', [DBNull]::Value))) -OutputPath $path
        (Get-Content -LiteralPath $path -Raw) | Should -Match '<td>x</td><td></td>'
        Export-CollectionHtml -DataTable (New-RollupTable -Column A -Row @()) -OutputPath $path
        (Get-Content -LiteralPath $path -Raw) | Should -Match 'Rows: 0'
    }
}

Describe 'Export from collection objects (the Export buttons)' {
    BeforeAll {
        $script:rows = @(
            [PSCustomObject]@{ Name = 'A'; CollectionID = 'MCM00001'; MemberCount = 3; Comment = $null; TypeGlyph = 'x' },
            [PSCustomObject]@{ Name = 'B, "q"'; CollectionID = 'MCM00002'; MemberCount = 0; Comment = 'c' })
        $script:cols = 'Name', 'CollectionID', 'MemberCount', 'Comment'
    }

    It 'writes only the requested columns to CSV, in order' {
        $path = Join-Path $TestDrive 'o.csv'
        Export-CollectionCsv -Collections $script:rows -Columns $script:cols -OutputPath $path
        $r = @(Import-Csv -LiteralPath $path)
        ($r[0].PSObject.Properties.Name -join ',') | Should -Be 'Name,CollectionID,MemberCount,Comment'
        $r[1].Name | Should -Be 'B, "q"'
        $r[0].Comment | Should -Be ''
        $r[0].MemberCount | Should -Be '3'
    }

    It 'takes every property of the first object when no columns are given' {
        $path = Join-Path $TestDrive 'o2.csv'
        Export-CollectionCsv -Collections $script:rows -OutputPath $path
        (@(Import-Csv -LiteralPath $path))[0].PSObject.Properties.Name | Should -Contain 'TypeGlyph'
    }

    It 'leaves a cell empty when the object has no such property' {
        $path = Join-Path $TestDrive 'o3.csv'
        Export-CollectionCsv -Collections @([PSCustomObject]@{ Name = 'A' }) -Columns 'Name', 'Missing' -OutputPath $path
        (@(Import-Csv -LiteralPath $path))[0].Missing | Should -Be ''
    }

    It 'writes a file for an empty list without an error' {
        $path = Join-Path $TestDrive 'o4.csv'
        { Export-CollectionCsv -Collections @() -Columns $script:cols -OutputPath $path } | Should -Not -Throw
        Test-Path -LiteralPath $path | Should -BeTrue
    }

    It 'writes the requested columns to HTML and encodes the values' {
        $path = Join-Path $TestDrive 'o.html'
        Export-CollectionHtml -Collections $script:rows -Columns $script:cols -OutputPath $path
        $html = Get-Content -LiteralPath $path -Raw
        ([regex]::Matches($html, '<th>')).Count | Should -Be 4
        $html | Should -Match 'Rows: 2'
        $html | Should -Match ([regex]::Escape('<td>B, &quot;q&quot;</td>'))
        $html | Should -Not -Match 'TypeGlyph'
    }

    It 'rejects a data table together with collection objects' {
        { Export-CollectionCsv -DataTable (New-RollupTable -Column A -Row @(, @('x'))) -Collections $script:rows -OutputPath (Join-Path $TestDrive 'x.csv') } | Should -Throw
    }
}
Describe 'Exporters bind positionally' {
    It 'Export-CollectionCsv takes the table and the path by position' {
        $path = Join-Path $TestDrive 'pos.csv'
        Export-CollectionCsv (New-RollupTable -Column A -Row @(, @('x'))) $path
        (Import-Csv -LiteralPath $path)[0].A | Should -Be 'x'
    }

    It 'Export-CollectionHtml takes the table, the path and the title by position' {
        $path = Join-Path $TestDrive 'pos.html'
        Export-CollectionHtml (New-RollupTable -Column A -Row @(, @('x'))) $path 'My title'
        (Get-Content -LiteralPath $path -Raw) | Should -Match '<h1>My title</h1>'
    }
}
