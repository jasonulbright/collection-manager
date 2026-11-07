#Requires -Modules Pester
# Read side: folders, collections, detail, rules, members and saved site queries.
# The ConfigMgr cmdlets are replaced by stubs and mocked. Fixture values reach
# the mock bodies through $global:Fx, because a mock body runs in the module scope.

BeforeAll {
    . "$PSScriptRoot\TestHelpers.ps1"
    Import-AppModule
    $script:Stubs = New-CmStub -Name Get-CMDeviceCollection, Get-CMDeviceCollectionQueryMembershipRule, Get-CMCollectionMember, Get-CMQuery, New-CMQuery, Set-CMQuery, Remove-CMQuery
    Mock Write-Log {} -ModuleName CollectionManagerCommon
}

AfterAll {
    Remove-CmStub -Name $script:Stubs
    Remove-Variable -Name Fx, TestLogs -Scope Global -ErrorAction SilentlyContinue
}

Describe 'Get-CMCollectionFolderTree' {
    It 'maps the container node columns to integers and a name' {
        Mock Get-CimInstance { [PSCustomObject]@{ ContainerNodeID = '7'; Name = 'Servers'; ParentContainerNodeID = '0' } } -ModuleName CollectionManagerCommon
        $r = @(Get-CMCollectionFolderTree -SMSProvider 'cm01' -SiteCode 'MCM')
        $r.Count | Should -Be 1
        $r[0].FolderID | Should -Be 7
        $r[0].Name | Should -Be 'Servers'
        $r[0].ParentID | Should -Be 0
        $r[0].FolderID | Should -BeOfType [int]
    }

    It 'asks the provider for device collection folders only' {
        Mock Get-CimInstance { } -ModuleName CollectionManagerCommon
        $null = Get-CMCollectionFolderTree -SMSProvider 'cm01.contoso.com' -SiteCode 'MCM'
        Should -Invoke Get-CimInstance -ModuleName CollectionManagerCommon -Times 1 -Exactly -ParameterFilter {
            $ComputerName -eq 'cm01.contoso.com' -and $Namespace -eq 'root\sms\site_MCM' -and $ClassName -eq 'SMS_ObjectContainerNode' -and $Filter -eq 'ObjectType = 5000'
        }
    }

    It 'returns an empty array when there are no folders' {
        Mock Get-CimInstance { } -ModuleName CollectionManagerCommon
        $r = Get-CMCollectionFolderTree -SMSProvider 'cm01' -SiteCode 'MCM'
        , $r | Should -BeOfType [object[]]
        @($r).Count | Should -Be 0
    }

    It 'keeps the unicode and quote characters of a folder name' {
        $global:Fx = "Ünï 'quoted' <b>"
        Mock Get-CimInstance { [PSCustomObject]@{ ContainerNodeID = 1; Name = $global:Fx; ParentContainerNodeID = 0 } } -ModuleName CollectionManagerCommon
        (@(Get-CMCollectionFolderTree -SMSProvider 'cm01' -SiteCode 'MCM'))[0].Name | Should -Be $global:Fx
    }

    It 'lets a provider failure reach the caller' {
        Mock Get-CimInstance { throw 'RPC server unavailable' } -ModuleName CollectionManagerCommon
        { Get-CMCollectionFolderTree -SMSProvider 'cm01' -SiteCode 'MCM' } | Should -Throw '*RPC server unavailable*'
    }

    It 'requires <Missing>' -ForEach @(
        @{ Missing = 'a provider'; Params = @{ SMSProvider = ''; SiteCode = 'MCM' } }
        @{ Missing = 'a site code'; Params = @{ SMSProvider = 'cm01'; SiteCode = '' } }
    ) {
        { Get-CMCollectionFolderTree @Params } | Should -Throw
    }
}

Describe 'Get-CMCollectionFolderMap' {
    It 'maps each collection id to its folder' {
        Mock Get-CimInstance {
            [PSCustomObject]@{ InstanceKey = 'MCM00001'; ContainerNodeID = 3 }
            [PSCustomObject]@{ InstanceKey = 'MCM00002'; ContainerNodeID = '4' }
        } -ModuleName CollectionManagerCommon
        $m = Get-CMCollectionFolderMap -SMSProvider 'cm01' -SiteCode 'MCM'
        $m['MCM00001'] | Should -Be 3
        $m['MCM00002'] | Should -Be 4
        $m.Count | Should -Be 2
    }

    It 'gives an empty map when no collection sits in a folder' {
        Mock Get-CimInstance { } -ModuleName CollectionManagerCommon
        (Get-CMCollectionFolderMap -SMSProvider 'cm01' -SiteCode 'MCM').Count | Should -Be 0
    }

    It 'asks for the container items of device collections' {
        Mock Get-CimInstance { } -ModuleName CollectionManagerCommon
        $null = Get-CMCollectionFolderMap -SMSProvider 'cm01' -SiteCode 'MCM'
        Should -Invoke Get-CimInstance -ModuleName CollectionManagerCommon -ParameterFilter { $ClassName -eq 'SMS_ObjectContainerItem' -and $Filter -eq 'ObjectType = 5000' }
    }

    It 'lets a provider failure reach the caller' {
        Mock Get-CimInstance { throw 'access denied' } -ModuleName CollectionManagerCommon
        { Get-CMCollectionFolderMap -SMSProvider 'cm01' -SiteCode 'MCM' } | Should -Throw '*access denied*'
    }
}

Describe 'Get-AllDeviceCollections' {
    It 'names refresh type <Type> as <Name>' -ForEach @(
        @{ Type = 1; Name = 'Manual' }
        @{ Type = 2; Name = 'Periodic' }
        @{ Type = 4; Name = 'Continuous' }
        @{ Type = 6; Name = 'Both' }
        @{ Type = 0; Name = 'Unknown (0)' }
        @{ Type = 3; Name = 'Unknown (3)' }
        @{ Type = 99; Name = 'Unknown (99)' }
    ) {
        $global:Fx = $Type
        Mock Get-CMDeviceCollection { [PSCustomObject]@{ CollectionID = 'MCM00001'; Name = 'A'; MemberCount = 5; RefreshType = $global:Fx; LimitToCollectionName = 'All'; LimitToCollectionID = 'SMS00001'; Comment = ''; CollectionRules = @() } } -ModuleName CollectionManagerCommon
        (@(Get-AllDeviceCollections))[0].RefreshType | Should -Be $Name
    }

    It 'flags built-in collections by id <Id>' -ForEach @(
        @{ Id = 'SMS00001'; Builtin = $true }
        @{ Id = 'SMSDM003'; Builtin = $true }
        @{ Id = 'sms00001'; Builtin = $true }
        @{ Id = 'MCM00001'; Builtin = $false }
        @{ Id = 'ZZ1SMS00'; Builtin = $false }
    ) {
        $global:Fx = $Id
        Mock Get-CMDeviceCollection { [PSCustomObject]@{ CollectionID = $global:Fx; Name = 'A'; MemberCount = 0; RefreshType = 1 } } -ModuleName CollectionManagerCommon
        (@(Get-AllDeviceCollections))[0].IsBuiltIn | Should -Be $Builtin
    }

    It 'turns a null member count into 0' {
        Mock Get-CMDeviceCollection { [PSCustomObject]@{ CollectionID = 'MCM00001'; Name = 'A'; MemberCount = $null; RefreshType = 1 } } -ModuleName CollectionManagerCommon
        (@(Get-AllDeviceCollections))[0].MemberCount | Should -Be 0
    }

    It 'returns nothing for a site without device collections' {
        Mock Get-CMDeviceCollection { } -ModuleName CollectionManagerCommon
        @(Get-AllDeviceCollections).Count | Should -Be 0
    }

    It 'handles 3000 collections' {
        Mock Get-CMDeviceCollection { 1..3000 | ForEach-Object { [PSCustomObject]@{ CollectionID = "MCM$_"; Name = "C$_"; MemberCount = $_; RefreshType = 2 } } } -ModuleName CollectionManagerCommon
        @(Get-AllDeviceCollections).Count | Should -Be 3000
    }

    It 'lets a cmdlet failure reach the caller' {
        Mock Get-CMDeviceCollection { throw 'not connected' } -ModuleName CollectionManagerCommon
        { Get-AllDeviceCollections } | Should -Throw '*not connected*'
    }
}

Describe 'Get-CollectionDetail' {
    It 'returns null when the collection does not exist' {
        Mock Get-CMDeviceCollection { $null } -ModuleName CollectionManagerCommon
        Get-CollectionDetail -CollectionId 'MCM99999' | Should -BeNullOrEmpty
    }

    It 'counts query rules and direct rules' {
        Mock Get-CMDeviceCollection {
            [PSCustomObject]@{
                CollectionID = 'MCM00001'; Name = 'A'; MemberCount = 3; LimitToCollectionName = 'All'; Comment = 'c'; RefreshType = 6
                CollectionRules = @(
                    [PSCustomObject]@{ SmsProviderObjectPath = 'SMS_CollectionRuleDirect.RuleName="x"' },
                    [PSCustomObject]@{ SmsProviderObjectPath = 'SMS_CollectionRuleDirect.RuleName="y"' },
                    [PSCustomObject]@{ SmsProviderObjectPath = 'SMS_CollectionRuleQuery.RuleName="q"' },
                    $null)
            }
        } -ModuleName CollectionManagerCommon
        Mock Get-CMDeviceCollectionQueryMembershipRule { [PSCustomObject]@{ RuleName = 'q'; QueryExpression = 'select 1' } } -ModuleName CollectionManagerCommon
        $d = Get-CollectionDetail -CollectionId 'MCM00001'
        $d.DirectRuleCount | Should -Be 2
        $d.QueryRuleCount | Should -Be 1
        $d.RefreshType | Should -Be 6
        $d.Name | Should -Be 'A'
        $d.IsBuiltIn | Should -BeFalse
    }

    It 'falls back to the type name when a rule has no object path' {
        Mock Get-CMDeviceCollection {
            [PSCustomObject]@{ CollectionID = 'MCM00001'; Name = 'A'; MemberCount = 0; RefreshType = 1; CollectionRules = @([PSCustomObject]@{ SmsProviderObjectPath = $null }) }
        } -ModuleName CollectionManagerCommon
        Mock Get-CMDeviceCollectionQueryMembershipRule { } -ModuleName CollectionManagerCommon
        (Get-CollectionDetail -CollectionId 'MCM00001').DirectRuleCount | Should -Be 0
    }

    It 'reports zero rules for a collection without rules' {
        Mock Get-CMDeviceCollection { [PSCustomObject]@{ CollectionID = 'SMS00001'; Name = 'All Systems'; MemberCount = 9; RefreshType = 4 } } -ModuleName CollectionManagerCommon
        Mock Get-CMDeviceCollectionQueryMembershipRule { } -ModuleName CollectionManagerCommon
        $d = Get-CollectionDetail -CollectionId 'SMS00001'
        $d.QueryRuleCount | Should -Be 0
        $d.DirectRuleCount | Should -Be 0
        $d.IsBuiltIn | Should -BeTrue
    }

    It 'lets a collection lookup failure reach the caller' {
        Mock Get-CMDeviceCollection { throw 'no such collection' } -ModuleName CollectionManagerCommon
        { Get-CollectionDetail -CollectionId 'MCM00001' } | Should -Throw '*no such collection*'
    }
}

Describe 'Get-CollectionQueryRules and Get-CollectionMembers' {
    It 'returns the rule name and expression' {
        Mock Get-CMDeviceCollectionQueryMembershipRule { [PSCustomObject]@{ RuleName = 'R1'; QueryExpression = 'select * from SMS_R_System'; Extra = 1 } } -ModuleName CollectionManagerCommon
        $r = @(Get-CollectionQueryRules -CollectionId 'MCM00001')
        $r[0].RuleName | Should -Be 'R1'
        $r[0].QueryExpression | Should -Be 'select * from SMS_R_System'
        $r[0].PSObject.Properties.Name | Should -Not -Contain 'Extra'
    }

    It 'returns no rules for a collection without query rules' {
        Mock Get-CMDeviceCollectionQueryMembershipRule { } -ModuleName CollectionManagerCommon
        @(Get-CollectionQueryRules -CollectionId 'MCM00001').Count | Should -Be 0
    }

    It 'lets a rule query failure reach the caller' {
        Mock Get-CMDeviceCollectionQueryMembershipRule { throw 'boom' } -ModuleName CollectionManagerCommon
        { Get-CollectionQueryRules -CollectionId 'MCM00001' } | Should -Throw '*boom*'
    }

    It 'returns the member columns' {
        Mock Get-CMCollectionMember { [PSCustomObject]@{ Name = 'PC1'; ResourceID = 16777220; Domain = 'CONTOSO'; IsClient = $true; IsActive = $false; Other = 'x' } } -ModuleName CollectionManagerCommon
        $m = @(Get-CollectionMembers -CollectionId 'MCM00001')
        $m[0].Name | Should -Be 'PC1'
        $m[0].ResourceID | Should -Be 16777220
        $m[0].IsClient | Should -BeTrue
        $m[0].IsActive | Should -BeFalse
    }

    It 'returns no members for an empty collection' {
        Mock Get-CMCollectionMember { } -ModuleName CollectionManagerCommon
        @(Get-CollectionMembers -CollectionId 'MCM00001').Count | Should -Be 0
    }

    It 'returns 5000 members' {
        Mock Get-CMCollectionMember { 1..5000 | ForEach-Object { [PSCustomObject]@{ Name = "PC$_"; ResourceID = $_; Domain = 'D'; IsClient = $true; IsActive = $true } } } -ModuleName CollectionManagerCommon
        @(Get-CollectionMembers -CollectionId 'MCM00001').Count | Should -Be 5000
    }

    It 'lets a member query failure reach the caller' {
        Mock Get-CMCollectionMember { throw 'timeout' } -ModuleName CollectionManagerCommon
        { Get-CollectionMembers -CollectionId 'MCM00001' } | Should -Throw '*timeout*'
    }
}

Describe 'ConvertTo-SiteSavedQueryInfo' {
    It 'turns null fields into empty strings' {
        $i = InModuleScope CollectionManagerCommon { ConvertTo-SiteSavedQueryInfo -Query ([PSCustomObject]@{ QueryID = $null; Name = $null; Expression = $null; Comments = $null; TargetClassName = $null; LimitToCollectionID = $null }) }
        foreach ($p in 'QueryID', 'Name', 'Expression', 'Comment', 'TargetClassName', 'LimitToCollectionID') { $i.$p | Should -Be '' }
    }

    It 'maps Comments to Comment' {
        (InModuleScope CollectionManagerCommon { ConvertTo-SiteSavedQueryInfo -Query ([PSCustomObject]@{ QueryID = 'Q1'; Name = 'n'; Expression = 'e'; Comments = 'note'; TargetClassName = 'SMS_R_System'; LimitToCollectionID = 'SMS00001' }) }).Comment | Should -Be 'note'
    }
}

Describe 'Get-SiteSavedQueries' {
    It 'sorts by name, then by id' {
        Mock Get-CMQuery {
            [PSCustomObject]@{ QueryID = 'MCM00003'; Name = 'B'; Expression = ''; Comments = '' }
            [PSCustomObject]@{ QueryID = 'MCM00002'; Name = 'A'; Expression = ''; Comments = '' }
            [PSCustomObject]@{ QueryID = 'MCM00001'; Name = 'A'; Expression = ''; Comments = '' }
        } -ModuleName CollectionManagerCommon
        (@(Get-SiteSavedQueries) | ForEach-Object QueryID) -join ',' | Should -Be 'MCM00001,MCM00002,MCM00003'
    }

    It 'returns nothing when the site has no saved query' {
        Mock Get-CMQuery { } -ModuleName CollectionManagerCommon
        @(Get-SiteSavedQueries).Count | Should -Be 0
    }
}

Describe 'New-SiteSavedQuery' {
    BeforeEach {
        Mock New-CMQuery { [PSCustomObject]@{ QueryID = 'MCM00009'; Name = $Name; Expression = $Expression; Comments = $Comment; TargetClassName = $TargetClassName; LimitToCollectionID = $LimitToCollectionId } } -ModuleName CollectionManagerCommon
    }

    It 'leaves out the optional parameters that are blank: <Case>' -ForEach @(
        @{ Case = 'empty'; Cmt = ''; Tgt = ''; Lim = '' }
        @{ Case = 'white space'; Cmt = '   '; Tgt = ' '; Lim = "`t" }
    ) {
        $null = New-SiteSavedQuery -Name 'Q' -Expression 'select 1' -Comment $Cmt -TargetClassName $Tgt -LimitToCollectionId $Lim
        Should -Invoke New-CMQuery -ModuleName CollectionManagerCommon -ParameterFilter { -not $Comment -and -not $TargetClassName -and -not $LimitToCollectionId -and $Name -eq 'Q' }
    }

    It 'passes the optional parameters that are set' {
        $r = New-SiteSavedQuery -Name 'Q' -Expression 'select 1' -Comment 'c' -TargetClassName 'SMS_R_System' -LimitToCollectionId 'SMS00001'
        $r.Comment | Should -Be 'c'
        $r.TargetClassName | Should -Be 'SMS_R_System'
        $r.LimitToCollectionID | Should -Be 'SMS00001'
    }

    It 'looks the query up by exact name when the cmdlet returns nothing' {
        Mock New-CMQuery { } -ModuleName CollectionManagerCommon
        Mock Get-CMQuery {
            [PSCustomObject]@{ QueryID = 'X1'; Name = 'Q2'; Expression = '' }
            [PSCustomObject]@{ QueryID = 'X2'; Name = 'Q'; Expression = 'select 1' }
        } -ModuleName CollectionManagerCommon
        (New-SiteSavedQuery -Name 'Q' -Expression 'select 1').QueryID | Should -Be 'X2'
    }

    It 'throws when the site returns no saved query' {
        Mock New-CMQuery { } -ModuleName CollectionManagerCommon
        Mock Get-CMQuery { } -ModuleName CollectionManagerCommon
        { New-SiteSavedQuery -Name 'Q' -Expression 'select 1' } | Should -Throw "*did not return the saved query 'Q'*"
    }

    It 'lets a create failure reach the caller' {
        Mock New-CMQuery { throw 'duplicate name' } -ModuleName CollectionManagerCommon
        { New-SiteSavedQuery -Name 'Q' -Expression 'select 1' } | Should -Throw '*duplicate name*'
    }

    It 'rejects an empty name or expression' -ForEach @(
        @{ P = @{ Name = ''; Expression = 'x' } }
        @{ P = @{ Name = 'x'; Expression = '' } }
    ) {
        { New-SiteSavedQuery @P } | Should -Throw
    }

    It 'does nothing under -WhatIf' {
        $r = New-SiteSavedQuery -Name 'Q' -Expression 'select 1' -WhatIf
        $r | Should -BeNullOrEmpty
        Should -Invoke New-CMQuery -ModuleName CollectionManagerCommon -Times 0 -Exactly
    }
}

Describe 'Set-SiteSavedQuery and Remove-SiteSavedQuery' {
    It 'updates by id and returns the query' {
        Mock Set-CMQuery { [PSCustomObject]@{ QueryID = $Id; Name = $NewName; Expression = $Expression; Comments = $Comment } } -ModuleName CollectionManagerCommon
        $r = Set-SiteSavedQuery -QueryId 'MCM00001' -Name 'New' -Expression 'select 2' -Comment 'c'
        $r.QueryID | Should -Be 'MCM00001'
        $r.Name | Should -Be 'New'
    }

    It 'reads the query back when the update returns nothing' {
        Mock Set-CMQuery { } -ModuleName CollectionManagerCommon
        Mock Get-CMQuery { [PSCustomObject]@{ QueryID = 'MCM00001'; Name = 'Back'; Expression = '' } } -ModuleName CollectionManagerCommon
        (Set-SiteSavedQuery -QueryId 'MCM00001' -Name 'x' -Expression 'y').Name | Should -Be 'Back'
    }

    It 'throws when the update and the read both return nothing' {
        Mock Set-CMQuery { } -ModuleName CollectionManagerCommon
        Mock Get-CMQuery { } -ModuleName CollectionManagerCommon
        { Set-SiteSavedQuery -QueryId 'MCM00001' -Name 'x' -Expression 'y' } | Should -Throw "*did not return the updated query 'MCM00001'*"
    }

    It 'does nothing under -WhatIf' {
        Mock Set-CMQuery { } -ModuleName CollectionManagerCommon
        $null = Set-SiteSavedQuery -QueryId 'MCM00001' -Name 'x' -Expression 'y' -WhatIf
        Should -Invoke Set-CMQuery -ModuleName CollectionManagerCommon -Times 0 -Exactly
    }

    It 'removes by id and returns true' {
        Mock Remove-CMQuery { } -ModuleName CollectionManagerCommon
        Remove-SiteSavedQuery -QueryId 'MCM00001' | Should -BeTrue
        Should -Invoke Remove-CMQuery -ModuleName CollectionManagerCommon -ParameterFilter { $Id -eq 'MCM00001' -and $Force }
    }

    It 'lets a remove failure reach the caller' {
        Mock Remove-CMQuery { throw 'in use' } -ModuleName CollectionManagerCommon
        { Remove-SiteSavedQuery -QueryId 'MCM00001' } | Should -Throw '*in use*'
    }

    It 'removes nothing under -WhatIf' {
        Mock Remove-CMQuery { } -ModuleName CollectionManagerCommon
        Remove-SiteSavedQuery -QueryId 'MCM00001' -WhatIf | Should -BeFalse
        Should -Invoke Remove-CMQuery -ModuleName CollectionManagerCommon -Times 0 -Exactly
    }
}
