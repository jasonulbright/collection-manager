#Requires -Modules Pester
# Write actions: collections, direct members, rules, evaluation.
# ConfigMgr cmdlets are stubs and mocks. Covers success, failure, partial failure,
# -WhatIf, built-in protection and hostile input.

BeforeAll {
    . "$PSScriptRoot\TestHelpers.ps1"
    Import-AppModule
    $script:Stubs = New-CmStub -Name Get-CMDeviceCollection, New-CMDeviceCollection, Copy-CMCollection, Remove-CMDeviceCollection, Set-CMCollection, New-CMSchedule,
        Get-CMDevice, Add-CMDeviceCollectionDirectMembershipRule, Remove-CMDeviceCollectionDirectMembershipRule,
        Add-CMDeviceCollectionQueryMembershipRule, Remove-CMDeviceCollectionQueryMembershipRule, Get-CMDeviceCollectionQueryMembershipRule,
        Add-CMDeviceCollectionIncludeMembershipRule, Add-CMDeviceCollectionExcludeMembershipRule, Invoke-CMCollectionUpdate
    $global:TestLogs = [System.Collections.ArrayList]::new()
    Mock Write-Log { [void]$global:TestLogs.Add(@{ Message = $Message; Level = $Level }) } -ModuleName CollectionManagerCommon
}

AfterAll {
    Remove-CmStub -Name $script:Stubs
    Remove-Variable -Name Fx, TestLogs -Scope Global -ErrorAction SilentlyContinue
}

Describe 'New-ManagedCollection' {
    BeforeEach {
        Mock New-CMSchedule { 'schedule' } -ModuleName CollectionManagerCommon
        Mock New-CMDeviceCollection { [PSCustomObject]@{ CollectionID = 'MCM00042'; Name = $Name } } -ModuleName CollectionManagerCommon
    }

    It 'maps <Type> to refresh value <Value> and <Sched> a schedule' -ForEach @(
        @{ Type = 'Manual';     Value = 1; Sched = 'no';   Times = 0 }
        @{ Type = 'Periodic';   Value = 2; Sched = 'adds'; Times = 1 }
        @{ Type = 'Continuous'; Value = 4; Sched = 'no';   Times = 0 }
        @{ Type = 'Both';       Value = 6; Sched = 'adds'; Times = 1 }
    ) {
        $null = New-ManagedCollection -Name 'C' -LimitingCollectionName 'All Systems' -RefreshType $Type
        Should -Invoke New-CMDeviceCollection -ModuleName CollectionManagerCommon -Times 1 -Exactly -ParameterFilter { $RefreshType -eq $Value -and $LimitingCollectionName -eq 'All Systems' }
        Should -Invoke New-CMSchedule -ModuleName CollectionManagerCommon -Times $Times -Exactly
    }

    It 'uses Both when no refresh type is given' {
        $null = New-ManagedCollection -Name 'C' -LimitingCollectionName 'All Systems'
        Should -Invoke New-CMDeviceCollection -ModuleName CollectionManagerCommon -ParameterFilter { $RefreshType -eq 6 }
    }

    It 'passes the comment only when it is not empty' -ForEach @(
        @{ Comment = 'hello'; Passed = $true }
        @{ Comment = '';      Passed = $false }
    ) {
        $null = New-ManagedCollection -Name 'C' -LimitingCollectionName 'L' -Comment $Comment
        Should -Invoke New-CMDeviceCollection -ModuleName CollectionManagerCommon -ParameterFilter { [bool]$Comment -eq $Passed }
    }

    It 'returns the new collection' {
        (New-ManagedCollection -Name 'C' -LimitingCollectionName 'L').CollectionID | Should -Be 'MCM00042'
    }

    It 'passes the hostile name <Name> through unchanged' -ForEach @(
        @{ Name = "O'Brien ""Q"" <b>" }
        @{ Name = ('n' * 300) }
        @{ Name = 'Ünï 日本語' }
        @{ Name = '..\..\x' }
    ) {
        $global:Fx = $Name
        $null = New-ManagedCollection -Name $Name -LimitingCollectionName 'L'
        Should -Invoke New-CMDeviceCollection -ModuleName CollectionManagerCommon -ParameterFilter { $Name -eq $global:Fx }
    }

    It 'rejects <Case>' -ForEach @(
        @{ Case = 'an empty name'; P = @{ Name = ''; LimitingCollectionName = 'L' } }
        @{ Case = 'an empty limiting collection'; P = @{ Name = 'C'; LimitingCollectionName = '' } }
        @{ Case = 'an unknown refresh type'; P = @{ Name = 'C'; LimitingCollectionName = 'L'; RefreshType = 'Never' } }
    ) {
        { New-ManagedCollection @P } | Should -Throw
        Should -Invoke New-CMDeviceCollection -ModuleName CollectionManagerCommon -Times 0 -Exactly
    }

    It 'lets a create failure reach the caller' {
        Mock New-CMDeviceCollection { throw 'name already exists' } -ModuleName CollectionManagerCommon
        { New-ManagedCollection -Name 'C' -LimitingCollectionName 'L' } | Should -Throw '*name already exists*'
    }

    It 'creates nothing under -WhatIf' {
        $r = New-ManagedCollection -Name 'C' -LimitingCollectionName 'L' -WhatIf
        $r | Should -BeNullOrEmpty
        Should -Invoke New-CMDeviceCollection -ModuleName CollectionManagerCommon -Times 0 -Exactly
    }
}

Describe 'Copy-ManagedCollection' {
    It 'clones by id and returns the clone that it reads back by name' {
        # The cmdlet returns the source collection with -PassThru, as it does on the lab site.
        Mock Copy-CMCollection { [PSCustomObject]@{ CollectionID = 'MCM00001'; Name = 'Source' } } -ModuleName CollectionManagerCommon
        Mock Get-CMDeviceCollection {
            [PSCustomObject]@{ CollectionID = 'MCM00050'; Name = 'Copy 2' }
            [PSCustomObject]@{ CollectionID = 'MCM00051'; Name = 'Copy' }
        } -ModuleName CollectionManagerCommon
        (Copy-ManagedCollection -SourceCollectionId 'MCM00001' -NewName 'Copy').CollectionID | Should -Be 'MCM00051'
        Should -Invoke Copy-CMCollection -ModuleName CollectionManagerCommon -ParameterFilter { $Id -eq 'MCM00001' -and $NewName -eq 'Copy' -and $PassThru }
    }

    It 'throws when the clone is not found by name' {
        Mock Copy-CMCollection { } -ModuleName CollectionManagerCommon
        Mock Get-CMDeviceCollection { } -ModuleName CollectionManagerCommon
        { Copy-ManagedCollection -SourceCollectionId 'MCM00001' -NewName 'Copy' } | Should -Throw "*did not return the cloned collection 'Copy'*"
    }

    It 'lets a clone failure reach the caller' {
        Mock Copy-CMCollection { throw 'source missing' } -ModuleName CollectionManagerCommon
        { Copy-ManagedCollection -SourceCollectionId 'MCM00001' -NewName 'x' } | Should -Throw '*source missing*'
    }

    It 'clones nothing under -WhatIf' {
        Mock Copy-CMCollection { } -ModuleName CollectionManagerCommon
        $null = Copy-ManagedCollection -SourceCollectionId 'MCM00001' -NewName 'x' -WhatIf
        Should -Invoke Copy-CMCollection -ModuleName CollectionManagerCommon -Times 0 -Exactly
    }

    It 'rejects an empty <Field>' -ForEach @(
        @{ Field = 'source'; P = @{ SourceCollectionId = ''; NewName = 'x' } }
        @{ Field = 'new name'; P = @{ SourceCollectionId = 'MCM00001'; NewName = '' } }
    ) {
        { Copy-ManagedCollection @P } | Should -Throw
    }
}
Describe 'Remove-ManagedCollection' {
    BeforeEach {
        $global:TestLogs.Clear()
        Mock Remove-CMDeviceCollection { } -ModuleName CollectionManagerCommon
    }

    It 'blocks the built-in collection "<Id>"' -ForEach @(
        @{ Id = 'SMS00001' }
        @{ Id = 'SMSDM003' }
        @{ Id = 'sms00001' }
        @{ Id = ' SMS00001' }
        @{ Id = "SMS00001`t" }
    ) {
        Remove-ManagedCollection -CollectionId $Id | Should -BeFalse
        Should -Invoke Remove-CMDeviceCollection -ModuleName CollectionManagerCommon -Times 0 -Exactly
        @($global:TestLogs | Where-Object { $_.Level -eq 'WARN' -and $_.Message -like 'BLOCKED*' }).Count | Should -Be 1
    }

    It 'removes a custom collection and returns true' {
        Remove-ManagedCollection -CollectionId 'MCM00001' | Should -BeTrue
        Should -Invoke Remove-CMDeviceCollection -ModuleName CollectionManagerCommon -Times 1 -Exactly -ParameterFilter { $CollectionId -eq 'MCM00001' -and $Force }
    }

    It 'lets a remove failure reach the caller' {
        Mock Remove-CMDeviceCollection { throw 'collection is in use by a deployment' } -ModuleName CollectionManagerCommon
        { Remove-ManagedCollection -CollectionId 'MCM00001' } | Should -Throw '*in use by a deployment*'
    }

    It 'removes nothing under -WhatIf' {
        Remove-ManagedCollection -CollectionId 'MCM00001' -WhatIf | Should -BeFalse
        Should -Invoke Remove-CMDeviceCollection -ModuleName CollectionManagerCommon -Times 0 -Exactly
    }

    It 'rejects an empty id' {
        { Remove-ManagedCollection -CollectionId '' } | Should -Throw
    }
}

Describe 'Set-CollectionProperties' {
    BeforeEach { Mock Set-CMCollection { } -ModuleName CollectionManagerCommon }

    It 'passes the new name and the comment' {
        Set-CollectionProperties -CollectionId 'MCM00001' -NewName 'N' -Comment 'C'
        Should -Invoke Set-CMCollection -ModuleName CollectionManagerCommon -Times 1 -Exactly -ParameterFilter { $CollectionId -eq 'MCM00001' -and $NewName -eq 'N' -and $Comment -eq 'C' }
    }

    It 'clears the comment when an empty comment is given' {
        Set-CollectionProperties -CollectionId 'MCM00001' -Comment ''
        Should -Invoke Set-CMCollection -ModuleName CollectionManagerCommon -ParameterFilter { $null -ne $Comment -and $Comment -eq '' }
    }

    It 'leaves the comment alone when it is not given' {
        Set-CollectionProperties -CollectionId 'MCM00001' -NewName 'N'
        Should -Invoke Set-CMCollection -ModuleName CollectionManagerCommon -ParameterFilter { $null -eq $Comment }
    }

    It 'applies refresh type <Value>' -ForEach @(
        @{ Value = 1 }, @{ Value = 2 }, @{ Value = 4 }, @{ Value = 6 }
    ) {
        Set-CollectionProperties -CollectionId 'MCM00001' -RefreshType $Value
        Should -Invoke Set-CMCollection -ModuleName CollectionManagerCommon -Times 1 -Exactly -ParameterFilter { $RefreshType -eq $Value }
    }

    It 'passes a refresh schedule for refresh type <Value> only when it is periodic: <Sched>' -ForEach @(@{ Value = 1; Sched = $false }, @{ Value = 2; Sched = $true }, @{ Value = 4; Sched = $false }, @{ Value = 6; Sched = $true }) {
        Mock New-CMSchedule { 'schedule' } -ModuleName CollectionManagerCommon
        Set-CollectionProperties -CollectionId 'MCM00001' -RefreshType $Value
        Should -Invoke Set-CMCollection -ModuleName CollectionManagerCommon -Times 1 -Exactly -ParameterFilter { [bool]$RefreshSchedule -eq $Sched }
    }

    It 'rejects the refresh type <Value> before it calls the site' -ForEach @(
        @{ Value = 0 }, @{ Value = 3 }, @{ Value = 5 }, @{ Value = 7 }, @{ Value = -1 }
    ) {
        { Set-CollectionProperties -CollectionId 'MCM00001' -RefreshType $Value } | Should -Throw '*refresh type*'
        Should -Invoke Set-CMCollection -ModuleName CollectionManagerCommon -Times 0 -Exactly
    }

    It 'lets an update failure reach the caller' {
        Mock Set-CMCollection { throw 'rename failed' } -ModuleName CollectionManagerCommon
        { Set-CollectionProperties -CollectionId 'MCM00001' -NewName 'x' } | Should -Throw '*rename failed*'
    }

    It 'changes nothing under -WhatIf' {
        Set-CollectionProperties -CollectionId 'MCM00001' -NewName 'x' -WhatIf
        Should -Invoke Set-CMCollection -ModuleName CollectionManagerCommon -Times 0 -Exactly
    }
}

Describe 'Add-DirectMember' {
    BeforeEach {
        $global:TestLogs.Clear()
        $global:Fx = @([PSCustomObject]@{ Name = 'PC1'; ResourceID = 16777220 })
        Mock Get-CMDevice { $global:Fx } -ModuleName CollectionManagerCommon
        Mock Add-CMDeviceCollectionDirectMembershipRule { } -ModuleName CollectionManagerCommon
    }

    It 'resolves the device and adds its resource id' {
        Add-DirectMember -CollectionId 'MCM00001' -DeviceName 'PC1' | Should -BeTrue
        Should -Invoke Add-CMDeviceCollectionDirectMembershipRule -ModuleName CollectionManagerCommon -Times 1 -Exactly -ParameterFilter { $CollectionId -eq 'MCM00001' -and $ResourceId -eq 16777220 }
    }

    It 'returns false and logs an error when the device is not found' {
        $global:Fx = @()
        Add-DirectMember -CollectionId 'MCM00001' -DeviceName 'NOPE' | Should -BeFalse
        Should -Invoke Add-CMDeviceCollectionDirectMembershipRule -ModuleName CollectionManagerCommon -Times 0 -Exactly
        @($global:TestLogs | Where-Object Level -eq 'ERROR').Count | Should -Be 1
    }

    It 'turns wildcard handling off for the device name' {
        $null = Add-DirectMember -CollectionId 'MCM00001' -DeviceName 'PC1'
        Should -Invoke Get-CMDevice -ModuleName CollectionManagerCommon -ParameterFilter { $DisableWildcardHandling }
    }

    It 'adds only the exact match when the lookup returns similar names' {
        $global:Fx = @([PSCustomObject]@{ Name = 'PC10'; ResourceID = 10 }, [PSCustomObject]@{ Name = 'pc1'; ResourceID = 1 }, [PSCustomObject]@{ Name = 'PC11'; ResourceID = 11 })
        Add-DirectMember -CollectionId 'MCM00001' -DeviceName 'PC1' | Should -BeTrue
        Should -Invoke Add-CMDeviceCollectionDirectMembershipRule -ModuleName CollectionManagerCommon -Times 1 -Exactly -ParameterFilter { $ResourceId -eq 1 }
    }

    It 'adds nothing when a wildcard name matches several devices' -ForEach @(
        @{ Name = '*' }, @{ Name = 'PC*' }, @{ Name = 'PC?' }
    ) {
        $global:Fx = @([PSCustomObject]@{ Name = 'PC1'; ResourceID = 1 }, [PSCustomObject]@{ Name = 'PC2'; ResourceID = 2 })
        Add-DirectMember -CollectionId 'MCM00001' -DeviceName $Name | Should -BeFalse
        Should -Invoke Add-CMDeviceCollectionDirectMembershipRule -ModuleName CollectionManagerCommon -Times 0 -Exactly
    }

    It 'adds nothing when two devices share the exact name' {
        $global:Fx = @([PSCustomObject]@{ Name = 'PC1'; ResourceID = 1 }, [PSCustomObject]@{ Name = 'PC1'; ResourceID = 2 })
        Add-DirectMember -CollectionId 'MCM00001' -DeviceName 'PC1' | Should -BeFalse
        Should -Invoke Add-CMDeviceCollectionDirectMembershipRule -ModuleName CollectionManagerCommon -Times 0 -Exactly
        @($global:TestLogs | Where-Object { $_.Level -eq 'ERROR' -and $_.Message -like '*more than one*' }).Count | Should -Be 1
    }

    It 'lets a membership failure reach the caller' {
        Mock Add-CMDeviceCollectionDirectMembershipRule { throw 'rule exists' } -ModuleName CollectionManagerCommon
        { Add-DirectMember -CollectionId 'MCM00001' -DeviceName 'PC1' } | Should -Throw '*rule exists*'
    }

    It 'adds nothing under -WhatIf' {
        Add-DirectMember -CollectionId 'MCM00001' -DeviceName 'PC1' -WhatIf | Should -BeFalse
        Should -Invoke Add-CMDeviceCollectionDirectMembershipRule -ModuleName CollectionManagerCommon -Times 0 -Exactly
    }

    It 'rejects an empty <Field>' -ForEach @(
        @{ Field = 'collection id'; P = @{ CollectionId = ''; DeviceName = 'PC1' } }
        @{ Field = 'device name'; P = @{ CollectionId = 'MCM00001'; DeviceName = '' } }
    ) {
        { Add-DirectMember @P } | Should -Throw
    }
}

Describe 'Remove-DirectMember' {
    BeforeEach { Mock Remove-CMDeviceCollectionDirectMembershipRule { } -ModuleName CollectionManagerCommon }

    It 'removes the resource id with -Force' {
        Remove-DirectMember -CollectionId 'MCM00001' -ResourceId 16777220
        Should -Invoke Remove-CMDeviceCollectionDirectMembershipRule -ModuleName CollectionManagerCommon -Times 1 -Exactly -ParameterFilter { $CollectionId -eq 'MCM00001' -and $ResourceId -eq 16777220 -and $Force }
    }

    It 'lets a failure reach the caller' {
        Mock Remove-CMDeviceCollectionDirectMembershipRule { throw 'no such rule' } -ModuleName CollectionManagerCommon
        { Remove-DirectMember -CollectionId 'MCM00001' -ResourceId 1 } | Should -Throw '*no such rule*'
    }

    It 'removes nothing under -WhatIf' {
        Remove-DirectMember -CollectionId 'MCM00001' -ResourceId 1 -WhatIf
        Should -Invoke Remove-CMDeviceCollectionDirectMembershipRule -ModuleName CollectionManagerCommon -Times 0 -Exactly
    }

    It 'rejects a resource id that is not a number' {
        { Remove-DirectMember -CollectionId 'MCM00001' -ResourceId 'abc' } | Should -Throw
    }
}

Describe 'Add-QueryRule, Remove-QueryRule, Add-IncludeRule, Add-ExcludeRule, Invoke-CollectionEvaluation' {
    BeforeEach {
        foreach ($c in 'Add-CMDeviceCollectionQueryMembershipRule', 'Remove-CMDeviceCollectionQueryMembershipRule', 'Add-CMDeviceCollectionIncludeMembershipRule', 'Add-CMDeviceCollectionExcludeMembershipRule', 'Invoke-CMCollectionUpdate') {
            Mock -CommandName $c -ModuleName CollectionManagerCommon -MockWith { }
        }
    }

    It 'adds a query rule' {
        Add-QueryRule -CollectionId 'MCM00001' -RuleName 'R' -QueryExpression 'select 1'
        Should -Invoke Add-CMDeviceCollectionQueryMembershipRule -ModuleName CollectionManagerCommon -Times 1 -Exactly -ParameterFilter { $CollectionId -eq 'MCM00001' -and $RuleName -eq 'R' -and $QueryExpression -eq 'select 1' }
    }

    It 'passes a hostile query expression through as data: <Q>' -ForEach @(
        @{ Q = "select * from SMS_R_System where Name = 'a'' or ''1''=''1'" }
        @{ Q = ('x' * 20000) }
        @{ Q = "select * -- comment`r`nfrom x" }
    ) {
        $global:Fx = $Q
        Add-QueryRule -CollectionId 'MCM00001' -RuleName 'R' -QueryExpression $Q
        Should -Invoke Add-CMDeviceCollectionQueryMembershipRule -ModuleName CollectionManagerCommon -ParameterFilter { $QueryExpression -ceq $global:Fx }
    }

    It 'removes a query rule with -Force' {
        Remove-QueryRule -CollectionId 'MCM00001' -RuleName 'R'
        Should -Invoke Remove-CMDeviceCollectionQueryMembershipRule -ModuleName CollectionManagerCommon -ParameterFilter { $RuleName -eq 'R' -and $Force }
    }

    It 'adds an include rule' {
        Add-IncludeRule -CollectionId 'MCM00001' -IncludeCollectionId 'MCM00002'
        Should -Invoke Add-CMDeviceCollectionIncludeMembershipRule -ModuleName CollectionManagerCommon -ParameterFilter { $CollectionId -eq 'MCM00001' -and $IncludeCollectionId -eq 'MCM00002' }
    }

    It 'adds an exclude rule' {
        Add-ExcludeRule -CollectionId 'MCM00001' -ExcludeCollectionId 'MCM00002'
        Should -Invoke Add-CMDeviceCollectionExcludeMembershipRule -ModuleName CollectionManagerCommon -ParameterFilter { $CollectionId -eq 'MCM00001' -and $ExcludeCollectionId -eq 'MCM00002' }
    }

    It 'forces an evaluation' {
        Invoke-CollectionEvaluation -CollectionId 'MCM00001'
        Should -Invoke Invoke-CMCollectionUpdate -ModuleName CollectionManagerCommon -ParameterFilter { $CollectionId -eq 'MCM00001' }
    }

    It '<Fn> lets a failure reach the caller' -ForEach @(
        @{ Fn = 'Add-QueryRule';  Cmd = 'Add-CMDeviceCollectionQueryMembershipRule';   P = @{ CollectionId = 'C'; RuleName = 'R'; QueryExpression = 'q' } }
        @{ Fn = 'Remove-QueryRule'; Cmd = 'Remove-CMDeviceCollectionQueryMembershipRule'; P = @{ CollectionId = 'C'; RuleName = 'R' } }
        @{ Fn = 'Add-IncludeRule'; Cmd = 'Add-CMDeviceCollectionIncludeMembershipRule'; P = @{ CollectionId = 'C'; IncludeCollectionId = 'D' } }
        @{ Fn = 'Add-ExcludeRule'; Cmd = 'Add-CMDeviceCollectionExcludeMembershipRule'; P = @{ CollectionId = 'C'; ExcludeCollectionId = 'D' } }
        @{ Fn = 'Invoke-CollectionEvaluation'; Cmd = 'Invoke-CMCollectionUpdate'; P = @{ CollectionId = 'C' } }
    ) {
        Mock -CommandName $Cmd -ModuleName CollectionManagerCommon -MockWith { throw 'site said no' }
        { & $Fn @P } | Should -Throw '*site said no*'
    }

    It '<Fn> changes nothing under -WhatIf' -ForEach @(
        @{ Fn = 'Add-QueryRule';  Cmd = 'Add-CMDeviceCollectionQueryMembershipRule';   P = @{ CollectionId = 'C'; RuleName = 'R'; QueryExpression = 'q' } }
        @{ Fn = 'Remove-QueryRule'; Cmd = 'Remove-CMDeviceCollectionQueryMembershipRule'; P = @{ CollectionId = 'C'; RuleName = 'R' } }
        @{ Fn = 'Add-IncludeRule'; Cmd = 'Add-CMDeviceCollectionIncludeMembershipRule'; P = @{ CollectionId = 'C'; IncludeCollectionId = 'D' } }
        @{ Fn = 'Add-ExcludeRule'; Cmd = 'Add-CMDeviceCollectionExcludeMembershipRule'; P = @{ CollectionId = 'C'; ExcludeCollectionId = 'D' } }
        @{ Fn = 'Invoke-CollectionEvaluation'; Cmd = 'Invoke-CMCollectionUpdate'; P = @{ CollectionId = 'C' } }
    ) {
        & $Fn @P -WhatIf
        Should -Invoke -CommandName $Cmd -ModuleName CollectionManagerCommon -Times 0 -Exactly
    }

    It 'rejects an empty rule name' {
        { Add-QueryRule -CollectionId 'C' -RuleName '' -QueryExpression 'q' } | Should -Throw
    }
}

Describe 'Update-QueryRule' {
    BeforeEach {
        $global:TestLogs.Clear()
        $script:order = [System.Collections.ArrayList]::new()
        Mock Get-CMDeviceCollectionQueryMembershipRule { [PSCustomObject]@{ RuleName = 'Old'; QueryExpression = 'select old' } } -ModuleName CollectionManagerCommon
        Mock Remove-CMDeviceCollectionQueryMembershipRule { [void]$script:order.Add("remove:$RuleName") } -ModuleName CollectionManagerCommon
        Mock Add-CMDeviceCollectionQueryMembershipRule { [void]$script:order.Add("add:$RuleName=$QueryExpression") } -ModuleName CollectionManagerCommon
    }

    It 'removes the old rule, then adds the new one' {
        Update-QueryRule -CollectionId 'MCM00001' -OldRuleName 'Old' -NewRuleName 'New' -NewQueryExpression 'select new'
        $script:order -join ';' | Should -Be 'remove:Old;add:New=select new'
    }

    It 'restores the old rule and throws when the new rule is refused' {
        Mock Add-CMDeviceCollectionQueryMembershipRule {
            [void]$script:order.Add("add:$RuleName=$QueryExpression")
            if ($RuleName -eq 'New') { throw 'invalid WQL' }
        } -ModuleName CollectionManagerCommon
        { Update-QueryRule -CollectionId 'MCM00001' -OldRuleName 'Old' -NewRuleName 'New' -NewQueryExpression 'bad' } | Should -Throw '*invalid WQL*restored*'
        $script:order -join ';' | Should -Be 'remove:Old;add:New=bad;add:Old=select old'
    }

    It 'says when the old rule could not be restored' {
        Mock Add-CMDeviceCollectionQueryMembershipRule { throw 'site down' } -ModuleName CollectionManagerCommon
        { Update-QueryRule -CollectionId 'MCM00001' -OldRuleName 'Old' -NewRuleName 'New' -NewQueryExpression 'x' } | Should -Throw '*not restored*'
    }

    It 'changes nothing when the old rule does not exist' {
        Mock Get-CMDeviceCollectionQueryMembershipRule { } -ModuleName CollectionManagerCommon
        { Update-QueryRule -CollectionId 'MCM00001' -OldRuleName 'Gone' -NewRuleName 'New' -NewQueryExpression 'x' } | Should -Throw "*'Gone' not found*"
        $script:order.Count | Should -Be 0
    }

    It 'stops when the old rule cannot be removed and adds nothing' {
        Mock Remove-CMDeviceCollectionQueryMembershipRule { throw 'locked' } -ModuleName CollectionManagerCommon
        { Update-QueryRule -CollectionId 'MCM00001' -OldRuleName 'Old' -NewRuleName 'New' -NewQueryExpression 'x' } | Should -Throw '*locked*'
        Should -Invoke Add-CMDeviceCollectionQueryMembershipRule -ModuleName CollectionManagerCommon -Times 0 -Exactly
    }

    It 'changes nothing under -WhatIf' {
        Update-QueryRule -CollectionId 'MCM00001' -OldRuleName 'Old' -NewRuleName 'New' -NewQueryExpression 'x' -WhatIf
        $script:order.Count | Should -Be 0
    }
}
