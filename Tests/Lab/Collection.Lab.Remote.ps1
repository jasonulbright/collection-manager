# Runs ON the lab site server (Windows PowerShell 5.1). Started by Invoke-CollectionLab.ps1.
#
# Creates ZZTEST- device collections and one ZZTEST- saved query, runs the module's
# write actions against them (create, copy, properties, direct member, query rules
# including a refused update, include and exclude rules, evaluation, WhatIf, built-in
# protection, remove), and removes everything it created, also when a step fails. It
# compares every other collection with a snapshot taken before the first change.
# Only objects named ZZTEST-* change.
param(
    [Parameter(Mandatory)][string]$ModuleRoot,
    [Parameter(Mandatory)][string]$SiteCode,
    [Parameter(Mandatory)][string]$Provider
)

$ErrorActionPreference = 'Stop'
$results = [System.Collections.Generic.List[object]]::new()
function Add-Result { param([string]$Step, [bool]$Ok, [string]$Detail = '') $results.Add([PSCustomObject]@{ Step = $Step; Ok = $Ok; Detail = $Detail }) }

$tag = [guid]::NewGuid().ToString('N').Substring(0, 8)
$nameA = "ZZTEST-CM-$tag-A"
$nameB = "ZZTEST-CM-$tag-B"
$nameCopy = "ZZTEST-CM-$tag-Copy"
$queryName = "ZZTEST-CM-$tag-Query"
$snapshot = $null
$ns = "root\SMS\site_$SiteCode"

function Get-CollectionSnapshot {
    @(Get-CimInstance -Namespace $ns -ClassName SMS_Collection -Filter 'CollectionType = 2' | Where-Object { $_.Name -notlike 'ZZTEST-*' } |
            ForEach-Object { '{0}|{1}|{2}|{3}' -f $_.CollectionID, $_.Name, $_.LimitToCollectionID, $_.RefreshType } | Sort-Object)
}

try {
    Import-Module (Join-Path $ModuleRoot 'Module\CollectionManagerCommon.psd1') -Force -DisableNameChecking
    Initialize-Logging -LogPath (Join-Path $ModuleRoot 'lab.log')

    $others = @(Get-CimInstance -Namespace $ns -ClassName SMS_Collection -Filter "Name LIKE 'ZZTEST-%'")
    $q = @(Get-CimInstance -Namespace $ns -ClassName SMS_Query -Filter "Name LIKE 'ZZTEST-%'" -ErrorAction SilentlyContinue)
    if ($others.Count -gt 0 -or $q.Count -gt 0) { throw 'ZZTEST- objects already exist. Another test may be running.' }

    Add-Result 'connect' ([bool](Connect-CMSite -SiteCode $SiteCode -SMSProvider $Provider)) 'Connect-CMSite'
    $snapshot = Get-CollectionSnapshot

    # --- create, properties, copy -------------------------------------------------------------------
    $a = New-ManagedCollection -Name $nameA -LimitingCollectionName 'All Systems' -Comment 'ZZTEST temporary' -RefreshType Manual
    Add-Result 'create collection' ([bool]($a.CollectionID -match '^[A-Z0-9]{8}$')) $a.CollectionID
    $b = New-ManagedCollection -Name $nameB -LimitingCollectionName 'All Systems' -RefreshType Periodic
    Add-Result 'create collection with a refresh schedule' ([bool]$b.CollectionID) $b.CollectionID

    $d = Get-CollectionDetail -CollectionId $a.CollectionID
    Add-Result 'detail: name, refresh type, built-in flag' ($d.Name -eq $nameA -and $d.RefreshType -eq 1 -and -not $d.IsBuiltIn) "RefreshType=$($d.RefreshType)"

    Set-CollectionProperties -CollectionId $a.CollectionID -Comment 'ZZTEST changed' -RefreshType 4
    $d = Get-CollectionDetail -CollectionId $a.CollectionID
    Add-Result 'set refresh type to Continuous' ($d.RefreshType -eq 4 -and $d.Comment -eq 'ZZTEST changed') "RefreshType=$($d.RefreshType) Comment=$($d.Comment)"
    $threw = $false; try { Set-CollectionProperties -CollectionId $a.CollectionID -RefreshType 3 } catch { $threw = $true }
    Add-Result 'refresh type 3 is rejected before the site is called' $threw ''

    Set-CollectionProperties -CollectionId $a.CollectionID -NewName "$nameA-renamed"
    $renamed = @(Get-CMDeviceCollection -CollectionId $a.CollectionID)[0].Name -eq "$nameA-renamed"
    Add-Result 'rename' $renamed ''
    Set-CollectionProperties -CollectionId $a.CollectionID -NewName $nameA

    $c = Copy-ManagedCollection -SourceCollectionId $a.CollectionID -NewName $nameCopy
    Add-Result 'copy collection' ([bool]($c.CollectionID -and $c.CollectionID -ne $a.CollectionID)) $c.CollectionID

    # --- direct member --------------------------------------------------------------------------------
    $ok = Add-DirectMember -CollectionId $a.CollectionID -DeviceName 'CLIENT01'
    Add-Result 'add direct member CLIENT01' ([bool]$ok) ''
    $notFound = Add-DirectMember -CollectionId $a.CollectionID -DeviceName 'ZZTEST-NO-SUCH-DEVICE'
    Add-Result 'unknown device returns false' ($notFound -eq $false) ''
    $star = Add-DirectMember -CollectionId $a.CollectionID -DeviceName '*'
    Add-Result 'a wildcard name adds nothing' ($star -eq $false) ''
    $rules = @((Get-CMDeviceCollection -CollectionId $a.CollectionID).CollectionRules)
    Add-Result 'exactly one direct rule exists' ($rules.Count -eq 1) "rules=$($rules.Count)"
    $whatif = Add-DirectMember -CollectionId $b.CollectionID -DeviceName 'CLIENT01' -WhatIf
    Add-Result 'WhatIf adds no member' ($whatif -eq $false -and @((Get-CMDeviceCollection -CollectionId $b.CollectionID).CollectionRules | Where-Object { $_ }).Count -eq 0) ''

    # --- query rules ---------------------------------------------------------------------------------
    $wql = 'select SMS_R_System.ResourceId, SMS_R_System.Name from SMS_R_System where SMS_R_System.Name like "ZZTEST%"'
    $v = Test-WqlQuery -QueryExpression $wql
    Add-Result 'Test-WqlQuery accepts a valid query' ([bool]$v.IsValid) $v.ErrorMessage
    $bad = Test-WqlQuery -QueryExpression 'selec nonsense from'
    Add-Result 'Test-WqlQuery rejects a broken query with a message' (-not $bad.IsValid -and $bad.ErrorMessage) $bad.ErrorMessage
    $pv = Invoke-WqlPreview -QueryExpression 'select SMS_R_System.ResourceId, SMS_R_System.Name from SMS_R_System' -MaxResults 2
    Add-Result 'Invoke-WqlPreview limits the rows' ($pv.TotalCount -ge 1 -and @($pv.Results).Count -le 2) "total=$($pv.TotalCount) shown=$(@($pv.Results).Count)"

    Add-QueryRule -CollectionId $b.CollectionID -RuleName 'ZZTEST rule' -QueryExpression $wql
    $qr = @(Get-CollectionQueryRules -CollectionId $b.CollectionID)
    Add-Result 'add query rule' ($qr.Count -eq 1 -and $qr[0].RuleName -eq 'ZZTEST rule') "rules=$($qr.Count)"

    $wql2 = 'select SMS_R_System.ResourceId, SMS_R_System.Name from SMS_R_System where SMS_R_System.Name like "ZZTEST2%"'
    Update-QueryRule -CollectionId $b.CollectionID -OldRuleName 'ZZTEST rule' -NewRuleName 'ZZTEST rule 2' -NewQueryExpression $wql2
    $qr = @(Get-CollectionQueryRules -CollectionId $b.CollectionID)
    Add-Result 'update query rule' ($qr.Count -eq 1 -and $qr[0].RuleName -eq 'ZZTEST rule 2') ($qr | ForEach-Object RuleName)

    $refused = $null
    try { Update-QueryRule -CollectionId $b.CollectionID -OldRuleName 'ZZTEST rule 2' -NewRuleName 'ZZTEST rule 3' -NewQueryExpression 'this is not WQL' } catch { $refused = $_.Exception.Message }
    $qr = @(Get-CollectionQueryRules -CollectionId $b.CollectionID)
    Add-Result 'a refused update keeps the old rule' ([bool]($refused -and $qr.Count -eq 1 -and $qr[0].RuleName -eq 'ZZTEST rule 2')) "refused=[$refused] rules=$(($qr | ForEach-Object RuleName) -join ',')"

    Remove-QueryRule -CollectionId $b.CollectionID -RuleName 'ZZTEST rule 2'
    Add-Result 'remove query rule' (@(Get-CollectionQueryRules -CollectionId $b.CollectionID).Count -eq 0) ''

    # --- include, exclude, evaluation ------------------------------------------------------------------
    Add-IncludeRule -CollectionId $b.CollectionID -IncludeCollectionId $a.CollectionID
    Add-ExcludeRule -CollectionId $b.CollectionID -ExcludeCollectionId $c.CollectionID
    $rules = @((Get-CMDeviceCollection -CollectionId $b.CollectionID).CollectionRules)
    Add-Result 'add include and exclude rule' ($rules.Count -eq 2) "rules=$($rules.Count)"
    Invoke-CollectionEvaluation -CollectionId $b.CollectionID
    Add-Result 'evaluation triggers without an error' $true ''

    # --- built-in protection and WhatIf -----------------------------------------------------------------
    $blocked = Remove-ManagedCollection -CollectionId 'SMS00001'
    Add-Result 'the built-in collection SMS00001 is not removed' ($blocked -eq $false -and [bool](Get-CMDeviceCollection -CollectionId 'SMS00001')) ''
    $w = Remove-ManagedCollection -CollectionId $c.CollectionID -WhatIf
    Add-Result 'remove under WhatIf keeps the collection' ($w -eq $false -and [bool](Get-CMDeviceCollection -CollectionId $c.CollectionID)) ''

    # --- saved query -----------------------------------------------------------------------------------
    $sq = New-SiteSavedQuery -Name $queryName -Expression 'select SMS_R_System.ResourceId from SMS_R_System' -Comment 'ZZTEST' -TargetClassName 'SMS_R_System'
    Add-Result 'create saved query' ([bool]$sq.QueryID) $sq.QueryID
    $sq2 = Set-SiteSavedQuery -QueryId $sq.QueryID -Name "$queryName-2" -Expression 'select SMS_R_System.ResourceId, SMS_R_System.Name from SMS_R_System' -Comment 'ZZTEST 2' -TargetClassName 'SMS_R_System'
    Add-Result 'update saved query' ($sq2.Name -eq "$queryName-2") $sq2.Name
    Add-Result 'list saved queries includes it' ([bool](Get-SiteSavedQueries | Where-Object QueryID -eq $sq.QueryID)) ''
    Add-Result 'remove saved query' ((Remove-SiteSavedQuery -QueryId $sq.QueryID) -eq $true -and -not (Get-SiteSavedQueries | Where-Object QueryID -eq $sq.QueryID)) ''

    # --- remove ----------------------------------------------------------------------------------------
    # B includes A and excludes the copy, so B goes first.
    foreach ($id in $b.CollectionID, $c.CollectionID, $a.CollectionID) {
        $r = Remove-ManagedCollection -CollectionId $id
        Add-Result "remove collection $id" ($r -eq $true -and -not (Get-CMDeviceCollection -CollectionId $id -ErrorAction SilentlyContinue)) ''
    }
}
catch {
    Add-Result 'run' $false $_.Exception.Message
}
finally {
    try {
        if (-not (Test-CMConnection)) { [void](Connect-CMSite -SiteCode $SiteCode -SMSProvider $Provider) }
        # A collection that another one includes or excludes cannot go first: retry in passes.
        for ($pass = 1; $pass -le 3; $pass++) {
            foreach ($coll in @(Get-CMDeviceCollection -Name 'ZZTEST-*' -ErrorAction SilentlyContinue)) {
                if ($coll.Name -notlike 'ZZTEST-*') { continue }
                try { Remove-CMDeviceCollection -CollectionId $coll.CollectionID -Force -ErrorAction Stop }
                catch { if ($pass -eq 3) { Add-Result "cleanup collection $($coll.Name)" $false $_.Exception.Message } }
            }
        }
        foreach ($sq in @(Get-CMQuery -Name 'ZZTEST-*' -ErrorAction SilentlyContinue)) {
            if ($sq.Name -like 'ZZTEST-*') {
                try { Remove-CMQuery -Id $sq.QueryID -Force -ErrorAction Stop } catch { Add-Result "cleanup query $($sq.Name)" $false $_.Exception.Message }
            }
        }
    } catch { Add-Result 'cleanup' $false $_.Exception.Message }
    try {
        $left = @(Get-CimInstance -Namespace $ns -ClassName SMS_Collection -Filter "Name LIKE 'ZZTEST-%'")
        $leftQ = @(Get-CimInstance -Namespace $ns -ClassName SMS_Query -Filter "Name LIKE 'ZZTEST-%'" -ErrorAction SilentlyContinue)
        Add-Result 'no ZZTEST- object remains' ($left.Count -eq 0 -and $leftQ.Count -eq 0) ((@($left | ForEach-Object Name) + @($leftQ | ForEach-Object Name)) -join ',')
        if ($snapshot) {
            $after = Get-CollectionSnapshot
            Add-Result 'other collections are unchanged' ((@($snapshot) -join "`n") -ceq (@($after) -join "`n")) ''
        }
    } catch { Add-Result 'verify cleanup' $false $_.Exception.Message }
}

$results
