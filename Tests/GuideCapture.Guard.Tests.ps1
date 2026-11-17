#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.1.0' }
# Guard tests for the guide capture tool. They parse the script text. They connect to no
# computer and open no window, so they run on Windows PowerShell 5.1 and on PowerShell 7.
#
# Rules:
#   1. The tool takes C:\guide\capture.lock with New-Item before it registers the scheduled task.
#   2. The tool waits for another capture (-LockWaitMinutes, default 90) and treats a lock
#      older than 30 minutes as stale.
#   3. The poll loop touches the lock, so a live capture never looks stale.
#   4. The finally block releases the lock, and only when this capture took it.

BeforeAll {
    $script:ScriptPath = Join-Path (Split-Path -Path $PSScriptRoot -Parent) 'tools\guide-capture\Invoke-GuideCapture.ps1'
    $script:Text = [System.IO.File]::ReadAllText($script:ScriptPath)
    $script:Ast = [System.Management.Automation.Language.Parser]::ParseFile($script:ScriptPath, [ref]$null, [ref]$null)
    $script:Finally = @($script:Ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.TryStatementAst] -and $n.Finally }, $true))[0].Finally
}

Describe 'Invoke-GuideCapture capture lock' {
    It 'names C:\guide\capture.lock and has a LockWaitMinutes parameter with the default 90' {
        $script:Text | Should -Match '\$lockPath = ''C:\\guide\\capture\.lock'''
        $p = $script:Ast.ParamBlock.Parameters | Where-Object { $_.Name.VariablePath.UserPath -eq 'LockWaitMinutes' }
        $p | Should -Not -BeNullOrEmpty
        $p.DefaultValue.Extent.Text | Should -Be '90'
    }

    It 'takes the lock with New-Item and an IOException catch, and treats a lock as stale only after 30 minutes with no running owner' {
        $newItem = @($script:Ast.FindAll({
                    param($n)
                    $n -is [System.Management.Automation.Language.CommandAst] -and $n.GetCommandName() -eq 'New-Item' -and
                    $n.Extent.Text -match '-ItemType File' -and $n.Extent.Text -match '-ErrorAction Stop'
                }, $true))
        $newItem.Count | Should -Be 1
        $catch = @($script:Ast.FindAll({
                    param($n)
                    $n -is [System.Management.Automation.Language.CatchClauseAst] -and
                    @($n.CatchTypes | Where-Object { $_.TypeName.FullName -eq 'System.IO.IOException' }).Count -gt 0
                }, $true))
        $catch.Count | Should -Be 1
        $script:Text | Should -Match '-ArgumentList \$lockPath, \$lockText'
        $script:Text | Should -Match '\[int\]\$age -gt 30 -and \$aliveLab -ne ''True'' -and -not \(Get-Process -Id \$ownerPid'
        $script:Text | Should -Match 'The stale lab lock .* was removed'
        $script:Text | Should -Match '''unknown: '' \+ \$held'
    }

    It 'writes the lock text as owner|pid|start' {
        $script:Text | Should -Match "\`$lockText = '\{0\}\|\{1\}\|\{2\}' -f \`$app, \`$PID, \(Get-Date\)"
    }

    It 'takes the lock before it registers the scheduled task' {
        $iLock = $script:Text.IndexOf('$lockEnd = (Get-Date).AddMinutes')
        $iRegister = $script:Text.IndexOf('Register-ScheduledTask')
        $iLock | Should -BeGreaterThan -1
        $iRegister | Should -BeGreaterThan $iLock
    }

    It 'touches the lock in the poll loop' {
        $script:Text | Should -Match '\(Get-Item -LiteralPath \$lock -ErrorAction SilentlyContinue\)\.LastWriteTime = Get-Date'
        $script:Text | Should -Match '-ArgumentList \$root, \$taskName, \$lockPath'
    }

    It 'releases the lock in the finally block, only when this capture took it, and never removes C:\guide' {
        $script:Finally | Should -Not -BeNullOrEmpty
        $fin = $script:Finally.Extent.Text
        $fin | Should -Match 'if \(\$haveLock\)'
        $fin | Should -Match '-eq \$text\)'
        $removals = @($script:Ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.CommandAst] -and $n.GetCommandName() -eq 'Remove-Item' }, $true))
        foreach ($c in $removals) { $c.Extent.Text | Should -Not -Match 'C:\\guide\b(?!\\)' -Because $c.Extent.Text }
    }

    It 'releases the lock after it stops the capture and removes the folder of this capture' {
        $fin = $script:Finally.Extent.Text
        $iFolder = $fin.IndexOf('Remove-Item $r -Recurse')
        $iLock = $fin.IndexOf('Remove-Item -LiteralPath $path -Force')
        $iFolder | Should -BeGreaterThan -1
        $iLock | Should -BeGreaterThan $iFolder
    }
}
