# Test copy only: make the module path lookup return nothing, as on a computer without the console.
$f = Join-Path $appDir 'Lib\SuiteCommon\SuiteCommon.psm1'
$t = Get-Content -LiteralPath $f -Raw
$patched = [regex]::Replace($t, '(function Resolve-ConfigurationManagerModulePath \{)', '${1} return $null;')
if ($patched -eq $t) { throw 'The module lookup function was not found. The missing-console test cannot run.' }
Set-Content -LiteralPath $f -Value $patched -Encoding UTF8
Remove-Item Env:SMS_ADMIN_UI_PATH -ErrorAction SilentlyContinue
$win = Start-App 'start-collectionmanager.ps1' 'Collection Manager'
Size-Win $win 1440 900
$win = Find-Top 'Collection Manager' 10
Invoke-El (Get-El $win 'btnRefresh')
# Wait for the log to show the missing-console message. A fixed sleep is not evidence.
if (-not (Wait-Until { (Get-ElText (Get-El $win 'txtLog')) -match 'ConfigurationManager module not found' } 30)) { throw 'The missing-console message did not show in the log.' }
Start-Sleep -Seconds 1
Shot $win 'console-missing-refresh'
Stop-App 'start-collectionmanager.ps1'
