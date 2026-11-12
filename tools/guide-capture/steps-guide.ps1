$win = Start-App 'start-collectionmanager.ps1' 'Collection Manager'
Size-Win $win 1440 900
$win = Find-Top 'Collection Manager' 10
function Idle([string]$id = 'btnRefresh', [int]$sec = 300) { Start-Sleep -Seconds 2; if (-not (Wait-Until { $b = Get-El $win $id -Quiet; $b -and $b.Current.IsEnabled } $sec)) { throw ("Timeout: $id did not become enabled in $sec seconds.") }; Start-Sleep -Seconds 2 }
function Rows($grid) { @((Get-El $win $grid).FindAll('Descendants', (New-Object System.Windows.Automation.PropertyCondition($AE::ControlTypeProperty, [System.Windows.Automation.ControlType]::DataItem)))) }
function Find-Row($grid, [string]$name) { @(Rows $grid) | Where-Object { $_.Current.Name -match ('Name=' + [regex]::Escape($name) + ';') } | Select-Object -First 1 }
function Select-Row($grid, [string]$name) { $r = Find-Row $grid $name; if (-not $r) { Log ('ERROR: no row ' + $name + ' in ' + $grid); return $false }; Select-El $r; Start-Sleep -Milliseconds 600; return $true }
function Pick($combo, [string]$item) { Expand-El (Get-El $win $combo); Start-Sleep -Milliseconds 500; Click-El (Find-Role $win 'ListItem' $item); Start-Sleep -Milliseconds 700 }
function Tab([string]$name) { Select-El (Find-Role $win 'TabItem' $name); Start-Sleep -Milliseconds 900 }
function WinNames { $c = New-Object System.Windows.Automation.PropertyCondition($AE::ControlTypeProperty, [System.Windows.Automation.ControlType]::Window); (@($root.FindAll('Descendants', $c)) | ForEach-Object { $_.Current.Name }) -join ' ; ' }
function Dlg([string]$btn, [string]$shot) {
  Invoke-El (Get-El $win $btn); Start-Sleep -Seconds 2
  Log ('after ' + $btn + ' windows: ' + (WinNames))
  $c = New-Object System.Windows.Automation.PropertyCondition($AE::ControlTypeProperty, [System.Windows.Automation.ControlType]::Window)
  $d = @($root.FindAll('Descendants', $c)) | Where-Object { $_.Current.Name -notin 'Collection Manager', 'Program Manager' -and $_.Current.Name -ne '' -and $_.Current.Name -notlike 'Administrator*' -and $_.Current.Name -notlike 'GlowWindow*' } | Select-Object -First 1
  if ($d) { Shot $d $shot; Dump-Tree $d $shot; Focus-Win $d } else { Log ('ERROR: no dialog after ' + $btn); Shot $win $shot }
  Send-Keys '{ESC}'; Start-Sleep -Seconds 1
}
Section 'first run' {
  Shot $win 'start-first-run'
  Invoke-El (Get-El $win 'btnRefresh'); Start-Sleep -Seconds 2; Shot $win 'refresh-without-connection'
  Invoke-El (Get-El $win 'btnOptions'); $d = Find-Last 'Options' 10; Shot $d 'options-empty'; Focus-Win $d; Send-Keys '{ESC}'; Start-Sleep -Seconds 1
}
Section 'options' {
  Invoke-El (Get-El $win 'btnOptions'); $d = Find-Last 'Options' 10
  Set-Text (Get-El $d 'txtSiteCode') $Capture.SiteCode; Set-Text (Get-El $d 'txtSmsProvider') $Capture.SMSProvider; Shot $d 'options-connection'
  Invoke-El (Get-El $d 'btnCatAbout'); Start-Sleep -Milliseconds 600; Shot $d 'options-about'
  Invoke-El (Get-El $d 'btnCatConnection'); Invoke-El (Get-El $d 'btnOk'); Start-Sleep -Seconds 1
}
Section 'connect' {
  Invoke-El (Get-El $win 'btnConnect'); Start-Sleep -Seconds 3; Shot $win 'connect-running'; Idle 'btnConnect'; Shot $win 'connected'
}
Section 'refresh' {
  Invoke-El (Get-El $win 'btnRefresh'); Start-Sleep -Seconds 3; Shot $win 'refresh-running'; Idle; Shot $win 'collections'
}
Section 'collection detail' {
  if (-not (Select-Row 'gridCollections' 'All Systems')) { return }
  Shot $win 'collection-properties'
  foreach ($t in @(@('Direct Members', 'collection-direct-members'), @('Query Rules', 'collection-query-rules'), @('Include Rules', 'collection-include-rules'), @('Exclude Rules', 'collection-exclude-rules'))) { Tab $t[0]; Shot $win $t[1] }
  Tab 'Properties'
}
Section 'exports' {
  Invoke-El (Get-El $win 'btnExportCsv'); File-Dialog $null $null; Shot $win 'collections-exported'
  Invoke-El (Get-El $win 'btnExportHtml'); File-Dialog $null $null; Start-Sleep -Seconds 1
}
Section 'dialogs' {
  Dlg 'btnNewCollection' 'dialog-new-collection'
  if (Select-Row 'gridCollections' 'All Workstations') { Shot $win 'collection-custom-selected' }
  Dlg 'btnCopyCollection' 'dialog-copy-collection'
  Dlg 'btnEditMembership' 'dialog-edit-membership'
}
Section 'blocked on built-in' {
  if (-not (Select-Row 'gridCollections' 'All Systems')) { return }
  Invoke-El (Get-El $win 'btnEditMembership'); Start-Sleep -Seconds 1; Invoke-El (Get-El $win 'btnRemoveColl'); Start-Sleep -Seconds 1; Shot $win 'builtin-blocked'
}
Section 'remove confirm' {
  # The confirmation is captured and answered No. Nothing is removed.
  # The step stops when the intended row is not selected or the confirmation does not show.
  if (-not (Select-Row 'gridCollections' 'All Workstations')) { return }
  Invoke-El (Get-El $win 'btnRemoveColl'); Start-Sleep -Seconds 2
  $no = Get-El $win 'No' -Quiet
  if (-not $no) { Log 'ERROR: the Remove confirmation did not show. Nothing was answered.'; return }
  Log ('windows: ' + (WinNames)); Shot $win 'dialog-remove-collection'
  Invoke-El $no; Start-Sleep -Seconds 2
  if (Get-El $win 'No' -Quiet) { Log 'ERROR: the Remove confirmation is still open after No.'; return }
  Shot $win 'after-remove-cancel'
}
Section 'filters' {
  Set-Text (Get-El $win 'txtFilter') 'All'; Start-Sleep -Seconds 1; Shot $win 'collections-filter-text'; Set-Text (Get-El $win 'txtFilter') ''
  Expand-El (Get-El $win 'cboCollectionFilter'); Start-Sleep -Milliseconds 500; Shot $win 'collections-filter-list' -Hover; Collapse-El (Get-El $win 'cboCollectionFilter')
}
Section 'wql' {
  Invoke-El (Get-El $win 'btnViewWqlEditor'); Start-Sleep -Seconds 2; Shot $win 'wql-editor'
  $c = New-Object System.Windows.Automation.PropertyCondition($AE::ControlTypeProperty, [System.Windows.Automation.ControlType]::TreeItem)
  $it = @((Get-El $win 'treeWqlCollections').FindAll('Descendants', $c)) | Where-Object { $_.Current.Name -like '*All Workstations*' } | Select-Object -First 1
  if (-not $it) { $c2 = New-Object System.Windows.Automation.PropertyCondition($AE::ControlTypeProperty, [System.Windows.Automation.ControlType]::ListItem); $it = @((Get-El $win 'treeWqlCollections').FindAll('Descendants', $c2)) | Where-Object { $_.Current.Name -like '*All Workstations*' } | Select-Object -First 1 }
  Log ('tree item: ' + $it.Current.Name); Select-El $it; Start-Sleep -Seconds 2; Shot $win 'wql-collection-picked'
  $l = (Get-El $win 'lstWqlRules').FindFirst('Descendants', (New-Object System.Windows.Automation.PropertyCondition($AE::ControlTypeProperty, [System.Windows.Automation.ControlType]::ListItem)))
  if ($l) { Select-El $l; Start-Sleep -Seconds 1; Shot $win 'wql-rule-selected' }
  Invoke-El (Get-El $win 'btnValidateWql'); Start-Sleep -Seconds 4; Shot $win 'wql-validated'
  Invoke-El (Get-El $win 'btnPreviewWql'); Start-Sleep -Seconds 5; Shot $win 'wql-previewed'
  Tab 'Site Queries'; Start-Sleep -Seconds 1; Shot $win 'wql-site-queries-before'; Invoke-El (Get-El $win 'btnRefreshSavedQueries'); Start-Sleep -Seconds 5; Shot $win 'wql-site-queries'
  $r = Rows 'gridSavedQueries'; Log ('saved queries: ' + $r.Count); if ($r.Count) { Select-El $r[0]; Start-Sleep -Seconds 1; Shot $win 'wql-site-query-selected' }
}
Section 'templates' {
  Invoke-El (Get-El $win 'btnViewTemplates'); Start-Sleep -Seconds 2; Shot $win 'templates-operational'
  $r = Rows 'gridOperationalTemplates'; Select-El $r[0]; Start-Sleep -Seconds 1; Shot $win 'template-selected'
  Tab 'Parameterized'; Shot $win 'templates-parameterized'
  $r = Rows 'gridParameterizedTemplates'; Select-El $r[0]; Start-Sleep -Seconds 1; Shot $win 'template-parameterized-selected'
  $e = @($win.FindAll('Descendants', (New-Object System.Windows.Automation.PropertyCondition($AE::ControlTypeProperty, [System.Windows.Automation.ControlType]::Edit)))) | Where-Object { $_.Current.AutomationId -eq '' -and $_.Current.BoundingRectangle.Width -gt 200 -and $_.Current.BoundingRectangle.Y -gt 250 } | Select-Object -First 1
  if ($e) { Set-Text $e '7-Zip'; Start-Sleep -Seconds 1; Shot $win 'template-parameter-filled' } else { Log 'ERROR: no parameter edit' }
  Dlg 'btnApplyTemplate' 'dialog-apply-template'
}
Section 'tab order' { Invoke-El (Get-El $win 'btnViewCollections'); Focus-Win $win; Click-El (Get-El $win 'btnViewCollections'); 1..14 | ForEach-Object { Send-Keys '{TAB}'; $f = [System.Windows.Automation.AutomationElement]::FocusedElement; Log ('tab -> ' + $f.Current.AutomationId + ' / ' + $f.Current.Name) } }
Section 'light' {
  Toggle-El (Get-El $win 'toggleTheme'); Start-Sleep -Seconds 1
  foreach ($v in @(@('Collections', 'light-collections'), @('Templates', 'light-templates'))) { Invoke-El (Get-El $win ('btnView' + $v[0])); Start-Sleep -Seconds 2; Shot $win $v[1] }
  Invoke-El (Get-El $win 'btnOptions'); $d = Find-Last 'Options' 10; Shot $d 'light-options'; Focus-Win $d; Send-Keys '{ESC}'; Start-Sleep -Seconds 1
  # Selected states in the light theme: every blue background must carry white text.
  Invoke-El (Get-El $win 'btnViewCollections'); Start-Sleep -Seconds 1
  Select-Row 'gridCollections' 'All Workstations' | Out-Null; Shot $win 'light-collection-selected'
  Expand-El (Get-El $win 'cboCollectionFilter'); Start-Sleep -Milliseconds 500; Shot $win 'light-collections-filter-list' -Hover; Collapse-El (Get-El $win 'cboCollectionFilter')
  Dlg 'btnNewCollection' 'light-dialog-new-collection'
  Dlg 'btnCopyCollection' 'light-dialog-copy-collection'
  Dlg 'btnEditMembership' 'light-dialog-edit-membership'
  Select-Row 'gridCollections' 'All Workstations' | Out-Null
  Invoke-El (Get-El $win 'btnRemoveColl'); Start-Sleep -Seconds 2
  $no = Get-El $win 'No' -Quiet
  Log ('before WQL: ' + (WinNames))
  if ($no) { Shot $win 'light-dialog-remove-collection'; Invoke-El $no; Start-Sleep -Seconds 2 } else { Log 'no confirmation in light'; Send-Keys '{ESC}'; Start-Sleep -Seconds 1 }
  Start-Sleep -Seconds 2; Log ('windows: ' + (WinNames))
  Invoke-El (Get-El $win 'btnViewWqlEditor'); Start-Sleep -Seconds 2
  Tab 'Collections'
  $c = New-Object System.Windows.Automation.PropertyCondition($AE::ControlTypeProperty, [System.Windows.Automation.ControlType]::TreeItem)
  $it = @((Get-El $win 'treeWqlCollections').FindAll('Descendants', $c)) | Where-Object { $_.Current.Name -like '*All Workstations*' } | Select-Object -First 1
  if ($it) { Select-El $it; Start-Sleep -Seconds 2 }
  $l = (Get-El $win 'lstWqlRules').FindFirst('Descendants', (New-Object System.Windows.Automation.PropertyCondition($AE::ControlTypeProperty, [System.Windows.Automation.ControlType]::ListItem)))
  if ($l) { Select-El $l; Start-Sleep -Seconds 1 }
  Shot $win 'light-wql-selected'
  Tab 'Site Queries'; Invoke-El (Get-El $win 'btnRefreshSavedQueries'); Start-Sleep -Seconds 5
  $r = Rows 'gridSavedQueries'; if ($r.Count) { Select-El $r[0]; Start-Sleep -Seconds 1 }; Shot $win 'light-site-query-selected'
  Tab 'Collections'
  Invoke-El (Get-El $win 'btnViewTemplates'); Start-Sleep -Seconds 2
  Tab 'Operational'
  $r = Rows 'gridOperationalTemplates'; Select-El $r[0]; Start-Sleep -Seconds 1; Shot $win 'light-template-selected'
  Toggle-El (Get-El $win 'toggleTheme')
}
Section 'persist' { Invoke-El (Get-El $win 'btnViewTemplates'); Start-Sleep -Milliseconds 600; Invoke-El (Get-El $win 'Close'); Start-Sleep -Seconds 8; $win = Start-App 'start-collectionmanager.ps1' 'Collection Manager'; Shot $win 'start-reopen' }
Section 'files' {
  New-Item -ItemType Directory "$OutDir\files" -Force | Out-Null
  Get-ChildItem $appDir -Recurse -File -ErrorAction SilentlyContinue | Where-Object { $_.FullName -notlike '*\ui\*' -and $_.FullName -notlike '*\Lib\*' -and $_.FullName -notlike '*\Module\*' -and $_.FullName -notlike '*\Catalog\*' -and $_.FullName -notlike '*\Templates\*' -and $_.Extension -in '.json', '.log', '.csv', '.html', '.txt' } | ForEach-Object { Log ('file: ' + $_.FullName.Substring($appDir.Length) + ' ' + $_.Length); Copy-Item $_.FullName ("$OutDir\files\" + $_.Name) -Force }
}
Stop-App 'start-collectionmanager.ps1'
