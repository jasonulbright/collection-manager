# Collection Manager: test plan

Measured 2026-11-06 with `Invoke-Checks.ps1` (Pester 5.7.1, PSScriptAnalyzer 1.25.0). Counts are tests as run; each data row counts as one test. Coverage is the share of module commands that the tests run.

| | Before | After |
|---|---|---|
| Tests | 27 | 457 on Windows PowerShell 5.1 and on PowerShell 7 (32 contract tests skip on a machine without the console) |
| Module coverage | 31.9% (109 of 342 commands) | 98.8% (427 of 432 commands) |
| Lab steps | none | 33 (`Tests/Lab`) |

Run: `.\Invoke-Checks.ps1` (lint, tests, coverage floor 80%). It runs in CI on both hosts (`.github/workflows/checks.yml`).

Window tests: `Window.Tests.ps1` opens a window only when `SUITE_ALLOW_WINDOW_TESTS` is `1`. The CI workflow and the lab run scripts set it. Otherwise Pester reports those tests as Skipped, with the reason `SUITE_ALLOW_WINDOW_TESTS is not 1`, not as Failed. Pester does not store the reason, so `Invoke-Checks.ps1` (and `Window.Tests.ps1`, when it runs alone) prints it to the console. `Test-WindowTestsAllowed` in `Tests/WindowHarness.ps1` makes the decision, and `Get-WpfSession` refuses to create a runspace when it says no. The CI workflow sets the variable in both jobs. `Invoke-Checks.ps1` prints that reason once, and the run still passes. The coverage floor stays in force: the window tests run no module code, so skipping them does not lower the module coverage. `WindowGuard.Tests.ps1` tests the guard with an injected opt-in value and checks that every test file that can open a window carries it. It opens no window.

## What is covered

| Area | File | Contents |
|---|---|---|
| Read side | `Module.Queries.Tests.ps1` | Folder tree and map, all collections (refresh type names, built-in flag, 3000 rows), detail and rule counts, query rules, members (5000), saved queries |
| Write actions | `Module.Write.Tests.ps1` | Create (all refresh types, schedule), copy, remove with built-in protection, properties, direct member (exact match, wildcards, duplicates), query, include and exclude rules, refused update, evaluation. Success, failure, `-WhatIf`, hostile input |
| WQL, templates, export | `Module.Wql.Templates.Export.Tests.ps1` | Validate and preview limits, template files (corrupt, empty, unicode), parameter expansion, CSV and HTML from tables and from objects |
| Script logic | `Shell.Logic.Tests.ps1` | Filters of the Collections and Templates views, status bar text, export columns, preferences file |
| Script calls | `Script.Contract.Tests.ps1` | Every call from the script and the module into an app function or SuiteCommon names only parameters that exist |
| Real cmdlet signatures | `CmCmdlet.Contract.Tests.ps1` | Every ConfigMgr call names only parameters that the real cmdlet has. Runs where the console is installed |
| Start without the console | `Startup.Tests.ps1` | Plain message, `$false`, no raw exception |
| Windows | `Window.Tests.ps1` | Main window: names, titles, the dialogs' XAML parses and has named controls and distinct titles, no clipped text and nothing off screen at 1024x720 and 1440x900 in both themes and all three views, text and control contrast |
| Lab | `Lab/Invoke-CollectionLab.ps1` | ZZTEST- collections and one ZZTEST- saved query on CM01: create, copy, properties, member, query rules (including a refused update), include, exclude, evaluate, built-in protection, WhatIf, remove. Checks that no other collection changed |

## Defects found

| # | What | Fix |
|---|---|---|
| 1 | Export CSV and HTML called the export functions with `-Collections`, which they did not have | Both functions accept `-Collections` and `-Columns` |
| 2 | The template loader passed `-Path` to functions that take `-TemplatesPath`, so no template loaded | `-TemplatesPath` |
| 3 | Update Rule removed the old rule first and lost it when the site refused the new one | Read the old rule, add it back on failure |
| 4 | `Set-CollectionProperties` ignored `-RefreshType` | Applied; values other than 1, 2, 4, 6 throw |
| 5 | `Add-DirectMember` took a wildcard name and added every match | Exact name only; zero or several matches add nothing |
| 6 | `Copy-ManagedCollection` returned the source collection (found in the lab) | Read the clone back by name |
| 7 | Template files read without an encoding (garbled on 5.1); an empty file gave one empty template | UTF-8 and null filter |
| 8 | WQL editor buttons Update Rule and Remove were off screen at 1024 pixels | `WrapPanel` |
| 9 | A built-in id with a leading space passed the `SMS*` guard | Trim the id |
| 10 | Edit Membership cleared the device box and logged success when `Add-DirectMember` added nothing | Check the result, keep the box, say why |
| 11 | Template expansion was sequential: a value that contained `{B}` pulled in parameter B | One pass over the template |
| 12 | `Set-CollectionProperties` with refresh type 2 or 6 sent no schedule | Pass a 7-day schedule, as the create function does |
| 13 | The exporters lost positional binding when `-Collections` was added | `Position` on the table, path and title |

## Not covered, and why

- Contrast of disabled controls, the sidebar button border, the light text boxes and the unselected tab headers misses the standard (text 4.5:1, borders 3:1). The cause is the shared SuiteCommon palette and the MahApps styles. `Tests/KnownContrastDefects.ps1` lists each violation. The window test fails on a new one and on a listed one that is fixed. The fix belongs in `app-packager-suite`.
- The dialogs are not shown with data, so their layout and contrast are not measured. The title bar of an inactive window misses 4.5:1 (listed as `titlebar:inactive`). The border of a disabled sidebar button is read from the button, not from the template border at half opacity.
- The connect and refresh handlers, the dialogs' button logic and the WQL tree run inside the window with a site. The lab run covers the module calls that they make, not the clicks.
- The settings save in SuiteCommon is not atomic. The fix belongs in `app-packager-suite`.
