# Refresh the shipped version lists

The tool ships lists of Windows builds, ConfigMgr client versions and Microsoft 365 Apps builds. Microsoft releases a new value in each list every month, every half year or every year. This page lists each list, where it lives and which Microsoft page gives the newest value. Do this procedure after each ConfigMgr release, after each Windows feature update and each month for Microsoft 365 Apps.

## Lists

| List | File | Entry name | Source page |
|---|---|---|---|
| ConfigMgr client version | `Templates/operational-collections.json` | `Clients Version \| <yymm>` and `Clients Version \| Not Latest (<yymm>)` | https://learn.microsoft.com/intune/configmgr/core/servers/manage/updates |
| Microsoft 365 Apps build | `Templates/operational-collections.json` | `Microsoft 365 Apps Build \| <yymm>` | https://learn.microsoft.com/officeupdates/update-history-microsoft365-apps-by-date |
| Microsoft 365 Apps channel | `Templates/operational-collections.json` | `Microsoft 365 Apps Channel \| <name>` | https://learn.microsoft.com/intune/device-configuration/settings-catalog/update-office |
| Windows 10 and 11 build | `Templates/operational-collections.json` | `Workstations \| Windows <10 or 11> v<release>` | https://learn.microsoft.com/windows/release-health/windows11-release-information and https://learn.microsoft.com/windows/release-health/release-information |
| Windows Server build | `Templates/operational-collections.json` | `Servers \| Windows <year>` | https://learn.microsoft.com/windows/release-health/windows-server-release-info |
| Template help text | `Templates/parameterized-templates.json` | `HelpText` of `Client Version` and `OS Build Number` | the pages above |
| Reference table for the tests | `Tests/ShippedVersions.psd1` | every list above | the pages above |

The tool carries no list for .NET, Edge, WinPE, ADK or SQL Server.

## Steps

1. Open each source page. Find the newest value and any value that the list lacks.
2. Add a row to `Tests/ShippedVersions.psd1`. Change `VerifiedOn` and the `Newest` section.
3. Add the entry to `Templates/operational-collections.json`. Copy the entry of the previous release. Change the name and the build.
   - ConfigMgr client: the query is `ClientVersion like '5.00.<build>.10%'`. The build is the second part of the site version on the page, for example 5.00.9152.1000 gives `9152`.
   - Microsoft 365 Apps: the query is `VersionToReport like '16.0.<build>.%'`. The build is the first number of the Build column.
   - Windows: the query is `SMS_R_System.Build = '10.0.<build>'`. For a server it is `BuildNumber = '<build>'`.
4. Rename `Clients Version | Not Latest (<yymm>)` and change its build to the newest release.
5. Change the example in `HelpText` in `Templates/parameterized-templates.json`.
6. Update the template count. Search for the old count with `git grep` and change it in `README.md`, `docs/index.md`, `docs/core/templates.md` and the two Templates strings in `start-collectionmanager.ps1`.
7. Run `Invoke-Checks.ps1` on a lab computer or in CI. `Tests/ShippedLists.Tests.ps1` fails when:
   - a template field is missing,
   - a template names a build, version or channel that the reference table does not know,
   - a template build differs from the reference build,
   - the template list lacks the newest value.
8. Add a line to `CHANGELOG.md` and make a release (`RELEASING.md`).

Mark a value that you cannot confirm on the source page as Unverified in `Tests/ShippedVersions.psd1` (section `Unverified`). Do not guess a build.
