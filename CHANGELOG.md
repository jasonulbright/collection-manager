# Changelog

All notable changes to the Collection Manager are documented in this file.

## [1.0.1] - 2026-03-04

### Fixed
- `Copy-CMCollection` missing `-PassThru` -- clone result was always `$null`, so the logged CollectionID was blank
- SplitContainer `SplitterDistance` on WQL Editor and Templates tabs deferred to `Shown` event -- splitters were always 100px because `Width` is 0 before layout

---

## [1.0.0] - 2026-03-03

### Added
- **WinForms GUI** (`start-collectionmanager.ps1`) for managing MECM device collections with offline WQL editing
  - Header panel, connection bar with Load Collections button
  - 4 summary cards: Total Collections, Query Rules, Direct Members, Include/Exclude
  - Text filter across collection name, ID, and comment
  - Log console with timestamped progress messages
  - Dark/light theme, window state persistence, preferences dialog

- **5 tabbed views**
  - **Collections** -- master-detail grid of all device collections with name, ID, member count, limiting collection, refresh type, comment; detail panel on selection
  - **WQL Editor** -- collection selector, query rule list, monospace WQL text editor (Consolas), Validate Query button (via `Invoke-CMWmiQuery`), Preview Results button, Add/Remove/Update Rule buttons
  - **Create / Clone** -- create new collections with name, limiting collection, comment, refresh type (Manual/Periodic/Continuous/Both); clone existing collections via `Copy-CMCollection`
  - **Direct Members** -- collection selector, current members grid, add devices by name (multi-line), remove selected members
  - **Templates** -- TreeView with two root categories (Ready-Made and Parameterized); template detail panel with dynamic parameter inputs; generated WQL preview; "Copy to WQL Editor" and "Create Collection from Template" buttons

- **Template library**
  - 155 ready-made operational collection queries (`operational-collections.json`) adapted from SystemCenterDudes, organized by category (Clients, Clients Version, Hardware Inventory, Laptops, Mobile Devices, Office 365, Others, SCCM Infrastructure, Servers, Software Inventory, System Health, Systems, Windows Update Agent, Workstations)
  - 20 parameterized WQL templates (`parameterized-templates.json`) with `{Placeholder}` substitution: Software Installed/Not Installed by Name, OS Build, Client Version, Manufacturer, Model, HW/SW Inventory Not Reporting, AD OU, Created Within N Days, Chassis Type, TPM, Disk Space, RAM, IP Subnet, Domain, BIOS Version, BitLocker, Last Boot

- **Core module** (`CollectionManagerCommon.psm1`) with 27 exported functions
  - Logging: Initialize-Logging, Write-Log
  - CM Connection: Connect-CMSite, Disconnect-CMSite, Test-CMConnection
  - Collection queries: Get-AllDeviceCollections, Get-CollectionDetail, Get-CollectionQueryRules, Get-CollectionMembers
  - CRUD: New-ManagedCollection, Copy-ManagedCollection, Remove-ManagedCollection (blocks SMS* built-in), Set-CollectionProperties
  - Membership: Add-DirectMember (auto-resolves ResourceId), Remove-DirectMember, Add-QueryRule, Remove-QueryRule, Update-QueryRule, Add-IncludeRule, Add-ExcludeRule, Invoke-CollectionEvaluation
  - WQL: Test-WqlQuery (Invoke-CMWmiQuery validation), Invoke-WqlPreview (first N results)
  - Templates: Get-OperationalTemplates, Get-ParameterizedTemplates, Expand-TemplateParameters
  - Export: Export-CollectionCsv, Export-CollectionHtml

- **Safety checks**: blocks deletion of built-in collections (SMS* prefix), confirmation dialogs before destructive operations

- `CollectionManagerCommon.Tests.ps1` -- 22 Pester 5.x tests covering logging, template loading (operational + parameterized), parameter expansion, and export functions
