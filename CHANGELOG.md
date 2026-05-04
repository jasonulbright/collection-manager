# Changelog

All notable changes to Collection Manager are documented in this
file.

## [1.0.1] - 2026-05-04

### Changed

- **Operational template library expanded from 157 to 225 entries.**
  Source material (SystemCenterDudes) was last updated mid-2022 and
  did not cover incremental OS, Microsoft 365 Apps, or Configuration
  Manager releases since. Legacy entries (Windows XP/7/8, Server
  2003/2008, Windows Phone, retired Win10 builds, "CB/CBB" naming)
  are preserved for environments that still inventory them.
- **Clients Version | Not Latest** baseline updated from 2207 to 2509.

### Added

- **Configuration Manager clients**: Clients Version | 2409, 2503,
  2509 (build prefixes 5.00.9132 / 9135 / 9141).
- **Servers**: Windows Server 2022 (build 20348) and Windows Server
  2025 (build 26100).
- **Workstations**: Windows 11 v25H2 (build 10.0.26200), Defender
  for Endpoint Onboarded / Not Onboarded (current naming for the
  same `AdvancedThreatProtectionHealthStatus` queries).
- **Systems | ARM64**: detects ARM64-based PCs via
  `SMS_G_System_COMPUTER_SYSTEM.SystemType`.
- **Microsoft 365 Apps Build Version**: monthly Current Channel
  builds 2208 through 2603 (43 entries).
- **Microsoft 365 Apps Channel**: Current Channel, Current Channel
  (Preview), Monthly Enterprise Channel, Semi-Annual Enterprise
  Channel, Semi-Annual Enterprise Channel (Preview), Beta Channel.
  These supersede the legacy "Office 365 Channel | Monthly /
  Semi-Annual" naming retired by Microsoft in September 2020.
- **Visio**: Installed (Any Edition), Subscription (Plan 1 / Plan 2),
  LTSC 2024, LTSC 2021, 2019.
- **Project**: Installed (Any Edition), Subscription (Plan 1 / Plan
  3 / Plan 5), LTSC 2024, LTSC 2021, 2019.

## [1.0.0] - 2026-05-02

Collection Manager is a MECM device collection manager with an
offline WQL editor and a 177-template library (157 operational +
20 parameterized; expanded to 225 + 20 in subsequent releases). Ships as a zip + `install.ps1` wrapper; no MSI,
no code signing required.

### Features

- **Collections view** -- master-detail grid of every device
  collection with name, ID, member count, limiting collection,
  refresh type, and comment. Text + membership-shape filters
  (Built-in / Custom / Empty / Has direct members / Has query
  rules). Detail panel tabs: Properties, Direct Members, Query
  Rules, Include Rules, Exclude Rules. Action bar: New, Copy,
  Edit Membership, Evaluate, Remove.
- **WQL Editor view** -- collection picker, query-rule list,
  monospace editor with Validate (`Invoke-CMWmiQuery`) and
  Preview Results actions. Add, Update, and Remove rules.
- **Templates view** -- Operational and Parameterized tabs over
  the shipped library; live parameter form with placeholder
  expansion; Copy to WQL Editor and Apply to Collection actions.
- **Template library**
  - 157 ready-made operational collection queries
    (`operational-collections.json`) adapted from SystemCenterDudes,
    grouped by category (Clients, Hardware Inventory, Laptops,
    Mobile Devices, Office 365, Servers, Software Inventory, System
    Health, Windows Update Agent, Workstations).
  - 20 parameterized WQL templates (`parameterized-templates.json`)
    with `{Placeholder}` substitution: Software Installed / Not
    Installed, OS Build, Client Version, Manufacturer, Model,
    HW/SW Inventory Not Reporting, AD OU, Created Within N Days,
    Chassis Type, TPM, Disk Space, RAM, IP Subnet, Domain, BIOS
    Version, BitLocker, Last Boot.
- **Modal dialogs** -- New Collection, Copy Collection, Edit
  Membership (direct add/remove with bulk confirm), Apply Template
  to Collection, themed Remove confirmations. All MetroWindow
  inline-XAML, theme-honoring, drag-fallback installed.
- **Safety** -- blocks deletion of built-in collections (SMS prefix),
  confirmation dialogs before every destructive operation, lazy
  member loads to keep refresh times bounded for large environments.
- **Export** -- CSV and HTML reports of collection metadata.
- **Core module** `CollectionManagerCommon.psm1` with 28 exported
  functions covering logging, CM connection, collection queries,
  CRUD, membership, WQL validation and preview, templates, and
  export. 23 Pester 5.x tests in
  `Module/CollectionManagerCommon.Tests.ps1`.
- **WPF brand alignment** -- MahApps.Metro shell with sidebar
  navigation, glyph status (system-managed `...` for built-in
  collections), animated ProgressRing during refresh, log drawer,
  status bar, dark and light themes with runtime toggle, window-
  state persistence including legacy WinForms schema bridge.

### Stack

- PowerShell 5.1 + .NET Framework 4.7.2+
- WPF + MahApps.Metro 2.4.10 (vendored under `Lib/`)
- ConfigurationManager PowerShell module (CM console required on
  the host machine running this app)
