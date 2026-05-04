# Collection Manager

A MahApps.Metro WPF GUI for managing MECM (Configuration Manager) device collections with an offline WQL editor. Bypass the slow console query editor -- load all collections into a local grid, edit WQL queries in a fast monospace editor, validate, preview results, and apply changes in bulk.

Ships with a template library: 225 ready-made operational queries and 20 parameterized WQL templates.

![Collection Manager](screenshot.png)

## Requirements

- Windows 10 / 11
- PowerShell 5.1
- .NET Framework 4.7.2+
- Configuration Manager console installed (provides the `ConfigurationManager` PowerShell module)

## Quick Start

1. Download the release zip and extract it to a working folder.
2. Right-click `start-collectionmanager.ps1` -> **Run with PowerShell**, or from a PowerShell prompt:

   ```powershell
   powershell -ExecutionPolicy Bypass -File start-collectionmanager.ps1
   ```
3. Click the **Options** button on the sidebar and set your Site Code and SMS Provider.
4. Click **Refresh** to load every device collection.

## Layout

The shell uses a sidebar layout with three views and an Options modal:

- **Collections** -- master-detail grid of every device collection: Name, ID, Member Count, Limiting Collection, Refresh Type, Comment. Detail panel tabs: Properties, Direct Members, Query Rules, Include Rules, Exclude Rules. Filter by membership shape (Built-in / Custom / Empty / Has direct members / Has query rules) or free text.
- **WQL Editor** -- pick a collection, pick a rule, edit the WQL in the monospace editor, click **Validate Query** to syntax-check via `Invoke-CMWmiQuery`, click **Preview Results** to see how many devices match. Add, Update, or Remove rules with single clicks.
- **Templates** -- two tabs (Operational, Parameterized). Pick a template; for parameterized ones, fill in the parameters and watch the expanded WQL update live. Then **Copy to WQL Editor** for hand-tuning, or **Apply to Collection** to push the rule onto a target collection.

## Workflows

### Create or copy a collection

1. **Collections** view -> **New Collection**. Pick name, limiting collection, comment, refresh type. Click Create.
2. To clone an existing one: select the source row, click **Copy...**, name the clone, click Copy.

### Edit direct membership

1. Select a collection on the **Collections** view.
2. Click **Edit Membership...** to open the modal: add devices by name (auto-resolves Resource ID), remove selected members in bulk with confirmation.

### Author a new WQL rule

1. Switch to **WQL Editor** view.
2. Pick a target collection.
3. Type or paste the WQL. Click **Validate Query** to syntax-check.
4. Click **Preview Results** to count matches.
5. Type a rule name and click **Add Rule**.

### Apply a ready-made template

1. Switch to **Templates** view -> **Operational** tab.
2. Pick a template (e.g., "Workstations | Windows 11 24H2"). The expanded WQL appears in the right pane.
3. Click **Apply to Collection...**, pick a target, confirm the rule name, click Apply.

### Use a parameterized template

1. **Templates** view -> **Parameterized** tab.
2. Pick a template (e.g., "Software Installed by Name").
3. Fill in the parameters (e.g., SoftwareName = `7-Zip`). The expanded WQL updates live.
4. **Apply to Collection...** or **Copy to WQL Editor** for further tuning.

## Template Categories

**Operational**: Clients, Clients Version, Hardware Inventory, Laptops, Microsoft 365 Apps Build Version / Channel, Mobile Devices, Office 365 Build / Channel (legacy), Project, SCCM Infrastructure, Servers, Software Inventory, System Health, Systems, Visio, Windows Update Agent, Workstations.

**Parameterized**: Software (Installed / Not Installed by Name, Version), OS Build, Client Version, Manufacturer, Model, HW / SW Inventory Reporting, AD OU, Created Within N Days, Chassis Type, TPM, Disk Space, RAM, IP Subnet, Domain, BIOS Version, BitLocker, Last Boot.

## Project Structure

```
collectionmanager/
+- start-collectionmanager.ps1               # WPF shell
+- MainWindow.xaml                           # Main window layout
+- Lib/                                      # Vendored MahApps.Metro 2.4.10
+- Module/
|  +- CollectionManagerCommon.psd1           # Module manifest
|  \- CollectionManagerCommon.psm1           # Business logic (28 functions)
+- Templates/
|  +- operational-collections.json           # 225 ready-made queries
|  \- parameterized-templates.json           # 20 parameterized templates
+- Logs/                                     # Session logs (per-run)
+- Reports/                                  # CSV / HTML exports
+- CHANGELOG.md
+- LICENSE
\- README.md
```

## Safety

- Built-in collections (SMS prefix) cannot be deleted.
- Confirmation dialogs before every destructive operation (collection removal, member removal, query-rule removal).
- WQL validation before applying query rules.
- Lazy loading of direct members keeps refreshes bounded for large environments.

## License

This project is licensed under the [MIT License](LICENSE).

## Author

Jason Ulbright

## Credits

Ready-made operational collection queries adapted from [SystemCenterDudes](https://www.systemcenterdudes.com/create-operational-sccm-collection-using-powershell-script/) / [prae1809/PowerShell-Scripts](https://github.com/prae1809/PowerShell-Scripts/tree/master/OperationalCollections).
