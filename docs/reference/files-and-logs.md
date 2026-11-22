# Files and logs

This page lists every folder and file that Collection Manager writes.

## What it does

The tool writes its global files in the folder of the script and its site data in the data folder of the site. It writes nothing in the registry. The buttons that change the site change the site, not these files.

## When to use it

Use it to find a log or a report, or to clean up.

## Requirements

None.

## Files

| Path | Content | Written when |
|---|---|---|
| `CollectionManager.prefs.json` | The settings. | You select **OK** in Options, or change the theme. |
| `CollectionManager.windowstate.json` | The size, the place and the open view of the window. | You close the window. |
| `Logs\CollectionManager-<yyyyMMdd-HHmmss>.log` | The log of one session. Each read, each change and each failure. | During the session. |
| `Logs\CollectionManager-startup-<yyyyMMdd-HHmmss>.txt` | The PowerShell transcript of the start. | At the start. |
| `%LOCALAPPDATA%\AppPackagerData\sites\<SITE>\tools\collection-manager\prefs.json` | The site code and the SMS Provider of the site. | You select **OK** in Options. |

With no site chosen, the tool uses `Reports\` beside the script for the exports and keeps the site settings in `CollectionManager.prefs.json`. The first start for a site copies the old `Reports` files into the site folder and keeps the old files. See [Sites](../core/options.md#sites).\operational-collections.json` and `Templates\parameterized-templates.json`. It does not write to them.

## Steps

### Clean up

1. Close the tool.
2. Delete the logs and the reports that you do not need. The reports of a site are in the `reports` folder of that site.

## Limits

- The tool never deletes a log or a report.

## Troubleshooting

| Message or symptom | Cause | Fix |
|---|---|---|
| A start leaves no log. | The folder is not writable. | Move the tool to a writable folder. |

## Links

- [Settings](settings.md)
- [Collections](../core/collections.md)
