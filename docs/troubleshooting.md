# Troubleshooting

This page lists the most common problems of Collection Manager with cause and fix. See [Error messages](reference/error-messages.md) for every message.

## What it does

It gives a short fix for each common problem.

## When to use it

Use it when a connection, a refresh, a change or an export does not work.

## Requirements

None.

## Problems

| Message or symptom | Cause | Fix |
|---|---|---|
| Refresh does nothing and the status bar asks you to open Options. | The site code or the provider is empty. | Set the connection. |
| `ConfigurationManager module not found`. | The console is not installed on this computer, or `SMS_ADMIN_UI_PATH` does not point to the `bin\i386` folder of the console. | Install the console. If it is installed, set `SMS_ADMIN_UI_PATH` to its `bin\i386` folder and start the tool again. |
| The WQL tree is empty. | No full refresh has run. | Select **Refresh** on the Collections view, or use **Open by ID**. |
| **Edit Membership...** or **Remove...** opens nothing. | The collection is built-in. | Select a custom collection. |
| The Site Queries list is empty. | The list was not refreshed, or the site saves no query. | Select **Refresh** on the Site Queries tab. |
| The settings are lost. | The folder is not writable. | Move the tool to a writable folder. |

## Steps

1. Read the log pane. It names the cause of most problems.
2. Read the session log in `Logs`.
3. Do the fix of the table.

## Limits

None.

## Links

- [Error messages](reference/error-messages.md)
- [Files and logs](reference/files-and-logs.md)
