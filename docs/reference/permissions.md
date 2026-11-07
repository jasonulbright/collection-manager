# Permissions

This page lists the rights that Collection Manager needs.

## What it does

A refresh reads the site. The buttons that change the site use the rights of the account that started the tool.

## When to use it

Use it to set up an account, or to find out why a read or a change is refused.

## Requirements

| Action | Right |
|---|---|
| Start the tool | Read and write in the folder of the script. |
| Every read | The Configuration Manager console on the computer, and read access to the SMS Provider under your Configuration Manager role. |
| **Validate Query** and **Preview Results** | Read access. They run a query and change nothing. |
| **New Collection**, **Copy...**, **Evaluate**, **Edit Membership...**, **Remove...** | Create, modify and delete rights for device collections. |
| **Add Rule**, **Update Rule**, **Remove...** in the WQL Editor, **Apply to Collection...** | Modify rights for the device collection. |

## Steps

1. Start the tool as the account that has the Configuration Manager role.
2. Select **Refresh** to test the read access.

## Limits

- The tool does not check the rights before a change. A refused change shows in the log as a failure.
- The tool does not change a built-in collection.

## Troubleshooting

| Message or symptom | Cause | Fix |
|---|---|---|
| `Refresh failed:` | The account has no access to the provider, or the console is missing. | Check the role, the provider name and the console. |
| `<action> failed:` in the log | The action failed. | Read the exception message. Check the permissions only when the message shows an access failure. |

## Links

- [Network](network.md)
- [Collections](../core/collections.md)
