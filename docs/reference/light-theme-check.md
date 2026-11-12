# Light theme check

This page shows the light theme states of the tool. Every blue background carries white text.

## What it does

The capture tool opens each state in the light theme, reads the screen colors of the control and writes them to its log. This page lists the colors from the last run.

## When to use it

Use it after a change to a style, a dialog or the shared theme library.

## Requirements

The same as [Capture the images again](capture-images.md).

## Steps

1. Run the capture command with the `-OutDir` parameter.
2. Read the `steps-guide.run.log` file in that folder. Each `SAMPLE` line has the background color, the text color and the contrast ratio.
3. Open the images that this page shows and compare them with the colors.

## Colors read from the screen

| State | Background | Text | Contrast |
|---|---|---|---|
| grid selected row | `#3393DF` | `#FFFFFF` | 3.3 |
| inline-button normal | `#FFFFFF` | `#000000` | 21.0 |
| inline-button hover | `#E0E0E0` | `#000000` | 15.9 |
| inline-button pressed | `#000000` | `#FFFFFF` | 21.0 |
| sidebar active | `#005A9E` | `#FFFFFF` | 7.1 |
| sidebar options | `#0078D4` | `#FFFFFF` | 4.5 |
| selected tab | `#FFFFFF` | `#000000` | 21.0 |
| combo selected item | `#3393DF` | `#FFF3EA` | 3.0 |
| dialog-button normal | `#FFFFFF` | `#000000` | 21.0 |
| dialog-button hover | `#E0E0E0` | `#000000` | 15.9 |
| dialog-button pressed | `#000000` | `#FFFFFF` | 21.0 |
| dialog-accent-button | `#3393DF` | `#FFFFFF` | 3.3 |
| members grid selected row | `#3393DF` | `#FFFFFF` | 3.3 |
| collection tree selected item | `#3393DF` | `#FFFFFF` | 3.3 |
| rule list selected item | `#3393DF` | `#FFFFFF` | 3.3 |
| saved query selected row | `#3393DF` | `#FCFDFD` | 3.2 |
| template selected row | `#3393DF` | `#FCFDFD` | 3.2 |

A selected row, list item or tree item shows white text on the accent blue. A button that is pressed shows white text on black. A button that the pointer is over shows black text on light gray. The sidebar buttons show white text on blue.

## Images

![A selected collection in the grid](../images/light-collection-selected.png)

![The filter list with the highlighted item](../images/light-collections-filter-list.png)

![The New Collection dialog](../images/light-dialog-new-collection.png)

![The Copy Collection dialog](../images/light-dialog-copy-collection.png)

![The Edit Membership dialog with a selected member](../images/light-dialog-edit-membership.png)

![The Remove Collection confirmation](../images/light-dialog-remove-collection.png)

![The WQL Editor with a selected collection and rule](../images/light-wql-selected.png)

![A selected saved site query](../images/light-site-query-selected.png)

![A selected template](../images/light-template-selected.png)

![A toolbar button with the pointer over it](../images/light-state-inline-button-hover.png)

![A toolbar button that is pressed](../images/light-state-inline-button-pressed.png)

![A dialog button with the pointer over it](../images/light-state-dialog-button-hover.png)

![A dialog button that is pressed](../images/light-state-dialog-button-pressed.png)

## Limits

- The capture does not show a disabled button in this theme state, so the page has no disabled sample.
- The grid, list and tree selections show the accent blue at the contrast ratio in the table.

## Links

- [Index](../index.md)
