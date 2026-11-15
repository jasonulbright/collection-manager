# Contrast violations that exist today and are not fixed yet. Each entry is
# "<control>:<metric>" or "text:<visible text>". The window test fails when a
# violation is not listed (a new defect) and when a listed one no longer happens
# (remove it here). Run Show-ContrastViolations.ps1 to print the current list.
#
# Enabled text must reach 4.5:1 and an enabled control border 3:1, in both themes.
# Disabled controls are exempt: WCAG 1.4.3 exempts inactive components, and the shared button
# templates dim a disabled button by opacity. The window test requires a disabled control to look
# dimmed instead. The causes of the listed defects are shared by every suite tool (the light text
# box border and the unselected tab header). The fix belongs in app-packager-suite, then a sync
# of the vendored copy, not here.

function Get-KnownContrastDefect {
    param([Parameter(Mandatory)][ValidateSet('dark', 'light')][string]$Theme)
    switch ($Theme) {
        'dark' {
            @(
                'txtTemplateExpandedQuery:EnabledBorder', 'txtWqlEditor:EnabledBorder'
            )
        }
        'light' {
            @(
                'cboCollectionFilter:EnabledText', 'txtTemplateExpandedQuery:EnabledBorder', 'txtWqlEditor:EnabledBorder'
            )
        }
    }
}
