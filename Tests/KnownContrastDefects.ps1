# Contrast violations that exist today and are not fixed yet. Each entry is
# "<control>:<metric>" or "text:<visible text>". The window test fails when a
# violation is not listed (a new defect) and when a listed one no longer happens
# (remove it here). Run Show-ContrastViolations.ps1 to print the current list.
#
# Text must reach 4.5:1 and a control border 3:1, enabled and disabled, in both
# themes. The causes are shared by every suite tool: the SuiteCommon palette
# (sidebar button border #555555 on #1E1E1E is 2.06:1), the MahApps disabled
# style, the light text box border and the unselected tab header. The fix
# belongs in app-packager-suite, then a sync of the vendored copy, not here.

function Get-KnownContrastDefect {
    param([Parameter(Mandatory)][ValidateSet('dark', 'light')][string]$Theme)
    switch ($Theme) {
        'dark' {
            @(
                'btnAddRule:DisabledText', 'btnApplyTemplate:DisabledText', 'btnConnect:DisabledText',
                'btnCopyCollection:DisabledText', 'btnCopyToWqlEditor:DisabledText', 'btnEditMembership:DisabledText',
                'btnEvaluateColl:DisabledText', 'btnExportCsv:DisabledText', 'btnExportHtml:DisabledText',
                'btnNewCollection:DisabledText', 'btnOpenWqlCollection:DisabledText', 'btnOptions:DisabledBorder',
                'btnOptions:EnabledBorder', 'btnPreviewWql:DisabledText', 'btnRefresh:DisabledText',
                'btnRemoveColl:DisabledText', 'btnRemoveRule:DisabledText', 'btnUpdateRule:DisabledText',
                'btnValidateWql:DisabledText', 'btnViewCollections:DisabledBorder', 'btnViewCollections:EnabledBorder',
                'btnViewTemplates:DisabledBorder', 'btnViewTemplates:EnabledBorder', 'btnViewWqlEditor:DisabledBorder',
                'btnViewWqlEditor:EnabledBorder', 'text:Collections', 'text:Operational',
                'text:Properties', 'txtTemplateExpandedQuery:DisabledBorder', 'txtTemplateExpandedQuery:EnabledBorder',
                'txtWqlEditor:DisabledBorder', 'txtWqlEditor:EnabledBorder'
            )
        }
        'light' {
            @(
                'btnAddRule:DisabledText', 'btnApplyTemplate:DisabledText', 'btnConnect:DisabledText',
                'btnCopyCollection:DisabledText', 'btnCopyToWqlEditor:DisabledText', 'btnEditMembership:DisabledText',
                'btnEvaluateColl:DisabledText', 'btnExportCsv:DisabledText', 'btnExportHtml:DisabledText',
                'btnNewCollection:DisabledText', 'btnOpenWqlCollection:DisabledText', 'btnOptions:DisabledText',
                'btnPreviewWql:DisabledText', 'btnRefresh:DisabledText', 'btnRemoveColl:DisabledText',
                'btnRemoveRule:DisabledText', 'btnUpdateRule:DisabledText', 'btnValidateWql:DisabledText',
                'btnViewCollections:DisabledText', 'btnViewTemplates:DisabledText', 'btnViewWqlEditor:DisabledText',
                'cboCollectionFilter:DisabledBorder', 'cboCollectionFilter:DisabledText', 'cboCollectionFilter:EnabledBorder',
                'cboCollectionFilter:EnabledText', 'text:Direct Members', 'text:Exclude Rules',
                'text:Include Rules', 'text:Parameterized', 'text:Query Rules',
                'text:Site Queries', 'txtFilter:DisabledBorder', 'txtFilter:EnabledBorder',
                'txtRuleName:DisabledBorder', 'txtRuleName:EnabledBorder', 'txtTemplateExpandedQuery:DisabledBorder',
                'txtTemplateExpandedQuery:EnabledBorder', 'txtWqlCollectionId:DisabledBorder', 'txtWqlCollectionId:EnabledBorder',
                'txtWqlEditor:DisabledBorder', 'txtWqlEditor:EnabledBorder', 'txtWqlTreeFilter:DisabledBorder',
                'txtWqlTreeFilter:EnabledBorder'
            )
        }
    }
}