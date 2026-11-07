# App-specific names for Window.Tests.ps1.
$script:AppConfig = @{
    Title         = 'Collection Manager'
    Shell         = 'start-collectionmanager.ps1'
    Xaml          = 'MainWindow.xaml'
    OptionsButton = 'btnOptions'
    LogLabel      = 'lblLogOutput'
    Views         = @(
        @{ Name = 'Collections'; Container = 'viewCollections'; Button = 'btnViewCollections' }
        @{ Name = 'WQL Editor';  Container = 'viewWqlEditor';   Button = 'btnViewWqlEditor' }
        @{ Name = 'Templates';   Container = 'viewTemplates';   Button = 'btnViewTemplates' }
    )
}
