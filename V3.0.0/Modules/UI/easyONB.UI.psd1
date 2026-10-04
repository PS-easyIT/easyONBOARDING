@{
    RootModule           = 'easyONB.UI.psm1'
    ModuleVersion        = '3.0.2'
    GUID                 = '5f150602-797f-463f-83a3-ea7df5dd90d2'
    Author               = 'Andreas Hepp'
    CompanyName          = 'PHINIT.DE'
    Copyright            = '(c) Andreas Hepp'
    Description          = 'easyONBOARDING Oberfläche (WPF): Navigation, Assistenten, kooperative Ausführung, zentrale Dialoge, Farbschemata.'
    PowerShellVersion    = '7.2'
    CompatiblePSEditions = @('Core')
    FunctionsToExport    = @(
        'Get-EobAccentPalette'
        'Get-EobXamlElementName'
        'Start-EobGui'
        'Test-EobGuiEnvironment'
        'Test-EobXamlContent'
    )
    CmdletsToExport      = @()
    VariablesToExport    = @()
    AliasesToExport      = @()
    PrivateData          = @{
        PSData = @{
            Tags       = @('easyONBOARDING', 'ActiveDirectory', 'Onboarding', 'Offboarding')
            ProjectUri = 'https://github.com/PS-easyIT/easyONBOARDING'
        }
    }
}
