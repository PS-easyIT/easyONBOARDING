@{
    RootModule           = 'easyONB.Setup.psm1'
    ModuleVersion        = '3.0.2'
    GUID                 = '6f0b3b1e-8f2a-4c5e-9d7b-2a1c4e5f6a70'
    Author               = 'Andreas Hepp'
    CompanyName          = 'PHINIT.DE'
    Copyright            = '(c) Andreas Hepp'
    Description          = 'easyONBOARDING Installer: Ersteinrichtung der Konfiguration (WPF), Vorbelegung auf Domänencontrollern.'
    PowerShellVersion    = '7.2'
    CompatiblePSEditions = @('Core')
    FunctionsToExport    = @(
        'ConvertFrom-EobEntraConnectAccountDescription'
        'ConvertTo-EobSetupIniText'
        'Get-EobSetupDefaultValue'
        'Get-EobSetupDirectoryDefault'
        'Get-EobSetupEnvironment'
        'Get-EobSetupExchangeServerName'
        'Get-EobSetupField'
        'Get-EobSetupIniValue'
        'Get-EobSetupPage'
        'New-EobSetupConfiguration'
        'Select-EobSetupOu'
        'Start-EobSetupGui'
        'Test-EobSetupFieldValue'
    )
    CmdletsToExport      = @()
    VariablesToExport    = @()
    AliasesToExport      = @()
    PrivateData          = @{
        PSData = @{
            Tags       = @('easyONBOARDING', 'ActiveDirectory', 'Onboarding', 'Setup')
            ProjectUri = 'https://github.com/PS-easyIT/easyONBOARDING'
        }
    }
}
