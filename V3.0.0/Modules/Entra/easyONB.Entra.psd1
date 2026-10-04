@{
    RootModule           = 'easyONB.Entra.psm1'
    ModuleVersion        = '3.0.2'
    GUID                 = '8ac71686-aa9b-4026-b152-1db10aaa4659'
    Author               = 'Andreas Hepp'
    CompanyName          = 'PHINIT.DE'
    Copyright            = '(c) Andreas Hepp'
    Description          = 'easyONBOARDING Entra-Adapter (vorbereitet): Microsoft Graph (Sitzungen widerrufen) und Entra Connect Sync.'
    PowerShellVersion    = '7.2'
    CompatiblePSEditions = @('Core')
    FunctionsToExport    = @(
        'Connect-EobGraph'
        'Disconnect-EobGraph'
        'Get-EobEntraConnectStatus'
        'Get-EobGraphStatus'
        'Revoke-EobEntraUserSession'
        'Start-EobEntraConnectSync'
        'Test-EobGraphAppOnlyConfigured'
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
