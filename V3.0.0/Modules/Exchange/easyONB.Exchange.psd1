@{
    RootModule           = 'easyONB.Exchange.psm1'
    ModuleVersion        = '3.0.2'
    GUID                 = '03957092-37b0-4f4c-b959-5756fdb5964c'
    Author               = 'Andreas Hepp'
    CompanyName          = 'PHINIT.DE'
    Copyright            = '(c) Andreas Hepp'
    Description          = 'easyONBOARDING Exchange-Adapter (vorbereitet): Status, Verbindung, Postfach, Abwesenheitsnotiz, Weiterleitung, freigegebenes Postfach, Berechtigungsdokumentation.'
    PowerShellVersion    = '7.2'
    CompatiblePSEditions = @('Core')
    FunctionsToExport    = @(
        'Connect-EobExchange'
        'ConvertTo-EobSharedMailbox'
        'Disconnect-EobExchange'
        'Enable-EobExchangeMailbox'
        'Get-EobExchangeStatus'
        'Get-EobMailboxInfo'
        'Get-EobMailboxPermissionReport'
        'Set-EobMailboxAutoReply'
        'Set-EobMailboxForwarding'
        'Set-EobMailboxHidden'
        'Test-EobExchangeAppOnlyConfigured'
        'Test-EobExchangeConnected'
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
