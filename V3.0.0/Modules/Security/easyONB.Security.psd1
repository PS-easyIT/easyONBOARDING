@{
    RootModule           = 'easyONB.Security.psm1'
    ModuleVersion        = '3.0.0'
    GUID                 = '56a91ca1-6955-4d00-81af-a012891aec6c'
    Author               = 'Andreas Hepp'
    CompanyName          = 'PHINIT.DE'
    Copyright            = '(c) Andreas Hepp'
    Description          = 'easyONBOARDING Sicherheit: Kennwörter (CSPRNG), Richtlinien, privilegierte Gruppen, geschützte Konten, Escaping, Eingabeformate.'
    PowerShellVersion    = '7.2'
    CompatiblePSEditions = @('Core')
    FunctionsToExport    = @(
        'ConvertTo-EobAdFilterLiteral'
        'ConvertTo-EobLdapFilterValue'
        'ConvertTo-EobPlainText'
        'Get-EobPasswordPolicy'
        'Get-EobPasswordPolicyText'
        'New-EobPassword'
        'Test-EobPasswordPolicy'
        'Test-EobPersonName'
        'Test-EobPhoneNumber'
        'Test-EobPrivilegedGroup'
        'Test-EobProtectedAccount'
        'Test-EobSafeText'
        'Test-EobSamAccountName'
        'Test-EobUserPrincipalName'
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
