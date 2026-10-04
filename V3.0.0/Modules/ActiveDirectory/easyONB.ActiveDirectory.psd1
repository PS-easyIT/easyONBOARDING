@{
    RootModule           = 'easyONB.ActiveDirectory.psm1'
    ModuleVersion        = '3.0.2'
    GUID                 = 'bf9bba6c-b0bb-4d68-974b-4ee73b375fec'
    Author               = 'Andreas Hepp'
    CompanyName          = 'PHINIT.DE'
    Copyright            = '(c) Andreas Hepp'
    Description          = 'easyONBOARDING AD-Adapter: Verbindung, Suche, Konflikt- und Schutzprüfung, Schreiboperationen mit ShouldProcess.'
    PowerShellVersion    = '7.2'
    CompatiblePSEditions = @('Core')
    FunctionsToExport    = @(
        'Add-EobAdGroupMembership'
        'ConvertTo-EobAdErrorMessage'
        'ConvertTo-EobAdUserInfo'
        'Disable-EobAdUserAccount'
        'Enable-EobAdUserAccount'
        'Find-EobAdUser'
        'Get-EobAdConnectionInfo'
        'Get-EobAdDirectReport'
        'Get-EobAdGroup'
        'Get-EobAdGroupProtection'
        'Get-EobAdIdentityConflict'
        'Get-EobAdOrganizationalUnitList'
        'Get-EobAdPasswordPolicy'
        'Get-EobAdPrivilegedMembership'
        'Get-EobAdStatus'
        'Get-EobAdTransitiveGroup'
        'Get-EobAdUser'
        'Get-EobAdUserGroupMembership'
        'Get-EobAdUserSnapshot'
        'Initialize-EobAdConnection'
        'Move-EobAdUserAccount'
        'New-EobAdUserAccount'
        'Remove-EobAdGroupMembership'
        'Remove-EobAdUserAccount'
        'Reset-EobAdUserPassword'
        'Resolve-EobAdUser'
        'Resolve-EobAdUserReference'
        'Set-EobAdUserAttribute'
        'Set-EobAdUserExpiration'
        'Set-EobAdUserLogonHours'
        'Set-EobAdUserManager'
        'Set-EobAdUserPasswordChangeRequired'
        'Test-EobAdModuleAvailable'
        'Test-EobAdNameInContainer'
        'Test-EobAdOrganizationalUnit'
        'Test-EobAdRecycleBinEnabled'
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
