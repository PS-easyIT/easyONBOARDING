@{
    RootModule           = 'easyONB.Onboarding.psm1'
    ModuleVersion        = '3.0.2'
    GUID                 = 'd0a567c8-deaf-4b1e-a463-e77324ad18f3'
    Author               = 'Andreas Hepp'
    CompanyName          = 'PHINIT.DE'
    Copyright            = '(c) Andreas Hepp'
    Description          = 'easyONBOARDING Onboarding: Identitäten, Anfragen, Pläne, Benutzer-Update, Kennwort-Reset, CSV-Massenverarbeitung.'
    PowerShellVersion    = '7.2'
    CompatiblePSEditions = @('Core')
    FunctionsToExport    = @(
        'Clear-EobPlanSecret'
        'ConvertFrom-EobLegacyTemplate'
        'ConvertTo-EobAsciiName'
        'ConvertTo-EobIdentifierPart'
        'ConvertTo-EobOnboardingRequest'
        'Export-EobBatchResult'
        'Format-EobDisplayName'
        'Format-EobIdentityTemplate'
        'Get-EobDisplayNameTemplate'
        'Get-EobOnboardingGroupSet'
        'Get-EobPlanCredential'
        'Get-EobRoleTemplate'
        'Get-EobTransliterationMap'
        'Import-EobOnboardingCsv'
        'Invoke-EobOnboardingBatch'
        'Invoke-EobOnboardingPlan'
        'New-EobHomeDirectory'
        'New-EobIdentityProposal'
        'New-EobOnboardingBatch'
        'New-EobOnboardingPlan'
        'New-EobPasswordResetPlan'
        'New-EobUserUpdatePlan'
        'Resolve-EobConfiguredGroupName'
        'Resolve-EobDisplayNameTemplate'
        'Resolve-EobIdentity'
        'Set-EobUserPhoto'
        'Test-EobOnboardingRequest'
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
