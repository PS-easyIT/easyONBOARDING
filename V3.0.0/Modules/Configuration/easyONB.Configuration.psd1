@{
    RootModule           = 'easyONB.Configuration.psm1'
    ModuleVersion        = '3.0.0'
    GUID                 = 'a129d9d1-804e-4cf1-878a-474f9de996d8'
    Author               = 'Andreas Hepp'
    CompanyName          = 'PHINIT.DE'
    Copyright            = '(c) Andreas Hepp'
    Description          = 'easyONBOARDING Konfiguration: INI-Parser, Schema, Validierung, Companies, Migration.'
    PowerShellVersion    = '7.2'
    CompatiblePSEditions = @('Core')
    FunctionsToExport    = @(
        'Convert-EobLegacyConfiguration'
        'Copy-EobConfigurationTemplate'
        'Get-EobCompany'
        'Get-EobConfigFindingSummary'
        'Get-EobConfigSchema'
        'Get-EobConfigSection'
        'Get-EobConfigSectionName'
        'Get-EobConfigValue'
        'Get-EobDefaultConfigurationPath'
        'Get-EobLoggingParameter'
        'Get-EobMailDomain'
        'Get-EobSchemaKeyDefinition'
        'Get-EobSchemaSectionDefinition'
        'Get-EobTextFileContent'
        'Import-EobConfiguration'
        'Read-EobIniFile'
        'Set-EobIniValue'
        'Test-EobConfiguration'
        'Test-EobConfigValueType'
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
