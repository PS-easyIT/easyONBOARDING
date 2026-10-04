@{
    RootModule           = 'easyONB.Core.psm1'
    ModuleVersion        = '3.0.2'
    GUID                 = '0543229d-bdfc-473e-9af3-5365b00cacf5'
    Author               = 'Andreas Hepp'
    CompanyName          = 'PHINIT.DE'
    Copyright            = '(c) Andreas Hepp'
    Description          = 'easyONBOARDING Basisfunktionen: Version, Pfade, Logging/Audit mit Redaktion, Plan-Engine.'
    PowerShellVersion    = '7.2'
    CompatiblePSEditions = @('Core')
    FunctionsToExport    = @(
        'Add-EobPlanFinding'
        'Add-EobPlanStep'
        'Add-EobRedactionValue'
        'Complete-EobPlanRun'
        'ConvertTo-EobBoolean'
        'ConvertTo-EobCsvSafeObject'
        'ConvertTo-EobCsvSafeValue'
        'ConvertTo-EobDate'
        'ConvertTo-EobHtmlEncoded'
        'Get-EobAppRoot'
        'Get-EobAuditEntry'
        'Get-EobCurrentIdentity'
        'Get-EobIntegrationStatus'
        'Get-EobLogStatus'
        'Get-EobPlanOutcome'
        'Get-EobPlanStatistic'
        'Get-EobPropertyValue'
        'Get-EobSafeFileName'
        'Get-EobVersion'
        'Initialize-EobLogging'
        'Invoke-EobLogMaintenance'
        'Invoke-EobPlan'
        'Invoke-EobPlanStep'
        'New-EobFinding'
        'New-EobIntegrationStatus'
        'New-EobOperationContext'
        'New-EobPlan'
        'New-EobResult'
        'Protect-EobSensitiveData'
        'Protect-EobSensitiveText'
        'Receive-EobLogEntry'
        'Remove-EobRedactionValue'
        'Resolve-EobPath'
        'Start-EobPlanRun'
        'Stop-EobPlan'
        'Test-EobAuditWritable'
        'Test-EobBooleanText'
        'Test-EobDirectoryWritable'
        'Test-EobDistinguishedName'
        'Test-EobDomainName'
        'Test-EobEmailAddress'
        'Test-EobPathWithin'
        'Test-EobPlanExecutable'
        'Test-EobWindowsStylePath'
        'Write-EobLog'
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
