@{
    RootModule           = 'easyONB.Reporting.psm1'
    ModuleVersion        = '3.0.0'
    GUID                 = '760966c1-277d-41e7-a899-4ba240f131d4'
    Author               = 'Andreas Hepp'
    CompanyName          = 'PHINIT.DE'
    Copyright            = '(c) Andreas Hepp'
    Description          = 'easyONBOARDING Berichte: Vorgangsberichte (HTML/JSON/CSV/TXT/PDF), Willkommensdokument und Welcome-Mail ohne Kennwort, Audit-Auswertung, Status SMTP/Dateiserver/PDF.'
    PowerShellVersion    = '7.2'
    CompatiblePSEditions = @('Core')
    FunctionsToExport    = @(
        'ConvertTo-EobPdf'
        'ConvertTo-EobReportHtml'
        'ConvertTo-EobReportText'
        'Expand-EobTemplate'
        'Export-EobAuditReport'
        'Export-EobPlanReport'
        'Export-EobReport'
        'Get-EobAssetDataUri'
        'Get-EobAuditSummary'
        'Get-EobFileServerStatus'
        'Get-EobPdfEngine'
        'Get-EobPdfEngineStatus'
        'Get-EobReportDirectory'
        'Get-EobReportFile'
        'Get-EobSmtpStatus'
        'Get-EobWelcomePlaceholderValue'
        'New-EobPlanReport'
        'New-EobWelcomeDocument'
        'Send-EobWelcomeMail'
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
