@{
    RootModule           = 'easyONB.Offboarding.psm1'
    ModuleVersion        = '3.0.0'
    GUID                 = '2211f256-7765-44d0-81ea-1d5ec3381a53'
    Author               = 'Andreas Hepp'
    CompanyName          = 'PHINIT.DE'
    Copyright            = '(c) Andreas Hepp'
    Description          = 'easyONBOARDING Offboarding: Vorlagen, Schutzprüfung, Zustandsberichte, Phasen mit Warteschlange, abgesicherte Endlöschung, CSV-Massenverarbeitung.'
    PowerShellVersion    = '7.2'
    CompatiblePSEditions = @('Core')
    FunctionsToExport    = @(
        'ConvertTo-EobOffboardingRequest'
        'Export-EobOffboardingBatchResult'
        'Export-EobOffboardingMailboxPermission'
        'Format-EobOffboardingText'
        'Get-EobDeletionConfirmationText'
        'Get-EobOffboardingDuePhase'
        'Get-EobOffboardingGroupClassification'
        'Get-EobOffboardingPhaseDate'
        'Get-EobOffboardingQueue'
        'Get-EobOffboardingTemplate'
        'Grant-EobHomeDirectoryAccess'
        'Import-EobOffboardingCsv'
        'Invoke-EobDueOffboardingPhase'
        'Invoke-EobOffboardingBatch'
        'Invoke-EobOffboardingFinalDeletion'
        'Invoke-EobOffboardingPlan'
        'Move-EobHomeDirectoryToArchive'
        'New-EobOffboardingBatch'
        'New-EobOffboardingDeletionPlan'
        'New-EobOffboardingPlan'
        'New-EobOffboardingPlanFromQueue'
        'Save-EobOffboardingQueueEntry'
        'Save-EobOffboardingSnapshot'
        'Stop-EobOffboardingQueueEntry'
        'Test-EobOffboardingRequest'
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
