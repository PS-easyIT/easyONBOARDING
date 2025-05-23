<#
.SYNOPSIS
    Exports onboarding data to CSV for ONBOARDING Tool import
.DESCRIPTION
    This script exports completed onboarding records to a CSV file formatted
    for import into the easyONBOARDING desktop application.
.NOTES
    File Name      : ExportCSV.ps1
    Author         : easyONBOARDING team
    Prerequisite   : PowerShell 5.1 or later
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$false)]
    [string]$OutputPath,
    
    [Parameter(Mandatory=$false)]
    [string]$FilterState = "Completed",
    
    [Parameter(Mandatory=$false)]
    [string]$MainScriptPath
)

# Set default script path if not provided
if (-not $MainScriptPath) {
    $MainScriptPath = Join-Path -Path $PSScriptRoot -ChildPath "easyONB_HR-AL_V0.1.1.ps1"
}

# Set default output path if not provided
if (-not $OutputPath) {
    $OutputPath = Join-Path -Path $env:TEMP -ChildPath "onboarding_export_$(Get-Date -Format 'yyyyMMddHHmmss').csv"
}

function Write-Log {
    param(
        [string]$Message,
        [string]$Level = "INFO"
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] [$Level] $Message"
    
    Write-Host $logEntry
}

function Export-OnboardingData {
    param(
        [string]$OutputFile,
        [string]$FilterState
    )
    
    try {
        # Import the main script to access its functions
        Import-Module $MainScriptPath -Force
        
        # Get records based on filter state
        $records = Get-OnboardingRecords | Where-Object { 
            if ($FilterState -eq "Completed") {
                $_.WorkflowState -eq 4
            } else {
                $true
            }
        }
        
        if (-not $records -or $records.Count -eq 0) {
            Write-Log "No records found matching filter criteria" -Level "WARNING"
            return $false
        }
        
        # Format records for export
        $exportRecords = $records | ForEach-Object {
            # Create a new object with the fields needed for the desktop app
            [PSCustomObject]@{
                'Vorname' = $_.FirstName
                'Nachname' = $_.LastName
                'Beschreibung' = $_.Description
                'Position' = $_.Position
                'Abteilung' = $_.DepartmentField
                'StartDatum' = $_.StartWorkDate
                'Personalnummer' = $_.PersonalNumber
                'Büro' = $_.OfficeRoom
                'Telefon' = $_.PhoneNumber
                'Mobil' = $_.MobileNumber
                'EMail' = $_.EmailAddress
                'TeamLeiter' = if ($_.TL) { "Ja" } else { "Nein" }
                'AbteilungsLeiter' = if ($_.AL) { "Ja" } else { "Nein" }
                'Software_Sage' = if ($_.SoftwareSage) { "Ja" } else { "Nein" }
                'Software_Genesis' = if ($_.SoftwareGenesis) { "Ja" } else { "Nein" }
                'Zugang_Lizenzmanager' = if ($_.ZugangLizenzmanager) { "Ja" } else { "Nein" }
                'Zugang_MS365' = if ($_.ZugangMS365) { "Ja" } else { "Nein" }
                'Weitere_Zugriffe' = $_.Zugriffe
                'Ausstattung' = $_.Equipment
                'Hinweis_HR' = $_.HRNotes
                'Hinweis_Manager' = $_.ManagerNotes
                'Hinweis_IT' = $_.ITNotes
                'Bearbeitet_von' = $_.LastUpdatedBy
                'Bearbeitet_am' = $_.LastUpdated
                'AccountErstellt' = if ($_.AccountCreated) { "Ja" } else { "Nein" }
                'AusstattungBereit' = if ($_.EquipmentReady) { "Ja" } else { "Nein" }
                'ExternerMitarbeiter' = if ($_.External) { "Ja" } else { "Nein" }
                'ExterneFirma' = $_.ExternalCompany
                'Ablaufdatum' = $_.Ablaufdatum
            }
        }
        
        # Export to CSV
        $exportRecords | Export-Csv -Path $OutputFile -NoTypeInformation -Encoding UTF8
        
        Write-Log "Successfully exported $($exportRecords.Count) records to $OutputFile" -Level "INFO"
        return $true
    }
    catch {
        Write-Log "Error exporting data: $_" -Level "ERROR"
        return $false
    }
}

# Main execution
try {
    $result = Export-OnboardingData -OutputFile $OutputPath -FilterState $FilterState
    
    if ($result) {
        # Return success information
        [PSCustomObject]@{
            success = $true
            filePath = $OutputPath
            recordCount = (Import-Csv -Path $OutputPath -Encoding UTF8 | Measure-Object).Count
            message = "Export completed successfully"
        } | ConvertTo-Json -Compress
    } else {
        # Return error information
        [PSCustomObject]@{
            success = $false
            message = "Export failed"
        } | ConvertTo-Json -Compress
    }
}
catch {
    # Return error information
    [PSCustomObject]@{
        success = $false
        message = "Exception: $_"
    } | ConvertTo-Json -Compress
}
