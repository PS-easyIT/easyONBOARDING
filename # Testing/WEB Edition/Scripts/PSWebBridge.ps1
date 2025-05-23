<#
.SYNOPSIS
    PowerShell Web Bridge for easyONBOARDING
.DESCRIPTION
    This script serves as a bridge between the web API and the easyONBOARDING PowerShell backend.
    It handles API requests from the web interface and routes them to the appropriate PowerShell functionality.
.NOTES
    File Name      : PSWebBridge.ps1
    Author         : easyONBOARDING team
    Prerequisite   : PowerShell 5.1 or later
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)]
    [string]$Action,
    
    [Parameter(Mandatory=$false)]
    [string]$Data,
    
    [Parameter(Mandatory=$false)]
    [string]$RequestId,
    
    [Parameter(Mandatory=$false)]
    [string]$User,
    
    [Parameter(Mandatory=$false)]
    [string]$Role,
    
    [Parameter(Mandatory=$false)]
    [string]$OutputPath
)

# Initialize error handling
$ErrorActionPreference = "Stop"
$global:result = $null
$global:error_message = $null
$global:error_code = 0

# Configure logging
$logFolder = Join-Path -Path $PSScriptRoot -ChildPath "..\App_Data\Logs"
$logFile = Join-Path -Path $logFolder -ChildPath "PSWebBridge_$(Get-Date -Format 'yyyyMMdd').log"

# Create log folder if it doesn't exist
if (-not (Test-Path -Path $logFolder)) {
    New-Item -ItemType Directory -Path $logFolder -Force | Out-Null
}

# Function to write log entries
function Write-Log {
    param(
        [string]$Message,
        [string]$Level = "INFO"
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] [$Level] $Message"
    
    try {
        Add-Content -Path $logFile -Value $logEntry -ErrorAction SilentlyContinue
    }
    catch {
        # If writing to log fails, output to console
        Write-Host $logEntry
    }
}

# Function to return results as JSON
function Write-JsonResponse {
    param(
        [Parameter(Mandatory=$true)]
        [PSObject]$ResponseObject
    )
    
    try {
        $json = ConvertTo-Json -InputObject $ResponseObject -Depth 10 -Compress
        
        if (-not [string]::IsNullOrEmpty($OutputPath)) {
            Set-Content -Path $OutputPath -Value $json -Encoding UTF8
            return $true
        }
        else {
            Write-Output $json
            return $true
        }
    }
    catch {
        Write-Log "Error converting response to JSON: $_" -Level "ERROR"
        return $false
    }
}

# Log the request
Write-Log "Received request: Action=$Action, RequestId=$RequestId, User=$User, Role=$Role"

try {
    # Path to the main easyONBOARDING script
    $mainScriptPath = Join-Path -Path $PSScriptRoot -ChildPath "easyONB_HR-AL_V0.1.1.ps1"
    
    if (-not (Test-Path -Path $mainScriptPath)) {
        throw "Main script not found at: $mainScriptPath"
    }
    
    # Parse JSON data if provided
    $dataObject = $null
    if (-not [string]::IsNullOrEmpty($Data)) {
        try {
            $dataObject = ConvertFrom-Json -InputObject $Data
        }
        catch {
            throw "Invalid JSON data: $_"
        }
    }
    
    # Process the request based on action
    switch ($Action) {
        "GetRecords" {
            # Get filtering parameters from data
            $filterState = if ($dataObject.state) { $dataObject.state } else { "" }
            $searchText = if ($dataObject.search) { $dataObject.search } else { "" }
            $department = if ($dataObject.department) { $dataObject.department } else { "" }
            $fromDate = if ($dataObject.fromDate) { [DateTime]::Parse($dataObject.fromDate) } else { $null }
            $toDate = if ($dataObject.toDate) { [DateTime]::Parse($dataObject.toDate) } else { $null }
            
            # Create script block for execution
            $scriptBlock = {
                param($mainScriptPath, $User, $Role, $filterState, $searchText, $department, $fromDate, $toDate)
                
                # Import the script as a module
                Import-Module $mainScriptPath -Force
                
                # Set global variables needed by the script
                $global:currentUserName = $User
                $global:userRole = $Role
                
                # Call the function from the script to get records
                Get-UserRelevantRecords -FilterState $filterState -SearchText $searchText -Department $department -FromDate $fromDate -ToDate $toDate
            }
            
            # Create a new PowerShell instance to run the script
            $ps = [PowerShell]::Create()
            $ps.AddScript($scriptBlock).AddParameters(@{
                mainScriptPath = $mainScriptPath
                User = $User
                Role = $Role
                filterState = $filterState
                searchText = $searchText
                department = $department
                fromDate = $fromDate
                toDate = $toDate
            }) | Out-Null
            
            # Execute and get the results
            $global:result = $ps.Invoke()
            
            # Check for errors
            if ($ps.HadErrors) {
                $errorMsg = $ps.Streams.Error[0].Exception.Message
                throw $errorMsg
            }
            
            Write-Log "Retrieved $($global:result.Count) records for user $User with role $Role"
        }
        
        "GetRecordById" {
            if ($null -eq $dataObject -or [string]::IsNullOrEmpty($dataObject.recordId)) {
                throw "Record ID is required"
            }
            
            $recordId = $dataObject.recordId
            
            $scriptBlock = {
                param($mainScriptPath, $User, $Role, $recordId)
                
                # Import the script as a module
                Import-Module $mainScriptPath -Force
                
                # Set global variables needed by the script
                $global:currentUserName = $User
                $global:userRole = $Role
                
                # Get all records and find by ID
                $allRecords = Get-OnboardingRecords
                $record = $allRecords | Where-Object { $_.RecordID -eq $recordId }
                
                # Check permissions
                if ($record) {
                    # Check if user has permission to view this record
                    $hasPermission = $false
                    
                    if ($Role -eq "Admin") {
                        $hasPermission = $true
                    }
                    elseif ($Role -eq "HR") {
                        $hasPermission = $true
                    }
                    elseif ($Role -eq "Manager" -and $record.assignedManager -eq $User) {
                        $hasPermission = $true
                    }
                    elseif ($Role -eq "IT" -and $record.workflowState -ge 3) {
                        $hasPermission = $true
                    }
                    
                    if ($hasPermission) {
                        return $record
                    }
                    else {
                        throw "You do not have permission to view this record"
                    }
                }
                else {
                    throw "Record not found"
                }
            }
            
            # Create a new PowerShell instance to run the script
            $ps = [PowerShell]::Create()
            $ps.AddScript($scriptBlock).AddParameters(@{
                mainScriptPath = $mainScriptPath
                User = $User
                Role = $Role
                recordId = $recordId
            }) | Out-Null
            
            # Execute and get the results
            $global:result = $ps.Invoke()
            
            if ($ps.HadErrors) {
                $errorMsg = $ps.Streams.Error[0].Exception.Message
                throw $errorMsg
            }
            
            Write-Log "Retrieved record details for ID: $recordId"
        }
        
        "CreateRecord" {
            # Validate that user has HR or Admin role
            if ($Role -ne "HR" -and $Role -ne "Admin") {
                throw "Only HR or Admin users can create records"
            }
            
            $scriptBlock = {
                param($mainScriptPath, $User, $Role, $recordData)
                
                # Import the script as a module
                Import-Module $mainScriptPath -Force
                
                # Set global variables needed by the script
                $global:currentUserName = $User
                $global:userRole = $Role
                
                # Prepare the new record
                $newRecord = [PSCustomObject]@{
                    RecordID = [Guid]::NewGuid().ToString()
                    firstName = $recordData.firstName
                    lastName = $recordData.lastName
                    external = $recordData.external -eq $true
                    externalCompany = $recordData.externalCompany
                    description = $recordData.description
                    officeRoom = $recordData.officeRoom
                    phoneNumber = $recordData.phoneNumber
                    mobileNumber = $recordData.mobileNumber
                    emailAddress = $recordData.emailAddress
                    startWorkDate = $recordData.startWorkDate
                    assignedManager = $recordData.assignedManager
                    position = ""
                    departmentField = ""
                    personalNumber = ""
                    ablaufdatum = $null
                    tl = $false
                    al = $false
                    softwareSage = $false
                    softwareGenesis = $false
                    zugangLizenzmanager = $false
                    zugangMS365 = $false
                    zugriffe = ""
                    equipment = ""
                    accountCreated = $false
                    equipmentReady = $false
                    hrVerified = $false
                    workflowState = 0
                    createdDate = Get-Date
                    createdBy = $User
                    lastUpdated = Get-Date
                    lastUpdatedBy = $User
                    hrNotes = ""
                    managerNotes = ""
                    verificationNotes = ""
                    itNotes = ""
                    adminNotes = ""
                }
                
                # Save the record
                $allRecords = Get-OnboardingRecords
                $allRecords += $newRecord
                $result = Save-OnboardingRecords -Records $allRecords
                
                if ($result) {
                    return $newRecord
                }
                else {
                    throw "Failed to save new record"
                }
            }
            
            # Create a new PowerShell instance to run the script
            $ps = [PowerShell]::Create()
            $ps.AddScript($scriptBlock).AddParameters(@{
                mainScriptPath = $mainScriptPath
                User = $User
                Role = $Role
                recordData = $dataObject
            }) | Out-Null
            
            # Execute and get the results
            $global:result = $ps.Invoke()
            
            if ($ps.HadErrors) {
                $errorMsg = $ps.Streams.Error[0].Exception.Message
                throw $errorMsg
            }
            
            Write-Log "Created new record for $($dataObject.firstName) $($dataObject.lastName)"
        }
        
        "UpdateRecord" {
            if ($null -eq $dataObject -or [string]::IsNullOrEmpty($dataObject.recordId)) {
                throw "Record ID is required"
            }
            
            $recordId = $dataObject.recordId
            
            $scriptBlock = {
                param($mainScriptPath, $User, $Role, $recordId, $updateData)
                
                # Import the script as a module
                Import-Module $mainScriptPath -Force
                
                # Set global variables needed by the script
                $global:currentUserName = $User
                $global:userRole = $Role
                
                # Get all records and find by ID
                $allRecords = Get-OnboardingRecords
                $recordIndex = 0
                $found = $false
                
                for ($i = 0; $i -lt $allRecords.Count; $i++) {
                    if ($allRecords[$i].RecordID -eq $recordId) {
                        $recordIndex = $i
                        $found = $true
                        break
                    }
                }
                
                if ($found) {
                    $record = $allRecords[$recordIndex]
                    
                    # Check if user has permission to update this record
                    $hasPermission = $false
                    
                    if ($Role -eq "Admin") {
                        $hasPermission = $true
                    }
                    elseif ($Role -eq "HR" -and ($record.workflowState -eq 0 -or $record.workflowState -eq 2)) {
                        $hasPermission = $true
                    }
                    elseif ($Role -eq "Manager" -and $record.workflowState -eq 1 -and $record.assignedManager -eq $User) {
                        $hasPermission = $true
                    }
                    elseif ($Role -eq "IT" -and $record.workflowState -eq 3) {
                        $hasPermission = $true
                    }
                    
                    if ($hasPermission) {
                        # Update fields based on role
                        switch ($Role) {
                            "HR" {
                                if ($record.workflowState -eq 0) {
                                    # Initial HR data entry
                                    $record.firstName = $updateData.firstName
                                    $record.lastName = $updateData.lastName
                                    $record.external = $updateData.external -eq $true
                                    $record.externalCompany = $updateData.externalCompany
                                    $record.description = $updateData.description
                                    $record.officeRoom = $updateData.officeRoom
                                    $record.phoneNumber = $updateData.phoneNumber
                                    $record.mobileNumber = $updateData.mobileNumber
                                    $record.emailAddress = $updateData.emailAddress
                                    $record.startWorkDate = $updateData.startWorkDate
                                    $record.assignedManager = $updateData.assignedManager
                                    $record.hrNotes = $updateData.hrNotes
                                    
                                    # Change state to waiting for manager input
                                    $record.workflowState = 1
                                }
                                elseif ($record.workflowState -eq 2) {
                                    # HR verification
                                    $record.hrVerified = $updateData.hrVerified -eq $true
                                    $record.verificationNotes = $updateData.verificationNotes
                                    
                                    if ($updateData.workflowState -ne $null) {
                                        $record.workflowState = $updateData.workflowState
                                    }
                                }
                            }
                            "Manager" {
                                $record.position = $updateData.position
                                $record.departmentField = $updateData.departmentField
                                $record.personalNumber = $updateData.personalNumber
                                $record.ablaufdatum = $updateData.ablaufdatum
                                $record.tl = $updateData.tl -eq $true
                                $record.al = $updateData.al -eq $true
                                $record.softwareSage = $updateData.softwareSage -eq $true
                                $record.softwareGenesis = $updateData.softwareGenesis -eq $true
                                $record.zugangLizenzmanager = $updateData.zugangLizenzmanager -eq $true
                                $record.zugangMS365 = $updateData.zugangMS365 -eq $true
                                $record.zugriffe = $updateData.zugriffe
                                $record.equipment = $updateData.equipment
                                $record.managerNotes = $updateData.managerNotes
                                
                                # Change state to waiting for HR verification
                                $record.workflowState = 2
                            }
                            "IT" {
                                $record.accountCreated = $updateData.accountCreated -eq $true
                                $record.equipmentReady = $updateData.equipmentReady -eq $true
                                $record.itNotes = $updateData.itNotes
                                
                                if ($updateData.workflowState -ne $null) {
                                    $record.workflowState = $updateData.workflowState
                                }
                            }
                            "Admin" {
                                # Admin can update any field
                                foreach ($prop in $updateData.PSObject.Properties) {
                                    if ($prop.Name -ne "recordId" -and $prop.Name -ne "action") {
                                        $record.($prop.Name) = $prop.Value
                                    }
                                }
                            }
                        }
                        
                        # Update common fields
                        $record.lastUpdated = Get-Date
                        $record.lastUpdatedBy = $User
                        
                        # Save changes
                        $allRecords[$recordIndex] = $record
                        $result = Save-OnboardingRecords -Records $allRecords
                        
                        if ($result) {
                            return $record
                        }
                        else {
                            throw "Failed to save record changes"
                        }
                    }
                    else {
                        throw "You do not have permission to update this record"
                    }
                }
                else {
                    throw "Record not found"
                }
            }
            
            # Create a new PowerShell instance to run the script
            $ps = [PowerShell]::Create()
            $ps.AddScript($scriptBlock).AddParameters(@{
                mainScriptPath = $mainScriptPath
                User = $User
                Role = $Role
                recordId = $recordId
                updateData = $dataObject
            }) | Out-Null
            
            # Execute and get the results
            $global:result = $ps.Invoke()
            
            if ($ps.HadErrors) {
                $errorMsg = $ps.Streams.Error[0].Exception.Message
                throw $errorMsg
            }
            
            Write-Log "Updated record ID: $recordId by user $User with role $Role"
        }
        
        "DeleteRecord" {
            # Only Admin can delete records
            if ($Role -ne "Admin") {
                throw "Only administrators can delete records"
            }
            
            if ($null -eq $dataObject -or [string]::IsNullOrEmpty($dataObject.recordId)) {
                throw "Record ID is required"
            }
            
            $recordId = $dataObject.recordId
            
            $scriptBlock = {
                param($mainScriptPath, $User, $recordId)
                
                # Import the script as a module
                Import-Module $mainScriptPath -Force
                
                # Set global variables needed by the script
                $global:currentUserName = $User
                $global:userRole = "Admin"
                
                # Get all records and filter out the one to delete
                $allRecords = Get-OnboardingRecords
                $newRecords = $allRecords | Where-Object { $_.RecordID -ne $recordId }
                
                # Check if a record was actually found and removed
                if ($newRecords.Count -eq $allRecords.Count) {
                    throw "Record not found"
                }
                
                # Save the updated records list
                $result = Save-OnboardingRecords -Records $newRecords
                
                if ($result) {
                    return @{ success = $true; message = "Record deleted successfully" }
                }
                else {
                    throw "Failed to delete record"
                }
            }
            
            # Create a new PowerShell instance to run the script
            $ps = [PowerShell]::Create()
            $ps.AddScript($scriptBlock).AddParameters(@{
                mainScriptPath = $mainScriptPath
                User = $User
                recordId = $recordId
            }) | Out-Null
            
            # Execute and get the results
            $global:result = $ps.Invoke()
            
            if ($ps.HadErrors) {
                $errorMsg = $ps.Streams.Error[0].Exception.Message
                throw $errorMsg
            }
            
            Write-Log "Deleted record ID: $recordId by admin user $User"
        }
        
        "GetManagers" {
            $scriptBlock = {
                param($mainScriptPath)
                
                # Import the script as a module
                Import-Module $mainScriptPath -Force
                
                # Get managers from the script function
                $managers = Get-ManagersList
                
                if ($managers -eq $null) {
                    return @()
                }
                
                return $managers
            }
            
            # Create a new PowerShell instance to run the script
            $ps = [PowerShell]::Create()
            $ps.AddScript($scriptBlock).AddParameters(@{
                mainScriptPath = $mainScriptPath
            }) | Out-Null
            
            # Execute and get the results
            $global:result = $ps.Invoke()
            
            if ($ps.HadErrors) {
                $errorMsg = $ps.Streams.Error[0].Exception.Message
                throw $errorMsg
            }
            
            Write-Log "Retrieved manager list"
        }
        
        "GetAuditLog" {
            # Only Admin can view audit log
            if ($Role -ne "Admin") {
                throw "Only administrators can view the audit log"
            }
            
            $scriptBlock = {
                param($mainScriptPath)
                
                # Import the script as a module
                Import-Module $mainScriptPath -Force
                
                # Get audit log from the script function
                $auditLog = Get-AuditLog
                
                if ($auditLog -eq $null) {
                    return @()
                }
                
                return $auditLog
            }
            
            # Create a new PowerShell instance to run the script
            $ps = [PowerShell]::Create()
            $ps.AddScript($scriptBlock).AddParameters(@{
                mainScriptPath = $mainScriptPath
            }) | Out-Null
            
            # Execute and get the results
            $global:result = $ps.Invoke()
            
            if ($ps.HadErrors) {
                $errorMsg = $ps.Streams.Error[0].Exception.Message
                throw $errorMsg
            }
            
            Write-Log "Retrieved audit log for admin user $User"
        }
        
        "ExportCSV" {
            # Only IT or Admin can export CSV
            if ($Role -ne "IT" -and $Role -ne "Admin") {
                throw "Only IT or administrators can export CSV data"
            }
            
            $filterState = if ($dataObject -and $dataObject.filterState) { $dataObject.filterState } else { "Completed" }
            
            $scriptBlock = {
                param($mainScriptPath, $filterState, $tempExportFile, $User, $Role)
                
                # Import the script as a module
                Import-Module $mainScriptPath -Force
                
                # Set global variables needed by the script
                $global:currentUserName = $User
                $global:userRole = $Role
                
                # Get filtered records
                $records = Get-UserRelevantRecords -FilterState $filterState
                
                if ($records -eq $null -or $records.Count -eq 0) {
                    return @{ success = $false; message = "No records found for export" }
                }
                
                # Prepare data for export
                $exportRecords = $records | ForEach-Object {
                    [PSCustomObject]@{
                        Name = "$($_.firstName) $($_.lastName)"
                        Position = $_.position
                        Department = $_.departmentField
                        OfficeRoom = $_.officeRoom
                        PhoneNumber = $_.phoneNumber
                        MobileNumber = $_.mobileNumber
                        Email = $_.emailAddress
                        StartDate = $_.startWorkDate
                        Manager = $_.assignedManager
                        PersonalNumber = $_.personalNumber
                        TeamLeader = $_.tl
                        DepartmentHead = $_.al
                        Equipment = $_.equipment
                        Software = $(if($_.softwareSage){"Sage;"})$(if($_.softwareGenesis){"Genesis;"})
                        AccessRights = $(if($_.zugangLizenzmanager){"Lizenzmanager;"})$(if($_.zugangMS365){"MS365;"})
                        AdditionalAccess = $_.zugriffe
                        CompletedDate = $(if($_.workflowState -eq 4){$_.lastUpdated}else{""})
                    }
                }
                
                # Export to CSV
                try {
                    $exportRecords | Export-Csv -Path $tempExportFile -NoTypeInformation -Encoding UTF8
                    return @{ success = $true; filePath = $tempExportFile }
                }
                catch {
                    return @{ success = $false; message = "Failed to create CSV: $_" }
                }
            }
            
            # Create temporary file for export
            $tempExportFile = Join-Path -Path $env:TEMP -ChildPath "onboarding_export_$(Get-Date -Format 'yyyyMMddHHmmss').csv"
            
            # Create a new PowerShell instance to run the script
            $ps = [PowerShell]::Create()
            $ps.AddScript($scriptBlock).AddParameters(@{
                mainScriptPath = $mainScriptPath
                filterState = $filterState
                tempExportFile = $tempExportFile
                User = $User
                Role = $Role
            }) | Out-Null
            
            # Execute and get the results
            $exportResult = $ps.Invoke()
            
            if ($ps.HadErrors) {
                $errorMsg = $ps.Streams.Error[0].Exception.Message
                throw $errorMsg
            }
            
            if ($exportResult.success) {
                $global:result = @{
                    success = $true
                    filePath = $tempExportFile
                    downloadUrl = "/api/download?file=" + [System.IO.Path]::GetFileName($tempExportFile)
                }
                Write-Log "CSV export created successfully at $tempExportFile"
            } else {
                throw $exportResult.message
            }
        }
        
        default {
            throw "Unknown action: $Action"
        }
    }
    
    # Prepare successful response
    $response = @{
        success = $true
        requestId = $RequestId
        data = $global:result
    }
    
    Write-JsonResponse -ResponseObject $response
    Write-Log "Request processed successfully: $Action"
}
catch {
    $global:error_message = $_.Exception.Message
    $global:error_code = 500
    
    # Log the error
    Write-Log "Error processing request: $($_.Exception.Message)" -Level "ERROR"
    if ($_.ScriptStackTrace) {
        Write-Log "Stack trace: $($_.ScriptStackTrace)" -Level "ERROR"
    }
    
    # Prepare error response
    $errorResponse = @{
        success = $false
        requestId = $RequestId
        error = @{
            message = $global:error_message
            code = $global:error_code
        }
    }
    
    Write-JsonResponse -ResponseObject $errorResponse
}
