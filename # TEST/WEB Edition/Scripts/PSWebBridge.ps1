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

# SIG # Begin signature block
# MIIbywYJKoZIhvcNAQcCoIIbvDCCG7gCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCoGAmlCLsfFQBw
# e/QQSSvdpkw5qkdoluq88dUSrr28aaCCFhcwggMQMIIB+KADAgECAhB3jzsyX9Cg
# jEi+sBC2rBMTMA0GCSqGSIb3DQEBCwUAMCAxHjAcBgNVBAMMFVBoaW5JVC1QU3Nj
# cmlwdHNfU2lnbjAeFw0yNTA3MDUwODI4MTZaFw0yNzA3MDUwODM4MTZaMCAxHjAc
# BgNVBAMMFVBoaW5JVC1QU3NjcmlwdHNfU2lnbjCCASIwDQYJKoZIhvcNAQEBBQAD
# ggEPADCCAQoCggEBALmz3o//iDA5MvAndTjGX7/AvzTSACClfuUR9WYK0f6Ut2dI
# mPxn+Y9pZlLjXIpZT0H2Lvxq5aSI+aYeFtuJ8/0lULYNCVT31Bf+HxervRBKsUyi
# W9+4PH6STxo3Pl4l56UNQMcWLPNjDORWRPWHn0f99iNtjI+L4tUC/LoWSs3obzxN
# 3uTypzlaPBxis2qFSTR5SWqFdZdRkcuI5LNsJjyc/QWdTYRrfmVqp0QrvcxzCv8u
# EiVuni6jkXfiE6wz+oeI3L2iR+ywmU6CUX4tPWoS9VTtmm7AhEpasRTmrrnSg20Q
# jiBa1eH5TyLAH3TcYMxhfMbN9a2xDX5pzM65EJUCAwEAAaNGMEQwDgYDVR0PAQH/
# BAQDAgeAMBMGA1UdJQQMMAoGCCsGAQUFBwMDMB0GA1UdDgQWBBQO7XOqiE/EYi+n
# IaR6YO5M2MUuVTANBgkqhkiG9w0BAQsFAAOCAQEAjYOKIwBu1pfbdvEFFaR/uY88
# peKPk0NnvNEc3dpGdOv+Fsgbz27JPvItITFd6AKMoN1W48YjQLaU22M2jdhjGN5i
# FSobznP5KgQCDkRsuoDKiIOTiKAAknjhoBaCCEZGw8SZgKJtWzbST36Thsdd/won
# ihLsuoLxfcFnmBfrXh3rTIvTwvfujob68s0Sf5derHP/F+nphTymlg+y4VTEAijk
# g2dhy8RAsbS2JYZT7K5aEJpPXMiOLBqd7oTGfM7y5sLk2LIM4cT8hzgz3v5yPMkF
# H2MdR//K403e1EKH9MsGuGAJZddVN8ppaiESoPLoXrgnw2SY5KCmhYw1xRFdjTCC
# BY0wggR1oAMCAQICEA6bGI750C3n79tQ4ghAGFowDQYJKoZIhvcNAQEMBQAwZTEL
# MAkGA1UEBhMCVVMxFTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3
# LmRpZ2ljZXJ0LmNvbTEkMCIGA1UEAxMbRGlnaUNlcnQgQXNzdXJlZCBJRCBSb290
# IENBMB4XDTIyMDgwMTAwMDAwMFoXDTMxMTEwOTIzNTk1OVowYjELMAkGA1UEBhMC
# VVMxFTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0
# LmNvbTEhMB8GA1UEAxMYRGlnaUNlcnQgVHJ1c3RlZCBSb290IEc0MIICIjANBgkq
# hkiG9w0BAQEFAAOCAg8AMIICCgKCAgEAv+aQc2jeu+RdSjwwIjBpM+zCpyUuySE9
# 8orYWcLhKac9WKt2ms2uexuEDcQwH/MbpDgW61bGl20dq7J58soR0uRf1gU8Ug9S
# H8aeFaV+vp+pVxZZVXKvaJNwwrK6dZlqczKU0RBEEC7fgvMHhOZ0O21x4i0MG+4g
# 1ckgHWMpLc7sXk7Ik/ghYZs06wXGXuxbGrzryc/NrDRAX7F6Zu53yEioZldXn1RY
# jgwrt0+nMNlW7sp7XeOtyU9e5TXnMcvak17cjo+A2raRmECQecN4x7axxLVqGDgD
# EI3Y1DekLgV9iPWCPhCRcKtVgkEy19sEcypukQF8IUzUvK4bA3VdeGbZOjFEmjNA
# vwjXWkmkwuapoGfdpCe8oU85tRFYF/ckXEaPZPfBaYh2mHY9WV1CdoeJl2l6SPDg
# ohIbZpp0yt5LHucOY67m1O+SkjqePdwA5EUlibaaRBkrfsCUtNJhbesz2cXfSwQA
# zH0clcOP9yGyshG3u3/y1YxwLEFgqrFjGESVGnZifvaAsPvoZKYz0YkH4b235kOk
# GLimdwHhD5QMIR2yVCkliWzlDlJRR3S+Jqy2QXXeeqxfjT/JvNNBERJb5RBQ6zHF
# ynIWIgnffEx1P2PsIV/EIFFrb7GrhotPwtZFX50g/KEexcCPorF+CiaZ9eRpL5gd
# LfXZqbId5RsCAwEAAaOCATowggE2MA8GA1UdEwEB/wQFMAMBAf8wHQYDVR0OBBYE
# FOzX44LScV1kTN8uZz/nupiuHA9PMB8GA1UdIwQYMBaAFEXroq/0ksuCMS1Ri6en
# IZ3zbcgPMA4GA1UdDwEB/wQEAwIBhjB5BggrBgEFBQcBAQRtMGswJAYIKwYBBQUH
# MAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBDBggrBgEFBQcwAoY3aHR0cDov
# L2NhY2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEUm9vdENBLmNy
# dDBFBgNVHR8EPjA8MDqgOKA2hjRodHRwOi8vY3JsMy5kaWdpY2VydC5jb20vRGln
# aUNlcnRBc3N1cmVkSURSb290Q0EuY3JsMBEGA1UdIAQKMAgwBgYEVR0gADANBgkq
# hkiG9w0BAQwFAAOCAQEAcKC/Q1xV5zhfoKN0Gz22Ftf3v1cHvZqsoYcs7IVeqRq7
# IviHGmlUIu2kiHdtvRoU9BNKei8ttzjv9P+Aufih9/Jy3iS8UgPITtAq3votVs/5
# 9PesMHqai7Je1M/RQ0SbQyHrlnKhSLSZy51PpwYDE3cnRNTnf+hZqPC/Lwum6fI0
# POz3A8eHqNJMQBk1RmppVLC4oVaO7KTVPeix3P0c2PR3WlxUjG/voVA9/HYJaISf
# b8rbII01YBwCA8sgsKxYoA5AY8WYIsGyWfVVa88nq2x2zm8jLfR+cWojayL/ErhU
# LSd+2DrZ8LaHlv1b0VysGMNNn3O3AamfV6peKOK5lDCCBq4wggSWoAMCAQICEAc2
# N7ckVHzYR6z9KGYqXlswDQYJKoZIhvcNAQELBQAwYjELMAkGA1UEBhMCVVMxFTAT
# BgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNvbTEh
# MB8GA1UEAxMYRGlnaUNlcnQgVHJ1c3RlZCBSb290IEc0MB4XDTIyMDMyMzAwMDAw
# MFoXDTM3MDMyMjIzNTk1OVowYzELMAkGA1UEBhMCVVMxFzAVBgNVBAoTDkRpZ2lD
# ZXJ0LCBJbmMuMTswOQYDVQQDEzJEaWdpQ2VydCBUcnVzdGVkIEc0IFJTQTQwOTYg
# U0hBMjU2IFRpbWVTdGFtcGluZyBDQTCCAiIwDQYJKoZIhvcNAQEBBQADggIPADCC
# AgoCggIBAMaGNQZJs8E9cklRVcclA8TykTepl1Gh1tKD0Z5Mom2gsMyD+Vr2EaFE
# FUJfpIjzaPp985yJC3+dH54PMx9QEwsmc5Zt+FeoAn39Q7SE2hHxc7Gz7iuAhIoi
# GN/r2j3EF3+rGSs+QtxnjupRPfDWVtTnKC3r07G1decfBmWNlCnT2exp39mQh0YA
# e9tEQYncfGpXevA3eZ9drMvohGS0UvJ2R/dhgxndX7RUCyFobjchu0CsX7LeSn3O
# 9TkSZ+8OpWNs5KbFHc02DVzV5huowWR0QKfAcsW6Th+xtVhNef7Xj3OTrCw54qVI
# 1vCwMROpVymWJy71h6aPTnYVVSZwmCZ/oBpHIEPjQ2OAe3VuJyWQmDo4EbP29p7m
# O1vsgd4iFNmCKseSv6De4z6ic/rnH1pslPJSlRErWHRAKKtzQ87fSqEcazjFKfPK
# qpZzQmiftkaznTqj1QPgv/CiPMpC3BhIfxQ0z9JMq++bPf4OuGQq+nUoJEHtQr8F
# nGZJUlD0UfM2SU2LINIsVzV5K6jzRWC8I41Y99xh3pP+OcD5sjClTNfpmEpYPtMD
# iP6zj9NeS3YSUZPJjAw7W4oiqMEmCPkUEBIDfV8ju2TjY+Cm4T72wnSyPx4Jduyr
# XUZ14mCjWAkBKAAOhFTuzuldyF4wEr1GnrXTdrnSDmuZDNIztM2xAgMBAAGjggFd
# MIIBWTASBgNVHRMBAf8ECDAGAQH/AgEAMB0GA1UdDgQWBBS6FtltTYUvcyl2mi91
# jGogj57IbzAfBgNVHSMEGDAWgBTs1+OC0nFdZEzfLmc/57qYrhwPTzAOBgNVHQ8B
# Af8EBAMCAYYwEwYDVR0lBAwwCgYIKwYBBQUHAwgwdwYIKwYBBQUHAQEEazBpMCQG
# CCsGAQUFBzABhhhodHRwOi8vb2NzcC5kaWdpY2VydC5jb20wQQYIKwYBBQUHMAKG
# NWh0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRSb290
# RzQuY3J0MEMGA1UdHwQ8MDowOKA2oDSGMmh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNv
# bS9EaWdpQ2VydFRydXN0ZWRSb290RzQuY3JsMCAGA1UdIAQZMBcwCAYGZ4EMAQQC
# MAsGCWCGSAGG/WwHATANBgkqhkiG9w0BAQsFAAOCAgEAfVmOwJO2b5ipRCIBfmbW
# 2CFC4bAYLhBNE88wU86/GPvHUF3iSyn7cIoNqilp/GnBzx0H6T5gyNgL5Vxb122H
# +oQgJTQxZ822EpZvxFBMYh0MCIKoFr2pVs8Vc40BIiXOlWk/R3f7cnQU1/+rT4os
# equFzUNf7WC2qk+RZp4snuCKrOX9jLxkJodskr2dfNBwCnzvqLx1T7pa96kQsl3p
# /yhUifDVinF2ZdrM8HKjI/rAJ4JErpknG6skHibBt94q6/aesXmZgaNWhqsKRcnf
# xI2g55j7+6adcq/Ex8HBanHZxhOACcS2n82HhyS7T6NJuXdmkfFynOlLAlKnN36T
# U6w7HQhJD5TNOXrd/yVjmScsPT9rp/Fmw0HNT7ZAmyEhQNC3EyTN3B14OuSereU0
# cZLXJmvkOHOrpgFPvT87eK1MrfvElXvtCl8zOYdBeHo46Zzh3SP9HSjTx/no8Zhf
# +yvYfvJGnXUsHicsJttvFXseGYs2uJPU5vIXmVnKcPA3v5gA3yAWTyf7YGcWoWa6
# 3VXAOimGsJigK+2VQbc61RWYMbRiCQ8KvYHZE/6/pNHzV9m8BPqC3jLfBInwAM1d
# wvnQI38AC+R2AibZ8GV2QqYphwlHK+Z/GqSFD/yYlvZVVCsfgPrA8g4r5db7qS9E
# FUrnEw4d2zc4GqEr9u3WfPwwgga8MIIEpKADAgECAhALrma8Wrp/lYfG+ekE4zME
# MA0GCSqGSIb3DQEBCwUAMGMxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2Vy
# dCwgSW5jLjE7MDkGA1UEAxMyRGlnaUNlcnQgVHJ1c3RlZCBHNCBSU0E0MDk2IFNI
# QTI1NiBUaW1lU3RhbXBpbmcgQ0EwHhcNMjQwOTI2MDAwMDAwWhcNMzUxMTI1MjM1
# OTU5WjBCMQswCQYDVQQGEwJVUzERMA8GA1UEChMIRGlnaUNlcnQxIDAeBgNVBAMT
# F0RpZ2lDZXJ0IFRpbWVzdGFtcCAyMDI0MIICIjANBgkqhkiG9w0BAQEFAAOCAg8A
# MIICCgKCAgEAvmpzn/aVIauWMLpbbeZZo7Xo/ZEfGMSIO2qZ46XB/QowIEMSvgjE
# dEZ3v4vrrTHleW1JWGErrjOL0J4L0HqVR1czSzvUQ5xF7z4IQmn7dHY7yijvoQ7u
# jm0u6yXF2v1CrzZopykD07/9fpAT4BxpT9vJoJqAsP8YuhRvflJ9YeHjes4fduks
# THulntq9WelRWY++TFPxzZrbILRYynyEy7rS1lHQKFpXvo2GePfsMRhNf1F41nyE
# g5h7iOXv+vjX0K8RhUisfqw3TTLHj1uhS66YX2LZPxS4oaf33rp9HlfqSBePejlY
# eEdU740GKQM7SaVSH3TbBL8R6HwX9QVpGnXPlKdE4fBIn5BBFnV+KwPxRNUNK6lY
# k2y1WSKour4hJN0SMkoaNV8hyyADiX1xuTxKaXN12HgR+8WulU2d6zhzXomJ2Ple
# I9V2yfmfXSPGYanGgxzqI+ShoOGLomMd3mJt92nm7Mheng/TBeSA2z4I78JpwGpT
# RHiT7yHqBiV2ngUIyCtd0pZ8zg3S7bk4QC4RrcnKJ3FbjyPAGogmoiZ33c1HG93V
# p6lJ415ERcC7bFQMRbxqrMVANiav1k425zYyFMyLNyE1QulQSgDpW9rtvVcIH7Wv
# G9sqYup9j8z9J1XqbBZPJ5XLln8mS8wWmdDLnBHXgYly/p1DhoQo5fkCAwEAAaOC
# AYswggGHMA4GA1UdDwEB/wQEAwIHgDAMBgNVHRMBAf8EAjAAMBYGA1UdJQEB/wQM
# MAoGCCsGAQUFBwMIMCAGA1UdIAQZMBcwCAYGZ4EMAQQCMAsGCWCGSAGG/WwHATAf
# BgNVHSMEGDAWgBS6FtltTYUvcyl2mi91jGogj57IbzAdBgNVHQ4EFgQUn1csA3cO
# KBWQZqVjXu5Pkh92oFswWgYDVR0fBFMwUTBPoE2gS4ZJaHR0cDovL2NybDMuZGln
# aWNlcnQuY29tL0RpZ2lDZXJ0VHJ1c3RlZEc0UlNBNDA5NlNIQTI1NlRpbWVTdGFt
# cGluZ0NBLmNybDCBkAYIKwYBBQUHAQEEgYMwgYAwJAYIKwYBBQUHMAGGGGh0dHA6
# Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBYBggrBgEFBQcwAoZMaHR0cDovL2NhY2VydHMu
# ZGlnaWNlcnQuY29tL0RpZ2lDZXJ0VHJ1c3RlZEc0UlNBNDA5NlNIQTI1NlRpbWVT
# dGFtcGluZ0NBLmNydDANBgkqhkiG9w0BAQsFAAOCAgEAPa0eH3aZW+M4hBJH2UOR
# 9hHbm04IHdEoT8/T3HuBSyZeq3jSi5GXeWP7xCKhVireKCnCs+8GZl2uVYFvQe+p
# PTScVJeCZSsMo1JCoZN2mMew/L4tpqVNbSpWO9QGFwfMEy60HofN6V51sMLMXNTL
# fhVqs+e8haupWiArSozyAmGH/6oMQAh078qRh6wvJNU6gnh5OruCP1QUAvVSu4kq
# VOcJVozZR5RRb/zPd++PGE3qF1P3xWvYViUJLsxtvge/mzA75oBfFZSbdakHJe2B
# VDGIGVNVjOp8sNt70+kEoMF+T6tptMUNlehSR7vM+C13v9+9ZOUKzfRUAYSyyEmY
# tsnpltD/GWX8eM70ls1V6QG/ZOB6b6Yum1HvIiulqJ1Elesj5TMHq8CWT/xrW7tw
# ipXTJ5/i5pkU5E16RSBAdOp12aw8IQhhA/vEbFkEiF2abhuFixUDobZaA0VhqAsM
# HOmaT3XThZDNi5U2zHKhUs5uHHdG6BoQau75KiNbh0c+hatSF+02kULkftARjsyE
# pHKsF7u5zKRbt5oK5YGwFvgc4pEVUNytmB3BpIiowOIIuDgP5M9WArHYSAR16gc0
# dP2XdkMEP5eBsX7bf/MGN4K3HP50v/01ZHo/Z5lGLvNwQ7XHBx1yomzLP8lx4Q1z
# ZKDyHcp4VQJLu2kWTsKsOqQxggUKMIIFBgIBATA0MCAxHjAcBgNVBAMMFVBoaW5J
# VC1QU3NjcmlwdHNfU2lnbgIQd487Ml/QoIxIvrAQtqwTEzANBglghkgBZQMEAgEF
# AKCBhDAYBgorBgEEAYI3AgEMMQowCKACgAChAoAAMBkGCSqGSIb3DQEJAzEMBgor
# BgEEAYI3AgEEMBwGCisGAQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3
# DQEJBDEiBCAGpTousZI//WpZwknPR+SRf81r9ug5CpM8AUPYZ9JyKzANBgkqhkiG
# 9w0BAQEFAASCAQCAiho8DXoSBMc1ppXTMvTCxJ3vnmkB9ndDqIoIRya5jrg13tgt
# cTu33Xnb2lk7bF+NlPVE7dQfSKwjsxu59RRpdFRtYlu9kBwDDULQ1sq1/+CYTuhV
# i4JmnFvEjd97eQgR64UibApuocYDUX3JAkI6RnJsazH8Y8jI5oWqr4mQFeeMguJj
# LpP2UQ8seLv+dZoASLcs8mXfUOj9TTbtw3T/34UMvYbHEoVnyZFt1MTN/seEtI3X
# SPMSdzf+kKXO7gW+PBc324Rl/JyXCzArVh7OdpdQ+roMuZXrY7UpdLg7aZ62OXkE
# N9VR/D+5qcO6+/pTCM6PQV/Fn1zGcIjuTngooYIDIDCCAxwGCSqGSIb3DQEJBjGC
# Aw0wggMJAgEBMHcwYzELMAkGA1UEBhMCVVMxFzAVBgNVBAoTDkRpZ2lDZXJ0LCBJ
# bmMuMTswOQYDVQQDEzJEaWdpQ2VydCBUcnVzdGVkIEc0IFJTQTQwOTYgU0hBMjU2
# IFRpbWVTdGFtcGluZyBDQQIQC65mvFq6f5WHxvnpBOMzBDANBglghkgBZQMEAgEF
# AKBpMBgGCSqGSIb3DQEJAzELBgkqhkiG9w0BBwEwHAYJKoZIhvcNAQkFMQ8XDTI1
# MDcwNTEwMTIwNVowLwYJKoZIhvcNAQkEMSIEIEbRue1k9MDa/UUwni1LpEh/4haS
# 3TIfBGWM4NgXwg9cMA0GCSqGSIb3DQEBAQUABIICAGa76dwqafu6ZgqpzT/DzFEp
# Op89+OR8FWybJXyoTQYhFce1Wl4YpNTPTgE1lX2pKUF8otsd2olsOto/9c88WLJB
# w28bdB011uIYhD4KCAqDrg2qii+eTsoxUfJMzjJrrQp+EwhgPSa9ZNoZtW98TwVQ
# 8iRj0VxKOQlyAFLSiGcD2EdmaiTqEs+npWlYoU74zp9oYnxtQpvgJxjVIK2OULNY
# bP1Lj2kUGHYdWS7lzh2BYiMLS5eRLWJ4vjX5jkCGSGb4r1gu9IGh/wgRhRW8Mb+0
# cboZoTt5TSD3EtiBKlEEvtEH2vDRWu+f+4D5nY2m4iEu/w5otUyKbcf7CeLG2cEJ
# VSBhuE7d20bXMoqz8x8Lo3GpDcWK+HMt3+lXGNK+6xiqKtJBSa0RRvyZemhrtqDX
# 4mBPzkZkQ3h9ikjlCX0S+Upw/X48t1Yk9w99dM/Q3PI/BBkIJWRx3xMwOkWZ9+rA
# GILsIHsURLPxItFyTCmpxPVsKzGzjUBYFI10uTgK6Tb8EkrpkNLria4VHpcGQvXr
# Y0kGxskjvVEA7RAZf05cTRB3HqyLAUysKmaWln/QvwknC4SBBnLLx2RjYxi/dJrS
# y7+iVGgG518mEekS/lLszcvijtxRxFG3Ly2VbcdBU9hEGKb5PlsImGVFJHVyBpI6
# OE/OkUHF+6xOkiFBkj/Y
# SIG # End signature block
