[CmdletBinding()]
param(
    [switch]$DebugOutput,
    [bool]$AccountDisabled = $true,
    
    # Web-specific parameters
    [Parameter(Mandatory=$false)]
    [string]$Action,
    
    [Parameter(Mandatory=$false)]
    [string]$RequestData,
    
    [Parameter(Mandatory=$false)]
    [string]$WebUser,
    
    [Parameter(Mandatory=$false)]
    [string]$WebRole,
    
    [Parameter(Mandatory=$false)]
    [string]$OutputFile
)

# General settings
$DebugPreference = 'SilentlyContinue' # or 'Continue' if you want to see non-terminating errors
$Debug = $PSBoundParameters.ContainsKey('DebugOutput') # Use the switch to control debugging

#############################################
# Helper functions for debugging and logging
#############################################
function Write-DebugMessage {
    param(
        [string]$Message
    )
    if ($DebugOutput -or $Debug) {
        Write-Host "[DEBUG] $Message" -ForegroundColor Cyan
    }
}

function Initialize-IISIntegration {
    # Prüfe, ob das Skript unter IIS (z. B. über APP_POOL_ID) läuft
    if ($env:APP_POOL_ID) {
        Write-DebugMessage "Running under IIS. Windows Authentication und IIS-spezifische Optionen konfigurieren..."
        # Hier können Sie weitere IIS-bezogene Einstellungen vornehmen
    }
}
Initialize-IISIntegration

function Write-Log {
    param(
        [string]$Message,
        [string]$Level = "INFO"
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"
    
    # Console output based on level
    switch ($Level) {
        "ERROR" { Write-Host $logMessage -ForegroundColor Red }
        "WARNING" { Write-Host $logMessage -ForegroundColor Yellow }
        "INFO" { Write-Host $logMessage -ForegroundColor Green }
        default { Write-Host $logMessage }
    }
    
    # For Web mode, also write to a log file
    if ($Action) {
        $logFolder = Join-Path -Path $PSScriptRoot -ChildPath "Logs"
        if (-not (Test-Path $logFolder)) {
            New-Item -ItemType Directory -Path $logFolder -Force | Out-Null
        }
        
        $logFile = Join-Path -Path $logFolder -ChildPath "easyONB_$(Get-Date -Format 'yyyyMMdd').log"
        try {
            Add-Content -Path $logFile -Value $logMessage -Encoding UTF8
        }
        catch {
            Write-Host "Failed to write to log file: $_" -ForegroundColor Red
        }
    }

    # Zusätzliche Log-Schreibung ins Windows Event Log, wenn unter IIS ausgeführt
    if ($env:APP_POOL_ID) {
        try {
            Write-EventLog -LogName Application -Source "easyONBOARDING" -EntryType $Level -EventId 1000 -Message $Message
        }
        catch {
            Write-DebugMessage "Fehler beim Schreiben ins Event Log: $($_.Exception.Message)"
        }
    }
}

Write-DebugMessage "Starting script with AccountDisabled=$AccountDisabled, Action=$Action"

#############################################
# Parameters - Set all settings here
#############################################
# CSV settings
$csvFolder = "C:\easyIT\DATA\easyONBOARDING\CSVData"
# For web mode, use a different location that is accessible to the IIS application pool
if ($Action) {
    $csvFolder = Join-Path -Path $PSScriptRoot -ChildPath "Data"
}

$csvFileName = "HROnboardingData.csv"
$backupFolder = Join-Path -Path $csvFolder -ChildPath "Backups"
$auditLogFile = Join-Path -Path $csvFolder -ChildPath "AuditLog.csv"

# Data security settings
$encryptionKey = "easyOnboardingSecureKey2023" # Simple encryption key
$maxBackups = 10 # Maximum number of backups to keep

# Create required folders
if (-not (Test-Path $csvFolder)) {
    New-Item -ItemType Directory -Path $csvFolder -Force | Out-Null
    Write-Log "Created CSV data folder: $csvFolder" -Level "INFO"
}
if (-not (Test-Path $backupFolder)) {
    New-Item -ItemType Directory -Path $backupFolder -Force | Out-Null
    Write-Log "Created backup folder: $backupFolder" -Level "INFO"
}

# Full path to the CSV file
$csvFile = Join-Path -Path $csvFolder -ChildPath $csvFileName

#############################################
# Enhanced Role-based permissions
#############################################
# Role definitions - can be combinations
$roleDefinitions = @{
    "HR" = @{
        "Users" = @("HR1", "HR2", "HRAdmin")
        "ADGroups" = @("Domain HR", "HR Department")
        "Permissions" = @("CreateRecord", "ViewAll", "Verify", "EditHRData")
    }
    "IT" = @{
        "Users" = @("ITAdmin", "ITSupport", "SysAdmin")
        "ADGroups" = @("Domain Admins", "IT Support")
        "Permissions" = @("CompleteOnboarding", "ViewITTasks", "CreateAccount", "AssignEquipment")
    }
    "Manager" = @{
        "Users" = @()  # Will be populated dynamically based on department assignments
        "ADGroups" = @("Department Managers", "Team Leaders")
        "Permissions" = @("EditTeamData", "ApproveRequest", "ViewDepartmentRecords")
    }
    "Admin" = @{
        "Users" = @("HRAdmin", "ITAdmin", "SystemAdmin")
        "ADGroups" = @("System Administrators")
        "Permissions" = @("ManageAll", "DeleteRecords", "ViewAuditLog")
    }
}

# Config file path for external role configuration
$roleConfigPath = Join-Path -Path $PSScriptRoot -ChildPath "RoleConfig.json"

# Load external role configuration if available
if (Test-Path $roleConfigPath) {
    try {
        $externalRoleConfig = Get-Content -Path $roleConfigPath -Raw | ConvertFrom-Json
        
        # Merge external configuration with default roles
        foreach ($role in $externalRoleConfig.PSObject.Properties.Name) {
            if ($roleDefinitions.ContainsKey($role)) {
                # Update existing role
                if ($externalRoleConfig.$role.Users) {
                    $roleDefinitions[$role].Users += $externalRoleConfig.$role.Users
                }
                if ($externalRoleConfig.$role.ADGroups) {
                    $roleDefinitions[$role].ADGroups += $externalRoleConfig.$role.ADGroups
                }
                if ($externalRoleConfig.$role.Permissions) {
                    $roleDefinitions[$role].Permissions += $externalRoleConfig.$role.Permissions
                }
            } else {
                # Add new role
                $roleDefinitions[$role] = @{
                    "Users" = $externalRoleConfig.$role.Users
                    "ADGroups" = $externalRoleConfig.$role.ADGroups
                    "Permissions" = $externalRoleConfig.$role.Permissions
                }
            }
        }
        Write-DebugMessage "External role configuration loaded and merged"
    } catch {
        Write-Log "Error loading external role configuration: $_" -Level "WARNING"
    }
}

# Current user information
if ($WebUser) {
    # Use provided web user if available
    $currentUserName = $WebUser
    $currentUser = $WebUser
} else {
    $currentUser = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
    $currentUserName = $currentUser.Split('\')[-1] # Extract username without domain
}

$currentDomain = if ($currentUser.Contains('\')) { $currentUser.Split('\')[0] } else { [System.Environment]::MachineName }
Write-DebugMessage "Current user: $currentUserName, Domain: $currentDomain"

# Function to check AD group membership
function Test-ADGroupMembership {
    param (
        [string[]]$GroupNames,
        [string]$Username = $currentUserName
    )
    
    Write-DebugMessage "Checking AD group membership for user $Username in groups: $($GroupNames -join ', ')"
    
    try {
        # Load Active Directory module if available
        if (-not (Get-Module -Name ActiveDirectory -ErrorAction SilentlyContinue)) {
            if (Get-Module -ListAvailable -Name ActiveDirectory) {
                Import-Module -Name ActiveDirectory -ErrorAction Stop
                Write-DebugMessage "ActiveDirectory module loaded"
            } else {
                Write-DebugMessage "ActiveDirectory module not available, using .NET methods"
                # Return false as we can't check with AD module
                return $false
            }
        }
        
        foreach ($groupName in $GroupNames) {
            $group = Get-ADGroup -Filter "Name -eq '$groupName'" -ErrorAction SilentlyContinue
            if ($group) {
                $isMember = Get-ADGroupMember -Identity $group -Recursive | Where-Object { $_.SamAccountName -eq $Username }
                if ($isMember) {
                    Write-DebugMessage "User is member of AD group: $groupName"
                    return $true
                }
            }
        }
        
        return $false
    }
    catch {
        Write-Log "Error checking AD group membership: $_" -Level "WARNING"
        
        # Fall back to .NET method
        try {
            $identity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
            foreach ($groupName in $GroupNames) {
                $groupSid = (New-Object System.Security.Principal.NTAccount($groupName)).Translate([System.Security.Principal.SecurityIdentifier])
                if ($identity.Groups -contains $groupSid) {
                    Write-DebugMessage "User is member of group: $groupName (using .NET method)"
                    return $true
                }
            }
        }
        catch {
            Write-Log "Error in fallback group membership check: $_" -Level "WARNING"
        }
        
        return $false
    }
}

# Determine user roles (can have multiple) - for web mode, use provided role
if ($WebRole) {
    $userRoles = @($WebRole)
    $userRole = $WebRole
    $userPermissions = $roleDefinitions[$WebRole].Permissions
    Write-DebugMessage "Using web-provided role: $WebRole"
} else {
    # Determine all user roles (can have multiple)
    $userRoles = @()
    $userPermissions = @()

    foreach ($roleName in $roleDefinitions.Keys) {
        $roleUsers = $roleDefinitions[$roleName].Users
        $roleADGroups = $roleDefinitions[$roleName].ADGroups
        
        # Check direct username match
        if ($roleUsers -contains $currentUserName) {
            $userRoles += $roleName
            $userPermissions += $roleDefinitions[$roleName].Permissions
            Write-DebugMessage "User assigned role '$roleName' based on username match"
        }
        # Check AD group membership
        elseif (($roleADGroups -and $roleADGroups.Count -gt 0) -and (Test-ADGroupMembership -GroupNames $roleADGroups)) {
            $userRoles += $roleName
            $userPermissions += $roleDefinitions[$roleName].Permissions
            Write-DebugMessage "User assigned role '$roleName' based on AD group membership"
        }
    }

    # Handle case where no roles were assigned
    if ($userRoles.Count -eq 0) {
        # Default to Manager role for backward compatibility
        $userRoles = @("Manager")
        $userPermissions = $roleDefinitions["Manager"].Permissions
        Write-Log "No explicit roles found for user $currentUserName. Defaulting to Manager role." -Level "WARNING"
    }

    # Primary role is the first in the list - for backward compatibility
    $userRole = $userRoles[0]

    # Make permissions unique
    $userPermissions = $userPermissions | Select-Object -Unique

    Write-Log "User roles determined: $($userRoles -join ', ')" -Level "INFO"
    Write-DebugMessage "User permissions: $($userPermissions -join ', ')"
}

# Function to check if user has a specific permission
function Test-UserPermission {
    param (
        [string]$Permission
    )
    
    $hasPermission = $userPermissions -contains $Permission
    Write-DebugMessage "Permission check: $Permission = $hasPermission"
    return $hasPermission
}

# Function to check if user has a specific role
function Test-UserRole {
    param (
        [string]$Role
    )
    
    $hasRole = $userRoles -contains $Role
    Write-DebugMessage "Role check: $Role = $hasRole"
    return $hasRole
}

# Workflow states
$workflowStates = @{
    "New" = 0                      # Initial state when HR creates record
    "PendingManagerInput" = 1      # Waiting for manager to complete their section
    "PendingHRVerification" = 2    # Waiting for HR to verify all information
    "ReadyForIT" = 3               # Ready for IT to process
    "Completed" = 4                # Onboarding process completed
}

# Data storage for in-progress record
$global:CurrentRecord = $null
$global:RecordModified = $false

#############################################
# Data Security Functions
#############################################
# Function to encrypt string data
function Protect-Data {
    param (
        [string]$Data
    )
    
    if ([string]::IsNullOrEmpty($Data)) {
        return $Data
    }
    
    try {
        $dataBytes = [System.Text.Encoding]::UTF8.GetBytes($Data)
        $keyBytes = [System.Text.Encoding]::UTF8.GetBytes($encryptionKey)
        
        # Ensure key is appropriate length by hashing if needed
        if ($keyBytes.Length -ne 32) {
            $sha = New-Object System.Security.Cryptography.SHA256Managed
            $keyBytes = $sha.ComputeHash($keyBytes)
        }
        
        # Simple XOR encryption (for demonstration - not for highly sensitive data)
        $encryptedBytes = @()
        for ($i = 0; $i -lt $dataBytes.Length; $i++) {
            $encryptedBytes += $dataBytes[$i] -bxor $keyBytes[$i % $keyBytes.Length]
        }
        
        # Convert to Base64 for storage
        $encryptedData = [Convert]::ToBase64String($encryptedBytes)
        return $encryptedData
    }
    catch {
        Write-Log "Error encrypting data: $_" -Level "ERROR"
        return $Data # Return original data if encryption fails
    }
}

# Function to decrypt string data
function Unprotect-Data {
    param (
        [string]$EncryptedData
    )
    
    if ([string]::IsNullOrEmpty($EncryptedData)) {
        return $EncryptedData
    }
    
    try {
        # Check if data is Base64 encoded
        try {
            $encryptedBytes = [Convert]::FromBase64String($EncryptedData)
        }
        catch {
            # If not Base64, return as is (probably not encrypted)
            return $EncryptedData
        }
        
        $keyBytes = [System.Text.Encoding]::UTF8.GetBytes($encryptionKey)
        
        # Ensure key is appropriate length by hashing if needed
        if ($keyBytes.Length -ne 32) {
            $sha = New-Object System.Security.Cryptography.SHA256Managed
            $keyBytes = $sha.ComputeHash($keyBytes)
        }
        
        # Simple XOR decryption
        $decryptedBytes = @()
        for ($i = 0; $i -lt $encryptedBytes.Length; $i++) {
            $decryptedBytes += $encryptedBytes[$i] -bxor $keyBytes[$i % $keyBytes.Length]
        }
        
        $decryptedData = [System.Text.Encoding]::UTF8.GetString($decryptedBytes)
        return $decryptedData
    }
    catch {
        Write-Log "Error decrypting data: $_" -Level "ERROR"
        return $EncryptedData # Return encrypted data if decryption fails
    }
}

# Function to create CSV backup
function Backup-CsvFile {
    param(
        [string]$SourcePath
    )
    
    if (-not (Test-Path $SourcePath)) {
        Write-Log "Cannot backup file that doesn't exist: $SourcePath" -Level "WARNING"
        return $false
    }
    
    # Create backup folder if it doesn't exist
    if (-not (Test-Path $backupFolder)) {
        Write-DebugMessage "Creating backup folder: $backupFolder"
        New-Item -ItemType Directory -Path $backupFolder -Force | Out-Null
    }
    
    # Generate backup filename with timestamp
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $backupFile = Join-Path -Path $backupFolder -ChildPath "HROnboardingData_$timestamp.bak"
    
    try {
        # Copy the file
        Copy-Item -Path $SourcePath -Destination $backupFile -Force
        Write-Log "Created backup: $backupFile" -Level "INFO"
        
        # Cleanup old backups if needed
        $allBackups = Get-ChildItem -Path $backupFolder -Filter "*.bak" | Sort-Object LastWriteTime -Descending
        if ($allBackups.Count -gt $maxBackups) {
            $backupsToDelete = $allBackups | Select-Object -Skip $maxBackups
            foreach ($backup in $backupsToDelete) {
                Remove-Item $backup.FullName -Force
                Write-DebugMessage "Removed old backup: $($backup.FullName)"
            }
        }
        
        return $true
    }
    catch {
        Write-Log "Error creating backup: $_" -Level "ERROR"
        return $false
    }
}

# Function to log data changes
function Write-AuditLog {
    param(
        [string]$Action,
        [string]$RecordID,
        [string]$Details
    )
    
    $auditEntry = [PSCustomObject]@{
        Timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
        User = $currentUserName
        Action = $Action
        RecordID = $RecordID
        Details = $Details
    }
    
    try {
        # Create audit log header if it doesn't exist
        if (-not (Test-Path $auditLogFile)) {
            $auditEntry | Export-Csv -Path $auditLogFile -NoTypeInformation -Encoding UTF8
        } else {
            $auditEntry | Export-Csv -Path $auditLogFile -NoTypeInformation -Append -Encoding UTF8
        }
        Write-DebugMessage "Audit logged: $Action for record $RecordID"
        return $true
    }
    catch {
        Write-Log "Error writing to audit log: $_" -Level "ERROR"
        return $false
    }
}

# Enhanced function to save records with encryption
function Save-OnboardingRecords {
    param(
        [PSCustomObject[]]$Records
    )
    
    # Backup existing file if it exists
    if (Test-Path $csvFile) {
        Backup-CsvFile -SourcePath $csvFile
    }
    
    try {
        # Create a copy of records with sensitive fields encrypted
        $encryptedRecords = $Records | ForEach-Object {
            $record = $_ | Select-Object * # Clone the record
            
            # Encrypt sensitive fields
            if ($record.PhoneNumber) { $record.PhoneNumber = Protect-Data -Data $record.PhoneNumber }
            if ($record.MobileNumber) { $record.MobileNumber = Protect-Data -Data $record.MobileNumber }
            if ($record.EmailAddress) { $record.EmailAddress = Protect-Data -Data $record.EmailAddress }
            if ($record.PersonalNumber) { $record.PersonalNumber = Protect-Data -Data $record.PersonalNumber }
            if ($record.ManagerNotes) { $record.ManagerNotes = Protect-Data -Data $record.ManagerNotes }
            if ($record.ITNotes) { $record.ITNotes = Protect-Data -Data $record.ITNotes }
            
            return $record
        }
        
        # Save the encrypted records to CSV
        $encryptedRecords | Export-Csv -Path $csvFile -NoTypeInformation -Encoding UTF8
        Write-Log "Records saved successfully with encryption" -Level "INFO"
        return $true
    }
    catch {
        Write-Log "Error saving records with encryption: $_" -Level "ERROR"
        
        # Try to restore from backup if save failed
        $latestBackup = Get-ChildItem -Path $backupFolder -Filter "*.bak" | Sort-Object LastWriteTime -Descending | Select-Object -First 1
        if ($latestBackup) {
            try {
                Copy-Item -Path $latestBackup.FullName -Destination $csvFile -Force
                Write-Log "Restored from backup: $($latestBackup.FullName)" -Level "INFO"
            }
            catch {
                Write-Log "Error restoring from backup: $_" -Level "ERROR"
            }
        }
        
        return $false
    }
}

# Function to load existing records
function Get-OnboardingRecords {
    if (-not (Test-Path $csvFile)) {
        Write-DebugMessage "No CSV file found. Returning empty array."
        return @()
    }
    
    try {
        $encryptedRecords = Import-Csv -Path $csvFile -Encoding UTF8
        
        # Decrypt sensitive fields
        $decryptedRecords = $encryptedRecords | ForEach-Object {
            $record = $_ | Select-Object * # Clone the record
            
            # Decrypt sensitive fields
            if ($record.PhoneNumber) { $record.PhoneNumber = Unprotect-Data -EncryptedData $record.PhoneNumber }
            if ($record.MobileNumber) { $record.MobileNumber = Unprotect-Data -EncryptedData $record.MobileNumber }
            if ($record.EmailAddress) { $record.EmailAddress = Unprotect-Data -EncryptedData $record.EmailAddress }
            if ($record.PersonalNumber) { $record.PersonalNumber = Unprotect-Data -EncryptedData $record.PersonalNumber }
            if ($record.ManagerNotes) { $record.ManagerNotes = Unprotect-Data -EncryptedData $record.ManagerNotes }
            if ($record.ITNotes) { $record.ITNotes = Unprotect-Data -EncryptedData $record.ITNotes }
            
            return $record
        }
        
        Write-DebugMessage "Loaded and decrypted $(($decryptedRecords | Measure-Object).Count) records from CSV."
        return $decryptedRecords
    }
    catch {
        Write-Log "Error loading records from CSV: $_" -Level "ERROR"
        return @()
    }
}

# Function to get records relevant to current user based on role
function Get-UserRelevantRecords {
    param(
        [string]$FilterState,
        [string]$SearchText = "",
        [DateTime]$FromDate = [DateTime]::MinValue,
        [DateTime]$ToDate = [DateTime]::MaxValue,
        [string]$Department = ""
    )
    
    $allRecords = Get-OnboardingRecords
    $filteredRecords = @()
    
    # Apply date filtering if specified
    if ($FromDate -ne [DateTime]::MinValue -or $ToDate -ne [DateTime]::MaxValue) {
        $allRecords = $allRecords | Where-Object {
            $recordDate = $null
            if ([DateTime]::TryParse($_.CreatedDate, [ref]$recordDate)) {
                return ($recordDate -ge $FromDate -and $recordDate -le $ToDate)
            }
            return $true # Include records with invalid dates
        }
    }
    
    # Apply text search if specified
    if (-not [string]::IsNullOrWhiteSpace($SearchText)) {
        $SearchText = $SearchText.ToLower()
        $allRecords = $allRecords | Where-Object {
            $_.FirstName -like "*$SearchText*" -or 
            $_.LastName -like "*$SearchText*" -or
            $_.Description -like "*$SearchText*" -or
            $_.AssignedManager -like "*$SearchText*"
        }
    }
    
    # Apply department filtering if specified
    if (-not [string]::IsNullOrWhiteSpace($Department)) {
        $allRecords = $allRecords | Where-Object {
            $_.DepartmentField -eq $Department
        }
    }
    
    # Filter based on user role and specific state if provided
    switch ($userRole) {
        "HR" {
            if (-not [string]::IsNullOrWhiteSpace($FilterState)) {
                # Filter HR records by specific workflow state
                switch ($FilterState) {
                    "New" { 
                        $filteredRecords = $allRecords | Where-Object { $_.WorkflowState -eq $workflowStates["New"] }
                    }
                    "PendingManagerInput" { 
                        $filteredRecords = $allRecords | Where-Object { $_.WorkflowState -eq $workflowStates["PendingManagerInput"] }
                    }
                    "PendingHRVerification" { 
                        $filteredRecords = $allRecords | Where-Object { $_.WorkflowState -eq $workflowStates["PendingHRVerification"] }
                    }
                    "ReadyForIT" { 
                        $filteredRecords = $allRecords | Where-Object { $_.WorkflowState -eq $workflowStates["ReadyForIT"] }
                    }
                    "Completed" { 
                        $filteredRecords = $allRecords | Where-Object { $_.WorkflowState -eq $workflowStates["Completed"] }
                    }
                    default {
                        $filteredRecords = $allRecords # Show all records if filter state is not recognized
                    }
                }
            } else {
                $filteredRecords = $allRecords # HR sees all records by default
            }
        }
        "IT" {
            if (-not [string]::IsNullOrWhiteSpace($FilterState)) {
                # Filter IT records by specific workflow state
                switch ($FilterState) {
                    "ReadyForIT" { 
                        $filteredRecords = $allRecords | Where-Object { $_.WorkflowState -eq $workflowStates["ReadyForIT"] }
                    }
                    "Completed" { 
                        $filteredRecords = $allRecords | Where-Object { $_.WorkflowState -eq $workflowStates["Completed"] }
                    }
                    default {
                        # Default to showing records ready for IT if state is not recognized
                        $filteredRecords = $allRecords | Where-Object { 
                            $_.WorkflowState -eq $workflowStates["ReadyForIT"] -or 
                            $_.WorkflowState -eq $workflowStates["Completed"]
                        }
                    }
                }
            } else {
                # IT sees records that are ready for IT or completed
                $filteredRecords = $allRecords | Where-Object { 
                    $_.WorkflowState -eq $workflowStates["ReadyForIT"] -or 
                    $_.WorkflowState -eq $workflowStates["Completed"]
                }
            }
        }
        "Manager" {
            if (-not [string]::IsNullOrWhiteSpace($FilterState)) {
                # Filter manager records by specific workflow state
                switch ($FilterState) {
                    "PendingManagerInput" { 
                        $filteredRecords = $allRecords | Where-Object {
                            $_.AssignedManager -eq $currentUserName -and 
                            $_.WorkflowState -eq $workflowStates["PendingManagerInput"]
                        }
                    }
                    "PendingHRVerification" { 
                        $filteredRecords = $allRecords | Where-Object {
                            $_.AssignedManager -eq $currentUserName -and 
                            $_.WorkflowState -eq $workflowStates["PendingHRVerification"]
                        }
                    }
                    "All" {
                        # Show all records assigned to this manager
                        $filteredRecords = $allRecords | Where-Object {
                            $_.AssignedManager -eq $currentUserName
                        }
                    }
                    default {
                        # Default to showing records pending manager input
                        $filteredRecords = $allRecords | Where-Object {
                            $_.AssignedManager -eq $currentUserName -and (
                                $_.WorkflowState -eq $workflowStates["PendingManagerInput"]
                            )
                        }
                    }
                }
            } else {
                # Managers see records assigned to them that are pending input
                $filteredRecords = $allRecords | Where-Object {
                    $_.AssignedManager -eq $currentUserName -and (
                        $_.WorkflowState -eq $workflowStates["PendingManagerInput"]
                    )
                }
            }
        }
        "Admin" {
            # Admin sees all records, but can filter by state if needed
            if (-not [string]::IsNullOrWhiteSpace($FilterState) -and $FilterState -ne "All") {
                $filteredRecords = $allRecords | Where-Object { 
                    $_.WorkflowState -eq $workflowStates[$FilterState] 
                }
            } else {
                $filteredRecords = $allRecords
            }
        }
        default {
            Write-Log "Unknown user role. Limited records returned." -Level "WARNING"
            $filteredRecords = @()
        }
    }
    
    Write-DebugMessage "Filtered records: Found $($filteredRecords.Count) records matching criteria"
    return $filteredRecords
}

# Function to get workflow state display name
function Get-WorkflowStateDisplayName {
    param(
        [int]$StateValue
    )
    
    $stateName = $workflowStates.GetEnumerator() | Where-Object { $_.Value -eq $StateValue } | Select-Object -ExpandProperty Key -First 1
    
    # Map state names to user-friendly display names
    switch ($stateName) {
        "New" { return "Neu" }
        "PendingManagerInput" { return "Warte auf Manager" }
        "PendingHRVerification" { return "Warte auf HR Prüfung" }
        "ReadyForIT" { return "Bereit für IT" }
        "Completed" { return "Abgeschlossen" }
        default { return "Unbekannt ($StateValue)" }
    }
}

# Function to update workflow state
function Update-WorkflowState {
    param(
        [PSCustomObject]$Record,
        [string]$NewState,
        [string]$AssignedTo = ""
    )
    
    $Record.WorkflowState = $workflowStates[$NewState]
    $Record.LastUpdated = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    $Record.LastUpdatedBy = $currentUserName
    
    if ($AssignedTo) {
        $Record.AssignedTo = $AssignedTo
    }
    
    # Send notifications based on new state
    switch ($NewState) {
        "PendingManagerInput" {
            # Notify manager
            $managerEmail = "$($Record.AssignedManager)@yourcompany.com"
            Send-WorkflowNotification -RecipientEmail $managerEmail -Subject "Action Required: New Onboarding Request" -Body "A new onboarding request for $($Record.FirstName) $($Record.LastName) requires your input."
        }
        "PendingHRVerification" {
            # Notify HR
            $hrEmail = "hr@yourcompany.com"
            Send-WorkflowNotification -RecipientEmail $hrEmail -Subject "Manager Input Completed: Onboarding Request" -Body "Manager $($Record.AssignedManager) has completed their input for $($Record.FirstName) $($Record.LastName)'s onboarding."
        }
        "ReadyForIT" {
            # Notify IT
            $itEmail = "it@yourcompany.com"
            Send-WorkflowNotification -RecipientEmail $itEmail -Subject "New Onboarding Request Ready for Processing" -Body "A new onboarding request for $($Record.FirstName) $($Record.LastName) is ready for IT processing."
        }
    }
    
    return $Record
}

# Notification functions
function Send-WorkflowNotification {
    param(
        [string]$RecipientEmail,
        [string]$Subject,
        [string]$Body
    )
    
    Write-DebugMessage "Sending notification to $RecipientEmail with subject: $Subject"
    
    # For now, we'll just log the notification
    # In a production environment, replace with actual email sending code
    Write-Log "NOTIFICATION - To: $RecipientEmail, Subject: $Subject, Body: $Body" -Level "INFO"
    
    # Uncomment and configure for actual email sending
    <#
    $smtpServer = "your.smtp.server"
    $smtpPort = 25
    $smtpFrom = "onboarding@yourcompany.com"
    
    $smtpClient = New-Object Net.Mail.SmtpClient($smtpServer, $smtpPort)
    $message = New-Object Net.Mail.MailMessage
    $message.From = $smtpFrom
    $message.To.Add($RecipientEmail)
    $message.Subject = $Subject
    $message.Body = $Body
    $message.IsBodyHtml = $true
    
    try {
        $smtpClient.Send($message)
        Write-DebugMessage "Email notification sent successfully"
        return $true
    }
    catch {
        Write-Log "Failed to send email notification: $_" -Level "ERROR"
        return $false
    }
    #>
    
    return $true
}

# Function to restore from a backup
function Restore-FromBackup {
    param(
        [string]$BackupFile
    )
    
    try {
        if (-not (Test-Path $BackupFile)) {
            Write-Log "Backup file not found: $BackupFile" -Level "ERROR"
            return $false
        }
        
        # First, create a backup of the current data just in case
        if (Test-Path $csvFile) {
            Backup-CsvFile -SourcePath $csvFile
        }
        
        # Copy the backup file to the main CSV file
        Copy-Item -Path $BackupFile -Destination $csvFile -Force
        
        Write-Log "Successfully restored from backup: $BackupFile" -Level "INFO"
        Write-AuditLog -Action "Restore" -RecordID "System" -Details "Data restored from backup: $BackupFile"
        
        return $true
    }
    catch {
        Write-Log "Error restoring from backup: $_" -Level "ERROR"
        return $false
    }
}

# Function to get available backups
function Get-Backups {
    try {
        if (-not (Test-Path $backupFolder)) {
            return @()
        }
        
        $backups = Get-ChildItem -Path $backupFolder -Filter "*.bak" | 
                   Sort-Object LastWriteTime -Descending | 
                   Select-Object Name, FullName, LastWriteTime
        
        return $backups
    }
    catch {
        Write-Log "Error getting backups: $_" -Level "ERROR"
        return @()
    }
}

# Function to get audit logs
function Get-AuditLogs {
    param(
        [int]$Count = 100  # Default to the last 100 logs
    )
    
    try {
        if (-not (Test-Path $auditLogFile)) {
            return @()
        }
        
        $logs = Import-Csv -Path $auditLogFile -Encoding UTF8 | 
                Select-Object -Last $Count
        
        return $logs
    }
    catch {
        Write-Log "Error getting audit logs: $_" -Level "ERROR"
        return @()
    }
}

# Function to get a specific record by ID
function Get-RecordById {
    param(
        [string]$RecordId
    )
    
    $allRecords = Get-OnboardingRecords
    $record = $allRecords | Where-Object { $_.RecordID -eq $RecordId }
    
    return $record
}

# Function to update a specific record
function Update-Record {
    param(
        [string]$RecordId,
        [hashtable]$Properties
    )
    
    $allRecords = Get-OnboardingRecords
    $recordIndex = 0
    $found = $false
    
    # Find the record
    foreach ($record in $allRecords) {
        if ($record.RecordID -eq $RecordId) {
            $found = $true
            break
        }
        $recordIndex++
    }
    
    if (-not $found) {
        Write-Log "Record with ID $RecordId not found for update" -Level "ERROR"
        return $false
    }
    
    # Update the record properties
    foreach ($key in $Properties.Keys) {
        $allRecords[$recordIndex].$key = $Properties[$key]
    }
    
    # Update last modified information
    $allRecords[$recordIndex].LastUpdated = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    $allRecords[$recordIndex].LastUpdatedBy = $currentUserName
    
    # Save the updated records
    $result = Save-OnboardingRecords -Records $allRecords
    
    if ($result) {
        Write-AuditLog -Action "Update" -RecordID $RecordId -Details "Record updated with new properties"
        return $true
    } else {
        return $false
    }
}

# Function to delete a record
function Remove-Record {
    param(
        [string]$RecordId
    )
    
    # Check for admin permission
    if (-not (Test-UserRole -Role "Admin")) {
        Write-Log "Permission denied. Only administrators can delete records." -Level "ERROR"
        return $false
    }
    
    $allRecords = Get-OnboardingRecords
    $recordToDelete = $allRecords | Where-Object { $_.RecordID -eq $RecordId }
    
    if (-not $recordToDelete) {
        Write-Log "Record with ID $RecordId not found for deletion" -Level "ERROR"
        return $false
    }
    
    # Remove the record
    $updatedRecords = $allRecords | Where-Object { $_.RecordID -ne $RecordId }
    
    # Save the updated records
    $result = Save-OnboardingRecords -Records $updatedRecords
    
    if ($result) {
        Write-AuditLog -Action "Delete" -RecordID $RecordId -Details "Record deleted for $($recordToDelete.FirstName) $($recordToDelete.LastName)"
        return $true
    } else {
        return $false
    }
}

# Function to create a new record
function New-OnboardingRecord {
    param(
        [string]$FirstName,
        [string]$LastName,
        [string]$Description = "",
        [string]$OfficeRoom = "",
        [string]$PhoneNumber = "",
        [string]$MobileNumber = "",
        [string]$EmailAddress = "",
        [bool]$External = $false,
        [string]$ExternalCompany = "",
        [string]$StartWorkDate = "",
        [string]$AssignedManager = "",
        [hashtable]$AdditionalProperties = @{}
    )
    
    # Create a new record
    $newRecord = [PSCustomObject]@{
        'RecordID'        = [Guid]::NewGuid().ToString()
        'FirstName'       = $FirstName
        'LastName'        = $LastName
        'Description'     = $Description
        'OfficeRoom'      = $OfficeRoom
        'PhoneNumber'     = $PhoneNumber
        'MobileNumber'    = $MobileNumber
        'EmailAddress'    = $EmailAddress
        'External'        = $External
        'ExternalCompany' = $ExternalCompany
        'StartWorkDate'   = $StartWorkDate
        'AssignedManager' = $AssignedManager
        'WorkflowState'   = $workflowStates["PendingManagerInput"]
        'CreatedBy'       = $currentUserName
        'CreatedDate'     = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
        'LastUpdated'     = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
        'LastUpdatedBy'   = $currentUserName
        'HRNotes'         = ""
        'AccountDisabled' = $AccountDisabled
    }
    
    # Add any additional properties
    foreach ($key in $AdditionalProperties.Keys) {
        $newRecord | Add-Member -NotePropertyName $key -NotePropertyValue $AdditionalProperties[$key]
    }
    
    # Add to records
    $allRecords = Get-OnboardingRecords
    $allRecords += $newRecord
    
    # Save records
    $result = Save-OnboardingRecords -Records $allRecords
    
    if ($result) {
        Write-AuditLog -Action "Create" -RecordID $newRecord.RecordID -Details "New onboarding record created for $FirstName $LastName"
        return $newRecord
    } else {
        return $null
    }
}

# Function to get managers from AD
function Get-Managers {
    try {
        # Demo values - in a real environment, you'd query AD
        $managers = @(
            [PSCustomObject]@{ Username = "Manager1"; DisplayName = "Michael Müller" },
            [PSCustomObject]@{ Username = "Manager2"; DisplayName = "Sarah Schmidt" },
            [PSCustomObject]@{ Username = "Manager3"; DisplayName = "Thomas Weber" }
        )
        
        return $managers
    }
    catch {
        Write-Log "Error getting managers: $_" -Level "ERROR"
        return @()
    }
}

# Function to convert object to JSON
function ConvertTo-JsonResult {
    param(
        [object]$InputObject,
        [bool]$Success = $true,
        [string]$ErrorMessage = ""
    )
    
    $result = [PSCustomObject]@{
        success = $Success
        data = $InputObject
        error = $ErrorMessage
    }
    
    try {
        $jsonResult = ConvertTo-Json -InputObject $result -Depth 10 -Compress
        return $jsonResult
    }
    catch {
        Write-Log "Error converting result to JSON: $_" -Level "ERROR"
        return "{}"
    }
}

#############################################
# Web mode specific functions
#############################################
# Process web request
function Process-WebRequest {
    param(
        [string]$Action,
        [object]$RequestData
    )
    
    Write-DebugMessage "Processing web request: $Action"
    
    try {
        switch ($Action) {
            "GetRecords" {
                $filterState = if ($RequestData.state) { $RequestData.state } else { "" }
                $searchText = if ($RequestData.search) { $RequestData.search } else { "" }
                $department = if ($RequestData.department) { $RequestData.department } else { "" }
                $fromDate = if ($RequestData.fromDate) { [DateTime]::Parse($RequestData.fromDate) } else { [DateTime]::MinValue }
                $toDate = if ($RequestData.toDate) { [DateTime]::Parse($RequestData.toDate) } else { [DateTime]::MaxValue }
                
                $records = Get-UserRelevantRecords -FilterState $filterState -SearchText $searchText -Department $department -FromDate $fromDate -ToDate $toDate
                return ConvertTo-JsonResult -InputObject $records
            }
            
            "GetRecordById" {
                $recordId = $RequestData.recordId
                
                if ([string]::IsNullOrEmpty($recordId)) {
                    return ConvertTo-JsonResult -Success $false -ErrorMessage "Record ID is required"
                }
                
                $record = Get-RecordById -RecordId $recordId
                
                if ($record) {
                    return ConvertTo-JsonResult -InputObject $record
                } else {
                    return ConvertTo-JsonResult -Success $false -ErrorMessage "Record not found"
                }
            }
            
            "CreateRecord" {
                # Validate that user has permission
                if (-not (Test-UserRole -Role "HR") -and -not (Test-UserRole -Role "Admin")) {
                    return ConvertTo-JsonResult -Success $false -ErrorMessage "Permission denied. Only HR or Admin users can create records."
                }
                
                # Create a new record
                $record = New-OnboardingRecord `
                    -FirstName $RequestData.firstName `
                    -LastName $RequestData.lastName `
                    -Description $RequestData.description `
                    -OfficeRoom $RequestData.officeRoom `
                    -PhoneNumber $RequestData.phoneNumber `
                    -MobileNumber $RequestData.mobileNumber `
                    -EmailAddress $RequestData.emailAddress `
                    -External $RequestData.external `
                    -ExternalCompany $RequestData.externalCompany `
                    -StartWorkDate $RequestData.startWorkDate `
                    -AssignedManager $RequestData.assignedManager `
                    -AdditionalProperties @{ "HRNotes" = $RequestData.hrNotes }
                
                if ($record) {
                    return ConvertTo-JsonResult -InputObject $record
                } else {
                    return ConvertTo-JsonResult -Success $false -ErrorMessage "Failed to create record"
                }
            }
            
            "UpdateRecord" {
                $recordId = $RequestData.recordId
                
                if ([string]::IsNullOrEmpty($recordId)) {
                    return ConvertTo-JsonResult -Success $false -ErrorMessage "Record ID is required"
                }
                
                $record = Get-RecordById -RecordId $recordId
                if (-not $record) {
                    return ConvertTo-JsonResult -Success $false -ErrorMessage "Record not found"
                }
                
                # Check permission to update
                $canUpdate = $false
                
                # Admin can update any record
                if (Test-UserRole -Role "Admin") {
                    $canUpdate = $true
                }
                # HR can update records in New or PendingHRVerification state
                elseif (Test-UserRole -Role "HR" -and ($record.WorkflowState -eq $workflowStates["New"] -or $record.WorkflowState -eq $workflowStates["PendingHRVerification"])) {
                    $canUpdate = $true
                }
                # Manager can update assigned records in PendingManagerInput state
                elseif (Test-UserRole -Role "Manager" -and $record.AssignedManager -eq $currentUserName -and $record.WorkflowState -eq $workflowStates["PendingManagerInput"]) {
                    $canUpdate = $true
                }
                # IT can update records in ReadyForIT state
                elseif (Test-UserRole -Role "IT" -and $record.WorkflowState -eq $workflowStates["ReadyForIT"]) {
                    $canUpdate = $true
                }
                
                if (-not $canUpdate) {
                    return ConvertTo-JsonResult -Success $false -ErrorMessage "Permission denied. You cannot update this record."
                }
                
                # Create a properties hashtable
                $properties = @{}
                
                # Process update based on role and action
                switch ($RequestData.action) {
                    "HRSubmit" {
                        if (Test-UserRole -Role "HR" -or Test-UserRole -Role "Admin") {
                            $properties["FirstName"] = $RequestData.firstName
                            $properties["LastName"] = $RequestData.lastName
                            $properties["Description"] = $RequestData.description
                            $properties["OfficeRoom"] = $RequestData.officeRoom
                            $properties["PhoneNumber"] = $RequestData.phoneNumber
                            $properties["MobileNumber"] = $RequestData.mobileNumber
                            $properties["EmailAddress"] = $RequestData.emailAddress
                            $properties["External"] = $RequestData.external
                            $properties["ExternalCompany"] = $RequestData.externalCompany
                            $properties["StartWorkDate"] = $RequestData.startWorkDate
                            $properties["AssignedManager"] = $RequestData.assignedManager
                            $properties["HRNotes"] = $RequestData.hrNotes
                            $properties["WorkflowState"] = $workflowStates["PendingManagerInput"]
                        } else {
                            return ConvertTo-JsonResult -Success $false -ErrorMessage "Permission denied. Only HR or Admin users can submit HR data."
                        }
                    }
                    
                    "HRVerify" {
                        if (Test-UserRole -Role "HR" -or Test-UserRole -Role "Admin") {
                            $properties["HRVerified"] = $true
                            $properties["VerificationNotes"] = $RequestData.verificationNotes
                            $properties["WorkflowState"] = $workflowStates["ReadyForIT"]
                        } else {
                            return ConvertTo-JsonResult -Success $false -ErrorMessage "Permission denied. Only HR or Admin users can verify records."
                        }
                    }
                    
                    "ManagerSubmit" {
                        if ((Test-UserRole -Role "Manager" -and $record.AssignedManager -eq $currentUserName) -or Test-UserRole -Role "Admin") {
                            $properties["Position"] = $RequestData.position
                            $properties["DepartmentField"] = $RequestData.departmentField
                            $properties["PersonalNumber"] = $RequestData.personalNumber
                            $properties["Ablaufdatum"] = $RequestData.ablaufdatum
                            $properties["TL"] = $RequestData.tl
                            $properties["AL"] = $RequestData.al
                            $properties["ManagerNotes"] = $RequestData.managerNotes
                            $properties["SoftwareSage"] = $RequestData.softwareSage
                            $properties["SoftwareGenesis"] = $RequestData.softwareGenesis
                            $properties["ZugangLizenzmanager"] = $RequestData.zugangLizenzmanager
                            $properties["ZugangMS365"] = $RequestData.zugangMS365
                            $properties["Zugriffe"] = $RequestData.zugriffe
                            $properties["WorkflowState"] = $workflowStates["PendingHRVerification"]
                        } else {
                            return ConvertTo-JsonResult -Success $false -ErrorMessage "Permission denied. Only the assigned manager or Admin users can submit manager data."
                        }
                    }
                    
                    "ITComplete" {
                        if (Test-UserRole -Role "IT" -or Test-UserRole -Role "Admin") {
                            $properties["AccountCreated"] = $RequestData.accountCreated
                            $properties["EquipmentReady"] = $RequestData.equipmentReady
                            $properties["ITNotes"] = $RequestData.itNotes
                            $properties["WorkflowState"] = $workflowStates["Completed"]
                            $properties["Processed"] = $true
                            $properties["ProcessedBy"] = $currentUserName
                        } else {
                            return ConvertTo-JsonResult -Success $false -ErrorMessage "Permission denied. Only IT or Admin users can complete records."
                        }
                    }
                    
                    "AdminUpdate" {
                        if (Test-UserRole -Role "Admin") {
                            $properties["WorkflowState"] = $RequestData.workflowState
                            $properties["AdminNotes"] = $RequestData.adminNotes
                        } else {
                            return ConvertTo-JsonResult -Success $false -ErrorMessage "Permission denied. Only Admin users can perform admin updates."
                        }
                    }
                    
                    default {
                        return ConvertTo-JsonResult -Success $false -ErrorMessage "Unknown action: $($RequestData.action)"
                    }
                }
                
                # Update the record
                $updateResult = Update-Record -RecordId $recordId -Properties $properties
                
                if ($updateResult) {
                    # Get the updated record
                    $updatedRecord = Get-RecordById -RecordId $recordId
                    return ConvertTo-JsonResult -InputObject $updatedRecord
                } else {
                    return ConvertTo-JsonResult -Success $false -ErrorMessage "Failed to update record"
                }
            }
            
            "DeleteRecord" {
                # Only Admin can delete records
                if (-not (Test-UserRole -Role "Admin")) {
                    return ConvertTo-JsonResult -Success $false -ErrorMessage "Permission denied. Only Admin users can delete records."
                }
                
                $recordId = $RequestData.recordId
                
                if ([string]::IsNullOrEmpty($recordId)) {
                    return ConvertTo-JsonResult -Success $false -ErrorMessage "Record ID is required"
                }
                
                $deleteResult = Remove-Record -RecordId $recordId
                
                if ($deleteResult) {
                    return ConvertTo-JsonResult -InputObject @{ deleted = $true }
                } else {
                    return ConvertTo-JsonResult -Success $false -ErrorMessage "Failed to delete record"
                }
            }
            
            "GetManagers" {
                $managers = Get-Managers
                return ConvertTo-JsonResult -InputObject $managers
            }
            
            "GetAuditLog" {
                # Only Admin can view audit log
                if (-not (Test-UserRole -Role "Admin")) {
                    return ConvertTo-JsonResult -Success $false -ErrorMessage "Permission denied. Only Admin users can view the audit log."
                }
                
                $count = if ($RequestData.count) { $RequestData.count } else { 100 }
                
                $logs = Get-AuditLogs -Count $count
                return ConvertTo-JsonResult -InputObject $logs
            }
            
            "GetBackups" {
                # Only Admin can view backups
                if (-not (Test-UserRole -Role "Admin")) {
                    return ConvertTo-JsonResult -Success $false -ErrorMessage "Permission denied. Only Admin users can view backups."
                }
                
                $backups = Get-Backups
                return ConvertTo-JsonResult -InputObject $backups
            }
            
            "RestoreBackup" {
                # Only Admin can restore backups
                if (-not (Test-UserRole -Role "Admin")) {
                    return ConvertTo-JsonResult -Success $false -ErrorMessage "Permission denied. Only Admin users can restore backups."
                }
                
                $backupFile = $RequestData.backupFile
                
                if ([string]::IsNullOrEmpty($backupFile)) {
                    return ConvertTo-JsonResult -Success $false -ErrorMessage "Backup file path is required"
                }
                
                $restoreResult = Restore-FromBackup -BackupFile $backupFile
                
                if ($restoreResult) {
                    return ConvertTo-JsonResult -InputObject @{ restored = $true }
                } else {
                    return ConvertTo-JsonResult -Success $false -ErrorMessage "Failed to restore from backup"
                }
            }
            
            default {
                return ConvertTo-JsonResult -Success $false -ErrorMessage "Unknown action: $Action"
            }
        }
    }
    catch {
        Write-Log "Error processing web request: $_" -Level "ERROR"
        return ConvertTo-JsonResult -Success $false -ErrorMessage "Internal error: $_"
    }
}

# Main execution for web mode
if ($Action) {
    Write-DebugMessage "Running in web mode with action: $Action"
    
    try {
        # Parse request data if provided
        $requestDataObj = $null
        if (-not [string]::IsNullOrEmpty($RequestData)) {
            $requestDataObj = ConvertFrom-Json -InputObject $RequestData
        }
        
        # Process the request
        $result = Process-WebRequest -Action $Action -RequestData $requestDataObj
        
        # Output the result
        if ($OutputFile) {
            $result | Out-File -FilePath $OutputFile -Encoding UTF8
        } else {
            Write-Output $result
        }
    }
    catch {
        Write-Log "Error in web mode: $_" -Level "ERROR"
        $errorResult = ConvertTo-JsonResult -Success $false -ErrorMessage "Internal server error: $_"
        
        if ($OutputFile) {
            $errorResult | Out-File -FilePath $OutputFile -Encoding UTF8
        } else {
            Write-Output $errorResult
        }
    }
    
    exit
}

# If we get here, script is being run in interactive mode for testing
Write-Log "Running in interactive mode - no web action specified" -Level "INFO"
Write-Log "Script completed successfully" -Level "INFO"

# SIG # Begin signature block
# MIIcCAYJKoZIhvcNAQcCoIIb+TCCG/UCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAggGJe5/UX+D2m
# I/QyYUDh4sUWJNdDl3qPP88yXlB+f6CCFk4wggMQMIIB+KADAgECAhB3jzsyX9Cg
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
# LSd+2DrZ8LaHlv1b0VysGMNNn3O3AamfV6peKOK5lDCCBrQwggScoAMCAQICEA3H
# rFcF/yGZLkBDIgw6SYYwDQYJKoZIhvcNAQELBQAwYjELMAkGA1UEBhMCVVMxFTAT
# BgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNvbTEh
# MB8GA1UEAxMYRGlnaUNlcnQgVHJ1c3RlZCBSb290IEc0MB4XDTI1MDUwNzAwMDAw
# MFoXDTM4MDExNDIzNTk1OVowaTELMAkGA1UEBhMCVVMxFzAVBgNVBAoTDkRpZ2lD
# ZXJ0LCBJbmMuMUEwPwYDVQQDEzhEaWdpQ2VydCBUcnVzdGVkIEc0IFRpbWVTdGFt
# cGluZyBSU0E0MDk2IFNIQTI1NiAyMDI1IENBMTCCAiIwDQYJKoZIhvcNAQEBBQAD
# ggIPADCCAgoCggIBALR4MdMKmEFyvjxGwBysddujRmh0tFEXnU2tjQ2UtZmWgyxU
# 7UNqEY81FzJsQqr5G7A6c+Gh/qm8Xi4aPCOo2N8S9SLrC6Kbltqn7SWCWgzbNfiR
# +2fkHUiljNOqnIVD/gG3SYDEAd4dg2dDGpeZGKe+42DFUF0mR/vtLa4+gKPsYfwE
# u7EEbkC9+0F2w4QJLVSTEG8yAR2CQWIM1iI5PHg62IVwxKSpO0XaF9DPfNBKS7Za
# zch8NF5vp7eaZ2CVNxpqumzTCNSOxm+SAWSuIr21Qomb+zzQWKhxKTVVgtmUPAW3
# 5xUUFREmDrMxSNlr/NsJyUXzdtFUUt4aS4CEeIY8y9IaaGBpPNXKFifinT7zL2gd
# FpBP9qh8SdLnEut/GcalNeJQ55IuwnKCgs+nrpuQNfVmUB5KlCX3ZA4x5HHKS+rq
# BvKWxdCyQEEGcbLe1b8Aw4wJkhU1JrPsFfxW1gaou30yZ46t4Y9F20HHfIY4/6vH
# espYMQmUiote8ladjS/nJ0+k6MvqzfpzPDOy5y6gqztiT96Fv/9bH7mQyogxG9QE
# PHrPV6/7umw052AkyiLA6tQbZl1KhBtTasySkuJDpsZGKdlsjg4u70EwgWbVRSX1
# Wd4+zoFpp4Ra+MlKM2baoD6x0VR4RjSpWM8o5a6D8bpfm4CLKczsG7ZrIGNTAgMB
# AAGjggFdMIIBWTASBgNVHRMBAf8ECDAGAQH/AgEAMB0GA1UdDgQWBBTvb1NK6eQG
# fHrK4pBW9i/USezLTjAfBgNVHSMEGDAWgBTs1+OC0nFdZEzfLmc/57qYrhwPTzAO
# BgNVHQ8BAf8EBAMCAYYwEwYDVR0lBAwwCgYIKwYBBQUHAwgwdwYIKwYBBQUHAQEE
# azBpMCQGCCsGAQUFBzABhhhodHRwOi8vb2NzcC5kaWdpY2VydC5jb20wQQYIKwYB
# BQUHMAKGNWh0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0
# ZWRSb290RzQuY3J0MEMGA1UdHwQ8MDowOKA2oDSGMmh0dHA6Ly9jcmwzLmRpZ2lj
# ZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRSb290RzQuY3JsMCAGA1UdIAQZMBcwCAYG
# Z4EMAQQCMAsGCWCGSAGG/WwHATANBgkqhkiG9w0BAQsFAAOCAgEAF877FoAc/gc9
# EXZxML2+C8i1NKZ/zdCHxYgaMH9Pw5tcBnPw6O6FTGNpoV2V4wzSUGvI9NAzaoQk
# 97frPBtIj+ZLzdp+yXdhOP4hCFATuNT+ReOPK0mCefSG+tXqGpYZ3essBS3q8nL2
# UwM+NMvEuBd/2vmdYxDCvwzJv2sRUoKEfJ+nN57mQfQXwcAEGCvRR2qKtntujB71
# WPYAgwPyWLKu6RnaID/B0ba2H3LUiwDRAXx1Neq9ydOal95CHfmTnM4I+ZI2rVQf
# jXQA1WSjjf4J2a7jLzWGNqNX+DF0SQzHU0pTi4dBwp9nEC8EAqoxW6q17r0z0noD
# js6+BFo+z7bKSBwZXTRNivYuve3L2oiKNqetRHdqfMTCW/NmKLJ9M+MtucVGyOxi
# Df06VXxyKkOirv6o02OoXN4bFzK0vlNMsvhlqgF2puE6FndlENSmE+9JGYxOGLS/
# D284NHNboDGcmWXfwXRy4kbu4QFhOm0xJuF2EZAOk5eCkhSxZON3rGlHqhpB/8Ml
# uDezooIs8CVnrpHMiD2wL40mm53+/j7tFaxYKIqL0Q4ssd8xHZnIn/7GELH3IdvG
# 2XlM9q7WP/UwgOkw/HQtyRN62JK4S1C8uw3PdBunvAZapsiI5YKdvlarEvf8EA+8
# hcpSM9LHJmyrxaFtoza2zNaQ9k+5t1wwggbtMIIE1aADAgECAhAKgO8YS43xBYLR
# xHanlXRoMA0GCSqGSIb3DQEBCwUAMGkxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5E
# aWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1l
# U3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYgMjAyNSBDQTEwHhcNMjUwNjA0MDAwMDAw
# WhcNMzYwOTAzMjM1OTU5WjBjMQswCQYDVQQGEwJVUzEXMBUGA1UEChMORGlnaUNl
# cnQsIEluYy4xOzA5BgNVBAMTMkRpZ2lDZXJ0IFNIQTI1NiBSU0E0MDk2IFRpbWVz
# dGFtcCBSZXNwb25kZXIgMjAyNSAxMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIIC
# CgKCAgEA0EasLRLGntDqrmBWsytXum9R/4ZwCgHfyjfMGUIwYzKomd8U1nH7C8Dr
# 0cVMF3BsfAFI54um8+dnxk36+jx0Tb+k+87H9WPxNyFPJIDZHhAqlUPt281mHrBb
# ZHqRK71Em3/hCGC5KyyneqiZ7syvFXJ9A72wzHpkBaMUNg7MOLxI6E9RaUueHTQK
# WXymOtRwJXcrcTTPPT2V1D/+cFllESviH8YjoPFvZSjKs3SKO1QNUdFd2adw44wD
# cKgH+JRJE5Qg0NP3yiSyi5MxgU6cehGHr7zou1znOM8odbkqoK+lJ25LCHBSai25
# CFyD23DZgPfDrJJJK77epTwMP6eKA0kWa3osAe8fcpK40uhktzUd/Yk0xUvhDU6l
# vJukx7jphx40DQt82yepyekl4i0r8OEps/FNO4ahfvAk12hE5FVs9HVVWcO5J4dV
# mVzix4A77p3awLbr89A90/nWGjXMGn7FQhmSlIUDy9Z2hSgctaepZTd0ILIUbWuh
# KuAeNIeWrzHKYueMJtItnj2Q+aTyLLKLM0MheP/9w6CtjuuVHJOVoIJ/DtpJRE7C
# e7vMRHoRon4CWIvuiNN1Lk9Y+xZ66lazs2kKFSTnnkrT3pXWETTJkhd76CIDBbTR
# ofOsNyEhzZtCGmnQigpFHti58CSmvEyJcAlDVcKacJ+A9/z7eacCAwEAAaOCAZUw
# ggGRMAwGA1UdEwEB/wQCMAAwHQYDVR0OBBYEFOQ7/PIx7f391/ORcWMZUEPPYYzo
# MB8GA1UdIwQYMBaAFO9vU0rp5AZ8esrikFb2L9RJ7MtOMA4GA1UdDwEB/wQEAwIH
# gDAWBgNVHSUBAf8EDDAKBggrBgEFBQcDCDCBlQYIKwYBBQUHAQEEgYgwgYUwJAYI
# KwYBBQUHMAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBdBggrBgEFBQcwAoZR
# aHR0cDovL2NhY2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0VHJ1c3RlZEc0VGlt
# ZVN0YW1waW5nUlNBNDA5NlNIQTI1NjIwMjVDQTEuY3J0MF8GA1UdHwRYMFYwVKBS
# oFCGTmh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRHNFRp
# bWVTdGFtcGluZ1JTQTQwOTZTSEEyNTYyMDI1Q0ExLmNybDAgBgNVHSAEGTAXMAgG
# BmeBDAEEAjALBglghkgBhv1sBwEwDQYJKoZIhvcNAQELBQADggIBAGUqrfEcJwS5
# rmBB7NEIRJ5jQHIh+OT2Ik/bNYulCrVvhREafBYF0RkP2AGr181o2YWPoSHz9iZE
# N/FPsLSTwVQWo2H62yGBvg7ouCODwrx6ULj6hYKqdT8wv2UV+Kbz/3ImZlJ7YXwB
# D9R0oU62PtgxOao872bOySCILdBghQ/ZLcdC8cbUUO75ZSpbh1oipOhcUT8lD8QA
# GB9lctZTTOJM3pHfKBAEcxQFoHlt2s9sXoxFizTeHihsQyfFg5fxUFEp7W42fNBV
# N4ueLaceRf9Cq9ec1v5iQMWTFQa0xNqItH3CPFTG7aEQJmmrJTV3Qhtfparz+BW6
# 0OiMEgV5GWoBy4RVPRwqxv7Mk0Sy4QHs7v9y69NBqycz0BZwhB9WOfOu/CIJnzkQ
# TwtSSpGGhLdjnQ4eBpjtP+XB3pQCtv4E5UCSDag6+iX8MmB10nfldPF9SVD7weCC
# 3yXZi/uuhqdwkgVxuiMFzGVFwYbQsiGnoa9F5AaAyBjFBtXVLcKtapnMG3VH3EmA
# p/jsJ3FVF3+d1SVDTmjFjLbNFZUWMXuZyvgLfgyPehwJVxwC+UpX2MSey2ueIu9T
# HFVkT+um1vshETaWyQo8gmBto/m3acaP9QsuLj3FNwFlTxq25+T4QwX9xa6ILs84
# ZPvmpovq90K8eWyG2N01c4IhSOxqt81nMYIFEDCCBQwCAQEwNDAgMR4wHAYDVQQD
# DBVQaGluSVQtUFNzY3JpcHRzX1NpZ24CEHePOzJf0KCMSL6wELasExMwDQYJYIZI
# AWUDBAIBBQCggYQwGAYKKwYBBAGCNwIBDDEKMAigAoAAoQKAADAZBgkqhkiG9w0B
# CQMxDAYKKwYBBAGCNwIBBDAcBgorBgEEAYI3AgELMQ4wDAYKKwYBBAGCNwIBFTAv
# BgkqhkiG9w0BCQQxIgQgsYnFo5WOZXexePvVDrpJcSdN3OBVjybp3rEULUoYW4Yw
# DQYJKoZIhvcNAQEBBQAEggEAjBqSCcC5dwbSBKu6L9nt7E8uPrhQBd7iNzeUEMzF
# vTUbyh+IOnlIod6gYxNHrc5SVrhRBqVuWezBabh2WNPdNqz7MQyNBK0Ana89WufS
# 5RtyZTrIFreus0S2qyiKDI/q45tHL35IbW7uFbX8n0Az0xhSDUllMyqaAHJbcXCl
# 4WKbMTPk9FzboXAU5auPGfppZlJLk5VRFo7zkuKNepVNcwL/g9LHsQTRwAebOz/i
# LWO9P2AOEQz5+BEFsOaJo1qpPS3vsf3wfwOW6M9ZfV2FAKWp0I6atCRLEQspJ1J+
# iX1Cvy54QK5XgoGFZIHuz00VRzjoEyZE5R4fzmSCp7UsJ6GCAyYwggMiBgkqhkiG
# 9w0BCQYxggMTMIIDDwIBATB9MGkxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdp
# Q2VydCwgSW5jLjFBMD8GA1UEAxM4RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3Rh
# bXBpbmcgUlNBNDA5NiBTSEEyNTYgMjAyNSBDQTECEAqA7xhLjfEFgtHEdqeVdGgw
# DQYJYIZIAWUDBAIBBQCgaTAYBgkqhkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqG
# SIb3DQEJBTEPFw0yNTA3MTIwODAyMzRaMC8GCSqGSIb3DQEJBDEiBCAQkL91k7FA
# +uEwdI/gueZPwFLehU1OYtJXRzTBUAXGWDANBgkqhkiG9w0BAQEFAASCAgBBVdRG
# ooSxTTxtIpEEpFACO278Z2rt9+dmiozXd+P2J29PAitw68PRGpwJ/HTSvVgJFkRr
# JeM85CLryKP6T3EOVEhMWGcnY0GMJYvm19zY6GZMGd0xgePSw9jyJd9K9HSCkb0M
# 8tENEFtVtWFkxPe3aF9n1+xv5XMgMF0pne6sSfogsPfcRP5tDu6kHqVxOgDQcKB7
# cWjL+21M0w2SzO/Li3aE+IdvxUas45rmrFg7DCRbmkaqvYJcdFLDDpUGo7izIXoP
# /FdeQ6N+o6Wmu6hxuo6VeLCWDG+oJAaTMv1+p3RUarL4GFYPUuCqI85qugLnNLjR
# MDWlit3c/JBBnA8q7QNJgDGCqiEXptyL/4vrckyUPLBzStZDMRrTVT98nMYd7xDr
# qqZngr4AuybMKH4p9UTjmj1rA2tRdHZtu466YV1EfqyoQjcPIru8kFJSO2F4HY7Q
# ubEBUnBh4n6W757A+VdDL8hssS+Zx3bI79rFOaz44U1/57k2SuJBp5V/xoHzSD2E
# /VtIsGjjGsQ5dqoTDGiDCaJayDmbqwiip1bN9zOY8mUwbOdikasDn6JhKDPeLTS3
# GyAVgxE9WqREFl/rWPJeC1WL48H8WVhEWQ7DZ/9KADHJPWTgr0r9tosFkNs6tIvp
# m4O6rNG7LjOq82Y25wFNIvwQCbgpuTG5eay+vA==
# SIG # End signature block
