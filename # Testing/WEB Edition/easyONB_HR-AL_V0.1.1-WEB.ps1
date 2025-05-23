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
