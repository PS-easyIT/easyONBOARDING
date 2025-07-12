[CmdletBinding()]
param(
    [switch]$DebugOutput,
    [bool]$AccountDisabled = $true
)

#region SCRIPT PARAMETERS AND SETTINGS
# General settings
$DebugPreference = 'SilentlyContinue' # Set to 'Continue' to display non-terminating errors
$Debug = $PSBoundParameters.ContainsKey('DebugOutput') # Use the switch to control debugging

#region Debugging and Logging Functions
# Helper functions for debugging and logging
function Write-DebugMessage {
    param(
        [string]$Message
    )
    if ($DebugOutput -or $Debug) {
        Write-Host "[DEBUG] $Message" -ForegroundColor Cyan
    }
}

function Write-Log {
    param(
        [string]$Message,
        [string]$Level = "INFO"
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"
    
    # Console output based on log severity level
    switch ($Level) {
        "ERROR" { Write-Host $logMessage -ForegroundColor Red }
        "WARNING" { Write-Host $logMessage -ForegroundColor Yellow }
        "INFO" { Write-Host $logMessage -ForegroundColor Green }
        default { Write-Host $logMessage }
    }
}
#endregion Debugging and Logging Functions

Write-DebugMessage "Starting script with AccountDisabled=$AccountDisabled"

#region CONFIGURATION PARAMETERS
# CSV settings for data storage
$csvFolder = "C:\easyIT\DATA\easyONBOARDING\CSVData"
$csvFileName = "HROnboardingData.csv"
$backupFolder = Join-Path -Path $csvFolder -ChildPath "Backups"
$auditLogFile = Join-Path -Path $csvFolder -ChildPath "AuditLog.csv"

# Data security settings
$encryptionKey = "easyOnboardingSecureKey2023" # Simple encryption key for data protection
$maxBackups = 10 # Maximum number of backup files to retain

# GUI settings
$fontSize = 10
$formBackColor = "Silver"

# Logo settings
$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
$headerLogo = Join-Path -Path $scriptPath -ChildPath "APPICON1.PNG"
#endregion CONFIGURATION PARAMETERS
#endregion SCRIPT PARAMETERS AND SETTINGS

#region ROLE-BASED PERMISSIONS SYSTEM
# Enhanced Role-based permissions with hierarchical structure
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
$roleConfigPath = Join-Path -Path $scriptPath -ChildPath "RoleConfig.json"

# Load external role configuration if available for more flexibility
if (Test-Path $roleConfigPath) {
    try {
        $externalRoleConfig = Get-Content -Path $roleConfigPath -Raw | ConvertFrom-Json
        
        # Merge external configuration with default roles
        foreach ($role in $externalRoleConfig.PSObject.Properties.Name) {
            if ($roleDefinitions.ContainsKey($role)) {
                # Update existing role with external configuration
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
                # Add completely new role from external configuration
                $roleDefinitions[$role] = @{
                    "Users" = $externalRoleConfig.$role.Users
                    "ADGroups" = $externalRoleConfig.$role.ADGroups
                    "Permissions" = $externalRoleConfig.$role.Permissions
                }
            }
        }
        Write-DebugMessage "External role configuration loaded and merged successfully"
    } catch {
        Write-Log "Error loading external role configuration: $_" -Level "WARNING"
    }
}

# Current user information for authentication and authorization
$currentUser = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
$currentUserName = $currentUser.Split('\')[-1] # Extract username without domain prefix
$currentDomain = if ($currentUser.Contains('\')) { $currentUser.Split('\')[0] } else { [System.Environment]::MachineName }
Write-DebugMessage "Current user: $currentUserName, Domain: $currentDomain"

#region Permission Functions
# Function to check Active Directory group membership for role assignment
function Test-ADGroupMembership {
    param (
        [string[]]$GroupNames
    )
    
    Write-DebugMessage "Checking AD group membership for groups: $($GroupNames -join ', ')"
    
    try {
        # Load Active Directory module if available for group membership checks
        if (-not (Get-Module -Name ActiveDirectory -ErrorAction SilentlyContinue)) {
            if (Get-Module -ListAvailable -Name ActiveDirectory) {
                Import-Module -Name ActiveDirectory -ErrorAction Stop
                Write-DebugMessage "ActiveDirectory module loaded successfully"
            } else {
                Write-DebugMessage "ActiveDirectory module not available, using .NET methods instead"
                # Return false as we can't check with AD module
                return $false
            }
        }
        
        foreach ($groupName in $GroupNames) {
            $group = Get-ADGroup -Filter "Name -eq '$groupName'" -ErrorAction SilentlyContinue
            if ($group) {
                $isMember = Get-ADGroupMember -Identity $group -Recursive | Where-Object { $_.SamAccountName -eq $currentUserName }
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
        
        # Fall back to .NET method if AD module fails
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

# Determine all user roles (user can have multiple roles)
$userRoles = @()
$userPermissions = @()

foreach ($roleName in $roleDefinitions.Keys) {
    $roleUsers = $roleDefinitions[$roleName].Users
    $roleADGroups = $roleDefinitions[$roleName].ADGroups
    
    # Check direct username match in role definition
    if ($roleUsers -contains $currentUserName) {
        $userRoles += $roleName
        $userPermissions += $roleDefinitions[$roleName].Permissions
        Write-DebugMessage "User assigned role '$roleName' based on username match"
    }
    # Check AD group membership for role assignment
    elseif (($roleADGroups -and $roleADGroups.Count -gt 0) -and (Test-ADGroupMembership -GroupNames $roleADGroups)) {
        $userRoles += $roleName
        $userPermissions += $roleDefinitions[$roleName].Permissions
        Write-DebugMessage "User assigned role '$roleName' based on AD group membership"
    }
}

# Handle case where no roles were assigned - default to Manager for backward compatibility
if ($userRoles.Count -eq 0) {
    # Default to Manager role for backward compatibility, but we'll verify later
    $userRoles = @("Manager")
    $userPermissions = $roleDefinitions["Manager"].Permissions
    Write-Log "No explicit roles found for user $currentUserName. Defaulting to Manager role pending verification." -Level "WARNING"
}

# Primary role is the first in the list - for backward compatibility
$userRole = $userRoles[0]

# Make permissions unique to avoid duplicates
$userPermissions = $userPermissions | Select-Object -Unique

Write-Log "User roles determined: $($userRoles -join ', ')" -Level "INFO"
Write-DebugMessage "User permissions: $($userPermissions -join ', ')"

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

# Function to require a permission to continue with an operation
function Require-Permission {
    param (
        [string]$Permission,
        [string]$Action = "perform this action"
    )
    
    if (-not (Test-UserPermission -Permission $Permission)) {
        $errorMsg = "Access denied: You don't have permission to $Action"
        Write-Log $errorMsg -Level "ERROR"
        [System.Windows.Forms.MessageBox]::Show(
            $errorMsg,
            "Permission Denied",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
        return $false
    }
    return $true
}
#endregion Permission Functions
#endregion ROLE-BASED PERMISSIONS SYSTEM

#region WORKFLOW DEFINITIONS
# Workflow states for the onboarding process
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
#endregion WORKFLOW DEFINITIONS

#region ASSEMBLY LOADING
# Load required assemblies for GUI and processing
try {
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    Add-Type -AssemblyName PresentationFramework
    Add-Type -AssemblyName PresentationCore
    Add-Type -AssemblyName WindowsBase
    Write-DebugMessage "Windows.Forms, Drawing, PresentationFramework, PresentationCore and WindowsBase assemblies loaded"
}
catch {
    Write-Log "Error loading assemblies: $_" -Level "ERROR"
    [System.Windows.Forms.MessageBox]::Show("Error loading assemblies: $_", "Error", 
        [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    exit
}
#endregion ASSEMBLY LOADING

#region DATA DIRECTORY SETUP
# Check if CSV folder exists, create it if needed
try {
    if (-not (Test-Path $csvFolder)) {
        Write-DebugMessage "CSV folder not found. Creating folder: $csvFolder"
        New-Item -ItemType Directory -Path $csvFolder -Force | Out-Null
        Write-Log "CSV folder created: $csvFolder" -Level "INFO"
    } else {
        Write-DebugMessage "CSV folder already exists: $csvFolder"
    }
    $csvFile = Join-Path -Path $csvFolder -ChildPath $csvFileName
    Write-DebugMessage "Full CSV path: $csvFile"
} catch {
    Write-Log "Error creating CSV folder ($csvFolder): $_" -Level "ERROR"
    [System.Windows.Forms.MessageBox]::Show("Error creating CSV folder ($csvFolder): $_", "Error", 
        [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        exit
}
#endregion DATA DIRECTORY SETUP

#region NOTIFICATION SYSTEM
# Notification functions for workflow status changes
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
    
    # Uncomment and configure for actual email sending in production environment
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
#endregion NOTIFICATION SYSTEM

#region DATA MANAGEMENT FUNCTIONS
# Function to load existing onboarding records from CSV
function Get-OnboardingRecords {
    if (-not (Test-Path $csvFile)) {
        Write-DebugMessage "No CSV file found. Returning empty array."
        return @()
    }
    
    try {
        $encryptedRecords = Import-Csv -Path $csvFile -Encoding UTF8
        
        # Decrypt sensitive fields for in-memory usage
        $decryptedRecords = $encryptedRecords | ForEach-Object {
            $record = $_ | Select-Object * # Clone the record
            
            # Decrypt sensitive fields for data processing
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

# Function to get records relevant to current user based on role and filters
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

# Function to get workflow state display name for UI presentation
function Get-WorkflowStateDisplayName {
    param(
        [int]$StateValue
    )
    
    $stateName = $workflowStates.GetEnumerator() | Where-Object { $_.Value -eq $StateValue } | Select-Object -ExpandProperty Key -First 1
    
    # Map state names to user-friendly display names
    switch ($stateName) {
        "New" { return "New" }
        "PendingManagerInput" { return "Waiting for Manager" }
        "PendingHRVerification" { return "Waiting for HR Verification" }
        "ReadyForIT" { return "Ready for IT" }
        "Completed" { return "Completed" }
        default { return "Unknown ($StateValue)" }
    }
}

# Function to refresh records list with filtering
function Update-RecordsList {
    param(
        [string]$FilterState = "",
        [string]$SearchText = "",
        [string]$Department = ""
    )
    
    Write-DebugMessage "Updating records list with filter: $FilterState, search: $SearchText, department: $Department"
    
    # Clear existing items in the list
    if ($controls.lstOnboardingRecords -ne $null) {
        $controls.lstOnboardingRecords.Items.Clear()
        
        # Get filtered records based on criteria
        $records = Get-UserRelevantRecords -FilterState $FilterState -SearchText $SearchText -Department $Department
        
        if ($records.Count -eq 0) {
            $noRecordsItem = New-Object System.Windows.Controls.ListBoxItem
            $noRecordsItem.Content = "No records found"
            $noRecordsItem.IsEnabled = $false
            $controls.lstOnboardingRecords.Items.Add($noRecordsItem)
        } else {
            # Sort records by workflow state and then by name
            $records = $records | Sort-Object -Property @{Expression={$_.WorkflowState}; Ascending=$true}, @{Expression={$_.LastName}; Ascending=$true}
            
            foreach ($record in $records) {
                # Get user-friendly state name for display
                $stateName = Get-WorkflowStateDisplayName -StateValue $record.WorkflowState
                
                # Create formatted display name with workflow state
                $displayName = "$($record.FirstName) $($record.LastName)"
                
                # Create a new ListBoxItem for UI display
                $item = New-Object System.Windows.Controls.ListBoxItem
                
                # Create a StackPanel for better content formatting
                $stackPanel = New-Object System.Windows.Controls.StackPanel
                $stackPanel.Orientation = [System.Windows.Controls.Orientation]::Horizontal
                
                # Create main text with name
                $textBlock = New-Object System.Windows.Controls.TextBlock
                $textBlock.Text = $displayName
                $textBlock.FontWeight = [System.Windows.FontWeights]::SemiBold
                $textBlock.Margin = New-Object System.Windows.Thickness(0, 0, 8, 0)
                
                # Add state indicator with appropriate color based on workflow state
                $stateBlock = New-Object System.Windows.Controls.Border
                $stateBlock.Background = switch ($record.WorkflowState) {
                    $workflowStates["New"] { [System.Windows.Media.Brushes]::LightBlue }
                    $workflowStates["PendingManagerInput"] { [System.Windows.Media.Brushes]::Gold }
                    $workflowStates["PendingHRVerification"] { [System.Windows.Media.Brushes]::Orange }
                    $workflowStates["ReadyForIT"] { [System.Windows.Media.Brushes]::LightGreen }
                    $workflowStates["Completed"] { [System.Windows.Media.Brushes]::DarkGreen }
                    default { [System.Windows.Media.Brushes]::Gray }
                }
                $stateBlock.CornerRadius = New-Object System.Windows.CornerRadius(3)
                $stateBlock.Padding = New-Object System.Windows.Thickness(5, 2, 5, 2)
                $stateText = New-Object System.Windows.Controls.TextBlock
                $stateText.Text = $stateName
                $stateText.Foreground = [System.Windows.Media.Brushes]::White
                $stateText.FontSize = 11
                $stateBlock.Child = $stateText
                
                # Add department info if available
                if (-not [string]::IsNullOrWhiteSpace($record.DepartmentField)) {
                    $deptBlock = New-Object System.Windows.Controls.Border
                    $deptBlock.Background = [System.Windows.Media.Brushes]::LightGray
                    $deptBlock.CornerRadius = New-Object System.Windows.CornerRadius(3)
                    $deptBlock.Padding = New-Object System.Windows.Thickness(5, 2, 5, 2)
                    $deptBlock.Margin = New-Object System.Windows.Thickness(6, 0, 0, 0)
                    $deptText = New-Object System.Windows.Controls.TextBlock
                    $deptText.Text = $record.DepartmentField
                    $deptText.FontSize = 11
                    $deptBlock.Child = $deptText
                }
                
                # Add date info if available
                $dateBlock = New-Object System.Windows.Controls.TextBlock
                $dateBlock.Margin = New-Object System.Windows.Thickness(8, 0, 0, 0)
                $dateBlock.FontStyle = [System.Windows.FontStyles]::Italic
                $dateBlock.FontSize = 11
                try {
                    if (-not [string]::IsNullOrWhiteSpace($record.CreatedDate)) {
                        $createdDate = [DateTime]::Parse($record.CreatedDate).ToString("MM/dd/yyyy")
                        $dateBlock.Text = "Created: $createdDate"
                    }
                } catch {
                    $dateBlock.Text = "Date unknown"
                    Write-DebugMessage "Could not parse date: $($record.CreatedDate)"
                }
                
                # Add tooltip with additional information for hover details
                $tooltipContent = "Name: $($record.FirstName) $($record.LastName)"
                if (-not [string]::IsNullOrWhiteSpace($record.Description)) {
                    $tooltipContent += "`nDescription: $($record.Description)"
                }
                if (-not [string]::IsNullOrWhiteSpace($record.Position)) {
                    $tooltipContent += "`nPosition: $($record.Position)"
                }
                if (-not [string]::IsNullOrWhiteSpace($record.StartWorkDate)) {
                    $tooltipContent += "`nStart date: $($record.StartWorkDate)"
                }
                if (-not [string]::IsNullOrWhiteSpace($record.OfficeRoom)) {
                    $tooltipContent += "`nOffice: $($record.OfficeRoom)"
                }
                if (-not [string]::IsNullOrWhiteSpace($record.AssignedManager)) {
                    $tooltipContent += "`nManager: $($record.AssignedManager)"
                }
                $item.ToolTip = $tooltipContent
                
                # Assemble the stack panel with all elements
                $stackPanel.Children.Add($textBlock)
                $stackPanel.Children.Add($stateBlock)
                if ($deptBlock) {
                    $stackPanel.Children.Add($deptBlock)
                }
                $stackPanel.Children.Add($dateBlock)
                
                # Add the panel to the list item
                $item.Content = $stackPanel
                $item.Tag = $record
                
                # Add the item to the list
                $controls.lstOnboardingRecords.Items.Add($item)
            }
            
            Write-DebugMessage "Added $($records.Count) records to the list with enhanced formatting"
        }
    } else {
        Write-DebugMessage "lstOnboardingRecords control not found"
    }
}

# Function to update workflow state with notifications
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
    
    # Send notifications based on new state to relevant parties
    switch ($NewState) {
        "PendingManagerInput" {
            # Notify manager about new onboarding request
            $managerEmail = "$($Record.AssignedManager)@yourcompany.com"
            Send-WorkflowNotification -RecipientEmail $managerEmail -Subject "Action Required: New Onboarding Request" -Body "A new onboarding request for $($Record.FirstName) $($Record.LastName) requires your input."
        }
        "PendingHRVerification" {
            # Notify HR about completed manager input
            $hrEmail = "hr@yourcompany.com"
            Send-WorkflowNotification -RecipientEmail $hrEmail -Subject "Manager Input Completed: Onboarding Request" -Body "Manager $($Record.AssignedManager) has completed their input for $($Record.FirstName) $($Record.LastName)'s onboarding record and it requires HR verification."
        }
        "ReadyForIT" {
            # Notify IT about request ready for processing
            $itEmail = "it@yourcompany.com"
            Send-WorkflowNotification -RecipientEmail $itEmail -Subject "New Onboarding Request Ready for Processing" -Body "A new onboarding request for $($Record.FirstName) $($Record.LastName) is ready for IT processing."
        }
    }
    
    return $Record
}
#endregion DATA MANAGEMENT FUNCTIONS

#region DATA SECURITY FUNCTIONS
# Function to encrypt string data for sensitive information
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

# Function to decrypt string data for data processing
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
        
        # Simple XOR decryption for encrypted data
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

# Function to create CSV backup before making changes
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
        # Copy the file to backup location
        Copy-Item -Path $SourcePath -Destination $backupFile -Force
        Write-Log "Created backup: $backupFile" -Level "INFO"
        
        # Cleanup old backups if exceeding maximum count
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

# Function to log data changes for auditing
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

# Enhanced function to save records with encryption for sensitive fields
function Save-OnboardingRecords {
    param(
        [PSCustomObject[]]$Records
    )
    
    # Backup existing file if it exists before making changes
    if (Test-Path $csvFile) {
        Backup-CsvFile -SourcePath $csvFile
    }
    
    try {
        # Create a copy of records with sensitive fields encrypted
        $encryptedRecords = $Records | ForEach-Object {
            $record = $_ | Select-Object * # Clone the record
            
            # Encrypt sensitive fields before storage
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
#endregion DATA SECURITY FUNCTIONS

#region GUI LOADING AND SETUP
Write-DebugMessage "Loading GUI from MainGUI.xaml"
$useXamlGUI = $true
try {
    # Get the XAML file path
    $xamlPath = Join-Path -Path $scriptPath -ChildPath "MainGUI.xaml"
    
    if (-not (Test-Path $xamlPath)) {
        $errorMsg = "XAML file not found at: $xamlPath"
        Write-Log $errorMsg -Level "ERROR"
        
        # Try to find other XAML files in the same directory to help diagnose the issue
        try {
            $availableXamlFiles = Get-ChildItem -Path $scriptPath -Filter "*.xaml" | Select-Object -ExpandProperty Name
            $errorMsg += "`nFound XAML files in directory: $($availableXamlFiles -join ', ')"
        }
        catch {
            # Ignore errors in this diagnostic step
        }
        
        [System.Windows.Forms.MessageBox]::Show(
            $errorMsg,
            "Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error)
        exit
    }
    
    # Load the XAML content with special handling for margin issues
    try {
        # Load XAML content with special handling for margin issues
        $xamlContent = Get-Content -Path $xamlPath -Raw
        
        # Fix margin issues in XAML (3-value margin like "0,15,5")
        $xamlContent = $xamlContent -replace '(Margin=")(\d+),(\d+),(\d+)(")', '$1$2,$3,$4,0$5'
        
        # Load the modified XAML
        [xml]$xaml = $xamlContent
        
        # Remove the class attribute from the XAML if it exists to prevent errors
        if ($xaml.DocumentElement.HasAttribute("Class")) {
            Write-Host "Removing Class attribute from XAML for compatibility"
            $xaml.DocumentElement.RemoveAttribute("Class")
        }
        
        $reader = New-Object System.Xml.XmlNodeReader $xaml
        $window = [System.Windows.Markup.XamlReader]::Load($reader)
        
        # Get named elements from XAML
        $xaml.SelectNodes("//*[@Name]") | ForEach-Object {
            $name = $_.Name
            $element = $window.FindName($name)
            Set-Variable -Name $name -Value $element -Scope Script
        }
        
        Write-Host "XAML loaded successfully"
    }
    catch {
        Write-Host "[ERROR] Error loading XAML: $($_.Exception.Message)" -ForegroundColor Red
        if ($_.Exception.InnerException) {
            Write-Host "[ERROR] Inner exception: $($_.Exception.InnerException.Message)" -ForegroundColor Red
            
            # Check if the error is related to Margin format
            if ($_.Exception.InnerException.Message -match "Margin") {
                Write-Host "`n[INFO] Attempting to repair XAML margin values..." -ForegroundColor Yellow
                
                # Create temporary repair module path
                $repairModulePath = Join-Path -Path $scriptPath -ChildPath "XamlRepair.psm1"
                
                # Write a simple repair function to the temp module
                @'
function Repair-XamlMargins {
    param([string]$XamlContent)
    return $XamlContent -replace '(Margin=")(\d+),(\d+),(\d+)(")', '$1$2,$3,$4,0$5'
}
'@ | Out-File -FilePath $repairModulePath -Force
                
                # Import the repair module
                Import-Module $repairModulePath -Force
                
                # Read original XAML and repair it
                $xamlContent = Get-Content -Path $xamlPath -Raw
                $repairedXaml = Repair-XamlMargins -XamlContent $xamlContent
                
                # Save repaired XAML to a new file
                $repairedXamlPath = Join-Path -Path $scriptPath -ChildPath "MainGUI.repaired.xaml"
                $repairedXaml | Out-File -FilePath $repairedXamlPath -Force
                
                Write-Host "`n[INFO] Repaired XAML saved to: $repairedXamlPath" -ForegroundColor Green
                Write-Host "[INFO] Please use the repaired XAML file for your application." -ForegroundColor Green
            }
        }
        
        # Add additional error handling as needed
        exit
    }
    
    # Check for potential ResourceDictionary issues
    $resourceDicts = $xaml.SelectNodes("//*[local-name()='ResourceDictionary']")
    if ($resourceDicts.Count -gt 0) {
        Write-DebugMessage "Found $($resourceDicts.Count) ResourceDictionary elements - checking for duplicate keys"
        foreach ($dict in $resourceDicts) {
            $keys = @{} 
            $duplicates = @()
            foreach ($resource in $dict.ChildNodes) {
                if ($resource.Key -and $keys.ContainsKey($resource.Key)) {
                    $duplicates += $resource.Key
                } else {
                    $keys[$resource.Key] = $true
                }
            }
            if ($duplicates.Count -gt 0) {
                Write-Log "Warning: Found duplicate keys in ResourceDictionary: $($duplicates -join ', ')" -Level "WARNING"
            }
        }
    }

    # Helper function to safely get controls with error handling
    function Get-SafeControl {
        param (
            [string]$ControlName,
            [switch]$Required
        )
        
        $control = $window.FindName($ControlName)
        if ($null -eq $control) {
            if ($Required) {
                Write-Log "Control '$ControlName' not found in XAML" -Level "ERROR"
                [System.Windows.Forms.MessageBox]::Show(
                    "The element '$ControlName' is missing in the XAML file",
                    "Element not found",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Warning
                )
            } else {
                Write-DebugMessage "Control '$ControlName' not found"
            }
        }
        
        return $control
    }

    # Create the XAML reader
    $reader = New-Object System.Xml.XmlNodeReader $xaml
    
    # Load the window from XAML
    $window = [Windows.Markup.XamlReader]::Load($reader)
    Write-DebugMessage "GUI loaded successfully"
    
    # Access controls defined in the XAML
    $controls = @{
        "tabControl" = Get-SafeControl -ControlName "tabControl" -Required
        "tabHR" = Get-SafeControl -ControlName "tabHR" -Required
        "tabManager" = Get-SafeControl -ControlName "tabManager" -Required
        "tabVerification" = Get-SafeControl -ControlName "tabVerification" -Required
        "tabIT" = Get-SafeControl -ControlName "tabIT" -Required
        
        # HR Tab Controls
        "txtFirstName" = Get-SafeControl -ControlName "txtFirstName"
        "txtLastName" = Get-SafeControl -ControlName "txtLastName"
        "chkExternal" = Get-SafeControl -ControlName "chkExternal"
        "txtExtCompany" = Get-SafeControl -ControlName "txtExtCompany"
        "txtDescription" = Get-SafeControl -ControlName "txtDescription"
        "cmbOffice" = Get-SafeControl -ControlName "cmbOffice"
        "txtPhone" = Get-SafeControl -ControlName "txtPhone"
        "txtMobile" = Get-SafeControl -ControlName "txtMobile"
        "txtMail" = Get-SafeControl -ControlName "txtMail"
        "dtpStartWorkDate" = Get-SafeControl -ControlName "dtpStartWorkDate"
        "cmbAssignedManager" = Get-SafeControl -ControlName "cmbAssignedManager"
        "btnHRSubmit" = Get-SafeControl -ControlName "btnHRSubmit"
        "txtHRNotes" = Get-SafeControl -ControlName "txtHRNotes"
        
        # Manager Tab Controls - Updated names to match XAML
        "txtPosition" = Get-SafeControl -ControlName "txtPosition"
        "cmbBusinessUnit" = Get-SafeControl -ControlName "cmbBusinessUnit"
        "txtPersonalNumber" = Get-SafeControl -ControlName "txtPersonalNumber"
        "dtpTermination" = Get-SafeControl -ControlName "dtpTermination"
        "chkTL" = Get-SafeControl -ControlName "chkTL"
        "chkAL" = Get-SafeControl -ControlName "chkAL"
        "txtManagerNotes" = Get-SafeControl -ControlName "txtManagerNotes"
        "btnManagerSubmit" = Get-SafeControl -ControlName "btnManagerSubmit"
        
        # Additional Manager Tab Controls from XAML
        "chkSoftwareSage" = Get-SafeControl -ControlName "chkSoftwareSage"
        "chkSoftwareGenesis" = Get-SafeControl -ControlName "chkSoftwareGenesis"
        "chkSoftwareNavision" = Get-SafeControl -ControlName "chkSoftwareNavision"
        "chkSoftwareSAP" = Get-SafeControl -ControlName "chkSoftwareSAP"
        "chkSoftwareERP" = Get-SafeControl -ControlName "chkSoftwareERP"
        "chkSoftwareCRM" = Get-SafeControl -ControlName "chkSoftwareCRM"
        "chkSoftwareLexware" = Get-SafeControl -ControlName "chkSoftwareLexware"
        "chkSoftwareCustom1" = Get-SafeControl -ControlName "chkSoftwareCustom1"
        "chkSoftwareCustom2" = Get-SafeControl -ControlName "chkSoftwareCustom2"
        "chkZugangLizenzmanager" = Get-SafeControl -ControlName "chkZugangLizenzmanager" 
        "chkZugangTerminalserver" = Get-SafeControl -ControlName "chkZugangTerminalserver"
        "chkZugangVPN" = Get-SafeControl -ControlName "chkZugangVPN"
        "txtZugriffe" = Get-SafeControl -ControlName "txtZugriffe"
        "cmbAusstattung" = Get-SafeControl -ControlName "cmbAusstattung"
        "txtArbeitsplatz" = Get-SafeControl -ControlName "txtArbeitsplatz"
        "txtWeitereUnternehmenssoftware" = Get-SafeControl -ControlName "txtWeitereUnternehmenssoftware"
        "cmbMS365Lizenzen" = Get-SafeControl -ControlName "cmbMS365Lizenzen"
        "chkMS365Email" = Get-SafeControl -ControlName "chkMS365Email"
        "chkMS365Teams" = Get-SafeControl -ControlName "chkMS365Teams"
        "chkMS365OneDrive" = Get-SafeControl -ControlName "chkMS365OneDrive"
        "chkMS365SharePoint" = Get-SafeControl -ControlName "chkMS365SharePoint"
        "chkMS365PowerBI" = Get-SafeControl -ControlName "chkMS365PowerBI"
        "chkMS365PowerApps" = Get-SafeControl -ControlName "chkMS365PowerApps"
        "txtZusatzanforderungen" = Get-SafeControl -ControlName "txtZusatzanforderungen"
        
        # Verification Tab Controls
        "lstPendingVerifications" = Get-SafeControl -ControlName "lstPendingVerifications"
        "chkHRVerified" = Get-SafeControl -ControlName "chkHRVerified"
        "txtVerificationNotes" = Get-SafeControl -ControlName "txtVerificationNotes"
        "btnVerifySubmit" = Get-SafeControl -ControlName "btnVerifySubmit"
        
        # IT Tab Controls
        "lstPendingIT" = Get-SafeControl -ControlName "lstPendingIT"
        "chkAccountCreated" = Get-SafeControl -ControlName "chkAccountCreated"
        "chkEquipmentReady" = Get-SafeControl -ControlName "chkEquipmentReady"
        "txtITNotes" = Get-SafeControl -ControlName "txtITNotes"
        "btnITComplete" = Get-SafeControl -ControlName "btnITComplete"
        "spITChecklist" = Get-SafeControl -ControlName "spITChecklist"  # Corrected from lstITChecklist to spITChecklist
        "txtAssetID" = Get-SafeControl -ControlName "txtAssetID"
        "btnITChecklistUpdate" = Get-SafeControl -ControlName "btnITChecklistUpdate" 
        "chkIT_MS365" = Get-SafeControl -ControlName "chkIT_MS365" 
        "chkIT_Software1" = Get-SafeControl -ControlName "chkIT_Software1" 
        "chkIT_Software2" = Get-SafeControl -ControlName "chkIT_Software2" 
        "chkIT_Computer" = Get-SafeControl -ControlName "chkIT_Computer" 
        "chkIT_Peripherie" = Get-SafeControl -ControlName "chkIT_Peripherie" 
        "chkIT_Netzwerk" = Get-SafeControl -ControlName "chkIT_Netzwerk" 
        "chkIT_VPN" = Get-SafeControl -ControlName "chkIT_VPN" 
        "chkIT_Smartphone" = Get-SafeControl -ControlName "chkIT_Smartphone" 
        "chkIT_Tablet" = Get-SafeControl -ControlName "chkIT_Tablet"
        
        # Common Controls
        "lblCurrentUser" = Get-SafeControl -ControlName "lblCurrentUser"
        "lblUserRole" = Get-SafeControl -ControlName "lblUserRole"
        "btnClose" = Get-SafeControl -ControlName "btnClose"
        "picLogo" = Get-SafeControl -ControlName "picLogo"
        "lstOnboardingRecords" = Get-SafeControl -ControlName "lstOnboardingRecords"
        "txtAccessLevel" = Get-SafeControl -ControlName "txtAccessLevel"
        "btnExportCSV" = Get-SafeControl -ControlName "btnExportCSV"
        "btnRefresh" = Get-SafeControl -ControlName "btnRefresh"
        "btnNew" = Get-SafeControl -ControlName "btnNew"
        "btnViewAuditLog" = Get-SafeControl -ControlName "btnViewAuditLog"
        "btnRestore" = Get-SafeControl -ControlName "btnRestore"
        "btnHelp" = Get-SafeControl -ControlName "btnHelp"
    }
    
    # Enable tabs based on user role for proper security
    if ($controls.tabControl -ne $null) {
        switch ($userRole) {
            "HR" {
                $controls.tabHR.IsEnabled = $true
                $controls.tabManager.IsEnabled = $false
                $controls.tabVerification.IsEnabled = $true
                $controls.tabIT.IsEnabled = $false
                $controls.tabControl.SelectedIndex = 0  # Select HR tab
            }
            "Manager" {
                $controls.tabHR.IsEnabled = $false
                $controls.tabManager.IsEnabled = $true
                $controls.tabVerification.IsEnabled = $false
                $controls.tabIT.IsEnabled = $false
                $controls.tabControl.SelectedIndex = 1  # Select Manager tab
            }
            "IT" {
                $controls.tabHR.IsEnabled = $false
                $controls.tabManager.IsEnabled = $false
                $controls.tabVerification.IsEnabled = $false
                $controls.tabIT.IsEnabled = $true
                $controls.tabControl.SelectedIndex = 3  # Select IT tab
            }
            "Admin" {
                $controls.tabHR.IsEnabled = $true
                $controls.tabManager.IsEnabled = $true
                $controls.tabVerification.IsEnabled = $true
                $controls.tabIT.IsEnabled = $true
                $controls.tabControl.SelectedIndex = 0  # Default to HR tab
            }
        }
    } else {
        Write-DebugMessage "Tab control not found in XAML"
    }
    
    # Display current user info in the UI
    if ($controls.lblCurrentUser -ne $null) {
        $controls.lblCurrentUser.Text = $currentUserName  # WPF TextBlock uses .Text
    }
    if ($controls.lblUserRole -ne $null) {
        $controls.lblUserRole.Text = $userRole  # WPF TextBlock uses .Text
    }
    
    # Set access level text for status display
    if ($controls.txtAccessLevel -ne $null) {
        $controls.txtAccessLevel.Text = $userRole
    }

    # Set default values for DatePickers
    if ($controls.dtpTermination) {
        $controls.dtpTermination.SelectedDate = [DateTime]::Now.AddYears(1)
    }
    if ($controls.dtpStartWorkDate) {
        $controls.dtpStartWorkDate.SelectedDate = [DateTime]::Now.AddDays(14)
    }
    
    # Load logo if exists
    if (Test-Path $headerLogo) {
        try {
            $imageSource = New-Object System.Windows.Media.Imaging.BitmapImage
            $imageSource.BeginInit()
            $imageSource.CacheOption = [System.Windows.Media.Imaging.BitmapCacheOption]::OnLoad
            $imageSource.UriSource = New-Object System.Uri($headerLogo, [System.UriKind]::Absolute)
            $imageSource.EndInit()
            $controls.picLogo.Source = $imageSource
            Write-DebugMessage "Logo loaded successfully from $headerLogo"
        }
        catch {
            Write-Log "Error loading logo: $_" -Level "WARNING"
        }
    } else {
        Write-Log "Logo file not found at $headerLogo. Skipping logo loading." -Level "WARNING"
    }
    
    # Load records list based on user role
    if ($controls.lstOnboardingRecords -ne $null) {
        $records = Get-UserRelevantRecords
        foreach ($record in $records) {
            $displayName = "$($record.FirstName) $($record.LastName) - $($record.WorkflowState)"
            $item = New-Object System.Windows.Controls.ListBoxItem
            $item.Content = $displayName
            $item.Tag = $record
            $controls.lstOnboardingRecords.Items.Add($item)
        }
    }
}
catch [System.Xml.XmlException] {
    Write-Log "XML parsing error in XAML file: $($_.Exception.Message)" -Level "ERROR"
    $useXamlGUI = $false
    
    # More detailed diagnostic info for XML errors
    $lineInfo = if ($_.Exception.LineNumber -gt 0) { 
        "Line: $($_.Exception.LineNumber), Position: $($_.Exception.LinePosition)" 
    } else { 
        "Position unavailable" 
    }
    
    [System.Windows.Forms.MessageBox]::Show(
        "XML parsing error in XAML: $($_.Exception.Message)`n$lineInfo`n`nTry running the XAML repair tool.",
        "XML Error",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error)
}
catch {
    Write-Log "Error loading or accessing GUI elements: $_" -Level "ERROR"
    Write-Log "Exception type: $($_.Exception.GetType().FullName)" -Level "ERROR"
    
    if ($_.Exception.InnerException) {
        Write-Log "Inner exception: $($_.Exception.InnerException.Message)" -Level "ERROR"
    }
    $useXamlGUI = $false
    
    [System.Windows.Forms.MessageBox]::Show(
        "Error loading GUI elements: $_`n`nInner exception: $($_.Exception.InnerException.Message)`n`nTry running the XAML repair tool.",
        "GUI Error",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error)
}

# Handle XAML loading failure
if (-not $useXamlGUI) {
    Write-Log "XAML loading failed - exiting script" -Level "ERROR"
    [System.Windows.Forms.MessageBox]::Show(
        "Critical error: Could not load the XAML UI. The application will now exit. Please check the log for details.",
        "Fatal Error",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error)
    exit
}

# Function to update tab access based on user role and workflow state
function Update-TabAccess {
    param(
        [PSCustomObject]$SelectedRecord = $null
    )
    
    Write-DebugMessage "Updating tab access controls"
    
    if ($null -eq $controls.tabControl) {
        Write-DebugMessage "Tab control not found, skipping tab access update"
        return
    }
    
    # Default state - disable all tabs
    $controls.tabHR.IsEnabled = $false
    $controls.tabManager.IsEnabled = $false  
    $controls.tabVerification.IsEnabled = $false  
    $controls.tabIT.IsEnabled = $false
    
    # Grant access based on user role and record state   
    switch ($userRole) {
        "HR" {
            # HR always has access to create new records
            $controls.tabHR.IsEnabled = $true
            if ($null -ne $SelectedRecord) {
                # HR can access verification tab only when there's a record ready for verification
                if ($null -ne $SelectedRecord -and $SelectedRecord.WorkflowState -eq $workflowStates["PendingHRVerification"]) {
                    $controls.tabVerification.IsEnabled = $true
                }
            }
            # Set active tab based on selected record
            if ($null -ne $SelectedRecord -and $SelectedRecord.WorkflowState -eq $workflowStates["PendingHRVerification"]) {
                $controls.tabControl.SelectedIndex = 2  # Verification tab
            } else {
                $controls.tabControl.SelectedIndex = 0  # HR tab
            }
        }
        "Manager" {
            # Manager can only access Manager tab and only for assigned records
            $controls.tabManager.IsEnabled = $true  
            
            # Start on manager tab
            $controls.tabControl.SelectedIndex = 1
            
            # If a record is selected that isn't in the right state, disable the tab
            if ($null -ne $SelectedRecord) {
                if ($SelectedRecord.WorkflowState -ne $workflowStates["PendingManagerInput"] -or 
                    $SelectedRecord.AssignedManager -ne $currentUserName) {
                    $controls.tabManager.IsEnabled = $false  
                }
            }
        }
        "IT" {
            # IT can only access IT tab and only for ready records
            $controls.tabIT.IsEnabled = $true
            $controls.tabControl.SelectedIndex = 3
            
            # If a record is selected that isn't in the right state, disable the tab
            if ($null -ne $SelectedRecord -and $SelectedRecord.WorkflowState -ne $workflowStates["ReadyForIT"]) {
                $controls.tabIT.IsEnabled = $false
            }
        }
        "Admin" {
            # Admin has full access to all tabs
            $controls.tabHR.IsEnabled = $true
            $controls.tabManager.IsEnabled = $true  
            $controls.tabVerification.IsEnabled = $true
            $controls.tabIT.IsEnabled = $true
            
            # Set active tab based on record state if one is selected
            if ($null -ne $SelectedRecord) {
                switch ($SelectedRecord.WorkflowState) {
                    $workflowStates["New"] { $controls.tabControl.SelectedIndex = 0 }
                    $workflowStates["PendingManagerInput"] { $controls.tabControl.SelectedIndex = 1 }
                    $workflowStates["PendingHRVerification"] { $controls.tabControl.SelectedIndex = 2 }
                    $workflowStates["ReadyForIT"] { $controls.tabControl.SelectedIndex = 3 }
                    $workflowStates["Completed"] { $controls.tabControl.SelectedIndex = 3 }
                    default { $controls.tabControl.SelectedIndex = 0 }
                }
            }
        }
    }
    
    # Special case for users with multiple roles
    if ($userRoles.Count -gt 1) {
        Write-DebugMessage "User has multiple roles, extending tab access permissions"
        
        if ($userRoles -contains "Admin") {
            # Admin role supersedes all restrictions
            $controls.tabHR.IsEnabled = $true
            $controls.tabManager.IsEnabled = $true  
            $controls.tabVerification.IsEnabled = $true
            $controls.tabIT.IsEnabled = $true
        } else {
            # Enable tabs based on all assigned roles
            if ($userRoles -contains "HR") {
                $controls.tabHR.IsEnabled = $true
                if ($null -ne $SelectedRecord -and $SelectedRecord.WorkflowState -eq $workflowStates["PendingHRVerification"]) {
                    $controls.tabVerification.IsEnabled = $true
                }
            }
            if ($userRoles -contains "Manager") {
                $controls.tabManager.IsEnabled = $true  
                # Only for assigned records in the right state
                if ($null -ne $SelectedRecord) {
                    if ($SelectedRecord.WorkflowState -ne $workflowStates["PendingManagerInput"] -or 
                        $SelectedRecord.AssignedManager -ne $currentUserName) {
                        $controls.tabManager.IsEnabled = $false  
                    }
                }
            }
            if ($userRoles -contains "IT") {
                $controls.tabIT.IsEnabled = $true
                if ($null -ne $SelectedRecord -and $SelectedRecord.WorkflowState -ne $workflowStates["ReadyForIT"]) {
                    $controls.tabIT.IsEnabled = $false
                }
            }
        }
    }
    
    # Display visual indication of workflow state and available actions
    if ($null -ne $controls.txtAccessLevel) {
        if ($null -ne $SelectedRecord) {
            $stateName = Get-WorkflowStateDisplayName -StateValue $SelectedRecord.WorkflowState
            $controls.txtAccessLevel.Text = "$userRole - $stateName"
        } else {
            $controls.txtAccessLevel.Text = $userRole
        }
    }
    
    Write-DebugMessage "Tab access updated: HR=$($controls.tabHR.IsEnabled), Manager=$($controls.tabManager.IsEnabled), Verification=$($controls.tabVerification.IsEnabled), IT=$($controls.tabIT.IsEnabled)"
}

#############################################
# Event Handlers
#############################################
# Function to validate input fields for HR
function Validate-HRInputFields {
    $isValid = $true
    $errorMessage = @()
    
    # Check required fields
    if ([string]::IsNullOrWhiteSpace($controls.txtFirstName.Text)) {
        $isValid = $false
        $errorMessage += "- First Name is required"
    }
    if ([string]::IsNullOrWhiteSpace($controls.txtLastName.Text)) {
        $isValid = $false
        $errorMessage += "- Last Name is required"
    }
    
    # Phone validation
    if (-not [string]::IsNullOrWhiteSpace($controls.txtPhone.Text) -and 
        ($controls.txtPhone.Text -notmatch "^[\d\+\-\(\) \.]+$")) {
        $isValid = $false
        $errorMessage += "- Phone number contains invalid characters"
    }
    if (-not [string]::IsNullOrWhiteSpace($controls.txtMobile.Text) -and 
        ($controls.txtMobile.Text -notmatch "^[\d\+\-\(\) \.]+$")) {
        $isValid = $false
        $errorMessage += "- Mobile number contains invalid characters"
    }
    
    # Email validation
    if (-not [string]::IsNullOrWhiteSpace($controls.txtMail.Text) -and 
        ($controls.txtMail.Text -notmatch "^[\w\-\.]+(@[\w\-\.]+)?$")) {
        $isValid = $false
        $errorMessage += "- Email username contains invalid characters"
    }
    
    # Manager assignment validation
    if ([string]::IsNullOrWhiteSpace($controls.cmbAssignedManager.Text)) {
        $isValid = $false
        $errorMessage += "- Assigned manager is required"
    }
    
    return @{
        IsValid = $isValid
        ErrorMessage = ($errorMessage -join "`n")
    }
}

# Function to validate input fields for Manager
function Validate-ManagerInputFields {
    $isValid = $true
    $errorMessage = @()
    
    # Check required fields
    if ([string]::IsNullOrWhiteSpace($controls.txtPosition.Text)) {
        $isValid = $false
        $errorMessage += "- Position is required"
    }
    if ([string]::IsNullOrWhiteSpace($controls.cmbBusinessUnit.Text)) {
        $isValid = $false
        $errorMessage += "- Business Unit is required"
    }
    
    return @{
        IsValid = $isValid
        ErrorMessage = ($errorMessage -join "`n")
    }
}

# Event handler for HR Submit button
if ($controls.btnHRSubmit -ne $null) {
    $controls.btnHRSubmit.Add_Click({
        Write-DebugMessage "HR Submit button clicked"
        try {
            # Validate HR fields
            $validation = Validate-HRInputFields
            if (-not $validation.IsValid) {
                Write-Log "HR validation error: $($validation.ErrorMessage)" -Level "ERROR"
                [System.Windows.Forms.MessageBox]::Show(
                    "Please correct the following errors:`n$($validation.ErrorMessage)", 
                    "Validation Error", 
                    [System.Windows.Forms.MessageBoxButtons]::OK, 
                    [System.Windows.Forms.MessageBoxIcon]::Warning)
                return
            }
            
            # Create new onboarding record
            $newRecord = [PSCustomObject]@{
                'RecordID'        = [Guid]::NewGuid().ToString()
                'FirstName'       = $controls.txtFirstName.Text.Trim()
                'LastName'        = $controls.txtLastName.Text.Trim()
                'Description'     = $controls.txtDescription.Text.Trim()
                'OfficeRoom'      = $controls.cmbOffice.Text.Trim()
                'PhoneNumber'     = $controls.txtPhone.Text.Trim()
                'MobileNumber'    = $controls.txtMobile.Text.Trim()
                'EmailAddress'    = $controls.txtMail.Text.Trim()
                'External'        = $controls.chkExternal.IsChecked
                'ExternalCompany' = if ($controls.chkExternal.IsChecked) { $controls.txtExtCompany.Text.Trim() } else { "" }
                'StartWorkDate'   = if ($controls.dtpStartWorkDate.SelectedDate) { $controls.dtpStartWorkDate.SelectedDate.ToString("yyyy-MM-dd") } else { "" }
                'AssignedManager' = $controls.cmbAssignedManager.Text.Trim()
                'WorkflowState'   = $workflowStates["PendingManagerInput"]
                'CreatedBy'       = $currentUserName
                'CreatedDate'     = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
                'LastUpdated'     = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
                'LastUpdatedBy'   = $currentUserName
                'HRNotes'         = ""
                'AccountDisabled' = $AccountDisabled
                'HRVerified'      = $false                 # Added missing property
                'VerificationNotes' = ""                   # Added missing property
                'AccountCreated'  = $false                 # Added missing property
                'EquipmentReady'  = $false                 # Added missing property
                'Processed'       = $false                 # Added missing property
                'ProcessedBy'     = ""                     # Added missing property
            }
            
            # Save to CSV with security
            $allRecords = Get-OnboardingRecords
            $allRecords += $newRecord
            $saveResult = Save-OnboardingRecords -Records $allRecords
            
            if ($saveResult) {
                # Log the action
                Write-AuditLog -Action "Create" -RecordID $newRecord.RecordID -Details "New onboarding record created for $($newRecord.FirstName) $($newRecord.LastName)"
                
                # Notify assigned manager
                $managerEmail = "$($newRecord.AssignedManager)@yourcompany.com" # Replace with your email domain
                $notificationSent = Send-WorkflowNotification -RecipientEmail $managerEmail -Subject "New Onboarding Request" -Body "A new onboarding request has been created for $($newRecord.FirstName) $($newRecord.LastName) and requires your input."
                
                # Show success message
                [System.Windows.Forms.MessageBox]::Show(
                    "Onboarding record created successfully. Manager has been notified.", 
                    "Success", 
                    [System.Windows.Forms.MessageBoxButtons]::OK, 
                    [System.Windows.Forms.MessageBoxIcon]::Information)
                
                # Clear input fields
                $controls.txtFirstName.Text = ""
                $controls.txtLastName.Text = ""
                $controls.chkExternal.IsChecked = $false
                $controls.txtExtCompany.Text = ""
                $controls.txtDescription.Text = ""
                $controls.cmbOffice.Text = ""   
                $controls.txtPhone.Text = ""   
                $controls.txtMobile.Text = ""   
                $controls.txtMail.Text = ""   
                $controls.cmbAssignedManager.Text = ""
                if ($controls.dtpStartWorkDate) {
                    $controls.dtpStartWorkDate.SelectedDate = [DateTime]::Now.AddDays(14)
                }
                
                # Refresh records list
                Update-RecordsList
                
                # After successful submission, refresh tab access
                Update-TabAccess
            } else {
                [System.Windows.Forms.MessageBox]::Show(
                    "Error saving record. Please try again.", 
                    "Save Error", 
                    [System.Windows.Forms.MessageBoxButtons]::OK, 
                    [System.Windows.Forms.MessageBoxIcon]::Error)
            }
        }
        catch {
            Write-Log "Error processing HR submission: $_" -Level "ERROR"
            [System.Windows.Forms.MessageBox]::Show(
                "Error saving data: $_", 
                "Error", 
                [System.Windows.Forms.MessageBoxButtons]::OK, 
                [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    })
}

# Event handler for Manager Submit button
if ($controls.btnManagerSubmit -ne $null) {
    $controls.btnManagerSubmit.Add_Click({
        Write-DebugMessage "Manager Submit button clicked"
        try {
            # Make sure a record is selected
            if ($global:CurrentRecord -eq $null) {
                [System.Windows.Forms.MessageBox]::Show(
                    "Please select an onboarding record from the list first.", 
                    "No Record Selected", 
                    [System.Windows.Forms.MessageBoxButtons]::OK, 
                    [System.Windows.Forms.MessageBoxIcon]::Warning)
                return
            }
            
            # Validate manager fields
            $validation = Validate-ManagerInputFields
            if (-not $validation.IsValid) {
                Write-Log "Manager validation error: $($validation.ErrorMessage)" -Level "ERROR"
                [System.Windows.Forms.MessageBox]::Show(
                    "Please correct the following errors:`n$($validation.ErrorMessage)", 
                    "Validation Error", 
                    [System.Windows.Forms.MessageBoxButtons]::OK, 
                    [System.Windows.Forms.MessageBoxIcon]::Warning)
                return
            }
            
            # Store original values for audit
            $originalState = $global:CurrentRecord.WorkflowState
            
            # Update record with manager data
            $global:CurrentRecord.Position = $controls.txtPosition.Text.Trim()
            $global:CurrentRecord.DepartmentField = $controls.cmbBusinessUnit.Text.Trim()
            $global:CurrentRecord.PersonalNumber = $controls.txtPersonalNumber.Text.Trim()
            $global:CurrentRecord.Ablaufdatum = if ($controls.dtpTermination.SelectedDate) { 
                $controls.dtpTermination.SelectedDate.ToString("yyyy-MM-dd") 
            } else { "" }
            $global:CurrentRecord.TL = $controls.chkTL.IsChecked
            $global:CurrentRecord.AL = $controls.chkAL.IsChecked
            $global:CurrentRecord.ManagerNotes = $controls.txtManagerNotes.Text.Trim()
            $global:CurrentRecord.ADGroup = $controls.cmbBusinessUnit.Text.Trim()
            
            # Add additional fields if controls exist
            if ($controls.txtZugriffe -ne $null) {
                $global:CurrentRecord.Zugriffe = $controls.txtZugriffe.Text.Trim()
            }
            
            # Check for software checkbox controls
            if ($controls.chkSoftwareSage -ne $null) {
                $global:CurrentRecord.SoftwareSage = $controls.chkSoftwareSage.IsChecked
            }
            if ($controls.chkSoftwareGenesis -ne $null) {
                $global:CurrentRecord.SoftwareGenesis = $controls.chkSoftwareGenesis.IsChecked
            }
            if ($controls.chkZugangLizenzmanager -ne $null) {
                $global:CurrentRecord.ZugangLizenzmanager = $controls.chkZugangLizenzmanager.IsChecked
            }
            
            # Update workflow state
            $global:CurrentRecord.WorkflowState = $workflowStates["PendingHRVerification"]
            $global:CurrentRecord.LastUpdated = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
            $global:CurrentRecord.LastUpdatedBy = $currentUserName
            
            # Update the CSV file securely
            $allRecords = Get-OnboardingRecords
            $updatedRecords = $allRecords | ForEach-Object {
                if ($_.RecordID -eq $global:CurrentRecord.RecordID) {
                    $global:CurrentRecord
                } else {
                    $_
                }
            }
            
            # Save with encryption
            $saveResult = Save-OnboardingRecords -Records $updatedRecords
            
            if ($saveResult) {
                # Log the change in audit
                Write-AuditLog -Action "Update" -RecordID $global:CurrentRecord.RecordID -Details "Manager data updated for $($global:CurrentRecord.FirstName) $($global:CurrentRecord.LastName). Status changed from $originalState to $($global:CurrentRecord.WorkflowState)"
                
                # Notify HR
                $hrEmail = "hr@yourcompany.com" # Replace with your HR email
                $notificationSent = Send-WorkflowNotification -RecipientEmail $hrEmail -Subject "Manager Input Completed" -Body "Manager $($currentUserName) has completed their input for $($global:CurrentRecord.FirstName) $($global:CurrentRecord.LastName)'s onboarding record and it requires HR verification."
                
                # Show success message
                [System.Windows.Forms.MessageBox]::Show(
                    "Manager information saved successfully. HR has been notified for verification.", 
                    "Success", 
                    [System.Windows.Forms.MessageBoxButtons]::OK, 
                    [System.Windows.Forms.MessageBoxIcon]::Information)
                
                # Clear current record
                $global:CurrentRecord = $null
                
                # Clear manager input fields
                $controls.txtPosition.Text = ""
                $controls.cmbBusinessUnit.Text = ""   
                $controls.txtPersonalNumber.Text = ""   
                $controls.chkTL.IsChecked = $false
                $controls.chkAL.IsChecked = $false
                $controls.txtManagerNotes.Text = ""
                
                # Clear additional fields if controls exist
                if ($controls.txtZugriffe -ne $null) {
                    $controls.txtZugriffe.Text = ""
                }
                if ($controls.chkSoftwareSage -ne $null) {
                    $controls.chkSoftwareSage.IsChecked = $false
                }
                if ($controls.chkSoftwareGenesis -ne $null) {
                    $controls.chkSoftwareGenesis.IsChecked = $false
                }
                if ($controls.chkZugangLizenzmanager -ne $null) {
                    $controls.chkZugangLizenzmanager.IsChecked = $false
                }
                
                if ($controls.dtpTermination) {
                    $controls.dtpTermination.SelectedDate = [DateTime]::Now.AddYears(1)
                }
                
                # Refresh records list
                Update-RecordsList
                
                # After successful submission, refresh tab access
                Update-TabAccess
            } else {
                [System.Windows.Forms.MessageBox]::Show(
                    "Error saving record. Please try again.", 
                    "Save Error", 
                    [System.Windows.Forms.MessageBoxButtons]::OK, 
                    [System.Windows.Forms.MessageBoxIcon]::Error)
            }
        }
        catch {
            Write-Log "Error processing manager submission: $_" -Level "ERROR"
            [System.Windows.Forms.MessageBox]::Show(
                "Error saving data: $_", 
                "Error", 
                [System.Windows.Forms.MessageBoxButtons]::OK, 
                [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    })
}

# Event handler for HR Verification Submit button
if ($controls.btnVerifySubmit -ne $null) {
    $controls.btnVerifySubmit.Add_Click({
        Write-DebugMessage "HR Verification Submit button clicked"
        try {
            # Make sure a record is selected
            if ($global:CurrentRecord -eq $null) {
                [System.Windows.Forms.MessageBox]::Show(
                    "Please select an onboarding record from the list first.", 
                    "No Record Selected", 
                    [System.Windows.Forms.MessageBoxButtons]::OK, 
                    [System.Windows.Forms.MessageBoxIcon]::Warning)
                return
            }
            
            # Check verification status
            if (-not $controls.chkHRVerified.IsChecked) {
                [System.Windows.Forms.MessageBox]::Show(
                    "Please check the 'Verified by HR' box to confirm verification.", 
                    "Verification Required", 
                    [System.Windows.Forms.MessageBoxButtons]::OK, 
                    [System.Windows.Forms.MessageBoxIcon]::Warning)
                return
            }
            
            # Store original state for audit
            $originalState = $global:CurrentRecord.WorkflowState
            
            # Update record with verification data - Add property if missing
            if (-not ($global:CurrentRecord.PSObject.Properties.Name -contains 'HRVerified')) {
                $global:CurrentRecord | Add-Member -MemberType NoteProperty -Name 'HRVerified' -Value $true
            } else {
                $global:CurrentRecord.HRVerified = $true
            }
            
            if (-not ($global:CurrentRecord.PSObject.Properties.Name -contains 'VerificationNotes')) {
                $global:CurrentRecord | Add-Member -MemberType NoteProperty -Name 'VerificationNotes' -Value $controls.txtVerificationNotes.Text.Trim()
            } else {
                $global:CurrentRecord.VerificationNotes = $controls.txtVerificationNotes.Text.Trim()
            }
            
            $global:CurrentRecord.WorkflowState = $workflowStates["ReadyForIT"]
            $global:CurrentRecord.LastUpdated = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
            $global:CurrentRecord.LastUpdatedBy = $currentUserName
            
            # Update the CSV file securely
            $allRecords = Get-OnboardingRecords
            $updatedRecords = $allRecords | ForEach-Object {
                if ($_.RecordID -eq $global:CurrentRecord.RecordID) {
                    $global:CurrentRecord
                } else {
                    $_
                }
            }
            
            # Save with encryption
            $saveResult = Save-OnboardingRecords -Records $updatedRecords
            
            if ($saveResult) {
                # Log the change in audit
                Write-AuditLog -Action "Verify" -RecordID $global:CurrentRecord.RecordID -Details "HR verified record for $($global:CurrentRecord.FirstName) $($global:CurrentRecord.LastName). Status changed from $originalState to $($global:CurrentRecord.WorkflowState)"
                
                # Notify IT
                $itEmail = "it@yourcompany.com" # Replace with your IT email
                $notificationSent = Send-WorkflowNotification -RecipientEmail $itEmail -Subject "Onboarding Record Ready for IT" -Body "An onboarding record for $($global:CurrentRecord.FirstName) $($global:CurrentRecord.LastName) has been verified by HR and is ready for IT processing."
                
                # Show success message
                [System.Windows.Forms.MessageBox]::Show(
                    "Record verified successfully. IT has been notified for processing.", 
                    "Success", 
                    [System.Windows.Forms.MessageBoxButtons]::OK, 
                    [System.Windows.Forms.MessageBoxIcon]::Information)
                
                # Clear current record
                $global:CurrentRecord = $null
                
                # Clear verification fields
                $controls.chkHRVerified.IsChecked = $false
                $controls.txtVerificationNotes.Text = ""
                
                # Refresh records list
                Update-RecordsList
                
                # After successful submission, refresh tab access
                Update-TabAccess
            } else {
                [System.Windows.Forms.MessageBox]::Show(
                    "Error saving verification. Please try again.", 
                    "Save Error", 
                    [System.Windows.Forms.MessageBoxButtons]::OK, 
                    [System.Windows.Forms.MessageBoxIcon]::Error)
            }
        }
        catch {
            Write-Log "Error processing HR verification: $_" -Level "ERROR"
            [System.Windows.Forms.MessageBox]::Show(
                "Error saving verification: $_", 
                "Error", 
                [System.Windows.Forms.MessageBoxButtons]::OK, 
                [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    })
}

# Event handler for IT Complete button
if ($controls.btnITComplete -ne $null) {
    $controls.btnITComplete.Add_Click({
        Write-DebugMessage "IT Complete button clicked"
        try {
            # Make sure a record is selected
            if ($global:CurrentRecord -eq $null) {
                [System.Windows.Forms.MessageBox]::Show(
                    "Please select an onboarding record from the list first.", 
                    "No Record Selected", 
                    [System.Windows.Forms.MessageBoxButtons]::OK, 
                    [System.Windows.Forms.MessageBoxIcon]::Warning)
                return
            }
            
            # Check required checkboxes
            if (-not $controls.chkAccountCreated.IsChecked) {
                [System.Windows.Forms.MessageBox]::Show(
                    "Please check the 'Account Created' box to confirm account creation.", 
                    "Account Creation Required", 
                    [System.Windows.Forms.MessageBoxButtons]::OK, 
                    [System.Windows.Forms.MessageBoxIcon]::Warning)
                return
            }
            
            if (-not $controls.chkEquipmentReady.IsChecked) {
                [System.Windows.Forms.MessageBox]::Show(
                    "Please check the 'Equipment Ready' box to confirm equipment preparation.", 
                    "Equipment Preparation Required", 
                    [System.Windows.Forms.MessageBoxButtons]::OK, 
                    [System.Windows.Forms.MessageBoxIcon]::Warning)
                return
            }
            
            # Store original state for audit
            $originalState = $global:CurrentRecord.WorkflowState
            
            # Update record with IT completion data - Add properties if missing
            if (-not ($global:CurrentRecord.PSObject.Properties.Name -contains 'AccountCreated')) {
                $global:CurrentRecord | Add-Member -MemberType NoteProperty -Name 'AccountCreated' -Value $true
            } else {
                $global:CurrentRecord.AccountCreated = $true
            }
            
            if (-not ($global:CurrentRecord.PSObject.Properties.Name -contains 'EquipmentReady')) {
                $global:CurrentRecord | Add-Member -MemberType NoteProperty -Name 'EquipmentReady' -Value $true
            } else {
                $global:CurrentRecord.EquipmentReady = $true
            }
            
            if (-not ($global:CurrentRecord.PSObject.Properties.Name -contains 'ITNotes')) {
                $global:CurrentRecord | Add-Member -MemberType NoteProperty -Name 'ITNotes' -Value $controls.txtITNotes.Text.Trim()
            } else {
                $global:CurrentRecord.ITNotes = $controls.txtITNotes.Text.Trim()
            }
            
            if (-not ($global:CurrentRecord.PSObject.Properties.Name -contains 'Processed')) {
                $global:CurrentRecord | Add-Member -MemberType NoteProperty -Name 'Processed' -Value $true
            } else {
                $global:CurrentRecord.Processed = $true
            }
            
            if (-not ($global:CurrentRecord.PSObject.Properties.Name -contains 'ProcessedBy')) {
                $global:CurrentRecord | Add-Member -MemberType NoteProperty -Name 'ProcessedBy' -Value $currentUserName
            } else {
                $global:CurrentRecord.ProcessedBy = $currentUserName
            }
            
            $global:CurrentRecord.WorkflowState = $workflowStates["Completed"]
            $global:CurrentRecord.LastUpdated = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
            $global:CurrentRecord.LastUpdatedBy = $currentUserName
            
            # Update the CSV file securely
            $allRecords = Get-OnboardingRecords
            $updatedRecords = $allRecords | ForEach-Object {
                if ($_.RecordID -eq $global:CurrentRecord.RecordID) {
                    $global:CurrentRecord
                } else {
                    $_
                }
            }
            
            # Save with encryption
            $saveResult = Save-OnboardingRecords -Records $updatedRecords
            
            if ($saveResult) {
                # Log the change in audit
                Write-AuditLog -Action "Complete" -RecordID $global:CurrentRecord.RecordID -Details "IT completed onboarding for $($global:CurrentRecord.FirstName) $($global:CurrentRecord.LastName). Status changed from $originalState to $($global:CurrentRecord.WorkflowState)"
                
                # Notify HR about completion
                $hrEmail = "hr@yourcompany.com" # Replace with your HR email
                $notificationSent = Send-WorkflowNotification -RecipientEmail $hrEmail -Subject "Onboarding Complete" -Body "The onboarding process for $($global:CurrentRecord.FirstName) $($global:CurrentRecord.LastName) has been completed by IT."
                
                # Show success message
                [System.Windows.Forms.MessageBox]::Show(
                    "Onboarding process completed successfully. HR has been notified.", 
                    "Success", 
                    [System.Windows.Forms.MessageBoxButtons]::OK, 
                    [System.Windows.Forms.MessageBoxIcon]::Information)
                
                # Clear current record
                $global:CurrentRecord = $null
                
                # Clear IT fields
                $controls.chkAccountCreated.IsChecked = $false
                $controls.chkEquipmentReady.IsChecked = $false
                $controls.txtITNotes.Text = ""
                
                # Refresh records list
                Update-RecordsList
                
                # After successful submission, refresh tab access
                Update-TabAccess
            } else {
                [System.Windows.Forms.MessageBox]::Show(
                    "Error saving completion. Please try again.", 
                    "Save Error", 
                    [System.Windows.Forms.MessageBoxButtons]::OK, 
                    [System.Windows.Forms.MessageBoxIcon]::Error)
            }
        }
        catch {
            Write-Log "Error completing IT process: $_" -Level "ERROR"
            [System.Windows.Forms.MessageBox]::Show(
                "Error saving completion: $_", 
                "Error", 
                [System.Windows.Forms.MessageBoxButtons]::OK, 
                [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    })
}

# Event handler for record selection
if ($controls.lstOnboardingRecords -ne $null) {
    $controls.lstOnboardingRecords.Add_SelectionChanged({
        $selectedItem = $controls.lstOnboardingRecords.SelectedItem
        if ($selectedItem -ne $null) {
            $global:CurrentRecord = $selectedItem.Tag
            Write-DebugMessage "Selected record: $($global:CurrentRecord.FirstName) $($global:CurrentRecord.LastName)"
            
            # Update tab access based on the selected record
            Update-TabAccess -SelectedRecord $global:CurrentRecord
            
            # Load record data into appropriate form based on user role and workflow state
            switch ($userRole) {
                "Manager" {
                    if ($global:CurrentRecord.WorkflowState -eq $workflowStates["PendingManagerInput"]) {
                        # Load data into manager form
                        $controls.txtPosition.Text = $global:CurrentRecord.Position
                        $controls.cmbBusinessUnit.Text = $global:CurrentRecord.DepartmentField
                        $controls.txtPersonalNumber.Text = $global:CurrentRecord.PersonalNumber
                        
                        try {
                            if (-not [string]::IsNullOrWhiteSpace($global:CurrentRecord.Ablaufdatum)) {
                                $controls.dtpTermination.SelectedDate = [DateTime]::ParseExact($global:CurrentRecord.Ablaufdatum, "yyyy-MM-dd", $null)
                            }
                        } catch {
                            Write-DebugMessage "Error parsing termination date: $_"
                        }
                        
                        # Use safer boolean parsing with default values
                        try {
                            $controls.chkTL.IsChecked = if ([bool]::TryParse($global:CurrentRecord.TL, [ref]$null)) { [bool]::Parse($global:CurrentRecord.TL) } else { $false }
                            $controls.chkAL.IsChecked = if ([bool]::TryParse($global:CurrentRecord.AL, [ref]$null)) { [bool]::Parse($global:CurrentRecord.AL) } else { $false }
                        } catch {
                            Write-DebugMessage "Error parsing boolean values: $_"
                            $controls.chkTL.IsChecked = $false
                            $controls.chkAL.IsChecked = $false
                        }
                        
                        # Set additional fields if controls exist
                        if ($controls.chkSoftwareSage -ne $null) {
                            $controls.chkSoftwareSage.IsChecked = $global:CurrentRecord.SoftwareSage -eq $true
                        }
                        if ($controls.chkSoftwareGenesis -ne $null) {
                            $controls.chkSoftwareGenesis.IsChecked = $global:CurrentRecord.SoftwareGenesis -eq $true
                        }
                        if ($controls.chkZugangLizenzmanager -ne $null) {
                            $controls.chkZugangLizenzmanager.IsChecked = $global:CurrentRecord.ZugangLizenzmanager -eq $true
                        }
                        if ($controls.txtZugriffe -ne $null) {
                            $controls.txtZugriffe.Text = $global:CurrentRecord.Zugriffe
                        }
                        
                        $controls.txtManagerNotes.Text = $global:CurrentRecord.ManagerNotes
                    }
                }
                "HR" {
                    if ($global:CurrentRecord.WorkflowState -eq $workflowStates["PendingHRVerification"]) {
                        # Load data into verification form
                        $controls.chkHRVerified.IsChecked = $false
                        $controls.txtVerificationNotes.Text = ""
                    }
                }
                "IT" {
                    if ($global:CurrentRecord.WorkflowState -eq $workflowStates["ReadyForIT"]) {
                        # Load data into IT form
                        $controls.chkAccountCreated.IsChecked = $false
                        $controls.chkEquipmentReady.IsChecked = $false
                        $controls.txtITNotes.Text = ""
                        
                        # Initialize IT checklist if needed
                        if ($controls.lstITChecklist -ne $null) {
                            $controls.lstITChecklist.Items.Clear()
                            $checklistItems = @(
                                "AD Account erstellen",
                                "Mailbox konfigurieren", 
                                "Berechtigungen einrichten",
                                "Hardware vorbereiten",
                                "Software installieren"
                            )
                            
                            foreach ($item in $checklistItems) {
                                $listItem = New-Object System.Windows.Controls.ListBoxItem
                                $checkBox = New-Object System.Windows.Controls.CheckBox
                                $checkBox.Content = $item
                                $checkBox.Margin = New-Object System.Windows.Thickness(5)
                                $listItem.Content = $checkBox
                                $controls.lstITChecklist.Items.Add($listItem)
                            }
                        }
                    }
                }
            }
        } else {
            $global:CurrentRecord = $null
            # Reset tab access to default state based on role
            Update-TabAccess
        }
    })
}

# Close button click handler
$controls.btnClose.Add_Click({
    Write-DebugMessage "Button 'Close' was clicked. Closing application."
    $window.Close()
})

# Set focus to the form - conditional for WPF vs WinForms
if ($useXamlGUI -and $window.Dispatcher) {
    Write-DebugMessage "Setting focus using WPF Dispatcher"
    $window.Dispatcher.Invoke([Action]{
        $window.Activate()
        if ($controls.lblCurrentUser) {
            $controls.lblCurrentUser.Text = $currentUserName  # WPF TextBlock verwendet .Text
        }
        if ($controls.lblUserRole) {
            $controls.lblUserRole.Text = $userRole  # WPF TextBlock verwendet .Text
        }
        if ($controls.txtAccessLevel) {
            $controls.txtAccessLevel.Text = $userRole
        }
    }, [System.Windows.Threading.DispatcherPriority]::ApplicationIdle)
} else {
    Write-DebugMessage "Setting focus using Windows Forms"
    $window.Activate()
}

# Show the form
Write-DebugMessage "Showing form..."
[void] $window.ShowDialog()
Write-DebugMessage "Script finished."

# Event handler for Refresh button
if ($controls.btnRefresh -ne $null) {
    $controls.btnRefresh.Add_Click({
        Write-DebugMessage "Refresh button clicked"
        
        # Get the current tab to determine appropriate filtering
        $currentTab = $controls.tabControl.SelectedItem
        $tabName = $currentTab.Name
        
        # Set filter based on current tab
        $filterState = switch ($tabName) {
            "tabHR" { "New" } # HR tab shows new records by default
            "tabManager" { "PendingManagerInput" } # Manager tab shows records pending manager input
            "tabVerification" { "PendingHRVerification" } # Verification tab shows records pending HR verification
            "tabIT" { "ReadyForIT" } # IT tab shows records ready for IT
            default { "" } # No filtering by default
        }
        
        # Update the records list with appropriate filter
        Update-RecordsList -FilterState $filterState
        
        # Provide feedback to user
        [System.Windows.Forms.MessageBox]::Show(
            "Die Liste wurde aktualisiert.",
            "Aktualisiert",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information)
    })
}

# Event handler for Tab selection change
$controls.tabControl.Add_SelectionChanged({
    if ($_ -eq $null -or $_.Source -eq $null) {
        return
    }
    
    Write-DebugMessage "Tab selection changed"
    
    # Get the name of the selected tab
    $selectedTab = $controls.tabControl.SelectedItem
    if ($selectedTab -eq $null) {
        return
    }
    $tabName = $selectedTab.Name
    
    Write-DebugMessage "Selected tab: $tabName"
    
    # Set filter based on selected tab
    $filterState = switch ($tabName) {
        "tabHR" { "New" } # HR tab shows new records by default
        "tabManager" { "PendingManagerInput" } # Manager tab shows records pending manager input
        "tabVerification" { "PendingHRVerification" } # Verification tab shows records pending HR verification
        "tabIT" { "ReadyForIT" } # IT tab shows records ready for IT
        default { "" } # No filtering by default
    }
    
    # Update the records list with appropriate filter
    Update-RecordsList -FilterState $filterState
})

# Update the initial list loading code
if ($controls.lstOnboardingRecords -ne $null) {
    # Initial list population - based on default tab
    $initialTab = $controls.tabControl.SelectedItem
    $initialTabName = if ($initialTab -ne $null) { $initialTab.Name } else { "" }
    
    $initialFilterState = switch ($initialTabName) {
        "tabHR" { "New" }
        "tabManager" { "PendingManagerInput" }
        "tabVerification" { "PendingHRVerification" }
        "tabIT" { "ReadyForIT" }
        default { "" }
    }
    
    Update-RecordsList -FilterState $initialFilterState
}

# Add CSV export functionality for IT - Modified to use existing button
if ((Test-UserRole -Role "IT") -or (Test-UserRole -Role "Admin")) {
    # Use the existing export button in the XAML
    $exportButton = $window.FindName("btnExportCSV")
    
    if ($exportButton -ne $null) {
        # Add event handler for export button
        $exportButton.Add_Click({
            Write-DebugMessage "Export CSV button clicked"
            try {
                # Get records ready for export (completed or ready for IT)
                $exportRecords = Get-OnboardingRecords | Where-Object {
                    $_.WorkflowState -eq $workflowStates["ReadyForIT"] -or
                    $_.WorkflowState -eq $workflowStates["Completed"]
                }
                
                if ($exportRecords.Count -eq 0) {
                    [System.Windows.Forms.MessageBox]::Show(
                        "Keine Datensätze für den Export gefunden.",
                        "Export Information",
                        [System.Windows.Forms.MessageBoxButtons]::OK,
                        [System.Windows.Forms.MessageBoxIcon]::Information)
                    return
                }
                
                # Create a SaveFileDialog to let user choose export location
                $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
                $saveFileDialog.Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
                $saveFileDialog.Title = "Export für ONBOARDING Tool speichern"
                $saveFileDialog.FileName = "ONBOARDING_Import_$(Get-Date -Format 'yyyyMMdd').csv"
                
                if ($saveFileDialog.ShowDialog() -eq 'OK') {
                    $exportPath = $saveFileDialog.FileName
                    
                    # Format data for ONBOARDING tool import
                    $formattedRecords = foreach ($record in $exportRecords) {
                        [PSCustomObject]@{
                            'SamAccountName' = "$($record.FirstName.ToLower()).$($record.LastName.ToLower())"
                            'FirstName' = $record.FirstName
                            'LastName' = $record.LastName
                            'Department' = $record.DepartmentField
                            'Position' = $record.Position
                            'Office' = $record.OfficeRoom
                            'Phone' = $record.PhoneNumber
                            'Mobile' = $record.MobileNumber
                            'Email' = $record.EmailAddress
                            'Manager' = $record.AssignedManager
                            'StartDate' = $record.StartWorkDate
                            'ExpiryDate' = $record.Ablaufdatum
                            'Description' = $record.Description
                            'IsExternal' = $record.External
                            'ExternalCompany' = $record.ExternalCompany
                            'AccountDisabled' = $record.AccountDisabled -eq $true -or $record.AccountDisabled -eq "True"
                            'IsTeamLead' = $record.TL -eq $true -or $record.TL -eq "True"
                            'IsDeptHead' = $record.AL -eq $true -or $record.AL -eq "True"
                            'PersonalNumber' = $record.PersonalNumber
                            'ADGroup' = $record.ADGroup
                        }
                    }
                    
                    # Export to CSV
                    $formattedRecords | Export-Csv -Path $exportPath -NoTypeInformation -Encoding UTF8
                    
                    [System.Windows.Forms.MessageBox]::Show(
                        "$($formattedRecords.Count) Datensätze wurden nach '$exportPath' exportiert.",
                        "Export erfolgreich",
                        [System.Windows.Forms.MessageBoxButtons]::OK,
                        [System.Windows.Forms.MessageBoxIcon]::Information)
                }
            }
            catch {
                Write-Log "Error exporting CSV: $_" -Level "ERROR"
                [System.Windows.Forms.MessageBox]::Show(
                    "Fehler beim Export der CSV-Datei: $_",
                    "Export Fehler",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Error)
            }
        })
    } else {
        Write-Log "Export CSV button not found in XAML" -Level "WARNING"
    }
}

# Initialize IT checklist at startup with enhanced error handling
try {
    # Check if controls dictionary exists
    if ($null -eq $controls) {
        Write-Log "Controls dictionary is null" -Level "ERROR"
        Write-DebugMessage "Controls dictionary is null, cannot initialize IT checklist"
        return
    }
    
    # Check if the IT checklist control exists in the dictionary
    if (-not $controls.ContainsKey("spITChecklist")) {  
        Write-Log "IT checklist key not found in controls dictionary" -Level "WARNING"
        Write-DebugMessage "spITChecklist key not found in controls dictionary"  
        
        # Add the key with null value to prevent further errors
        $controls["spITChecklist"] = $null  
        return
    }
    
    # Check if the control reference is valid
    if ($null -eq $controls.spITChecklist) {  
        Write-Log "IT checklist control reference is null" -Level "WARNING"
        Write-DebugMessage "IT checklist control reference is null"
        return
    }
    
    Write-DebugMessage "IT checklist already initialized in XAML, no need to add items programmatically"
    
    # Die IT-Checkliste ist bereits in der XAML definiert, daher müssen wir keine Elemente dynamisch hinzufügen
    # Stattdessen können wir die vorhandenen CheckBox-Steuerelemente verwenden und ihre Sichtbarkeit basierend auf der Auswahl steuern
}
catch {
    Write-Log "Critical error initializing IT checklist: $($_.Exception.Message)" -Level "ERROR"
    Write-DebugMessage "Critical error in IT checklist initialization: $($_.Exception.Message)"
    if ($_.Exception.InnerException) {
        Write-Log "Inner exception: $($_.Exception.InnerException.Message)" -Level "ERROR"
        Write-DebugMessage "Inner exception: $($_.Exception.InnerException.Message)"
    }
    
    if ($useXamlGUI -and $null -ne $window) {
        try {
            [System.Windows.Forms.MessageBox]::Show(
                "The IT checklist could not be initialized. The application will continue but some features may be limited.",
                "Initialization Warning",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning)
        }
        catch {
            # At this point we can't do much more
            Write-DebugMessage "Could not show error message to user: $($_.Exception.Message)"
        }
    }
}

# Event handler for IT checklist update button
if ($controls.btnITChecklistUpdate -ne $null) {
    $controls.btnITChecklistUpdate.Add_Click({
        Write-DebugMessage "IT checklist update button clicked"
        try {
            # Aktualisiere die Sichtbarkeit der VPN-Option basierend auf dem ausgewählten Datensatz
            if ($global:CurrentRecord -ne $null -and $controls.chkIT_VPN -ne $null) {
                if ($global:CurrentRecord.ZugangVPN -eq $true) {
                    $controls.chkIT_VPN.Visibility = [System.Windows.Visibility]::Visible
                } else {
                    $controls.chkIT_VPN.Visibility = [System.Windows.Visibility]::Collapsed
                }
            }
            
            # Aktualisiere die Sichtbarkeit der Smartphone-Option
            if ($global:CurrentRecord -ne $null -and $controls.chkIT_Smartphone -ne $null) {
                if ($global:CurrentRecord.RequiresSmartphone -eq $true) {
                    $controls.chkIT_Smartphone.Visibility = [System.Windows.Visibility]::Visible
                } else {
                    $controls.chkIT_Smartphone.Visibility = [System.Windows.Visibility]::Collapsed
                }
            }
            
            # Aktualisiere die Sichtbarkeit der Tablet-Option
            if ($global:CurrentRecord -ne $null -and $controls.chkIT_Tablet -ne $null) {
                if ($global:CurrentRecord.RequiresTablet -eq $true) {
                    $controls.chkIT_Tablet.Visibility = [System.Windows.Visibility]::Visible
                } else {
                    $controls.chkIT_Tablet.Visibility = [System.Windows.Visibility]::Collapsed
                }
            }
            
            [System.Windows.Forms.MessageBox]::Show(
                "Die IT-Checkliste wurde erfolgreich aktualisiert.",
                "Checkliste aktualisiert",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information)
        }
        catch {
            Write-Log "Error updating IT checklist: $_" -Level "ERROR"
            [System.Windows.Forms.MessageBox]::Show(
                "Fehler bei der Aktualisierung der IT-Checkliste: $_",
                "Fehler",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    })
}

# Event handler für den Help Button
# Überprüfe zuerst, ob der Hilfebutton in der XAML-Datei definiert ist
$btnHelp = $window.FindName("btnHelp")
if ($btnHelp -ne $null) {
    $btnHelp.Add_Click({
        Write-DebugMessage "Hilfe-Button geklickt"
        
        # Hilfesystem-Script laden
        $helpScriptPath = Join-Path -Path $scriptPath -ChildPath "HelpSystem.ps1"
        
        # Prüfen, ob die Hilfedatei existiert
        if (Test-Path $helpScriptPath) {
            try {
                # Hilfemodul importieren
                . $helpScriptPath
                
                # Hilfefenster anzeigen
                Show-HelpWindow
            } catch {
                Write-Log "Fehler beim Laden des Hilfesystems: $_" -Level "ERROR"
                [System.Windows.Forms.MessageBox]::Show(
                    "Das Hilfesystem konnte nicht geladen werden: $_",
                    "Fehler beim Laden der Hilfe",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Error)
            }
        } else {
            # Einfache Hilfe anzeigen, wenn die Hilfedatei nicht gefunden wurde
            $currentTab = $controls.tabControl.SelectedItem
            $tabName = if ($null -ne $currentTab) { $currentTab.Header } else { "Allgemein" }
            
            # Hilfsthema basierend auf dem aktuellen Tab bestimmen
            $helpTopic = switch ($tabName) {
                "HR" { "HR" }
                "Manager" { "Manager" }
                "Verifikation" { "Verification" }
                "IT" { "IT" }
                default { "Allgemein" }
            }
            
            # Einfache Hilfsnachricht anzeigen
            [System.Windows.Forms.MessageBox]::Show(
                "Die $helpTopic-Funktion hilft Ihnen bei der Verwaltung von Onboarding-Anfragen. Weitere Informationen finden Sie in der Dokumentation.",
                "Hilfe zu $helpTopic",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information)
        }
    })
}

# Event handler for the restore button
$btnRestore = $window.FindName("btnRestore")
if ($btnRestore -ne $null) {
    $btnRestore.Add_Click({
        Write-DebugMessage "Restore button clicked"
        
        # Überprüfen, ob der Benutzer die Berechtigung hat, Wiederherstellungen durchzuführen
        if (Test-UserPermission -Permission "ManageAll") {
            # Wiederherstellungsfunktion aufrufen
            Restore-FromBackup
        } else {
            [System.Windows.Forms.MessageBox]::Show(
                "Sie haben keine Berechtigung, Datensicherungen wiederherzustellen.",
                "Zugriff verweigert",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning)
        }
    })
}

# Function to recover from backup
function Restore-FromBackup {
    # Get a list of all available backups
    if (-not (Test-Path $backupFolder)) {
        [System.Windows.Forms.MessageBox]::Show(
            "No backup folder found.",
            "Restore Failed",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    
    $backups = Get-ChildItem -Path $backupFolder -Filter "*.bak" | Sort-Object LastWriteTime -Descending
    if ($backups.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show(
            "No backups found in backup folder.",
            "No Backups",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    
    # Create a simple form to select a backup
    $restoreForm = New-Object System.Windows.Forms.Form
    $restoreForm.Text = "Restore from Backup"
    $restoreForm.Size = New-Object System.Drawing.Size(500, 400)
    $restoreForm.StartPosition = "CenterScreen"
    
    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point(10, 10)
    $label.Size = New-Object System.Drawing.Size(480, 20)
    $label.Text = "Select a backup to restore:"
    $restoreForm.Controls.Add($label)
    
    $listBox = New-Object System.Windows.Forms.ListBox
    $listBox.Location = New-Object System.Drawing.Point(10, 40)
    $listBox.Size = New-Object System.Drawing.Size(480, 250)
    $listBox.SelectionMode = 'One'
    
    # Add backups to the list
    foreach ($backup in $backups) {
        $date = $backup.LastWriteTime.ToString("yyyy-MM-dd HH:mm:ss")
        $listBox.Items.Add("$date - $($backup.Name)")
        $listBox.Tag = $backups # Store the actual backup objects
    }
    
    $restoreForm.Controls.Add($listBox)
    
    # Restore button
    $restoreButton = New-Object System.Windows.Forms.Button
    $restoreButton.Location = New-Object System.Drawing.Point(10, 300)
    $restoreButton.Size = New-Object System.Drawing.Size(100, 30)
    $restoreButton.Text = "Restore"
    
    $restoreButton.Add_Click({
        $selectedIndex = $listBox.SelectedIndex
        if ($selectedIndex -ge 0) {
            $backupToRestore = ($listBox.Tag)[$selectedIndex]
            
            $result = [System.Windows.Forms.MessageBox]::Show(
                "Are you sure you want to restore from backup: $($backupToRestore.Name)?`nThis will overwrite the current data file.",
                "Confirm Restore",
                [System.Windows.Forms.MessageBoxButtons]::YesNo,
                [System.Windows.Forms.MessageBoxIcon]::Warning)
                
            if ($result -eq 'Yes') {
                try {
                    # Backup current file before overwriting
                    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
                    $preRestoreBackup = Join-Path -Path $backupFolder -ChildPath "PreRestore_$timestamp.bak"
                    
                    if (Test-Path $csvFile) {
                        Copy-Item -Path $csvFile -Destination $preRestoreBackup -Force
                    }
                    
                    # Restore the selected backup
                    Copy-Item -Path $backupToRestore.FullName -Destination $csvFile -Force
                    
                    # Log the restore
                    Write-AuditLog -Action "Restore" -RecordID "System" -Details "Data restored from backup: $($backupToRestore.Name)"
                    
                    [System.Windows.Forms.MessageBox]::Show(
                        "Backup restored successfully. The application will now reload the data.",
                        "Restore Complete",
                        [System.Windows.Forms.MessageBoxButtons]::OK,
                        [System.Windows.Forms.MessageBoxIcon]::Information)
                        
                    $restoreForm.DialogResult = [System.Windows.Forms.DialogResult]::OK
                    $restoreForm.Close()
                    
                    # Refresh the data in the application
                    Update-RecordsList
                }
                catch {
                    Write-Log "Error restoring from backup: $_" -Level "ERROR"
                    [System.Windows.Forms.MessageBox]::Show(
                        "Error restoring from backup: $_",
                        "Restore Failed",
                        [System.Windows.Forms.MessageBoxButtons]::OK,
                        [System.Windows.Forms.MessageBoxIcon]::Error)
                }
            }
        }
        else {
            [System.Windows.Forms.MessageBox]::Show(
                "Please select a backup to restore.",
                "No Selection",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning)
        }
    })
    
    $restoreForm.Controls.Add($restoreButton)
    
    # Cancel button
    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Location = New-Object System.Drawing.Point(120, 300)
    $cancelButton.Size = New-Object System.Drawing.Size(100, 30)
    $cancelButton.Text = "Cancel"
    
    $cancelButton.Add_Click({
        $restoreForm.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        $restoreForm.Close()
    })
    
    $restoreForm.Controls.Add($cancelButton)
    
    # Show the form
    $restoreForm.ShowDialog() | Out-Null
}

# Add a help button to the form if not already present
try {
    if ($null -ne $controls.btnClose -and $null -eq $btnHelp) {
        $placeForHelpButton = $true
        
        # Hilfe-Button erstellen
        $btnHelp = New-Object System.Windows.Controls.Button
        $btnHelp.Content = "Hilfe"
        $btnHelp.Margin = New-Object System.Windows.Thickness(0, 10, 10, 10)
        $btnHelp.FontWeight = [System.Windows.FontWeights]::Bold
        $btnHelp.Background = [System.Windows.Media.Brushes]::SteelBlue
        $btnHelp.Foreground = [System.Windows.Media.Brushes]::White
        $btnHelp.ToolTip = "Hilfe anzeigen"
        
        # Button im gleichen Container wie der Close-Button platzieren
        $parent = [System.Windows.Controls.Panel]::GetParent($controls.btnClose)
        if ($null -ne $parent) {
            if ($parent -is [System.Windows.Controls.StackPanel]) {
                # In StackPanel links vom Close-Button einfügen
                $index = $parent.Children.IndexOf($controls.btnClose)
                if ($index -ge 0) {
                    $parent.Children.Insert($index, $btnHelp)
                } else {
                    $parent.Children.Add($btnHelp)
                }
            } elseif ($parent -is [System.Windows.Controls.Grid]) {
                # In Grid mit gleichem Row/Column wie Close-Button, aber rechts davon
                $grid = $parent
                $row = [System.Windows.Controls.Grid]::GetRow($controls.btnClose)
                $col = [System.Windows.Controls.Grid]::GetColumn($controls.btnClose)
                
                # Prüfen, ob die Spaltenposition valide ist
                if ($col -gt 0) {
                    # Button zum Grid hinzufügen
                    $grid.Children.Add($btnHelp)
                    [System.Windows.Controls.Grid]::SetRow($btnHelp, $row)
                    [System.Windows.Controls.Grid]::SetColumn($btnHelp, $col - 1)
                } else {
                    # Fallback: Button direkt neben Close-Button platzieren
                    $btnHelp.HorizontalAlignment = [System.Windows.HorizontalAlignment]::Right
                    $btnHelp.Margin = New-Object System.Windows.Thickness(0, 10, 50, 10)
                    $parent.Children.Add($btnHelp)
                }
            } else {
                # Fallback: Button direkt neben Close-Button platzieren
                $btnHelp.HorizontalAlignment = [System.Windows.HorizontalAlignment]::Right
                $btnHelp.Margin = New-Object System.Windows.Thickness(0, 10, 50, 10)
                $parent.Children.Add($btnHelp)
            }
        }
        
        if (-not $placeForHelpButton) {
            Write-DebugMessage "Kein geeigneter Platz für den Hilfe-Button gefunden"
        }
        
        # Event handler für den neuen Help-Button
        $btnHelp.Add_Click({
            Write-DebugMessage "Neu erstellter Hilfe-Button geklickt"
            
            # Hilfesystem-Script laden
            $helpScriptPath = Join-Path -Path $scriptPath -ChildPath "HelpSystem.ps1"
            
            # Prüfen, ob die Hilfedatei existiert
            if (Test-Path $helpScriptPath) {
                try {
                    # Hilfemodul importieren
                    Import-Module $helpScriptPath -Force
                    
                    # Hilfefenster anzeigen
                    Show-HelpWindow
                } catch {
                    Write-Log "Fehler beim Laden des Hilfesystems: $_" -Level "ERROR"
                    [System.Windows.Forms.MessageBox]::Show(
                        "Das Hilfesystem konnte nicht geladen werden: $_",
                        "Fehler beim Laden der Hilfe",
                        [System.Windows.Forms.MessageBoxButtons]::OK,
                        [System.Windows.Forms.MessageBoxIcon]::Error)
                }
            } else {
                # Einfache Hilfe anzeigen, wenn die Hilfedatei nicht gefunden wurde
                $currentTab = $controls.tabControl.SelectedItem
                $tabName = if ($null -ne $currentTab) { $currentTab.Name } else { "General" }
                
                # Hilfsthema basierend auf dem aktuellen Tab bestimmen
                $helpTopic = switch ($tabName) {
                    "tabHR" { "HR" }
                    "tabManager" { "Manager" }
                    "tabVerification" { "Verification" }
                    "tabIT" { "IT" }
                    default { "General" }
                }
                
                # Hilfetext anzeigen
                $helpTitle = "Hilfe zu $helpTopic"
                $helpContent = "Die $helpTopic-Funktion hilft Ihnen bei der Verwaltung von Onboarding-Anfragen. Weitere Informationen finden Sie in der Dokumentation."
                
                [System.Windows.Forms.MessageBox]::Show(
                    $helpContent,
                    $helpTitle,
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Information)
            }
        })
    }
} catch {
    Write-DebugMessage "Fehler beim Erstellen des Hilfe-Buttons: $_"
}

# Event-Handler für den "New Request" Button hinzufügen
if ($null -ne $controls.btnNew) {
    $controls.btnNew.Add_Click({
        Write-DebugMessage "New Request button clicked"
        
        # Prüfen, ob der Benutzer die Berechtigung hat, neue Anfragen zu erstellen
        if (-not (Test-UserPermission -Permission "CreateRecord")) {
            [System.Windows.Forms.MessageBox]::Show(
                "Sie haben keine Berechtigung, neue Onboarding-Anfragen zu erstellen.",
                "Zugriff verweigert",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }
        
        # HR-Tab aktivieren und auswählen
        $controls.tabControl.SelectedIndex = 0 # HR tab
        $controls.tabHR.IsEnabled = $true
        
        # Eingabefelder zurücksetzen
        $controls.txtFirstName.Text = ""
        $controls.txtLastName.Text = ""
        $controls.chkExternal.IsChecked = $false
        $controls.txtExtCompany.Text = ""
        $controls.txtDescription.Text = ""
        $controls.cmbOffice.SelectedIndex = 0 # Standardwert wählen
        $controls.txtPhone.Text = ""
        $controls.txtMobile.Text = ""
        $controls.txtMail.Text = ""
        $controls.cmbAssignedManager.Text = ""
        
        # Standarddaten für Datumfelder setzen
        if ($controls.dtpStartWorkDate) {
            $controls.dtpStartWorkDate.SelectedDate = [DateTime]::Now.AddDays(14)
        }
        
        # Fokus auf das erste Feld setzen
        $controls.txtFirstName.Focus()
        
        # Bestätigung anzeigen
        [System.Windows.Forms.MessageBox]::Show(
            "Sie können nun eine neue Onboarding-Anfrage erstellen. Bitte füllen Sie alle erforderlichen Felder aus.",
            "Neue Anfrage",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information)
    })
}

# Manager-Liste für den Dropdown laden
function Load-ManagerList {
    if ($null -ne $controls.cmbAssignedManager) {
        $controls.cmbAssignedManager.Items.Clear()
        
        # Hier könnten Sie Manager dynamisch aus AD laden
        # Für die Demo verwenden wir statische Einträge
        $managerList = @(
            "Andreas Hepp",
            "Michael Müller",
            "Sabine Schmidt",
            "Thomas Becker",
            "Julia Weber",
            "Stefan König"
        )
        foreach ($manager in $managerList) {
            $controls.cmbAssignedManager.Items.Add($manager)
        }
        
        Write-DebugMessage "Manager-Liste geladen mit $($managerList.Count) Einträgen"
    }
}

# Initialisierung der Manager-Liste beim Programmstart
Load-ManagerList

# Event-Handler für MS365-Lizenzauswahl
$cmbMS365Lizenzen = $window.FindName("cmbMS365Lizenzen")
if ($null -ne $cmbMS365Lizenzen) {
    $cmbMS365Lizenzen.Add_SelectionChanged({
        if ($cmbMS365Lizenzen.SelectedItem -ne $null) {
            $selectedLicense = $cmbMS365Lizenzen.SelectedItem.ToString()
            Write-DebugMessage "MS365-Lizenz ausgewählt: $selectedLicense"
            
            # Optional: Hier könnten Sie basierend auf der gewählten Lizenz andere Felder aktualisieren
            if ($selectedLicense -eq "E3" -or $selectedLicense -eq "E5") {
                if ($controls.chkMS365Email -ne $null) {
                    $controls.chkMS365Email.IsChecked = $true
                }
            }
        }
    })
}

# Implementierung für ein einheitliches Ereignissystem
function Register-FormEvent {
    param (
        [string]$ControlName,
        [string]$EventName,
        [scriptblock]$Handler
    )
    
    $control = $window.FindName($ControlName)
    if ($null -ne $control) {
        $eventMethod = "add_$EventName"
        if ($control | Get-Member -MemberType Method -Name $eventMethod) {
            $control.$eventMethod.Invoke($Handler)
            Write-DebugMessage "Ereignishandler für $ControlName.$EventName registriert"
            return $true
        } else {
            Write-DebugMessage "Ereignismethode $eventMethod nicht gefunden für $ControlName"
            return $false
        }
    } else {
        Write-DebugMessage "Kontrollelement $ControlName nicht gefunden"
        return $false
    }
}

# Zusatzfunktion für externe Zugänge (wird aktiviert, wenn "External" angekreuzt wird)
if ($controls.chkExternal -ne $null -and $controls.txtExtCompany -ne $null) {
    Register-FormEvent -ControlName "chkExternal" -EventName "Click" -Handler {
        if ($controls.chkExternal.IsChecked) {
            $controls.txtExtCompany.IsEnabled = $true
            $controls.txtExtCompany.Focus()
        } else {
            $controls.txtExtCompany.IsEnabled = $false
            $controls.txtExtCompany.Text = ""
        }
    }
}

# Validierungsfunktionen für verschiedene Felder
function Validate-EmailFormat {
    param (
        [string]$Email
    )
    
    if ([string]::IsNullOrWhiteSpace($Email)) {
        return $true # Leere E-Mail ist in Ordnung (wird separat auf "Erforderlich" geprüft)
    }
    
    # Einfache E-Mail-Validierung
    return $Email -match '^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
}

# Validieren des E-Mail-Feldes bei Verlassen
Register-FormEvent -ControlName "txtMail" -EventName "LostFocus" -Handler {
    if (-not [string]::IsNullOrWhiteSpace($controls.txtMail.Text) -and -not (Validate-EmailFormat -Email $controls.txtMail.Text)) {
        [System.Windows.Forms.MessageBox]::Show(
            "Die eingegebene E-Mail-Adresse hat ein ungültiges Format.",
            "Ungültiges Format",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning)
        $controls.txtMail.Focus()
    }
}

# Hilfsfunktion zum Generieren von Vorschlägen für E-Mail und Benutzernamen
function Update-SuggestedUsername {
    if (-not [string]::IsNullOrWhiteSpace($controls.txtFirstName.Text) -and 
        -not [string]::IsNullOrWhiteSpace($controls.txtLastName.Text)) {
        
        $firstName = $controls.txtFirstName.Text.Trim().ToLower()
        $lastName = $controls.txtLastName.Text.Trim().ToLower()
        
        # E-Mail-Vorschlag generieren
        if ([string]::IsNullOrWhiteSpace($controls.txtMail.Text)) {
            $suggestedEmail = "$firstName.$lastName@yourcompany.com"
            $controls.txtMail.Text = $suggestedEmail
            Write-DebugMessage "E-Mail-Vorschlag generiert: $suggestedEmail"
        }
        
        # SAM-Account-Vorschlag generieren (für IT-Tab)
        $userIDField = $window.FindName("txtUserID")
        if ($null -ne $userIDField -and [string]::IsNullOrWhiteSpace($userIDField.Text)) {
            $suggestedSAM = "$firstName.$lastName"
            $userIDField.Text = $suggestedSAM
            Write-DebugMessage "SAM-Account-Vorschlag generiert: $suggestedSAM"
        }
    }
}

# Bei Änderung des Nachnamens Vorschläge aktualisieren
Register-FormEvent -ControlName "txtLastName" -EventName "LostFocus" -Handler {
    Update-SuggestedUsername
}

# Status-Dashboard-Aktualisierung (für visuelle Workflow-Indikatoren)
function Update-WorkflowStatusDisplay {
    param(
        [int]$CurrentStep = 1
    )
    
    # Workflow-Status-Anzeigen finden
    $steps = 1..4 | ForEach-Object {
        $window.FindName("workflowStep$_")
    }
    
    # Wenn die Elemente nicht direkt gefunden werden, versuchen wir es mit den Workflow-Badges
    if ($steps -contains $null) {
        Write-DebugMessage "Keine direkten Workflow-Step-Elemente gefunden, suche nach Workflow-Badges"
        $steps = @()
        
        # In der XAML haben wir Border-Elemente mit StackPanels für die Workflow-Anzeige
        # Wir suchen alle Border-Elemente mit dem entsprechenden Style
        $borders = $xaml.SelectNodes("//Border[@Style='{StaticResource WorkflowBadge}']")
        
        for ($i = 0; $i -lt $borders.Count; $i++) {
            $borderElement = $borders[$i]
            # Der entsprechende visuelle Baum muss hier durchsucht werden
            # Da dies komplex ist, vereinfachen wir und aktualisieren direkt den Hintergrund
            
            if ($i -lt ($CurrentStep - 1)) {
                # Bereits abgeschlossene Schritte
                $steps += @{ Status = "Completed" }
            }
            elseif ($i -eq ($CurrentStep - 1)) {
                # Aktueller Schritt
                $steps += @{ Status = "Current" }
            }
            else {
                # Zukünftige Schritte
                $steps += @{ Status = "Future" }
            }
        }
    }
    
    # Status-Anzeigen aktualisieren
    for ($i = 0; $i -lt $steps.Count; $i++) {
        $step = $steps[$i]
        if ($null -ne $step -and $step -is [System.Collections.Hashtable]) {
            # Visuellen Status aktualisieren
            $status = $step.Status
            
            # Hier würde die visuelle Aktualisierung erfolgen
            # Da wir keine direkten Referenzen haben, geben wir nur eine Debug-Nachricht aus
            Write-DebugMessage "Workflow-Schritt $($i+1) Status: $status"
        }
    }
    
    # Fortschrittsbalken aktualisieren, falls vorhanden
    $progressBar = $window.FindName("progressBar")
    if ($null -ne $progressBar) {
        $progressBar.Value = ($CurrentStep / 4) * 100
        Write-DebugMessage "Fortschrittsbalken aktualisiert auf $($progressBar.Value)%"
    }
}

# Initialisierung des Fortschrittsbalkens
Update-WorkflowStatusDisplay -CurrentStep 1

# SIG # Begin signature block
# MIIcCAYJKoZIhvcNAQcCoIIb+TCCG/UCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAUKaEDfR4OSrnL
# znIMwodc9H6CLIPUFoQwgXpaLX/zXqCCFk4wggMQMIIB+KADAgECAhB3jzsyX9Cg
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
# BgkqhkiG9w0BCQQxIgQgpPk2Yxxl1lyNupnuVnHavylGYM2r6ghdRe5BaPjEbzow
# DQYJKoZIhvcNAQEBBQAEggEAI50Ycu2AWvkchIfDSmxv/zeHy2w2N6fLtEuVyuEP
# qpYPnehqCeDdqskMHX1IQOKaJCJJFXCA3xHoFHwk5+CEwtxKNnYcj1FKr2RU85Za
# 01/c+rqBq9/Rf8SGes9VCJDjLIIgXfAgb6tV77rPQ6oBCRlHfzIqKlJWUO0mRZQt
# J4ex9mAh5Iym1iiPiBNfNgIvp3v+K5ohD0pEKNHyuUmryd5vO6Qs4R0HDsAR6TJE
# uJ3Ur+boAUPIlcz67z5tU3v82JxSkiawympxBJROJMSxEDS6C4ZqSS7DnHqznWde
# q3nw9Zc5+Ex8F8vYi4Y8J54FaNXk4PceXOaGogeyHodhd6GCAyYwggMiBgkqhkiG
# 9w0BCQYxggMTMIIDDwIBATB9MGkxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdp
# Q2VydCwgSW5jLjFBMD8GA1UEAxM4RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3Rh
# bXBpbmcgUlNBNDA5NiBTSEEyNTYgMjAyNSBDQTECEAqA7xhLjfEFgtHEdqeVdGgw
# DQYJYIZIAWUDBAIBBQCgaTAYBgkqhkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqG
# SIb3DQEJBTEPFw0yNTA3MTIwODAwMjhaMC8GCSqGSIb3DQEJBDEiBCAymT3MpqF+
# O9TaC5t4exh+Y1O8bJ215Dpvxg2xhSLodDANBgkqhkiG9w0BAQEFAASCAgCutn+U
# ocCa5wn6QwC2Z4zzfbOqcfD/eD6IHUa0vSsPxoW1VgsgMCMHyIOk5g/XWbKYPIwb
# 8Sh9jgBrDFWaA+XxshiSmjoPTvnZuGDFUYXE1f8mEuDVYM0yGDE4+QOb3jhEJ/C4
# dxkNPQw6mBKqi7bix6q8mhcX4QCH+e/hSca/Lg3LmxqvD1ut1fOZz3qvoC+hjDQY
# Sl5Cd7l872kPRQNLofFdFHdFnN0SiQa1FzrqqIHk35+MwayK1jDiXe0hnumMNrdG
# YPniJA+LaTMPMpDk/JqhP2e/w4vbN6eAWgQAYiO1AohqCKbHmMC1klIlCcyTSgBs
# /jpW2pg9dsIUiyDYFeJGyZ2jWBkQB9PbtujUDGw+ww6DmDsooIsNpH7GIgH4fdF4
# MEiPNVdrfz/5UOtBwS4fb7olNJKJfjKMXTv/XaWYj/zT47b2ZLoVn19WXee9tF4r
# DaEn9Ky13fvJI8RWvQptLwK1Lf+6XzhhqOIspFCpMv/6iC8pG6Y4HmWVWFlnTyna
# ZM+oN7qjvzuM6z33ghDFpZqlY0ACgvoeFp7tMbxOr/0gA3S4gaK4/SuZ2P0r3Csz
# oxqVjxWWLiaaM4ytaUhtMWVOzXMV5WWyBO8DZf7c9NFo9dudPy2k40/oQ0TJ6lYg
# n7xioOQX96a7oUYU3az8URvclRfqME67xiXkeQ==
# SIG # End signature block
