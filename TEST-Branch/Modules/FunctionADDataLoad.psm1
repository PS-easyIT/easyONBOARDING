Write-DebugMessage "Get-ADData"

# Function to retrieve and process data from Active Directory
Write-DebugMessage "Retrieving AD data."
function Get-ADData {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$false)]
        [System.Windows.Window]$Window
    )
    
    try {
        # Retrieve all OUs and convert to a sorted array
        $allOUs = Get-ADOrganizationalUnit -Filter * -Properties Name, DistinguishedName |
            ForEach-Object {
                [PSCustomObject]@{
                    Name = $_.Name
                    DN   = $_.DistinguishedName
                }
            } | Sort-Object Name

        # Debug output for OUs if DebugMode=1
        if ($global:Config.Logging.DebugMode -eq "1") {
            Write-Log " All OUs loaded:" "DEBUG"
            $allOUs | Format-Table | Out-String | Write-Log
        }
        Write-DebugMessage "Get-ADData: Retrieving all users"
        # Retrieve all users (here the first 200 users sorted by DisplayName are displayed)
        $allUsers = Get-ADUser -Filter * -Properties DisplayName, SamAccountName |
            ForEach-Object {
                [PSCustomObject]@{
                    DisplayName    = if ($_.DisplayName) { $_.DisplayName } else { $_.SamAccountName }
                    SamAccountName = $_.SamAccountName
                }
            } | Sort-Object DisplayName | Select-Object

        # Debug output for users if DebugMode=1
        if ($global:Config.Logging.DebugMode -eq "1") {
            Write-Log " User list loaded:" "DEBUG"
            $allUsers | Format-Table | Out-String | Write-Log
        }
        Write-DebugMessage "Get-ADData: comboBoxOU and listBoxUsers"
        
        # Find the ComboBox and ListBox in the Window if provided
        if ($PSBoundParameters.ContainsKey('Window') -and $Window) {
            $comboBoxOU = $Window.FindName("cmbOU") 
            $listBoxUsers = $Window.FindName("lstUsers")
            
            if ($comboBoxOU) {
                $comboBoxOU.ItemsSource = $null
                $comboBoxOU.Items.Clear()
                $comboBoxOU.ItemsSource = $allOUs
                $comboBoxOU.DisplayMemberPath = "Name"
                $comboBoxOU.SelectedValuePath = "DN"
                if ($allOUs.Count -gt 0) {
                    $comboBoxOU.SelectedIndex = 0
                }
            }

            if ($listBoxUsers) {
                $listBoxUsers.ItemsSource = $null
                $listBoxUsers.Items.Clear()
                $listBoxUsers.ItemsSource = $allUsers
                $listBoxUsers.DisplayMemberPath = "DisplayName"
                $listBoxUsers.SelectedValuePath = "SamAccountName"
            }
        }
        
        # Return the data for further use
        return @{
            OUs = $allOUs
            Users = $allUsers
        }
    }
    catch {
        Write-Error "Error in Get-ADData: $($_.Exception.Message)"
        Write-Log "Error in Get-ADData: $($_.Exception.Message)" "ERROR"
        [System.Windows.MessageBox]::Show("AD connection error: $($_.Exception.Message)")
        return $null
    }
}

# Functions for loading data from Active Directory

function Load-ADOrganizationalUnits {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$false)]
        [string]$ComboBoxName = "cmbOU"
    )
    
    try {
        if (Get-Command -Name "Write-DebugOutput" -ErrorAction SilentlyContinue) {
            Write-DebugOutput "Loading organizational units into $ComboBoxName" -Level INFO
        }
        
        $comboBox = $global:window.FindName($ComboBoxName)
        if ($null -eq $comboBox) {
            if (Get-Command -Name "Write-DebugOutput" -ErrorAction SilentlyContinue) {
                Write-DebugOutput "ComboBox $ComboBoxName not found" -Level WARNING
            }
            return $false
        }
        
        # Get OUs from Active Directory
        $ous = Get-ADOrganizationalUnit -Filter * -Properties Name, DistinguishedName | 
               Sort-Object -Property Name
        
        # Clear and populate the combobox
        $comboBox.Items.Clear()
        
        foreach ($ou in $ous) {
            $comboBox.Items.Add($ou.DistinguishedName)
        }
        
        # Select the first item if available
        if ($comboBox.Items.Count -gt 0) {
            $comboBox.SelectedIndex = 0
        }
        
        if (Get-Command -Name "Write-DebugOutput" -ErrorAction SilentlyContinue) {
            Write-DebugOutput "Loaded $($ous.Count) organizational units" -Level INFO
        }
        
        return $true
    }
    catch {
        if (Get-Command -Name "Write-DebugOutput" -ErrorAction SilentlyContinue) {
            Write-DebugOutput "Error loading organizational units: $($_.Exception.Message)" -Level ERROR
        }
        return $false
    }
}

function Load-TeamLeaderGroups {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$false)]
        [string]$ComboBoxName = "cmbTLGroup"
    )
    
    try {
        if (Get-Command -Name "Write-DebugOutput" -ErrorAction SilentlyContinue) {
            Write-DebugOutput "Loading team leader groups into $ComboBoxName" -Level INFO
        }
        
        $comboBox = $global:window.FindName($ComboBoxName)
        if ($null -eq $comboBox) {
            if (Get-Command -Name "Write-DebugOutput" -ErrorAction SilentlyContinue) {
                Write-DebugOutput "ComboBox $ComboBoxName not found" -Level WARNING
            }
            return $false
        }
        
        # Get groups that might be team leader groups (using naming convention like *TL* or *TeamLead*)
        $filter = "Name -like '*TL*' -or Name -like '*TeamLead*' -or Name -like '*Team*'"
        $groups = Get-ADGroup -Filter $filter -Properties Name, Description | 
                 Sort-Object -Property Name
        
        # Clear and populate the combobox
        $comboBox.Items.Clear()
        
        foreach ($group in $groups) {
            $comboBox.Items.Add($group.DistinguishedName)
        }
        
        # Select the first item if available
        if ($comboBox.Items.Count -gt 0) {
            $comboBox.SelectedIndex = 0
        }
        
        if (Get-Command -Name "Write-DebugOutput" -ErrorAction SilentlyContinue) {
            Write-DebugOutput "Loaded $($groups.Count) team leader groups" -Level INFO
        }
        
        return $true
    }
    catch {
        if (Get-Command -Name "Write-DebugOutput" -ErrorAction SilentlyContinue) {
            Write-DebugOutput "Error loading team leader groups: $($_.Exception.Message)" -Level ERROR
        }
        return $false
    }
}

# Export all functions
Export-ModuleMember -Function Get-ADData, Load-ADOrganizationalUnits, Load-TeamLeaderGroups
