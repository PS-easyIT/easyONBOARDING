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

# Export the function for use in other scripts
Export-ModuleMember -Function Get-ADData
