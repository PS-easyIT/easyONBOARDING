# Functions for managing AD groups in the GUI

function Load-ADGroups {
    [CmdletBinding()]
    param()
    
    try {
        if (Get-Command -Name "Write-DebugOutput" -ErrorAction SilentlyContinue) {
            Write-DebugOutput "Loading AD groups" -Level INFO
        }
        
        $icADGroups = $global:window.FindName("icADGroups")
        if ($null -eq $icADGroups) {
            if (Get-Command -Name "Write-DebugOutput" -ErrorAction SilentlyContinue) {
                Write-DebugOutput "icADGroups control not found" -Level WARNING
            }
            return $false
        }
        
        # Clear existing items
        $icADGroups.Items.Clear()
        
        # Get AD groups
        try {
            $groups = Get-ADGroup -Filter * -Properties Name, Description -ResultSetSize 500 | 
                      Sort-Object -Property Name
        }
        catch {
            if (Get-Command -Name "Write-DebugOutput" -ErrorAction SilentlyContinue) {
                Write-DebugOutput "Error retrieving AD groups: $($_.Exception.Message)" -Level ERROR
            }
            return $false
        }
        
        # Create the items control panel to host checkboxes
        foreach ($group in $groups) {
            $checkBox = New-Object System.Windows.Controls.CheckBox
            $checkBox.Content = $group.Name
            $checkBox.Tag = $group.DistinguishedName
            $checkBox.Margin = New-Object System.Windows.Thickness(2)
            
            # Add tooltip if description is available
            if (-not [string]::IsNullOrEmpty($group.Description)) {
                $checkBox.ToolTip = $group.Description
            }
            
            $icADGroups.Items.Add($checkBox)
        }
        
        if (Get-Command -Name "Write-DebugOutput" -ErrorAction SilentlyContinue) {
            Write-DebugOutput "Loaded $($groups.Count) AD groups" -Level INFO
        }
        
        return $true
    }
    catch {
        if (Get-Command -Name "Write-DebugOutput" -ErrorAction SilentlyContinue) {
            Write-DebugOutput "Error loading AD groups: $($_.Exception.Message)" -Level ERROR
        }
        return $false
    }
}

function Get-SelectedADGroups {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [System.Windows.Controls.ItemsControl]$Panel
    )
    
    try {
        $selectedGroups = @()
        
        foreach ($item in $Panel.Items) {
            if ($item.IsChecked) {
                $selectedGroups += $item.Tag.ToString()
            }
        }
        
        if (Get-Command -Name "Write-DebugOutput" -ErrorAction SilentlyContinue) {
            Write-DebugOutput "Selected $($selectedGroups.Count) AD groups" -Level INFO
        }
        
        return $selectedGroups
    }
    catch {
        if (Get-Command -Name "Write-DebugOutput" -ErrorAction SilentlyContinue) {
            Write-DebugOutput "Error getting selected AD groups: $($_.Exception.Message)" -Level ERROR
        }
        return @()
    }
}

# Export all functions
Export-ModuleMember -Function Load-ADGroups, Get-SelectedADGroups


