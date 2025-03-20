# Loads AD groups from the INI file and binds them to the GUI
Write-DebugMessage "Loading AD groups from INI."
function Load-ADGroups {
    try {
        Write-DebugMessage "Loading AD groups from INI..."
        
        # Check if global config is null
        if (-not $global:Config) {
            Write-Error "Global configuration is null."
            return
        }
        
        if (-not $global:Config.Contains("ADGroups")) {
            Write-Error "The [ADGroups] section is missing in the INI."
            return
        }
        
        $adGroupsIni = $global:Config["ADGroups"]
        $icADGroups = $window.FindName("icADGroups")
        
        if ($null -eq $icADGroups) {
            Write-Error "ItemsControl 'icADGroups' not found. Check the XAML name."
            return
        }
        
        # Create a list of AD group items for binding
        $adGroupItems = New-Object System.Collections.ObjectModel.ObservableCollection[PSObject]
        
        # Check if we have numbered entries like ADGroup1, ADGroup2 etc.
        $numberedGroupKeys = $adGroupsIni.Keys | Where-Object { $_ -match '^ADGroup\d+$' }
        if ($numberedGroupKeys.Count -gt 0) {
            # Sort numerically by the number after "ADGroup"
            $sortedKeys = $numberedGroupKeys | Sort-Object { [int]($_ -replace 'ADGroup', '') }
            
            # Add any remaining keys
            $otherKeys = $adGroupsIni.Keys | Where-Object { $_ -notmatch '^ADGroup\d+$' }
            $sortedKeys += $otherKeys
        } else {
            # Use original order from the OrderedDictionary
            $sortedKeys = $adGroupsIni.Keys
        }
        
        Write-DebugMessage "Found AD groups: $($sortedKeys.Count)"
        
        # For each entry in the ADGroups section create a data object
        foreach ($key in $sortedKeys) {
            # The key (e.g., "DEV") is the display name
            $displayName = $key
            
            # The value (e.g., "TEAM-1") is the actual group name
            $groupName = $adGroupsIni[$key]
            
            # Check if the group name is not empty
            if (-not [string]::IsNullOrWhiteSpace($groupName)) {
                # Create a custom object with display name and actual group name
                $groupItem = [PSCustomObject]@{
                    DisplayName = $displayName
                    GroupName = $groupName
                    IsChecked = $false
                    Key = $key
                }
                
                $adGroupItems.Add($groupItem)
                Write-DebugMessage "Added group: Display='$displayName', Group='$groupName'"
            }
        }
        
        # Set the ItemsSource to our collection
        $icADGroups.ItemsSource = $adGroupItems
        
        # Override the ItemTemplate to bind to DisplayName and store GroupName
        $factory = New-Object System.Windows.FrameworkElementFactory([System.Windows.Controls.CheckBox])
        $factory.SetBinding([System.Windows.Controls.ContentControl]::ContentProperty, (New-Object System.Windows.Data.Binding("DisplayName")))
        $factory.SetBinding([System.Windows.Controls.Primitives.ToggleButton]::IsCheckedProperty, (New-Object System.Windows.Data.Binding("IsChecked") -Property @{Mode = [System.Windows.Data.BindingMode]::TwoWay}))
        $factory.SetResourceReference([System.Windows.Controls.Control]::MarginProperty, "Margin")
        $factory.SetValue([System.Windows.Controls.ToolTipService]::ToolTipProperty, (New-Object System.Windows.Data.Binding("GroupName") -Property @{StringFormat = "AD Group: {0}"}))
        
        $template = New-Object System.Windows.DataTemplate
        $template.VisualTree = $factory
        
        $icADGroups.ItemTemplate = $template
        
        Write-DebugMessage "AD groups successfully loaded from INI: $($adGroupItems.Count) groups."
        Write-Log "AD groups successfully loaded from INI." "DEBUG"
        
        # Store reference to the items collection
        $global:ADGroupItems = $adGroupItems
    }
    catch {
        $errorMessage = "Error loading AD groups from INI: $($_.Exception.Message)"
        Write-Error $errorMessage
        Write-Log $errorMessage "ERROR"
    }
}

# Function to collect selected AD groups
function Get-SelectedADGroups {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [System.Windows.Controls.ItemsControl]$Panel,
        
        [Parameter(Mandatory=$false)]
        [string]$Separator = ";"
    )
    
    # Use a more efficient ArrayList to avoid output on add
    [System.Collections.ArrayList]$selectedGroups = @()
    
    if ($null -eq $Panel) {
        Write-DebugMessage "Error: Panel is null"
        return "NONE"
    }
    
    if ($null -eq $Panel.ItemsSource -or ($Panel.ItemsSource | Measure-Object).Count -eq 0) {
        Write-DebugMessage "Panel contains no items to select from"
        return "NONE"
    }
    
    # Get selected groups directly from data source
    foreach ($item in $Panel.ItemsSource) {
        if ($item.IsChecked) {
            Write-DebugMessage "Selected group: $($item.GroupName)"
            [void]$selectedGroups.Add($item.GroupName)
        }
    }
    
    Write-DebugMessage "Total selected groups: $($selectedGroups.Count)"
    
    if ($selectedGroups.Count -eq 0) {
        return "NONE"
    } else {
        return ($selectedGroups -join $Separator)
    }
}

# Export functions for use in other modules
Export-ModuleMember -Function Load-ADGroups, Get-SelectedADGroups


