# Function to set values and options for dropdown controls
Write-DebugMessage "Set-DropDownValues"
function Set-DropDownValues {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        $DropDownControl,    # The dropdown/combobox control (WinForms or WPF)

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$DataType,   # Required type (e.g., "OU", "Location", "License", etc.)

        [Parameter(Mandatory = $true)]
        [ValidateNotNull()]
        $ConfigData        # Configuration data source (e.g., Hashtable)
    )
    
    try {
        Write-Verbose "Populating dropdown for type '$DataType'..."
        Write-DebugMessage "Set-DropDownValues: Removing existing bindings and items"
        
        # Remove existing bindings and items (regardless of framework used)
        if ($DropDownControl -is [System.Windows.Forms.ComboBox]) {
            $DropDownControl.DataSource = $null
            $DropDownControl.Items.Clear()
        }
        elseif ($DropDownControl -is [System.Windows.Controls.ItemsControl]) {
            # For WPF, ensure we're operating on the UI thread
            if ($DropDownControl.Dispatcher.CheckAccess() -eq $false) {
                $DropDownControl.Dispatcher.Invoke({
                    $DropDownControl.ItemsSource = $null
                    $DropDownControl.Items.Clear()
                })
            } else {
                $DropDownControl.ItemsSource = $null
                $DropDownControl.Items.Clear()
            }
        }
        else {
            try { $DropDownControl.Items.Clear() } catch { }
        }
        Write-Verbose "Existing items removed."
        
        Write-DebugMessage "Set-DropDownValues: Retrieving items and default value from Config"
        $items = @()
        $defaultValue = $null
        switch ($DataType.ToLower()) {
            "ou" {
                $items = $ConfigData.OUList
                $defaultValue = $ConfigData.DefaultOU
            }
            "location" {
                $items = $ConfigData.LocationList
                $defaultValue = $ConfigData.DefaultLocation
            }
            "license" {
                $items = $ConfigData.LicenseList
                $defaultValue = $ConfigData.DefaultLicense
            }
            "mailsuffix" {
                $items = $ConfigData.MailSuffixList
                $defaultValue = $ConfigData.DefaultMailSuffix
            }
            "displaynametemplate" {
                $items = $ConfigData.DisplayNameTemplates
                $defaultValue = $ConfigData.DefaultDisplayNameTemplate
            }
            "tlgroups" {
                $items = $ConfigData.TLGroupsList
                $defaultValue = $ConfigData.DefaultTLGroup
            }
            default {
                Write-Warning "Unknown dropdown type '$DataType'. Aborting."
                return
            }
        }
    
        Write-DebugMessage "Set-DropDownValues: Ensuring items is a collection"
        if ($items) {
            if ($items -is [string] -or -not ($items -is [System.Collections.IEnumerable])) {
                $items = @($items)
            }
        }
        else {
            $items = @()
        }
        Write-Verbose "$($items.Count) entries for '$DataType' retrieved from configuration."
    
        Write-DebugMessage "Set-DropDownValues: Setting new data binding/items"
        if ($DropDownControl -is [System.Windows.Forms.ComboBox]) {
            $DropDownControl.DataSource = $items
        }
        else {
            # For WPF, ensure we're operating on the UI thread
            if ($DropDownControl.Dispatcher.CheckAccess() -eq $false) {
                $DropDownControl.Dispatcher.Invoke({
                    $DropDownControl.ItemsSource = $items
                })
            } else {
                $DropDownControl.ItemsSource = $items
            }
        }
        Write-Verbose "Data binding set for '$DataType' dropdown."
    
        Write-DebugMessage "Set-DropDownValues: Setting default value if available"
        if ($defaultValue) {
            # Ensure the default value exists in the items
            if ($items -contains $defaultValue) {
                if ($DropDownControl -is [System.Windows.Forms.ComboBox]) {
                    $DropDownControl.SelectedItem = $defaultValue
                } else {
                    # For WPF, ensure we're operating on the UI thread
                    if ($DropDownControl.Dispatcher.CheckAccess() -eq $false) {
                        $DropDownControl.Dispatcher.Invoke({
                            $DropDownControl.SelectedItem = $defaultValue
                        })
                    } else {
                        $DropDownControl.SelectedItem = $defaultValue
                    }
                }
                Write-Verbose "Default value for '$DataType' set to '$defaultValue'."
            } else {
                Write-Warning "Default value '$defaultValue' for '$DataType' not found in available options."
            }
        }
        elseif ($items.Count -gt 0) {
            try {
                 if ($DropDownControl -is [System.Windows.Forms.ComboBox]) {
                    $DropDownControl.SelectedIndex = 0
                } else {
                    # For WPF, ensure we're operating on the UI thread
                    if ($DropDownControl.Dispatcher.CheckAccess() -eq $false) {
                        $DropDownControl.Dispatcher.Invoke({
                            $DropDownControl.SelectedIndex = 0
                        })
                    } else {
                        $DropDownControl.SelectedIndex = 0
                    }
                }
            } catch {
                 Write-Warning "Could not set selected index to 0 for '$DataType'."
            }
        }
    }
    catch {
        Write-Error "Error populating dropdown '$DataType': $($_.Exception.Message)"
    }
}

# Export the function for use in other scripts
Export-ModuleMember -Function Set-DropDownValues