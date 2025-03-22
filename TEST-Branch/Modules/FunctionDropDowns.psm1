# Functions for managing dropdown controls

function Initialize-DisplayNameTemplateDropdown {
    [CmdletBinding()]
    param()
    
    try {
        $cmbDisplayTemplate = $global:window.FindName("cmbDisplayTemplate")
        if ($null -eq $cmbDisplayTemplate) {
            if (Get-Command -Name "Write-DebugOutput" -ErrorAction SilentlyContinue) {
                Write-DebugOutput "cmbDisplayTemplate control not found" -Level WARNING
            }
            return $false
        }
        
        $templates = @("Vorname.Nachname", "NachnameV", "Nachname.Vorname", "VornameN")
        $cmbDisplayTemplate.Items.Clear()
        foreach ($template in $templates) {
            $cmbDisplayTemplate.Items.Add($template)
        }
        
        if ($cmbDisplayTemplate.Items.Count -gt 0) {
            $cmbDisplayTemplate.SelectedIndex = 0
        }
        
        if (Get-Command -Name "Write-DebugOutput" -ErrorAction SilentlyContinue) {
            Write-DebugOutput "Display name template dropdown initialized" -Level INFO
        }
        
        return $true
    }
    catch {
        if (Get-Command -Name "Write-DebugOutput" -ErrorAction SilentlyContinue) {
            Write-DebugOutput "Error initializing display name template dropdown: $($_.Exception.Message)" -Level ERROR
        }
        return $false
    }
}

function Initialize-LicenseDropdown {
    [CmdletBinding()]
    param()
    
    try {
        $cmbLicense = $global:window.FindName("cmbLicense")
        if ($null -eq $cmbLicense) {
            if (Get-Command -Name "Write-DebugOutput" -ErrorAction SilentlyContinue) {
                Write-DebugOutput "cmbLicense control not found" -Level WARNING
            }
            return $false
        }
        
        $licenses = @("Standard", "Premium", "Enterprise", "Keine")
        $cmbLicense.Items.Clear()
        foreach ($license in $licenses) {
            $cmbLicense.Items.Add($license)
        }
        
        if ($cmbLicense.Items.Count -gt 0) {
            $cmbLicense.SelectedIndex = 0
        }
        
        if (Get-Command -Name "Write-DebugOutput" -ErrorAction SilentlyContinue) {
            Write-DebugOutput "License dropdown initialized" -Level INFO
        }
        
        return $true
    }
    catch {
        if (Get-Command -Name "Write-DebugOutput" -ErrorAction SilentlyContinue) {
            Write-DebugOutput "Error initializing license dropdown: $($_.Exception.Message)" -Level ERROR
        }
        return $false
    }
}

function Initialize-EmailDomainSuffixDropdown {
    [CmdletBinding()]
    param()
    
    try {
        $cmbSuffix = $global:window.FindName("cmbSuffix")
        if ($null -eq $cmbSuffix) {
            if (Get-Command -Name "Write-DebugOutput" -ErrorAction SilentlyContinue) {
                Write-DebugOutput "cmbSuffix control not found" -Level WARNING
            }
            return $false
        }
        
        # Load domains from config if available
        $domains = @("example.com", "example.org")
        if ($global:Config -and $global:Config.Contains("Email") -and $global:Config["Email"].Contains("Domains")) {
            $configDomains = $global:Config["Email"]["Domains"] -split ','
            if ($configDomains.Count -gt 0) {
                $domains = $configDomains
            }
        }
        
        $cmbSuffix.Items.Clear()
        foreach ($domain in $domains) {
            $cmbSuffix.Items.Add($domain.Trim())
        }
        
        if ($cmbSuffix.Items.Count -gt 0) {
            $cmbSuffix.SelectedIndex = 0
        }
        
        if (Get-Command -Name "Write-DebugOutput" -ErrorAction SilentlyContinue) {
            Write-DebugOutput "Email domain suffix dropdown initialized" -Level INFO
        }
        
        return $true
    }
    catch {
        if (Get-Command -Name "Write-DebugOutput" -ErrorAction SilentlyContinue) {
            Write-DebugOutput "Error initializing email domain suffix dropdown: $($_.Exception.Message)" -Level ERROR
        }
        return $false
    }
}

function Initialize-AllDropdowns {
    [CmdletBinding()]
    param()
    
    $success = $true
    
    if (-not (Initialize-DisplayNameTemplateDropdown)) {
        $success = $false
    }
    
    if (-not (Initialize-LicenseDropdown)) {
        $success = $false
    }
    
    if (-not (Initialize-EmailDomainSuffixDropdown)) {
        $success = $false
    }
    
    return $success
}

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

# Export all functions
Export-ModuleMember -Function Initialize-DisplayNameTemplateDropdown, Initialize-LicenseDropdown, Initialize-EmailDomainSuffixDropdown, Initialize-AllDropdowns, Set-DropDownValues