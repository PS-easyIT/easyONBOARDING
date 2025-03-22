# Functions to populate various dropdown menus from configuration data
# Ensure Debug function exists
if (-not (Get-Command -Name Write-DebugMessage -ErrorAction SilentlyContinue)) {
    function Write-DebugMessage {
        [CmdletBinding()]
        param([string]$Message)
        Write-Verbose $Message
    }
    Write-Verbose "Created local Write-DebugMessage function in GUIDropDowns module"
}

#region [Region 14.1 | OU DROPDOWN]
# Registers event handler for OU dropdown refresh button
function Register-OUDropdownRefreshEvent {
    [CmdletBinding()]
    param()
    
    Write-DebugMessage "Initializing OU dropdown refresh functionality"
    
    # Ensure window object exists
    if (-not (Get-Variable -Name window -Scope Global -ErrorAction SilentlyContinue)) {
        Write-DebugMessage "Error: Global window object not found"
        return $false
    }
    
    try {
        $btnRefreshOU = $global:window.FindName("btnRefreshOU")
        if (-not $btnRefreshOU) {
            Write-DebugMessage "Button 'btnRefreshOU' not found in the window"
            return $false
        }
        
        # Register click event for the refresh button
        $btnRefreshOU.Add_Click({
            try {
                if (-not (Get-Variable -Name Config -Scope Global -ErrorAction SilentlyContinue)) {
                    [System.Windows.MessageBox]::Show("Configuration data not loaded.", "Error")
                    return
                }
                
                $defaultOUFromINI = $global:Config.ADUserDefaults["DefaultOU"]
                if ([string]::IsNullOrEmpty($defaultOUFromINI)) {
                    [System.Windows.MessageBox]::Show("Default OU not configured in INI file.", "Error")
                    return
                }
                
                $OUList = Get-ADOrganizationalUnit -Filter * -SearchBase $defaultOUFromINI |
                    Select-Object -ExpandProperty DistinguishedName
                
                Write-DebugMessage "Found OUs: $($OUList.Count)"
                
                $cmbOU = $global:window.FindName("cmbOU")
                if ($cmbOU) {
                    # First set ItemsSource to null
                    $cmbOU.ItemsSource = $null
                    # Then clear the ItemsCollection
                    $cmbOU.Items.Clear()
                    # Now assign the new ItemsSource
                    $cmbOU.ItemsSource = $OUList
                    Write-DebugMessage "Dropdown: OU-DropDown successfully populated."
                } else {
                    Write-DebugMessage "ComboBox 'cmbOU' not found in the window"
                }
            }
            catch {
                $errorMsg = "Error loading OUs: $($_.Exception.Message)"
                Write-DebugMessage $errorMsg
                [System.Windows.MessageBox]::Show($errorMsg, "Error")
            }
        })
        
        Write-DebugMessage "OU refresh button event registered successfully"
        return $true
    }
    catch {
        Write-DebugMessage "Error registering OU dropdown refresh event: $($_.Exception.Message)"
        return $false
    }
}
#endregion

#region [Region 14.2 | DISPLAY NAME TEMPLATES]
# Populates the display name template dropdown from INI configuration
function Initialize-DisplayNameTemplateDropdown {
    [CmdletBinding()]
    param()
    
    Write-DebugMessage "Initializing display name template dropdown"
    
    # Ensure window object exists
    if (-not (Get-Variable -Name window -Scope Global -ErrorAction SilentlyContinue)) {
        Write-DebugMessage "Error: Global window object not found"
        return $false
    }
    
    # Ensure Config exists
    if (-not (Get-Variable -Name Config -Scope Global -ErrorAction SilentlyContinue)) {
        Write-DebugMessage "Error: Global Config object not found"
        return $false
    }
    
    try {
        $comboBoxDisplayTemplate = $global:window.FindName("cmbDisplayTemplate")
        if (-not $comboBoxDisplayTemplate) {
            Write-DebugMessage "ComboBox 'cmbDisplayTemplate' not found in the window"
            return $false
        }
        
        if (-not $global:Config.Contains("DisplayNameUPNTemplates")) {
            Write-DebugMessage "DisplayNameUPNTemplates section missing in configuration"
            return $false
        }
        
        # Clear previous items and bindings
        $comboBoxDisplayTemplate.Items.Clear()
        [System.Windows.Data.BindingOperations]::ClearBinding($comboBoxDisplayTemplate, [System.Windows.Controls.ComboBox]::ItemsSourceProperty)

        $DisplayNameTemplateList = @()
        $displayTemplates = $global:Config["DisplayNameUPNTemplates"]

        # Add the default entry (from DefaultDisplayNameFormat)
        $defaultDisplayNameFormat = if ($displayTemplates.Contains("DefaultDisplayNameFormat")) {
            $displayTemplates["DefaultDisplayNameFormat"]
        } else {
            ""
        }
        
        Write-DebugMessage "Default DisplayName Format: $defaultDisplayNameFormat"

        if ($defaultDisplayNameFormat) {
            $DisplayNameTemplateList += [PSCustomObject]@{
                Name     = $defaultDisplayNameFormat
                Template = $defaultDisplayNameFormat
            }
        }

        # Add all other DisplayName templates (only those starting with "DisplayNameTemplate")
        if ($displayTemplates -and $displayTemplates.Keys.Count -gt 0) {
            foreach ($key in $displayTemplates.Keys) {
                if ($key -like "DisplayNameTemplate*") {
                    $pattern = $displayTemplates[$key]
                    Write-DebugMessage "DisplayName Template found: $pattern"
                    $DisplayNameTemplateList += [PSCustomObject]@{
                        Name     = $pattern
                        Template = $pattern
                    }
                }
            }
        }

        # Set the ItemsSource, DisplayMemberPath and SelectedValuePath
        $comboBoxDisplayTemplate.ItemsSource = $DisplayNameTemplateList
        $comboBoxDisplayTemplate.DisplayMemberPath = "Name"
        $comboBoxDisplayTemplate.SelectedValuePath = "Template"

        # Set the default value if available
        if ($defaultDisplayNameFormat) {
            $comboBoxDisplayTemplate.SelectedValue = $defaultDisplayNameFormat
        }
        
        Write-DebugMessage "Display name template dropdown populated successfully"
        return $true
    }
    catch {
        Write-DebugMessage "Error initializing display name template dropdown: $($_.Exception.Message)"
        return $false
    }
}
#endregion

#region [Region 14.3 | LICENSE OPTIONS]
# Populates the license dropdown with available license options
function Initialize-LicenseDropdown {
    [CmdletBinding()]
    param()
    
    Write-DebugMessage "Initializing license dropdown"
    
    # Ensure window and config objects exist
    if (-not (Get-Variable -Name window -Scope Global -ErrorAction SilentlyContinue) -or
        -not (Get-Variable -Name Config -Scope Global -ErrorAction SilentlyContinue)) {
        Write-DebugMessage "Error: Required global variables not found"
        return $false
    }
    
    try {
        $comboBoxLicense = $global:window.FindName("cmbLicense")
        if (-not $comboBoxLicense) {
            Write-DebugMessage "ComboBox 'cmbLicense' not found in the window"
            return $false
        }
        
        if (-not $global:Config.Contains("LicensesGroups")) {
            Write-DebugMessage "LicensesGroups section missing in configuration"
            return $false
        }
        
        # Clear bindings and items
        [System.Windows.Data.BindingOperations]::ClearBinding($comboBoxLicense, [System.Windows.Controls.ComboBox]::ItemsSourceProperty)
        $comboBoxLicense.Items.Clear()
        
        $LicenseListFromINI = @()
        foreach ($licenseKey in $global:Config["LicensesGroups"].Keys) {
            $LicenseListFromINI += [PSCustomObject]@{
                Name  = $licenseKey -replace '^MS365_', ''
                Value = $licenseKey
            }
        }
        
        $comboBoxLicense.ItemsSource = $LicenseListFromINI
        $comboBoxLicense.DisplayMemberPath = "Name"
        $comboBoxLicense.SelectedValuePath = "Value"
        
        if ($LicenseListFromINI.Count -gt 0) {
            $comboBoxLicense.SelectedIndex = 0
        }
        
        Write-DebugMessage "License dropdown populated successfully"
        return $true
    }
    catch {
        Write-DebugMessage "Error initializing license dropdown: $($_.Exception.Message)"
        return $false
    }
}
#endregion

#region [Region 14.4 | TEAM LEADER GROUPS]
# Populates the team leader group dropdown from configuration
function Initialize-TeamLeaderGroupDropdown {
    [CmdletBinding()]
    param()
    
    Write-DebugMessage "Initializing team leader group dropdown"
    
    # Ensure window and config objects exist
    if (-not (Get-Variable -Name window -Scope Global -ErrorAction SilentlyContinue) -or
        -not (Get-Variable -Name Config -Scope Global -ErrorAction SilentlyContinue)) {
        Write-DebugMessage "Error: Required global variables not found"
        return $false
    }
    
    try {
        $comboBoxTLGroup = $global:window.FindName("cmbTLGroup")
        if (-not $comboBoxTLGroup) {
            Write-DebugMessage "ComboBox 'cmbTLGroup' not found in the window"
            return $false
        }
        
        if (-not $global:Config.Contains("TLGroups")) {
            Write-DebugMessage "TLGroups section missing in configuration"
            return $false
        }
        
        # Clear old bindings and items
        $comboBoxTLGroup.Items.Clear()
        [System.Windows.Data.BindingOperations]::ClearBinding($comboBoxTLGroup, [System.Windows.Controls.ComboBox]::ItemsSourceProperty)
        
        $TLGroupOptions = @()
        foreach ($key in $global:Config["TLGroups"].Keys) {
            # Here the key (e.g. DEV) is used as display text,
            # while the corresponding value (e.g. TEAM1TL) is stored separately.
            $TLGroupOptions += [PSCustomObject]@{
                Name  = $key
                Group = $global:Config["TLGroups"][$key]
            }
        }
        
        $comboBoxTLGroup.ItemsSource = $TLGroupOptions
        $comboBoxTLGroup.DisplayMemberPath = "Name"
        $comboBoxTLGroup.SelectedValuePath = "Group"
        
        if ($TLGroupOptions.Count -gt 0) {
            $comboBoxTLGroup.SelectedIndex = 0
        }
        
        Write-DebugMessage "Team leader group dropdown populated successfully"
        return $true
    }
    catch {
        Write-DebugMessage "Error initializing team leader group dropdown: $($_.Exception.Message)"
        return $false
    }
}
#endregion

#region [Region 14.5 | EMAIL DOMAIN SUFFIXES]
# Populates the email domain suffix dropdown from configuration
function Initialize-EmailDomainSuffixDropdown {
    [CmdletBinding()]
    param()
    
    Write-DebugMessage "Initializing email domain suffix dropdown"
    
    # Ensure window and config objects exist
    if (-not (Get-Variable -Name window -Scope Global -ErrorAction SilentlyContinue) -or
        -not (Get-Variable -Name Config -Scope Global -ErrorAction SilentlyContinue)) {
        Write-DebugMessage "Error: Required global variables not found"
        return $false
    }
    
    try {
        $comboBoxSuffix = $global:window.FindName("cmbMailSuffix")
        if (-not $comboBoxSuffix) {
            # Try alternative name from original code
            $comboBoxSuffix = $global:window.FindName("cmbSuffix")
            if (-not $comboBoxSuffix) {
                Write-DebugMessage "ComboBox for mail suffix not found in the window"
                return $false
            }
        }
        
        if (-not $global:Config.Contains("MailEndungen")) {
            Write-DebugMessage "MailEndungen section missing in configuration"
            return $false
        }
        
        # Clear binding and items
        $comboBoxSuffix.Items.Clear()
        [System.Windows.Data.BindingOperations]::ClearBinding($comboBoxSuffix, [System.Windows.Controls.ComboBox]::ItemsSourceProperty)
        
        $MailSuffixListFromINI = @()
        foreach ($domainKey in $global:Config["MailEndungen"].Keys) {
            $domainValue = $global:Config["MailEndungen"][$domainKey]
            $MailSuffixListFromINI += [PSCustomObject]@{
                Key   = $domainValue
                Value = $domainValue
            }
        }
        
        $comboBoxSuffix.ItemsSource = $MailSuffixListFromINI
        $comboBoxSuffix.DisplayMemberPath = "Key"
        $comboBoxSuffix.SelectedValuePath = "Value"
        
        # Set default value from the Company section if available
        if ($global:Config.Contains("Company") -and $global:Config.Company.Contains("CompanyMailDomain")) {
            $defaultSuffix = $global:Config.Company["CompanyMailDomain"]
            $comboBoxSuffix.SelectedValue = $defaultSuffix
        } elseif ($MailSuffixListFromINI.Count -gt 0) {
            $comboBoxSuffix.SelectedIndex = 0
        }
        
        Write-DebugMessage "Email domain suffix dropdown populated successfully"
        return $true
    }
    catch {
        Write-DebugMessage "Error initializing email domain suffix dropdown: $($_.Exception.Message)"
        return $false
    }
}
#endregion

#region [Region 14.6 | LOCATION OPTIONS]
# Populates the location dropdown with available office locations
function Initialize-LocationDropdown {
    [CmdletBinding()]
    param()
    
    Write-DebugMessage "Initializing location dropdown"
    
    # Ensure window and config objects exist
    if (-not (Get-Variable -Name window -Scope Global -ErrorAction SilentlyContinue) -or
        -not (Get-Variable -Name Config -Scope Global -ErrorAction SilentlyContinue)) {
        Write-DebugMessage "Error: Required global variables not found"
        return $false
    }
    
    try {
        if (-not $global:Config.Contains("STANDORTE")) {
            Write-DebugMessage "STANDORTE section missing in configuration"
            return $false
        }
        
        $locationList = @()
        foreach ($key in $global:Config["STANDORTE"].Keys) {
            if ($key -match '^(STANDORTE_\d+)$') {
                $bezKey = $key + "_Bez"
                $locationObj = [PSCustomObject]@{
                    Key = $global:Config["STANDORTE"][$key]
                    Bez = if ($global:Config["STANDORTE"].Contains($bezKey)) { 
                        $global:Config["STANDORTE"][$bezKey] 
                    } else { 
                        $global:Config["STANDORTE"][$key] 
                    }
                }
                $locationList += $locationObj
            }
        }
        
        # Make location list available globally
        $global:LocationListFromINI = $locationList
        
        $cmbLocation = $global:window.FindName("cmbLocation")
        if (-not $cmbLocation) {
            Write-DebugMessage "ComboBox 'cmbLocation' not found in the window"
            return $false
        }
        
        # Clear binding and items
        $cmbLocation.Items.Clear()
        [System.Windows.Data.BindingOperations]::ClearBinding($cmbLocation, [System.Windows.Controls.ComboBox]::ItemsSourceProperty)
        
        $cmbLocation.ItemsSource = $global:LocationListFromINI
        $cmbLocation.DisplayMemberPath = "Bez"
        $cmbLocation.SelectedValuePath = "Key"
        
        if ($global:LocationListFromINI.Count -gt 0) {
            $cmbLocation.SelectedIndex = 0
        }
        
        Write-DebugMessage "Location dropdown populated successfully"
        return $true
    }
    catch {
        Write-DebugMessage "Error initializing location dropdown: $($_.Exception.Message)"
        return $false
    }
}
#endregion

#region [Initialize All Dropdowns]
# Initializes all dropdowns with a single function call
function Initialize-AllDropdowns {
    [CmdletBinding()]
    param()
    
    Write-DebugMessage "Initializing all dropdowns"
    
    $results = @{
        DisplayNameTemplate = Initialize-DisplayNameTemplateDropdown
        License = Initialize-LicenseDropdown
        TeamLeaderGroup = Initialize-TeamLeaderGroupDropdown
        EmailDomainSuffix = Initialize-EmailDomainSuffixDropdown
        Location = Initialize-LocationDropdown
    }
    
    # Register OU dropdown refresh event separately
    $results.OUDropdownRefresh = Register-OUDropdownRefreshEvent
    
    # Count successful initializations
    $successCount = ($results.Values | Where-Object { $_ -eq $true }).Count
    $totalCount = $results.Count
    
    Write-DebugMessage "Dropdown initialization complete: $successCount of $totalCount successful"
    
    return $results
}
#endregion

# Export module functions
Export-ModuleMember -Function Register-OUDropdownRefreshEvent, 
                             Initialize-DisplayNameTemplateDropdown,
                             Initialize-LicenseDropdown,
                             Initialize-TeamLeaderGroupDropdown,
                             Initialize-EmailDomainSuffixDropdown,
                             Initialize-LocationDropdown,
                             Initialize-AllDropdowns
#endregion