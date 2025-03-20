#region [Region 14 | DROPDOWN POPULATION]
# Functions to populate various dropdown menus from configuration data

Write-DebugMessage "Dropdown: OU Refresh"

#region [Region 14.1 | OU DROPDOWN]
# Populates the Organizational Unit dropdown with data from AD
# --- OU Refresh --- 
# [14.1.1 - Fetches OUs from beneath the configured default OU]
Write-DebugMessage "Refreshing OU dropdown."
$btnRefreshOU = $window.FindName("btnRefreshOU")
if ($btnRefreshOU) {
    $btnRefreshOU.Add_Click({
        try {
            $defaultOUFromINI = $global:Config.ADUserDefaults["DefaultOU"]
            $OUList = Get-ADOrganizationalUnit -Filter * -SearchBase $defaultOUFromINI |
                Select-Object -ExpandProperty DistinguishedName
                Write-DebugMessage "Found OUs: $($OUList.Count)"
            $cmbOU = $window.FindName("cmbOU")
            if ($cmbOU) {
                # First set ItemsSource to null
                $cmbOU.ItemsSource = $null
                # Then clear the ItemsCollection
                $cmbOU.Items.Clear()
                # Now assign the new ItemsSource
                $cmbOU.ItemsSource = $OUList
                Write-DebugMessage "Dropdown: OU-DropDown successfully populated."
            }
        }
        catch {
            [System.Windows.MessageBox]::Show("Error loading OUs: $($_.Exception.Message)", "Error")
        }
    })
}
#endregion

Write-DebugMessage "Dropdown: DisplayName Dropdown populate"

#region [Region 14.2 | DISPLAY NAME TEMPLATES]
# Populates the display name template dropdown from INI configuration
Write-DebugMessage "Populating display name template dropdown."
$comboBoxDisplayTemplate = $window.FindName("cmbDisplayTemplate")

if ($comboBoxDisplayTemplate -and $global:Config.Contains("DisplayNameUPNTemplates")) {
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
}
Write-DebugMessage "Display name template dropdown populated successfully."
#endregion

Write-DebugMessage "Dropdown: License Dropdown populate"

#region [Region 14.3 | LICENSE OPTIONS]
# Populates the license dropdown with available license options
Write-DebugMessage "Populating license dropdown."
$comboBoxLicense = $window.FindName("cmbLicense")
if ($comboBoxLicense -and $global:Config.Contains("LicensesGroups")) {
    [System.Windows.Data.BindingOperations]::ClearBinding($comboBoxLicense, [System.Windows.Controls.ComboBox]::ItemsSourceProperty)
    $LicenseListFromINI = @()
    foreach ($licenseKey in $global:Config["LicensesGroups"].Keys) {
        $LicenseListFromINI += [PSCustomObject]@{
            Name  = $licenseKey -replace '^MS365_', ''
            Value = $licenseKey
        }
    }
    $comboBoxLicense.ItemsSource = $LicenseListFromINI
    if ($LicenseListFromINI.Count -gt 0) {
        $comboBoxLicense.SelectedValue = $LicenseListFromINI[0].Value
    }
}
Write-DebugMessage "License dropdown populated successfully."
#endregion

Write-DebugMessage "Dropdown: TLGroups Dropdown populate"

#region [Region 14.4 | TEAM LEADER GROUPS]
# Populates the team leader group dropdown from configuration
Write-DebugMessage "Populating team leader group dropdown."
$comboBoxTLGroup = $window.FindName("cmbTLGroup")
if ($comboBoxTLGroup -and $global:Config.Contains("TLGroups")) {
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
}
Write-DebugMessage "Team leader group dropdown populated successfully."
#endregion

Write-DebugMessage "Dropdown: MailSuffix Dropdown populate"

#region [Region 14.5 | EMAIL DOMAIN SUFFIXES]
# Populates the email domain suffix dropdown from configuration
Write-DebugMessage "Populating email domain suffix dropdown."
$comboBoxSuffix = $window.FindName("cmbSuffix")
if ($comboBoxSuffix -and $global:Config.Contains("MailEndungen")) {
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
    # Default value from the [Company] section:
    $defaultSuffix = $global:Config.Company["CompanyMailDomain"]
    $comboBoxSuffix.SelectedValue = $defaultSuffix
}
Write-DebugMessage "Email domain suffix dropdown populated successfully."
#endregion

Write-DebugMessage "Dropdown: Location Dropdown populate"

#region [Region 14.6 | LOCATION OPTIONS]
# Populates the location dropdown with available office locations
Write-DebugMessage "Populating location dropdown."
if ($global:Config.Contains("STANDORTE")) {
    $locationList = @()
    foreach ($key in $global:Config["STANDORTE"].Keys) {
        if ($key -match '^(STANDORTE_\d+)$') {
            $bezKey = $key + "_Bez"
            $locationObj = [PSCustomObject]@{
                Key = $global:Config["STANDORTE"][$key]
                Bez = if ($global:Config["STANDORTE"].Contains($bezKey)) { $global:Config["STANDORTE"][$bezKey] } else { $global:Config["STANDORTE"][$key] }
            }
            $locationList += $locationObj
        }
    }
    $global:LocationListFromINI = $locationList
    $cmbLocation = $window.FindName("cmbLocation")
    if ($cmbLocation -and $global:LocationListFromINI) {
        [System.Windows.Data.BindingOperations]::ClearBinding($cmbLocation, [System.Windows.Controls.ComboBox]::ItemsSourceProperty)
        $cmbLocation.ItemsSource = $global:LocationListFromINI
        if ($global:LocationListFromINI.Count -gt 0) {
            $cmbLocation.SelectedIndex = 0
        }
    }
    Write-DebugMessage "Location dropdown populated successfully."
}
#endregion
#endregion