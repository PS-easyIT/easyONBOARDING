#region [Region 14.2.1 | UPN TEMPLATE DISPLAY]
# Updates the UPN template display with the current template from INI
Write-DebugMessage "Initializing UPN Template Display."
function Update-UPNTemplateDisplay {
    [CmdletBinding()]
    param()
    
    try {
        $txtUPNTemplateDisplay = $window.FindName("txtUPNTemplateDisplay")
        if ($null -eq $txtUPNTemplateDisplay) {
            Write-DebugMessage "UPN Template Display TextBlock not found in XAML"
            return
        }
        
        # Default UPN template format
        $defaultUPNFormat = "FIRSTNAME.LASTNAME"
        
        # Read from INI if available
        if ($global:Config.Contains("DisplayNameUPNTemplates") -and 
            $global:Config.DisplayNameUPNTemplates.Contains("DefaultUserPrincipalNameFormat")) {
            $upnTemplate = $global:Config.DisplayNameUPNTemplates.DefaultUserPrincipalNameFormat
            if (-not [string]::IsNullOrWhiteSpace($upnTemplate)) {
                $defaultUPNFormat = $upnTemplate.ToUpper()
                Write-DebugMessage "Found UPN template in INI: $defaultUPNFormat"
            }
        }
        
        # Create viewmodel for binding
        if (-not $global:viewModel) {
            $global:viewModel = [PSCustomObject]@{
                UPNTemplate = $defaultUPNFormat
            }
        }
        else {
            $global:viewModel.UPNTemplate = $defaultUPNFormat
        }
        
        # Set DataContext for binding
        $txtUPNTemplateDisplay.DataContext = $global:viewModel
        
        Write-DebugMessage "UPN Template Display initialized with: $defaultUPNFormat"
    }
    catch {
        Write-DebugMessage "Error updating UPN Template Display: $($_.Exception.Message)"
    }
}

# Call the function to initialize the UPN template display
Update-UPNTemplateDisplay

# Make sure UPN Template Display gets updated when dropdown selection changes
$comboBoxDisplayTemplate = $window.FindName("cmbDisplayTemplate")
if ($comboBoxDisplayTemplate) {
    $comboBoxDisplayTemplate.Add_SelectionChanged({
        $selectedTemplate = $comboBoxDisplayTemplate.SelectedValue
        if ($global:viewModel -and -not [string]::IsNullOrWhiteSpace($selectedTemplate)) {
            $global:viewModel.UPNTemplate = $selectedTemplate.ToUpper()
            Write-DebugMessage "UPN Template updated to: $selectedTemplate"
        }
    })
}
#endregion
