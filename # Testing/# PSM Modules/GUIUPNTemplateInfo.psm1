# Check if Write-DebugMessage function exists, if not create a simple version for module testing
if (-not (Get-Command -Name Write-DebugMessage -ErrorAction SilentlyContinue)) {
    function Write-DebugMessage {
        [CmdletBinding()]
        param([Parameter(Mandatory=$true, Position=0)][string]$Message)
        Write-Verbose $Message
    }
}

Write-DebugMessage "Initializing UPN Template Display."

# Updates the UPN template display with the current template from INI
function Update-UPNTemplateDisplay {
    [CmdletBinding()]
    param(
        [Parameter(ValueFromPipeline=$false)]
        [System.Windows.Controls.ComboBox]$ComboBox
    )
    
    try {
        # Validate $window is available
        if ($null -eq $window) {
            Write-DebugMessage "Window object not found. Make sure this module is loaded after the main window is initialized."
            return
        }
        
        $txtUPNTemplateDisplay = $window.FindName("txtUPNTemplateDisplay")
        if ($null -eq $txtUPNTemplateDisplay) {
            Write-DebugMessage "UPN Template Display TextBlock not found in XAML"
            return
        }
        
        # Default UPN template format
        $defaultUPNFormat = "FIRSTNAME.LASTNAME"
        
        # Read from INI if available
        if ($null -ne $global:Config -and $global:Config.Contains("DisplayNameUPNTemplates") -and 
            $global:Config.DisplayNameUPNTemplates.Contains("DefaultUserPrincipalNameFormat")) {
            $upnTemplate = $global:Config.DisplayNameUPNTemplates.DefaultUserPrincipalNameFormat
            if (-not [string]::IsNullOrWhiteSpace($upnTemplate)) {
                $defaultUPNFormat = $upnTemplate.ToUpper()
                Write-DebugMessage "Found UPN template in INI: $defaultUPNFormat"
            }
        } else {
            Write-DebugMessage "Using default UPN format: $defaultUPNFormat"
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
        # Re-throw the exception if we're in verbose mode for better debugging
        if ($VerbosePreference -eq 'Continue') {
            throw
        }
    }
}

# Function to handle combobox selection changes
function Register-UPNTemplateSelectionHandler {
    [CmdletBinding()]
    param()
    
    try {
        # Get the combobox control
        $comboBoxDisplayTemplate = $window.FindName("cmbDisplayTemplate")
        if ($null -eq $comboBoxDisplayTemplate) {
            Write-DebugMessage "ComboBox control cmbDisplayTemplate not found in XAML"
            return
        }
        
        # Remove previous event handlers to prevent duplicates
        $comboBoxDisplayTemplate.add_SelectionChanged({
            $selectedTemplate = $comboBoxDisplayTemplate.SelectedValue
            if ($global:viewModel -and -not [string]::IsNullOrWhiteSpace($selectedTemplate)) {
                $global:viewModel.UPNTemplate = $selectedTemplate.ToUpper()
                Write-DebugMessage "UPN Template updated to: $selectedTemplate"
            }
        })
        
        Write-DebugMessage "UPN Template selection handler registered"
    }
    catch {
        Write-DebugMessage "Error registering UPN Template selection handler: $($_.Exception.Message)"
    }
}

# Initialize the module
function Initialize-UPNTemplateModule {
    [CmdletBinding()]
    param()
    
    Write-DebugMessage "Initializing UPN Template module"
    Update-UPNTemplateDisplay
    Register-UPNTemplateSelectionHandler
}

# Call the initialization function
Initialize-UPNTemplateModule

# Export module members
Export-ModuleMember -Function Update-UPNTemplateDisplay, Register-UPNTemplateSelectionHandler
