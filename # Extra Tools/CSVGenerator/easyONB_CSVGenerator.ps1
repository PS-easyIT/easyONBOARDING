[CmdletBinding()]
param(
    [switch]$DebugOutput,
    [bool]$AccountDisabled = $true
)

# General settings
$DebugPreference = 'SilentlyContinue' # or 'Continue' if you want to see non-terminating errors
$Debug = $PSBoundParameters.ContainsKey('DebugOutput') # Use the switch to control debugging

#############################################
# Helper functions for debugging and logging
#############################################
function Write-DebugMessage {
    param(
        [string]$Message
    )
    if ($DebugOutput -or $Debug) {
        Write-Host "[DEBUG] $Message" -ForegroundColor Cyan
    }
}

function Write-Log {
    param(
        [string]$Message,
        [string]$Level = "INFO"
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"
    
    # Console output based on level
    switch ($Level) {
        "ERROR" { Write-Host $logMessage -ForegroundColor Red }
        "WARNING" { Write-Host $logMessage -ForegroundColor Yellow }
        "INFO" { Write-Host $logMessage -ForegroundColor Green }
        default { Write-Host $logMessage }
    }
}

Write-DebugMessage "Starting script with AccountDisabled=$AccountDisabled"

#############################################
# Parameters - Set all settings here
#############################################
# CSV settings
$csvFolder = "C:\easyIT\DATA\easyONBOARDING\CSVData"
$csvFileName = "HROnboardingData.csv"

# GUI settings
$fontSize = 10
# Using named color from System.Windows.Media.Colors instead of hex code
# See: https://learn.microsoft.com/de-de/dotnet/api/system.windows.media.colors
$formBackColor = "Silver"

# Logo settings
$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
$headerLogo = Join-Path -Path $scriptPath -ChildPath "APPICON1.PNG" # Path to logo

Write-DebugMessage "Parameters set: CSV path=$csvFolder, File=$csvFileName, Logo=$headerLogo"

# Load required assemblies
try {
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    Add-Type -AssemblyName PresentationFramework
    Add-Type -AssemblyName PresentationCore
    Add-Type -AssemblyName WindowsBase
    Write-DebugMessage "Windows.Forms, Drawing, PresentationFramework, PresentationCore and WindowsBase assemblies loaded"
}
catch {
    Write-Log "Error loading assemblies: $_" -Level "ERROR"
    [System.Windows.Forms.MessageBox]::Show("Error loading assemblies: $_", "Error", 
        [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    exit
}

# Check if CSV folder exists, create it if needed
try {
    if (-not (Test-Path $csvFolder)) {
        Write-DebugMessage "CSV folder not found. Creating folder: $csvFolder"
        New-Item -ItemType Directory -Path $csvFolder -Force | Out-Null
        Write-Log "CSV folder created: $csvFolder" -Level "INFO"
    } else {
        Write-DebugMessage "CSV folder already exists: $csvFolder"
    }
    $csvFile = Join-Path -Path $csvFolder -ChildPath $csvFileName
    Write-DebugMessage "Full CSV path: $csvFile"
} catch {
    Write-Log "Error creating CSV folder ($csvFolder): $_" -Level "ERROR"
    [System.Windows.Forms.MessageBox]::Show("Error creating CSV folder ($csvFolder): $_", "Error", 
        [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        exit
}

#############################################
# Load GUI from XAML file
#############################################
Write-DebugMessage "Loading GUI from MainGUI.xaml"
$useXamlGUI = $true
try {
    # Get the XAML file path
    $xamlPath = Join-Path -Path $scriptPath -ChildPath "MainGUI.xaml"
    
    if (-not (Test-Path $xamlPath)) {
        throw "XAML file not found at: $xamlPath"
    }
    
    # Load the XAML content
    [xml]$xaml = Get-Content -Path $xamlPath -Raw -ErrorAction Stop
    
    # Check for potential ResourceDictionary issues
    $resourceDicts = $xaml.SelectNodes("//*[local-name()='ResourceDictionary']")
    if ($resourceDicts.Count -gt 0) {
        Write-DebugMessage "Found $($resourceDicts.Count) ResourceDictionary elements - checking for duplicate keys"
        foreach ($dict in $resourceDicts) {
            $keys = @{} 
            $duplicates = @()
            foreach ($resource in $dict.ChildNodes) {
                if ($resource.Key -and $keys.ContainsKey($resource.Key)) {
                    $duplicates += $resource.Key
                } else {
                    $keys[$resource.Key] = $true
                }
            }
            if ($duplicates.Count -gt 0) {
                Write-Log "Warning: Found duplicate keys in ResourceDictionary: $($duplicates -join ', ')" -Level "WARNING"
            }
        }
    }
    
    # Create the XAML reader
    $reader = New-Object System.Xml.XmlNodeReader $xaml
    
    try {
        # Load the window
        $window = [Windows.Markup.XamlReader]::Load($reader)
        Write-DebugMessage "GUI loaded successfully"
    }
    catch [System.Exception] {
        Write-Log "Error loading XAML: $($_.Exception.Message)" -Level "ERROR"
        Write-Log "Inner exception: $($_.Exception.InnerException.Message)" -Level "ERROR"
        $useXamlGUI = $false
        throw
    }
    
    # Access controls
    $controls = @{
        "txtFirstName" = $window.FindName("txtFirstName")
        "txtLastName" = $window.FindName("txtLastName")
        "chkExternal" = $window.FindName("chkExternal")
        "txtExtCompany" = $window.FindName("txtExtCompany")
        "txtDescription" = $window.FindName("txtDescription")
        "cmbOffice" = $window.FindName("cmbOffice")
        "txtPhone" = $window.FindName("txtPhone")
        "txtMobile" = $window.FindName("txtMobile")
        "txtMail" = $window.FindName("txtMail")
        "txtPosition" = $window.FindName("txtPosition")
        "cmbBusinessUnit" = $window.FindName("cmbBusinessUnit")
        "txtPersonalNumber" = $window.FindName("txtPersonalNumber")
        "dtpTermination" = $window.FindName("dtpTermination")
        "dtpStartWorkDate" = $window.FindName("dtpStartWorkDate")  # Add StartWorkDate control
        "chkTL" = $window.FindName("chkTL")
        "chkAL" = $window.FindName("chkAL")
        "btnSave" = $window.FindName("btnSave")
        "btnClose" = $window.FindName("btnClose")
        "picLogo" = $window.FindName("picLogo")
        # New status controls
        "cmbStatus" = $window.FindName("cmbStatus")
        "chkProcessed" = $window.FindName("chkProcessed")
        "txtProcessedBy" = $window.FindName("txtProcessedBy")
        "txtNotes" = $window.FindName("txtNotes")
    }
    
    # Define required vs optional controls
    $requiredControls = @(
        "txtFirstName", "txtLastName", "chkExternal", "txtExtCompany", "txtDescription",
        "cmbOffice", "txtPhone", "txtMobile", "txtMail", "txtPosition", "cmbBusinessUnit",
        "txtPersonalNumber", "dtpTermination", "dtpStartWorkDate", "chkTL", "chkAL", 
        "btnSave", "btnClose", "picLogo", "chkProcessed", "txtNotes"
    )
    
    $optionalControls = @("cmbStatus", "txtProcessedBy")
    
    # Validate required controls were found
    $missingRequiredControls = $requiredControls | Where-Object { $controls[$_] -eq $null } 
    if ($missingRequiredControls) {
        throw "Could not find the following required controls in XAML: $($missingRequiredControls -join ', ')"
    }
    
    # Log warnings for missing optional controls but continue execution
    $missingOptionalControls = $optionalControls | Where-Object { $controls[$_] -eq $null }
    if ($missingOptionalControls) {
        Write-Log "Some optional controls were not found in XAML: $($missingOptionalControls -join ', ')" -Level "WARNING"
        # Initialize missing controls to null so the rest of the code can check for their existence
        foreach ($controlName in $missingOptionalControls) {
            Write-DebugMessage "Setting $controlName to null (optional control)"
            $controls[$controlName] = $null
        }
    }
    
    # Optional: Set default values for DatePickers
    if ($controls.dtpTermination) {
        $controls.dtpTermination.SelectedDate = [DateTime]::Now.AddYears(1)
    }
    if ($controls.dtpStartWorkDate) {
        $controls.dtpStartWorkDate.SelectedDate = [DateTime]::Now.AddDays(14)
    }
    
    # Load logo if exists
    if (Test-Path $headerLogo) {
        try {
            $imageSource = New-Object System.Windows.Media.Imaging.BitmapImage
            $imageSource.BeginInit()
            $imageSource.CacheOption = [System.Windows.Media.Imaging.BitmapCacheOption]::OnLoad
            $imageSource.UriSource = New-Object System.Uri($headerLogo, [System.UriKind]::Absolute)
            $imageSource.EndInit()
            $controls.picLogo.Source = $imageSource
            Write-DebugMessage "Logo loaded successfully from $headerLogo"
        }
        catch {
            Write-Log "Error loading logo: $_" -Level "WARNING"
        }
    } else {
        Write-Log "Logo file not found at $headerLogo. Skipping logo loading." -Level "WARNING"
    }
} 
catch [System.Xml.XmlException] {
    Write-Log "XML parsing error in XAML file: $($_.Exception.Message)" -Level "ERROR"
    $useXamlGUI = $false
}
catch {
    Write-Log "Error loading or accessing GUI elements: $_" -Level "ERROR"
    Write-Log "Exception type: $($_.Exception.GetType().FullName)" -Level "ERROR"
    
    if ($_.Exception.InnerException) {
        Write-Log "Inner exception: $($_.Exception.InnerException.Message)" -Level "ERROR"
    }
    
    $useXamlGUI = $false
}

# Create fallback UI if XAML loading failed
if (-not $useXamlGUI) {
    Write-Log "Creating fallback UI due to XAML loading failure" -Level "WARNING"
    try {
        # Create a basic Windows Forms UI instead
        $form = New-Object System.Windows.Forms.Form
        $form.Text = "Easy Onboarding CSV Generator v1.3 (Fallback UI)"
        $form.Size = New-Object System.Drawing.Size(500, 650)
        $form.StartPosition = "CenterScreen"
        $form.BackColor = [System.Drawing.Color]::$formBackColor
        $form.Font = New-Object System.Drawing.Font("Segoe UI", $fontSize)

        # Create panel for scrolling
        $panel = New-Object System.Windows.Forms.Panel
        $panel.Dock = [System.Windows.Forms.DockStyle]::Fill
        $panel.AutoScroll = $true
        $form.Controls.Add($panel)

        $currentY = 20
        $labelWidth = 120
        $controlWidth = 250
        $spacing = 30
        $controlHeight = 23

        # Helper function to add a labeled control
        function Add-LabeledControl {
            param(
                [string]$LabelText,
                [System.Windows.Forms.Control]$Control,
                [int]$Y,
                [switch]$FullWidth
            )
            
            $label = New-Object System.Windows.Forms.Label
            $label.Location = New-Object System.Drawing.Point(20, $Y)
            $label.Size = New-Object System.Drawing.Size($labelWidth, $controlHeight)
            $label.Text = $LabelText
            $panel.Controls.Add($label)
            
            $controlX = 20 + $labelWidth + 10
            $actualWidth = if ($FullWidth) { 390 } else { $controlWidth }
            $Control.Location = New-Object System.Drawing.Point($controlX, $Y)
            $Control.Size = New-Object System.Drawing.Size($actualWidth, $controlHeight)
            $panel.Controls.Add($Control)
            
            return $Y + $spacing
        }

        # Create form controls
        $txtFirstName = New-Object System.Windows.Forms.TextBox
        $currentY = Add-LabeledControl -LabelText "First Name:" -Control $txtFirstName -Y $currentY
        
        $txtLastName = New-Object System.Windows.Forms.TextBox
        $currentY = Add-LabeledControl -LabelText "Last Name:" -Control $txtLastName -Y $currentY
        
        $chkExternal = New-Object System.Windows.Forms.CheckBox
        $chkExternal.Text = ""
        $currentY = Add-LabeledControl -LabelText "External:" -Control $chkExternal -Y $currentY
        
        $txtExtCompany = New-Object System.Windows.Forms.TextBox
        $currentY = Add-LabeledControl -LabelText "External Company:" -Control $txtExtCompany -Y $currentY
        
        $txtDescription = New-Object System.Windows.Forms.TextBox
        $currentY = Add-LabeledControl -LabelText "Description:" -Control $txtDescription -Y $currentY
        
        $cmbOffice = New-Object System.Windows.Forms.ComboBox
        $cmbOffice.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDown
        $currentY = Add-LabeledControl -LabelText "Office:" -Control $cmbOffice -Y $currentY
        
        $txtPhone = New-Object System.Windows.Forms.TextBox
        $currentY = Add-LabeledControl -LabelText "Phone:" -Control $txtPhone -Y $currentY
        
        $txtMobile = New-Object System.Windows.Forms.TextBox
        $currentY = Add-LabeledControl -LabelText "Mobile:" -Control $txtMobile -Y $currentY
        
        $txtMail = New-Object System.Windows.Forms.TextBox
        $currentY = Add-LabeledControl -LabelText "Email:" -Control $txtMail -Y $currentY
        
        $txtPosition = New-Object System.Windows.Forms.TextBox
        $currentY = Add-LabeledControl -LabelText "Position:" -Control $txtPosition -Y $currentY
        
        $cmbBusinessUnit = New-Object System.Windows.Forms.ComboBox
        $cmbBusinessUnit.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDown
        $currentY = Add-LabeledControl -LabelText "Business Unit:" -Control $cmbBusinessUnit -Y $currentY
        
        $txtPersonalNumber = New-Object System.Windows.Forms.TextBox
        $currentY = Add-LabeledControl -LabelText "Personal Number:" -Control $txtPersonalNumber -Y $currentY
        
        $dtpTermination = New-Object System.Windows.Forms.DateTimePicker
        $dtpTermination.Value = [DateTime]::Now.AddYears(1)
        $currentY = Add-LabeledControl -LabelText "Termination Date:" -Control $dtpTermination -Y $currentY
        
        $chkTL = New-Object System.Windows.Forms.CheckBox
        $chkTL.Text = ""
        $currentY = Add-LabeledControl -LabelText "Team Leader:" -Control $chkTL -Y $currentY
        
        $chkAL = New-Object System.Windows.Forms.CheckBox
        $chkAL.Text = ""
        $currentY = Add-LabeledControl -LabelText "Department Head:" -Control $chkAL -Y $currentY
        
        # Add Status fields to fallback UI
        $cmbStatus = New-Object System.Windows.Forms.ComboBox
        $cmbStatus.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDown
        $cmbStatus.Items.AddRange(@("Neu", "In Bearbeitung", "Abgeschlossen", "Storniert"))
        $cmbStatus.SelectedIndex = 0
        $currentY = Add-LabeledControl -LabelText "Status:" -Control $cmbStatus -Y $currentY
        
        $chkProcessed = New-Object System.Windows.Forms.CheckBox
        $chkProcessed.Text = ""
        $currentY = Add-LabeledControl -LabelText "Bearbeitet:" -Control $chkProcessed -Y $currentY
        
        $txtProcessedBy = New-Object System.Windows.Forms.TextBox
        $currentY = Add-LabeledControl -LabelText "Bearbeitet von:" -Control $txtProcessedBy -Y $currentY
        
        $txtNotes = New-Object System.Windows.Forms.TextBox
        $txtNotes.Multiline = $true
        $txtNotes.Height = 40
        $currentY = Add-LabeledControl -LabelText "Notizen:" -Control $txtNotes -Y $currentY
        
        # Add to controls dictionary
        $controls["cmbStatus"] = $cmbStatus
        $controls["chkProcessed"] = $chkProcessed
        $controls["txtProcessedBy"] = $txtProcessedBy
        $controls["txtNotes"] = $txtNotes
        
        # Add buttons
        $currentY += 20
        $btnSave = New-Object System.Windows.Forms.Button
        $btnSave.Location = New-Object System.Drawing.Point(120, $currentY)
        $btnSave.Size = New-Object System.Drawing.Size(100, 30)
        $btnSave.Text = "Save to CSV"
        $panel.Controls.Add($btnSave)
        
        $btnClose = New-Object System.Windows.Forms.Button
        $btnClose.Location = New-Object System.Drawing.Point(240, $currentY)
        $btnClose.Size = New-Object System.Drawing.Size(100, 30)
        $btnClose.Text = "Close"
        $panel.Controls.Add($btnClose)
        
        # Create a controls dictionary to match the XAML structure
        $controls = @{
            "txtFirstName" = $txtFirstName
            "txtLastName" = $txtLastName
            "chkExternal" = $chkExternal
            "txtExtCompany" = $txtExtCompany
            "txtDescription" = $txtDescription
            "cmbOffice" = $cmbOffice
            "txtPhone" = $txtPhone
            "txtMobile" = $txtMobile
            "txtMail" = $txtMail
            "txtPosition" = $txtPosition
            "cmbBusinessUnit" = $cmbBusinessUnit
            "txtPersonalNumber" = $txtPersonalNumber
            "dtpTermination" = $dtpTermination
            "chkTL" = $chkTL
            "chkAL" = $chkAL
            "btnSave" = $btnSave
            "btnClose" = $btnClose
            "picLogo" = $null # No logo in fallback UI
            "cmbStatus" = $cmbStatus
            "chkProcessed" = $chkProcessed
            "txtProcessedBy" = $txtProcessedBy
            "txtNotes" = $txtNotes
        }

        # Attach event handlers (we'll define these later)
        $btnSave.Add_Click({
            # The Save button handler will be attached below - content will be the same
        })
        
        $btnClose.Add_Click({
            $form.Close()
        })
        
        # Set the form as our "window" variable to keep the rest of the code working
        $window = $form
        
        Write-Log "Fallback UI created successfully" -Level "INFO"
    }
    catch {
        Write-Log "Error creating fallback UI: $_" -Level "ERROR"
        [System.Windows.Forms.MessageBox]::Show("Critical error: Could not create UI interface. $_", "Fatal Error", 
            [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        exit
    }
}

#############################################
# Event Handlers
#############################################
# Function to validate input fields
function Validate-InputFields {
    $isValid = $true
    $errorMessage = @()
    
    # Check required fields
    if ([string]::IsNullOrWhiteSpace($controls.txtFirstName.Text)) {
        $isValid = $false
        $errorMessage += "- First Name is required"
    }
    if ([string]::IsNullOrWhiteSpace($controls.txtLastName.Text)) {
        $isValid = $false
        $errorMessage += "- Last Name is required"
    }
    
    # Validate phone numbers (if provided)
    if (-not [string]::IsNullOrWhiteSpace($controls.txtPhone.Text) -and 
        ($controls.txtPhone.Text -notmatch "^[\d\+\-\(\) \.]+$")) {
        $isValid = $false
        $errorMessage += "- Phone number contains invalid characters"
    }
    if (-not [string]::IsNullOrWhiteSpace($controls.txtMobile.Text) -and 
        ($controls.txtMobile.Text -notmatch "^[\d\+\-\(\) \.]+$")) {
        $isValid = $false
        $errorMessage += "- Mobile number contains invalid characters"
    }
    
    # Email validation - allowing partial addresses (just username part is fine)
    if (-not [string]::IsNullOrWhiteSpace($controls.txtMail.Text) -and 
        ($controls.txtMail.Text -notmatch "^[\w\-\.]+(@[\w\-\.]+)?$")) {
        $isValid = $false
        $errorMessage += "- Email username contains invalid characters"
    }
    
    return @{
        IsValid = $isValid
        ErrorMessage = ($errorMessage -join "`n")
    }
}

# Save button click handler
$controls.btnSave.Add_Click({
    Write-DebugMessage "Button 'Save to CSV' was clicked"
    try {
        # Validate required fields
        Write-DebugMessage "Validating required fields..."
        $validation = Validate-InputFields
        if (-not $validation.IsValid) {
            Write-Log "Validation error: $($validation.ErrorMessage)" -Level "ERROR"
            [System.Windows.Forms.MessageBox]::Show(
                "Please correct the following errors:`n$($validation.ErrorMessage)", 
                "Validation Error", 
                [System.Windows.Forms.MessageBoxButtons]::OK, 
                [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }
        
        # Format dates
        $terminationDate = $null
        if ($controls.dtpTermination -and $controls.dtpTermination.SelectedDate) {
            $terminationDate = $controls.dtpTermination.SelectedDate.ToString("yyyy-MM-dd")
            Write-DebugMessage "Termination Date: $terminationDate"
        }
        
        $startWorkDate = $null
        if ($controls.dtpStartWorkDate -and $controls.dtpStartWorkDate.SelectedDate) {
            $startWorkDate = $controls.dtpStartWorkDate.SelectedDate.ToString("yyyy-MM-dd")
            Write-DebugMessage "Start Work Date: $startWorkDate"
        }
        
        # Create CSV header if file doesn't exist
        $headerExists = Test-Path $csvFile
        
        # Create data object with fields in the specific order
        Write-DebugMessage "Creating data object for CSV export..."
        $data = [PSCustomObject]@{
            'FirstName'      = $controls.txtFirstName.Text.Trim()
            'LastName'       = $controls.txtLastName.Text.Trim()
            'Description'    = $controls.txtDescription.Text.Trim()
            'OfficeRoom'     = $controls.cmbOffice.Text.Trim()
            'PhoneNumber'    = $controls.txtPhone.Text.Trim()
            'MobileNumber'   = $controls.txtMobile.Text.Trim()
            'Position'       = $controls.txtPosition.Text.Trim()
            'DepartmentField'= $controls.cmbBusinessUnit.Text.Trim()
            'EmailAddress'   = $controls.txtMail.Text.Trim()
            'Ablaufdatum'    = $terminationDate
            'StartWorkDate'  = $startWorkDate
            'PersonalNumber' = $controls.txtPersonalNumber.Text.Trim()
            'External'       = $controls.chkExternal.IsChecked
            'ExternalCompany'= if ($controls.chkExternal.IsChecked) { $controls.txtExtCompany.Text.Trim() } else { "" }
            'TL'             = $controls.chkTL.IsChecked
            'AL'             = $controls.chkAL.IsChecked
            'AccountDisabled'= $AccountDisabled
            'ADGroup'        = $controls.cmbBusinessUnit.Text.Trim() 
            'AdminAccount'   = $controls.chkProcessed.IsChecked 
            'Notes'          = $controls.txtNotes.Text.Trim()
        }
        
        # Write to CSV file with correct encoding
        if (-not $headerExists) {
            # Create a new file with headers
            $data | Export-Csv -Path $csvFile -NoTypeInformation -Encoding UTF8
            Write-DebugMessage "Created new CSV file with headers"
        } else {
            # Append to existing file
            $data | Export-Csv -Path $csvFile -NoTypeInformation -Append -Encoding UTF8
            Write-DebugMessage "Appended data to existing CSV file"
        }
        
        Write-Log "Data was successfully saved to CSV file: $csvFile" -Level "INFO"
        [System.Windows.Forms.MessageBox]::Show(
            "Data has been saved to:`n$csvFile", 
            "Success", 
            [System.Windows.Forms.MessageBoxButtons]::OK, 
            [System.Windows.Forms.MessageBoxIcon]::Information)
        
        # Clear input fields after successful save
        $controls.txtFirstName.Text = ""
        $controls.txtLastName.Text = ""
        $controls.chkExternal.IsChecked = $false
        $controls.txtExtCompany.Text = ""
        $controls.txtDescription.Text = ""
        $controls.cmbOffice.Text = ""
        $controls.txtPhone.Text = ""
        $controls.txtMobile.Text = ""
        $controls.txtMail.Text = ""
        $controls.txtPosition.Text = ""
        $controls.cmbBusinessUnit.Text = ""
        $controls.txtPersonalNumber.Text = ""
        if ($controls.dtpTermination) {
            $controls.dtpTermination.SelectedDate = [DateTime]::Now.AddYears(1)
        }
        if ($controls.dtpStartWorkDate) {
            $controls.dtpStartWorkDate.SelectedDate = [DateTime]::Now.AddDays(14)
        }
        $controls.chkTL.IsChecked = $false
        $controls.chkAL.IsChecked = $false
        # Reset admin account checkbox and notes
        $controls.chkProcessed.IsChecked = $false
        $controls.txtNotes.Text = ""
    } catch {
        Write-Log "Error saving data: $_" -Level "ERROR"
        [System.Windows.Forms.MessageBox]::Show(
            "Error saving data: $_", 
            "Error", 
            [System.Windows.Forms.MessageBoxButtons]::OK, 
            [System.Windows.Forms.MessageBoxIcon]::Error)
    }
})

# Close button click handler
$controls.btnClose.Add_Click({
    Write-DebugMessage "Button 'Close' was clicked. Closing application."
    $window.Close()
})

# Set focus to the form - conditional for WPF vs WinForms
if ($useXamlGUI -and $window.Dispatcher) {
    Write-DebugMessage "Setting focus using WPF Dispatcher"
    $window.Dispatcher.Invoke([Action]{
        $window.Activate()
        $controls.txtFirstName.Focus()
    }, [System.Windows.Threading.DispatcherPriority]::ApplicationIdle)
} else {
    Write-DebugMessage "Setting focus using Windows Forms"
    $window.Activate()
    $controls.txtFirstName.Focus()
}

# Show the form
Write-DebugMessage "Showing form..."
# Apply background color to window - conditional for WPF vs WinForms
if ($useXamlGUI -and $window.PSObject.Properties.Name -contains "Background") {
    Write-DebugMessage "Setting background for WPF window"
    $window.Background = [System.Windows.Media.Brushes]::$formBackColor
} elseif (-not $useXamlGUI) {
    Write-DebugMessage "Setting background for Windows Forms"
    $window.BackColor = [System.Drawing.Color]::$formBackColor
}

[void] $window.ShowDialog()

Write-DebugMessage "Script finished."

# SIG # Begin signature block
# MIIcCAYJKoZIhvcNAQcCoIIb+TCCG/UCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDcBYCFaHMJVlcK
# /OWpDsKiy0vCsOvAOFwAWQazTRhMX6CCFk4wggMQMIIB+KADAgECAhB3jzsyX9Cg
# jEi+sBC2rBMTMA0GCSqGSIb3DQEBCwUAMCAxHjAcBgNVBAMMFVBoaW5JVC1QU3Nj
# cmlwdHNfU2lnbjAeFw0yNTA3MDUwODI4MTZaFw0yNzA3MDUwODM4MTZaMCAxHjAc
# BgNVBAMMFVBoaW5JVC1QU3NjcmlwdHNfU2lnbjCCASIwDQYJKoZIhvcNAQEBBQAD
# ggEPADCCAQoCggEBALmz3o//iDA5MvAndTjGX7/AvzTSACClfuUR9WYK0f6Ut2dI
# mPxn+Y9pZlLjXIpZT0H2Lvxq5aSI+aYeFtuJ8/0lULYNCVT31Bf+HxervRBKsUyi
# W9+4PH6STxo3Pl4l56UNQMcWLPNjDORWRPWHn0f99iNtjI+L4tUC/LoWSs3obzxN
# 3uTypzlaPBxis2qFSTR5SWqFdZdRkcuI5LNsJjyc/QWdTYRrfmVqp0QrvcxzCv8u
# EiVuni6jkXfiE6wz+oeI3L2iR+ywmU6CUX4tPWoS9VTtmm7AhEpasRTmrrnSg20Q
# jiBa1eH5TyLAH3TcYMxhfMbN9a2xDX5pzM65EJUCAwEAAaNGMEQwDgYDVR0PAQH/
# BAQDAgeAMBMGA1UdJQQMMAoGCCsGAQUFBwMDMB0GA1UdDgQWBBQO7XOqiE/EYi+n
# IaR6YO5M2MUuVTANBgkqhkiG9w0BAQsFAAOCAQEAjYOKIwBu1pfbdvEFFaR/uY88
# peKPk0NnvNEc3dpGdOv+Fsgbz27JPvItITFd6AKMoN1W48YjQLaU22M2jdhjGN5i
# FSobznP5KgQCDkRsuoDKiIOTiKAAknjhoBaCCEZGw8SZgKJtWzbST36Thsdd/won
# ihLsuoLxfcFnmBfrXh3rTIvTwvfujob68s0Sf5derHP/F+nphTymlg+y4VTEAijk
# g2dhy8RAsbS2JYZT7K5aEJpPXMiOLBqd7oTGfM7y5sLk2LIM4cT8hzgz3v5yPMkF
# H2MdR//K403e1EKH9MsGuGAJZddVN8ppaiESoPLoXrgnw2SY5KCmhYw1xRFdjTCC
# BY0wggR1oAMCAQICEA6bGI750C3n79tQ4ghAGFowDQYJKoZIhvcNAQEMBQAwZTEL
# MAkGA1UEBhMCVVMxFTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3
# LmRpZ2ljZXJ0LmNvbTEkMCIGA1UEAxMbRGlnaUNlcnQgQXNzdXJlZCBJRCBSb290
# IENBMB4XDTIyMDgwMTAwMDAwMFoXDTMxMTEwOTIzNTk1OVowYjELMAkGA1UEBhMC
# VVMxFTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0
# LmNvbTEhMB8GA1UEAxMYRGlnaUNlcnQgVHJ1c3RlZCBSb290IEc0MIICIjANBgkq
# hkiG9w0BAQEFAAOCAg8AMIICCgKCAgEAv+aQc2jeu+RdSjwwIjBpM+zCpyUuySE9
# 8orYWcLhKac9WKt2ms2uexuEDcQwH/MbpDgW61bGl20dq7J58soR0uRf1gU8Ug9S
# H8aeFaV+vp+pVxZZVXKvaJNwwrK6dZlqczKU0RBEEC7fgvMHhOZ0O21x4i0MG+4g
# 1ckgHWMpLc7sXk7Ik/ghYZs06wXGXuxbGrzryc/NrDRAX7F6Zu53yEioZldXn1RY
# jgwrt0+nMNlW7sp7XeOtyU9e5TXnMcvak17cjo+A2raRmECQecN4x7axxLVqGDgD
# EI3Y1DekLgV9iPWCPhCRcKtVgkEy19sEcypukQF8IUzUvK4bA3VdeGbZOjFEmjNA
# vwjXWkmkwuapoGfdpCe8oU85tRFYF/ckXEaPZPfBaYh2mHY9WV1CdoeJl2l6SPDg
# ohIbZpp0yt5LHucOY67m1O+SkjqePdwA5EUlibaaRBkrfsCUtNJhbesz2cXfSwQA
# zH0clcOP9yGyshG3u3/y1YxwLEFgqrFjGESVGnZifvaAsPvoZKYz0YkH4b235kOk
# GLimdwHhD5QMIR2yVCkliWzlDlJRR3S+Jqy2QXXeeqxfjT/JvNNBERJb5RBQ6zHF
# ynIWIgnffEx1P2PsIV/EIFFrb7GrhotPwtZFX50g/KEexcCPorF+CiaZ9eRpL5gd
# LfXZqbId5RsCAwEAAaOCATowggE2MA8GA1UdEwEB/wQFMAMBAf8wHQYDVR0OBBYE
# FOzX44LScV1kTN8uZz/nupiuHA9PMB8GA1UdIwQYMBaAFEXroq/0ksuCMS1Ri6en
# IZ3zbcgPMA4GA1UdDwEB/wQEAwIBhjB5BggrBgEFBQcBAQRtMGswJAYIKwYBBQUH
# MAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBDBggrBgEFBQcwAoY3aHR0cDov
# L2NhY2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEUm9vdENBLmNy
# dDBFBgNVHR8EPjA8MDqgOKA2hjRodHRwOi8vY3JsMy5kaWdpY2VydC5jb20vRGln
# aUNlcnRBc3N1cmVkSURSb290Q0EuY3JsMBEGA1UdIAQKMAgwBgYEVR0gADANBgkq
# hkiG9w0BAQwFAAOCAQEAcKC/Q1xV5zhfoKN0Gz22Ftf3v1cHvZqsoYcs7IVeqRq7
# IviHGmlUIu2kiHdtvRoU9BNKei8ttzjv9P+Aufih9/Jy3iS8UgPITtAq3votVs/5
# 9PesMHqai7Je1M/RQ0SbQyHrlnKhSLSZy51PpwYDE3cnRNTnf+hZqPC/Lwum6fI0
# POz3A8eHqNJMQBk1RmppVLC4oVaO7KTVPeix3P0c2PR3WlxUjG/voVA9/HYJaISf
# b8rbII01YBwCA8sgsKxYoA5AY8WYIsGyWfVVa88nq2x2zm8jLfR+cWojayL/ErhU
# LSd+2DrZ8LaHlv1b0VysGMNNn3O3AamfV6peKOK5lDCCBrQwggScoAMCAQICEA3H
# rFcF/yGZLkBDIgw6SYYwDQYJKoZIhvcNAQELBQAwYjELMAkGA1UEBhMCVVMxFTAT
# BgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNvbTEh
# MB8GA1UEAxMYRGlnaUNlcnQgVHJ1c3RlZCBSb290IEc0MB4XDTI1MDUwNzAwMDAw
# MFoXDTM4MDExNDIzNTk1OVowaTELMAkGA1UEBhMCVVMxFzAVBgNVBAoTDkRpZ2lD
# ZXJ0LCBJbmMuMUEwPwYDVQQDEzhEaWdpQ2VydCBUcnVzdGVkIEc0IFRpbWVTdGFt
# cGluZyBSU0E0MDk2IFNIQTI1NiAyMDI1IENBMTCCAiIwDQYJKoZIhvcNAQEBBQAD
# ggIPADCCAgoCggIBALR4MdMKmEFyvjxGwBysddujRmh0tFEXnU2tjQ2UtZmWgyxU
# 7UNqEY81FzJsQqr5G7A6c+Gh/qm8Xi4aPCOo2N8S9SLrC6Kbltqn7SWCWgzbNfiR
# +2fkHUiljNOqnIVD/gG3SYDEAd4dg2dDGpeZGKe+42DFUF0mR/vtLa4+gKPsYfwE
# u7EEbkC9+0F2w4QJLVSTEG8yAR2CQWIM1iI5PHg62IVwxKSpO0XaF9DPfNBKS7Za
# zch8NF5vp7eaZ2CVNxpqumzTCNSOxm+SAWSuIr21Qomb+zzQWKhxKTVVgtmUPAW3
# 5xUUFREmDrMxSNlr/NsJyUXzdtFUUt4aS4CEeIY8y9IaaGBpPNXKFifinT7zL2gd
# FpBP9qh8SdLnEut/GcalNeJQ55IuwnKCgs+nrpuQNfVmUB5KlCX3ZA4x5HHKS+rq
# BvKWxdCyQEEGcbLe1b8Aw4wJkhU1JrPsFfxW1gaou30yZ46t4Y9F20HHfIY4/6vH
# espYMQmUiote8ladjS/nJ0+k6MvqzfpzPDOy5y6gqztiT96Fv/9bH7mQyogxG9QE
# PHrPV6/7umw052AkyiLA6tQbZl1KhBtTasySkuJDpsZGKdlsjg4u70EwgWbVRSX1
# Wd4+zoFpp4Ra+MlKM2baoD6x0VR4RjSpWM8o5a6D8bpfm4CLKczsG7ZrIGNTAgMB
# AAGjggFdMIIBWTASBgNVHRMBAf8ECDAGAQH/AgEAMB0GA1UdDgQWBBTvb1NK6eQG
# fHrK4pBW9i/USezLTjAfBgNVHSMEGDAWgBTs1+OC0nFdZEzfLmc/57qYrhwPTzAO
# BgNVHQ8BAf8EBAMCAYYwEwYDVR0lBAwwCgYIKwYBBQUHAwgwdwYIKwYBBQUHAQEE
# azBpMCQGCCsGAQUFBzABhhhodHRwOi8vb2NzcC5kaWdpY2VydC5jb20wQQYIKwYB
# BQUHMAKGNWh0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0
# ZWRSb290RzQuY3J0MEMGA1UdHwQ8MDowOKA2oDSGMmh0dHA6Ly9jcmwzLmRpZ2lj
# ZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRSb290RzQuY3JsMCAGA1UdIAQZMBcwCAYG
# Z4EMAQQCMAsGCWCGSAGG/WwHATANBgkqhkiG9w0BAQsFAAOCAgEAF877FoAc/gc9
# EXZxML2+C8i1NKZ/zdCHxYgaMH9Pw5tcBnPw6O6FTGNpoV2V4wzSUGvI9NAzaoQk
# 97frPBtIj+ZLzdp+yXdhOP4hCFATuNT+ReOPK0mCefSG+tXqGpYZ3essBS3q8nL2
# UwM+NMvEuBd/2vmdYxDCvwzJv2sRUoKEfJ+nN57mQfQXwcAEGCvRR2qKtntujB71
# WPYAgwPyWLKu6RnaID/B0ba2H3LUiwDRAXx1Neq9ydOal95CHfmTnM4I+ZI2rVQf
# jXQA1WSjjf4J2a7jLzWGNqNX+DF0SQzHU0pTi4dBwp9nEC8EAqoxW6q17r0z0noD
# js6+BFo+z7bKSBwZXTRNivYuve3L2oiKNqetRHdqfMTCW/NmKLJ9M+MtucVGyOxi
# Df06VXxyKkOirv6o02OoXN4bFzK0vlNMsvhlqgF2puE6FndlENSmE+9JGYxOGLS/
# D284NHNboDGcmWXfwXRy4kbu4QFhOm0xJuF2EZAOk5eCkhSxZON3rGlHqhpB/8Ml
# uDezooIs8CVnrpHMiD2wL40mm53+/j7tFaxYKIqL0Q4ssd8xHZnIn/7GELH3IdvG
# 2XlM9q7WP/UwgOkw/HQtyRN62JK4S1C8uw3PdBunvAZapsiI5YKdvlarEvf8EA+8
# hcpSM9LHJmyrxaFtoza2zNaQ9k+5t1wwggbtMIIE1aADAgECAhAKgO8YS43xBYLR
# xHanlXRoMA0GCSqGSIb3DQEBCwUAMGkxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5E
# aWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1l
# U3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYgMjAyNSBDQTEwHhcNMjUwNjA0MDAwMDAw
# WhcNMzYwOTAzMjM1OTU5WjBjMQswCQYDVQQGEwJVUzEXMBUGA1UEChMORGlnaUNl
# cnQsIEluYy4xOzA5BgNVBAMTMkRpZ2lDZXJ0IFNIQTI1NiBSU0E0MDk2IFRpbWVz
# dGFtcCBSZXNwb25kZXIgMjAyNSAxMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIIC
# CgKCAgEA0EasLRLGntDqrmBWsytXum9R/4ZwCgHfyjfMGUIwYzKomd8U1nH7C8Dr
# 0cVMF3BsfAFI54um8+dnxk36+jx0Tb+k+87H9WPxNyFPJIDZHhAqlUPt281mHrBb
# ZHqRK71Em3/hCGC5KyyneqiZ7syvFXJ9A72wzHpkBaMUNg7MOLxI6E9RaUueHTQK
# WXymOtRwJXcrcTTPPT2V1D/+cFllESviH8YjoPFvZSjKs3SKO1QNUdFd2adw44wD
# cKgH+JRJE5Qg0NP3yiSyi5MxgU6cehGHr7zou1znOM8odbkqoK+lJ25LCHBSai25
# CFyD23DZgPfDrJJJK77epTwMP6eKA0kWa3osAe8fcpK40uhktzUd/Yk0xUvhDU6l
# vJukx7jphx40DQt82yepyekl4i0r8OEps/FNO4ahfvAk12hE5FVs9HVVWcO5J4dV
# mVzix4A77p3awLbr89A90/nWGjXMGn7FQhmSlIUDy9Z2hSgctaepZTd0ILIUbWuh
# KuAeNIeWrzHKYueMJtItnj2Q+aTyLLKLM0MheP/9w6CtjuuVHJOVoIJ/DtpJRE7C
# e7vMRHoRon4CWIvuiNN1Lk9Y+xZ66lazs2kKFSTnnkrT3pXWETTJkhd76CIDBbTR
# ofOsNyEhzZtCGmnQigpFHti58CSmvEyJcAlDVcKacJ+A9/z7eacCAwEAAaOCAZUw
# ggGRMAwGA1UdEwEB/wQCMAAwHQYDVR0OBBYEFOQ7/PIx7f391/ORcWMZUEPPYYzo
# MB8GA1UdIwQYMBaAFO9vU0rp5AZ8esrikFb2L9RJ7MtOMA4GA1UdDwEB/wQEAwIH
# gDAWBgNVHSUBAf8EDDAKBggrBgEFBQcDCDCBlQYIKwYBBQUHAQEEgYgwgYUwJAYI
# KwYBBQUHMAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBdBggrBgEFBQcwAoZR
# aHR0cDovL2NhY2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0VHJ1c3RlZEc0VGlt
# ZVN0YW1waW5nUlNBNDA5NlNIQTI1NjIwMjVDQTEuY3J0MF8GA1UdHwRYMFYwVKBS
# oFCGTmh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRHNFRp
# bWVTdGFtcGluZ1JTQTQwOTZTSEEyNTYyMDI1Q0ExLmNybDAgBgNVHSAEGTAXMAgG
# BmeBDAEEAjALBglghkgBhv1sBwEwDQYJKoZIhvcNAQELBQADggIBAGUqrfEcJwS5
# rmBB7NEIRJ5jQHIh+OT2Ik/bNYulCrVvhREafBYF0RkP2AGr181o2YWPoSHz9iZE
# N/FPsLSTwVQWo2H62yGBvg7ouCODwrx6ULj6hYKqdT8wv2UV+Kbz/3ImZlJ7YXwB
# D9R0oU62PtgxOao872bOySCILdBghQ/ZLcdC8cbUUO75ZSpbh1oipOhcUT8lD8QA
# GB9lctZTTOJM3pHfKBAEcxQFoHlt2s9sXoxFizTeHihsQyfFg5fxUFEp7W42fNBV
# N4ueLaceRf9Cq9ec1v5iQMWTFQa0xNqItH3CPFTG7aEQJmmrJTV3Qhtfparz+BW6
# 0OiMEgV5GWoBy4RVPRwqxv7Mk0Sy4QHs7v9y69NBqycz0BZwhB9WOfOu/CIJnzkQ
# TwtSSpGGhLdjnQ4eBpjtP+XB3pQCtv4E5UCSDag6+iX8MmB10nfldPF9SVD7weCC
# 3yXZi/uuhqdwkgVxuiMFzGVFwYbQsiGnoa9F5AaAyBjFBtXVLcKtapnMG3VH3EmA
# p/jsJ3FVF3+d1SVDTmjFjLbNFZUWMXuZyvgLfgyPehwJVxwC+UpX2MSey2ueIu9T
# HFVkT+um1vshETaWyQo8gmBto/m3acaP9QsuLj3FNwFlTxq25+T4QwX9xa6ILs84
# ZPvmpovq90K8eWyG2N01c4IhSOxqt81nMYIFEDCCBQwCAQEwNDAgMR4wHAYDVQQD
# DBVQaGluSVQtUFNzY3JpcHRzX1NpZ24CEHePOzJf0KCMSL6wELasExMwDQYJYIZI
# AWUDBAIBBQCggYQwGAYKKwYBBAGCNwIBDDEKMAigAoAAoQKAADAZBgkqhkiG9w0B
# CQMxDAYKKwYBBAGCNwIBBDAcBgorBgEEAYI3AgELMQ4wDAYKKwYBBAGCNwIBFTAv
# BgkqhkiG9w0BCQQxIgQgf0NwyAGtfVpbp5DMS7qI91SYOe2+U12hm17PuIbg/R4w
# DQYJKoZIhvcNAQEBBQAEggEAtE+01fHYOLvvQUtVR6ytONnEAxZRHMi2GsuIBxYh
# 6aVp6jlPSpN5NG0ueu7wlKVCuDoacD5e+KXDb4mjTNHNtO9bNpVXq92x6Y44UYGD
# fWEEDSSu4+vNKuLK9hPvah3/PfEMe8zohTX4UOSIUi2Ro6aYcTzxsAFp3M6D/wrd
# doZ4YtD2inGBU5l27qMOzDvlzENlywjh7Pjd2LINu8VRtHqaVOsYJ4vEnxVfiBsL
# eMZ3NBVMT4cQ7uV/56HaBfnW3IKALqLBbg+Iv3wdwioIbCrLy6WnKu4PzixJ+HAb
# AbE6FAWYqoAiMTxMKyu7wl8nL2rK1fPMhxNVwj0RnfaPX6GCAyYwggMiBgkqhkiG
# 9w0BCQYxggMTMIIDDwIBATB9MGkxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdp
# Q2VydCwgSW5jLjFBMD8GA1UEAxM4RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3Rh
# bXBpbmcgUlNBNDA5NiBTSEEyNTYgMjAyNSBDQTECEAqA7xhLjfEFgtHEdqeVdGgw
# DQYJYIZIAWUDBAIBBQCgaTAYBgkqhkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqG
# SIb3DQEJBTEPFw0yNTA3MTIwODAwMjBaMC8GCSqGSIb3DQEJBDEiBCCB59K5RWch
# jnpKTKIbu8NBs3sTLsZ3d9WXskc01aHe8zANBgkqhkiG9w0BAQEFAASCAgCF1vxZ
# Ig6SwLn6P+NBAKvVeQM/1T2uULScuRdtUHZ38DqrjgNSI4ZeVLl/h78Q8rGn+BKw
# QRdo6S5xcH7NOb4EmBnIX2ofds3cwmmySC6cYsBeChpG8FwGMXG2zlD4tjFuJicS
# MplgxI/vciO4SSiEnduXKNivggc2WT/+rGkwK7WbYVv7d4bQWpe82f0k10C4WuF3
# cePvLq6gs16/mW7g0/cxP0lInI6M246LjZTwCRj2/xrIlhCBc7wiklQRL+vxr/YM
# owy29fOsmui3iijmyFDwGiRjYpkfTx9R4hw4zekoQaquFSfCm+Ug7KwybLuRJLJC
# DWL8aNM07cBLafpqXy5eWfi3/2m7zL2HRwCfDskeb8EHu7hhoApQx/RrMByfqIEY
# w7udwuoNWKvv7p0qIOrzBOW99P2IzCr9pMZAeCHU/EJKuXqJ8lEf5LCC4ZTbOJnx
# KgzRr26aG5xX4lQB46f7jL5+ewct0gYXHzhNDxJcEWN9GbUhquraK6MsBOd8D3rE
# ZrZt/iZrmBjTHtpX9NQzPCLOyjqMWDE8oEyLX6Y+ZQYrQuMNQyobN4nFeMxRN/o0
# /JwpVqJoXny9lpfQdsSY2ePh3y1IvB8TGYKb/dOex6pC0ejRdohkXuWeq6SFVjEL
# qOsuglXvoUefj0qRbyY3Uyi0JEHreJ6bJSZ9Gw==
# SIG # End signature block
