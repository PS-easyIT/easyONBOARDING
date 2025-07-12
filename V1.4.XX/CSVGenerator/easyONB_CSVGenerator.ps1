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
# MIIbywYJKoZIhvcNAQcCoIIbvDCCG7gCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDcBYCFaHMJVlcK
# /OWpDsKiy0vCsOvAOFwAWQazTRhMX6CCFhcwggMQMIIB+KADAgECAhB3jzsyX9Cg
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
# LSd+2DrZ8LaHlv1b0VysGMNNn3O3AamfV6peKOK5lDCCBq4wggSWoAMCAQICEAc2
# N7ckVHzYR6z9KGYqXlswDQYJKoZIhvcNAQELBQAwYjELMAkGA1UEBhMCVVMxFTAT
# BgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNvbTEh
# MB8GA1UEAxMYRGlnaUNlcnQgVHJ1c3RlZCBSb290IEc0MB4XDTIyMDMyMzAwMDAw
# MFoXDTM3MDMyMjIzNTk1OVowYzELMAkGA1UEBhMCVVMxFzAVBgNVBAoTDkRpZ2lD
# ZXJ0LCBJbmMuMTswOQYDVQQDEzJEaWdpQ2VydCBUcnVzdGVkIEc0IFJTQTQwOTYg
# U0hBMjU2IFRpbWVTdGFtcGluZyBDQTCCAiIwDQYJKoZIhvcNAQEBBQADggIPADCC
# AgoCggIBAMaGNQZJs8E9cklRVcclA8TykTepl1Gh1tKD0Z5Mom2gsMyD+Vr2EaFE
# FUJfpIjzaPp985yJC3+dH54PMx9QEwsmc5Zt+FeoAn39Q7SE2hHxc7Gz7iuAhIoi
# GN/r2j3EF3+rGSs+QtxnjupRPfDWVtTnKC3r07G1decfBmWNlCnT2exp39mQh0YA
# e9tEQYncfGpXevA3eZ9drMvohGS0UvJ2R/dhgxndX7RUCyFobjchu0CsX7LeSn3O
# 9TkSZ+8OpWNs5KbFHc02DVzV5huowWR0QKfAcsW6Th+xtVhNef7Xj3OTrCw54qVI
# 1vCwMROpVymWJy71h6aPTnYVVSZwmCZ/oBpHIEPjQ2OAe3VuJyWQmDo4EbP29p7m
# O1vsgd4iFNmCKseSv6De4z6ic/rnH1pslPJSlRErWHRAKKtzQ87fSqEcazjFKfPK
# qpZzQmiftkaznTqj1QPgv/CiPMpC3BhIfxQ0z9JMq++bPf4OuGQq+nUoJEHtQr8F
# nGZJUlD0UfM2SU2LINIsVzV5K6jzRWC8I41Y99xh3pP+OcD5sjClTNfpmEpYPtMD
# iP6zj9NeS3YSUZPJjAw7W4oiqMEmCPkUEBIDfV8ju2TjY+Cm4T72wnSyPx4Jduyr
# XUZ14mCjWAkBKAAOhFTuzuldyF4wEr1GnrXTdrnSDmuZDNIztM2xAgMBAAGjggFd
# MIIBWTASBgNVHRMBAf8ECDAGAQH/AgEAMB0GA1UdDgQWBBS6FtltTYUvcyl2mi91
# jGogj57IbzAfBgNVHSMEGDAWgBTs1+OC0nFdZEzfLmc/57qYrhwPTzAOBgNVHQ8B
# Af8EBAMCAYYwEwYDVR0lBAwwCgYIKwYBBQUHAwgwdwYIKwYBBQUHAQEEazBpMCQG
# CCsGAQUFBzABhhhodHRwOi8vb2NzcC5kaWdpY2VydC5jb20wQQYIKwYBBQUHMAKG
# NWh0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRSb290
# RzQuY3J0MEMGA1UdHwQ8MDowOKA2oDSGMmh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNv
# bS9EaWdpQ2VydFRydXN0ZWRSb290RzQuY3JsMCAGA1UdIAQZMBcwCAYGZ4EMAQQC
# MAsGCWCGSAGG/WwHATANBgkqhkiG9w0BAQsFAAOCAgEAfVmOwJO2b5ipRCIBfmbW
# 2CFC4bAYLhBNE88wU86/GPvHUF3iSyn7cIoNqilp/GnBzx0H6T5gyNgL5Vxb122H
# +oQgJTQxZ822EpZvxFBMYh0MCIKoFr2pVs8Vc40BIiXOlWk/R3f7cnQU1/+rT4os
# equFzUNf7WC2qk+RZp4snuCKrOX9jLxkJodskr2dfNBwCnzvqLx1T7pa96kQsl3p
# /yhUifDVinF2ZdrM8HKjI/rAJ4JErpknG6skHibBt94q6/aesXmZgaNWhqsKRcnf
# xI2g55j7+6adcq/Ex8HBanHZxhOACcS2n82HhyS7T6NJuXdmkfFynOlLAlKnN36T
# U6w7HQhJD5TNOXrd/yVjmScsPT9rp/Fmw0HNT7ZAmyEhQNC3EyTN3B14OuSereU0
# cZLXJmvkOHOrpgFPvT87eK1MrfvElXvtCl8zOYdBeHo46Zzh3SP9HSjTx/no8Zhf
# +yvYfvJGnXUsHicsJttvFXseGYs2uJPU5vIXmVnKcPA3v5gA3yAWTyf7YGcWoWa6
# 3VXAOimGsJigK+2VQbc61RWYMbRiCQ8KvYHZE/6/pNHzV9m8BPqC3jLfBInwAM1d
# wvnQI38AC+R2AibZ8GV2QqYphwlHK+Z/GqSFD/yYlvZVVCsfgPrA8g4r5db7qS9E
# FUrnEw4d2zc4GqEr9u3WfPwwgga8MIIEpKADAgECAhALrma8Wrp/lYfG+ekE4zME
# MA0GCSqGSIb3DQEBCwUAMGMxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2Vy
# dCwgSW5jLjE7MDkGA1UEAxMyRGlnaUNlcnQgVHJ1c3RlZCBHNCBSU0E0MDk2IFNI
# QTI1NiBUaW1lU3RhbXBpbmcgQ0EwHhcNMjQwOTI2MDAwMDAwWhcNMzUxMTI1MjM1
# OTU5WjBCMQswCQYDVQQGEwJVUzERMA8GA1UEChMIRGlnaUNlcnQxIDAeBgNVBAMT
# F0RpZ2lDZXJ0IFRpbWVzdGFtcCAyMDI0MIICIjANBgkqhkiG9w0BAQEFAAOCAg8A
# MIICCgKCAgEAvmpzn/aVIauWMLpbbeZZo7Xo/ZEfGMSIO2qZ46XB/QowIEMSvgjE
# dEZ3v4vrrTHleW1JWGErrjOL0J4L0HqVR1czSzvUQ5xF7z4IQmn7dHY7yijvoQ7u
# jm0u6yXF2v1CrzZopykD07/9fpAT4BxpT9vJoJqAsP8YuhRvflJ9YeHjes4fduks
# THulntq9WelRWY++TFPxzZrbILRYynyEy7rS1lHQKFpXvo2GePfsMRhNf1F41nyE
# g5h7iOXv+vjX0K8RhUisfqw3TTLHj1uhS66YX2LZPxS4oaf33rp9HlfqSBePejlY
# eEdU740GKQM7SaVSH3TbBL8R6HwX9QVpGnXPlKdE4fBIn5BBFnV+KwPxRNUNK6lY
# k2y1WSKour4hJN0SMkoaNV8hyyADiX1xuTxKaXN12HgR+8WulU2d6zhzXomJ2Ple
# I9V2yfmfXSPGYanGgxzqI+ShoOGLomMd3mJt92nm7Mheng/TBeSA2z4I78JpwGpT
# RHiT7yHqBiV2ngUIyCtd0pZ8zg3S7bk4QC4RrcnKJ3FbjyPAGogmoiZ33c1HG93V
# p6lJ415ERcC7bFQMRbxqrMVANiav1k425zYyFMyLNyE1QulQSgDpW9rtvVcIH7Wv
# G9sqYup9j8z9J1XqbBZPJ5XLln8mS8wWmdDLnBHXgYly/p1DhoQo5fkCAwEAAaOC
# AYswggGHMA4GA1UdDwEB/wQEAwIHgDAMBgNVHRMBAf8EAjAAMBYGA1UdJQEB/wQM
# MAoGCCsGAQUFBwMIMCAGA1UdIAQZMBcwCAYGZ4EMAQQCMAsGCWCGSAGG/WwHATAf
# BgNVHSMEGDAWgBS6FtltTYUvcyl2mi91jGogj57IbzAdBgNVHQ4EFgQUn1csA3cO
# KBWQZqVjXu5Pkh92oFswWgYDVR0fBFMwUTBPoE2gS4ZJaHR0cDovL2NybDMuZGln
# aWNlcnQuY29tL0RpZ2lDZXJ0VHJ1c3RlZEc0UlNBNDA5NlNIQTI1NlRpbWVTdGFt
# cGluZ0NBLmNybDCBkAYIKwYBBQUHAQEEgYMwgYAwJAYIKwYBBQUHMAGGGGh0dHA6
# Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBYBggrBgEFBQcwAoZMaHR0cDovL2NhY2VydHMu
# ZGlnaWNlcnQuY29tL0RpZ2lDZXJ0VHJ1c3RlZEc0UlNBNDA5NlNIQTI1NlRpbWVT
# dGFtcGluZ0NBLmNydDANBgkqhkiG9w0BAQsFAAOCAgEAPa0eH3aZW+M4hBJH2UOR
# 9hHbm04IHdEoT8/T3HuBSyZeq3jSi5GXeWP7xCKhVireKCnCs+8GZl2uVYFvQe+p
# PTScVJeCZSsMo1JCoZN2mMew/L4tpqVNbSpWO9QGFwfMEy60HofN6V51sMLMXNTL
# fhVqs+e8haupWiArSozyAmGH/6oMQAh078qRh6wvJNU6gnh5OruCP1QUAvVSu4kq
# VOcJVozZR5RRb/zPd++PGE3qF1P3xWvYViUJLsxtvge/mzA75oBfFZSbdakHJe2B
# VDGIGVNVjOp8sNt70+kEoMF+T6tptMUNlehSR7vM+C13v9+9ZOUKzfRUAYSyyEmY
# tsnpltD/GWX8eM70ls1V6QG/ZOB6b6Yum1HvIiulqJ1Elesj5TMHq8CWT/xrW7tw
# ipXTJ5/i5pkU5E16RSBAdOp12aw8IQhhA/vEbFkEiF2abhuFixUDobZaA0VhqAsM
# HOmaT3XThZDNi5U2zHKhUs5uHHdG6BoQau75KiNbh0c+hatSF+02kULkftARjsyE
# pHKsF7u5zKRbt5oK5YGwFvgc4pEVUNytmB3BpIiowOIIuDgP5M9WArHYSAR16gc0
# dP2XdkMEP5eBsX7bf/MGN4K3HP50v/01ZHo/Z5lGLvNwQ7XHBx1yomzLP8lx4Q1z
# ZKDyHcp4VQJLu2kWTsKsOqQxggUKMIIFBgIBATA0MCAxHjAcBgNVBAMMFVBoaW5J
# VC1QU3NjcmlwdHNfU2lnbgIQd487Ml/QoIxIvrAQtqwTEzANBglghkgBZQMEAgEF
# AKCBhDAYBgorBgEEAYI3AgEMMQowCKACgAChAoAAMBkGCSqGSIb3DQEJAzEMBgor
# BgEEAYI3AgEEMBwGCisGAQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3
# DQEJBDEiBCB/Q3DIAa19WlunkMxLuoj3VJg57b5TXaGbXs+4huD9HjANBgkqhkiG
# 9w0BAQEFAASCAQC0T7TV8dg4u+9BS1VHrK042cQDFlEcyLYay4gHFiHppWnqOU9K
# k3k0bS567vCUpUK4OhpwPl74pcNviaNM0c2071s2lVer3bHpjjhRgYN9YQQNJK7j
# 680q4sr2E+9qHf898Qx7zOiFNfhQ5IhSLZGjpphxPPGwAWnczoP/Ct12hnhi0PaK
# cYFTmXbuow7MO+XMQ2XLCOHs+N3Ysg27xVG0eppU6xgni8SfFV+IGwt4xnc0FUxP
# hxDu5X/nodoF+dbcgoAuosFuD4i/fB3CKghsKsvLpacq7g/OLEn4cBsBsToUBZiq
# gCIxPEwrK7vCXycvasrV88yHE1XCPRGd9o9foYIDIDCCAxwGCSqGSIb3DQEJBjGC
# Aw0wggMJAgEBMHcwYzELMAkGA1UEBhMCVVMxFzAVBgNVBAoTDkRpZ2lDZXJ0LCBJ
# bmMuMTswOQYDVQQDEzJEaWdpQ2VydCBUcnVzdGVkIEc0IFJTQTQwOTYgU0hBMjU2
# IFRpbWVTdGFtcGluZyBDQQIQC65mvFq6f5WHxvnpBOMzBDANBglghkgBZQMEAgEF
# AKBpMBgGCSqGSIb3DQEJAzELBgkqhkiG9w0BBwEwHAYJKoZIhvcNAQkFMQ8XDTI1
# MDcwODE1MDk0OFowLwYJKoZIhvcNAQkEMSIEIIHn0rlFZyGOekpMohu7w0GzexMu
# xnd31ZeyRzTVod7zMA0GCSqGSIb3DQEBAQUABIICAC3uJixl4JsX2CHDnQuwp9C1
# 2LmZvZs6ERrpY1/9bbwZkIOmCoqqvIWog9G85SxT+DrK6ZRk6lcZXcsMykidvpyU
# OeVP1EVTJ+KlQYNTSyKii7u6dMFZ7/W6XiYx3ufhBdVGjOEhNcGROB57VBKOejsB
# ANlsLdxsxvxS/2PnqNH42mT01bvtMkYN3h8RbuMHQYWvis1NTeWKA6cg2HnydUXI
# T/rM0tjYFHEYAhH8M3EXYHeY1wICkrlP7SCzlg/UtQbaZzXFanmayRRcwXFQidk/
# mpq0btyqpwsrHUwqzfI2la6OjsNUXi5jPfRMo4ByDpMxYhot+9NUg92KwVCQ68+W
# iCRcDnaqHmCBIwfsRJwiqU/BuEcXrMVYGr0upO1PaJNZdnEFkzG1OaRoxfGxvi4K
# GDosR1whF7hcFoG//OBJfj6vKOSX7gTbznwXt9OWEUSeA8eV/xS0vT/wkSg/L7EE
# D3q8ot3F+HxtmNbGSWmNptjTn2KHcSjw74g97uQHx/xEo6rH+wQwhdgfSZIx2Afo
# k/Rf3VWc7UpsV7lCK0wxZ9yuCe68jnzL3m7iB0Z0C6tuF4Mn1QLtC+0ebYn3PFby
# D3rJNijzEAD5Y8YXVpgq64FbzbVgJFQ/ZonwlFA4GNJFmLeGCFdBqGeNkYWlq8J0
# MJXHFKZfJlLDYSmzCSv0
# SIG # End signature block
