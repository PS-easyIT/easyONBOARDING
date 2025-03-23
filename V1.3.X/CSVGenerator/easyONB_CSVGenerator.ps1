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
