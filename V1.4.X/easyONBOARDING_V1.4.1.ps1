#region [Region 00 | SCRIPT INITIALIZATION]
# easyONBOARDING_V1.3.8.ps1
# Script for automated user onboarding in Active Directory environments
# Version: 1.3.8

# Get script directory for relative path resolution
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
$ModulesDir = Join-Path -Path $ScriptDir -ChildPath 'Modules'

# Define standard paths needed by modules (globally accessible)
$global:BasePath = $ScriptDir
$global:INIPath = Join-Path -Path $ScriptDir -ChildPath 'easyONB.ini'
$global:LogsPath = Join-Path -Path $ScriptDir -ChildPath 'Logs'
$global:ReportsPath = Join-Path -Path $ScriptDir -ChildPath 'Reports'
$global:TemplatesPath = Join-Path -Path $ScriptDir -ChildPath 'Templates'
$global:DataPath = Join-Path -Path $ScriptDir -ChildPath 'Data'

# Create directories if they don't exist
$pathsToCreate = @($LogsPath, $ReportsPath, (Split-Path -Parent $INIPath))
foreach ($path in $pathsToCreate) {
    if (-not (Test-Path -Path $path)) {
        try {
            New-Item -Path $path -ItemType Directory -Force | Out-Null
            Write-Host "Created directory: $path" -ForegroundColor Green
        }
        catch {
            Write-Host "Error creating directory $path : $($_.Exception.Message)" -ForegroundColor Red
        }
    }
}

# Initial debug handling - will be updated after INI file is loaded
$script:DebugEnabled = $false
$script:DebugMode = 0

# Simple debug function to use before modules are loaded
function Write-InitialDebug {
    param (
        [string]$Message
    )
    
    if ($script:DebugEnabled) {
        Write-Host "[DEBUG] $Message" -ForegroundColor Cyan
    }
}

Write-Host "easyONBOARDING v1.3.8 Initializing..." -ForegroundColor Cyan
Write-InitialDebug "Script directory: $ScriptDir"
Write-InitialDebug "Modules directory: $ModulesDir"
Write-InitialDebug "INI file path: $global:INIPath"
Write-InitialDebug "Logs directory: $global:LogsPath"
Write-InitialDebug "Reports directory: $global:ReportsPath"

# Function to import modules with error handling
function Import-CustomModule {
    param (
        [Parameter(Mandatory = $true)]
        [string]$ModuleName,
        
        [Parameter(Mandatory = $false)]
        [switch]$Critical
    )
    
    $modulePath = Join-Path -Path $ModulesDir -ChildPath "$ModuleName.psm1"
    
    try {
        Write-InitialDebug "Attempting to import module: $ModuleName from $modulePath"
        
        if (Test-Path $modulePath) {
            Import-Module $modulePath -Force -ErrorAction Stop
            Write-Host "Module loaded: $ModuleName" -ForegroundColor Green
            
            # Special handling for LoggingDebuging module - set up debug mode
            if ($ModuleName -eq "LoggingDebuging") {
                Write-InitialDebug "Logging and debugging module loaded, configuring debug handlers"
            }
            # Special handling for LoadConfigINI - read debug settings
            elseif ($ModuleName -eq "LoadConfigINI" -and (Test-Path variable:global:Config)) {
                # Update debug settings based on INI file
                if ($global:Config.Contains("General") -and $global:Config.General.Contains("Debug")) {
                    $script:DebugMode = [int]($global:Config.General.Debug)
                    $script:DebugEnabled = $script:DebugMode -gt 0
                    
                    # Now use the proper debug function if available
                    if (Get-Command -Name "Write-DebugMessage" -ErrorAction SilentlyContinue) {
                        Write-DebugMessage "Debug mode enabled (Level: $script:DebugMode)" -LogLevel "INFO"
                        Write-DebugMessage "INI Configuration loaded - Debug mode set to $script:DebugMode" -LogLevel "DEBUG"
                    } else {
                        Write-InitialDebug "Debug mode enabled (Level: $script:DebugMode)"
                        Write-InitialDebug "INI Configuration loaded - Debug mode set to $script:DebugMode"
                    }
                }
            }
            
            return $true
        } else {
            throw "Module file not found: $modulePath"
        }
    } catch {
        $errorMsg = "Error loading module '$ModuleName': $($_.Exception.Message)"
        Write-Host $errorMsg -ForegroundColor Red
        
        if (Get-Command -Name "Write-LogMessage" -ErrorAction SilentlyContinue) {
            Write-LogMessage -Message $errorMsg -LogLevel "ERROR"
        }
        
        if ($Critical) {
            $criticalMsg = "Critical module '$ModuleName' could not be loaded. The script cannot continue."
            Write-Host $criticalMsg -ForegroundColor Red -BackgroundColor Black
            
            if (Get-Command -Name "Write-LogMessage" -ErrorAction SilentlyContinue) {
                Write-LogMessage -Message $criticalMsg -LogLevel "FATAL"
            }
            
            throw $criticalMsg
        }
        return $false
    }
}

# Import modules in the correct order (dependencies first)
Write-Host "Loading easyONBOARDING modules..." -ForegroundColor Cyan
Write-InitialDebug "Module loading sequence started"

# 1. Core system modules (critical)
Write-InitialDebug "Loading core system modules (critical dependencies)"
Import-CustomModule -ModuleName "CheckAdminRights" -Critical
Import-CustomModule -ModuleName "FunctionLogDebug" -Critical

# After LoggingDebuging is loaded, we can use the proper debug functions
# 2. Configuration and assemblies (critical)
Write-DebugMessage "Loading configuration and assembly modules (critical dependencies)" -LogLevel "DEBUG"
Import-CustomModule -ModuleName "CheckModulesAssemblies" -Critical
Import-CustomModule -ModuleName "CheckModulesAD" -Critical
Import-CustomModule -ModuleName "LoadConfigINI" -Critical

# 3. Basic helpers and functions
Write-DebugMessage "Loading helper and function modules" -LogLevel "DEBUG"
Import-CustomModule -ModuleName "FunctionADCommand"
Import-CustomModule -ModuleName "GUIEventHandler"

# 4. AD functionality
Write-DebugMessage "Loading Active Directory functionality modules" -LogLevel "DEBUG"
Import-CustomModule -ModuleName "FunctionPassword"
Import-CustomModule -ModuleName "FunctionUPNCreate"
Import-CustomModule -ModuleName "FunctionUPNTemplate"
Import-CustomModule -ModuleName "FunctionADUserCreate"

# 5. GUI and display functionality
Write-DebugMessage "Loading GUI and display modules" -LogLevel "DEBUG"
Import-CustomModule -ModuleName "LoadConfigXAML" -Critical
Import-CustomModule -ModuleName "FunctionsSetLogoSetPassword"
Import-CustomModule -ModuleName "FunctionDropDowns"
Import-CustomModule -ModuleName "GUIDropDowns"
Import-CustomModule -ModuleName "GUIADGroup"
Import-CustomModule -ModuleName "GUIUPNTemplateInfo"

# 6. AD data and management
Write-DebugMessage "Loading AD data and setting modules" -LogLevel "DEBUG"
Import-CustomModule -ModuleName "FunctionADDataLoad"
Import-CustomModule -ModuleName "FunctionSettingsImportINI"
Import-CustomModule -ModuleName "FunctionSettingsLoadINI"
Import-CustomModule -ModuleName "FunctionSettingsSaveINI"

# 7. Core onboarding functionality (critical)
Write-DebugMessage "Loading core onboarding functionality (critical)" -LogLevel "DEBUG"
Import-CustomModule -ModuleName "CoreEASYONBOARDING" -Critical

Write-Host "All modules loaded successfully!" -ForegroundColor Green
Write-DebugMessage "Module loading sequence completed successfully" -LogLevel "INFO"

# Helper function to check debug status (can be called anywhere in script)
function Test-DebugEnabled {
    if ($null -ne (Get-Variable -Name 'Config' -Scope 'Global' -ErrorAction SilentlyContinue)) {
        if ($global:Config.Contains("General") -and $global:Config.General.Contains("Debug")) {
            return [int]($global:Config.General.Debug) -gt 0
        }
    }
    # Fallback to script variable if global:Config is not available
    return $script:DebugEnabled
}

# Enhance Write-DebugMessage function if it exists
if (Get-Command -Name "Write-DebugMessage" -ErrorAction SilentlyContinue) {
    # Create a wrapper that respects debug flag in INI
    $originalWriteDebugMessage = Get-Item -Path function:Write-DebugMessage
    
    # Redefine the function to respect debug settings
    function global:Write-DebugMessage {
        param(
            [string]$Message,
            [string]$LogLevel = "DEBUG"
        )
        
        if (Test-DebugEnabled) {
            & $originalWriteDebugMessage -Message $Message -LogLevel $LogLevel
        }
    }
    
    Write-DebugMessage "Enhanced debug messaging enabled and configured" -LogLevel "DEBUG"
}
#endregion

# Allow time for UI to render
$Panel.Dispatcher.Invoke([System.Action]{
    $Panel.UpdateLayout()
}, "Render")
    
try {
    # Get all checkboxes in the panel
    $checkBoxes = $Panel.Dispatcher.Invoke([System.Func[array]]{
        Find-CheckBoxes -Parent $Panel
    })
    
    Write-DebugMessage "Found $($checkBoxes.Count) checkboxes in panel"
    
    # Process each checkbox
    for ($i = 0; $i -lt $checkBoxes.Count; $i++) {
        $cb = $checkBoxes[$i]
        
        if ($cb.IsChecked) {
            $groupName = $cb.Content.ToString()
            Write-DebugMessage "Selected group: $groupName"
            [void]$selectedGroups.Add($groupName)
        }
    }
} catch {
    Write-DebugMessage "Error processing AD group checkboxes: $($_.Exception.Message)"
    # Continue anyway with any groups we've found
}

Write-DebugMessage "Total selected groups: $($selectedGroups.Count)"

# Return appropriate result
if ($selectedGroups.Count -eq 0) {
    return "NONE"
} else {
    return ($selectedGroups -join $Separator)
}

# Helper function to find a specific type of child in the visual tree
function FindVisualChild {
    param(
        [Parameter(Mandatory=$true)]
        [System.Windows.DependencyObject]$parent,
        
        [Parameter(Mandatory=$true)]
        [type]$childType
    )
    
    for ($i = 0; $i -lt [System.Windows.Media.VisualTreeHelper]::GetChildrenCount($parent); $i++) {
        $child = [System.Windows.Media.VisualTreeHelper]::GetChild($parent, $i)
        
        if ($child -and $child -is $childType) {
            return $child
        } else {
            $result = FindVisualChild -parent $child -childType $childType
            if ($result) {
                return $result
            }
        }
    }
    
    return $null
}
#endregion

Write-DebugMessage "Setting up logo upload button."

#region [Region 21.2 | LOGO UPLOAD BUTTON]
# Sets up the logo upload button and its functionality
$btnUploadLogo = $window.FindName("btnUploadLogo")
if ($btnUploadLogo) {
    # Make sure this script block is properly formatted
    $btnUploadLogo.Add_Click({
        if (-not $global:Config -or -not $global:Config.Contains("Report")) {
            [System.Windows.MessageBox]::Show("The section [Report] is missing in the INI file.", "Error", "OK", "Error")
            return
        }
        $brandingConfig = $global:Config["Report"]
        # Call the function directly, not with &
        Set-Logo -brandingConfig $brandingConfig
    })
}

#GUI: START "easyONBOARDING BUTTON"
Write-DebugMessage "GUI: TAB easyONBOARDING loaded"

#region [Region 21.3 | ONBOARDING TAB]
# Implements the main onboarding tab functionality
Write-DebugMessage "Onboarding tab loaded."
$onboardingTab = $window.FindName("Tab_Onboarding")
# [21.3.1 - Validates the tab exists in the XAML interface]
if (-not $onboardingTab) {
    Write-DebugMessage "ERROR: Onboarding-Tab missing!"
    Write-Error "The Onboarding-Tab (x:Name='Tab_Onboarding') was not found in the XAML."
    exit
}
#endregion

Write-DebugMessage "Setting up onboarding button handler."

#region [Region 21.4 | ONBOARDING BUTTON HANDLER]
# Implements the start onboarding button click handler
$btnStartOnboarding = $onboardingTab.FindName("btnOnboard")
# [21.4.1 - Sets up the primary function to execute when user initiates onboarding]
if (-not $btnStartOnboarding) {
    Write-DebugMessage "ERROR: btnOnboard was NOT found!"
    exit
}

Write-DebugMessage "Registering event handler for btnOnboard"
$btnStartOnboarding.Add_Click({
    Write-DebugMessage "Onboarding button clicked!"

    try {
        # Find the AD groups panel for selection
        $icADGroups = $onboardingTab.FindName("icADGroups")
        $ADGroupsSelected = if ($icADGroups) { 
            # Get selected groups as a string
            Get-SelectedADGroups -Panel $icADGroups 
        } else { 
            "NONE" 
        }
        
        Write-DebugMessage "Loading GUI elements..."
        Write-DebugMessage "Selected AD groups: $ADGroupsSelected"
        $txtFirstName = $onboardingTab.FindName("txtFirstName")
        $txtLastName = $onboardingTab.FindName("txtLastName")
        $txtDisplayName = $onboardingTab.FindName("txtDisplayName")
        $txtEmail = $onboardingTab.FindName("txtEmail")
        $txtOffice = $onboardingTab.FindName("txtOffice")
        $txtPhone = $onboardingTab.FindName("txtPhone")
        $txtMobile = $onboardingTab.FindName("txtMobile")
        $txtPosition = $onboardingTab.FindName("txtPosition")
        $txtDepartment = $onboardingTab.FindName("txtDepartment")
        $txtTermination = $onboardingTab.FindName("txtTermination")
        $chkExternal = $onboardingTab.FindName("chkExternal")
        $chkTL = $onboardingTab.FindName("chkTL")
        $chkAL = $onboardingTab.FindName("chkAL")
        $txtDescription = $onboardingTab.FindName("txtDescription")
        $chkAccountDisabled = $onboardingTab.FindName("chkAccountDisabled")
        $chkMustChangePassword = $onboardingTab.FindName("chkPWChangeLogon")
        $chkPasswordNeverExpires = $onboardingTab.FindName("chkPWNeverExpires")
        $comboBoxDisplayTemplate = $onboardingTab.FindName("cmbDisplayTemplate")
        $comboBoxMailSuffix = $onboardingTab.FindName("cmbMailSuffix")
        $comboBoxLicense = $onboardingTab.FindName("cmbLicense")
        $comboBoxOU = $onboardingTab.FindName("cmbOU")
        $comboBoxTLGroup = $onboardingTab.FindName("cmbTLGroup")

        Write-DebugMessage "GUI elements successfully loaded."

        # **Validate mandatory fields**
        if (-not $txtFirstName -or -not $txtFirstName.Text) {
            Write-DebugMessage "ERROR: First name missing!"
            [System.Windows.MessageBox]::Show("Error: First name cannot be empty!", "Error", 'OK', 'Error')
            return
        }
        if (-not $txtLastName -or -not $txtLastName.Text) {
            Write-DebugMessage "ERROR: Last name missing!"
            [System.Windows.MessageBox]::Show("Error: Last name cannot be empty!", "Error", 'OK', 'Error')
            return
        }

        # **Set values, avoid NULL values**
        $firstName = $txtFirstName.Text.Trim()
        $lastName = $txtLastName.Text.Trim()
        $displayName = if ($txtDisplayName -and $txtDisplayName.Text) { $txtDisplayName.Text.Trim() } else { "$firstName $lastName" }
        $email = if ($txtEmail -and $txtEmail.Text) { $txtEmail.Text.Trim() } else { "" }
        $office = if ($txtOffice -and $txtOffice.Text) { $txtOffice.Text.Trim() } else { "" }
        $phone = if ($txtPhone -and $txtPhone.Text) { $txtPhone.Text.Trim() } else { "" }
        $mobile = if ($txtMobile -and $txtMobile.Text) { $txtMobile.Text.Trim() } else { "" }
        $position = if ($txtPosition -and $txtPosition.Text) { $txtPosition.Text.Trim() } else { "" }
        $department = if ($txtDepartment -and $txtDepartment.Text) { $txtDepartment.Text.Trim() } else { "" }
        $terminationDate = if ($txtTermination -and $txtTermination.Text) { $txtTermination.Text.Trim() } else { "" }
        $external = if ($chkExternal) { $chkExternal.IsChecked } else { $false }
        $tl = if ($chkTL) { $chkTL.IsChecked } else { $false }
        $al = if ($chkAL) { $chkAL.IsChecked } else { $false }

        # **MailSuffix, License and OU**
        $mailSuffix = if ($comboBoxMailSuffix -and $comboBoxMailSuffix.SelectedValue) { 
            $comboBoxMailSuffix.SelectedValue.ToString() 
        } elseif ($global:Config.Contains("Company") -and $global:Config.Company.Contains("CompanyMailDomain")) { 
            $global:Config.Company["CompanyMailDomain"] 
        } else { "" }
        
        $selectedLicense = if ($comboBoxLicense -and $comboBoxLicense.SelectedValue) { 
            $comboBoxLicense.SelectedValue.ToString() 
        } else { "Standard" }
        
        $selectedOU = if ($comboBoxOU -and $comboBoxOU.SelectedValue) { 
            $comboBoxOU.SelectedValue.ToString() 
        } elseif ($global:Config.Contains("ADUserDefaults") -and $global:Config.ADUserDefaults.Contains("DefaultOU")) { 
            $global:Config.ADUserDefaults["DefaultOU"] 
        } else { "" }

        $selectedTLGroup = if ($comboBoxTLGroup -and $comboBoxTLGroup.SelectedValue) {
            $comboBoxTLGroup.SelectedValue.ToString()
        } else { "" }

        # **Create global object**
        $global:userData = [PSCustomObject]@{
            OU               = $selectedOU
            FirstName        = $firstName
            LastName         = $lastName
            DisplayName      = $displayName
            Description      = if ($txtDescription -and $txtDescription.Text) { $txtDescription.Text.Trim() } else { "" }
            EmailAddress     = $email
            OfficeRoom       = $office
            PhoneNumber      = $phone
            MobileNumber     = $mobile
            Position         = $position
            DepartmentField  = $department
            Ablaufdatum      = $terminationDate
            ADGroupsSelected = $ADGroupsSelected
            External         = $external
            TL               = $tl
            AL               = $al
            TLGroup          = $selectedTLGroup
            MailSuffix       = $mailSuffix
            License          = $selectedLicense
            UPNFormat        = if ($comboBoxDisplayTemplate -and $comboBoxDisplayTemplate.SelectedValue) { 
                $comboBoxDisplayTemplate.SelectedValue.ToString().Trim() 
            } elseif ($global:Config.Contains("DisplayNameUPNTemplates") -and $global:Config.DisplayNameUPNTemplates.Contains("DefaultUserPrincipalNameFormat")) { 
                $global:Config.DisplayNameUPNTemplates["DefaultUserPrincipalNameFormat"] 
            } else { "FIRSTNAME.LASTNAME" }
            AccountDisabled  = if ($chkAccountDisabled) { $chkAccountDisabled.IsChecked } else { $false }
            MustChangePassword   = if ($chkMustChangePassword) { $chkMustChangePassword.IsChecked } else { $false }
            PasswordNeverExpires = if ($chkPasswordNeverExpires) { $chkPasswordNeverExpires.IsChecked } else { $false }
            CompanyName      = if ($global:Config.Contains("Company") -and $global:Config.Company.Contains("CompanyNameFirma")) { 
                $global:Config.Company["CompanyNameFirma"] 
            } else { "" }
            CompanyStrasse   = if ($global:Config.Contains("Company") -and $global:Config.Company.Contains("CompanyStrasse")) { 
                $global:Config.Company["CompanyStrasse"] 
            } else { "" }
            CompanyPLZ       = if ($global:Config.Contains("Company") -and $global:Config.Company.Contains("CompanyPLZ")) { 
                $global:Config.Company["CompanyPLZ"] 
            } else { "" }
            CompanyOrt       = if ($global:Config.Contains("Company") -and $global:Config.Company.Contains("CompanyOrt")) { 
                $global:Config.Company["CompanyOrt"] 
            } else { "" }
            CompanyDomain    = if ($global:Config.Contains("Company") -and $global:Config.Company.Contains("CompanyActiveDirectoryDomain")) { 
                $global:Config.Company["CompanyActiveDirectoryDomain"] 
            } else { "" }
            CompanyTelefon   = if ($global:Config.Contains("Company") -and $global:Config.Company.Contains("CompanyTelefon")) { 
                $global:Config.Company["CompanyTelefon"] 
            } else { "" }
            CompanyCountry   = if ($global:Config.Contains("Company") -and $global:Config.Company.Contains("CompanyCountry")) { 
                $global:Config.Company["CompanyCountry"] 
            } else { "" }
            # Fallback for PasswordMode if used later
            PasswordMode     = 1 # Default: Generate password
        }        

        Write-DebugMessage "userData object created -> `n$($global:userData | Out-String)"
        Write-DebugMessage "Starting Invoke-Onboarding..."

        try {
            $global:result = Invoke-Onboarding -userData $global:userData -Config $global:Config
            if (-not $global:result) {
                throw "Invoke-Onboarding did not return a result!"
            }

            Write-DebugMessage "Invoke-Onboarding completed."
            [System.Windows.MessageBox]::Show("Onboarding successfully completed.`nSamAccountName: $($global:result.SamAccountName)`nUPN: $($global:result.UPN)", "Success")
        } catch {
            $errorMsg = "Error in Invoke-Onboarding: $($_.Exception.Message)"
            if ($_.Exception.InnerException) {
                $errorMsg += "`nInnerException: $($_.Exception.InnerException.Message)"
            }
            Write-DebugMessage "ERROR: $errorMsg"
            [System.Windows.MessageBox]::Show($errorMsg, "Error", 'OK', 'Error')
        }

    } catch {
        Write-DebugMessage "ERROR: Unhandled error: $($_.Exception.Message)"
        [System.Windows.MessageBox]::Show("Error: $($_.Exception.Message)", "Error", 'OK', 'Error')
    }
})
#endregion

# END "easyONBOARDING BUTTON"
Write-DebugMessage "TAB easyONBOARDING - BUTTON - CREATE PDF"

#region [Region 21.5 | PDF CREATION BUTTON]
# Implements the PDF creation button functionality
Write-DebugMessage "Setting up PDF creation button."
$btnPDF = $window.FindName("btnPDF")
if ($btnPDF) {
    Write-DebugMessage "TAB easyONBOARDING - BUTTON selected - CREATE PDF"
    $btnPDF.Add_Click({
        try {
            if (-not $global:result -or -not $global:result.SamAccountName) {
                [System.Windows.MessageBox]::Show("Please complete the onboarding process before creating the PDF.", "Error", 'OK', 'Error')
                Write-DebugMessage "$global:result is NULL or empty" "DEBUG"
                return
            }
            
            Write-DebugMessage "$global:result successfully loaded: $($global:result | Out-String)" "DEBUG"
            
            $SamAccountName = $global:result.SamAccountName

            # Check for PDFCreator module
            $pdfCreatorModule = Join-Path -Path $ModulesDir -ChildPath "PDFGenerator.psm1"
            $useExternalScript = $true
            
            if (Test-Path $pdfCreatorModule) {
                Write-DebugMessage "Found PDF Generator module at: $pdfCreatorModule"
                try {
                    Import-Module $pdfCreatorModule -Force -ErrorAction Stop
                    $useExternalScript = $false
                    Write-DebugMessage "Successfully imported PDF Generator module"
                } 
                catch {
                    Write-DebugMessage "Failed to import PDF Generator module: $($_.Exception.Message)"
                    Write-DebugMessage "Falling back to external script method"
                    $useExternalScript = $true
                }
            } else {
                Write-DebugMessage "PDF Generator module not found. Using external script method."
            }

            # Check for the Report section in config
            if (-not ($global:Config.Keys -contains "Report")) {
                [System.Windows.MessageBox]::Show("Error: The section [Report] is missing in the INI file.", "Error", 'OK', 'Error')
                return
            }

            $reportBranding = $global:Config["Report"]
            if (-not $reportBranding -or -not $reportBranding.Contains("ReportPath") -or [string]::IsNullOrWhiteSpace($reportBranding["ReportPath"])) {
                [System.Windows.MessageBox]::Show("ReportPath is missing or empty in [Report]. PDF cannot be created.", "Error", 'OK', 'Error')
                return
            }         

            # Create report directory if it doesn't exist
            $reportPath = $reportBranding["ReportPath"]
            if (-not (Test-Path $reportPath)) {
                try {
                    New-Item -Path $reportPath -ItemType Directory -Force | Out-Null
                    Write-DebugMessage "Created report directory: $reportPath"
                } catch {
                    [System.Windows.MessageBox]::Show("Failed to create report directory: $reportPath`nError: $($_.Exception.Message)", "Error", 'OK', 'Error')
                    return
                }
            }

            $htmlReportPath = Join-Path $reportPath "$SamAccountName.html"
            $pdfReportPath  = Join-Path $reportPath "$SamAccountName.pdf"

            # Check for HTML report (needs to be created by the onboarding process)
            if (-not (Test-Path $htmlReportPath)) {
                # Try to generate HTML if not found
                if ($global:userData -and (Get-Command -Name "Export-UserToHTML" -ErrorAction SilentlyContinue)) {
                    try {
                        Write-DebugMessage "HTML report not found, attempting to generate it now"
                        Export-UserToHTML -userData $global:userData -result $global:result -htmlPath $htmlReportPath
                        Write-DebugMessage "HTML report generated successfully"
                    } catch {
                        Write-DebugMessage "Failed to generate HTML report: $($_.Exception.Message)"
                        [System.Windows.MessageBox]::Show("HTML report not found: $htmlReportPath`nAttempt to generate it failed: $($_.Exception.Message)", "Error", 'OK', 'Error')
                        return
                    }
                } else {
                    [System.Windows.MessageBox]::Show("HTML report not found: $htmlReportPath`nThe PDF cannot be created.", "Error", 'OK', 'Error')
                    return
                }
            }

            # Use different methods based on availability
            if (-not $useExternalScript) {
                # Use the imported module directly
                try {
                    Write-DebugMessage "Converting HTML to PDF using PDF Generator module"
                    Convert-HTMLToPDF -HtmlFile $htmlReportPath -PdfFile $pdfReportPath
                    
                    if (Test-Path $pdfReportPath) {
                        [System.Windows.MessageBox]::Show("PDF successfully created: $pdfReportPath", "PDF Creation", 'OK', 'Information')
                        # Open the PDF automatically if configured
                        if ($reportBranding.Contains("OpenPDFAfterCreation") -and $reportBranding["OpenPDFAfterCreation"] -eq "true") {
                            Start-Process $pdfReportPath
                        }
                    } else {
                        throw "PDF file was not created"
                    }
                } catch {
                    Write-DebugMessage "Error in module-based PDF conversion: $($_.Exception.Message)"
                    Write-DebugMessage "Attempting fallback to external script method"
                    $useExternalScript = $true
                }
            }

            if ($useExternalScript) {
                # External script method with improved checking
                $wkhtmltopdfPath = $reportBranding["wkhtmltopdfPath"]
                if (-not (Test-Path $wkhtmltopdfPath)) {
                    [System.Windows.MessageBox]::Show("wkhtmltopdf.exe not found! Please check: $wkhtmltopdfPath", "Error", 'OK', 'Error')
                    return
                }          

                $pdfScript = Join-Path $PSScriptRoot "PDFCreator.ps1"
                if (-not (Test-Path $pdfScript)) {
                    # Try to find the script in alternate locations
                    $alternatePaths = @(
                        (Join-Path (Split-Path $PSScriptRoot -Parent) "PDFCreator.ps1"),
                        (Join-Path $ModulesDir "PDFCreator.ps1")
                    )
                    
                    $found = $false
                    foreach ($path in $alternatePaths) {
                        if (Test-Path $path) {
                            $pdfScript = $path
                            $found = $true
                            Write-DebugMessage "Found PDF creator script at alternate location: $pdfScript"
                            break
                        }
                    }
                    
                    if (-not $found) {
                        [System.Windows.MessageBox]::Show("PDF script 'PDFCreator.ps1' not found!", "Error", 'OK', 'Error')
                        return
                    }
                }

                $arguments = @(
                    "-NoProfile",
                    "-ExecutionPolicy", "Bypass",
                    "-File", "`"$pdfScript`"",
                    "-htmlFile", "`"$htmlReportPath`"",
                    "-pdfFile", "`"$pdfReportPath`"",
                    "-wkhtmltopdfPath", "`"$wkhtmltopdfPath`""
                )
                
                # Add any additional parameters from the INI file
                if ($reportBranding.Contains("PDFOptions")) {
                    $arguments += "-pdfOptions", "`"$($reportBranding["PDFOptions"])`""
                }
                
                Write-DebugMessage "Starting PDF conversion with external script: $pdfScript"
                $process = Start-Process -FilePath "powershell.exe" -ArgumentList $arguments -NoNewWindow -Wait -PassThru
                
                if ($process.ExitCode -eq 0 -and (Test-Path $pdfReportPath)) {
                    [System.Windows.MessageBox]::Show("PDF successfully created: $pdfReportPath", "PDF Creation", 'OK', 'Information')
                    # Open the PDF automatically if configured
                    if ($reportBranding.Contains("OpenPDFAfterCreation") -and $reportBranding["OpenPDFAfterCreation"] -eq "true") {
                        Start-Process $pdfReportPath
                    }
                } else {
                    Write-DebugMessage "PDF creation process exited with code: $($process.ExitCode)"
                    [System.Windows.MessageBox]::Show("Error creating PDF. Process exited with code: $($process.ExitCode)", "Error", 'OK', 'Error')
                }
            }
        } catch {
            Write-DebugMessage "Unhandled error in PDF creation: $($_.Exception.Message)"
            [System.Windows.MessageBox]::Show("Error creating PDF: $($_.Exception.Message)", "Error", 'OK', 'Error')
        }
    })
}
#endregion

Write-DebugMessage "TAB easyONBOARDING - BUTTON Info"

#region [Region 21.6 | INFO BUTTON]
# Implements the info button functionality
Write-DebugMessage "Setting up info button."
$btnInfo = $window.FindName("btnInfo")
if ($btnInfo) {
    Write-DebugMessage "TAB easyONBOARDING - BUTTON selected: Info"
    $btnInfo.Add_Click({
        $infoFilePath = Join-Path $PSScriptRoot "easyIT.txt"

        if (Test-Path $infoFilePath) {
            Start-Process -FilePath $infoFilePath
        } else {
            [System.Windows.MessageBox]::Show("The file easyIT.txt was not found!", "Error", 'OK', 'Error')
        }
    })
}
#endregion

Write-DebugMessage "TAB easyONBOARDING - BUTTON Close"

#region [Region 21.7 | CLOSE BUTTON]
# Implements the close button functionality
Write-DebugMessage "Setting up close button."
$btnClose = $window.FindName("btnClose")
if ($btnClose) {
    Write-DebugMessage "TAB easyONBOARDING - BUTTON selected: Close"
    $btnClose.Add_Click({
        $window.Close()
    })
}
#endregion

Write-DebugMessage "GUI: TAB easyADUpdate loaded"

#region [Region 21.8 | AD UPDATE TAB]
# Implements the AD Update tab functionality
Write-DebugMessage "AD Update tab loaded."

# Find the AD Update tab
$adUpdateTab = $window.FindName("Tab_ADUpdate")
if ($adUpdateTab) {

    try {
        # Retrieve search and basic controls – names must match the XAML definitions
        $txtSearchADUpdate   = $adUpdateTab.FindName("txtSearchADUpdate")
        $btnSearchADUpdate   = $adUpdateTab.FindName("btnSearchADUpdate")
        $lstUsersADUpdate    = $adUpdateTab.FindName("lstUsersADUpdate")
        
        # Corrected: The button is named btnRefreshOU_Kopieren in the XAML
        $btnRefreshADUserList = $adUpdateTab.FindName("btnRefreshOU_Kopieren") # Corrected name for the refresh button
        
        # Action buttons
        $btnADUserUpdate     = $adUpdateTab.FindName("btnADUserUpdate")
        $btnADUserCancel     = $adUpdateTab.FindName("btnADUserCancel")
        $btnAddGroupUpdate   = $adUpdateTab.FindName("btnAddGroupUpdate")
        $btnRemoveGroupUpdate= $adUpdateTab.FindName("btnRemoveGroupUpdate")
        
        # Basic information controls
        $txtFirstNameUpdate  = $adUpdateTab.FindName("txtFirstNameUpdate")
        $txtLastNameUpdate   = $adUpdateTab.FindName("txtLastNameUpdate")
        $txtDisplayNameUpdate= $adUpdateTab.FindName("txtDisplayNameUpdate")
        $txtEmailUpdate      = $adUpdateTab.FindName("txtEmailUpdate")
        $txtDepartmentUpdate = $adUpdateTab.FindName("txtDepartmentUpdate")
        
        # Contact information controls
        $txtPhoneUpdate      = $adUpdateTab.FindName("txtPhoneUpdate")
        $txtMobileUpdate     = $adUpdateTab.FindName("txtMobileUpdate")
        $txtOfficeUpdate     = $adUpdateTab.FindName("txtOfficeUpdate")
        
        # Account options
        $chkAccountEnabledUpdate      = $adUpdateTab.FindName("chkAccountEnabledUpdate")
        $chkPasswordNeverExpiresUpdate= $adUpdateTab.FindName("chkPasswordNeverExpiresUpdate")
        $chkMustChangePasswordUpdate  = $adUpdateTab.FindName("chkMustChangePasswordUpdate")
        
        # Extended properties controls
        $txtManagerUpdate    = $adUpdateTab.FindName("txtManagerUpdate")
        $txtJobTitleUpdate   = $adUpdateTab.FindName("txtJobTitleUpdate")
        $txtLocationUpdate   = $adUpdateTab.FindName("txtLocationUpdate")
        $txtEmployeeIDUpdate = $adUpdateTab.FindName("txtEmployeeIDUpdate")
        
        # Group management control
        $lstGroupsUpdate     = $adUpdateTab.FindName("lstGroupsUpdate")
        if ($lstGroupsUpdate) {
            $lstGroupsUpdate.Items.Clear()
        }
    }
    catch {
        Write-DebugMessage "Error loading AD Update tab controls: $($_.Exception.Message)"
        Write-LogMessage -Message "Failed to initialize AD Update tab controls: $($_.Exception.Message)" -LogLevel "ERROR"
    }
    
    # Register search event
    if ($btnSearchADUpdate -and $txtSearchADUpdate -and $lstUsersADUpdate) {
        Write-DebugMessage "Registering event for btnSearchADUpdate in Tab_ADUpdate"
        $btnSearchADUpdate.Add_Click({
            $searchTerm = $txtSearchADUpdate.Text.Trim()
            if ([string]::IsNullOrWhiteSpace($searchTerm)) {
                [System.Windows.MessageBox]::Show("Please enter a search term.", "Info", "OK", "Information")
                return
            }
            Write-DebugMessage "Searching AD users for '$searchTerm'..."
            try {
                # Improved search with correct filter
                $filter = "((DisplayName -like '*$searchTerm*') -or (SamAccountName -like '*$searchTerm*') -or (mail -like '*$searchTerm*'))"
                Write-DebugMessage "Using AD filter: $filter"
                
                $allMatches = Get-ADUser -Filter $filter -Properties DisplayName, SamAccountName, mail, EmailAddress -ErrorAction Stop

                if ($null -eq $allMatches -or ($allMatches.Count -eq 0)) {
                    Write-DebugMessage "No users found for search term: $searchTerm"
                    [System.Windows.MessageBox]::Show("No users found for: $searchTerm", "No Results", "OK", "Information")
                    return
                }

                # Create an array of custom objects for the ListView
                $results = @()
                foreach ($user in $allMatches) {
                    $emailAddress = if ($user.mail) { $user.mail } elseif ($user.EmailAddress) { $user.EmailAddress } else { "" }
                    $results += [PSCustomObject]@{
                        DisplayName = $user.DisplayName
                        SamAccountName = $user.SamAccountName
                        Email = $emailAddress
                    }
                }
                
                # Debug outputs for diagnostics
                Write-DebugMessage "Found $($results.Count) users matching '$searchTerm'"
                
                # First set ItemsSource to null and clear items
                $lstUsersADUpdate.ItemsSource = $null
                $lstUsersADUpdate.Items.Clear()
                
                # Manually add items to the ListView
                foreach ($item in $results) {
                    $lstUsersADUpdate.Items.Add($item)
                }

                if ($results.Count -eq 0) {
                    [System.Windows.MessageBox]::Show("No users found for: $searchTerm", "No Results", "OK", "Information")
                }
                else {
                    Write-DebugMessage "Successfully populated list with $($lstUsersADUpdate.Items.Count) items"
                }
            }
            catch {
                Write-DebugMessage "Error during search: $($_.Exception.Message)"
                [System.Windows.MessageBox]::Show("Error during search: $($_.Exception.Message)", "Error", "OK", "Error")
            }
        })
    }
    else {
        Write-DebugMessage "Required search controls missing in Tab_ADUpdate"
    }

    # Register Refresh User List event - We use the corrected button name
    if ($btnRefreshADUserList -and $lstUsersADUpdate) {
        Write-DebugMessage "Registering event for Refresh User List button in Tab_ADUpdate"
        $btnRefreshADUserList.Add_Click({
            Write-DebugMessage "Refreshing AD users list..."
            try {
                # For large AD environments, limit the number of loaded users
                # We filter by enabled users and sort by modification date
                $filter = "Enabled -eq 'True'"
                Write-DebugMessage "Loading active AD users with filter: $filter"

                
                # Load up to 500 active users
                $allUsers = Get-ADUser -Filter $filter -Properties DisplayName, SamAccountName, mail, EmailAddress, WhenChanged -ResultSetSize 500 |
                            Sort-Object -Property WhenChanged -Descending

                if ($null -eq $allUsers -or ($allUsers.Count -eq 0)) {
                    Write-DebugMessage "No users found using the filter: $filter"
                    [System.Windows.MessageBox]::Show("No users found.", "No Results", "OK", "Information")
                    return
                }

                # Create an array of custom objects for the ListView
                $results = @()
                foreach ($user in $allUsers) {
                    # Read email address from mail or EmailAddress attribute depending on availability
                    $emailAddress = if ($user.mail) { $user.mail } elseif ($user.EmailAddress) { $user.EmailAddress } else { "" }
                    
                    # Create a custom object with exactly the properties bound in the XAML
                    $results += [PSCustomObject]@{
                        DisplayName = $user.DisplayName
                        SamAccountName = $user.SamAccountName
                        Email = $emailAddress
                    }
                }
                
                # Debug output for the number of loaded users
                Write-DebugMessage "Loaded $($results.Count) active AD users"
                
                # Update ListView
                $lstUsersADUpdate.ItemsSource = $null
                $lstUsersADUpdate.Items.Clear()
                
                # Manually add items to the ListView
                foreach ($item in $results) {
                    $lstUsersADUpdate.Items.Add($item)
                }

                [System.Windows.MessageBox]::Show("$($lstUsersADUpdate.Items.Count) users were loaded.", "Refresh Complete", "OK", "Information")
                Write-DebugMessage "Successfully populated list with $($lstUsersADUpdate.Items.Count) users"
            }
            catch {
                Write-DebugMessage "Error loading user list: $($_.Exception.Message)"
                [System.Windows.MessageBox]::Show("Error loading user list: $($_.Exception.Message)", "Error", "OK", "Error")
            }
        })
    }

    # Populate update fields when a user is selected
    if ($lstUsersADUpdate) {
        $lstUsersADUpdate.Add_SelectionChanged({
            if ($null -ne $lstUsersADUpdate.SelectedItem) {
                $selectedUser = $lstUsersADUpdate.SelectedItem
                $samAccountName = $selectedUser.SamAccountName
                if ([string]::IsNullOrWhiteSpace($samAccountName)) {
                    Write-DebugMessage "Selected user has no SamAccountName"
                    return
                }
                try {
                    $adUser = Get-ADUser -Identity $samAccountName -Properties DisplayName, GivenName, Surname, EmailAddress, Department, Title, OfficePhone, Mobile, physicalDeliveryOfficeName, Enabled, PasswordNeverExpires, Manager, employeeID, MemberOf -ErrorAction Stop
                    if ($adUser) {
                        if ($txtFirstNameUpdate) { $txtFirstNameUpdate.Text = $(if ($adUser.GivenName) { $adUser.GivenName } else { "" }) }
                        if ($txtLastNameUpdate) { $txtLastNameUpdate.Text = $(if ($adUser.Surname) { $adUser.Surname } else { "" }) }
                        if ($txtDisplayNameUpdate) { $txtDisplayNameUpdate.Text = $(if ($adUser.DisplayName) { $adUser.DisplayName } else { "" }) }
                        if ($txtEmailUpdate) { $txtEmailUpdate.Text = $(if ($adUser.EmailAddress) { $adUser.EmailAddress } else { "" }) }
                        if ($txtDepartmentUpdate) { $txtDepartmentUpdate.Text = $(if ($adUser.Department) { $adUser.Department } else { "" }) }
                        if ($txtPhoneUpdate) { $txtPhoneUpdate.Text = $(if ($adUser.OfficePhone) { $adUser.OfficePhone } else { "" }) }
                        if ($txtMobileUpdate) { $txtMobileUpdate.Text = $(if ($adUser.Mobile) { $adUser.Mobile } else { "" }) }
                        if ($txtOfficeUpdate) { $txtOfficeUpdate.Text = $(if ($adUser.physicalDeliveryOfficeName) { $adUser.physicalDeliveryOfficeName } else { "" }) }
                        if ($txtJobTitleUpdate) { $txtJobTitleUpdate.Text = $(if ($adUser.Title) { $adUser.Title } else { "" }) }
                        if ($txtLocationUpdate) { $txtLocationUpdate.Text = $(if ($adUser.physicalDeliveryOfficeName) { $adUser.physicalDeliveryOfficeName } else { "" }) }
                        if ($txtEmployeeIDUpdate) { $txtEmployeeIDUpdate.Text = $(if ($adUser.employeeID) { $adUser.employeeID } else { "" }) }
                        
                        if ($chkAccountEnabledUpdate) { $chkAccountEnabledUpdate.IsChecked = $adUser.Enabled }
                        if ($chkPasswordNeverExpiresUpdate) { $chkPasswordNeverExpiresUpdate.IsChecked = $adUser.PasswordNeverExpires }
                        if ($chkMustChangePasswordUpdate) { $chkMustChangePasswordUpdate.IsChecked = $false }
                        
                        if ($txtManagerUpdate) {
                            if (-not [string]::IsNullOrEmpty($adUser.Manager)) {
                                try {
                                    $manager = Get-ADUser -Identity $adUser.Manager -Properties DisplayName -ErrorAction Stop
                                    $txtManagerUpdate.Text = $(if ($manager.DisplayName) { $manager.DisplayName } else { $manager.SamAccountName })
                                } catch {
                                    $txtManagerUpdate.Text = ""
                                    Write-DebugMessage "Error retrieving manager: $($_.Exception.Message)"
                                }
                            } else {
                                $txtManagerUpdate.Text = ""
                            }
                        }
                        
                        if ($lstGroupsUpdate) {
                            $lstGroupsUpdate.Items.Clear()
                            $memberOfGroups = @($adUser.MemberOf)
                            if ($memberOfGroups.Count -gt 0) {
                                foreach ($group in $memberOfGroups) {
                                    try {
                                        if (-not [string]::IsNullOrEmpty($group)) {
                                            $groupObj = Get-ADGroup -Identity $group -Properties Name -ErrorAction Stop
                                            if ($groupObj -and -not [string]::IsNullOrEmpty($groupObj.Name)) {
                                                [void]$lstGroupsUpdate.Items.Add($groupObj.Name)
                                            }
                                        }
                                    }
                                    catch {
                                        Write-DebugMessage "Error retrieving group details: $($_.Exception.Message)"
                                    }
                                }
                            }
                        }
                        Write-DebugMessage "Fields populated for user: $samAccountName"
                    }
                    else {
                        Write-DebugMessage "AD User object is null for user: $samAccountName"
                        [System.Windows.MessageBox]::Show("The selected user could not be found.", "Error", "OK", "Error")
                    }
                }
                catch {
                    Write-DebugMessage "Error loading user details: $($_.Exception.Message)"
                    [System.Windows.MessageBox]::Show("Error loading user details: $($_.Exception.Message)", "Error", "OK", "Error")
                }
            }
        })
    }

    # Update User button event
    if ($btnADUserUpdate) {
        Write-DebugMessage "Registering event handler for btnADUserUpdate"
        $btnADUserUpdate.Add_Click({
            try {
                if (-not $lstUsersADUpdate -or -not $lstUsersADUpdate.SelectedItem) {
                    [System.Windows.MessageBox]::Show("Please select a user from the list.", "Missing Selection", "OK", "Warning")
                    return
                }
                $userToUpdate = $lstUsersADUpdate.SelectedItem.SamAccountName
                if ([string]::IsNullOrWhiteSpace($userToUpdate)) {
                    [System.Windows.MessageBox]::Show("Invalid username selected.", "Error", "OK", "Warning")
                    return
                }
                Write-DebugMessage "Updating user: $userToUpdate"
                if (-not $txtDisplayNameUpdate -or [string]::IsNullOrWhiteSpace($txtDisplayNameUpdate.Text)) {
                    [System.Windows.MessageBox]::Show("The display name cannot be empty.", "Validation", "OK", "Warning")
                    return
                }
                
                # build parameters for Set-ADUser
                $paramUpdate = @{ Identity = $userToUpdate }
                if ($txtDisplayNameUpdate -and -not [string]::IsNullOrWhiteSpace($txtDisplayNameUpdate.Text)) {
                    $paramUpdate["DisplayName"] = $txtDisplayNameUpdate.Text.Trim()
                }
                if ($chkAccountEnabledUpdate) {
                    $paramUpdate["Enabled"] = $chkAccountEnabledUpdate.IsChecked
                }
                if ($chkPasswordNeverExpiresUpdate) {
                    $paramUpdate["PasswordNeverExpires"] = $chkPasswordNeverExpiresUpdate.IsChecked
                }
                if ($txtFirstNameUpdate -and -not [string]::IsNullOrWhiteSpace($txtFirstNameUpdate.Text)) {
                    $paramUpdate["GivenName"] = $txtFirstNameUpdate.Text.Trim()
                }
                if ($txtLastNameUpdate -and -not [string]::IsNullOrWhiteSpace($txtLastNameUpdate.Text)) {
                    $paramUpdate["Surname"] = $txtLastNameUpdate.Text.Trim()
                }
                if ($txtEmailUpdate -and -not [string]::IsNullOrWhiteSpace($txtEmailUpdate.Text)) {
                    $paramUpdate["EmailAddress"] = $txtEmailUpdate.Text.Trim()
                }
                if ($txtDepartmentUpdate -and -not [string]::IsNullOrWhiteSpace($txtDepartmentUpdate.Text)) {
                    $paramUpdate["Department"] = $txtDepartmentUpdate.Text.Trim()
                }
                if ($txtPhoneUpdate -and -not [string]::IsNullOrWhiteSpace($txtPhoneUpdate.Text)) {
                    $paramUpdate["OfficePhone"] = $txtPhoneUpdate.Text.Trim()
                }
                if ($txtMobileUpdate -and -not [string]::IsNullOrWhiteSpace($txtMobileUpdate.Text)) {
                    $paramUpdate["MobilePhone"] = $txtMobileUpdate.Text.Trim()
                }
                if ($txtJobTitleUpdate -and -not [string]::IsNullOrWhiteSpace($txtJobTitleUpdate.Text)) {
                    $paramUpdate["Title"] = $txtJobTitleUpdate.Text.Trim()
                }
                if ($txtManagerUpdate -and -not [string]::IsNullOrWhiteSpace($txtManagerUpdate.Text)) {
                    try {
                        $managerObj = Get-ADUser -Filter { (SamAccountName -eq $txtManagerUpdate.Text.Trim()) -or (DisplayName -eq $txtManagerUpdate.Text.Trim()) } -ErrorAction Stop
                        if ($managerObj) {
                            if ($managerObj -is [array]) { $managerObj = $managerObj[0] }
                            $paramUpdate["Manager"] = $managerObj.DistinguishedName
                        }
                        else {
                            [System.Windows.MessageBox]::Show("The specified manager could not be found: $($txtManagerUpdate.Text)", "Warning", "OK", "Warning")
                        }
                    }
                    catch {
                        Write-DebugMessage "Manager lookup error: $($_.Exception.Message)"
                        [System.Windows.MessageBox]::Show("The specified manager could not be found: $($txtManagerUpdate.Text)", "Warning", "OK", "Warning")
                    }
                }
                
                if ($paramUpdate.Count -gt 1) {
                    Set-ADUser @paramUpdate -ErrorAction Stop
                    Write-DebugMessage "Updated user attributes: $($paramUpdate.Keys -join ', ')"
                }
                
                # Update additional attributes using the -Replace parameter
                $replaceHash = @{ }
                if ($txtOfficeUpdate -and -not [string]::IsNullOrWhiteSpace($txtOfficeUpdate.Text)) {
                    $replaceHash["physicalDeliveryOfficeName"] = $txtOfficeUpdate.Text.Trim()
                }
                if ($txtLocationUpdate -and -not [string]::IsNullOrWhiteSpace($txtLocationUpdate.Text)) {
                    $replaceHash["l"] = $txtLocationUpdate.Text.Trim()
                }
                if ($txtEmployeeIDUpdate -and -not [string]::IsNullOrWhiteSpace($txtEmployeeIDUpdate.Text)) {
                    $replaceHash["employeeID"] = $txtEmployeeIDUpdate.Text.Trim()
                }
                if ($replaceHash.Count -gt 0) {
                    Set-ADUser -Identity $userToUpdate -Replace $replaceHash -ErrorAction Stop
                    Write-DebugMessage "Updated replace attributes: $($replaceHash.Keys -join ', ')"
                }
                if ($chkMustChangePasswordUpdate -ne $null) {
                    Set-ADUser -Identity $userToUpdate -ChangePasswordAtLogon $chkMustChangePasswordUpdate.IsChecked -ErrorAction Stop
                    Write-DebugMessage "Set ChangePasswordAtLogon to $($chkMustChangePasswordUpdate.IsChecked)"
                }
                [System.Windows.MessageBox]::Show("The user '$userToUpdate' was successfully updated.", "Success", "OK", "Information")
                Write-LogMessage -Message "AD update for $userToUpdate by $($env:USERNAME)" -LogLevel "INFO"
            }
            catch {
                Write-DebugMessage "AD User update error: $($_.Exception.Message)"
                [System.Windows.MessageBox]::Show("Error updating: $($_.Exception.Message)", "Error", "OK", "Error")
            }
        })
    }

    # Cancel button – reset form fields
    if ($btnADUserCancel) {
        Write-DebugMessage "Registering event handler for btnADUserCancel"
        $btnADUserCancel.Add_Click({
            Write-DebugMessage "Resetting update form"
            $textFields = @("txtFirstNameUpdate","txtLastNameUpdate","txtDisplayNameUpdate","txtEmailUpdate",
                              "txtPhoneUpdate","txtMobileUpdate","txtDepartmentUpdate","txtOfficeUpdate",
                              "txtManagerUpdate","txtJobTitleUpdate","txtLocationUpdate","txtEmployeeIDUpdate")
            foreach ($name in $textFields) {
                $ctrl = $adUpdateTab.FindName($name)
                if ($ctrl) { $ctrl.Text = "" }
            }
            if ($chkAccountEnabledUpdate) { $chkAccountEnabledUpdate.IsChecked = $false }
            if ($chkPasswordNeverExpiresUpdate) { $chkPasswordNeverExpiresUpdate.IsChecked = $false }
            if ($chkMustChangePasswordUpdate) { $chkMustChangePasswordUpdate.IsChecked = $false }
            if ($lstUsersADUpdate) {
                $lstUsersADUpdate.ItemsSource = $null
                $lstUsersADUpdate.Items.Clear()
                $lstUsersADUpdate.SelectedIndex = -1
            }
            if ($lstGroupsUpdate) { $lstGroupsUpdate.Items.Clear() }
            [System.Windows.MessageBox]::Show("Form reset.", "Reset", "OK", "Information")
        })
    }

    # Add Group button functionality
    if ($btnAddGroupUpdate) {
        $btnAddGroupUpdate.Add_Click({
            if (-not $lstUsersADUpdate -or -not $lstUsersADUpdate.SelectedItem) {
                [System.Windows.MessageBox]::Show("Please select a user first.", "Note", "OK", "Information")
                return
            }
            $selectedUser = $lstUsersADUpdate.SelectedItem.SamAccountName
            if ([string]::IsNullOrWhiteSpace($selectedUser)) {
                [System.Windows.MessageBox]::Show("Invalid username selected.", "Error", "OK", "Error")
                return
            }
            # Use an input box (WPF-friendly via VB) to get the group name
            $groupName = [Microsoft.VisualBasic.Interaction]::InputBox("Enter the name of the AD group:", "Add Group", "")
            if (-not [string]::IsNullOrEmpty($groupName)) {
                try {
                    $group = Get-ADGroup -Identity $groupName -ErrorAction Stop
                    if (-not $group) {
                        throw "Group '$groupName' could not be found."
                    }
                    $isMember = $false
                    try {
                        $members = Get-ADGroupMember -Identity $groupName -ErrorAction Stop
                        if ($members -is [array]) {
                            $isMember = ($members | Where-Object { $_.SamAccountName -eq $selectedUser }).Count -gt 0
                        }
                        else {
                            $isMember = $members.SamAccountName -eq $selectedUser
                        }
                    }
                    catch {
                        Write-DebugMessage "Error checking group membership: $($_.Exception.Message)"
                    }
                    if ($isMember) {
                        [System.Windows.MessageBox]::Show("The user is already a member of the group '$groupName'.", "Information", "OK", "Information")
                        return
                    }
                    Add-ADGroupMember -Identity $groupName -Members $selectedUser -ErrorAction Stop
                    if ($lstGroupsUpdate) { [void]$lstGroupsUpdate.Items.Add($group.Name) }
                    [System.Windows.MessageBox]::Show("User added to group '$groupName'.", "Success", "OK", "Information")
                    Write-LogMessage -Message "User $selectedUser added to group $groupName by $($env:USERNAME)" -LogLevel "INFO"
                }
                catch {
                    Write-DebugMessage "Error adding user to group: $($_.Exception.Message)"
                    [System.Windows.MessageBox]::Show("Error adding to group: $($_.Exception.Message)", "Error", "OK", "Error")
                }
            }
        })
    }

    # Remove Group button functionality
    if ($btnRemoveGroupUpdate) {
        $btnRemoveGroupUpdate.Add_Click({
            if (-not $lstUsersADUpdate -or -not $lstUsersADUpdate.SelectedItem) {
                [System.Windows.MessageBox]::Show("Please select a user first.", "Note", "OK", "Information")
                return
            }
            if (-not $lstGroupsUpdate -or -not $lstGroupsUpdate.SelectedItem) {
                [System.Windows.MessageBox]::Show("Please select a group from the list.", "Note", "OK", "Information")
                return
            }
            $selectedUser = $lstUsersADUpdate.SelectedItem.SamAccountName
            if ([string]::IsNullOrWhiteSpace($selectedUser)) {
                [System.Windows.MessageBox]::Show("Invalid username selected.", "Error", "OK", "Error")
                return
            }
            $selectedGroup = $lstGroupsUpdate.SelectedItem.ToString()
            if ([string]::IsNullOrWhiteSpace($selectedGroup)) {
                [System.Windows.MessageBox]::Show("Invalid group name selected.", "Error", "OK", "Error")
                return
            }
            try {
                $confirmation = [System.Windows.MessageBox]::Show(
                    "Do you really want to remove the user '$selectedUser' from the group '$selectedGroup'?",
                    "Confirmation",
                    "YesNo",
                    "Question"
                )
                if ($confirmation -eq [System.Windows.MessageBoxResult]::Yes) {
                    Remove-ADGroupMember -Identity $selectedGroup -Members $selectedUser -Confirm:$false -ErrorAction Stop
                    if ($lstGroupsUpdate) { $lstGroupsUpdate.Items.Remove($selectedGroup) }
                    [System.Windows.MessageBox]::Show("User removed from group '$selectedGroup'.", "Success", "OK", "Information")
                    Write-LogMessage -Message "User $selectedUser removed from group $selectedGroup by $($env:USERNAME)" -LogLevel "INFO"
                }
            }
            catch {
                Write-DebugMessage "Error removing user from group: $($_.Exception.Message)"
                [System.Windows.MessageBox]::Show("Error removing from group: $($_.Exception.Message)", "Error", "OK", "Error")
            }
        })
    }
}
else {
    Write-DebugMessage "AD Update tab (Tab_ADUpdate) not found in XAML."
}
#endregion

#region [Region 21.10 | SETTINGS TAB]
# Implements the Settings tab functionality for INI editing
Write-DebugMessage "Settings tab loaded."

$tabINIEditor = $window.FindName("Tab_INIEditor")
if ($tabINIEditor) {
    # Register event handlers for the INI editor UI
    $listViewINIEditor = $tabINIEditor.FindName("listViewINIEditor")
    $dataGridINIEditor = $tabINIEditor.FindName("dataGridINIEditor")
    
    if ($listViewINIEditor) {
        $listViewINIEditor.Add_SelectionChanged({
            if ($null -ne $listViewINIEditor.SelectedItem) {
                $selectedSection = $listViewINIEditor.SelectedItem.SectionName
                Write-DebugMessage "Section selected: $selectedSection"
                Import-SectionSettings -SectionName $selectedSection -DataGrid $dataGridINIEditor -INIPath $global:INIPath
            }
        })
    }

    # Setup DataGrid cell edit handling
    if ($dataGridINIEditor) {
        # Make sure DataGrid is editable
        $dataGridINIEditor.IsReadOnly = $false
        
        $dataGridINIEditor.Add_CellEditEnding({
            param($eventSender, $e)
            
            try {
                if ($e.EditAction -eq [System.Windows.Controls.DataGridEditAction]::Commit) {
                    $item = $e.Row.Item
                    $column = $e.Column
                    
                    if ($null -ne $listViewINIEditor.SelectedItem) {
                        $sectionName = $listViewINIEditor.SelectedItem.SectionName
                        
                        if ($column.Header -eq "Key") {
                            # Key was edited - update the dictionary
                            $oldKey = $item.OriginalKey
                            $newKey = $item.Key
                            
                            if ($oldKey -ne $newKey) {
                                Write-DebugMessage "Changing key from '$oldKey' to '$newKey' in section $sectionName"
                                
                                # Get the current value
                                $value = $global:Config[$sectionName][$oldKey]
                                
                                # Remove old key and add new key
                                $global:Config[$sectionName].Remove($oldKey)
                                $global:Config[$sectionName][$newKey] = $value
                                
                                # Update OriginalKey for future edits
                                $item.OriginalKey = $newKey
                            }
                        }
                        elseif ($column.Header -eq "Value") {
                            # Value was edited
                            $key = $item.Key
                            $newValue = $item.Value
                            
                            Write-DebugMessage "Updating value for [$sectionName] $key to: $newValue"
                            $global:Config[$sectionName][$key] = $newValue
                        }
                    }
                }
            }
            catch {
                Write-DebugMessage "Error handling cell edit: $($_.Exception.Message)"
            }
        })
    }

    # Add key button handler
    $btnAddKey = $tabINIEditor.FindName("btnAddKey")
    if ($btnAddKey) {
        $btnAddKey.Add_Click({
            try {
                if ($null -eq $listViewINIEditor.SelectedItem) {
                    [System.Windows.MessageBox]::Show("Please select a section first.", "Note", "OK", "Information")
                    return
                }
                
                $sectionName = $listViewINIEditor.SelectedItem.SectionName
                
                # Prompt for new key and value
                $keyName = [Microsoft.VisualBasic.Interaction]::InputBox("Enter the name of the new key:", "New Key", "")
                if ([string]::IsNullOrWhiteSpace($keyName)) { return }
                
                $keyValue = [Microsoft.VisualBasic.Interaction]::InputBox("Enter the value for '$keyName':", "New Value", "")
                
                # Add to config
                $global:Config[$sectionName][$keyName] = $keyValue
                
                # Refresh view
                Import-SectionSettings -SectionName $sectionName -DataGrid $dataGridINIEditor -INIPath $global:INIPath
                
                Write-DebugMessage "Added new key '$keyName' with value '$keyValue' to section '$sectionName'"
            }
            catch {
                Write-DebugMessage "Error adding new key: $($_.Exception.Message)"
                [System.Windows.MessageBox]::Show("Error adding: $($_.Exception.Message)", "Error", "OK", "Error")
            }
        })
    }

    # Remove key button handler
    $btnRemoveKey = $tabINIEditor.FindName("btnRemoveKey")
    if ($btnRemoveKey) {
        $btnRemoveKey.Add_Click({
            try {
                if ($null -eq $dataGridINIEditor.SelectedItem) {
                    [System.Windows.MessageBox]::Show("Please select a key first.", "Note", "OK", "Information")
                    return
                }
                
                $selectedItem = $dataGridINIEditor.SelectedItem
                $keyToRemove = $selectedItem.Key
                $sectionName = $listViewINIEditor.SelectedItem.SectionName
                
                $confirmation = [System.Windows.MessageBox]::Show(
                    "Do you really want to delete the key '$keyToRemove'?",
                    "Confirmation",
                    "YesNo",
                    "Question"
                )
                
                if ($confirmation -eq [System.Windows.MessageBoxResult]::Yes) {
                    # Remove from config
                    $global:Config[$sectionName].Remove($keyToRemove)
                    
                    # Refresh view
                    Import-SectionSettings -SectionName $sectionName -DataGrid $dataGridINIEditor -INIPath $global:INIPath
                    
                    Write-DebugMessage "Removed key '$keyToRemove' from section '$sectionName'"
                }
            }
            catch {
                Write-DebugMessage "Error removing key: $($_.Exception.Message)"
                [System.Windows.MessageBox]::Show("Error deleting: $($_.Exception.Message)", "Error", "OK", "Error")
            }
        })
    }

    # Register the save button handler
    $btnSaveINIChanges = $tabINIEditor.FindName("btnSaveINIChanges")
    if ($btnSaveINIChanges) {
        $btnSaveINIChanges.Add_Click({
            try {
                $result = Save-INIChanges -INIPath $global:INIPath
                if ($result) {
                    [System.Windows.MessageBox]::Show("INI file successfully saved.", "Success", "OK", "Information")
                    # Reload the INI to reflect any changes in the UI
                    $global:Config = Get-IniContent -Path $global:INIPath
                    Import-INIEditorData
                    
                    # Update current selection if available
                    if ($listViewINIEditor.SelectedItem -ne $null) {
                        $selectedSection = $listViewINIEditor.SelectedItem.SectionName
                        Import-SectionSettings -SectionName $selectedSection -DataGrid $dataGridINIEditor -INIPath $global:INIPath
                    }
                }
                else {
                    [System.Windows.MessageBox]::Show("There was a problem saving the INI file.", "Error", "OK", "Error")
                }
            }
            catch {
                Write-DebugMessage "Error in Save button handler: $($_.Exception.Message)"
                [System.Windows.MessageBox]::Show("Error: $($_.Exception.Message)", "Error", "OK", "Error")
            }
        })
    }
    
    # Info button event handler
    $btnINIEditorInfo = $tabINIEditor.FindName("btnINIEditorInfo")
    if ($btnINIEditorInfo) {
        $btnINIEditorInfo.Add_Click({
            [System.Windows.MessageBox]::Show(
                "INI Editor Help:
                
1. Select a section from the left list
2. Edit values directly in the table
3. Use the buttons to add or remove entries
4. Save your changes with the Save button

INI file: $global:INIPath", 
                "INI Editor Help", "OK", "Information")
        })
    }
    
    # Initialize the editor with data when the tab is first loaded
    $tabINIEditor.Add_Loaded({
        Write-DebugMessage "INI Editor loaded with file path: $global:INIPath"
        Import-INIEditorData
    })
    
    $btnINIEditorClose = $tabINIEditor.FindName("btnINIEditorClose")
    if ($btnINIEditorClose) {
        $btnINIEditorClose.Add_Click({
            $window.Close()
        })
    }
}
#endregion

#region [Region 22 | MAIN GUI EXECUTION]
# Main code block that starts and handles the GUI dialog
Write-DebugMessage "GUI: Main GUI execution"

# Apply branding settings from INI file before showing the window
try {
    Write-DebugMessage "Applying branding settings from INI file"
    
    # Set window properties
    if ($global:Config.Contains("WPFGUI")) {
        # Set window title with dynamic replacements
        $window.Title = $global:Config["WPFGUI"]["HeaderText"] -replace "{ScriptVersion}", 
            $global:Config["ScriptInfo"]["ScriptVersion"] -replace "{LastUpdate}", 
            $global:Config["ScriptInfo"]["LastUpdate"] -replace "{Author}", 
            $global:Config["ScriptInfo"]["Author"]
            
        # Set application name if specified
        if ($global:Config["WPFGUI"].Contains("APPName")) {
            $window.Title = $global:Config["WPFGUI"]["APPName"]
        }
        
        # Set window theme color if specified and valid
        if ($global:Config["WPFGUI"].Contains("ThemeColor")) {
            try {
                $color = $global:Config["WPFGUI"]["ThemeColor"]
                # Try to create a brush from the color string
                $themeColorBrush = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.ColorConverter]::ConvertFromString($color))
                $window.Background = $themeColorBrush
            } catch {
                Write-DebugMessage "Invalid ThemeColor: $color. Using named color."
                try {
                    $window.Background = [System.Windows.Media.Brushes]::$color
                } catch {
                    Write-DebugMessage "Named color also invalid. Using default."
                }
            }
        }
        
        # Set BoxColor for appropriate controls (like GroupBox, TextBox, etc.)
        if ($global:Config["WPFGUI"].Contains("BoxColor")) {
            try {
                $boxColor = $global:Config["WPFGUI"]["BoxColor"]
                $boxBrush = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.ColorConverter]::ConvertFromString($boxColor))
                
                # Function to recursively find and set background for appropriate controls
                function Set-BoxBackground {
                    param($parent)
                    
                    # Process all child elements recursively
                    for ($i = 0; $i -lt [System.Windows.Media.VisualTreeHelper]::GetChildrenCount($parent); $i++) {
                        $child = [System.Windows.Media.VisualTreeHelper]::GetChild($parent, $i)
                        
                        # Apply to specific control types
                        if ($child -is [System.Windows.Controls.GroupBox] -or 
                            $child -is [System.Windows.Controls.TextBox] -or 
                            $child -is [System.Windows.Controls.ComboBox] -or
                            $child -is [System.Windows.Controls.ListBox]) {
                            $child.Background = $boxBrush
                        }
                        
                        # Recursively process children
                        if ([System.Windows.Media.VisualTreeHelper]::GetChildrenCount($child) -gt 0) {
                            Set-BoxBackground -parent $child
                        }
                    }
                }
                
                # Apply when window is loaded to ensure visual tree is constructed
                $window.Add_Loaded({
                    Set-BoxBackground -parent $window
                })
                
            } catch {
                Write-DebugMessage "Error applying BoxColor: $($_.Exception.Message)"
            }
        }
        
        # Set RahmenColor (border color) for appropriate controls
        if ($global:Config["WPFGUI"].Contains("RahmenColor")) {
            try {
                $rahmenColor = $global:Config["WPFGUI"]["RahmenColor"]
                $rahmenBrush = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.ColorConverter]::ConvertFromString($rahmenColor))
                
                # Function to recursively find and set border color for appropriate controls
                function Set-BorderColor {
                    param($parent)
                    
                    # Process all child elements recursively
                    for ($i = 0; $i -lt [System.Windows.Media.VisualTreeHelper]::GetChildrenCount($parent); $i++) {
                        $child = [System.Windows.Media.VisualTreeHelper]::GetChild($parent, $i)
                        
                        # Apply to specific control types with BorderBrush property
                        if ($child -is [System.Windows.Controls.Border] -or 
                            $child -is [System.Windows.Controls.GroupBox] -or 
                            $child -is [System.Windows.Controls.TextBox] -or 
                            $child -is [System.Windows.Controls.ComboBox]) {
                            $child.BorderBrush = $rahmenBrush
                        }
                        
                        # Recursively process children
                        if ([System.Windows.Media.VisualTreeHelper]::GetChildrenCount($child) -gt 0) {
                            Set-BorderColor -parent $child
                        }
                    }
                }
                
                # Apply when window is loaded to ensure visual tree is constructed
                $window.Add_Loaded({
                    Set-BorderColor -parent $window
                })
                
            } catch {
                Write-DebugMessage "Error applying RahmenColor: $($_.Exception.Message)"
            }
        }
        
        # Set font properties
        if ($global:Config["WPFGUI"].Contains("FontFamily")) {
            $window.FontFamily = New-Object System.Windows.Media.FontFamily($global:Config["WPFGUI"]["FontFamily"])
        }
        
        if ($global:Config["WPFGUI"].Contains("FontSize")) {
            try {
                $window.FontSize = [double]$global:Config["WPFGUI"]["FontSize"]
            } catch {
                Write-DebugMessage "Invalid FontSize value"
            }
        }
        
        # Set background image if specified
        if ($global:Config["WPFGUI"].Contains("BackgroundImage")) {
            try {
                $bgImagePath = $global:Config["WPFGUI"]["BackgroundImage"]
                if (Test-Path $bgImagePath) {
                    $bgImage = New-Object System.Windows.Media.Imaging.BitmapImage
                    $bgImage.BeginInit()
                    $bgImage.UriSource = New-Object System.Uri($bgImagePath, [System.UriKind]::Absolute)
                    $bgImage.EndInit()
                    $window.Background = New-Object System.Windows.Media.ImageBrush($bgImage)
                }
            } catch {
                Write-DebugMessage "Error setting background image: $($_.Exception.Message)"
            }
        }
        
        # Set header logo
        $picLogo = $window.FindName("picLogo")
        if ($picLogo -and $global:Config["WPFGUI"].Contains("HeaderLogo")) {
            try {
                $logoPath = $global:Config["WPFGUI"]["HeaderLogo"]
                if (Test-Path $logoPath) {
                    $logo = New-Object System.Windows.Media.Imaging.BitmapImage
                    $logo.BeginInit()
                    $logo.UriSource = New-Object System.Uri($logoPath, [System.UriKind]::Absolute) 
                    $logo.EndInit()
                    $picLogo.Source = $logo
                    
                    # Add click event if URL is specified
                    if ($global:Config["WPFGUI"].Contains("HeaderLogoURL")) {
                        $picLogo.Cursor = [System.Windows.Input.Cursors]::Hand
                        $picLogo.Add_MouseLeftButtonUp({
                            $url = $global:Config["WPFGUI"]["HeaderLogoURL"]
                            if (-not [string]::IsNullOrEmpty($url)) {
                                Start-Process $url
                            }
                        })
                    }
                }
            } catch {
                Write-DebugMessage "Error setting header logo: $($_.Exception.Message)"
            }
        }
        
        # Set footer website hyperlink
        $linkLabel = $window.FindName("linkLabel")
        if ($linkLabel -and $global:Config["WPFGUI"].Contains("FooterWebseite")) {
            try {
                $websiteUrl = $global:Config["WPFGUI"]["FooterWebseite"]
                $linkLabel.Inlines.Clear()
                
                # Create a proper hyperlink with text
                $hyperlink = New-Object System.Windows.Documents.Hyperlink
                $hyperlink.NavigateUri = New-Object System.Uri($websiteUrl)
                $hyperlink.Inlines.Add($websiteUrl)
                $hyperlink.Add_RequestNavigate({
                    param($sender, $e)
                    Start-Process $e.Uri.AbsoluteUri
                    $e.Handled = $true
                })
                
                $linkLabel.Inlines.Add($hyperlink)
            } catch {
                Write-DebugMessage "Error setting footer website link: $($_.Exception.Message)"
            }
        }
        
        # Set footer info text
        $footerInfo = $window.FindName("footerInfo")
        if ($footerInfo -and $global:Config["WPFGUI"].Contains("GUI_ExtraText")) {
            $footerInfo.Text = $global:Config["WPFGUI"]["GUI_ExtraText"]
        }
    }
    
    Write-DebugMessage "Branding settings applied successfully"
} catch {
    Write-DebugMessage "Error applying branding settings: $($_.Exception.Message)"
    Write-DebugMessage "Stack trace: $($_.ScriptStackTrace)"
}

# Shows main window and handles any GUI initialization errors
try {
    $result = $window.ShowDialog()
    Write-DebugMessage "GUI started successfully, result: $result"
} catch {
    Write-DebugMessage "ERROR: GUI could not be started!"
    Write-DebugMessage "Error message: $($_.Exception.Message)"
    Write-DebugMessage "Error details: $($_.InvocationInfo.PositionMessage)"
    exit 1
}
#endregion

Write-DebugMessage "Main GUI execution completed."