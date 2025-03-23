# [Region 00 | SCRIPT INITIALIZATION]
# Sets up error handling and basic script environment
#requires -Version 7.0

$ErrorActionPreference = "Stop"
trap {
    $errorMessage = "ERROR: Unhandled error occurred! " +
                    "Error message: $($_.Exception.Message); " +
                    "Position: $($_.InvocationInfo.PositionMessage); " +
                    "StackTrace: $($_.ScriptStackTrace)"
    if ($null -ne $global:Config -and $null -ne $global:Config.Logging -and $global:Config.Logging.DebugMode -eq "1") {
        Write-Error $errorMessage
    } else {
        Write-Error "A critical error has occurred. Please check the log file."
    }
    exit 1
}

#region [Region 01 | ADMIN AND VERSION CHECK]
# Verifies administrator rights and PowerShell version before proceeding
function Test-Admin {
    $currentUser = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    return $currentUser.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Start-AsAdmin {
    $scriptPath = $MyInvocation.MyCommand.Definition
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = "pwsh.exe"
    $psi.Arguments = "-ExecutionPolicy Bypass -File `"$scriptPath`"" 
    $psi.Verb = "runas" # This triggers the UAC prompt
    try {
        [System.Diagnostics.Process]::Start($psi) | Out-Null
    }
    catch {
        Write-Warning "Failed to restart as administrator: $($_.Exception.Message)"
        return $false
    }
    return $true
}

# Check for admin privileges
if (-not (Test-Admin)) {
    # If not running as admin, show a dialog to restart with elevated permissions
    Add-Type -AssemblyName PresentationFramework
    $result = [System.Windows.MessageBox]::Show(
        "This script requires administrator privileges to run properly.`n`nDo you want to restart with elevated permissions?",
        "Administrator Rights Required",
        "YesNo",
        "Warning"
    )
    
    if ($result -eq "Yes") {
        if (Start-AsAdmin) {
            # Exit this instance as we've started a new elevated instance
            exit
        }
    }
    else {
        Write-Warning "The script will continue without administrator privileges. Some functions may not work properly."
    }
}

# Check PowerShell version
if ($PSVersionTable.PSVersion.Major -lt 7) {
    Add-Type -AssemblyName PresentationFramework
    [System.Windows.MessageBox]::Show(
        "This script requires PowerShell 7 or higher.`n`nCurrent version: $($PSVersionTable.PSVersion)",
        "Version Error",
        "OK",
        "Error"
    )
    exit
}

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
#endregion

#region [Region 02 | INI FILE LOADING]
# Loads configuration data from the external INI file
function Get-IniContent {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Path
    )
    
    # [2.1 | VALIDATE INI FILE]
    # Checks if the INI file exists and is accessible
    if (-not (Test-Path $Path)) {
        Throw "INI file not found: $Path"
    }
    
    $ini = [ordered]@{}
    $currentSection = "Global"
    $ini[$currentSection] = [ordered]@{}
    
    try {
        $lines = Get-Content -Path $Path -ErrorAction Stop
    } catch {
        Throw "Error reading INI file: $($_.Exception.Message)"
    }
    
    # [2.2 | PARSE INI CONTENT]
    # Processes the INI file line by line to extract sections and key-value pairs
    foreach ($line in $lines) {
        $line = $line.Trim()
        # Skip empty lines and comment lines (starting with ; or #)
        if ([string]::IsNullOrWhiteSpace($line) -or $line -match '^\s*[;#]') { continue }
    
        if ($line -match '^\[(.+)\]$') {
            $currentSection = $matches[1].Trim()
            if (-not $ini.Contains($currentSection)) {
                $ini[$currentSection] = [ordered]@{}
            }
        } elseif ($line -match '^(.*?)=(.*)$') {
            $key = $matches[1].Trim()
            $value = $matches[2].Trim()
            $ini[$currentSection][$key] = $value
        }
    }
    return $ini
}

$INIPath = Join-Path $ScriptDir "easyONB.ini"
try {
    $global:Config = Get-IniContent -Path $INIPath
}
catch {
    Write-Host "Error loading INI file: $_"
    exit
}
#endregion

#region [Region 03 | FUNCTION DEFINITIONS]
# Contains all helper and core functionality functions

#region [Region 03.1 | LOGGING FUNCTIONS]
# Defines logging capabilities for different message levels
function Write-LogMessage {
    # [03.1.1 - Primary logging wrapper for consistent message formatting]
    param(
        [string]$message,
        [string]$logLevel = "INFO"
    )
    Write-Log -message $message -logLevel $logLevel
}

function Write-DebugMessage {
    # [03.1.2 - Debug-specific logging with conditional execution]
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true, ValueFromPipeline=$true)]
        [AllowEmptyString()]
        [string]$Message
    )
    process {
        # Only log debug messages if DebugMode is enabled in config
        if ($null -ne $global:Config -and 
            $null -ne $global:Config.Logging -and 
            $global:Config.Logging.DebugMode -eq "1") {
            
            # Call Write-Log without capturing output to avoid pipeline return
            Write-Log -Message $Message -LogLevel "DEBUG"
            
            # Also output to console for immediate feedback during debugging
            Write-Host "[DEBUG] $Message" -ForegroundColor Cyan
        }
        
        # No return value to avoid unwanted pipeline output
    }
}
#endregion

# Log level (default is "INFO"): "WARN", "ERROR", "DEBUG".

#region [Region 03.2 | LOG FILE WRITER] 
# Core logging function that writes messages to log files
function Write-Log {
    # [03.2.1 - Low-level file logging implementation]
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true, ValueFromPipeline=$true)]
        [AllowEmptyString()]
        [string]$Message,
        
        [Parameter(Mandatory=$false)]
        [ValidateSet("INFO", "WARN", "ERROR", "DEBUG")]
        [string]$LogLevel = "INFO"
    )
    process {
        # log path determination using null-conditional and coalescing operators
        $logFilePath = $global:Config?.Logging?.ExtraFile ?? 
                      $global:Config?.Logging?.LogFile ?? 
                      (Join-Path $ScriptDir "Logs")
        
        # Ensure the log directory exists
        if (-not (Test-Path $logFilePath)) {
            try {
                # Use -Force to create parent directories if needed
                [void](New-Item -ItemType Directory -Path $logFilePath -Force -ErrorAction Stop)
                Write-Host "Created log directory: $logFilePath"
            } catch {
                # Handle directory creation errors gracefully
                Write-Warning "Error creating log directory: $($_.Exception.Message)"
                # Try writing to a fallback location if possible
                $logFilePath = $env:TEMP
            }
        }

        # Define log file paths
        $logFile = Join-Path -Path $logFilePath -ChildPath "easyOnboarding.log"
        $errorLogFile = Join-Path -Path $logFilePath -ChildPath "easyOnboarding_error.log"
        
        # Generate timestamp and log entry
        $timeStamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        $logEntry = "$timeStamp [$LogLevel] $Message"

        try {
            Add-Content -Path $logFile -Value $logEntry -ErrorAction Stop
        } catch {
            # Log additional error if writing to log file fails
            try {
                Add-Content -Path $errorLogFile -Value "$timeStamp [ERROR] Error writing to log file: $($_.Exception.Message)" -ErrorAction SilentlyContinue
            } catch {
                # Last resort: output to console
                Write-Warning "Critical error: Cannot write to any log file: $($_.Exception.Message)"
            }
        }

        # Output to console based on DebugMode setting
        if ($null -ne $global:Config -and $null -ne $global:Config.Logging) {
            # When DebugMode=1, output all messages regardless of log level
            if ($global:Config.Logging.DebugMode -eq "1") {
                # Use color coding based on log level
                switch ($LogLevel) {
                    "DEBUG" { Write-Host "$logEntry" -ForegroundColor Cyan }
                    "INFO"  { Write-Host "$logEntry" -ForegroundColor White }
                    "WARN"  { Write-Host "$logEntry" -ForegroundColor Yellow }
                    "ERROR" { Write-Host "$logEntry" -ForegroundColor Red }
                    default { Write-Host "$logEntry" }
                }
            }
            # When DebugMode=0, no console output
        }
        
        # No return value to avoid unwanted pipeline output
    }
}
#endregion

Write-DebugMessage "INI, LOG, Module, etc. regions initialized."

#region [Region 04 | ACTIVE DIRECTORY MODULE]
# Loads the Active Directory PowerShell module required for user management
Write-DebugMessage "Loading Active Directory module."
try {
    Write-DebugMessage "Loading AD module"
    Import-Module ActiveDirectory -SkipEditionCheck -ErrorAction Stop
    Write-DebugMessage "Active Directory module loaded successfully."
} catch {
    Write-Log -Message "ActiveDirectory module could not be loaded: $($_.Exception.Message)" -LogLevel "ERROR"
    Throw "Critical error: ActiveDirectory module missing!"
}
#endregion

#region [Region 05 | WPF ASSEMBLIES]
# Loads required WPF assemblies for the GUI interface
Write-DebugMessage "Loading WPF assemblies."
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Windows.Forms
Write-DebugMessage "WPF assemblies loaded successfully."
#endregion

Write-DebugMessage "Determining XAML file path."

#region [Region 06 | XAML LOADING]
# Determines XAML file path and loads the GUI definition
Write-DebugMessage "Loading XAML file: $xamlPath"
$xamlPath = Join-Path $ScriptDir "MainGUI.xaml"
if (-not (Test-Path $xamlPath)) {
    Write-Error "XAML file not found: $xamlPath"
    exit
}
try {
    [xml]$xaml = Get-Content -Path $xamlPath
    $reader = New-Object System.Xml.XmlNodeReader $xaml
    $window = [Windows.Markup.XamlReader]::Load($reader)
    if (-not $window) {
        Throw "XAML could not be loaded."
    }
    Write-DebugMessage "XAML file loaded successfully."
    if ($global:Config.Logging.DebugMode -eq "1") {
        Write-Log "XAML file successfully loaded." "DEBUG"
    }
}
catch {
    Write-Error "Error loading XAML file. Please check the file content. $_"
    exit
}
#endregion

Write-DebugMessage "XAML file loaded"
Write-DebugMessage "Defining UPN generation function."

#region [Region 08 | UPN GENERATION]
# Creates user principal names based on templates and user data
Write-DebugMessage "Generating UPN for user."
function New-UPN {
    param(
        [pscustomobject]$userData,
        [hashtable]$Config
    )

    # Check if FirstName and LastName are provided
    if ([string]::IsNullOrWhiteSpace($userData.FirstName) -or [string]::IsNullOrWhiteSpace($userData.LastName)) {
        Throw "Error: FirstName and LastName must be set!"
    }

    # 1) Generate SamAccountName (first letter of first name + entire last name, all lowercase)
    $SamAccountName = ''
    if ($userData.FirstName.Length -gt 0) {
        $SamAccountName = ($userData.FirstName.Substring(0,1) + $userData.LastName).ToLower()
    } else {
        $SamAccountName = $userData.LastName.ToLower()
    }
    Write-DebugMessage "SamAccountName= $SamAccountName"

    # 2) If a manual UPN was entered, use it immediately
    if (-not [string]::IsNullOrWhiteSpace($userData.UPNEntered)) {
        return @{
            SamAccountName = $SamAccountName
            UPN            = $userData.UPNEntered
            CompanySection = "Company"
        }
    }

    # 3) Determine the Company section (default: "Company")
    $companySection = "Company"  # Default value
    
    # Optimized check for CompanySection with simplified null value check
    if ($null -ne $userData.CompanySection) {
        # Case 1: CompanySection is directly a string
        if ($userData.CompanySection -is [string] -and -not [string]::IsNullOrWhiteSpace($userData.CompanySection)) {
            $companySection = $userData.CompanySection
        }
        # Case 2: CompanySection is an object with Section property
        elseif ($userData.CompanySection -is [PSObject] -and 
                $null -ne ($userData.CompanySection.PSObject.Properties.Match('Section') | Select-Object -First 1) -and
                -not [string]::IsNullOrWhiteSpace($userData.CompanySection.Section)) {
            $companySection = $userData.CompanySection.Section
        }
    }
    Write-DebugMessage "Using Company section: '$companySection'"

    # 4) Check if the desired section exists in the INI with improved error handling
    if ($null -eq $Config) {
        Throw "Error: Config object is NULL! No configuration available."
    }
    
    if (-not $Config.Contains($companySection)) {
        Write-DebugMessage "Section '$companySection' not found in Config, using 'Company'"
        $companySection = "Company" # Fallback to default value
        if (-not $Config.Contains($companySection)) {
            Throw "Error: Neither the requested section nor the default section 'Company' exists in the INI!"
        }
    }
    $companyData = $Config[$companySection]
    $suffix = ($companySection -replace "\D","")

    # 5) Determine domain key and read domain - with improved null checks
    $domainKey = "CompanyActiveDirectoryDomain$suffix"
    $fallbackDomainKey = "CompanyActiveDirectoryDomain"
    
    # Check for the specific key first, then fallback - with more precise checks
    if ($null -ne $companyData -and 
        $companyData.Contains($domainKey) -and 
        -not [string]::IsNullOrWhiteSpace($companyData[$domainKey])) {
        $adDomain = "@" + $companyData[$domainKey].Trim()
    }
    elseif ($null -ne $companyData -and 
            $companyData.Contains($fallbackDomainKey) -and 
            -not [string]::IsNullOrWhiteSpace($companyData[$fallbackDomainKey])) {
        $adDomain = "@" + $companyData[$fallbackDomainKey].Trim()
        Write-DebugMessage "Using fallback domain: $adDomain"
    }
    else {
        # Extended error message with more details about the problem
        $availableKeys = $null -ne $companyData ? ($companyData.Keys -join ", ") : "No keys available"
        Throw "Error: Domain information missing in the INI! Neither '$domainKey' nor '$fallbackDomainKey' found or valid. Available keys: $availableKeys"
    }

    # 6) Always use UPN template from INI (no GUI dropdown) - with optimized NULL check
    $upnTemplate = "FIRSTNAME.LASTNAME" # Safe default value
    
    # PowerShell 7 ?:-operator for NULL check and improved readability
    $displayNameTemplates = $null -ne $Config -and $Config.Contains("DisplayNameUPNTemplates") ? 
        $Config.DisplayNameUPNTemplates : $null
    
    if ($null -ne $displayNameTemplates -and 
        $displayNameTemplates.Contains("DefaultDisplayNameFormat") -and 
        -not [string]::IsNullOrWhiteSpace($displayNameTemplates["DefaultDisplayNameFormat"])) {
        $upnTemplate = $displayNameTemplates["DefaultDisplayNameFormat"].ToUpper()
        Write-DebugMessage "UPN template loaded from INI: $upnTemplate"
    }
    else {
        Write-DebugMessage "No valid UPN template found in INI, using default '$upnTemplate'"
    }
    
    Write-DebugMessage "UPN template (from INI): $upnTemplate"

    # 7) Generate UPN based on template - with safe substring processing and null checks
    # Make sure FirstName and LastName are not null
    $firstName = if ($null -eq $userData.FirstName) { "" } else { $userData.FirstName }
    $lastName = if ($null -eq $userData.LastName) { "" } else { $userData.LastName }
    
    $UPN = ""
    switch ($upnTemplate) {
        "FIRSTNAME.LASTNAME"    { $UPN = "$firstName.$lastName".ToLower() + $adDomain }
        "LASTNAME.FIRSTNAME"    { $UPN = "$lastName.$firstName".ToLower() + $adDomain }
        "FIRSTINITIAL.LASTNAME" { 
            if (-not [string]::IsNullOrEmpty($firstName) -and $firstName.Length -gt 0) {
                $UPN = "$($firstName.Substring(0,1)).$lastName".ToLower() + $adDomain
            } else {
                $UPN = "x.$lastName".ToLower() + $adDomain
            }
        }
        "FIRSTNAME.LASTINITIAL" { 
            if (-not [string]::IsNullOrWhiteSpace($lastName)) {
                $UPN = "$firstName.$($lastName.Substring(0,1))".ToLower() + $adDomain
            } else {
                $UPN = "$firstName.x".ToLower() + $adDomain
            }
        }
        "FIRSTNAME_LASTNAME"    { $UPN = "$firstName_$lastName".ToLower() + $adDomain }
        "LASTNAME_FIRSTNAME"    { $UPN = "$lastName_$firstName".ToLower() + $adDomain }
        "FIRSTINITIAL_LASTNAME" { 
            if (-not [string]::IsNullOrWhiteSpace($firstName)) {
                $UPN = "$($firstName.Substring(0,1))_$lastName".ToLower() + $adDomain
            } else {
                $UPN = "x_$lastName".ToLower() + $adDomain
            }
        }
        "FIRSTNAME_LASTINITIAL" { 
            if (-not [string]::IsNullOrWhiteSpace($lastName)) {
                $UPN = "$firstName_$($lastName.Substring(0,1))".ToLower() + $adDomain
            } else {
                $UPN = "$firstName_x".ToLower() + $adDomain
            }
        }
        "FIRSTNAMELASTNAME"     { $UPN = "$firstName$lastName".ToLower() + $adDomain }
        "LASTNAMEFIRSTNAME"     { $UPN = "$lastName$firstName".ToLower() + $adDomain }
        "FIRSTINITIALLASTNAME"  { 
            if (-not [string]::IsNullOrWhiteSpace($firstName)) {
                $UPN = "$($firstName.Substring(0,1))$lastName".ToLower() + $adDomain
            } else {
                $UPN = "x$lastName".ToLower() + $adDomain
            }
        }
        "FIRSTNAMELASTINITIAL"  { 
            if (-not [string]::IsNullOrWhiteSpace($lastName)) {
                $UPN = "$firstName$($lastName.Substring(0,1))".ToLower() + $adDomain
            } else {
                $UPN = "$firstName x".ToLower() + $adDomain
            }
        }
        default                 { $UPN = "$firstName.$lastName".ToLower() + $adDomain }
    }

    # 8) Clean UPN of umlauts and special characters
    if ($UPN) {
        $UPN = $UPN -replace "ä", "ae" -replace "ö", "oe" -replace "ü", "ue" -replace "ß", "ss" -replace "Ä", "ae" -replace "Ö", "oe" -replace "Ü", "ue"
        $UPN = $UPN -replace "[^a-zA-Z0-9._@-]", ""
    } else {
        # Fallback for empty UPN
        $UPN = "user@" + $companyData[$domainKey ?? $fallbackDomainKey].Trim()
    }

    # 9) Return result
    return @{
        SamAccountName = $SamAccountName
        UPN            = $UPN
        CompanySection = $companySection
    }
}
#endregion

#region [Region 09 | AD USER CREATION]
# Creates Active Directory user accounts based on input data
function New-ADUserAccount {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        [pscustomobject]$UserData,
        [Parameter(Mandatory=$true)]
        [hashtable]$Config
    )
    
    try {
        # 1) Generate UPN data
        $upnData = New-UPN -userData $UserData -Config $Config
        $samAccountName = $upnData.SamAccountName
        $userPrincipalName = $upnData.UPN
        $companySection = $upnData.CompanySection
        
        Write-DebugMessage "Generated UPN: $userPrincipalName"
        Write-DebugMessage "Using Company section: $companySection"
        
        # 2) Check if all required information is available
        if (-not $Config.Contains($companySection)) {
            throw "No configuration found for Company section '$companySection'."
        }
        
        $suffix = ($companySection -replace "\D","") 
        $adOUKey = "CompanyActiveDirectoryOU$suffix"
        $adPath = $Config[$companySection][$adOUKey]
        
        if (-not $adPath) {
            # Fall back to default OU if specified
            if ($Config.Contains("ADUserDefaults") -and $Config.ADUserDefaults.Contains("DefaultOU")) {
                $adPath = $Config.ADUserDefaults.DefaultOU
                Write-DebugMessage "Using fallback DefaultOU: $adPath"
            }
            else {
                throw "No AD path (OU) found for Company section '$companySection'."
            }
        }
        
        # 3) Generate DisplayName according to configured template (from INI)
        $displayNameFormat = Get-DisplayNameFormat -IniPath $INIPath -SelectedTemplate $UserData.DisplayNameTemplate

        # Default to "LastName, FirstName" if no valid format is found and nothing was entered or selected
        if ([string]::IsNullOrWhiteSpace($displayNameFormat) -and [string]::IsNullOrWhiteSpace($UserData.DisplayNameTemplate)) {
            $displayNameFormat = "{last}, {first}"
        }

        $displayName = Format-DisplayName -Format $displayNameFormat -FirstName $UserData.FirstName -LastName $UserData.LastName
        
        # 4) Check if user already exists
        $userExists = $false
        try {
            $existingUser = Get-ADUser -Identity $samAccountName
            if ($existingUser) {
                $userExists = $true
                Write-Log "User $samAccountName already exists." -LogLevel "WARN"
            }
        } catch {
            # User doesn't exist - that's good
            $userExists = $false
        }
        
        if ($userExists) {
            return @{
                Success = $false
                Message = "User $samAccountName already exists."
                SamAccountName = $samAccountName
            }
        }
        
        # 5) Generate secure password if not set
        if ([string]::IsNullOrWhiteSpace($UserData.Password)) {
            # Password length from Config or default 12 characters
            $passwordLength = 12
            if ($Config.Contains("PasswordFixGenerate") -and $Config.PasswordFixGenerate.Contains("DefaultPasswordLength")) {
                $tempLength = [int]::TryParse($Config.PasswordFixGenerate.DefaultPasswordLength, [ref]$passwordLength)
                if (-not $tempLength) { $passwordLength = 12 }
            }
            
            # Character pool for password
            $charPool = "abcdefghijkmnopqrstuvwxyzABCDEFGHJKLMNPQRSTUVWXYZ23456789!@#$%^&*_=+-"
            $securePassword = ""
            $random = New-Object System.Random
            
            # At least 1 uppercase, 1 lowercase, 1 number and 1 special character
            $securePassword += $charPool.Substring($random.Next(0, 25), 1)  # Lowercase
            $securePassword += $charPool.Substring($random.Next(26, 50), 1) # Uppercase
            $securePassword += $charPool.Substring($random.Next(51, 59), 1) # Number
            $securePassword += $charPool.Substring($random.Next(60, $charPool.Length-1), 1) # Special character
            
            # Fill up to desired length
            for ($i = 4; $i -lt $passwordLength; $i++) {
                $securePassword += $charPool.Substring($random.Next(0, $charPool.Length-1), 1)
            }
            
            # Randomize character order
            $securePasswordArray = $securePassword.ToCharArray()
            $randomizedPassword = ""
            for ($i = $securePasswordArray.Count; $i -gt 0; $i--) {
                $randomPosition = $random.Next(0, $i)
                $randomizedPassword += $securePasswordArray[$randomPosition]
                $securePasswordArray = $securePasswordArray[0..($randomPosition-1)] + $securePasswordArray[($randomPosition+1)..($securePasswordArray.Count-1)]
            }
            
            $UserData.Password = $randomizedPassword
            Write-DebugMessage "Generated password: $randomizedPassword"
        }
        
        # Email domain suffix handling - get directly from the dropdown control
        $mailSuffix = ""
        if ($global:SelectedMailSuffix) {
            $mailSuffix = $global:SelectedMailSuffix.Trim()
            Write-DebugMessage "Using mail suffix from global variable: '$mailSuffix'"
        } elseif (-not [string]::IsNullOrWhiteSpace($UserData.MailSuffix)) {
            $mailSuffix = $UserData.MailSuffix.Trim()
            Write-DebugMessage "Using mail suffix from UserData: '$mailSuffix'"
        } else {
            $comboBoxMailSuffix = $window.FindName("cmbSuffix")
            if ($comboBoxMailSuffix -and $comboBoxMailSuffix.SelectedValue) {
                $mailSuffix = $comboBoxMailSuffix.SelectedValue.ToString().Trim()
                Write-DebugMessage "Using mail suffix directly from dropdown: '$mailSuffix'"
            } else {
                # Fallback to Company config
                if ($Config[$companySection].Contains("CompanyMailDomain")) {
                    $mailSuffix = $Config[$companySection]["CompanyMailDomain"]
                    Write-DebugMessage "Using fallback mail suffix from [$companySection].CompanyMailDomain: '$mailSuffix'"
                }
            }
        }
        
        # Ensure mail suffix starts with @
        if (-not [string]::IsNullOrWhiteSpace($mailSuffix) -and -not $mailSuffix.StartsWith('@')) {
            $mailSuffix = "@" + $mailSuffix
            Write-DebugMessage "Added @ prefix to mail suffix: '$mailSuffix'"
        }
        
        # 6) Collect AD user parameters with all possible attributes
        $adUserParams = @{
            SamAccountName = $samAccountName
            UserPrincipalName = $userPrincipalName
            Name = $displayName
            DisplayName = $displayName
            GivenName = $UserData.FirstName
            Surname = $UserData.LastName
            Path = $adPath
            AccountPassword = (ConvertTo-SecureString -String $UserData.Password -AsPlainText -Force)
            Enabled = (-not [bool]$UserData.AccountDisabled)
        }
        
        # Add all other attributes from UserData or config if they exist
        
        # Basic attributes
        if (-not [string]::IsNullOrWhiteSpace($UserData.Description)) {
            $adUserParams.Description = $UserData.Description
        }
        
        if (-not [string]::IsNullOrWhiteSpace($UserData.OtherName)) {
            $adUserParams.OtherName = $UserData.OtherName
        }
        
        if (-not [string]::IsNullOrWhiteSpace($UserData.Initials)) {
            $adUserParams.Initials = $UserData.Initials
        }
        
        # Job information
        if (-not [string]::IsNullOrWhiteSpace($UserData.Position)) {
            $adUserParams.Title = $UserData.Position
        }
        
        if (-not [string]::IsNullOrWhiteSpace($UserData.DepartmentField)) {
            $adUserParams.Department = $UserData.DepartmentField
        }
        
        if (-not [string]::IsNullOrWhiteSpace($UserData.Division)) {
            $adUserParams.Division = $UserData.Division
        }
        
        # Company information
        if ($Config[$companySection].Contains("CompanyNameFirma$suffix")) {
            $adUserParams.Company = $Config[$companySection]["CompanyNameFirma$suffix"]
        } elseif (-not [string]::IsNullOrWhiteSpace($UserData.Company)) {
            $adUserParams.Company = $UserData.Company
        }
        
        # Employee identifiers
        if (-not [string]::IsNullOrWhiteSpace($UserData.EmployeeID)) {
            $adUserParams.EmployeeID = $UserData.EmployeeID
        }
        
        if (-not [string]::IsNullOrWhiteSpace($UserData.EmployeeNumber)) {
            $adUserParams.EmployeeNumber = $UserData.EmployeeNumber
        }
        
        # Address information
        if (-not [string]::IsNullOrWhiteSpace($UserData.OfficeRoom)) {
            $adUserParams.Office = $UserData.OfficeRoom
        }
        
        if (-not [string]::IsNullOrWhiteSpace($UserData.POBox)) {
            $adUserParams.POBox = $UserData.POBox
        }
        
        if (-not [string]::IsNullOrWhiteSpace($UserData.State)) {
            $adUserParams.State = $UserData.State
        }
        
        # Street address from user data or company config
        if (-not [string]::IsNullOrWhiteSpace($UserData.StreetAddress)) {
            $adUserParams.StreetAddress = $UserData.StreetAddress
        } elseif ($Config[$companySection].Contains("CompanyStrasse")) {
            $adUserParams.StreetAddress = $Config[$companySection]["CompanyStrasse"]
        }
        
        # Postal code from user data or company config
        if (-not [string]::IsNullOrWhiteSpace($UserData.PostalCode)) {
            $adUserParams.PostalCode = $UserData.PostalCode
        } elseif ($Config[$companySection].Contains("CompanyPLZ")) {
            $adUserParams.PostalCode = $Config[$companySection]["CompanyPLZ"]
        }
        
        # City from user data or company config
        if (-not [string]::IsNullOrWhiteSpace($UserData.City)) {
            $adUserParams.City = $UserData.City
        } elseif ($Config[$companySection].Contains("CompanyOrt")) {
            $adUserParams.City = $Config[$companySection]["CompanyOrt"]
        }
        
        # Country from user data or company config
        if (-not [string]::IsNullOrWhiteSpace($UserData.Country)) {
            $adUserParams.Country = $UserData.Country
        } elseif ($Config[$companySection].Contains("CompanyCountry$suffix")) {
            $adUserParams.Country = $Config[$companySection]["CompanyCountry$suffix"]
        }
        
        # Contact information
        if (-not [string]::IsNullOrWhiteSpace($UserData.EmailAddress)) {
            # Apply the mail suffix from the dropdown if it exists
            $emailAddress = $UserData.EmailAddress
            if (-not $emailAddress.Contains('@') -and -not [string]::IsNullOrWhiteSpace($mailSuffix)) {
                $emailAddress = "$emailAddress$mailSuffix"
            }
            $adUserParams.EmailAddress = $emailAddress
            Write-DebugMessage "Setting EmailAddress to: $emailAddress"
        }
        
        if (-not [string]::IsNullOrWhiteSpace($UserData.PhoneNumber)) {
            $adUserParams.OfficePhone = $UserData.PhoneNumber
        }
        
        if (-not [string]::IsNullOrWhiteSpace($UserData.MobileNumber)) {
            $adUserParams.MobilePhone = $UserData.MobileNumber
        }
        
        if (-not [string]::IsNullOrWhiteSpace($UserData.Fax)) {
            $adUserParams.Fax = $UserData.Fax
        }
        
        if (-not [string]::IsNullOrWhiteSpace($UserData.Pager)) {
            $adUserParams.Pager = $UserData.Pager
        }
        
        if (-not [string]::IsNullOrWhiteSpace($UserData.IPPhone)) {
            $adUserParams.IPPhone = $UserData.IPPhone
        } elseif ($Config[$companySection].Contains("CompanyTelefon")) {
            $adUserParams.IPPhone = $Config[$companySection]["CompanyTelefon"]
        }
        
        # File system paths from user data or defaults
        if (-not [string]::IsNullOrWhiteSpace($UserData.HomeDirectory)) {
            $adUserParams.HomeDirectory = $UserData.HomeDirectory
            if (-not [string]::IsNullOrWhiteSpace($UserData.HomeDrive)) {
                $adUserParams.HomeDrive = $UserData.HomeDrive
            }
        } elseif ($Config.ADUserDefaults.Contains("HomeDirectory")) {
            $homePath = $Config.ADUserDefaults.HomeDirectory
            # Replace placeholder %username% with actual SamAccountName
            $homePath = $homePath -replace "%username%", $samAccountName
            $adUserParams.HomeDirectory = $homePath
            
            if ($Config.ADUserDefaults.Contains("HomeDrive")) {
                $adUserParams.HomeDrive = $Config.ADUserDefaults.HomeDrive
            }
        }
        
        if (-not [string]::IsNullOrWhiteSpace($UserData.ProfilePath)) {
            $adUserParams.ProfilePath = $UserData.ProfilePath
        } elseif ($Config.ADUserDefaults.Contains("ProfilePath")) {
            $profilePath = $Config.ADUserDefaults.ProfilePath
            # Replace placeholder %username% with actual SamAccountName
            $profilePath = $profilePath -replace "%username%", $samAccountName
            $adUserParams.ProfilePath = $profilePath
        }
        
        if (-not [string]::IsNullOrWhiteSpace($UserData.ScriptPath)) {
            $adUserParams.ScriptPath = $UserData.ScriptPath
        } elseif ($Config.ADUserDefaults.Contains("LogonScript")) {
            $adUserParams.ScriptPath = $Config.ADUserDefaults.LogonScript
        }
        
        # Account expiration
        if (-not [string]::IsNullOrWhiteSpace($UserData.Ablaufdatum)) {
            try {
                $expirationDate = [DateTime]::Parse($UserData.Ablaufdatum)
                $adUserParams.AccountExpirationDate = $expirationDate
            }
            catch {
                Write-DebugMessage "Invalid expiration date format: $($UserData.Ablaufdatum). Error: $($_.Exception.Message)"
            }
        }
        
        # Add other attributes with mail information
        $otherAttributes = @{}
        
        # Create the 'mail' attribute with the selected mail suffix
        if (-not [string]::IsNullOrWhiteSpace($UserData.EmailAddress)) {
            $email = $UserData.EmailAddress
            if (-not $email.Contains('@') -and -not [string]::IsNullOrWhiteSpace($mailSuffix)) {
                $email = "$email$mailSuffix"
            }
            $otherAttributes["mail"] = $email
            Write-DebugMessage "Setting mail attribute to: $email"
            
            # If setProxyMailAddress is enabled, add proxyAddresses
            if ($UserData.setProxyMailAddress) {
                # Initialize proxyAddresses array
                $proxyAddresses = @()
                
                # Add the primary SMTP address
                $proxyAddresses += "SMTP:$email"
                
                # Add MS365 address if configured
                if ($Config[$companySection].Contains("CompanyMS365Domain") -and 
                    -not [string]::IsNullOrWhiteSpace($Config[$companySection]["CompanyMS365Domain"])) {
                    
                    $ms365Domain = $Config[$companySection]["CompanyMS365Domain"]
                    if (-not $ms365Domain.StartsWith('@')) {
                        $ms365Domain = "@$ms365Domain"
                    }
                    
                    # Extract username part from email
                    $username = ""
                    if ($email -match '@') {
                        $username = $email.Substring(0, $email.IndexOf('@'))
                    } else {
                        $username = $UserData.EmailAddress
                    }
                    
                    $ms365Address = "$username$ms365Domain"
                    $proxyAddresses += "smtp:$ms365Address"
                    Write-DebugMessage "Added secondary MS365 proxy address: smtp:$ms365Address"
                }
                
                if ($proxyAddresses.Count -gt 0) {
                    $otherAttributes["proxyAddresses"] = $proxyAddresses
                    Write-DebugMessage "Setting proxyAddresses attribute with $($proxyAddresses.Count) addresses"
                }
            }
        }
        
        # Add otherAttributes to the user parameters if we have any
        if ($otherAttributes.Count -gt 0) {
            $adUserParams["OtherAttributes"] = $otherAttributes
        }
        
        # 7) Create the user account
        Write-DebugMessage "Creating AD user: $samAccountName with attributes: $(($adUserParams.Keys | Sort-Object) -join ', ')"
        try {
            $newUser = New-ADUser @adUserParams -PassThru
            if (-not $newUser) {
                throw "User creation appeared to succeed but no user object was returned."
            }
        }
        catch {
            throw "Error creating user: $($_.Exception.Message)"
        }
        
        # 8) Set manager if specified
        if (-not [string]::IsNullOrWhiteSpace($UserData.Manager)) {
            try {
                # Find manager by SamAccountName or DisplayName
                $manager = Get-ADUser -Filter {(SamAccountName -eq $UserData.Manager) -or (DisplayName -eq $UserData.Manager)}
                if ($manager) {
                    Set-ADUser -Identity $samAccountName -Manager $manager
                    Write-DebugMessage "Manager set to: $($manager.DistinguishedName)"
                }
            }
            catch {
                Write-DebugMessage "Failed to set manager: $($_.Exception.Message)"
            }
        }
        
        # 9) Add the selected mail suffix to the return data for reference
        Write-Log "User $samAccountName was successfully created with mail suffix: $mailSuffix" -LogLevel "INFO"
        return @{
            Success = $true
            Message = "User was successfully created."
            SamAccountName = $samAccountName
            UserPrincipalName = $userPrincipalName
            Password = $UserData.Password
            MailSuffix = $mailSuffix
        }
    }
    catch {
        Write-Log "Error creating user: $($_.Exception.Message)" -LogLevel "ERROR"
        return @{
            Success = $false
            Message = "Error creating user: $($_.Exception.Message)"
            SamAccountName = $samAccountName
        }
    }
}

# Add this to ensure the email suffix dropdown selection is captured
# Run this when the script initializes
$comboBoxMailSuffix = $window.FindName("cmbSuffix")
if ($comboBoxMailSuffix) {
    # Initialize the global variable with the default selection
    if ($comboBoxMailSuffix.SelectedValue) {
        $global:SelectedMailSuffix = $comboBoxMailSuffix.SelectedValue.ToString()
        Write-DebugMessage "Initialized mail suffix to: $global:SelectedMailSuffix"
    }
    
    # Register the selection change event
    $comboBoxMailSuffix.Add_SelectionChanged({
        $global:SelectedMailSuffix = $comboBoxMailSuffix.SelectedValue
        Write-DebugMessage "Mail suffix selection changed to: $global:SelectedMailSuffix"
    })
}

Write-DebugMessage "Registering GUI event handlers."
#endregion

#region [Region 09 | GUI EVENT HANDLER]
# Function to register and handle GUI button click events with error handling
Write-DebugMessage "Registering GUI event handler for $Control."
function Register-GUIEvent {
    param (
        [Parameter(Mandatory=$true)]
        $Control,
        [Parameter(Mandatory=$true)]
        [ScriptBlock]$EventAction,
        [string]$ErrorMessagePrefix = "Error executing event"
    )
    if ($Control) {
        $Control.Add_Click({
            try {
                & $EventAction
            }
            catch {
                [System.Windows.MessageBox]::Show("${ErrorMessagePrefix}: $($_.Exception.Message)", "Error", "OK", "Error")
            }
        })
    }
}
#endregion

Write-DebugMessage "Defining AD command execution function."

#region [Region 10 | AD COMMAND EXECUTION]
# Consolidates error handling for Active Directory commands
Write-DebugMessage "Executing AD command."
function Invoke-ADCommand {
    param (
        [Parameter(Mandatory=$true)]
        [ScriptBlock]$Command,
        
        [Parameter(Mandatory=$true)]
        [string]$ErrorContext
    )
    try {
        & $Command
    }
    catch {
        $errMsg = "Error in Invoke-ADCommand"
        Write-Error $errMsg
        Throw $errMsg
    }
}
#endregion

Write-DebugMessage "Defining configuration value access function."

#region [Region 11 | CONFIGURATION VALUE ACCESS]
# Helper function for safely accessing configuration values
Write-DebugMessage "Accessing configuration value: $Key."
function Get-ConfigValue {
    param (
        [Parameter(Mandatory = $true)]
        [hashtable]$Section,
        
        [Parameter(Mandatory = $true)]
        [string]$Key,
        
        [bool]$Mandatory = $true
    )
    
    if ([string]::IsNullOrWhiteSpace($Key)) {
        Throw "Error: The provided key is null or empty."
    }
    
    if ($Section.Contains($Key)) {
        return $Section[$Key]
    }
    elseif ($Mandatory) {
        Throw "Error: The key is missing in the configuration!"
    }
    return $null
}
#endregion

Write-DebugMessage "Defining template processing functions."

#region [Region 12 | TEMPLATE PROCESSING]
# Functions to replace placeholders in templates with user data
Write-DebugMessage "Resolving template placeholders."
function Resolve-TemplatePlaceholders {
    # [12.1 - Generic placeholder replacement for string templates]
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Template,

        [Parameter(Mandatory = $true)]
        [ValidateNotNull()]
        [pscustomobject]$userData
    )
    # Replaces placeholders {first} and {last} with corresponding values from $userData
    $result = $Template -replace '{first}', $userData.FirstName `
                          -replace '{last}', $userData.LastName
    return $result
}

Write-DebugMessage "Function to replace placeholders - UPN"
# In the UPN part:
if (-not [string]::IsNullOrWhiteSpace($userData.UPNFormat)) {
    # Normalize the template: trim and convert to lowercase
    $upnTemplate = $userData.UPNFormat.Trim().ToLower()
    Write-DebugMessage "Invoke-Onboarding: UPN Format from userData: $upnTemplate"
    if ($upnTemplate -like "*{first}*") {
        # Dynamic replacement of placeholders from the template
        $upnBase = Resolve-TemplatePlaceholders -Template $upnTemplate -userData $userData
        $UPN = "$upnBase$adDomain"
    }
    else {
        # Fixed cases as fallback – add more cases if needed
        switch ($upnTemplate) {
            "firstname.lastname"    { $UPN = "$($userData.FirstName).$($userData.LastName)$adDomain" }
            "f.lastname"            { $UPN = "$($userData.FirstName.Substring(0,1)).$($userData.LastName)$adDomain" }
            "firstnamelastname"     { $UPN = "$($userData.FirstName)$($userData.LastName)$adDomain" }
            "flastname"             { $UPN = "$($userData.FirstName.Substring(0,1))$($userData.LastName)$adDomain" }
            Default                 { $UPN = "$SamAccountName$adDomain" }
        }
    }
}
else {
    Write-DebugMessage "No UPNFormat given, fallback to SamAccountName + domain"
    $UPN = "$SamAccountName$adDomain"
}

function Get-DisplayNameFormat {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$IniPath,

        [Parameter(Mandatory = $true)]
        [string]$SelectedTemplate
    )

    # Direkt Format zurückgeben, wenn es bereits ein Format-String ist (z.B. "{first} {last}")
    if ($SelectedTemplate -match '(\{first\}|\{last\})') {
        Write-DebugMessage "Using direct template format: $SelectedTemplate"
        return $SelectedTemplate
    }

    # Wenn SelectedTemplate nicht wie ein Template-Name aussieht und nicht "Custom" ist,
    # könnte es bereits ein DisplayName-Format sein (z.B. aus der ComboBox)
    if ($SelectedTemplate -notmatch '^Template\d+$' -and $SelectedTemplate -ne "Custom") {
        Write-DebugMessage "Selected template doesn't match expected pattern: $SelectedTemplate"
        # Falls es ein DisplayName ohne Formatierung ist, direkt zurückgeben
        if (-not [string]::IsNullOrWhiteSpace($SelectedTemplate)) {
            return $SelectedTemplate
        }
    }

    # INI-Datei einlesen
    $iniContent = Get-Content -Path $IniPath -Raw

    # Extrahiere die [DisplayNameUPNTemplates] Sektion
    if ($iniContent -match '\[DisplayNameUPNTemplates\](.*?)(?=\[|$)') {
        $displayNameSection = $matches[1]
        Write-DebugMessage "Found DisplayNameUPNTemplates section in INI"

        # Je nach ausgewähltem Template den richtigen Wert zurückgeben
        if ($SelectedTemplate -eq "Template1" -and $displayNameSection -match 'DisplayNameTemplate1=(.+)') {
            $value = $matches[1]
            if ($value -match '\|') {
                $value = $value.Split('|', 2)[1].Trim()
            }
            Write-DebugMessage "Using Template1: $value"
            return $value
        }
        elseif ($SelectedTemplate -eq "Template2" -and $displayNameSection -match 'DisplayNameTemplate2=(.+)') {
            $value = $matches[1]
            if ($value -match '\|') {
                $value = $value.Split('|', 2)[1].Trim()
            }
            Write-DebugMessage "Using Template2: $value"
            return $value
        }
        elseif ($SelectedTemplate -eq "Template3" -and $displayNameSection -match 'DisplayNameTemplate3=(.+)') {
            $value = $matches[1].Trim()
            Write-DebugMessage "Using Template3: $value"
            return $value
        }
        elseif ($SelectedTemplate -eq "Template4" -and $displayNameSection -match 'DisplayNameTemplate4=(.+)') {
            $value = $matches[1].Trim()
            Write-DebugMessage "Using Template4: $value"
            return $value
        }
        elseif ($SelectedTemplate -eq "Template5" -and $displayNameSection -match 'DisplayNameTemplate5=(.+)') {
            $value = $matches[1].Trim()
            Write-DebugMessage "Using Template5: $value"
            return $value
        }
        elseif ($SelectedTemplate -eq "Template6" -and $displayNameSection -match 'DisplayNameTemplate6=(.+)') {
            $value = $matches[1].Trim()
            Write-DebugMessage "Using Template6: $value"
            return $value
        }
        elseif ($SelectedTemplate -eq "Custom") {
            # Bei benutzerdefinierter Eingabe wird der Benutzerwert zurückgegeben
            Write-DebugMessage "Using Custom template: CUSTOM_VALUE"
            return "CUSTOM_VALUE"
        }
        else {
            # Standardwert zurückgeben, wenn nichts passt
            if ($displayNameSection -match 'DefaultDisplayNameFormat=(.+)') {
                $value = $matches[1].Trim()
                Write-DebugMessage "Using DefaultDisplayNameFormat from INI: $value"
                return $value
            }
        }
    }

    # Fallback-Option falls etwas schiefgeht
    Write-DebugMessage "Falling back to default template: {first} {last}"
    return "{first} {last}"
}

function Format-DisplayName {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$Format,

        [Parameter(Mandatory = $true)]
        [string]$FirstName,

        [Parameter(Mandatory = $true)]
        [string]$LastName,

        [Parameter(Mandatory = $false)]
        [string]$CustomValue = ""
    )

    if ($Format -eq "CUSTOM_VALUE" -and $CustomValue -ne "") {
        $result = $CustomValue
    }
    else {
        $result = $Format -replace '{first}', $FirstName -replace '{last}', $LastName
    }

    return $result
}


Write-DebugMessage "Loading AD groups."

#region [Region 13 | AD GROUP LOADING]
# Loads AD groups from the INI file and binds them to the GUI
Write-DebugMessage "Loading AD groups from INI."
function Load-ADGroups {
    try {
        Write-DebugMessage "Loading AD groups from INI..."
        if (-not $global:Config.Contains("ADGroups")) {
            Write-Error "The [ADGroups] section is missing in the INI."
            return
        }
        
        $adGroupsIni = $global:Config["ADGroups"]
        $icADGroups = $window.FindName("icADGroups")
        
        if ($null -eq $icADGroups) {
            Write-Error "ItemsControl 'icADGroups' not found. Check the XAML name."
            return
        }
        
        # Create a list of AD group items for binding
        $adGroupItems = New-Object System.Collections.ObjectModel.ObservableCollection[PSObject]
        
        # Check if we have numbered entries like ADGroup1, ADGroup2 etc.
        $numberedGroupKeys = $adGroupsIni.Keys | Where-Object { $_ -match '^ADGroup\d+$' }
        if ($numberedGroupKeys.Count -gt 0) {
            # Sort numerically by the number after "ADGroup"
            $sortedKeys = $numberedGroupKeys | Sort-Object { [int]($_ -replace 'ADGroup', '') }
            
            # Add any remaining keys
            $otherKeys = $adGroupsIni.Keys | Where-Object { $_ -notmatch '^ADGroup\d+$' }
            $sortedKeys += $otherKeys
        } else {
            # Use original order from the OrderedDictionary
            $sortedKeys = $adGroupsIni.Keys
        }
        
        Write-DebugMessage "Found AD groups: $($sortedKeys.Count)"
        
        # For each entry in the ADGroups section create a data object
        foreach ($key in $sortedKeys) {
            # The key (e.g., "DEV") is the display name
            $displayName = $key
            
            # The value (e.g., "TEAM-1") is the actual group name
            $groupName = $adGroupsIni[$key]
            
            # Check if the group name is not empty
            if (-not [string]::IsNullOrWhiteSpace($groupName)) {
                # Create a custom object with display name and actual group name
                $groupItem = [PSCustomObject]@{
                    DisplayName = $displayName
                    GroupName = $groupName
                    IsChecked = $false
                    Key = $key
                }
                
                $adGroupItems.Add($groupItem)
                Write-DebugMessage "Added group: Display='$displayName', Group='$groupName'"
            }
        }
        
        # Set the ItemsSource to our collection
        $icADGroups.ItemsSource = $adGroupItems
        
        # Override the ItemTemplate to bind to DisplayName and store GroupName
        $factory = New-Object System.Windows.FrameworkElementFactory([System.Windows.Controls.CheckBox])
        $factory.SetBinding([System.Windows.Controls.ContentControl]::ContentProperty, (New-Object System.Windows.Data.Binding("DisplayName")))
        $factory.SetBinding([System.Windows.Controls.Primitives.ToggleButton]::IsCheckedProperty, (New-Object System.Windows.Data.Binding("IsChecked") -Property @{Mode = [System.Windows.Data.BindingMode]::TwoWay}))
        $factory.SetResourceReference([System.Windows.Controls.Control]::MarginProperty, "Margin")
        $factory.SetValue([System.Windows.Controls.ToolTipService]::ToolTipProperty, (New-Object System.Windows.Data.Binding("GroupName") -Property @{StringFormat = "AD Group: {0}"}))
        
        $template = New-Object System.Windows.DataTemplate
        $template.VisualTree = $factory
        
        $icADGroups.ItemTemplate = $template
        
        Write-DebugMessage "AD groups successfully loaded from INI: $($adGroupItems.Count) groups."
        Write-Log "AD groups successfully loaded from INI." "DEBUG"
        
        # Store reference to the items collection
        $global:ADGroupItems = $adGroupItems
    }
    catch {
        $errorMessage = "Error loading AD groups from INI: $($_.Exception.Message)"
        Write-Error $errorMessage
        Write-Log $errorMessage "ERROR"
    }
}
# Function to update logos based on the selected tab
function Update-TabLogos {
    param (
        [Parameter(Mandatory = $true)]
        [string]$selectedTabName,
        
        [Parameter(Mandatory = $true)]
        $window
    )
    
    Write-DebugMessage "Updating logos for tab: $selectedTabName"
    
    # Logo-Update basierend auf ausgewähltem Tab
    switch ($selectedTabName) {
        "Tab_Onboarding" {
            Update-SingleLogo -window $window -logoControl "picLogo1" -configKey "HeaderLogo1"
        }
        "Tab_ADUpdate" {
            Update-SingleLogo -window $window -logoControl "picLogo2" -configKey "HeaderLogo1"
        }
        "Tab_INIEditor" {
            Update-SingleLogo -window $window -logoControl "picLogo3" -configKey "HeaderLogo1"
        }
    }
}

function Update-SingleLogo {
    param (
        [Parameter(Mandatory = $true)]
        $window,
        
        [Parameter(Mandatory = $true)]
        [string]$logoControl,
        
        [Parameter(Mandatory = $true)]
        [string]$configKey
    )
    
    $picLogo = $window.FindName($logoControl)
    
    if (-not $picLogo) {
        Write-DebugMessage "Logo control '$logoControl' not found in window"
        return
    }
    
    if (-not $global:Config["WPFGUI"].Contains($configKey)) {
        Write-DebugMessage "Config key '$configKey' not found in WPFGUI section"
        return
    }
    
    try {
        $logoPath = $global:Config["WPFGUI"][$configKey]
        Write-DebugMessage "Attempting to load logo from: $logoPath"
        
        if (Test-Path $logoPath) {
            $logo = New-Object System.Windows.Media.Imaging.BitmapImage
            $logo.BeginInit()
            $logo.CacheOption = [System.Windows.Media.Imaging.BitmapCacheOption]::OnLoad
            $logo.UriSource = New-Object System.Uri($logoPath, [System.UriKind]::Absolute)
            $logo.EndInit()
            $logo.Freeze() # Verbessert die Leistung
            
            # Set the image source
            $picLogo.Source = $logo
            Write-DebugMessage "Logo successfully set for $logoControl"
            
            # Add click event if URL is specified
            if ($global:Config["WPFGUI"].Contains("HeaderLogoURL")) {
                $picLogo.Cursor = [System.Windows.Input.Cursors]::Hand
                $picLogo.Tag = $global:Config["WPFGUI"]["HeaderLogoURL"] # Speichern der URL in Tag
                
                # Entfernen Sie zuerst alle vorhandenen Event-Handler
                $picLogo.RemoveHandler([System.Windows.Controls.Image]::MouseLeftButtonUpEvent, [System.Windows.Input.MouseButtonEventHandler]::new({param($sender,$e) Start-Process $sender.Tag}))
                
                # Fügen Sie den neuen Event-Handler hinzu
                $picLogo.AddHandler([System.Windows.Controls.Image]::MouseLeftButtonUpEvent, [System.Windows.Input.MouseButtonEventHandler]::new({
                    param($sender,$e)
                    $url = $sender.Tag
                    if (-not [string]::IsNullOrEmpty($url)) {
                        try {
                            Start-Process $url
                        } catch {
                            Write-DebugMessage "Error opening URL: $($_.Exception.Message)"
                        }
                    }
                }))
                Write-DebugMessage "Click event added for $logoControl"
            }
        } else {
            Write-DebugMessage "Logo file not found: $logoPath"
        }
    } catch {
        Write-DebugMessage "Error setting logo for $($_.Exception.Message)"
    }
}

# Funktion zum Initialisieren der Logo-Anzeige nach dem Laden des Fensters
function Initialize-LogoHandling {
    param (
        [Parameter(Mandatory = $true)]
        $window
    )

    # Verzögerte Ausführung sicherstellen
    $window.Dispatcher.InvokeAsync({
        # Korrekter Name des TabControls aus XAML (MainTabControl statt tabControl)
        $tabControl = $window.FindName("MainTabControl")
        
        if ($tabControl) {
            Write-DebugMessage "MainTabControl found, setting up event handlers"
            
            # Aktualisiere das Logo für den ersten Tab sofort
            $selectedTab = $tabControl.SelectedItem
            if ($selectedTab) {
                Update-TabLogos -selectedTabName $selectedTab.Name -window $window
                Write-DebugMessage "Updated logos for initial tab: $($selectedTab.Name)"
            } else {
                Write-DebugMessage "No tab selected initially"
            }
            
            # Handler für TabControl.SelectionChanged-Event
            $tabControl.Add_SelectionChanged({
                $selectedTab = $tabControl.SelectedItem
                if ($selectedTab) {
                    Update-TabLogos -selectedTabName $selectedTab.Name -window $window
                    Write-DebugMessage "Tab changed to: $($selectedTab.Name)"
                }
            })
            
            Write-DebugMessage "Tab selection event handler registered"
        } else {
            Write-DebugMessage "ERROR: MainTabControl not found in window"
        }
    }, [System.Windows.Threading.DispatcherPriority]::Loaded)
}
# Function to collect selected AD groups
function Get-SelectedADGroups {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [System.Windows.Controls.ItemsControl]$Panel,
        
        [Parameter(Mandatory=$false)]
        [string]$Separator = ";"
    )
    
    # Use a more efficient ArrayList to avoid output on add
    [System.Collections.ArrayList]$selectedGroups = @()
    
    if ($null -eq $Panel) {
        Write-DebugMessage "Error: Panel is null"
        return "NONE"
    }
    
    if ($null -eq $Panel.ItemsSource -or ($Panel.ItemsSource | Measure-Object).Count -eq 0) {
        Write-DebugMessage "Panel contains no items to select from"
        return "NONE"
    }
    
    # Get selected groups directly from data source
    foreach ($item in $Panel.ItemsSource) {
        if ($item.IsChecked) {
            Write-DebugMessage "Selected group: $($item.GroupName)"
            [void]$selectedGroups.Add($item.GroupName)
        }
    }
    
    Write-DebugMessage "Total selected groups: $($selectedGroups.Count)"
    
    if ($selectedGroups.Count -eq 0) {
        return "NONE"
    } else {
        return ($selectedGroups -join $Separator)
    }
}

# Call the function after loading the XAML:
Load-ADGroups
#endregion

Write-DebugMessage "Populating dropdowns (OU, Location, License, MailSuffix, DisplayName Template)."

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

#region [Region 14.2.1 | UPN TEMPLATE DISPLAY]
# Updates the UPN template display with the current template from INI
Write-DebugMessage "Initializing UPN Template Display."
function Update-UPNTemplateDisplay {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$false)]
        [string]$Template
    )
    
    try {
        $txtUPNTemplateDisplay = $window.FindName("txtUPNTemplateDisplay")
        if ($null -eq $txtUPNTemplateDisplay) {
            Write-DebugMessage "UPN Template Display TextBlock not found in XAML"
            return
        }
        
        # Default UPN template format
        $defaultUPNFormat = "FIRSTNAME.LASTNAME"
        
        # If a template was explicitly provided, use it
        if (-not [string]::IsNullOrWhiteSpace($Template)) {
            $defaultUPNFormat = $Template.ToUpper()
            Write-DebugMessage "Using provided template: $defaultUPNFormat"
        }
        # Otherwise read from INI if available
        elseif ($global:Config.Contains("DisplayNameUPNTemplates") -and 
            $global:Config.DisplayNameUPNTemplates.Contains("DefaultUserPrincipalNameFormat")) {
            $upnTemplate = $global:Config.DisplayNameUPNTemplates.DefaultUserPrincipalNameFormat
            if (-not [string]::IsNullOrWhiteSpace($upnTemplate)) {
                $defaultUPNFormat = $upnTemplate.ToUpper()
                Write-DebugMessage "Found UPN template in INI: $defaultUPNFormat"
            }
        }
        
        # Direct text update (most reliable method)
        $txtUPNTemplateDisplay.Text = $defaultUPNFormat
        Write-DebugMessage "UPN Template Display directly set to: $defaultUPNFormat"
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
        # Check what was selected
        if ($comboBoxDisplayTemplate.SelectedItem -ne $null) {
            $selectedItem = $comboBoxDisplayTemplate.SelectedItem
            Write-DebugMessage "Template dropdown selection changed: $($selectedItem | ConvertTo-Json -Compress)"
            
            # First try to get the Template property if it exists
            $selectedTemplate = if ($selectedItem.PSObject.Properties.Match('Template').Count -gt 0) {
                $selectedItem.Template
            } 
            # If that fails, try Name property
            elseif ($selectedItem.PSObject.Properties.Match('Name').Count -gt 0) {
                $selectedItem.Name
            }
            # If all else fails, try converting the whole item to string
            else {
                $selectedItem.ToString()
            }
            
            Write-DebugMessage "Selected template value: '$selectedTemplate'"
            
            # Update the display with the selected template
            if (-not [string]::IsNullOrWhiteSpace($selectedTemplate)) {
                Update-UPNTemplateDisplay -Template $selectedTemplate
            }
        }
        else {
            Write-DebugMessage "Template dropdown selection cleared or set to null"
        }
    })
    
    # Initial selection based on default from INI
    if ($global:Config.Contains("DisplayNameUPNTemplates") -and 
        $global:Config.DisplayNameUPNTemplates.Contains("DefaultDisplayNameFormat")) {
        
        $defaultFormat = $global:Config.DisplayNameUPNTemplates.DefaultDisplayNameFormat
        Write-DebugMessage "Setting initial selection to default format: $defaultFormat"
        
        # Find and select the matching item
        foreach ($item in $comboBoxDisplayTemplate.Items) {
            $itemTemplate = if ($item.PSObject.Properties.Match('Template').Count -gt 0) {
                $item.Template
            } elseif ($item.PSObject.Properties.Match('Name').Count -gt 0) {
                $item.Name
            } else {
                $item.ToString()
            }
            
            if ($itemTemplate -eq $defaultFormat) {
                $comboBoxDisplayTemplate.SelectedItem = $item
                Write-DebugMessage "Selected matching item for default format"
                break
            }
        }
    }
}
#endregion
#endregion

Write-DebugMessage "Defining utility functions."

#region [Region 15 | UTILITY FUNCTIONS]
# Miscellaneous utility functions for the application

Write-DebugMessage "Set-Logo"

#region [Region 15.1 | LOGO MANAGEMENT]
# Function to handle logo uploads and management for reports
Write-DebugMessage "Setting logo."
function Set-Logo {
    # [15.1.1 - Handles file selection and saving of logo images]
    param (
        [hashtable]$brandingConfig
    )
    try {
        Write-DebugMessage "GUI Opening file selection dialog for logo upload"
        $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
        $openFileDialog.Filter = "Image files (*.jpg;*.png;*.bmp)|*.jpg;*.jpeg;*.png;*.bmp"
        $openFileDialog.Title = "Select a logo for the onboarding document"
        if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $selectedFilePath = $openFileDialog.FileName
            Write-Log "Selected file: $selectedFilePath" "DEBUG"

            if (-not $brandingConfig.Contains("TemplatePath") -or [string]::IsNullOrWhiteSpace($brandingConfig["TemplatePath"])) {
                Throw "No 'TemplatePath' defined in the Report section."
            }

            $templatePath = $brandingConfig["TemplatePath"]
            if (-not (Test-Path $templatePath)) {
                try {
                    New-Item -ItemType Directory -Path $templatePath -Force -ErrorAction Stop | Out-Null
                } catch {
                    Throw "Could not create template directory for logo: $($_.Exception.Message)"
                }
            }

            $targetLogoPath = Join-Path -Path $templatePath -ChildPath "ReportHeader.jpg"
            Copy-Item -Path $selectedFilePath -Destination $targetLogoPath -Force -ErrorAction Stop
            Write-Log "Logo successfully saved as: $targetLogoPath" "DEBUG"
            [System.Windows.MessageBox]::Show("The logo was successfully saved!`nLocation: $targetLogoPath", "Success", "OK", "Information")
        }
    } catch {
        [System.Windows.MessageBox]::Show("Error uploading logo: $($_.Exception.Message)", "Error", "OK", "Error")
    }
}
#endregion

Write-DebugMessage "Defining advanced password generation function."

#region [Region 15.2 | PASSWORD GENERATION]
# Advanced password generation with security requirements
Write-DebugMessage "Generating advanced password."
function New-AdvancedPassword {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$false)]
        [int]$Length = 12,

        [Parameter(Mandatory=$false)]
        [int]$MinUpperCase = 2,

        [Parameter(Mandatory=$false)]
        [int]$MinDigits = 2,

        [Parameter(Mandatory=$false)]
        [bool]$AvoidAmbiguous = $false,

        [Parameter(Mandatory=$false)]
        [bool]$IncludeSpecial = $true,

        [Parameter(Mandatory=$false)]
        [ValidateRange(1,10)]
        [int]$MinNonAlpha = 2
    )

    if ($Length -lt ($MinUpperCase + $MinDigits)) {
        Throw "Error: The password length ($Length) is too short for the required minimum values (MinUpperCase + MinDigits = $($MinUpperCase + $MinDigits))."
    }

    Write-DebugMessage "New-AdvancedPassword: Defining character pools"
    $lower = 'abcdefghijklmnopqrstuvwxyz'
    $upper = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
    $digits = '0123456789'
    $special = '!@#$%^&*()'
    
    Write-DebugMessage "New-AdvancedPassword: Removing ambiguous characters if desired"
    if ($AvoidAmbiguous) {
        $ambiguous = 'Il1O0'
        $upper = -join ($upper.ToCharArray() | Where-Object { $ambiguous -notcontains $_ })
        $lower = -join ($lower.ToCharArray() | Where-Object { $ambiguous -notcontains $_ })
        $digits = -join ($digits.ToCharArray() | Where-Object { $ambiguous -notcontains $_ })
        $special = -join ($special.ToCharArray() | Where-Object { $ambiguous -notcontains $_ })
    }
    
    # Ensure that character pools are never empty
    if (-not $upper -or $upper.Length -eq 0) { $upper = 'ABCDEFGHJKLMNPQRSTUVWXYZ' }
    if (-not $lower -or $lower.Length -eq 0) { $lower = 'abcdefghijkmnopqrstuvwxyz' }
    if (-not $digits -or $digits.Length -eq 0) { $digits = '23456789' }
    if (-not $special -or $special.Length -eq 0) { $special = '!@#$%^&*()' }
    
    # Recalculate 'all' after ensuring pools are not empty
    $all = $lower + $upper + $digits
    if ($IncludeSpecial) { $all += $special }
    
    do {
        Write-DebugMessage "New-AdvancedPassword: Starting password generation"
        # Initialize as a simple string array
        $passwordChars = [System.Collections.ArrayList]::new()
        
        Write-DebugMessage "New-AdvancedPassword: Adding minimum number of uppercase letters"
        for ($i = 0; $i -lt $MinUpperCase; $i++) {
            [void]$passwordChars.Add($upper[(Get-Random -Minimum 0 -Maximum $upper.Length)].ToString())
        }
        
        Write-DebugMessage "New-AdvancedPassword: Adding minimum number of digits"
        for ($i = 0; $i -lt $MinDigits; $i++) {
            [void]$passwordChars.Add($digits[(Get-Random -Minimum 0 -Maximum $digits.Length)].ToString())
        }
        
        Write-DebugMessage "New-AdvancedPassword: Filling up to desired length"
        while ($passwordChars.Count -lt $Length) {
            [void]$passwordChars.Add($all[(Get-Random -Minimum 0 -Maximum $all.Length)].ToString())
        }
        
        Write-DebugMessage "New-AdvancedPassword: Randomizing order"
        # Get array of strings, then join them at the end
        $shuffledChars = $passwordChars | Get-Random -Count $passwordChars.Count
        $generatedPassword = -join $shuffledChars
        
        Write-DebugMessage "New-AdvancedPassword: Checking minimum number of non-alphabetic characters"
        # Count characters that don't match letters
        $nonAlphaCount = ($generatedPassword.ToCharArray() | Where-Object { $_ -notmatch '[a-zA-Z]' }).Count
        
    } while ($nonAlphaCount -lt $MinNonAlpha)
    
    Write-DebugMessage "Advanced password generated successfully."
    return $generatedPassword
}
#endregion

#region [Region 16 | CORE ONBOARDING PROCESS]
# Main function that handles the complete user onboarding process
Write-DebugMessage "Starting onboarding process."
function Invoke-Onboarding {
    # [16.1 - Processes AD user creation and configuration in a single workflow]
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [pscustomobject]$userData,

        [Parameter(Mandatory = $true)]
        [hashtable]$Config
    )

    Write-DebugMessage "Invoke-Onboarding: Start"

    # Debugging information
    if ($global:Config.Logging.DebugMode -eq "1") {
        Write-DebugMessage "[DEBUG] UPNFormat loaded from INI: $($Config.DisplayNameUPNTemplates.DefaultUserPrincipalNameFormat)"
        Write-DebugMessage "[DEBUG] userData: $($userData | ConvertTo-Json -Depth 1 -Compress)"
    }

    # [16.1.1 - ActiveDirectory Module Verification]
    # Checks if the ActiveDirectory module is loaded before proceeding with operations
    Write-DebugMessage "Invoke-Onboarding: Checking if ActiveDirectory module is already loaded"
    if (-not (Get-Module -Name ActiveDirectory)) {
        Throw "ActiveDirectory module is not loaded. Please ensure it is installed and pre-imported."
    }

    # [16.1.2 - Check mandatory fields: FirstName]
    Write-DebugMessage "Invoke-Onboarding: Checking mandatory fields"
    if ([string]::IsNullOrWhiteSpace($userData.FirstName)) {
        Throw "First Name is required!"
    }
    if ([string]::IsNullOrWhiteSpace($userData.LastName)) {
        Throw "Last Name is required!"
    }
    
    # Validate OU
    $selectedOU = if (-not [string]::IsNullOrWhiteSpace($userData.OU)) {
        $userData.OU
    } else {
        if (-not $Config.ADUserDefaults.Contains("DefaultOU") -or [string]::IsNullOrWhiteSpace($Config.ADUserDefaults["DefaultOU"])) {
            Throw "Error: No OU specified and no DefaultOU found in the configuration!"
        }
        $Config.ADUserDefaults["DefaultOU"]
    }
    Write-DebugMessage "Selected OU: $selectedOU"

    # Debug outputs with null checks
    Write-DebugMessage "DEBUG LINE 898: Checking all important variables"
    $SamAccountName = $null
    $ADGroupsSelected = if ($userData.PSObject.Properties.Match("ADGroupsSelected").Count -gt 0) { $userData.ADGroupsSelected } else { "NONE" }
    $License = if ($userData.PSObject.Properties.Match("License").Count -gt 0) { $userData.License } else { "None" }
    
    Write-DebugMessage "SamAccountName: Will be generated"
    Write-DebugMessage "UPNFormat: $($userData.UPNFormat ?? 'Not defined')"
    Write-DebugMessage "MailSuffix: $($userData.MailSuffix ?? 'Not defined')"
    Write-DebugMessage "ADGroupsSelected: $ADGroupsSelected"
    Write-DebugMessage "License: $License"
    
    # [16.1.3 - Password generation or fixed password]
    Write-DebugMessage "Invoke-Onboarding: Determining main password"
    $mainPW = $null
    if ($userData.PasswordMode -eq 1) {
        # [16.1.3.1 - Use advanced generation]
        Write-DebugMessage "Invoke-Onboarding: Checking [PasswordFixGenerate] keys"
        foreach ($key in @("DefaultPasswordLength","IncludeSpecialChars","AvoidAmbiguousChars","MinNonAlpha","MinUpperCase","MinDigits")) {
            if (-not $Config.PasswordFixGenerate.Contains($key)) {
                Throw "Error: '$key' is missing in [PasswordFixGenerate]!"
            }
        }
        Write-DebugMessage "Invoke-Onboarding: Reading password settings from INI"
        $includeSpecial = ($Config.PasswordFixGenerate["IncludeSpecialChars"] -match "^(?i:true|1)$")
        $avoidAmbiguous = ($Config.PasswordFixGenerate["AvoidAmbiguousChars"] -match "^(?i:true|1)$")

        Write-DebugMessage "Invoke-Onboarding: Generating password with New-AdvancedPassword"
        $mainPW = New-AdvancedPassword `
            -Length ([int]$Config.PasswordFixGenerate["DefaultPasswordLength"]) `
            -IncludeSpecial $includeSpecial `
            -AvoidAmbiguous $avoidAmbiguous `
            -MinUpperCase ([int]$Config.PasswordFixGenerate["MinUpperCase"]) `
            -MinDigits ([int]$Config.PasswordFixGenerate["MinDigits"]) `
            -MinNonAlpha ([int]$Config.PasswordFixGenerate["MinNonAlpha"])
        
        if ([string]::IsNullOrWhiteSpace($mainPW)) {
            Throw "Error: Generated password is empty. Please check [PasswordFixGenerate]."
        }
    }
    else {
        # [16.1.3.2 - Use fixPassword from INI]
        Write-DebugMessage "Invoke-Onboarding: Using fixPassword from INI"
        if (-not $Config.PasswordFixGenerate.Contains("fixPassword")) {
            Throw "Error: fixPassword is missing in PasswordFixGenerate!"
        }
        $mainPW = $Config.PasswordFixGenerate["fixPassword"]
        if ([string]::IsNullOrWhiteSpace($mainPW)) {
            Throw "Error: No fixed password provided in the INI!"
        }
    }

    # [16.1.3.3 - Convert to secure]
    $SecurePW = ConvertTo-SecureString $mainPW -AsPlainText -Force

    # [16.1.4 - Generate additional (custom) passwords (5 total)]
    Write-DebugMessage "Invoke-Onboarding: Generating additional custom passwords"
    $customPWLabels = if ($Config.Contains("CustomPWLabels")) { $Config["CustomPWLabels"] } else { @{} }
    # PowerShell 7+ null-coalescing Operator (??) replaced with alternative method for PowerShell 5 compatibility
    $pwLabel1 = if ($customPWLabels.Contains("CustomPW1_Label")) { $customPWLabels["CustomPW1_Label"] } else { "Custom PW #1" }
    $pwLabel2 = if ($customPWLabels.Contains("CustomPW2_Label")) { $customPWLabels["CustomPW2_Label"] } else { "Custom PW #2" }
    $pwLabel3 = if ($customPWLabels.Contains("CustomPW3_Label")) { $customPWLabels["CustomPW3_Label"] } else { "Custom PW #3" }
    $pwLabel4 = if ($customPWLabels.Contains("CustomPW4_Label")) { $customPWLabels["CustomPW4_Label"] } else { "Custom PW #4" }
    $pwLabel5 = if ($customPWLabels.Contains("CustomPW5_Label")) { $customPWLabels["CustomPW5_Label"] } else { "Custom PW #5" }

    $CustomPW1 = New-AdvancedPassword
    $CustomPW2 = New-AdvancedPassword
    $CustomPW3 = New-AdvancedPassword
    $CustomPW4 = New-AdvancedPassword
    $CustomPW5 = New-AdvancedPassword

    # [16.1.5 - SamAccountName and UPN creation logic]
    Write-DebugMessage "Invoke-Onboarding: Creating SamAccountName and UPN"
    # Use only the return value of the New-UPN function:
    $upnResult = New-UPN -userData $userData -Config $Config
    $SamAccountName = $upnResult.SamAccountName
    $UPN = $upnResult.UPN
    $companySection = $upnResult.CompanySection
    
    Write-DebugMessage "Invoke-Onboarding: Reading [Company] data from config using section '$companySection'"
    if (-not $Config.Contains($companySection)) {
        Throw "Error: Section '$companySection' is missing in the INI!"
    }
    $companyData = $Config[$companySection]   

    # [16.1.5.1 - Enhanced logging after UPN creation]
    Write-DebugMessage "Generated SamAccountName: $SamAccountName"
    Write-DebugMessage "Generated UPN: $UPN"
    Write-DebugMessage "Using company section: $companySection"
    
    # [16.1.5.2 - Additional validation based on configuration]
    if ($Config.Contains("ValidateUPN") -and $Config.ValidateUPN.Contains("CheckForDuplicates") -and 
        $Config.ValidateUPN["CheckForDuplicates"] -eq "1") {
        
        Write-DebugMessage "Checking for UPN duplicates: $UPN"
        try {
            $existingUser = Get-ADUser -Filter {UserPrincipalName -eq $UPN} -ErrorAction Stop
            if ($existingUser) {
                Write-DebugMessage "WARNING: UPN $UPN is already in use by $($existingUser.SamAccountName)!"
                
                if ($Config.ValidateUPN.Contains("AppendRandomOnConflict") -and 
                    $Config.ValidateUPN["AppendRandomOnConflict"] -eq "1") {
                    
                    $random = -join ((48..57) + (97..122) | Get-Random -Count 3 | ForEach-Object {[char]$_})
                    $UPNBase = $UPN.Split('@')[0]
                    $UPNDomain = $UPN.Split('@')[1]
                    $UPN = "$UPNBase$random@$UPNDomain"
                    Write-DebugMessage "UPN modified to avoid conflict: $UPN"
                }
                else {
                    Throw "Error: UPN $UPN is already in use and conflict resolution is disabled!"
                }
            }
        }
        catch {
            if ($_.Exception.Message -notmatch "No user matches filter criteria") {
                Throw "Error checking for duplicate UPN: $($_.Exception.Message)"
            }
        }
    }
    
    # [16.1.5.3 - Apply name normalization for special characters if configured]
    if ($Config.Contains("NameNormalization") -and $Config.NameNormalization.Contains("Enabled") -and 
        $Config.NameNormalization["Enabled"] -eq "1") {
        
        Write-DebugMessage "Applying name normalization rules"
        
        # Handle special characters in names based on configuration
        if ($Config.NameNormalization.Contains("ReplaceSpecialChars")) {
            $replacements = $Config.NameNormalization["ReplaceSpecialChars"].Split(';')
            foreach ($replacement in $replacements) {
                if ($replacement -match '^(.+?)=(.+?)$') {
                    $from = $matches[1]
                    $to = $matches[2]
                    $UPN = $UPN -replace [regex]::Escape($from), $to
                    $SamAccountName = $SamAccountName -replace [regex]::Escape($from), $to
                    Write-DebugMessage "Replaced '$from' with '$to' in identifiers"
                }
            }
        }
        
        # Additional character normalization
        if ($Config.NameNormalization.Contains("NormalizeCase") -and 
            $Config.NameNormalization["NormalizeCase"] -eq "1") {
            $SamAccountName = $SamAccountName.ToLower()
            $UPNPrefix = $UPN.Split('@')[0]
            $UPNDomain = $UPN.Split('@')[1]
            $UPN = "$($UPNPrefix.ToLower())@$UPNDomain"
            Write-DebugMessage "Normalized case: $SamAccountName, $UPN"
        }
    }
    
    # [16.1.5.4 - Initialize attributes collection for later use]
    $otherAttributes = @{ }
    
    # [16.1.5.5 - Check for custom attribute mappings]
    if ($Config.Contains("CustomAttributeMappings")) {
        Write-DebugMessage "Processing custom attribute mappings"
        foreach ($key in $Config.CustomAttributeMappings.Keys) {
            $attributeName = $key
            $userDataProperty = $Config.CustomAttributeMappings[$key]
            
            if ($userData.PSObject.Properties.Match($userDataProperty).Count -gt 0 -and 
                -not [string]::IsNullOrWhiteSpace($userData.$userDataProperty)) {
                $otherAttributes[$attributeName] = $userData.$userDataProperty
                Write-DebugMessage "Mapped custom attribute: $attributeName = $($userData.$userDataProperty)"
            }
        }
    }
    
    # [16.1.5.6 - Record normalized values back to userData]
    $userData | Add-Member -MemberType NoteProperty -Name "NormalizedSamAccountName" -Value $SamAccountName -Force
    $userData | Add-Member -MemberType NoteProperty -Name "NormalizedUPN" -Value $UPN -Force
    
    Write-DebugMessage "UPN and SamAccountName processing complete"
    # [16.1.6 - Read company data from [Company*] in the INI]
    Write-DebugMessage "Invoke-Onboarding: Reading [Company] data from config"
    if (-not $Config.Contains($companySection)) {
        Throw "Error: Section '$companySection' is missing in the INI!"
    }
    $companyData = $Config[$companySection]
    $suffix = ($companySection -replace "\D","")

    # [16.1.6.1 - Basic address info from company]
    $streetKey = "CompanyStrasse"
    $plzKey    = "CompanyPLZ"
    $ortKey    = "CompanyOrt"
    $Street = Get-ConfigValue -Section $companyData -Key "CompanyStrasse"
    $Zip    = Get-ConfigValue -Section $companyData -Key "CompanyPLZ"
    $City   = Get-ConfigValue -Section $companyData -Key "CompanyOrt"
    
    # [16.1.6.1.1 - Get CompanyDomain for user website attribute]
    $companyDomain = ""
    if ($companyData.Contains("CompanyDomain")) {
        $companyDomain = $companyData["CompanyDomain"]
        Write-DebugMessage "Found CompanyDomain: $companyDomain"
        
        # Add to otherAttributes for later use when creating the user
        if (-not [string]::IsNullOrWhiteSpace($companyDomain)) {
            $otherAttributes["wWWHomePage"] = $companyDomain
            Write-DebugMessage "Set user website attribute to: $companyDomain"
        }
    }

    # [16.1.6.2 - Set mail suffix from userData or use fallback]
    Write-DebugMessage "Invoke-Onboarding: Checking mail suffix from userData"

    # Extra debugging for ComboBox selection
    Write-DebugMessage "DEBUG: All userData properties: $($userData.PSObject.Properties.Name -join ', ')"
    Write-DebugMessage "DEBUG: userData.MailSuffix exists: $($userData.PSObject.Properties.Name -contains 'MailSuffix')"
    if ($userData.PSObject.Properties.Name -contains 'MailSuffix') {
        Write-DebugMessage "DEBUG: userData.MailSuffix value: '$($userData.MailSuffix)'"
    }

    $mailSuffix = $null

    # Direct access to MailSuffix property with stronger validation
    if ($userData.PSObject.Properties.Name -contains 'MailSuffix' -and 
        $null -ne $userData.MailSuffix -and
        -not [string]::IsNullOrWhiteSpace($userData.MailSuffix.ToString())) {
        
        $mailSuffix = $userData.MailSuffix.ToString().Trim()
        Write-DebugMessage "Using mail suffix directly from userData.MailSuffix: '$mailSuffix'"
    }
    # If dropdown selection wasn't captured properly, try getting from global variable
    elseif ($global:SelectedMailSuffix -and -not [string]::IsNullOrWhiteSpace($global:SelectedMailSuffix)) {
        $mailSuffix = $global:SelectedMailSuffix.Trim()
        Write-DebugMessage "Using mail suffix from global variable: '$mailSuffix'"
    }
    # Fallback to company configuration
    else {
        Write-DebugMessage "No mail suffix in userData, checking company section: $companySection"
        
        # First check the selected company section
        if ($companyData.Contains("CompanyMailDomain")) {
            $mailSuffix = $companyData["CompanyMailDomain"]
            Write-DebugMessage "Using mail suffix from [$companySection].CompanyMailDomain: '$mailSuffix'"
        }
        # Fall back to the default Company section if necessary
        elseif ($Config.Contains("Company") -and $Config["Company"].Contains("CompanyMailDomain")) {
            $mailSuffix = $Config["Company"]["CompanyMailDomain"]
            Write-DebugMessage "Using mail suffix from [Company].CompanyMailDomain: '$mailSuffix'"
        }
        else {
            $mailSuffix = ""
            Write-DebugMessage "No CompanyMailDomain found in configuration, proceeding with empty mail suffix"
        }
    }

    # Add the following code right before the "Onboarding" button handling
    # to capture combobox selection changes:
    $comboBoxMailSuffix = $window.FindName("cmbSuffix")
    if ($comboBoxMailSuffix) {
        $comboBoxMailSuffix.Add_SelectionChanged({
            $global:SelectedMailSuffix = $comboBoxMailSuffix.SelectedValue
            Write-DebugMessage "Mail suffix selection changed to: $global:SelectedMailSuffix"
        })
    }

    # Ensure mail suffix starts with @
    if (-not [string]::IsNullOrWhiteSpace($mailSuffix) -and -not $mailSuffix.StartsWith('@')) {
        $mailSuffix = "@" + $mailSuffix
        Write-DebugMessage "Added @ prefix to mail suffix: '$mailSuffix'"
    }

    # Initialize empty email variable for later use
    $email = if (-not [string]::IsNullOrWhiteSpace($userData.EmailAddress)) {
        $userData.EmailAddress
    } else {
        ""
    }
    Write-DebugMessage "Final values - Email address: '$email', Mail suffix: '$mailSuffix'"

    # Initialize an array to store proxy addresses
    $proxyAddresses = [System.Collections.ArrayList]@()

    # Initialize an array to store proxy addresses
    $proxyAddresses = [System.Collections.ArrayList]@()

    # Only set the mail attribute if EmailAddress is provided
    if (-not [string]::IsNullOrWhiteSpace($userData.EmailAddress)) {
        # Check if EmailAddress already contains @ symbol (complete email address)
        if ($userData.EmailAddress -match '@') {
            $primaryMail = $userData.EmailAddress
        } else {
            # Ensure mailSuffix starts with @ for proper concatenation
            if (-not [string]::IsNullOrWhiteSpace($mailSuffix) -and -not $mailSuffix.StartsWith('@')) {
                $mailSuffix = "@" + $mailSuffix
            }
            $primaryMail = "$($userData.EmailAddress)$mailSuffix"
        }
        
        # Set the primary mail attribute
        $otherAttributes["mail"] = $primaryMail
        
        # Add the primary SMTP proxy address (uppercase SMTP denotes primary)
        [void]$proxyAddresses.Add("SMTP:$primaryMail")
        Write-DebugMessage "Added primary proxy address: SMTP:$primaryMail"
        
        # Check if there's a MS365 domain defined in the Company section
        if ($companyData.Contains("CompanyMS365Domain") -and -not [string]::IsNullOrWhiteSpace($companyData["CompanyMS365Domain"])) {
            $ms365Domain = $companyData["CompanyMS365Domain"]
            # Ensure domain has @ prefix
            if (-not $ms365Domain.StartsWith('@')) {
                $ms365Domain = "@" + $ms365Domain
            }
            
            # Extract the username part from the primary email
            $username = ""
            if ($primaryMail -match '@') {
                $username = $primaryMail.Substring(0, $primaryMail.IndexOf('@'))
            } else {
                $username = $userData.EmailAddress
            }
            
            # Add the secondary MS365 proxy address (lowercase smtp for secondary)
            $ms365Address = "$username$ms365Domain"
            [void]$proxyAddresses.Add("smtp:$ms365Address")
            Write-DebugMessage "Added secondary MS365 proxy address: smtp:$ms365Address"
        }
        
        # Set the proxyAddresses attribute if we have any addresses
        if ($proxyAddresses.Count -gt 0) {
            $otherAttributes["proxyAddresses"] = $proxyAddresses.ToArray()
            Write-DebugMessage "Set proxyAddresses attribute with $($proxyAddresses.Count) addresses"
        }
    }

    # [16.1.6.3 - Country]
    $countryKey = "CompanyCountry$suffix"
    if (-not $companyData.Contains($countryKey) -or -not $companyData[$countryKey]) {
        Throw "Error: '$countryKey' is missing in the INI!"
    }
    $Country = $companyData[$countryKey]
    Write-DebugMessage "Country = $Country"

    # [16.1.6.4 - Company Display]
    $companyNameKey = "CompanyNameFirma$suffix"
    if (-not $companyData.Contains($companyNameKey) -or -not $companyData[$companyNameKey]) {
        Throw "Error: '$companyNameKey' is missing in the INI!"
    }
    $companyDisplay = $companyData[$companyNameKey]
    Write-DebugMessage "Company Display = $companyDisplay"

    # [16.1.7 - Integrate ADUserDefaults from INI (like mustChange, disabled, etc.)]
    Write-DebugMessage "Invoke-Onboarding: Checking [ADUserDefaults]"
    $adDefaults = $null
    if ($Config.Contains("ADUserDefaults")) {
        $adDefaults = $Config.ADUserDefaults
    }
    else {
        $adDefaults = @{ }
    }

    # [16.1.7.1 - OU from [ADUserDefaults].DefaultOU]
    if (-not $Config.ADUserDefaults.Contains("DefaultOU")) {
        Throw "Error: 'DefaultOU' is missing in [ADUserDefaults]!"
    }
    $defaultOU = $Config.ADUserDefaults["DefaultOU"]

    # [16.1.8 - Build AD user parameters]
    Write-DebugMessage "Invoke-Onboarding: Building userParams for AD"
    $userParams = [ordered]@{
        # Basic information
        Name                = $userData.DisplayName
        DisplayName         = $userData.DisplayName
        GivenName           = $userData.FirstName
        Surname             = $userData.LastName
        SamAccountName      = $SamAccountName
        UserPrincipalName   = $UPN
        AccountPassword     = $SecurePW

        # Enable/disable account
        Enabled             = (-not $userData.AccountDisabled)

        # AD User Defaults
        ChangePasswordAtLogon = if ($adDefaults.Contains("MustChangePasswordAtLogon")) {
                                    $adDefaults["MustChangePasswordAtLogon"] -eq "True"
                                } else {
                                    $userData.MustChangePassword
                                }
        PasswordNeverExpires  = if ($adDefaults.Contains("PasswordNeverExpires")) {
                                    $adDefaults["PasswordNeverExpires"] -eq "True"
                                } else {
                                    $userData.PasswordNeverExpires
                                }

        # OU (use the value selected from the dropdown – make sure it's not empty)
        Path                = $selectedOU

        # Address data
        City                = $City
        StreetAddress       = $Street
        Country             = $Country
        postalCode          = $Zip

        # Company name/display
        Company             = $companyDisplay
    }

    # [16.1.8.1 - Additional defaults from [ADUserDefaults] for home/profile/logonscript]
    if ($adDefaults.Contains("HomeDirectory")) {
        $userParams["HomeDirectory"] = $adDefaults["HomeDirectory"] -replace "%username%", $SamAccountName
    }
    if ($adDefaults.Contains("ProfilePath")) {
        $userParams["ProfilePath"] = $adDefaults["ProfilePath"] -replace "%username%", $SamAccountName
    }
    if ($adDefaults.Contains("LogonScript")) {
        $userParams["ScriptPath"] = $adDefaults["LogonScript"]
    }

    # [16.1.8.2 - Collect additional attributes from the form]
    Write-DebugMessage "Invoke-Onboarding: Building otherAttributes"
    $otherAttributes = @{ }

    # [16.1.8.3 - Determine the mail suffix: If available, use the one from the form, otherwise from the INI from [MailEndungen]/Domain1]
    if (-not [string]::IsNullOrWhiteSpace($userData.MailSuffix)) {
        $mailSuffix = $userData.MailSuffix.Trim()
    } else {
        Write-DebugMessage "No mail suffix specified, fallback to [MailEndungen].Domain1"
        if (-not ($Config.Contains("MailEndungen") -and $Config["MailEndungen"])) {
            Throw "Error: Section 'MailEndungen' is not defined in the INI!"
        }
        if (-not $Config["MailEndungen"].Contains("Domain1")) {
            Throw "Error: Key 'Domain1' is missing in section 'MailEndungen'!"
        }
        $mailSuffix = $Config["MailEndungen"]["Domain1"]
    }

    # [16.1.8.4 - Ensure EmailAddress has a valid (non-null) value]
    $emailAddress = $userData.EmailAddress
    if ($emailAddress -eq $null) { 
        $emailAddress = "" 
    }

    # [16.1.8.5 - First set the "mail" attribute from EmailAddress + mailSuffix]
    $otherAttributes["mail"] = "$emailAddress$mailSuffix"

    # [16.1.8.6 - If EmailAddress is present but does not contain "@", use an alternative suffix]
    if (-not [string]::IsNullOrWhiteSpace($userData.EmailAddress)) {
        if ($userData.EmailAddress -notmatch "@") {
            $otherAttributes["mail"] = $userData.EmailAddress + $mailSuffix
        }
        else {
            $otherAttributes["mail"] = $userData.EmailAddress
        }
    } else {
        # [16.1.8.7 - If Email is empty]
        $otherAttributes["mail"] = "placeholder$mailSuffix"
    }

    if ($userData.Description)    { $otherAttributes["description"] = $userData.Description }
    if ($userData.OfficeRoom)     { $otherAttributes["physicalDeliveryOfficeName"] = $userData.OfficeRoom }
    if ($userData.PhoneNumber)    { $otherAttributes["telephoneNumber"] = $userData.PhoneNumber }
    if ($userData.MobileNumber)   { $otherAttributes["mobile"] = $userData.MobileNumber }
    if ($userData.Position)       { $otherAttributes["title"] = $userData.Position }
    if ($userData.DepartmentField){ $otherAttributes["department"] = $userData.DepartmentField }

    # [16.1.8.8 - If there's a CompanyTelefon in the config (company phone), add as ipPhone]
    if ($companyData.Contains("CompanyTelefon") -and -not [string]::IsNullOrWhiteSpace($companyData["CompanyTelefon"])) {
        $otherAttributes["ipPhone"] = $companyData["CompanyTelefon"].Trim()
    }

    if ($otherAttributes.Count -gt 0) {
        $userParams["OtherAttributes"] = $otherAttributes
    }

    # [16.1.8.9 - If there's an Ablaufdatum (termination), set accountExpirationDate]
    if (-not [string]::IsNullOrWhiteSpace($userData.Ablaufdatum)) {
        try {
            $expirationDate = [DateTime]::Parse($userData.Ablaufdatum)
            $userParams["AccountExpirationDate"] = $expirationDate
        }
        catch {
            Write-Warning "Invalid termination date format: $($_.Exception.Message)"
        }
    }

    # [16.1.9 - Create or update AD user]
    Write-DebugMessage "Invoke-Onboarding: Checking if user already exists in AD"
    $existingUser = $null
    try {
        $existingUser = Get-ADUser -Filter { SamAccountName -eq $SamAccountName } -ErrorAction SilentlyContinue
    }
    catch {
        $existingUser = $null
    }

    if (-not $existingUser) {
        Write-DebugMessage "Invoke-Onboarding: Creating new user: $($userData.DisplayName)"
        Invoke-ADCommand -Command { New-ADUser @userParams -ErrorAction Stop } -ErrorContext "Creating AD user for $($userData.DisplayName)"
        
        # [16.1.9.1 - Smartcard Logon setting]
        if ($userData.SmartcardLogonRequired) {
            Write-DebugMessage "Invoke-Onboarding: Setting SmartcardLogonRequired for $($userData.DisplayName)"
            try {
                Set-ADUser -Identity $SamAccountName -SmartcardLogonRequired $true -ErrorAction Stop
                Write-DebugMessage "Smartcard logon requirement set successfully."
            }
            catch {
                Write-Warning "Error setting Smartcard logon"
            }
        }
        
        # [16.1.9.2 - "CannotChangePassword" – Note that this functionality is not yet implemented]
        if ($userData.CannotChangePassword) {
            Write-DebugMessage "Invoke-Onboarding: (Note) 'CannotChangePassword' via ACL not yet implemented."
        }
    }
    else {
        Write-DebugMessage "Invoke-Onboarding: User '$SamAccountName' already exists - updating attributes"
        try {
            Set-ADUser -Identity $existingUser.DistinguishedName `
                -GivenName $userData.FirstName `
                -Surname $userData.LastName `
                -City $City `
                -StreetAddress $Street `
                -Country $Country `
                -Enabled (-not $userData.AccountDisabled) -ErrorAction SilentlyContinue
    
            Set-ADUser -Identity $existingUser.DistinguishedName `
                -ChangePasswordAtLogon:$userData.MustChangePassword `
                -PasswordNeverExpires:$userData.PasswordNeverExpires
    
            if ($otherAttributes.Count -gt 0) {
                Set-ADUser -Identity $existingUser.DistinguishedName -Replace $otherAttributes -ErrorAction SilentlyContinue
            }
        }
        catch {
            Write-Warning "Error updating user: $($_.Exception.Message)"
        }
    }    

    # [16.1.9.3 - Setting AD account password (Reset)]
    Write-DebugMessage "Invoke-Onboarding: Setting AD account password (Reset)"
    try {
        Set-ADAccountPassword -Identity $SamAccountName -Reset -NewPassword $SecurePW -ErrorAction SilentlyContinue
    }
    catch {
        Write-Warning "Error setting password: $($_.Exception.Message)"
    }

    # [16.1.10 - Handling AD group assignments]
    Write-DebugMessage "Invoke-Onboarding: Handling AD group assignments"
    if ($userData.External) {
        Write-DebugMessage "External user: skipping default AD group assignment."
    }
    else {
        # [16.1.10.1 - Check any ADGroupsSelected from the GUI]
        Write-DebugMessage "Processing ADGroupsSelected: $($userData.ADGroupsSelected | Out-String)"

        # Ensure $adGroupsToProcess is always a string array
        $adGroupsToProcess = @()
        if ($null -ne $userData.ADGroupsSelected) {
            if ($userData.ADGroupsSelected -is [string]) {
                if ($userData.ADGroupsSelected -eq "NONE") {
                    Write-DebugMessage "No AD groups selected ('NONE')"
                    $adGroupsToProcess = @()
                } else {
                    # Split the semicolon-separated string into an array
                    $adGroupsToProcess = $userData.ADGroupsSelected.Split(';', [StringSplitOptions]::RemoveEmptyEntries)
                    Write-DebugMessage "Split ADGroupsSelected string into array of $($adGroupsToProcess.Count) items"
                }
            }
            elseif ($userData.ADGroupsSelected -is [array]) {
                $adGroupsToProcess = $userData.ADGroupsSelected
                Write-DebugMessage "Using array of $($adGroupsToProcess.Count) items directly from ADGroupsSelected"
            }
            else {
                # Force conversion to string then split (handles other object types)
                $groupString = $userData.ADGroupsSelected.ToString()
                $adGroupsToProcess = $groupString.Split(';', [StringSplitOptions]::RemoveEmptyEntries)
                Write-DebugMessage "Converted object to string and split into array of $($adGroupsToProcess.Count) items"
            }
        }

        foreach ($grpKey in $adGroupsToProcess) {
            Write-DebugMessage "Processing group key: '$grpKey'"
            
            # First try to match by display name directly in AD
            try {
                Write-DebugMessage "Attempting to add user directly to group '$grpKey' by name"
                Add-ADGroupMember -Identity $grpKey -Members $SamAccountName -ErrorAction Stop
                Write-DebugMessage "Successfully added user to group '$grpKey' directly."
                continue
            } catch {
                Write-DebugMessage "Could not add directly to group name '$grpKey': $($_.Exception.Message)"
                # Continue to try other methods
            }
            
            # Next try direct lookup using the value as a key
            if ($grpKey -and $Config.ADGroups.Contains($grpKey)) {
                $groupName = $Config.ADGroups[$grpKey]
                Write-DebugMessage "Found group name '$groupName' for key '$grpKey'"
                if ($groupName) {
                    try {
                        Add-ADGroupMember -Identity $groupName -Members $SamAccountName -ErrorAction Stop
                        Write-DebugMessage "Added user to group '$groupName' via config key."
                    }
                    catch {
                        Write-Warning "Error adding group '$groupName': $($_.Exception.Message)"
                    }
                }
            } else {
                # Try to find if this is a display name/label by searching through values
                $found = $false
                foreach ($adGroupKey in $Config.ADGroups.Keys) {
                    # Look for keys that match ADGroup pattern and have matching label or value
                    if ($adGroupKey -match '^ADGroup\d+$') {
                        $labelKey = "${adGroupKey}_Label"
                        # Check if display label matches
                        if (($Config.ADGroups.Contains($labelKey) -and $Config.ADGroups[$labelKey] -eq $grpKey) -or
                            $Config.ADGroups[$adGroupKey] -eq $grpKey) {
                            
                            $found = $true
                            $groupName = $Config.ADGroups[$adGroupKey]
                            Write-DebugMessage "Found group name '$groupName' via label/value lookup"
                            
                            try {
                                Add-ADGroupMember -Identity $groupName -Members $SamAccountName -ErrorAction Stop
                                Write-DebugMessage "Added user to group '$groupName'."
                            }
                            catch {
                                Write-Warning "Error adding group '$groupName': $($_.Exception.Message)"
                            }
                            break
                        }
                    }
                }
                
                if (-not $found) {
                    Write-DebugMessage "Group key '$grpKey' not found in Config.ADGroups and is not a valid AD group name"
                }
            }
        }
    }

    # [16.1.10.2 - Team Leader (TL) or Department Head (AL)]
    # Fix for TL group assignment
    if ($userData.TL -eq $true) {
        Write-DebugMessage "Invoke-Onboarding: TL is true - adding user to selected TLGroup: '$($userData.TLGroup)'"
        # Check both direct key lookup and if it exists as a key in the TLGroups section
        if (-not [string]::IsNullOrWhiteSpace($userData.TLGroup)) {
            $tlGroupName = ""
            if ($Config.TLGroups.Contains($userData.TLGroup)) {
                $tlGroupName = $Config.TLGroups[$userData.TLGroup]
                Write-DebugMessage "Found TL group '$tlGroupName' via key lookup"
            } elseif ($Config.TLGroups.Values -contains $userData.TLGroup) {
                $tlGroupName = $userData.TLGroup
                Write-DebugMessage "Using direct TL group name: $tlGroupName"
            }
            
            if (-not [string]::IsNullOrWhiteSpace($tlGroupName)) {
                try {
                    Add-ADGroupMember -Identity $tlGroupName -Members $SamAccountName -ErrorAction Stop
                    Write-DebugMessage "Added user to Team-Leader group '$tlGroupName'."
                }
                catch {
                    Write-Warning "Error adding TL group '$tlGroupName': $($_.Exception.Message)"
                }
            } else {
                Write-DebugMessage "Could not resolve TL group name from '$($userData.TLGroup)'"
            }
        }
        else {
            Write-DebugMessage "No TLGroup value specified though TL is true"
        }
    }

    # Fix for AL group assignment
    if ($userData.AL -eq $true) {
        Write-DebugMessage "Invoke-Onboarding: AL is true - adding user to ALGroup"
        if ($Config.Contains("ALGroup") -and $Config.ALGroup.Contains("Group")) {
            $alGroupName = $Config.ALGroup["Group"]
            Write-DebugMessage "Found ALGroup: $alGroupName"
            try {
                Add-ADGroupMember -Identity $alGroupName -Members $SamAccountName -ErrorAction Stop
                Write-DebugMessage "Added user to Abteilungsleiter group '$alGroupName'."
            }
            catch {
                Write-Warning "Error adding AL group '$alGroupName': $($_.Exception.Message)"
            }
        }
        else {
            Write-DebugMessage "ALGroup configuration is missing or invalid."
        }
    }            

    # [16.1.10.3 - Add user to default groups from [UserCreationDefaults]]
    if ($Config.Contains("UserCreationDefaults") -and $Config.UserCreationDefaults.Contains("InitialGroupMembership")) {
        Write-DebugMessage "Invoke-Onboarding: InitialGroupMembership"
        $defaultGroups = $Config.UserCreationDefaults["InitialGroupMembership"].Split(";") | ForEach-Object { $_.Trim() }
        foreach ($g in $defaultGroups) {
            if ($g) {
                try {
                    # Check if user is already a member before adding
                    $isMember = $false
                    try {
                        $groupMembers = Get-ADGroupMember -Identity $g -ErrorAction Stop
                        $isMember = ($groupMembers | Where-Object { $_.SamAccountName -eq $SamAccountName }) -ne $null
                    } catch {
                        # Group might not exist or other error
                        Write-DebugMessage "Error checking membership in group '$g': $($_.Exception.Message)"
                    }
                    
                    if (-not $isMember) {
                        Add-ADGroupMember -Identity $g -Members $SamAccountName -ErrorAction Stop
                        Write-DebugMessage "Added user to default group '$g'."
                    } else {
                        Write-DebugMessage "User is already a member of group '$g' - skipping."
                    }
                }
                catch {
                    Write-Warning "Error adding default group '$g': $($_.Exception.Message)"
                }
            }
        }
    }

    # [16.1.10.4 - License]
    # Fix for License group assignment 
    if (-not [string]::IsNullOrWhiteSpace($userData.License) -and $userData.License -notmatch "^(?i:none)$") {
        $licenseKey = "MS365_" + $userData.License.Trim().ToUpper()
        Write-DebugMessage "Processing license group with key: $licenseKey"
        
        if ($Config.Contains("LicensesGroups")) {
            # Try direct key lookup
            if ($Config.LicensesGroups.Contains($licenseKey)) {
                $licenseGroup = $Config.LicensesGroups[$licenseKey]
                Write-DebugMessage "Found license group '$licenseGroup' via key lookup"
            } 
            # Try to see if the selected value itself is a group name
            elseif ($Config.LicensesGroups.Values -contains $userData.License) {
                $licenseGroup = $userData.License
                Write-DebugMessage "Using direct license value as group name: $licenseGroup"
            }
            # Try without MS365_ prefix if key not found
            elseif ($Config.LicensesGroups.Contains($userData.License.Trim().ToUpper())) {
                $licenseGroup = $Config.LicensesGroups[$userData.License.Trim().ToUpper()]
                Write-DebugMessage "Found license group '$licenseGroup' via simplified key lookup"
            }
            else {
                Write-Warning "License key '$licenseKey' not found in [LicensesGroups]."
                $licenseGroup = $null
            }
            
            if ($licenseGroup) {
                try {
                    Add-ADGroupMember -Identity $licenseGroup -Members $SamAccountName -ErrorAction Stop
                    Write-DebugMessage "Added user to license group '$licenseGroup'."
                }
                catch {
                    Write-Warning "Error adding license group '$licenseGroup': $($_.Exception.Message)"
                }
            }
        }
        else {
            Write-Warning "[LicensesGroups] section not found in configuration."
        }
    }
   
    # [16.1.10.5 - MS365 AD Sync]
    if ($Config.Contains("ActivateUserMS365ADSync") -and $Config.ActivateUserMS365ADSync.Contains("ADSync") -and $Config.ActivateUserMS365ADSync["ADSync"] -eq "1") {
        $syncGroup = $Config.ActivateUserMS365ADSync["ADSyncADGroup"]
        if ($syncGroup) {
            try {
                Add-ADGroupMember -Identity $syncGroup -Members $SamAccountName -ErrorAction Stop
                Write-DebugMessage "Added user to AD-Sync group '$syncGroup'."
            }
            catch {
                Write-Warning "Error adding AD-Sync group '$syncGroup': $($_.Exception.Message)"
            }
        }
    }

    # [16.1.11 - Logging]
    Write-DebugMessage "Invoke-Onboarding: Logging to file"
    try {
        Write-LogMessage -Message ("ONBOARDING PERFORMED BY: " + ([System.Security.Principal.WindowsIdentity]::GetCurrent().Name) +
            ", SamAccountName: $SamAccountName, Display Name: '$($userData.DisplayName)', UPN: '$UPN', Location: '$($userData.Location)', Company: '$companyDisplay', License: '$($userData.License)', Password: '$mainPW', External: $($userData.External)") -Level Info
        Write-DebugMessage "Log written"
    }
    catch {
        Write-Warning "Error logging: $($_.Exception.Message)"
    }

    # [16.1.11.1 - Initialize variables for reports]
    $reportBranding = $global:Config["Report"]
    $reportPath = $reportBranding["ReportPath"]
    if (-not (Test-Path $reportPath)) {
        New-Item -ItemType Directory -Path $reportPath -Force | Out-Null
    }

    $finalReportFooter = $reportBranding["ReportFooter"]

    # [16.1.12 - Create Reports (HTML + TXT)]
    Write-DebugMessage "Invoke-Onboarding: Creating HTML and/or TXT reports"
    try {
        if (-not (Test-Path $reportPath)) {
            New-Item -ItemType Directory -Path $reportPath -Force | Out-Null
        }

        # [16.1.12.1 - Load the HTML template]
        $htmlFile = Join-Path $reportPath "$SamAccountName.html"
        
        # Check if TemplatePathHTML is defined and exists
        if (-not $reportBranding.Contains("TemplatePathHTML")) {
            Throw "Error: TemplatePathHTML is not defined in the [Report] section of the INI file."
        }
        
        $htmlTemplatePath = $reportBranding["TemplatePathHTML"]
        if (-not (Test-Path $htmlTemplatePath)) {
            Throw "Error: HTML template file not found at specified path: $htmlTemplatePath"
        }
        
        Write-DebugMessage "Reading HTML template from: $htmlTemplatePath"
        $htmlTemplate = Get-Content -Path $htmlTemplatePath -Raw -ErrorAction Stop
        if ([string]::IsNullOrWhiteSpace($htmlTemplate)) {
            Throw "Error: HTML template file is empty."
        }

        # [16.1.12.2 - Possibly build Websites HTML block if you want to show links in HTML]
        $websitesHTML = ""
        if ($Config.Contains("Websites")) {
            $websitesHTML += "<ul>`r`n"
            foreach ($key in $Config.Websites.Keys) {
                if ($key -match '^EmployeeLink\d+$') {
                    $line = $Config.Websites[$key]
                    # The format is "Label | URL | Description", or "Label;URL;Description"
                    if ($line -match '\|') {
                        $parts = $line -split '\|'
                    }
                    else {
                        $parts = $line -split ';'
                    }
                    if ($parts.Count -eq 3) {
                        $lbl = $parts[0].Trim()
                        $url = $parts[1].Trim()
                        $desc= $parts[2].Trim()
                        $websitesHTML += "<li><a href='$url' target='_blank'>$lbl</a> - $desc</li>`r`n"
                    }
                }
            }
            $websitesHTML += "</ul>`r`n"
        }

        # [16.1.12.3 - Replace placeholders in the HTML]
        $logoTag = ""
        if ($userData.CustomLogo -and (Test-Path $userData.CustomLogo)) {
            $logoTag = "<img src='$($userData.CustomLogo)' alt='Logo' />"
        }

        # Fix inline if expressions by assigning to variables first
        $description = if ($null -ne $userData.Description) { $userData.Description } else { "" }
        $officeRoom = if ($null -ne $userData.OfficeRoom) { $userData.OfficeRoom } else { "" }
        $phoneNumber = if ($null -ne $userData.PhoneNumber) { $userData.PhoneNumber } else { "" }
        $mobileNumber = if ($null -ne $userData.MobileNumber) { $userData.MobileNumber } else { "" }
        $position = if ($null -ne $userData.Position) { $userData.Position } else { "" }
        $departmentField = if ($null -ne $userData.DepartmentField) { $userData.DepartmentField } else { "" }
        $ablaufdatum = if ($null -ne $userData.Ablaufdatum) { $userData.Ablaufdatum } else { "" }

        $htmlContent = $htmlTemplate `
            -replace "{{ReportTitle}}",    ([string]$reportTitle) `
            -replace "{{Admin}}",          $env:USERNAME `
            -replace "{{ReportDate}}",     (Get-Date -Format "yyyy-MM-dd") `
            -replace "{{Vorname}}",        $userData.FirstName `
            -replace "{{Nachname}}",       $userData.LastName `
            -replace "{{DisplayName}}",    $userData.DisplayName `
            -replace "{{Description}}",    $description `
            -replace "{{Buero}}",          $officeRoom `
            -replace "{{Standort}}",       $userData.Location `
            -replace "{{Rufnummer}}",      $phoneNumber `
            -replace "{{Mobil}}",          $mobileNumber `
            -replace "{{Position}}",       $position `
            -replace "{{Abteilung}}",      $departmentField `
            -replace "{{Ablaufdatum}}",    $ablaufdatum `
            -replace "{{LoginName}}",      $SamAccountName `
            -replace "{{Passwort}}",       $mainPW `
            -replace "{{MailAddress}}",    $email `
            -replace "{{UPN}}",            $UPN `
            -replace "{{UPNFormat}}",      $userData.UPNFormat `
            -replace "{{Enabled}}",        (-not $userData.AccountDisabled) `
            -replace "{{External}}",       $userData.External `
            -replace "{{MailSuffix}}",     $mailSuffix `
            -replace "{{License}}",        $selectedLicense `
            -replace "{{ProxyMail}}",      $userData.setProxyMailAddress `
            -replace "{{TL}}",             $userData.TL `
            -replace "{{AL}}",             $userData.AL `
            -replace "{{TLGroup}}",        $selectedTLGroup `
            -replace "{{ADGroupsSelected}}", $ADGroupsSelected `
            -replace "{{DefaultOU}}",      $selectedOU `
            -replace "{{HomeDirectory}}",  $userData.HomeDirectory `
            -replace "{{ProfilePath}}",    $userData.ProfilePath `
            -replace "{{LoginScript}}",    $userData.ScriptPath `
            -replace "{{CompanyName}}",    $userData.CompanyName `
            -replace "{{CompanyStreet}}",  $userData.CompanyStreet `
            -replace "{{CompanyZIP}}",     $userData.CompanyPLZ `
            -replace "{{CompanyCity}}",    $userData.CompanyOrt `
            -replace "{{CompanyDomain}}",  $userData.CompanyDomain `
            -replace "{{CompanyPhone}}",   $userData.CompanyTelefon `
            -replace "{{CompanyCountry}}", $userData.CompanyCountry `
            -replace "{{WebsitesHTML}}",   $websitesHTML `
            -replace "{{ReportFooter}}",   $finalReportFooter `
            -replace "{{LogoTag}}",        $logoTag `
            -replace "{{CustomPWLabel1}}", $pwLabel1 `
            -replace "{{CustomPWLabel2}}", $pwLabel2 `
            -replace "{{CustomPWLabel3}}", $pwLabel3 `
            -replace "{{CustomPWLabel4}}", $pwLabel4 `
            -replace "{{CustomPWLabel5}}", $pwLabel5 `
            -replace "{{CustomPW1}}",      $CustomPW1 `
            -replace "{{CustomPW2}}",      $CustomPW2 `
            -replace "{{CustomPW3}}",      $CustomPW3 `
            -replace "{{CustomPW4}}",      $CustomPW4 `
            -replace "{{CustomPW5}}",      $CustomPW5

        Set-Content -Path $htmlFile -Value $htmlContent -Encoding UTF8
        Write-DebugMessage "HTML report created: $htmlFile"

        # [16.1.12.4 - Create the TXT if userData.OutputTXT or if the INI says to]
        $shouldCreateTXT = $userData.OutputTXT -or ($Config.Report["UserOnboardingCreateTXT"] -eq '1')
        if ($shouldCreateTXT) {
            # Check if TemplatePathTXT is defined and exists
            if (-not $reportBranding.Contains("TemplatePathTXT")) {
                Throw "Error: TemplatePathTXT is not defined in the [Report] section of the INI file."
            }
            
            $txtTemplatePath = $reportBranding["TemplatePathTXT"]
            if (-not (Test-Path $txtTemplatePath)) {
                Throw "Error: TXT template file not found at specified path: $txtTemplatePath"
            }
            
            $txtFile = Join-Path $reportPath "$SamAccountName.txt"

            Write-DebugMessage "Reading txtTemplate from: $txtTemplatePath"
            $txtContent = Get-Content -Path $txtTemplatePath -Raw -ErrorAction Stop
            if ([string]::IsNullOrWhiteSpace($txtContent)) {
                Throw "Error: TXT template file is empty."
            }

            # [16.1.12.4.2 - Replace placeholders in the template]
            # Fix placeholder replacements to avoid parsing issues
            $defaultOU = $Config.ADUserDefaults["DefaultOU"]
            #$txtContent = $txtContent -replace '\$Config\.ADUserDefaults\["DefaultOU"\]', $defaultOU
            #$txtContent = $txtContent -replace "\`$SamAccountName", $SamAccountName
            #$txtContent = $txtContent -replace "\`$UserPW", $mainPW
            #$txtContent = $txtContent -replace "\`$userData\.AccountDisabled", $userData.AccountDisabled
            # Safer pattern replacement approach
            #$txtContent = [regex]::Replace($txtContent, "\$\(.*?\)", "")
            
            # Replace PowerShell variable placeholders in the template
            $txtContent = $txtContent `
                -replace '\$Config\.ADUserDefaults\["DefaultOU"\]', $selectedOU `
                -replace '\$\(\$userData\.FirstName\)', $userData.FirstName `
                -replace '\$\(\$userData\.LastName\)', $userData.LastName `
                -replace '\$\(\$userData\.External\)', $($userData.External) `
                -replace '\$\(\$userData\.DisplayName\)', $userData.DisplayName `
                -replace '\$\(\$userData\.SamAccountName\)', $SamAccountName `
                -replace '\$\(\$userData\.UserPrincipalName\)', $UPN `
                -replace '\$\(\$userData\.UPNFormat\)', $($userData.UPNFormat) `
                -replace '\$\(\$userData\.PasswordMode\)', $passwordMode `
                -replace '\$\(\$mainPW\)', $mainPW `
                -replace '\$\(\$userData\.Ablaufdatum\)', $ablaufdatum `
                -replace '\$\(\$userData\.AccountDisabled\)', $($userData.AccountDisabled) `
                -replace '\$\(\$userData\.EmailAddress\)', $email `
                -replace '\$\(\$userData\.MailSuffix\)', $mailSuffix `
                -replace '\$\(\$userData\.setProxyMailAddress\)', $($userData.setProxyMailAddress) `
                -replace '\$\(\$userData\.Position\)', $position `
                -replace '\$\(\$userData\.DepartmentField\)', $departmentField `
                -replace '\$\(\$userData\.Division\)', $($userData.Division) `
                -replace '\$\(\$userData\.Manager\)', $($userData.Manager) `
                -replace '\$\(\$userData\.Company\)', $($userData.CompanyName) `
                -replace '\$\(\$userData\.CompanyTelefon\)', $($userData.CompanyTelefon) `
                -replace '\$\(\$userData\.CompanyStrasse\)', $($userData.CompanyStrasse) `
                -replace '\$\(\$userData\.CompanyPLZ\)', $($userData.CompanyPLZ) `
                -replace '\$\(\$userData\.CompanyOrt\)', $($userData.CompanyOrt) `
                -replace '\$\(\$userData\.CompanyCountry\)', $($userData.CompanyCountry) `
                -replace '\$\(\$userData\.PhoneNumber\)', $phoneNumber `
                -replace '\$\(\$userData\.MobileNumber\)', $mobileNumber `
                -replace '\$\(\$userData\.IPPhone\)', $($userData.IPPhone) `
                -replace '\$\(\$userData\.OfficeRoom\)', $officeRoom `
                -replace '\$\(\$userData\.StreetAddress\)', $($userData.StreetAddress) `
                -replace '\$\(\$userData\.City\)', $($userData.City) `
                -replace '\$\(\$userData\.PostalCode\)', $($userData.PostalCode) `
                -replace '\$\(\$userData\.State\)', $($userData.State) `
                -replace '\$\(\$userData\.Country\)', $($userData.Country) `
                -replace '\$\(\$userData\.ADGroupsSelected\)', $ADGroupsSelected `
                -replace '\$\(\$userData\.TL\)', $($userData.TL) `
                -replace '\$\(\$userData\.TLGroup\)', $selectedTLGroup `
                -replace '\$\(\$userData\.AL\)', $($userData.AL) `
                -replace '\$\(\$userData\.License\)', $selectedLicense `
                -replace '\$pwLabel1', $pwLabel1 `
                -replace '\$pwLabel2', $pwLabel2 `
                -replace '\$pwLabel3', $pwLabel3 `
                -replace '\$pwLabel4', $pwLabel4 `
                -replace '\$pwLabel5', $pwLabel5 `
                -replace '\$CustomPW1', $CustomPW1 `
                -replace '\$CustomPW2', $CustomPW2 `
                -replace '\$CustomPW3', $CustomPW3 `
                -replace '\$CustomPW4', $CustomPW4 `
                -replace '\$CustomPW5', $CustomPW5

            Write-DebugMessage "Writing final TXT to: $txtFile"
            Out-File -FilePath $txtFile -InputObject $txtContent -Encoding UTF8
            Write-DebugMessage "TXT report created: $txtFile"
        }
    }
    catch {
        Write-Warning "Error creating reports: $($_.Exception.Message)"
        Throw "Failed to create reports: $($_.Exception.Message)"
    }

    Write-DebugMessage "Invoke-Onboarding: Returning final result object"
    # [16.1.13 - Return a small object with final SamAccountName, UPN, and PW if needed]
    Write-DebugMessage "Onboarding process completed successfully."
    return [ordered]@{
        SamAccountName = $SamAccountName
        UPN            = $UPN
        Password       = $mainPW
    }
} 
Write-Debug ("userData: " + ( $userData | Out-String ))
#endregion

#region [Region 17 | DROPDOWN UTILITY]
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
            $DropDownControl.ItemsSource = $null
            $DropDownControl.Items.Clear()
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
            $DropDownControl.ItemsSource = $items
        }
        Write-Verbose "Data binding set for '$DataType' dropdown."
    
        Write-DebugMessage "Set-DropDownValues: Setting default value if available"
        if ($defaultValue) {
            $DropDownControl.SelectedItem = $defaultValue
            Write-Verbose "Default value for '$DataType' set to '$defaultValue'."
        }
        elseif ($items.Count -gt 0) {
            try { $DropDownControl.SelectedIndex = 0 } catch { }
        }
    }
    catch {
        Write-Error "Error populating dropdown '$DataType': $($_.Exception.Message)"
    }
}
#endregion

Write-DebugMessage "Process-ADGroups"

#region [Region 18 | AD GROUP CREATION]
# Function for handling AD group creation (placeholder)
Write-DebugMessage "Processing AD groups."
function Process-ADGroups {
    param(
        [Parameter(Mandatory=$true)] $GroupData,
        [Parameter(Mandatory=$true)] [hashtable]$Config
    )
    Write-Log "Process-ADGroups: Group creation is being performed..." "DEBUG"
    # Here the code for group creation can be added and optimized.
    return $true
}
#endregion

Write-DebugMessage "Get-ADData"

#region [Region 19 | AD DATA RETRIEVAL]
# Function to retrieve and process data from Active Directory
Write-DebugMessage "Retrieving AD data."
function Get-ADData {
    param(
        [Parameter(Mandatory=$false)]
        $Window
    )
    
    try {
        # Retrieve all OUs and convert to a sorted array
        $allOUs = Get-ADOrganizationalUnit -Filter * -Properties Name, DistinguishedName |
            ForEach-Object {
                [PSCustomObject]@{
                    Name = $_.Name
                    DN   = $_.DistinguishedName
                }
            } | Sort-Object Name

        # Debug output for OUs if DebugMode=1
        if ($global:Config.Logging.DebugMode -eq "1") {
            Write-Log " All OUs loaded:" "DEBUG"
            $allOUs | Format-Table | Out-String | Write-Log
        }
        Write-DebugMessage "Get-ADData: Retrieving all users"
        # Retrieve all users (here the first 200 users sorted by DisplayName are displayed)
        $allUsers = Get-ADUser -Filter * -Properties DisplayName, SamAccountName |
            ForEach-Object {
                [PSCustomObject]@{
                    DisplayName    = if ($_.DisplayName) { $_.DisplayName } else { $_.SamAccountName }
                    SamAccountName = $_.SamAccountName
                }
            } | Sort-Object DisplayName | Select-Object

        # Debug output for users if DebugMode=1
        if ($global:Config.Logging.DebugMode -eq "1") {
            Write-Log " User list loaded:" "DEBUG"
            $allUsers | Format-Table | Out-String | Write-Log
        }
        Write-DebugMessage "Get-ADData: comboBoxOU and listBoxUsers"
        
        # Find the ComboBox and ListBox in the Window if provided
        if ($Window) {
            $comboBoxOU = $Window.FindName("cmbOU") 
            $listBoxUsers = $Window.FindName("lstUsers")
            
            if ($comboBoxOU) {
                $comboBoxOU.ItemsSource = $null
                $comboBoxOU.Items.Clear()
                $comboBoxOU.ItemsSource = $allOUs
                $comboBoxOU.DisplayMemberPath = "Name"
                $comboBoxOU.SelectedValuePath = "DN"
                if ($allOUs.Count -gt 0) {
                    $comboBoxOU.SelectedIndex = 0
                }
            }

            if ($listBoxUsers) {
                $listBoxUsers.ItemsSource = $null
                $listBoxUsers.Items.Clear()
                $listBoxUsers.ItemsSource = $allUsers
                $listBoxUsers.DisplayMemberPath = "DisplayName"
                $listBoxUsers.SelectedValuePath = "SamAccountName"
            }
        }
        
        # Return the data for further use
        return @{
            OUs = $allOUs
            Users = $allUsers
        }
    }
    catch {
        Write-DebugMessage "Error in Get-ADData: $($_.Exception.Message)"
        [System.Windows.MessageBox]::Show("AD connection error: $($_.Exception.Message)")
        return $null
    }
}
#endregion

Write-DebugMessage "Open-INIEditor"

#region [Region 20 | INI EDITOR]
# Function to open the INI configuration editor
Write-DebugMessage "Opening INI editor."
function Open-INIEditor {
    try {
        if ($global:Config.Logging.DebugMode -eq "1") {
            Write-Log " INI Editor is starting." "DEBUG"
        }
        Write-DebugMessage "INI Editor has started."
        Write-LogMessage -Message "INI Editor has started." -Level "Info"
        [System.Windows.MessageBox]::Show("INI Editor (Settings) – Implement functionality here.", "Settings")
    }
    catch {
        Throw "Error opening INI Editor: $_"
    }
}
#endregion

Write-DebugMessage "Defining GUI implementation."

#region [Region 21 | GUI IMPLEMENTATION]
# Handles the entire GUI implementation and event handlers

Write-DebugMessage "GUI Collecting selected groups from the AD groups panel"

#region [Region 21.1 | AD GROUPS SELECTION]
# Collects selected groups from the AD groups panel
Write-DebugMessage "Collecting selected AD groups from panel."
function Get-SelectedADGroups {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [System.Windows.Controls.ItemsControl]$Panel,
        
        [Parameter(Mandatory=$false)]
        [string]$Separator = ";"
    )
    
    # Use System.Collections.ArrayList to avoid pipeline output when adding items
    [System.Collections.ArrayList]$selectedGroups = @()
    
    if ($null -eq $Panel) {
        Write-DebugMessage "Error: Panel is null"
        return "NONE"
    }
    
    if ($Panel.Items.Count -eq 0) {
        Write-DebugMessage "Panel contains no items to select from"
        return "NONE"
    }
    
    Write-DebugMessage "Processing $($Panel.Items.Count) items in AD groups panel"
    
    # Get all checkboxes regardless of nesting
    function Find-CheckBoxes {
        param (
            [Parameter(Mandatory=$true)]
            [System.Windows.DependencyObject]$Parent
        )
        
        $checkBoxes = @()
        
        for ($i = 0; $i -lt [System.Windows.Media.VisualTreeHelper]::GetChildrenCount($Parent); $i++) {
            $child = [System.Windows.Media.VisualTreeHelper]::GetChild($Parent, $i)
            
            if ($child -is [System.Windows.Controls.CheckBox]) {
                $checkBoxes += $child
            } else {
                # Recursively search children
                $checkBoxes += Find-CheckBoxes -Parent $child
            }
        }
        
        return $checkBoxes
    }
    
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
        $chkSetProxyMail = $onboardingTab.FindName("chkSetProxyMail")
        $comboBoxDisplayTemplate = $onboardingTab.FindName("cmbDisplayTemplate")
        $comboBoxMailSuffix = $onboardingTab.FindName("cmbMailSuffix")
        $comboBoxLicense = $onboardingTab.FindName("cmbLicense")
        $comboBoxOU = $onboardingTab.FindName("cmbOU")
        $comboBoxTLGroup = $onboardingTab.FindName("cmbTLGroup")
        $radAccountEnabled = $onboardingTab.FindName("radAccountEnabled")
        $radAccountDisabled = $onboardingTab.FindName("radAccountDisabled")

        Write-DebugMessage "GUI elements successfully loaded."

        # **Validate mandatory fields**
        if (-not $txtFirstName -or [string]::IsNullOrWhiteSpace($txtFirstName.Text)) {
            Write-DebugMessage "ERROR: First name missing!"
            [System.Windows.MessageBox]::Show("Error: First name cannot be empty!", "Validation Error", 'OK', 'Error')
            return
        }
        if (-not $txtLastName -or [string]::IsNullOrWhiteSpace($txtLastName.Text)) {
            Write-DebugMessage "ERROR: Last name missing!"
            [System.Windows.MessageBox]::Show("Error: Last name cannot be empty!", "Validation Error", 'OK', 'Error')
            return
        }

        # Validate OU selection - critical for user creation
        if (-not $comboBoxOU -or -not $comboBoxOU.SelectedValue) {
            # Check if there's a default in the configuration
            if (-not ($global:Config.Contains("ADUserDefaults") -and 
                     $global:Config.ADUserDefaults.Contains("DefaultOU") -and 
                     -not [string]::IsNullOrWhiteSpace($global:Config.ADUserDefaults["DefaultOU"]))) {
                Write-DebugMessage "ERROR: No OU selected and no default found!"
                [System.Windows.MessageBox]::Show("Error: Please select an Organizational Unit (OU) for the user!", "Validation Error", 'OK', 'Error')
                return
            }
        }

        # Validate TL Group selection if TL is checked
        if ($chkTL -and $chkTL.IsChecked -eq $true) {
            if (-not $comboBoxTLGroup -or -not $comboBoxTLGroup.SelectedValue) {
                Write-DebugMessage "ERROR: TL is checked but no TL Group selected!"
                [System.Windows.MessageBox]::Show("Error: Please select a Team Leader Group since the TL option is checked!", "Validation Error", 'OK', 'Error')
                return
            }
        }

        # Validate termination date format if provided
        if ($txtTermination -and -not [string]::IsNullOrWhiteSpace($txtTermination.Text)) {
            try {
                [DateTime]::Parse($txtTermination.Text)
            }
            catch {
                Write-DebugMessage "ERROR: Invalid termination date format!"
                [System.Windows.MessageBox]::Show("Error: The termination date format is invalid. Please use a valid date format.", "Validation Error", 'OK', 'Error')
                return
            }
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
        $accountDisabled = if ($radAccountDisabled -and $radAccountDisabled.IsChecked) { 
            $true 
        } elseif ($chkAccountDisabled -and $chkAccountDisabled.IsChecked) {
            $true
        } else { 
            $false 
        }

        # Determine password mode (generated or fixed)
        $radGeneratedPW = $onboardingTab.FindName("radGeneratedPW")
        $radFixedPW = $onboardingTab.FindName("radFixedPW")
        $txtFixedPassword = $onboardingTab.FindName("txtFixedPassword")
        
        $passwordMode = 1 # Default: Generate password
        $password = ""
        
        if ($radFixedPW -and $radFixedPW.IsChecked) {
            $passwordMode = 2 # Fixed password
            if ($txtFixedPassword -and -not [string]::IsNullOrWhiteSpace($txtFixedPassword.Text)) {
                $password = $txtFixedPassword.Text.Trim()
            }
            else {
                # If fixed password is selected but no password provided, show error
                Write-DebugMessage "ERROR: Fixed password selected but no password entered!"
                [System.Windows.MessageBox]::Show("Error: You selected to use a fixed password but did not enter one.", "Validation Error", 'OK', 'Error')
                return
            }
        }

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
        Write-DebugMessage "Selected DisplayTemplate Value: $($comboBoxDisplayTemplate.SelectedValue)"
    
        # Erst Format-String ermitteln
        $displayNameFormat = Get-DisplayNameFormat -IniPath $INIPath -SelectedTemplate $comboBoxDisplayTemplate.SelectedValue

        # Dann DisplayName formatieren
        $formattedDisplayName = Format-DisplayName -Format $displayNameFormat -FirstName $firstName -LastName $lastName
        Write-DebugMessage "Formatted DisplayName: $formattedDisplayName"

        $global:userData = [PSCustomObject]@{
            OU               = $selectedOU
            FirstName        = $firstName
            LastName         = $lastName
            DisplayName      = $formattedDisplayName
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
            setProxyMailAddress = if ($chkSetProxyMail) { $chkSetProxyMail.IsChecked } else { $false }
            UPNFormat        = if ($comboBoxDisplayTemplate -and $comboBoxDisplayTemplate.SelectedValue) { 
                # First priority: Use the user-selected template from dropdown
                $comboBoxDisplayTemplate.SelectedValue.ToString().Trim() 
            } elseif ($global:Config.Contains("DisplayNameUPNTemplates") -and 
                      $global:Config.DisplayNameUPNTemplates.Contains("DefaultUserPrincipalNameFormat") -and
                      -not [string]::IsNullOrWhiteSpace($global:Config.DisplayNameUPNTemplates["DefaultUserPrincipalNameFormat"])) { 
                # Second priority: Use DefaultUserPrincipalNameFormat from INI
                Write-DebugMessage "Using UPN format from INI: $($global:Config.DisplayNameUPNTemplates['DefaultUserPrincipalNameFormat'])"
                $global:Config.DisplayNameUPNTemplates["DefaultUserPrincipalNameFormat"] 
            } else { 
                # Fallback: Use standard format
                Write-DebugMessage "No UPN format in dropdown or INI, using default: FIRSTNAME.LASTNAME"
                "FIRSTNAME.LASTNAME" 
            }
            AccountDisabled  = $accountDisabled
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
            PasswordMode     = $passwordMode
            Password         = $password
        }        

        Write-DebugMessage "userData object created -> `n$($global:userData | Out-String)"
        Write-DebugMessage "Starting Invoke-Onboarding..."

        try {
            # Update UI to show we're processing
            $progressBar = $window.FindName("progressBar")
            $lblStatus = $window.FindName("lblStatus")
            
            if ($progressBar) {
                $progressBar.Value = 25
            }
            if ($lblStatus) {
                $lblStatus.Text = "Creating user account..."
            }

            $global:result = Invoke-Onboarding -userData $global:userData -Config $global:Config
            if (-not $global:result) {
                throw "Invoke-Onboarding did not return a result!"
            }

            # Update progress bar
            if ($progressBar) {
                $progressBar.Value = 100
            }
            if ($lblStatus) {
                $lblStatus.Text = "User created successfully"
            }

            Write-DebugMessage "Invoke-Onboarding completed."
            [System.Windows.MessageBox]::Show("Onboarding successfully completed.`nSamAccountName: $($global:result.SamAccountName)`nUPN: $($global:result.UPN)`nPassword: $($global:result.Password)", "Success")
        } catch {
            # Reset progress on error
            if ($progressBar) {
                $progressBar.Value = 0
            }
            if ($lblStatus) {
                $lblStatus.Text = "Error: Failed to create user"
            }

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
                Write-Log "$global:result is NULL or empty" "DEBUG"
                return
            }
            
            Write-Log "$global:result successfully loaded: $($global:result | Out-String)" "DEBUG"
            
            $SamAccountName = $global:result.SamAccountName

            if (-not ($global:Config.Keys -contains "Report")) {
                [System.Windows.MessageBox]::Show("Error: The section [Report] is missing in the INI file.", "Error", 'OK', 'Error')
                return
            }

            $reportBranding = $global:Config["Report"]
            if (-not $reportBranding -or -not $reportBranding.Contains("ReportPath") -or [string]::IsNullOrWhiteSpace($reportBranding["ReportPath"])) {
                [System.Windows.MessageBox]::Show("ReportPath is missing or empty in [Report]. PDF cannot be created.", "Error", 'OK', 'Error')
                return
            }         

            $htmlReportPath = Join-Path $reportBranding["ReportPath"] "$SamAccountName.html"
            $pdfReportPath  = Join-Path $reportBranding["ReportPath"] "$SamAccountName.pdf"

            $wkhtmltopdfPath = $reportBranding["wkhtmltopdfPath"]
            if (-not (Test-Path $wkhtmltopdfPath)) {
                [System.Windows.MessageBox]::Show("wkhtmltopdf.exe not found! Please check: $wkhtmltopdfPath", "Error", 'OK', 'Error')
                return
            }          

            if (-not (Test-Path $htmlReportPath)) {
                [System.Windows.MessageBox]::Show("HTML report not found: $htmlReportPath`nThe PDF cannot be created.", "Error", 'OK', 'Error')
                return
            }

            $pdfScript = Join-Path $PSScriptRoot "PDFCreator.ps1"
            if (-not (Test-Path $pdfScript)) {
                [System.Windows.MessageBox]::Show("PDF script 'easyOnboarding_PDFCreator.ps1' not found!", "Error", 'OK', 'Error')
                return
            }

            $arguments = @(
                "-NoProfile",
                "-ExecutionPolicy", "Bypass",
                "-File", "`"$pdfScript`"",
                "-htmlFile", "`"$htmlReportPath`"",
                "-pdfFile", "`"$pdfReportPath`"",
                "-wkhtmltopdfPath", "`"$wkhtmltopdfPath`""
            )
            Start-Process -FilePath "powershell.exe" -ArgumentList $arguments -NoNewWindow -Wait

            [System.Windows.MessageBox]::Show("PDF successfully created: $pdfReportPath", "PDF Creation", 'OK', 'Information')

        } catch {
            [System.Windows.MessageBox]::Show("Error creating PDF: $($_.Exception.Message)", "Error", 'OK', 'Error')
        }
    })
}
#endregion

$btnProxyMail = $window.FindName("btnProxyMail")
if ($btnProxyMail) {
    Write-DebugMessage "Registering event handler for ProxyMail button"
    $btnProxyMail.Add_Click({
        try {
            # Get the first name and last name from the GUI
            $txtFirstName = $onboardingTab.FindName("txtFirstName")
            $txtLastName = $onboardingTab.FindName("txtLastName")

            # Validate that the first name and last name are not empty
            if ([string]::IsNullOrWhiteSpace($txtFirstName.Text) -or [string]::IsNullOrWhiteSpace($txtLastName.Text)) {
                Write-Warning "First name or last name is null or empty. Cannot set proxy addresses."
                [System.Windows.MessageBox]::Show("First name or last name is null or empty. Cannot set proxy addresses.", "Warning", "OK", "Warning")
                return
            }

            # Get the email address from the GUI
            $txtEmail = $onboardingTab.FindName("txtEmail")
            $emailAddress = $txtEmail.Text.Trim()

            # Validate that the email address is not empty
            if ([string]::IsNullOrWhiteSpace($emailAddress)) {
                Write-Warning "Email address is null or empty. Cannot set proxy addresses."
                [System.Windows.MessageBox]::Show("Email address is null or empty. Cannot set proxy addresses.", "Warning", "OK", "Warning")
                return
            }
            
            # Get the mail suffix from the GUI
            $comboBoxSuffix = $onboardingTab.FindName("cmbSuffix")
            $mailSuffix = if ($comboBoxSuffix -and $comboBoxSuffix.SelectedValue) {
                $comboBoxSuffix.SelectedValue.ToString()
            } else {
                ""
            }

            # Construct the username from the first letter of the first name and the last name
            $username = ($txtFirstName.Text.Substring(0, 1) + $txtLastName.Text).ToLower()

            # Construct the primary SMTP address from the email address and mail suffix
            $smtpAddress = "SMTP:" + $emailAddress + $mailSuffix

            # Get the domains from the config
            $companySection = $global:Config.Company
            $ms365Domain = $companySection["CompanyMS365Domain"]
            $companyMailDomain = $companySection["CompanyMailDomain"]

            # Construct the smtp alias addresses
            $smtpAliasAddress1 = "smtp:" + $emailAddress + $ms365Domain
            $smtpAliasAddress2 = "smtp:" + $emailAddress + $companyMailDomain

            # Create an array of proxy addresses
            $proxyAddresses = @($smtpAddress, $smtpAliasAddress1, $smtpAliasAddress2)

            # Call the Set-ADUser function with the values from the GUI
            Set-ADUser -Identity $username -Add @{proxyAddresses = $proxyAddresses}
            
            [System.Windows.MessageBox]::Show(
                "Proxy addresses set successfully for user: $username`n`n" + 
                "Primary: $smtpAddress`n" +
                "Alias 1: $smtpAliasAddress1`n" +
                "Alias 2: $smtpAliasAddress2", 
                "Success", "OK", "Information")
            
            Write-DebugMessage "Set proxy addresses for ${username}: Primary=${smtpAddress}, Alias1=${smtpAliasAddress1}, Alias2=${smtpAliasAddress2}"
        }
        catch {
            Write-DebugMessage "Error executing Set-ADUser: $($_.Exception.Message)"
            [System.Windows.MessageBox]::Show(
                "Error executing Set-ADUser: $($_.Exception.Message)",
                "Error", "OK", "Error")
        }
    })
}

Write-DebugMessage "TAB easyONBOARDING - BUTTON Info"

#region [Region 21.6 | INFO BUTTON]
# Implements the info button functionality
Write-DebugMessage "Setting up info button."
$btnInfo = $window.FindName("btnInfo")
if ($btnInfo) {
    Write-DebugMessage "TAB easyONBOARDING - BUTTON selected: Info"
    $btnInfo.Add_Click({
        $infoFilePath = Join-Path $PSScriptRoot "easyONBOARDING.txt"

        if (Test-Path $infoFilePath) {
            Start-Process -FilePath $infoFilePath
        } else {
            [System.Windows.MessageBox]::Show("The file easyONBOARDING.txt was not found!", "Error", 'OK', 'Error')
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
    Write-DebugMessage "BUTTON selected: Close"
    $btnClose.Add_Click({
        $window.Close()
    })
}
#endregion

Write-DebugMessage "GUI: TAB easyADUpdate loaded"

#region [Region 21.8.1 | CSV IMPORT FUNCTIONALITY]
# Implement CSV import and preview functionality
Write-DebugMessage "Setting up CSV import functionality."

# Get references to the buttons
$btnImportCSV = $window.FindName("btnImportCSV")
$btnClearPreview = $window.FindName("btnClearPreview")

if ($btnImportCSV) {
    Write-DebugMessage "Registering event handler for Import CSV button"
    $btnImportCSV.Add_Click({
        try {
            Write-DebugMessage "Import CSV button clicked"
            
            # Create and configure OpenFileDialog
            $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
            $openFileDialog.Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
            $openFileDialog.Title = "Select CSV file with user information"
            $openFileDialog.Multiselect = $false
            
            # Show the dialog and check if user clicked OK
            if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                $csvPath = $openFileDialog.FileName
                Write-DebugMessage "CSV file selected: $csvPath"
                
                # Check if file exists and is not empty
                if (-not (Test-Path $csvPath) -or (Get-Item $csvPath).Length -eq 0) {
                    [System.Windows.MessageBox]::Show("The selected file is empty or could not be found.", "Error", "OK", "Error")
                    return
                }

                # Attempt to import the CSV
                try {
                    $csvData = Import-Csv -Path $csvPath -ErrorAction Stop
                    Write-DebugMessage "Successfully imported CSV with $($csvData.Count) records"
                    
                    # Check if we have any data
                    if ($csvData.Count -eq 0) {
                        [System.Windows.MessageBox]::Show("The CSV file contains no data.", "Warning", "OK", "Warning")
                        return
                    }
                    
                    # Get first row or prompt for selection if multiple records
                    $selectedRow = $null
                    if ($csvData.Count -eq 1) {
                        $selectedRow = $csvData[0]
                    } else {
                        # Multiple records found - ask user what to do
                        $userChoice = [System.Windows.MessageBox]::Show(
                            "The CSV file contains $($csvData.Count) records. Do you want to use the first record?`n`nClick 'Yes' to use the first record, or 'No' to cancel.",
                            "Multiple Records Found",
                            "YesNo",
                            "Question")
                            
                        if ($userChoice -eq [System.Windows.MessageBoxResult]::Yes) {
                            $selectedRow = $csvData[0]
                        } else {
                            return
                        }
                    }
                    
                    # Check for expected column format
                    $expectedColumns = @("FirstName", "LastName", "Description", "OfficeRoom", 
                                    "PhoneNumber", "MobileNumber", "Position", "DepartmentField", 
                                    "EmailAddress", "Ablaufdatum", "External", "TL", "AL", "AccountDisabled", "ADGroup")
                    
                    $missingColumns = $expectedColumns | Where-Object { $_ -notin $selectedRow.PSObject.Properties.Name }
                    
                    if ($missingColumns.Count -gt 0) {
                        $warningMessage = "The CSV is missing the following expected columns: $($missingColumns -join ', ')`n`n" +
                                        "Expected format: $($expectedColumns -join ', ')"
                        
                        $userChoice = [System.Windows.MessageBox]::Show(
                            $warningMessage + "`n`nDo you want to continue with the import anyway?", 
                            "Format Mismatch", 
                            "YesNo", 
                            "Warning")
                        
                        if ($userChoice -eq [System.Windows.MessageBoxResult]::No) {
                            return
                        }
                    }
                    
                    # Log what properties the row data has to help with debugging
                    $rowProperties = $selectedRow.PSObject.Properties.Name -join ", "
                    Write-DebugMessage "CSV row properties: $rowProperties"
                    
                    # Get all form controls that will receive data
                    $txtFirstName = $window.FindName("txtFirstName")
                    $txtLastName = $window.FindName("txtLastName")
                    $txtDisplayName = $window.FindName("txtDisplayName")
                    $txtEmail = $window.FindName("txtEmail")
                    $txtOffice = $window.FindName("txtOffice")
                    $txtPhone = $window.FindName("txtPhone")
                    $txtMobile = $window.FindName("txtMobile")
                    $txtPosition = $window.FindName("txtPosition")
                    $txtDepartment = $window.FindName("txtDepartment")
                    $txtDescription = $window.FindName("txtDescription")
                    $txtTermination = $window.FindName("txtTermination")
                    $chkExternal = $window.FindName("chkExternal")
                    $chkTL = $window.FindName("chkTL")
                    $chkAL = $window.FindName("chkAL")
                    $radAccountEnabled = $window.FindName("radAccountEnabled")
                    $radAccountDisabled = $window.FindName("radAccountDisabled")
                    $icADGroups = $window.FindName("icADGroups") # Get the ItemsControl for AD Groups

                    Write-DebugMessage "GUI elements successfully loaded."
                    
                    # Clear all values first to prevent data mixing
                    if ($txtFirstName) { $txtFirstName.Text = "" }
                    if ($txtLastName) { $txtLastName.Text = "" }
                    if ($txtDisplayName) { $txtDisplayName.Text = "" }
                    if ($txtEmail) { $txtEmail.Text = "" }
                    if ($txtOffice) { $txtOffice.Text = "" }
                    if ($txtPhone) { $txtPhone.Text = "" }
                    if ($txtMobile) { $txtMobile.Text = "" }
                    if ($txtPosition) { $txtPosition.Text = "" }
                    if ($txtDepartment) { $txtDepartment.Text = "" }
                    if ($txtDescription) { $txtDescription.Text = "" }
                    if ($txtTermination) { $txtTermination.Text = "" }
                    if ($chkExternal) { $chkExternal.IsChecked = $false }
                    if ($chkTL) { $chkTL.IsChecked = $false }
                    if ($chkAL) { $chkAL.IsChecked = $false }

                    # Reset AD Group selections
                    if ($icADGroups -and $global:ADGroupItems) {
                        foreach ($item in $global:ADGroupItems) {
                            $item.IsChecked = $false
                        }
                    }
                    
                    # Populate form fields with CSV data
                    if ($selectedRow.PSObject.Properties.Name -contains "FirstName" -and $txtFirstName) {
                        $txtFirstName.Text = $selectedRow.FirstName
                        Write-DebugMessage "Set FirstName to: $($selectedRow.FirstName)"
                    }
                    
                    if ($selectedRow.PSObject.Properties.Name -contains "LastName" -and $txtLastName) {
                        $txtLastName.Text = $selectedRow.LastName
                        Write-DebugMessage "Set LastName to: $($selectedRow.LastName)"
                    }
                    
                    # Set DisplayName based on FirstName and LastName
                    if ($txtFirstName -and $txtLastName -and $txtDisplayName -and 
                        -not [string]::IsNullOrWhiteSpace($txtFirstName.Text) -and 
                        -not [string]::IsNullOrWhiteSpace($txtLastName.Text)) {
                        $txtDisplayName.Text = "$($txtFirstName.Text) $($txtLastName.Text)"
                        Write-DebugMessage "Set DisplayName to: $($txtDisplayName.Text)"
                    }
                    
                    if ($selectedRow.PSObject.Properties.Name -contains "EmailAddress" -and $txtEmail) {
                        $txtEmail.Text = $selectedRow.EmailAddress
                        Write-DebugMessage "Set Email to: $($selectedRow.EmailAddress)"
                    } elseif ($selectedRow.PSObject.Properties.Name -contains "Email" -and $txtEmail) {
                        $txtEmail.Text = $selectedRow.Email
                        Write-DebugMessage "Set Email to: $($selectedRow.Email)"
                    }
                    
                    if ($selectedRow.PSObject.Properties.Name -contains "OfficeRoom" -and $txtOffice) {
                        $txtOffice.Text = $selectedRow.OfficeRoom
                        Write-DebugMessage "Set Office to: $($selectedRow.OfficeRoom)"
                    }
                    
                    if ($selectedRow.PSObject.Properties.Name -contains "PhoneNumber" -and $txtPhone) {
                        $txtPhone.Text = $selectedRow.PhoneNumber
                        Write-DebugMessage "Set Phone to: $($selectedRow.PhoneNumber)"
                    }
                    
                    if ($selectedRow.PSObject.Properties.Name -contains "MobileNumber" -and $txtMobile) {
                        $txtMobile.Text = $selectedRow.MobileNumber
                        Write-DebugMessage "Set Mobile to: $($selectedRow.MobileNumber)"
                    }
                    
                    if ($selectedRow.PSObject.Properties.Name -contains "Position" -and $txtPosition) {
                        $txtPosition.Text = $selectedRow.Position
                        Write-DebugMessage "Set Position to: $($selectedRow.Position)"
                    }
                    
                    if ($selectedRow.PSObject.Properties.Name -contains "DepartmentField" -and $txtDepartment) {
                        $txtDepartment.Text = $selectedRow.DepartmentField
                        Write-DebugMessage "Set Department to: $($selectedRow.DepartmentField)"
                    }
                    
                    if ($selectedRow.PSObject.Properties.Name -contains "Description" -and $txtDescription) {
                        $txtDescription.Text = $selectedRow.Description
                        Write-DebugMessage "Set Description to: $($selectedRow.Description)"
                    }
                    
                    if ($selectedRow.PSObject.Properties.Name -contains "Ablaufdatum" -and $txtTermination) {
                        $txtTermination.Text = $selectedRow.Ablaufdatum
                        Write-DebugMessage "Set Termination to: $($selectedRow.Ablaufdatum)"
                    }
                    
                    # Set checkboxes
                    if ($selectedRow.PSObject.Properties.Name -contains "External" -and $chkExternal) {
                        $externalValue = $selectedRow.External
                        if ($externalValue -eq "True" -or $externalValue -eq "1" -or 
                            $externalValue -eq "Yes" -or $externalValue -eq "Y") {
                            $chkExternal.IsChecked = $true
                            Write-DebugMessage "Set External to: True"
                        } else {
                            $chkExternal.IsChecked = $false
                            Write-DebugMessage "Set External to: False"
                        }
                    }
                    
                    if ($selectedRow.PSObject.Properties.Name -contains "TL" -and $chkTL) {
                        $tlValue = $selectedRow.TL
                        if ($tlValue -eq "True" -or $tlValue -eq "1" -or 
                            $tlValue -eq "Yes" -or $tlValue -eq "Y") {
                            $chkTL.IsChecked = $true
                            Write-DebugMessage "Set TL to: True"
                        } else {
                            $chkTL.IsChecked = $false
                            Write-DebugMessage "Set TL to: False"
                        }
                    }
                    
                    if ($selectedRow.PSObject.Properties.Name -contains "AL" -and $chkAL) {
                        $alValue = $selectedRow.AL
                        if ($alValue -eq "True" -or $alValue -eq "1" -or 
                            $alValue -eq "Yes" -or $alValue -eq "Y") {
                            $chkAL.IsChecked = $true
                            Write-DebugMessage "Set AL to: True"
                        } else {
                            $chkAL.IsChecked = $false
                            Write-DebugMessage "Set AL to: False"
                        }
                    }
                    
                    # Handle account status
                    if ($selectedRow.PSObject.Properties.Name -contains "AccountDisabled") {
                        $disabledValue = $selectedRow.AccountDisabled
                        $isDisabled = $disabledValue -eq "True" -or $disabledValue -eq "1" -or 
                                $disabledValue -eq "Yes" -or $disabledValue -eq "Y"
                        
                        if ($radAccountEnabled -and $radAccountDisabled) {
                            $radAccountEnabled.IsChecked = -not $isDisabled
                            $radAccountDisabled.IsChecked = $isDisabled
                            Write-DebugMessage "Set AccountDisabled to: $isDisabled"
                        }
                    }

                    # Handle AD Group selection
                    if ($selectedRow.PSObject.Properties.Name -contains "ADGroup" -and $icADGroups -and $global:ADGroupItems) {
                        $adGroupName = $selectedRow.ADGroup
                        Write-DebugMessage "Setting ADGroup to: $adGroupName"

                        # Iterate through the AD groups and check the matching one
                        foreach ($item in $global:ADGroupItems) {
                            if ($item.DisplayName -eq $adGroupName -or $item.GroupName -eq $adGroupName) {
                                $item.IsChecked = $true
                                Write-DebugMessage "Selected AD Group: $($item.DisplayName) - $($item.GroupName)"
                                break
                            }
                        }
                    }
                    
                    [System.Windows.MessageBox]::Show(
                        "Successfully imported user data from CSV.", 
                        "Import Complete", "OK", "Information")
                        
                    Write-LogMessage -Message "Imported user data from CSV file: $csvPath" -LogLevel "INFO"
                }
                catch {
                    Write-DebugMessage "Error parsing CSV: $($_.Exception.Message)"
                    [System.Windows.MessageBox]::Show(
                        "Error parsing CSV file: $($_.Exception.Message)", 
                        "CSV Import Error", "OK", "Error")
                }
            }
        }
        catch {
            Write-DebugMessage "Error during CSV import: $($_.Exception.Message)"
            [System.Windows.MessageBox]::Show(
                "An error occurred: $($_.Exception.Message)", 
                "Error", "OK", "Error")
        }
    })
}

if ($btnClearPreview) {
    Write-DebugMessage "Registering event handler for Clear Preview button"
    $btnClearPreview.Add_Click({
        try {
            # Clear all the text fields and checkbox states in the onboarding form
            $fieldsToReset = @(
                "txtFirstName", "txtLastName", "txtDisplayName", 
                "txtEmail", "txtOffice", "txtPhone", 
                "txtMobile", "txtPosition", "txtDepartment", 
                "txtDescription", "txtTermination"
            )
            
            foreach ($fieldName in $fieldsToReset) {
                $field = $onboardingTab.FindName($fieldName)
                if ($field) {
                    $field.Text = ""
                }
            }
            
            # Reset checkboxes
            $checkboxesToReset = @(
                "chkExternal", "chkTL", "chkAL"
            )
            
            foreach ($checkboxName in $checkboxesToReset) {
                $checkbox = $onboardingTab.FindName($checkboxName)
                if ($checkbox) {
                    $checkbox.IsChecked = $false
                }
            }

            # Reset AD Group selections
            if ($icADGroups -and $global:ADGroupItems) {
                foreach ($item in $global:ADGroupItems) {
                    $item.IsChecked = $false
                }
            }
            
            # Reset radio buttons for account status
            $radAccountEnabled = $onboardingTab.FindName("radAccountEnabled")
            $radAccountDisabled = $onboardingTab.FindName("radAccountDisabled")
            if ($radAccountEnabled) { $radAccountEnabled.IsChecked = $true }
            if ($radAccountDisabled) { $radAccountDisabled.IsChecked = $false }
            
            [System.Windows.MessageBox]::Show("Preview and form fields have been cleared.", "Cleared", "OK", "Information")
        }
        catch {
            Write-DebugMessage "Error clearing preview: $($_.Exception.Message)"
            [System.Windows.MessageBox]::Show(
                "Error clearing preview: $($_.Exception.Message)", 
                "Error", "OK", "Error")
        }
    })
}
#endregion

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
        
        # Get the refresh button using its actual name from XAML
        $btnRefreshADUserList = $adUpdateTab.FindName("btnRefreshADUserList")
        
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
        
        # Create a function to load logos that will be called initially and when tabs change
        function Update-TabLogos {
            param(
                [string]$selectedTabName
            )
            
            Write-DebugMessage "Updating logos for selected tab: $selectedTabName"
            
            # Handle Tab_Onboarding
            if ($selectedTabName -eq "Tab_Onboarding") {
                $picLogo1 = $window.FindName("picLogo1")
                $logoParent = if ($picLogo1) { $picLogo1.Parent } else { $null }
                
                if ($picLogo1 -and $global:Config["WPFGUI"].Contains("HeaderLogo1")) {
                    try {
                        $logoPath = $global:Config["WPFGUI"]["HeaderLogo1"]
                        if (Test-Path $logoPath) {
                            $logo = New-Object System.Windows.Media.Imaging.BitmapImage
                            $logo.BeginInit()
                            $logo.UriSource = New-Object System.Uri($logoPath, [System.UriKind]::Absolute)
                            $logo.EndInit()
                            
                            # Set both the image source and the parent StackPanel background
                            $picLogo1.Source = $logo
                            
                            # Also set the StackPanel background if possible
                            if ($logoParent -is [System.Windows.Controls.StackPanel]) {
                                $imgBrush = New-Object System.Windows.Media.ImageBrush($logo)
                                $logoParent.Background = $imgBrush
                            }
                            
                            # Add click event if URL is specified
                            if ($global:Config["WPFGUI"].Contains("HeaderLogoURL")) {
                                $picLogo1.Cursor = [System.Windows.Input.Cursors]::Hand
                                $picLogo1.Add_MouseLeftButtonUp({
                                    $url = $global:Config["WPFGUI"]["HeaderLogoURL"]
                                    if (-not [string]::IsNullOrEmpty($url)) {
                                        Start-Process $url
                                    }
                                })
                            }
                            Write-DebugMessage "Successfully set header logo for Tab_Onboarding"
                        } else {
                            Write-DebugMessage "Logo file not found: $logoPath"
                        }
                    } catch {
                        Write-DebugMessage "Error setting header logo for Tab_Onboarding: $($_.Exception.Message)"
                    }
                }
            }
            
            # Handle Tab_ADUpdate
            elseif ($selectedTabName -eq "Tab_ADUpdate") {
                $picLogo2 = $window.FindName("picLogo2")
                $logoParent = if ($picLogo2) { $picLogo2.Parent } else { $null }
                
                if ($picLogo2 -and $global:Config["WPFGUI"].Contains("HeaderLogo1")) {
                    try {
                        $logoPath = $global:Config["WPFGUI"]["HeaderLogo1"]
                        if (Test-Path $logoPath) {
                            $logo = New-Object System.Windows.Media.Imaging.BitmapImage
                            $logo.BeginInit()
                            $logo.UriSource = New-Object System.Uri($logoPath, [System.UriKind]::Absolute)
                            $logo.EndInit()
                            
                            # Set both the image source and the parent StackPanel background
                            $picLogo2.Source = $logo
                            
                            # Also set the StackPanel background if possible
                            if ($logoParent -is [System.Windows.Controls.StackPanel]) {
                                $imgBrush = New-Object System.Windows.Media.ImageBrush($logo)
                                $logoParent.Background = $imgBrush
                            }
                            
                            # Add click event if URL is specified
                            if ($global:Config["WPFGUI"].Contains("HeaderLogoURL")) {
                                $picLogo2.Cursor = [System.Windows.Input.Cursors]::Hand
                                $picLogo2.Add_MouseLeftButtonUp({
                                    $url = $global:Config["WPFGUI"]["HeaderLogoURL"]
                                    if (-not [string]::IsNullOrEmpty($url)) {
                                        Start-Process $url
                                    }
                                })
                            }
                            Write-DebugMessage "Successfully set header logo for Tab_ADUpdate"
                        } else {
                            Write-DebugMessage "Logo file not found: $logoPath"
                        }
                    } catch {
                        Write-DebugMessage "Error setting header logo for Tab_ADUpdate: $($_.Exception.Message)"
                    }
                }
            }
            
            # Handle Tab_INIEditor
            elseif ($selectedTabName -eq "Tab_INIEditor") {
                $picLogo3 = $window.FindName("picLogo3")
                $logoParent = if ($picLogo3) { $picLogo3.Parent } else { $null }
                
                if ($picLogo3 -and $global:Config["WPFGUI"].Contains("HeaderLogo1")) {
                    try {
                        $logoPath = $global:Config["WPFGUI"]["HeaderLogo1"]
                        if (Test-Path $logoPath) {
                            $logo = New-Object System.Windows.Media.Imaging.BitmapImage
                            $logo.BeginInit()
                            $logo.UriSource = New-Object System.Uri($logoPath, [System.UriKind]::Absolute)
                            $logo.EndInit()
                            
                            # Set both the image source and the parent StackPanel background
                            $picLogo3.Source = $logo
                            
                            # Also set the StackPanel background if possible
                            if ($logoParent -is [System.Windows.Controls.StackPanel]) {
                                $imgBrush = New-Object System.Windows.Media.ImageBrush($logo)
                                $logoParent.Background = $imgBrush
                            }
                            
                            # Add click event if URL is specified
                            if ($global:Config["WPFGUI"].Contains("HeaderLogoURL")) {
                                $picLogo3.Cursor = [System.Windows.Input.Cursors]::Hand
                                $picLogo3.Add_MouseLeftButtonUp({
                                    $url = $global:Config["WPFGUI"]["HeaderLogoURL"]
                                    if (-not [string]::IsNullOrEmpty($url)) {
                                        Start-Process $url
                                    }
                                })
                            }
                            Write-DebugMessage "Successfully set header logo for Tab_INIEditor"
                        } else {
                            Write-DebugMessage "Logo file not found: $logoPath"
                        }
                    } catch {
                        Write-DebugMessage "Error setting header logo for Tab_INIEditor: $($_.Exception.Message)"
                    }
                }
            }
        }

        # Load logos when the window loads
        $window.Add_Loaded({
            # Get the tab control and register for tab changes
            $tabControl = $window.FindName("tabControl")
            if ($tabControl) {
                # Get the initially selected tab
                $selectedTab = $tabControl.SelectedItem
                if ($selectedTab) {
                    Update-TabLogos -selectedTabName $selectedTab.Name
                }
                
                # Register for selection changed events
                $tabControl.Add_SelectionChanged({
                    $selectedTab = $tabControl.SelectedItem
                    if ($selectedTab) {
                        Update-TabLogos -selectedTabName $selectedTab.Name
                    }
                })
            }
    })
    
    # Set footer text from configuration
    $txtFooter = $window.FindName("txtFooter")
    if ($txtFooter) {
        try {
            # Combine footer text with website if both are available
            $footerText = ""
            
            if ($global:Config["WPFGUI"].Contains("FooterText")) {
                $footerText = $global:Config["WPFGUI"]["FooterText"]
            }
                
                if ($global:Config["WPFGUI"].Contains("FooterWebseite")) {
                    # If there's already text, add a separator
                    if (-not [string]::IsNullOrWhiteSpace($footerText)) {
                        $footerText += " | "
                    }
                    $footerText += $global:Config["WPFGUI"]["FooterWebseite"]
                }
                
                # Add any extra info if configured
                if ($global:Config["WPFGUI"].Contains("GUI_ExtraText")) {
                    if (-not [string]::IsNullOrWhiteSpace($footerText)) {
                        $footerText += " | "
                    }
                    $footerText += $global:Config["WPFGUI"]["GUI_ExtraText"]
                }
                
                # Set the footer text
                $txtFooter.Text = $footerText
                
                # Apply styling from configuration if available
                if ($global:Config["WPFGUI"].Contains("FooterTextColor")) {
                    try {
                        $colorValue = $global:Config["WPFGUI"]["FooterTextColor"]
                        $txtFooter.Foreground = New-Object System.Windows.Media.SolidColorBrush(
                            [System.Windows.Media.ColorConverter]::ConvertFromString($colorValue)
                        )
                    } catch {
                        Write-DebugMessage "Error setting footer text color: $($_.Exception.Message)"
                    }
                }
                
                if ($global:Config["WPFGUI"].Contains("FooterFontSize")) {
                    try {
                        $txtFooter.FontSize = [double]$global:Config["WPFGUI"]["FooterFontSize"]
                    } catch {
                        Write-DebugMessage "Error setting footer font size: $($_.Exception.Message)"
                    }
                }
                
                Write-DebugMessage "Footer text set successfully"
            } catch {
                Write-DebugMessage "Error setting footer text: $($_.Exception.Message)"
            }
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