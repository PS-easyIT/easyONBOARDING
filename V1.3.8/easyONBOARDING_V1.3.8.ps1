<#
  This script integrates functions for onboarding, AD group creation and the INI editor in a WPF interface.
  
  Requirements: Administrative rights and PowerShell 7 or higher
    - INI file     ("easyONBOARDINGConfig.ini")
    - XAML file    ("MainGUI.xaml")
#>

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

#region [Region 07 | PASSWORD MANAGEMENT]
# Functions for setting and removing password change restrictions

#region [Region 07.1 | PREVENT PASSWORD CHANGE]
# Sets ACL restrictions to prevent users from changing their passwords
Write-DebugMessage "Defining Set-CannotChangePassword function."
function Set-CannotChangePassword {
    # [07.1.1 - Modifies ACL settings to deny password change permissions]
    param(
        [Parameter(Mandatory=$true)]
        [string]$SamAccountName
    )

    try {
        $adUser = Get-ADUser -Identity $SamAccountName -Properties DistinguishedName -ErrorAction Stop
        if (-not $adUser) {
            Write-Warning "User $SamAccountName not found."
            return
        }

        $user = [ADSI]"LDAP://$($adUser.DistinguishedName)"
        $acl = $user.psbase.ObjectSecurity

        Write-DebugMessage "Set-CannotChangePassword: Defining AccessRule"
        # Define AccessRule: SELF is not allowed to 'Change Password'
        $denyRule = New-Object System.DirectoryServices.ActiveDirectoryAccessRule(
            [System.Security.Principal.NTAccount]"NT AUTHORITY\\SELF",
            [System.DirectoryServices.ActiveDirectoryRights]::WriteProperty,
            [System.Security.AccessControl.AccessControlType]::Deny,
            [GUID]"ab721a53-1e2f-11d0-9819-00aa0040529b"  # GUID for 'User-Change-Password'
        )
        $acl.AddAccessRule($denyRule)
        $user.psbase.ObjectSecurity = $acl
        $user.psbase.CommitChanges()
        Write-Log "Prevent Password Change has been set for $SamAccountName." "DEBUG"
    }
    catch {
        Write-Warning "Error setting password change restriction"
    }
}
#endregion

Write-DebugMessage "Defining Remove-CannotChangePassword function."

#region [Region 07.2 | ALLOW PASSWORD CHANGE]
# Removes ACL restrictions to allow users to change their passwords
Write-DebugMessage "Removing CannotChangePassword for $SamAccountName."
function Remove-CannotChangePassword {
    # [07.2.1 - Removes deny rules from user ACL for password change permission]
    param(
        [Parameter(Mandatory=$true)]
        [string]$SamAccountName
    )

    $adUser = Get-ADUser -Identity $SamAccountName -Properties DistinguishedName
    if (-not $adUser) {
        Write-Warning "User $SamAccountName not found."
        return
    }

    $user = [ADSI]"LDAP://$($adUser.DistinguishedName)"
    $acl = $user.psbase.ObjectSecurity

    Write-DebugMessage "Remove-CannotChangePassword: Removing all deny rules"
    # Remove all deny rules that affect SELF and have GUID ab721a53-1e2f-11d0-9819-00aa0040529b
    $rulesToRemove = $acl.Access | Where-Object {
        $_.IdentityReference -eq "NT AUTHORITY\\SELF" -and
        $_.AccessControlType -eq 'Deny' -and
        $_.ObjectType -eq "ab721a53-1e2f-11d0-9819-00aa0040529b"
    }
    foreach ($rule in $rulesToRemove) {
        $acl.RemoveAccessRule($rule) | Out-Null
    }
    $user.psbase.ObjectSecurity = $acl
    $user.psbase.CommitChanges()
    Write-Log "Prevent Password Change has been removed for $SamAccountName." "DEBUG"
}
#endregion
#endregion

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
        $displayNameTemplates.Contains("DefaultUserPrincipalNameFormat") -and 
        -not [string]::IsNullOrWhiteSpace($displayNameTemplates["DefaultUserPrincipalNameFormat"])) {
        $upnTemplate = $displayNameTemplates["DefaultUserPrincipalNameFormat"].ToUpper()
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
            if (-not [string]::IsNullOrEmpty($lastName)) {
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
            throw "No AD path (OU) found for Company section '$companySection'."
        }
        
        # 3) Generate DisplayName according to configured template (from INI)
        $displayNameFormat = ""
        if ($Config.Contains("DisplayNameUPNTemplates") -and 
            $Config.DisplayNameUPNTemplates.Contains("DefaultDisplayNameFormat")) {
            $displayNameFormat = $Config.DisplayNameUPNTemplates.DefaultDisplayNameFormat
        }
        
        # Default to "LastName, FirstName"
        if ([string]::IsNullOrWhiteSpace($displayNameFormat)) {
            $displayNameFormat = "LastName, FirstName"
        }
        
        $displayName = switch -Wildcard ($displayNameFormat) {
            "LastName, FirstName" { "$($UserData.LastName), $($UserData.FirstName)" }
            "FirstName LastName"  { "$($UserData.FirstName) $($UserData.LastName)" }
            "LastName FirstName"  { "$($UserData.LastName) $($UserData.FirstName)" }
            default               { "$($UserData.LastName), $($UserData.FirstName)" }
        }
        
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
        
        # 6) Collect more AD attributes
        $adUserParams = @{
            SamAccountName = $samAccountName
            UserPrincipalName = $userPrincipalName
            Name = $displayName
            DisplayName = $displayName
            GivenName = $UserData.FirstName
            Surname = $UserData.LastName
            Path = $adPath
            AccountPassword = (ConvertTo-SecureString -String $UserData.Password -AsPlainText -Force)
            Enabled = $true
            PasswordNeverExpires = $false
            ChangePasswordAtLogon = $true
        }
        
        # Optional: More attributes if available
        if (-not [string]::IsNullOrWhiteSpace($UserData.Email)) {
            $adUserParams.EmailAddress = $UserData.Email
        }
        
        if (-not [string]::IsNullOrWhiteSpace($UserData.Description)) {
            $adUserParams.Description = $UserData.Description
        }
        
        if (-not [string]::IsNullOrWhiteSpace($UserData.Phone)) {
            $adUserParams.OfficePhone = $UserData.Phone
        }
        
        # 7) Create user
        Write-DebugMessage "Creating AD user: $samAccountName"
        $newUser = New-ADUser @adUserParams -PassThru
        
        # 8) Set optional attributes
        # If certain attributes need to be set separately...
        
        # 9) Set password options (if configured)
        if ($UserData.PreventPasswordChange -eq $true) {
            Write-DebugMessage "Setting 'Prevent Password Change' for $samAccountName"
            Set-CannotChangePassword -SamAccountName $samAccountName
        }
        
        Write-Log "User $samAccountName was successfully created." -LogLevel "INFO"
        return @{
            Success = $true
            Message = "User was successfully created."
            SamAccountName = $samAccountName
            UserPrincipalName = $userPrincipalName
            Password = $UserData.Password
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
#endregion

Write-DebugMessage "Registering GUI event handlers."

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
#endregion

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

            if (-not $brandingConfig.Contains("TemplateLogo") -or [string]::IsNullOrWhiteSpace($brandingConfig["TemplateLogo"])) {
                Throw "No 'TemplateLogo' defined in the Report section."
            }

            $TemplateLogo = $brandingConfig["TemplateLogo"]
            if (-not (Test-Path $TemplateLogo)) {
                try {
                    New-Item -ItemType Directory -Path $TemplateLogo -Force -ErrorAction Stop | Out-Null
                } catch {
                    Throw "Could not create target directory for logo: $($_.Exception.Message)"
                }
            }

            $targetLogoTemplate = Join-Path -Path $TemplateLogo -ChildPath "ReportLogo.png"
            Copy-Item -Path $selectedFilePath -Destination $targetLogoTemplate -Force -ErrorAction Stop
            Write-Log "Logo successfully saved at: $targetLogoTemplate" "DEBUG"
            [System.Windows.MessageBox]::Show("The logo was successfully saved!`nLocation: $targetLogoTemplate", "Success", "OK", "Information")
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

    # [16.1.6.2 - If user did not specify any mailSuffix, fallback from [MailEndungen]/Domain1]
    Write-DebugMessage "Invoke-Onboarding: Checking mail suffix from userData"
    # If the user entered a MailSuffix, use it – otherwise use the INI default value
    if (-not [string]::IsNullOrWhiteSpace($userData.MailSuffix)) {
        $mailSuffix = $userData.MailSuffix.Trim()
        Write-DebugMessage "userData.MailSuffix= $($userData.MailSuffix)"
    }
    else {
        Write-DebugMessage "No mail suffix specified, fallback to [MailEndungen].Domain1"
        if (-not $Config.Contains("MailEndungen") -and -not $Config["MailEndungen"]) {
            Throw "Error: Section 'MailEndungen' is not defined in the INI!"
        }
        if (-not $Config["MailEndungen"].Contains("Domain1")) {
            Throw "Error: Key 'Domain1' is missing in section 'MailEndungen'!"
        }
        $mailSuffix = $Config["MailEndungen"]["Domain1"]
        Write-DebugMessage "userData.MailSuffix= $($userData.MailSuffix)"
    }
    $otherAttributes["mail"] = "$($userData.EmailAddress)$mailSuffix"

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
        $htmlTemplatePath = $reportBranding["TemplatePath"]
        $htmlTemplate = ""
        if (-not $htmlTemplatePath -or -not (Test-Path $htmlTemplatePath)) {
            Write-DebugMessage "No valid HTML template found. Using fallback text."
            $htmlTemplate = "INI or HTML Template not provided!"
        }
        else {
            Write-DebugMessage "Reading HTML template from: $htmlTemplatePath"
            $htmlTemplate = Get-Content -Path $htmlTemplatePath -Raw
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
            -replace "{{Rufnummer}}",      $phoneNumber `
            -replace "{{Mobil}}",          $mobileNumber `
            -replace "{{Position}}",       $position `
            -replace "{{Abteilung}}",      $departmentField `
            -replace "{{Ablaufdatum}}",    $ablaufdatum `
            -replace "{{LoginName}}",      $SamAccountName `
            -replace "{{Passwort}}",       $mainPW `
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
            # [16.1.12.4.1 - If you have a separate txtTemplate]
            $txtTemplatePath = Join-Path $PSScriptRoot "txtTemplate.txt"  # or from the config if desired
            $txtFile = Join-Path $reportPath "$SamAccountName.txt"

            [string]$txtContent = ""
            if (Test-Path $txtTemplatePath) {
                Write-DebugMessage "Reading txtTemplate from: $txtTemplatePath"
                $txtContent = Get-Content -Path $txtTemplatePath -Raw

                # [16.1.12.4.2 - Replace placeholders in the template]
                # Fix placeholder replacements to avoid parsing issues
                $defaultOU = $Config.ADUserDefaults["DefaultOU"]
                $txtContent = $txtContent -replace '\$Config\.ADUserDefaults\["DefaultOU"\]', $defaultOU
                $txtContent = $txtContent -replace "\`$SamAccountName", $SamAccountName
                $txtContent = $txtContent -replace "\`$UserPW", $mainPW
                $txtContent = $txtContent -replace "\`$userData\.AccountDisabled", $userData.AccountDisabled
                # Safer pattern replacement approach
                $txtContent = [regex]::Replace($txtContent, "\$\(.*?\)", "")
                
                # ...remaining code...
            }
            else {
                # [16.1.12.4.4 - If no external txtTemplate file, build it inline]
                Write-DebugMessage "No external txtTemplate found - using inline approach"
                $txtFile = Join-Path $reportPath "$SamAccountName.txt"

                $txtContent = @"
Onboarding Report for New Employees
Created by: $($env:USERNAME)
Date: $(Get-Date -Format 'yyyy-MM-dd')

Active Directory OU Path:
----------------------------------------------------
$($Config.ADUserDefaults["DefaultOU"])
====================================================

Credentials
====================================================
User Name:     $SamAccountName
Password:      $mainPW
----------------------------------------------------
Enabled:       $([bool](-not $userData.AccountDisabled))

User Details:
====================================================
First Name:    $($userData.FirstName)
Last  Name:    $($userData.LastName)
Display Name:  $($userData.DisplayName)
External:      $($userData.External)
Office:        $($userData.OfficeRoom)
Phone:         $($userData.PhoneNumber)
Mobile:        $($userData.MobileNumber)
Position:      $($userData.Position)
Department:    $($userData.DepartmentField)
====================================================

Additional Passwords:
----------------------------------------------------
$pwLabel1 : $CustomPW1
$pwLabel2 : $CustomPW2
$pwLabel3 : $CustomPW3
$pwLabel4 : $CustomPW4
$pwLabel5 : $CustomPW5

Useful Links:
----------------------------------------------------
"@

                # [16.1.12.4.5 - Add the websites from [Websites]]
                if ($Config.Contains("Websites")) {
                    foreach ($key in $Config.Websites.Keys) {
                        if ($key -match '^EmployeeLink\d+$') {
                            $line = $Config.Websites[$key]
                            $parts = $line -split '[\|\;]'
                            if ($parts.Count -eq 3) {
                                $title = $parts[0].Trim()
                                $url   = $parts[1].Trim()
                                $desc  = $parts[2].Trim()
                                $txtContent += "$title : $url  ($desc)`r`n"
                            }
                        }
                    }
                }
                $txtContent += "`r`n$finalReportFooter`r`n"
            }

            Write-DebugMessage "Writing final TXT to: $txtFile"
            Out-File -FilePath $txtFile -InputObject $txtContent -Encoding UTF8
            Write-DebugMessage "TXT report created: $txtFile"
        }
    }
    catch {
        Write-Warning "Error creating reports: $($_.Exception.Message)"
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

            $pdfScript = Join-Path $PSScriptRoot "easyOnboarding_PDFCreator.ps1"
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
    # Define a function to load the INI sections and settings
    function Import-INIEditorData {
        try {
            Write-DebugMessage "Loading INI sections and settings for the editor"
            
            # Get the ListView and DataGrid controls
            $listViewINIEditor = $tabINIEditor.FindName("listViewINIEditor")
            $dataGridINIEditor = $tabINIEditor.FindName("dataGridINIEditor")
            
            if ($null -eq $listViewINIEditor -or $null -eq $dataGridINIEditor) {
                Write-DebugMessage "ListView or DataGrid not found in XAML"
                return
            }
            
            # Create a collection to store the sections
            $sectionItems = New-Object System.Collections.ObjectModel.ObservableCollection[PSObject]
            
            # Add all sections from the global Config
            foreach ($sectionName in $global:Config.Keys) {
                $sectionItems.Add([PSCustomObject]@{
                    SectionName = $sectionName
                })
            }
            
            # Set the ListView ItemsSource
            $listViewINIEditor.ItemsSource = $sectionItems
            
            # Ensure ListView displays section names
            if ($listViewINIEditor.View -is [System.Windows.Controls.GridView]) {
                # GridView already defined in XAML
            } else {
                # Create GridView programmatically
                $gridView = New-Object System.Windows.Controls.GridView
                $column = New-Object System.Windows.Controls.GridViewColumn
                $column.Header = "Section"
                $column.DisplayMemberBinding = New-Object System.Windows.Data.Binding("SectionName")
                $gridView.Columns.Add($column)
                $listViewINIEditor.View = $gridView
            }
            
            # Ensure DataGrid has columns defined
            if ($dataGridINIEditor.Columns.Count -eq 0) {
                # Add columns programmatically
                $keyColumn = New-Object System.Windows.Controls.DataGridTextColumn
                $keyColumn.Header = "Key"
                $keyColumn.Binding = New-Object System.Windows.Data.Binding("Key")
                $keyColumn.Width = New-Object System.Windows.Controls.DataGridLength(1, [System.Windows.Controls.DataGridLengthUnitType]::Star)
                $dataGridINIEditor.Columns.Add($keyColumn)
                
                $valueColumn = New-Object System.Windows.Controls.DataGridTextColumn
                $valueColumn.Header = "Value"
                $valueColumn.Binding = New-Object System.Windows.Data.Binding("Value") 
                $valueColumn.Width = New-Object System.Windows.Controls.DataGridLength(2, [System.Windows.Controls.DataGridLengthUnitType]::Star)
                $dataGridINIEditor.Columns.Add($valueColumn)
                
                $commentColumn = New-Object System.Windows.Controls.DataGridTextColumn
                $commentColumn.Header = "Comment"
                $commentColumn.Binding = New-Object System.Windows.Data.Binding("Comment")
                $commentColumn.Width = New-Object System.Windows.Controls.DataGridLength(2, [System.Windows.Controls.DataGridLengthUnitType]::Star)
                $dataGridINIEditor.Columns.Add($commentColumn)
            }
            
            Write-DebugMessage "Loaded $($sectionItems.Count) INI sections"
        }
        catch {
            Write-DebugMessage "Error loading INI editor data: $($_.Exception.Message)"
            [System.Windows.MessageBox]::Show("Error loading INI data: $($_.Exception.Message)", "Error", "OK", "Error")
        }
    }
    
    # Function to load settings for a selected section
    function Import-SectionSettings {
        param(
            [Parameter(Mandatory=$true)]
            [string]$SectionName
        )
        
        try {
            Write-DebugMessage "Loading settings for section: $SectionName"
            
            $dataGridINIEditor = $tabINIEditor.FindName("dataGridINIEditor")
            if ($null -eq $dataGridINIEditor) {
                Write-DebugMessage "DataGrid not found in XAML"
                return
            }
            
            # Get the section data from the global Config
            $sectionData = $global:Config[$SectionName]
            if ($null -eq $sectionData) {
                Write-DebugMessage "Section $SectionName not found in Config"
                return
            }
            
            Write-DebugMessage "Found section with $($sectionData.Count) keys"
            
            # Create a collection to store the key-value pairs
            $settingsItems = New-Object System.Collections.ObjectModel.ObservableCollection[PSObject]
            
            # Extract comments from the original INI file if available
            $commentsByKey = @{ }
            try {
                $iniContent = Get-Content -Path $INIPath -ErrorAction SilentlyContinue
                $currentSection = ""
                $keyComment = ""
                
                foreach ($line in $iniContent) {
                    $line = $line.Trim()
                    
                    # Check if it's a section header
                    if ($line -match '^\[(.+)\]$') {
                        $currentSection = $matches[1].Trim()
                        continue
                    }
                    
                    # Check if it's a comment
                    if ($line -match '^[;#](.*)$') {
                        $keyComment = $matches[1].Trim()
                        continue
                    }
                    
                    # Check if it's a key-value pair
                    if ($line -match '^(.*?)=(.*)$' -and $currentSection -eq $SectionName) {
                        $key = $matches[1].Trim()
                        $commentsByKey[$key] = $keyComment
                        $keyComment = ""  # Reset for next key
                    }
                }
            }
            catch {
                Write-DebugMessage "Error extracting comments from INI: $($_.Exception.Message)"
            }
            
            # Add all key-value pairs from the section
            foreach ($key in $sectionData.Keys) {
                $value = $sectionData[$key]
                $comment = if ($commentsByKey.ContainsKey($key)) { $commentsByKey[$key] } else { "" }
                
                Write-DebugMessage "Adding key: '$key' with value: '$value'"
                
                $settingsItems.Add([PSCustomObject]@{
                    Key = $key
                    Value = $value
                    Comment = $comment
                    OriginalKey = $key  # Store original key in case it's edited
                })
            }
            
            # Clear current data and set the new DataGrid ItemsSource
            $dataGridINIEditor.ItemsSource = $null
            $dataGridINIEditor.Items.Clear()
            $dataGridINIEditor.ItemsSource = $settingsItems
            
            Write-DebugMessage "Loaded $($settingsItems.Count) settings for section $SectionName"
            
            # Force UI update
            $dataGridINIEditor.UpdateLayout()
        }
        catch {
            Write-DebugMessage "Error loading section settings: $($_.Exception.Message)"
            [System.Windows.MessageBox]::Show("Error loading settings for section '$SectionName': $($_.Exception.Message)", "Error", "OK", "Error")
        }
    }
    
    # Function to save INI changes back to the file
    function Save-INIChanges {
        try {
            Write-DebugMessage "Starting Save-INIChanges..."
            # Make sure INIPath is defined
            if (-not (Get-Variable -Name INIPath -Scope Script -ErrorAction SilentlyContinue)) {
                $script:INIPath = $global:INIPath
                Write-DebugMessage "Set INI path from global: $script:INIPath"
            }
            
            if (-not (Test-Path -Path $script:INIPath)) {
                Write-DebugMessage "INI file not found: $script:INIPath"
                throw "INI file not found at: $script:INIPath"
            }
            
            # Create a backup before making changes
            $backupPath = "$script:INIPath.backup"
            Copy-Item -Path $script:INIPath -Destination $backupPath -Force -ErrorAction Stop
            Write-DebugMessage "Created backup at: $backupPath"
            
            # Create a temporary string builder to build the INI content
            $iniContent = New-Object System.Text.StringBuilder
            
            # Read original INI to preserve comments and structure
            Write-DebugMessage "Reading original INI content from: $script:INIPath"
            $originalContent = Get-Content -Path $script:INIPath -ErrorAction Stop
            $currentSection = ""
            $inSection = $false
            $processedSections = @{ }
            $processedKeys = @{ }
            
            Write-DebugMessage "Processing original content with ${$originalContent.Count} lines"
            
            # First pass: Preserve structure and update existing values
            foreach ($line in $originalContent) {
                $trimmedLine = $line.Trim()
                
                # Check if it's a section header
                if ($trimmedLine -match '^\[(.+)\]$') {
                    $currentSection = $matches[1].Trim()
                    Write-DebugMessage "Found section: [$currentSection]"
                    
                    # Initialize tracking for this section if needed
                    if (-not $processedKeys.ContainsKey($currentSection)) {
                        $processedKeys[$currentSection] = @{ }
                    }
                    
                    $inSection = $global:Config.Contains($currentSection)
                    $processedSections[$currentSection] = $true
                    
                    # Add the section header
                    [void]$iniContent.AppendLine($line)
                    continue
                }
                
                # Check if it's a comment or empty line
                if ([string]::IsNullOrWhiteSpace($trimmedLine) -or $trimmedLine -match '^[;#]') {
                    [void]$iniContent.AppendLine($line)
                    continue
                }
                
                # Check if it's a key-value pair
                if ($trimmedLine -match '^(.*?)=(.*)$') {
                    $key = $matches[1].Trim()
                    
                    # If we're in a tracked section and the key exists in our config, use the updated value
                    if ($inSection -and $global:Config[$currentSection].Contains($key)) {
                        $value = $global:Config[$currentSection][$key]
                        [void]$iniContent.AppendLine("$key=$value")
                        
                        # Mark this key as processed
                        $processedKeys[$currentSection][$key] = $true
                        Write-DebugMessage "Updated key: [$currentSection] $key=$value"
                    }
                    else {
                        if ($inSection) {
                            Write-DebugMessage "Key no longer exists in config: [$currentSection] $key"
                        }
                        # Keep the original line for keys not in our config or not in tracked sections
                        [void]$iniContent.AppendLine($line)
                    }
                }
                else {
                    # Keep other lines as-is
                    [void]$iniContent.AppendLine($line)
                }
            }
            
            # Second pass: Add any new sections or keys that weren't in the original file
            Write-DebugMessage "Adding new sections and keys..."
            foreach ($sectionName in $global:Config.Keys) {
                # Add new section if it wasn't in the original file
                if (-not $processedSections.ContainsKey($sectionName)) {
                    Write-DebugMessage "Adding new section: [$sectionName]"
                    [void]$iniContent.AppendLine("")
                    [void]$iniContent.AppendLine("[$sectionName]")
                    
                    # Initialize tracking for this section
                    if (-not $processedKeys.ContainsKey($sectionName)) {
                        $processedKeys[$sectionName] = @{ }
                    }
                }
                
                # Add any keys that weren't in the original file
                foreach ($key in $global:Config[$sectionName].Keys) {
                    if (-not ($processedKeys.ContainsKey($sectionName) -and $processedKeys[$sectionName].ContainsKey($key))) {
                        $value = $global:Config[$sectionName][$key]
                        [void]$iniContent.AppendLine("$key=$value")
                        Write-DebugMessage "Added new key: [$sectionName] $key=$value"
                    }
                }
            }
            
            # Save the content back to the file
            Write-DebugMessage "Saving new content to INI file..."
            $finalContent = $iniContent.ToString()
            [System.IO.File]::WriteAllText($script:INIPath, $finalContent, [System.Text.Encoding]::UTF8)
            
            Write-DebugMessage "INI file saved successfully: $script:INIPath"
            Write-LogMessage -Message "INI file edited by $($env:USERNAME)" -LogLevel "INFO"
            
            return $true
        }
        catch {
            Write-DebugMessage "Error saving INI changes: $($_.Exception.Message)"
            Write-DebugMessage "Stack trace: $($_.ScriptStackTrace)"
            Write-LogMessage -Message "Error saving INI file: $($_.Exception.Message)" -LogLevel "ERROR"
            return $false
        }
    }
    # Register event handlers for the INI editor UI
    $listViewINIEditor = $tabINIEditor.FindName("listViewINIEditor")
    $dataGridINIEditor = $tabINIEditor.FindName("dataGridINIEditor")
    
    if ($listViewINIEditor) {
        $listViewINIEditor.Add_SelectionChanged({
            if ($null -ne $listViewINIEditor.SelectedItem) {
                $selectedSection = $listViewINIEditor.SelectedItem.SectionName
                Write-DebugMessage "Section selected: $selectedSection"
                Import-SectionSettings -SectionName $selectedSection
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
                Import-SectionSettings -SectionName $sectionName
                
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
                    Import-SectionSettings -SectionName $sectionName
                    
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
                $result = Save-INIChanges
                if ($result) {
                    [System.Windows.MessageBox]::Show("INI file successfully saved.", "Success", "OK", "Information")
                    # Reload the INI to reflect any changes in the UI
                    $global:Config = Get-IniContent -Path $script:INIPath
                    Import-INIEditorData
                    
                    # Update current selection if available
                    if ($listViewINIEditor.SelectedItem -ne $null) {
                        $selectedSection = $listViewINIEditor.SelectedItem.SectionName
                        Import-SectionSettings -SectionName $selectedSection
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

INI file: $script:INIPath", 
                "INI Editor Help", "OK", "Information")
        })
    }
    
    # Initialize the editor with data when the tab is first loaded
    $tabINIEditor.Add_Loaded({
        # Make sure we have the correct path to the INI file
        if (-not (Get-Variable -Name INIPath -Scope Script -ErrorAction SilentlyContinue)) {
            $script:INIPath = $global:INIPath
        }
        
        Write-DebugMessage "INI Editor loaded with file path: $script:INIPath"
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