#requires -Version 7.0
[CmdletBinding()]
param(
    [string]$FirstName,
    [string]$LastName,
    [string]$Location,
    [string]$Company,         # e.g. "1" for Company1, "2" for Company2, etc.
    [string]$License = "",
    [switch]$External,
    [string]$ScriptINIPath = "C:\SCRIPT\easyONBOARDINGConfig.ini"  # Bitte exakten Pfad angeben!
)

<#
  This script:
  1) Checks the OS version (Windows 10 / Server 2016+).
  2) Reads the INI file (sections such as [ScriptInfo], [General], [Logging], [Branding-GUI], [Branding-Report], etc.).
  3) Generates a main password as well as 5 additional custom passwords.
  4) Creates an AD user, writes log entries, and generates HTML/TXT reports.
  5) Displays a GUI with three side-by-side panels (Input, Advanced Settings, Info & Tools).
  6) The "CREATE PDF" button calls easyOnboarding_PDFCreator.ps1 and uses the wkhtmltopdfPath defined in [Branding-Report].
  7) UPN logic: Either manually entered (UPNEntered) or generated based on a template (dropdown value in UPNFormat) and the domain defined in the company section.
#>

# ---------------------------------------------------------------------------
# Define custom type "CompanyOption" which overrides ToString()
Add-Type -TypeDefinition @"
public class CompanyOption {
    public string Display { get; set; }
    public string Section { get; set; }
    public override string ToString() {
        return Display;
    }
}
"@

# ---------------------------------------------------------------------------------
# Global configuration (read once)
$global:Config = $null
# Global variable for GUI font size; default is 8 if not set in INI
$global:guiFontSize = 8

# ---------------------------------------------------------------------------------
# Function to convert hex color to System.Drawing.Color
function Convert-HexToColor {
    param (
        [Parameter(Mandatory=$true)]
        [string]$hex
    )
    if ($hex.StartsWith("#")) { $hex = $hex.Substring(1) }
    return [System.Drawing.Color]::FromArgb([Convert]::ToInt32("FF$hex", 16))
}

# ---------------------------------------------------------------------------------
# 1) OS Compatibility Check
# ---------------------------------------------------------------------------------
try {
    $os = Get-CimInstance -ClassName Win32_OperatingSystem
    if (-not ($os.Version -match "^(10\.|\d{2}\.)")) {
        Throw "This script only runs on Windows 10 or Windows Server 2016 or higher. Current OS version: $($os.Version)"
    }
}
catch {
    Write-Error "Error checking OS version: $_"
    exit
}

# ---------------------------------------------------------------------------------
# 2) Read INI file (robust INI parsing)
# ---------------------------------------------------------------------------------
function Get-IniContent {
    param(
        [Parameter(Mandatory)]
        [string]$Path
    )
    if (-not (Test-Path $Path)) {
        Throw "INI file not found: $Path"
    }
    $ini = @{}
    $currentSection = "Global"
    foreach ($line in Get-Content -Path $Path) {
        $line = $line.Trim()
        if (($line -eq "") -or ($line -match '^\s*[;#]')) { continue }
        if ($line -match '^\[(.+)\]$') {
            $currentSection = $matches[1].Trim()
            if (-not $ini.ContainsKey($currentSection)) { $ini[$currentSection] = @{} }
        }
        elseif ($line -match '^(.*?)=(.*)$') {
            $key = $matches[1].Trim()
            $value = $matches[2].Trim()
            if ($currentSection -and $ini[$currentSection]) { $ini[$currentSection][$key] = $value }
        }
    }
    return $ini
}

# ---------------------------------------------------------------------------------
# Central Logging Function (using log path from [Logging])
# ---------------------------------------------------------------------------------
function Log-Message {
    param(
        [Parameter(Mandatory=$true)] [string]$Message,
        [ValidateSet("Info", "Warning", "Error")] [string]$Level = "Info"
    )
    try {
        if ($global:Config.ContainsKey("Logging") -and $global:Config.Logging.ContainsKey("LogPath")) {
            $logPath = $global:Config.Logging["LogPath"]
        }
        else { $logPath = $global:Config.General["LogFilePath"] }
        if (-not (Test-Path $logPath)) { New-Item -ItemType Directory -Path $logPath -Force | Out-Null }
        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        $entry = "$timestamp [$Level] $Message"
        $logFile = Join-Path $logPath ("Onboarding_" + (Get-Date -Format "yyyyMMdd") + ".log")
        Add-Content -Path $logFile -Value $entry
    }
    catch { Write-Error "Error writing to log file: $_" }
}

# ---------------------------------------------------------------------------------
# Replace placeholders in text (e.g. GUI_Header)
# ---------------------------------------------------------------------------------
function Replace-Placeholders {
    param([Parameter(Mandatory=$true)] [string]$Template)
    try {
        $scriptInfo = $global:Config.ScriptInfo
        if ($scriptInfo) {
            if ($scriptInfo.ContainsKey("ScriptVersion")) { $Template = $Template -replace "\{ScriptVersion\}", $scriptInfo["ScriptVersion"] }
            if ($scriptInfo.ContainsKey("LastUpdate")) { $Template = $Template -replace "\{LastUpdate\}", $scriptInfo["LastUpdate"] }
            if ($scriptInfo.ContainsKey("Author")) { $Template = $Template -replace "\{Author\}", $scriptInfo["Author"] }
        }
        return $Template
    }
    catch {
        Log-Message -Message "Error replacing placeholders: $_" -Level Error
        return $Template
    }
}

# ---------------------------------------------------------------------------------
# GUI Helper Functions (AddLabel, AddTextBox, AddCheckBox, AddComboBox, Add-HorizontalLine)
# ---------------------------------------------------------------------------------
function AddLabel {
    param(
        [Parameter(Mandatory=$true)] $parent,
        [string]$text,
        [int]$x,
        [int]$y,
        [switch]$Bold
    )
    $fontSize = 8
    try { $fontSize = [int]$global:guiFontSize; if ($fontSize -lt 1) { $fontSize = 8 } }
    catch { $fontSize = 8 }
    if ($Bold) { 
        $lbl = New-Object System.Windows.Forms.Label 
        $lbl.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", $fontSize, [System.Drawing.FontStyle]::Bold) 
    }
    else { 
        $lbl = New-Object System.Windows.Forms.Label 
        $lbl.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", $fontSize) 
    }
    $lbl.Text = $text; $lbl.Location = New-Object System.Drawing.Point($x, $y); $lbl.AutoSize = $true
    $parent.Controls.Add($lbl)
    return $lbl
}

function AddTextBox {
    param(
        [Parameter(Mandatory=$true)] $parent,
        [string]$default,
        [int]$x,
        [int]$y,
        [int]$width = 250
    )
    $fontSize = [int]$global:guiFontSize
    $tb = New-Object System.Windows.Forms.TextBox
    $tb.Text = $default; $tb.Location = New-Object System.Drawing.Point($x, $y); $tb.Width = $width
    $tb.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", $fontSize)
    $parent.Controls.Add($tb)
    return $tb
}

function AddCheckBox {
    param(
        [Parameter(Mandatory=$true)] $parent,
        [string]$text,
        [bool]$checked,
        [int]$x,
        [int]$y
    )
    $fontSize = [int]$global:guiFontSize
    $cb = New-Object System.Windows.Forms.CheckBox
    $cb.Text = $text; $cb.Location = New-Object System.Drawing.Point($x, $y); $cb.Checked = $checked
    $cb.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", $fontSize)
    $cb.AutoSize = $true; $parent.Controls.Add($cb)
    return $cb
}

function AddComboBox {
    param(
        [Parameter(Mandatory=$true)] $parent,
        [string[]]$items,
        [int]$x,
        [int]$y,
        [int]$width = 150,
        [string]$default = ""
    )
    $cmb = New-Object System.Windows.Forms.ComboBox
    $cmb.DropDownStyle = 'DropDownList'
    $cmb.Location = New-Object System.Drawing.Point($x, $y)
    $cmb.Width = $width
    foreach ($i in $items) { [void]$cmb.Items.Add($i) }
    if ($default -ne "" -and $cmb.Items.Contains($default)) { $cmb.SelectedItem = $default }
    elseif ($cmb.Items.Count -gt 0) { $cmb.SelectedIndex = 0 }
    $parent.Controls.Add($cmb)
    return $cmb
}

function Add-HorizontalLine {
    param(
        [Parameter(Mandatory=$true)] [System.Windows.Forms.Panel]$Parent,
        [Parameter(Mandatory=$true)] [int]$y,
        [int]$x = 10,
        [int]$width = 0,
        [int]$height = 1,
        [System.Drawing.Color]$color = [System.Drawing.Color]::Gray
    )
    if ($width -eq 0) { $width = $Parent.Width - 20 }
    $line = New-Object System.Windows.Forms.Panel
    $line.Size = New-Object System.Drawing.Size($width, $height)
    $line.Location = New-Object System.Drawing.Point($x, $y)
    $line.BackColor = $color; $Parent.Controls.Add($line)
    return $line
}

# ---------------------------------------------------------------------------------
# Advanced Password Generation
# ---------------------------------------------------------------------------------
function Generate-AdvancedPassword {
    param(
        [int]$Length = 12,
        [bool]$IncludeSpecial = $true,
        [bool]$AvoidAmbiguous = $true,
        [int]$MinUpperCase = 2,
        [int]$MinDigits = 2,
        [int]$MinNonAlpha = 2
    )
    $lower = 'abcdefghijklmnopqrstuvwxyz'
    $upper = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
    $digits = '0123456789'
    $special = '!@#$%^&*()'
    if ($AvoidAmbiguous) {
         $ambiguous = 'Il1O0'
         $upper = -join ($upper.ToCharArray() | Where-Object { $ambiguous -notcontains $_ })
         $lower = -join ($lower.ToCharArray() | Where-Object { $ambiguous -notcontains $_ })
         $digits = -join ($digits.ToCharArray() | Where-Object { $ambiguous -notcontains $_ })
         $special = -join ($special.ToCharArray() | Where-Object { $ambiguous -notcontains $_ })
    }
    $all = $lower + $upper + $digits
    if ($IncludeSpecial) { $all += $special }
    do {
        $passwordChars = @()
        for ($i = 0; $i -lt $MinUpperCase; $i++) { $passwordChars += $upper[(Get-Random -Minimum 0 -Maximum $upper.Length)] }
        for ($i = 0; $i -lt $MinDigits; $i++) { $passwordChars += $digits[(Get-Random -Minimum 0 -Maximum $digits.Length)] }
        while ($passwordChars.Count -lt $Length) { $passwordChars += $all[(Get-Random -Minimum 0 -Maximum $all.Length)] }
        $passwordChars = $passwordChars | Sort-Object { Get-Random }
        $generatedPassword = -join $passwordChars
        $nonAlphaCount = ($generatedPassword.ToCharArray() | Where-Object { $_ -notmatch '[A-Za-z]' }).Count
    } while ($nonAlphaCount -lt $MinNonAlpha)
    return $generatedPassword
}

# ---------------------------------------------------------------------------------
# Process Onboarding (AD, Logging, Reports)
# ---------------------------------------------------------------------------------
function Process-Onboarding {
    param(
        [Parameter(Mandatory=$true)] [pscustomobject]$userData,
        [Parameter(Mandatory=$true)] [hashtable]$Config
    )
    # Load AD Module
    try { Import-Module ActiveDirectory -ErrorAction Stop }
    catch { Throw "Could not load AD module: $($_.Exception.Message)" }
    if ([string]::IsNullOrWhiteSpace($userData.FirstName)) { Throw "First Name is required!" }
    # Generate main password or use fixed value
    if ($userData.PasswordMode -eq 1) {
        foreach ($key in @("PasswordLaenge","IncludeSpecialChars","AvoidAmbiguousChars","MinNonAlpha","MinUpperCase","MinDigits")) {
            if (-not $Config.PasswordFixGenerate.ContainsKey($key)) { 
                Throw "Error: '$key' is missing in [PasswordFixGenerate]!" 
            }
        }
        $includeSpecial = ($Config.PasswordFixGenerate["IncludeSpecialChars"] -match "^(?i:true|1)$")
        $avoidAmbiguous = ($Config.PasswordFixGenerate["AvoidAmbiguousChars"] -match "^(?i:true|1)$")
        $UserPW = Generate-AdvancedPassword `
            -Length ([int]$Config.PasswordFixGenerate["PasswordLaenge"]) `
            -IncludeSpecial $includeSpecial `
            -AvoidAmbiguous $avoidAmbiguous `
            -MinUpperCase ([int]$Config.PasswordFixGenerate["MinUpperCase"]) `
            -MinDigits ([int]$Config.PasswordFixGenerate["MinDigits"]) `
            -MinNonAlpha ([int]$Config.PasswordFixGenerate["MinNonAlpha"])
        if ([string]::IsNullOrWhiteSpace($UserPW)) { Throw "Error: Generated password is empty. Please check [PasswordFixGenerate]." }
    }
    else {
        if (-not $Config.PasswordFixGenerate.ContainsKey("fixPassword")) { Throw "Error: 'fixPassword' is missing in [PasswordFixGenerate]!" }
        $UserPW = $Config.PasswordFixGenerate["fixPassword"]
        if ([string]::IsNullOrWhiteSpace($UserPW)) { Throw "Error: No fixed password provided in the INI!" }
    }
    $SecurePW = ConvertTo-SecureString $UserPW -AsPlainText -Force
    # Generate 5 custom passwords
    $customPWLabels = if ($Config.ContainsKey("CustomPWLabels")) { $Config["CustomPWLabels"] } else { @{} }
    $pwLabel1 = $customPWLabels["CustomPW1_Label"] ?? "Custom PW #1"
    $pwLabel2 = $customPWLabels["CustomPW2_Label"] ?? "Custom PW #2"
    $pwLabel3 = $customPWLabels["CustomPW3_Label"] ?? "Custom PW #3"
    $pwLabel4 = $customPWLabels["CustomPW4_Label"] ?? "Custom PW #4"
    $pwLabel5 = $customPWLabels["CustomPW5_Label"] ?? "Custom PW #5"
    $CustomPW1 = Generate-AdvancedPassword
    $CustomPW2 = Generate-AdvancedPassword
    $CustomPW3 = Generate-AdvancedPassword
    $CustomPW4 = Generate-AdvancedPassword
    $CustomPW5 = Generate-AdvancedPassword
    # Create SamAccountName / UPN
    $SamAccountName = ($userData.FirstName.Substring(0,1) + $userData.LastName).ToLower()
    if (-not [string]::IsNullOrWhiteSpace($userData.UPNEntered)) {
        $UPN = $userData.UPNEntered
    }
    else {
        $companySection = $userData.CompanySection.Section
        if (-not $Config.ContainsKey($companySection)) { Throw "Error: Section '$companySection' does not exist in the INI!" }
        $companyData = $Config[$companySection]
        $suffix = ($companySection -replace "\D", "")
        if (-not $companyData.ContainsKey("ActiveDirectoryDomain$suffix") -or -not $companyData["ActiveDirectoryDomain$suffix"]) {
            Throw "Error: 'ActiveDirectoryDomain$suffix' is missing in the INI!"
        }
        $adDomain = "@" + $companyData["ActiveDirectoryDomain$suffix"].Trim()
        if (-not [string]::IsNullOrWhiteSpace($userData.UPNFormat)) {
            # Convert to string and upper-case for comparison
            $upnTemplate = $userData.UPNFormat.ToString().ToUpperInvariant()
            switch -Wildcard ($upnTemplate) {
                "FIRSTNAME.LASTNAME"    { $UPN = "$($userData.FirstName).$($userData.LastName)$adDomain" }
                "F.LASTNAME"            { $UPN = "$($userData.FirstName.Substring(0,1)).$($userData.LastName)$adDomain" }
                "FIRSTNAMELASTNAME"     { $UPN = "$($userData.FirstName)$($userData.LastName)$adDomain" }
                "FLASTNAME"             { $UPN = "$($userData.FirstName.Substring(0,1))$($userData.LastName)$adDomain" }
                Default                 { $UPN = "$SamAccountName$adDomain" }
            }
        }
        else { $UPN = "$SamAccountName$adDomain" }
    }
    # Read company/location data from INI
    $companySection = $userData.CompanySection.Section
    if (-not $Config.ContainsKey($companySection)) { Throw "Error: Section '$companySection' is missing in the INI!" }
    $companyData = $Config[$companySection]
    $suffix = ($companySection -replace "\D", "")
    if (-not $companyData.ContainsKey("Strasse$suffix")) { Throw "Error: 'Strasse$suffix' is missing in the INI!" }
    if (-not $companyData.ContainsKey("PLZ$suffix"))     { Throw "Error: 'PLZ$suffix' is missing in the INI!" }
    if (-not $companyData.ContainsKey("Ort$suffix"))     { Throw "Error: 'Ort$suffix' is missing in the INI!" }
    $Street = $companyData["Strasse$suffix"]
    $Zip = $companyData["PLZ$suffix"]
    $City = $companyData["Ort$suffix"]
    if (-not $userData.MailSuffix) {
        if (-not $Config.ContainsKey("MailEndungen") -or -not $Config.MailEndungen.ContainsKey("Domain1")) {
            Throw "Error: Mail endings in [MailEndungen] are not defined!"
        }
        $userData.MailSuffix = $Config.MailEndungen["Domain1"]
    }
    if (-not $companyData.ContainsKey("Country$suffix") -or -not $companyData["Country$suffix"]) {
        Throw "Error: 'Country$suffix' is missing in the INI!"
    }
    $Country = $companyData["Country$suffix"]
    if (-not $companyData.ContainsKey("NameFirma$suffix") -or -not $companyData["NameFirma$suffix"]) {
        Throw "Error: 'NameFirma$suffix' is missing in the INI!"
    }
    $companyDisplay = $companyData["NameFirma$suffix"]
    # Integrate ADUserDefaults from INI
    if ($Config.ContainsKey("ADUserDefaults")) { $adDefaults = $Config.ADUserDefaults } else { $adDefaults = @{} }
    if (-not $Config.General.ContainsKey("DefaultOU")) { Throw "Error: 'DefaultOU' is missing in [General]!" }
    $defaultOU = $Config.General["DefaultOU"]
    # Report data from [Branding-Report]
    if (-not $Config.ContainsKey("Branding-Report")) { Throw "Error: Section [Branding-Report] is missing!" }
    $reportBranding = $Config["Branding-Report"]
    if (-not $reportBranding.ContainsKey("ReportPath"))   { Throw "Error: 'ReportPath' is missing in [Branding-Report]!" }
    if (-not $reportBranding.ContainsKey("ReportTitle"))  { Throw "Error: 'ReportTitle' is missing in [Branding-Report]!" }
    if (-not $reportBranding.ContainsKey("ReportFooter")) { Throw "Error: 'ReportFooter' is missing in [Branding-Report]!" }
    $reportPath         = $reportBranding["ReportPath"]
    $reportTitle        = $reportBranding["ReportTitle"]
    $finalReportFooter  = $reportBranding["ReportFooter"]
    # Build AD user parameters including defaults from [ADUserDefaults]
    $userParams = @{
        Name                  = $userData.DisplayName
        DisplayName           = $userData.DisplayName
        GivenName             = $userData.FirstName
        Surname               = $userData.LastName
        SamAccountName        = $SamAccountName
        UserPrincipalName     = $UPN
        AccountPassword       = $SecurePW
        Enabled               = (-not $userData.AccountDisabled)
        ChangePasswordAtLogon = if ($adDefaults.ContainsKey("MustChangePasswordAtLogon")) { $adDefaults["MustChangePasswordAtLogon"] -eq "True" } else { $userData.MustChangePassword }
        PasswordNeverExpires  = if ($adDefaults.ContainsKey("PasswordNeverExpires")) { $adDefaults["PasswordNeverExpires"] -eq "True" } else { $userData.PasswordNeverExpires }
        Path                  = $defaultOU
        City                  = $City
        StreetAddress         = $Street
        Country               = $Country
        postalCode            = $Zip
    }
    # Set company as Organization
    $userParams["Company"] = $companyDisplay
    # Additional defaults from [ADUserDefaults]
    if ($adDefaults.ContainsKey("HomeDirectory")) { $userParams["HomeDirectory"] = $adDefaults["HomeDirectory"] -replace "%username%", $SamAccountName }
    if ($adDefaults.ContainsKey("ProfilePath")) { $userParams["ProfilePath"] = $adDefaults["ProfilePath"] -replace "%username%", $SamAccountName }
    if ($adDefaults.ContainsKey("LogonScript")) { $userParams["ScriptPath"] = $adDefaults["LogonScript"] }
    # Do NOT override UPN if already set via the GUI template!
    # Collect additional attributes from the form
    $otherAttributes = @{}
    if ($userData.EmailAddress -and $userData.EmailAddress.Trim()) {
        if ($userData.EmailAddress -notmatch "@") {
            if ($userData.MailSuffix -eq "MailSuffix | Company") {
                $companyDomain = ""
                if ($companyData.ContainsKey("MailDomain$suffix")) { $companyDomain = $companyData["MailDomain$suffix"].Trim() }
                $otherAttributes["mail"] = "$($userData.EmailAddress)$companyDomain"
            }
            else { $otherAttributes["mail"] = "$($userData.EmailAddress)$($userData.MailSuffix)" }
        }
        else { $otherAttributes["mail"] = $userData.EmailAddress }
    }
    if ($userData.Description) { $otherAttributes["description"] = $userData.Description }
    if ($userData.OfficeRoom) { $otherAttributes["physicalDeliveryOfficeName"] = $userData.OfficeRoom }
    if ($userData.PhoneNumber) { $otherAttributes["telephoneNumber"] = $userData.PhoneNumber }
    if ($userData.MobileNumber) { $otherAttributes["mobile"] = $userData.MobileNumber }
    if ($userData.Position) { $otherAttributes["title"] = $userData.Position }
    if ($userData.DepartmentField) { $otherAttributes["department"] = $userData.DepartmentField }
    # Set IP-Telefon from Company (Telefon1)
    if ($companyData.ContainsKey("Telefon1") -and -not [string]::IsNullOrWhiteSpace($companyData["Telefon1"])) {
        $otherAttributes["ipPhone"] = $companyData["Telefon1"].Trim()
    }
    if ($otherAttributes.Count -gt 0) { $userParams["OtherAttributes"] = $otherAttributes }
    # Set account expiration date if provided
    if (-not [string]::IsNullOrWhiteSpace($userData.Ablaufdatum)) {
        try {
            $expirationDate = [DateTime]::Parse($userData.Ablaufdatum)
            $userParams["AccountExpirationDate"] = $expirationDate
        }
        catch { Write-Warning "Invalid termination date format: $($_.Exception.Message)" }
    }
    # Create or update user
    try { $existingUser = Get-ADUser -Filter { SamAccountName -eq $SamAccountName } -ErrorAction SilentlyContinue }
    catch { $existingUser = $null }
    if (-not $existingUser) {
        Write-Host "Creating new user: $($userData.DisplayName)"
        try {
            New-ADUser @userParams -ErrorAction Stop
            Write-Host "AD user created."
        }
        catch {
            $line = $_.InvocationInfo.ScriptLineNumber
            $function = $_.InvocationInfo.FunctionName
            $script = $_.InvocationInfo.ScriptName
            $errorDetails = $_.Exception.ToString()
            $detailedMessage = "Error creating user in script '$script', function '$function', line ${line}: $errorDetails"
            Throw $detailedMessage
        }
        if ($userData.SmartcardLogonRequired) { try { Set-ADUser -Identity $SamAccountName -SmartcardLogonRequired $true } catch {} }
        if ($userData.CannotChangePassword) { Write-Host "(Note: 'CannotChangePassword' via ACL still needs to be implemented.)" }
    }
    else {
        Write-Host "User '$SamAccountName' already exists – updating."
        try {
            Set-ADUser -Identity $existingUser.DistinguishedName `
                -GivenName $userData.FirstName `
                -Surname $userData.LastName `
                -City $City `
                -StreetAddress $Street `
                -Country $Country `
                -Enabled (-not $userData.AccountDisabled) -ErrorAction SilentlyContinue
            Set-ADUser -Identity $existingUser.DistinguishedName -ChangePasswordAtLogon:$userData.MustChangePassword -PasswordNeverExpires:$userData.PasswordNeverExpires
            if ($otherAttributes.Count -gt 0) { Set-ADUser -Identity $existingUser.DistinguishedName -Replace $otherAttributes -ErrorAction SilentlyContinue }
        }
        catch { Write-Warning "Error updating user: $($_.Exception.Message)" }
    }
    try { Set-ADAccountPassword -Identity $SamAccountName -Reset -NewPassword $SecurePW -ErrorAction SilentlyContinue }
    catch { Write-Warning "Error setting password: $($_.Exception.Message)" }
    # AD Group Assignment
    if ($userData.External) {
        Write-Host "External user: Skipping default AD group assignment."
    }
    else {
        foreach ($groupKey in $userData.ADGroupsSelected) {
            $groupName = $Config.ADGroups[$groupKey]
            if ($groupName) {
                try { Add-ADGroupMember -Identity $groupName -Members $SamAccountName -ErrorAction Stop; Write-Host "Group '$groupName' added." }
                catch { Write-Warning "Error adding group '$groupName': $($_.Exception.Message)" }
            }
        }
        if ($Config.ContainsKey("UserCreationDefaults") -and $Config.UserCreationDefaults.ContainsKey("InitialGroupMembership")) {
            $defaultGroups = $Config.UserCreationDefaults["InitialGroupMembership"].Split(";") | ForEach-Object { $_.Trim() }
            foreach ($grp in $defaultGroups) {
                if ($grp) {
                    try { Add-ADGroupMember -Identity $grp -Members $SamAccountName -ErrorAction Stop; Write-Host "Default group '$grp' added." }
                    catch { Write-Warning "Error adding default group '$grp': $($_.Exception.Message)" }
                }
            }
        }
        if ($userData.License -and $userData.License -ne "NONE") {
            $licenseKey = "MS365_" + $userData.License
            if ($Config.ContainsKey("LicensesGroups") -and $Config.LicensesGroups.ContainsKey($licenseKey)) {
                $licenseGroup = $Config.LicensesGroups[$licenseKey]
                try { Add-ADGroupMember -Identity $licenseGroup -Members $SamAccountName -ErrorAction Stop; Write-Host "License group '$licenseGroup' added." }
                catch { Write-Warning "Error adding license group '$licenseGroup': $($_.Exception.Message)" }
            }
            else { Write-Warning "License key '$licenseKey' not found in [LicensesGroups]." }
        }
        if ($Config.ContainsKey("ActivateUserMS365ADSync") -and $Config.ActivateUserMS365ADSync["ADSync"] -eq "1") {
            $syncGroup = $Config.ActivateUserMS365ADSync["ADSyncADGroup"]
            if ($syncGroup) {
                try { Add-ADGroupMember -Identity $syncGroup -Members $SamAccountName -ErrorAction Stop; Write-Host "AD-Sync group '$syncGroup' added." }
                catch { Write-Warning "Error adding AD-Sync group '$syncGroup': $($_.Exception.Message)" }
            }
        }
    }
    try {
        Log-Message -Message ("ONBOARDING PERFORMED BY: " + ([System.Security.Principal.WindowsIdentity]::GetCurrent().Name) +
        ", SamAccountName: $SamAccountName, Display Name: '$($userData.DisplayName)', UPN: '$UPN', Location: '$($userData.Location)', Company: '$companyDisplay', License: '$($userData.License)', Password: '$UserPW', External: $($userData.External)") -Level Info
        Write-Host "Log written."
    }
    catch { Write-Warning "Error logging: $($_.Exception.Message)" }
    # Create Reports (HTML and TXT)
    try {
        if (-not (Test-Path $reportPath)) { New-Item -ItemType Directory -Path $reportPath -Force | Out-Null }
        $htmlFile = Join-Path $reportPath "$SamAccountName.html"
        $htmlTemplatePath = $reportBranding["TemplatePath"]
        if (-not $htmlTemplatePath -or -not (Test-Path $htmlTemplatePath)) { $htmlTemplate = "INI or HTML Template not provided!" }
        else { $htmlTemplate = Get-Content -Path $htmlTemplatePath -Raw }
        $logoTag = ""
        if ($userData.CustomLogo -and (Test-Path $userData.CustomLogo)) { $logoTag = "<img src='$($userData.CustomLogo)' alt='Logo' />" }
        $htmlContent = $htmlTemplate `
            -replace "{{ReportTitle}}", ([string]$reportTitle) `
            -replace "{{Admin}}", $env:USERNAME `
            -replace "{{ReportDate}}", (Get-Date -Format "yyyy-MM-dd") `
            -replace "{{Vorname}}", $userData.FirstName `
            -replace "{{Nachname}}", $userData.LastName `
            -replace "{{DisplayName}}", $userData.DisplayName `
            -replace "{{Description}}", $userData.Description `
            -replace "{{Buero}}", $userData.OfficeRoom `
            -replace "{{Rufnummer}}", $userData.PhoneNumber `
            -replace "{{Mobil}}", $userData.MobileNumber `
            -replace "{{Position}}", $userData.Position `
            -replace "{{Abteilung}}", $userData.DepartmentField `
            -replace "{{Ablaufdatum}}", $userData.Ablaufdatum `
            -replace "{{LoginName}}", $SamAccountName `
            -replace "{{Passwort}}", $UserPW `
            -replace "{{WebsitesHTML}}", "" `
            -replace "{{ReportFooter}}", $finalReportFooter `
            -replace "{{LogoTag}}", $logoTag `
            -replace "{{CustomPWLabel1}}", $pwLabel1 `
            -replace "{{CustomPWLabel2}}", $pwLabel2 `
            -replace "{{CustomPWLabel3}}", $pwLabel3 `
            -replace "{{CustomPWLabel4}}", $pwLabel4 `
            -replace "{{CustomPWLabel5}}", $pwLabel5 `
            -replace "{{CustomPW1}}", $CustomPW1 `
            -replace "{{CustomPW2}}", $CustomPW2 `
            -replace "{{CustomPW3}}", $CustomPW3 `
            -replace "{{CustomPW4}}", $CustomPW4 `
            -replace "{{CustomPW5}}", $CustomPW5
        Set-Content -Path $htmlFile -Value $htmlContent -Encoding UTF8
        Write-Host "HTML report created: $htmlFile"
        if ($userData.OutputTXT) {
            $txtFile = Join-Path $reportPath "$SamAccountName.txt"
            $txtContent = @"
Onboarding Report for New Employees
Created by: $($env:USERNAME)
Date: $(Get-Date -Format 'yyyy-MM-dd')

User Details:
-------------
First Name:      $($userData.FirstName)
Last Name:       $($userData.LastName)
Display Name:    $($userData.DisplayName)
External User:   $($userData.External)
Description:     $($userData.Description)
Office:          $($userData.OfficeRoom)
Phone:           $($userData.PhoneNumber)
Mobile:          $($userData.MobileNumber)
Position:        $($userData.Position)
Department:      $($userData.DepartmentField)
Login Name:      $SamAccountName
Password:        $UserPW

Additional Passwords:
-----------------------
$pwLabel1 : $CustomPW1
$pwLabel2 : $CustomPW2
$pwLabel3 : $CustomPW3
$pwLabel4 : $CustomPW4
$pwLabel5 : $CustomPW5

Useful Links:
-------------
"@
            if ($Config.ContainsKey("Websites")) {
                foreach ($key in $Config.Websites.Keys) {
                    if ($key -match '^EmployeeLink\d+$') {
                        $line = $Config.Websites[$key]
                        $parts = $line -split ';'
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
            Out-File -FilePath $txtFile -InputObject $txtContent -Encoding UTF8
            Write-Host "TXT report created: $txtFile"
        }
    }
    catch { Write-Warning "Error creating reports: $($_.Exception.Message)" }
}

# ---------------------------------------------------------------------------------
# 6) GUI Creation with 3 Panels
# ---------------------------------------------------------------------------------
function Show-OnboardingForm {
    param([hashtable]$INIConfig)
    # Branding report (for PDF)
    $reportBranding = if ($INIConfig.ContainsKey("Branding-Report")) { $INIConfig["Branding-Report"] } else { @{} }
    # Laden der ADUserDefaults aus der INI – so dass diese in der GUI berücksichtigt werden
    if ($INIConfig.ContainsKey("ADUserDefaults")) { $adDefaults = $INIConfig.ADUserDefaults } else { $adDefaults = @{} }
    # Create object for user inputs
    $result = [PSCustomObject]@{
        FirstName               = ""
        LastName                = ""
        DisplayName             = ""
        Description             = ""
        OfficeRoom              = ""
        PhoneNumber             = ""
        MobileNumber            = ""
        Position                = ""
        DepartmentField         = ""
        Location                = ""
        CompanySection          = $null
        License                 = ""
        PasswordNeverExpires    = $false
        MustChangePassword      = $false
        AccountDisabled         = $false
        CannotChangePassword    = $false
        PasswordMode            = 1
        FixPassword             = ""
        PasswordLaenge          = 12
        IncludeSpecialChars     = $true
        AvoidAmbiguousChars     = $true
        OutputHTML              = $true
        OutputPDF               = $false
        OutputTXT               = $false
        UPNEntered              = ""
        UPNFormat               = ""
        EmailAddress            = ""
        MailSuffix              = ""
        Ablaufdatum             = ""
        Cancel                  = $false
        ADGroupsSelected        = @()
        External                = $false
        SmartcardLogonRequired  = $false
        CustomLogo              = ""
    }
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    $guiBranding = if ($INIConfig.ContainsKey("Branding-GUI")) { $INIConfig["Branding-GUI"] } else { @{} }
    $scriptInfo = if ($INIConfig.ContainsKey("ScriptInfo")) { $INIConfig["ScriptInfo"] } else { @{} }
    $guiHeaderRaw = $guiBranding["GUI_Header"]
    if ($guiHeaderRaw) { $guiHeaderRaw = Replace-Placeholders -Template $guiHeaderRaw }
    # If INI specifies a FontSize, set the global variable; fallback to 8
    if ($guiBranding.ContainsKey("FontSize")) {
        try { $global:guiFontSize = [int]$guiBranding["FontSize"]; if ($global:guiFontSize -lt 1) { $global:guiFontSize = 8 } }
        catch { $global:guiFontSize = 8 }
    }
    else { $global:guiFontSize = 8 }
    # Get ThemeColor from INI (if available) and convert it; default to WhiteSmoke
    $themeColor = [System.Drawing.Color]::WhiteSmoke
    if ($guiBranding.ContainsKey("ThemeColor")) {
        try { $themeColor = Convert-HexToColor -hex $guiBranding["ThemeColor"] }
        catch { Write-Warning "Error converting ThemeColor. Using default color." }
    }
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "easyONBOARDING - Create New Employee (AD User)"
    $form.StartPosition = "CenterScreen"
    $form.Size = New-Object System.Drawing.Size(1205,835)
    $form.AutoScroll = $true
    # LEFT PANEL: Input Fields (use themeColor as background)
    $panelLeft = New-Object System.Windows.Forms.Panel
    $panelLeft.Location = New-Object System.Drawing.Point(10,10)
    $panelLeft.Size = New-Object System.Drawing.Size(420,775)
    $panelLeft.BorderStyle = 'FixedSingle'
    $panelLeft.BackColor = $themeColor
    $form.Controls.Add($panelLeft)
    # MIDDLE PANEL: UPN, Email, AD Flags, Password Options, AD Groups (use themeColor)
    $panelMiddle = New-Object System.Windows.Forms.Panel
    $panelMiddle.Location = New-Object System.Drawing.Point(440,10)
    $panelMiddle.Size = New-Object System.Drawing.Size(450,775)
    $panelMiddle.BorderStyle = 'FixedSingle'
    $panelMiddle.BackColor = $themeColor
    $form.Controls.Add($panelMiddle)
    # RIGHT PANEL: Header, Logo, Tools, Info, Progress and Buttons (LightGray background)
    [int]$panelRightWidth = 280
    $panelRight = New-Object System.Windows.Forms.Panel
    $panelRight.Location = New-Object System.Drawing.Point(900,10)
    $panelRight.Size = New-Object System.Drawing.Size($panelRightWidth,775)
    $panelRight.BorderStyle = 'FixedSingle'
    $panelRight.BackColor = [System.Drawing.Color]::LightGray
    $form.Controls.Add($panelRight)
    # LEFT PANEL - Input Fields
    $yLeft = 10
    AddLabel $panelLeft "First Name *:" 10 $yLeft -Bold | Out-Null
    $txtVorname = AddTextBox $panelLeft "" 150 $yLeft; $yLeft += 30
    AddLabel $panelLeft "Last Name *:" 10 $yLeft -Bold | Out-Null
    $txtNachname = AddTextBox $panelLeft "" 150 $yLeft; $yLeft += 35
    AddLabel $panelLeft "External Employee:" 10 $yLeft -Bold | Out-Null
    $chkExternal = AddCheckBox $panelLeft "" $false 150 $yLeft; $yLeft += 35
    AddLabel $panelLeft "Display Name:" 10 $yLeft -Bold | Out-Null
    $txtDisplayName = AddTextBox $panelLeft "" 150 $yLeft; $yLeft += 30
    AddLabel $panelLeft "Template:" 37 $yLeft -Bold | Out-Null
    $templates = @()
    if ($INIConfig.ContainsKey("DisplayNameTemplates")) {
        $templates = $INIConfig["DisplayNameTemplates"].Keys | ForEach-Object { $INIConfig["DisplayNameTemplates"][$_] }
    }
    $cmbDisplayNameTemplate = AddComboBox $panelLeft $templates 150 $yLeft 250; $yLeft += 40
    AddLabel $panelLeft "Description:" 10 $yLeft -Bold | Out-Null
    $txtDescription = AddTextBox $panelLeft "" 150 $yLeft; $yLeft += 30
    AddLabel $panelLeft "Office:" 10 $yLeft -Bold | Out-Null
    $txtOffice = AddTextBox $panelLeft "" 150 $yLeft; $yLeft += 35
    AddLabel $panelLeft "Phone:" 10 $yLeft -Bold | Out-Null
    $txtPhone = AddTextBox $panelLeft "" 150 $yLeft; $yLeft += 30
    AddLabel $panelLeft "Mobile:" 10 $yLeft -Bold | Out-Null
    $txtMobile = AddTextBox $panelLeft "" 150 $yLeft; $yLeft += 35
    AddLabel $panelLeft "Position:" 10 $yLeft -Bold | Out-Null
    $txtPosition = AddTextBox $panelLeft "" 150 $yLeft; $yLeft += 30
    AddLabel $panelLeft "Department:" 10 $yLeft -Bold | Out-Null
    $txtDeptField = AddTextBox $panelLeft "" 150 $yLeft; $yLeft += 35
    AddLabel $panelLeft "Termination Date:" 10 $yLeft -Bold | Out-Null
    $txtAblaufdatum = AddTextBox $panelLeft "" 150 $yLeft; $yLeft += 40
    $lineAbove = Add-HorizontalLine -Parent $panelLeft -y $yLeft; $yLeft += 15
    AddLabel $panelLeft "Location *:" 10 $yLeft -Bold | Out-Null
    $locationDisplayList = $INIConfig.STANDORTE.Keys | Where-Object { $_ -match '_Bez$' } | ForEach-Object { $INIConfig.STANDORTE[$_] }
    $cmbLocation = AddComboBox $panelLeft $locationDisplayList 150 $yLeft 250; $yLeft += 35
    # -------------------- COMPANY - Only display company name --------------------
    AddLabel $panelLeft "Company *:" 10 $yLeft -Bold | Out-Null
    $companyOptions = New-Object System.Collections.ArrayList
    foreach ($section in $INIConfig.Keys | Where-Object { $_ -like "Company*" }) {
        $suffix = ($section -replace "\D", "")
        $visibleKey = "$section`_Visible"
        if ($INIConfig[$section].ContainsKey($visibleKey)) { if ($INIConfig[$section][$visibleKey] -ne "1") { continue } }
        if ($INIConfig[$section].ContainsKey("NameFirma$suffix") -and -not [string]::IsNullOrWhiteSpace($INIConfig[$section]["NameFirma$suffix"])) {
            $display = $INIConfig[$section]["NameFirma$suffix"].Trim()
            $co = New-Object CompanyOption; $co.Display = $display; $co.Section = $section
            [void]$companyOptions.Add($co)
        }
    }
    $cmbCompany = New-Object System.Windows.Forms.ComboBox
    $cmbCompany.DropDownStyle = 'DropDownList'
    $cmbCompany.FormattingEnabled = $true
    $cmbCompany.Location = New-Object System.Drawing.Point(150, $yLeft)
    $cmbCompany.Width = 250
    $cmbCompany.DataSource    = $companyOptions
    $cmbCompany.DisplayMember = "Display"
    $cmbCompany.ValueMember   = "Section"
    $panelLeft.Controls.Add($cmbCompany)
    $yLeft += 40
    # ------------------------------------------------------------------------
    $lineAbove = Add-HorizontalLine -Parent $panelLeft -y $yLeft; $yLeft += 15
    AddLabel $panelLeft "MS365 License:" 10 $yLeft -Bold | Out-Null
    $cmbMS365License = AddComboBox $panelLeft ( @("NONE") + ($INIConfig.LicensesGroups.Keys | ForEach-Object { $_ -replace '^MS365_','' } ) ) 150 $yLeft 250; $yLeft += 35
    AddLabel $panelLeft "* REQUIRED FIELDS" 12 $yLeft -Bold | Out-Null; $yLeft += 65
    $lineAbove = Add-HorizontalLine -Parent $panelLeft -y $yLeft; $yLeft += 20
    AddLabel $panelLeft "Onboarding Document > Logo" 10 $yLeft -Bold | Out-Null
    $btnBrowseLogo = New-Object System.Windows.Forms.Button; $btnBrowseLogo.Text = "Browse"
    $btnBrowseLogo.Location = New-Object System.Drawing.Point(225, $yLeft)
    $btnBrowseLogo.Size = New-Object System.Drawing.Size(175, 30)
    $btnBrowseLogo.BackColor = [System.Drawing.Color]::DarkGray; $btnBrowseLogo.ForeColor = [System.Drawing.Color]::White
    $btnBrowseLogo.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", $global:guiFontSize, [System.Drawing.FontStyle]::Bold)
    $panelLeft.Controls.Add($btnBrowseLogo)
    $btnBrowseLogo.Add_Click({
        $ofd = New-Object System.Windows.Forms.OpenFileDialog; $ofd.Filter = "Image Files|*.jpg;*.jpeg;*.png;*.gif;*.bmp"
        if ($ofd.ShowDialog() -eq "OK") {
            $result.CustomLogo = $ofd.FileName
            $btnBrowseLogo.Text = "Logo Selected"
            $btnBrowseLogo.BackColor = [System.Drawing.Color]::LightGreen
        }
    })
    $yLeft += 45
    $lineAbove = Add-HorizontalLine -Parent $panelLeft -y $yLeft; $yLeft += 20
    AddLabel $panelLeft "Create Onboarding Document?" 10 $yLeft -Bold | Out-Null; $yLeft += 25
    AddLabel $panelLeft "upn.HTML" 150 $yLeft -Bold | Out-Null
    $lblHTML = AddLabel $panelLeft "X" 250 $yLeft; $lblHTML.ForeColor = [System.Drawing.Color]::Gray; $yLeft += 20
    AddLabel $panelLeft "upn.TXT" 150 $yLeft -Bold | Out-Null
    $chkTXT_Left = AddCheckBox $panelLeft "" $true 250 $yLeft; $yLeft += 30

    # MIDDLE PANEL - UPN, Email, AD Flags, Password Options, AD Groups
    $yMid = 10
    AddLabel $panelMiddle "User Name (UPN):" 10 $yMid -Bold | Out-Null
    $txtUPN = AddTextBox $panelMiddle "" 180 $yMid 250; $yMid += 35
    AddLabel $panelMiddle "Template:" 62 $yMid -Bold | Out-Null
    $cmbUPNFormat = AddComboBox $panelMiddle @("FIRSTNAME.LASTNAME","F.LASTNAME","FIRSTNAMELASTNAME","FLASTNAME") 180 $yMid 250; $yMid += 45
    $lineAbove = Add-HorizontalLine -Parent $panelMiddle -y $yMid; $yMid += 15
    AddLabel $panelMiddle "Email Address:" 10 $yMid -Bold | Out-Null
    $txtEmail = AddTextBox $panelMiddle "" 180 $yMid 250; $yMid += 35
    AddLabel $panelMiddle "Mail Suffix:" 10 $yMid -Bold | Out-Null
    $cmbMailSuffix = AddComboBox $panelMiddle @() 180 $yMid 250
    function Update-MailSuffix {
        param($INIConfig, $cmbMailSuffix)
        $cmbMailSuffix.Items.Clear()
        [void]$cmbMailSuffix.Items.Add("MailSuffix | Company")
        if ($INIConfig.ContainsKey("MailEndungen")) {
            foreach ($key in $INIConfig.MailEndungen.Keys) { [void]$cmbMailSuffix.Items.Add($INIConfig.MailEndungen[$key]) }
        }
        if ($cmbMailSuffix.Items.Count -gt 0) { $cmbMailSuffix.SelectedIndex = 0 }
    }
    Update-MailSuffix -INIConfig $INIConfig -cmbMailSuffix $cmbMailSuffix; $yMid += 45
    $lineAbove = Add-HorizontalLine -Parent $panelMiddle -y $yMid; $yMid += 15
    # AD-User Flags aus [ADUserDefaults] in die GUI laden
    if ($adDefaults) {
        $chkAccountDisabled = AddCheckBox $panelMiddle "Account Disabled" ($adDefaults.AccountDisabled -match "^(?i:true|1)$") 180 $yMid; $yMid += 25
        $chkPWNeverExpires  = AddCheckBox $panelMiddle "Password Never Expires" ($adDefaults.PasswordNeverExpires -match "^(?i:true|1)$") 180 $yMid; $yMid += 25
        $chkSmartcardLogonRequired = AddCheckBox $panelMiddle "Smartcard Logon" $false 180 $yMid; $yMid += 35
        AddLabel $panelMiddle "Not effective when account disabled:" 10 $yMid -Bold | Out-Null; $yMid += 20
        $chkCannotChangePW  = AddCheckBox $panelMiddle "Prevent Password Change" ($adDefaults.CannotChangePassword -match "^(?i:true|1)$") 180 $yMid; $yMid += 25
        $chkMustChange      = AddCheckBox $panelMiddle "Change Password at Logon" ($adDefaults.MustChangePasswordAtLogon -match "^(?i:true|1)$") 180 $yMid; $yMid += 35
    }
    else {
        $chkAccountDisabled = AddCheckBox $panelMiddle "Account Disabled" $true 180 $yMid; $yMid += 25
        $chkPWNeverExpires  = AddCheckBox $panelMiddle "Password Never Expires" $false 180 $yMid; $yMid += 25
        $chkSmartcardLogonRequired = AddCheckBox $panelMiddle "Smartcard Logon" $false 180 $yMid; $yMid += 35
        AddLabel $panelMiddle "Not effective when account disabled:" 10 $yMid -Bold | Out-Null; $yMid += 20
        $chkCannotChangePW  = AddCheckBox $panelMiddle "Prevent Password Change" $false 180 $yMid; $yMid += 25
        $chkMustChange      = AddCheckBox $panelMiddle "Change Password at Logon" $true 180 $yMid; $yMid += 35
    }
    $lineAbove = Add-HorizontalLine -Parent $panelMiddle -y $yMid; $yMid += 10
    AddLabel $panelMiddle "PASSWORD?" 10 $yMid -Bold | Out-Null; $yMid += 15
    $rbFix = New-Object System.Windows.Forms.RadioButton; $rbFix.Text = "FIXED"; $rbFix.Location = New-Object System.Drawing.Point(180, $yMid)
    $panelMiddle.Controls.Add($rbFix)
    $rbRand = New-Object System.Windows.Forms.RadioButton; $rbRand.Text = "GENERATED"; $rbRand.Location = New-Object System.Drawing.Point(300, $yMid); $rbRand.Checked = $true
    $panelMiddle.Controls.Add($rbRand); $yMid += 35
    AddLabel $panelMiddle "Fixed Password:" 10 $yMid -Bold | Out-Null
    $txtFixPW = AddTextBox $panelMiddle "" 180 $yMid 150; $txtFixPW.Enabled = $false; $yMid += 35
    $rbFix.Add_CheckedChanged({ if ($rbFix.Checked) { $txtFixPW.Enabled = $true } })
    $rbRand.Add_CheckedChanged({ if ($rbRand.Checked) { $txtFixPW.Text = ""; $txtFixPW.Enabled = $false } })
    AddLabel $panelMiddle "RANDOM PASSWORD:" 10 $yMid -Bold | Out-Null; $yMid += 25
    AddLabel $panelMiddle "Number of Characters:" 10 $yMid -Bold | Out-Null
    $txtPWLen = AddTextBox $panelMiddle "12" 180 $yMid 50
    $chkIncludeSpecial = AddCheckBox $panelMiddle "Special Characters" $true 260 $yMid; $yMid += 5
    $chkAvoidAmbig = AddCheckBox $panelMiddle "Avoid Ambiguous Characters" $true 260 ($yMid + 20); $yMid += 50
    $lineAbove = Add-HorizontalLine -Parent $panelMiddle -y $yMid; $yMid += 15
    AddLabel $panelMiddle "AD Groups:" 10 $yMid -Bold | Out-Null; $yMid += 15
    $panelADGroups = New-Object System.Windows.Forms.Panel; $panelADGroups.Location = New-Object System.Drawing.Point(10, $yMid)
    $panelADGroups.Size = New-Object System.Drawing.Size(420,180); $panelADGroups.AutoScroll = $true; $panelADGroups.BorderStyle = 'None'
    $panelMiddle.Controls.Add($panelADGroups); $yMid += ($panelADGroups.Height + 10)
    $adGroupChecks = @{}
    if ($INIConfig.ContainsKey("ADGroups")) {
        $adGroupKeys = $INIConfig.ADGroups.Keys | Where-Object { $_ -notmatch '^(DefaultADGroup|.*_(Visible|Label))$' } | Sort-Object { [int]($_ -replace '\D','') }
        $groupCount = 0
        foreach ($g in $adGroupKeys) {
            $visibleKey = $g + "_Visible"; $isVisible = $true
            if ($INIConfig.ADGroups.ContainsKey($visibleKey) -and $INIConfig.ADGroups[$visibleKey] -eq '0') { $isVisible = $false }
            if ($isVisible) {
                $labelKey = $g + "_Label"; $displayText = $g
                if ($INIConfig.ADGroups.ContainsKey($labelKey) -and $INIConfig.ADGroups[$labelKey]) { $displayText = $INIConfig.ADGroups[$labelKey] }
                $col = $groupCount % 3; $row = [math]::Floor($groupCount / 3)
                $x = 10 + ($col * 130); $yPos = 10 + ($row * 30)
                $cbGroup = AddCheckBox $panelADGroups $displayText $false $x $yPos
                $adGroupChecks[$g] = $cbGroup; $groupCount++
            }
        }
    }
    else { AddLabel $panelMiddle "No [ADGroups] section found." 10 $yMid -Bold | Out-Null; $yMid += 25 }
    # RIGHT PANEL: Header, Logo, Tools, Info, Progress and Buttons
    $lblRightHeader = New-Object System.Windows.Forms.Label; $lblRightHeader.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", $global:guiFontSize, [System.Drawing.FontStyle]::Bold)
    $lblRightHeader.AutoSize = $true; $lblRightHeader.Location = New-Object System.Drawing.Point(10,10)
    $lblRightHeader.Text = if ($guiHeaderRaw) { $guiHeaderRaw } else { "HEADER TEXT" }
    $panelRight.Controls.Add($lblRightHeader)
    $picLogo = New-Object System.Windows.Forms.PictureBox; $picLogo.Size = New-Object System.Drawing.Size(($panelRightWidth - 20), 60)
    $picLogo.Location = New-Object System.Drawing.Point(10, 40); $picLogo.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::StretchImage
    if ($guiBranding.ContainsKey("HeaderLogo") -and (Test-Path $guiBranding["HeaderLogo"])) { $picLogo.Image = [System.Drawing.Image]::FromFile($guiBranding["HeaderLogo"]) }
    $picLogo.Cursor = [System.Windows.Forms.Cursors]::Hand
    $picLogo.Add_Click({ if ($guiBranding.ContainsKey("HeaderLogoURL") -and $guiBranding["HeaderLogoURL"]) { Start-Process $guiBranding["HeaderLogoURL"] } })
    $panelRight.Controls.Add($picLogo)
    [int]$toolsY = 110; [int]$toolsSpacing = 45
    for ($i = 1; $i -le 3; $i++) {
       $labelKey = "adminAPPGUI${i}_Label"; $pathKey  = "adminAPPGUI${i}_Path"
       if ($guiBranding.ContainsKey($labelKey) -and $guiBranding.ContainsKey($pathKey) -and
           -not [string]::IsNullOrWhiteSpace($guiBranding[$labelKey]) -and -not [string]::IsNullOrWhiteSpace($guiBranding[$pathKey])) {
           $btnTool = New-Object System.Windows.Forms.Button; $btnTool.Text = $guiBranding[$labelKey]
           $btnTool.Size = New-Object System.Drawing.Size(260, 30); $btnTool.Location = New-Object System.Drawing.Point(10, $toolsY)
           $btnTool.BackColor = [System.Drawing.Color]::DarkGray; $btnTool.ForeColor = [System.Drawing.Color]::White
           $btnTool.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", ([int]$global:guiFontSize -lt 1 ? 8 : [int]$global:guiFontSize), [System.Drawing.FontStyle]::Bold)
           $btnTool.Add_Click({ Start-Process $guiBranding[$pathKey] })
           $panelRight.Controls.Add($btnTool); $toolsY += $toolsSpacing
       }
    }
    [int]$infoTitleY = $toolsY + 35
    $lblInfoTitle = New-Object System.Windows.Forms.Label; $lblInfoTitle.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", ([int]$global:guiFontSize -lt 1 ? 8 : [int]$global:guiFontSize), [System.Drawing.FontStyle]::Bold)
    $lblInfoTitle.AutoSize = $true; $lblInfoTitle.Location = New-Object System.Drawing.Point(10, $infoTitleY)
    $lblInfoTitle.Text = "easyONBOARDING"; $panelRight.Controls.Add($lblInfoTitle)
    $infoBodyY = $infoTitleY + $lblInfoTitle.Height + 5
    if ($guiBranding.ContainsKey("GUI_ExtraText") -and -not [string]::IsNullOrWhiteSpace($guiBranding["GUI_ExtraText"])) {
        $rawText = $guiBranding["GUI_ExtraText"] -replace '\\n', "`n"
        $lblInfoBody = New-Object System.Windows.Forms.Label; $lblInfoBody.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", $global:guiFontSize, [System.Drawing.FontStyle]::Regular)
        $lblInfoBody.AutoSize = $true; $lblInfoBody.MaximumSize = New-Object System.Drawing.Size(260, 0)
        $lblInfoBody.Location = New-Object System.Drawing.Point(10, $infoBodyY); $lblInfoBody.Text = $rawText
        $panelRight.Controls.Add($lblInfoBody)
    }
    $widthForBar = $panelRightWidth - 20; $progressY = 450
    $progressBar = New-Object System.Windows.Forms.ProgressBar; $progressBar.Location = New-Object System.Drawing.Point(10, $progressY)
    $progressBar.Size = New-Object System.Drawing.Size($widthForBar,20); $progressBar.Minimum = 0; $progressBar.Maximum = 100; $progressBar.Value = 0
    $panelRight.Controls.Add($progressBar)
    $lblStatus = New-Object System.Windows.Forms.Label; $lblStatus.Location = New-Object System.Drawing.Point(10, ($progressY + 25))
    $lblStatus.Size = New-Object System.Drawing.Size($widthForBar,20); $lblStatus.Text = "Ready..."
    $panelRight.Controls.Add($lblStatus)
    $btnY = $progressY + 55; $btnHeight = 40; $btnWidth = $panelRightWidth - 20; $spacing = 10
    # ONBOARDING BUTTON
    $btnOnboard = New-Object System.Windows.Forms.Button; $btnOnboard.Text = "ONBOARDING"
    $btnOnboard.Size = New-Object System.Drawing.Size($btnWidth, $btnHeight); $btnOnboard.Location = New-Object System.Drawing.Point(10, $btnY)
    $btnOnboard.BackColor = [System.Drawing.Color]::LightGreen; $btnOnboard.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", $global:guiFontSize)
    $btnOnboard.Add_Click({
        # Collect inputs
        $result.FirstName            = $txtVorname.Text
        $result.LastName             = $txtNachname.Text
        $result.Description          = $txtDescription.Text
        $result.OfficeRoom           = $txtOffice.Text
        $result.PhoneNumber          = $txtPhone.Text
        $result.MobileNumber         = $txtMobile.Text
        $result.Position             = $txtPosition.Text
        $result.DepartmentField      = $txtDeptField.Text
        $result.Ablaufdatum          = $txtAblaufdatum.Text
        $result.Location             = $cmbLocation.SelectedItem
        $result.CompanySection       = $cmbCompany.SelectedItem
        $result.License              = $cmbMS365License.SelectedItem
        $result.UPNEntered           = $txtUPN.Text
        $result.UPNFormat            = $cmbUPNFormat.SelectedItem.ToString()
        $result.EmailAddress         = $txtEmail.Text.Trim()
        $result.MailSuffix           = $cmbMailSuffix.SelectedItem
        $result.OutputTXT            = $chkTXT_Left.Checked
        $result.PasswordNeverExpires = $chkPWNeverExpires.Checked
        $result.MustChangePassword   = $chkMustChange.Checked
        $result.AccountDisabled      = $chkAccountDisabled.Checked
        $result.CannotChangePassword = $chkCannotChangePW.Checked
        $result.SmartcardLogonRequired = $chkSmartcardLogonRequired.Checked

        # UPN and DisplayName handling:
        # DisplayName: if the field is empty, first check the dropdown template; if none, use the default from INI.
        if (-not [string]::IsNullOrWhiteSpace($txtDisplayName.Text)) {
            $result.DisplayName = $txtDisplayName.Text
        }
        elseif ($cmbDisplayNameTemplate.SelectedItem -and $cmbDisplayNameTemplate.SelectedItem -match '{first}') {
            $template = $cmbDisplayNameTemplate.SelectedItem
            $result.DisplayName = $template -replace '{first}', $txtVorname.Text -replace '{last}', $txtNachname.Text
        }
        elseif ($Config.UserCreationDefaults.ContainsKey("DisplayNameFormat") -and -not [string]::IsNullOrWhiteSpace($Config.UserCreationDefaults["DisplayNameFormat"])) {
            $template = $Config.UserCreationDefaults["DisplayNameFormat"]
            $result.DisplayName = $template -replace '{first}', $txtVorname.Text -replace '{last}', $txtNachname.Text
        }
        else {
            $result.DisplayName = "$($txtVorname.Text) $($txtNachname.Text)"
        }
        # UPN: Already calculated in Process-Onboarding based on UPNEntered / UPNFormat.
        if ($chkExternal.Checked) {
            $result.DisplayName = "EXTERNAL | $($txtVorname.Text) $($txtNachname.Text)"
            $result.ADGroupsSelected = @()
            $result.External = $true
        }
        else {
            if (-not $result.DisplayName) { $result.DisplayName = "$($txtVorname.Text) $($txtNachname.Text)" }
            $groupSel = @()
            foreach ($key in $adGroupChecks.Keys) {
                if ($adGroupChecks[$key].Checked) { $groupSel += $key }
            }
            $result.ADGroupsSelected = $groupSel
            $result.External = $false
        }
        $lblStatus.Text = "Onboarding in progress..."
        $progressBar.Value = 30
        try {
            Process-Onboarding -userData $result -Config $INIConfig
            $lblStatus.Text = "Onboarding successfully completed."
            $progressBar.Value = 100
        }
        catch {
            $lblStatus.Text = "Error: $($_.Exception.Message)"
            [System.Windows.Forms.MessageBox]::Show("Error during onboarding: $($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    })
    $panelRight.Controls.Add($btnOnboard)
    $btnY += ($btnHeight + $spacing)
    # CREATE PDF BUTTON
    $btnPDF = New-Object System.Windows.Forms.Button; $btnPDF.Text = "CREATE PDF"
    $btnPDF.Size = New-Object System.Drawing.Size($btnWidth, $btnHeight)
    $btnPDF.Location = New-Object System.Drawing.Point(10, $btnY)
    $btnPDF.BackColor = [System.Drawing.Color]::LightYellow; $btnPDF.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", $global:guiFontSize)
    $btnPDF.Add_Click({
        if (-not $result.FirstName -or -not $result.LastName) {
            [System.Windows.Forms.MessageBox]::Show("Please perform onboarding before creating the PDF.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            return
        }
        $SamAccountName = ($result.FirstName.Substring(0,1) + $result.LastName).ToLower()
        if (-not $reportBranding.ContainsKey("ReportPath")) {
            [System.Windows.Forms.MessageBox]::Show("ReportPath missing in [Branding-Report]. Cannot create PDF.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            return
        }
        $htmlReportPath = Join-Path $reportBranding["ReportPath"] "$SamAccountName.html"
        $pdfReportPath  = Join-Path $reportBranding["ReportPath"] "$SamAccountName.pdf"
        $wkhtmltopdfPath = $reportBranding["wkhtmltopdfPath"]
        if (-not $wkhtmltopdfPath -or -not (Test-Path $wkhtmltopdfPath)) { $wkhtmltopdfPath = "C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe" }
        $pdfScript = Join-Path $PSScriptRoot "easyOnboarding_PDFCreator.ps1"
        if (-not (Test-Path $pdfScript)) {
            [System.Windows.Forms.MessageBox]::Show("PDF script 'easyOnboarding_PDFCreator.ps1' not found!", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            return
        }
        Start-Process -FilePath "powershell.exe" -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$pdfScript`" -htmlFile `"$htmlReportPath`" -pdfFile `"$pdfReportPath`" -wkhtmltopdfPath `"$wkhtmltopdfPath`"" -NoNewWindow -Wait
        $lblStatus.Text = "PDF created (see $pdfReportPath)."
    })
    $panelRight.Controls.Add($btnPDF)
    $btnY += ($btnHeight + $spacing)
    # INFO BUTTON
    $btnInfo = New-Object System.Windows.Forms.Button; $btnInfo.Text = "INFO"
    $btnInfo.Size = New-Object System.Drawing.Size($btnWidth, $btnHeight)
    $btnInfo.Location = New-Object System.Drawing.Point(10, $btnY)
    $btnInfo.BackColor = [System.Drawing.Color]::LightBlue; $btnInfo.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", $global:guiFontSize)
    $btnInfo.Add_Click({ [System.Windows.Forms.MessageBox]::Show("Information file? Or info dialog...", "INFO") })
    $panelRight.Controls.Add($btnInfo)
    $btnY += ($btnHeight + $spacing)
    # CLOSE BUTTON
    $btnClose = New-Object System.Windows.Forms.Button; $btnClose.Text = "CLOSE"
    $btnClose.Size = New-Object System.Drawing.Size($btnWidth, $btnHeight)
    $btnClose.Location = New-Object System.Drawing.Point(10, $btnY)
    $btnClose.BackColor = [System.Drawing.Color]::LightCoral; $btnClose.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", $global:guiFontSize)
    $btnClose.Add_Click({ $result.Cancel = $true; $form.Close() })
    $panelRight.Controls.Add($btnClose)
    $btnY += ($btnHeight + $spacing)
    $lblFooterRight = New-Object System.Windows.Forms.Label
    $lblFooterRight.Text = if ($guiBranding.ContainsKey("FooterWebseite")) { $guiBranding["FooterWebseite"] } else { "www.PSscripts.de" }
    $lblFooterRight.AutoSize = $true; $lblFooterRight.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", $global:guiFontSize)
    $lblFooterRight.Location = New-Object System.Drawing.Point(10, ($panelRight.Height - 30))
    $panelRight.Controls.Add($lblFooterRight)
    [void]$form.ShowDialog()
    return $result
}

# ---------------------------------------------------------------------------------
# 7) Main Execution
# ---------------------------------------------------------------------------------
try {
    Write-Host "Loading INI: $ScriptINIPath"
    $global:Config = Get-IniContent -Path $ScriptINIPath
    Log-Message -Message "INI file loaded successfully." -Level Info
}
catch { Throw "Error loading INI file: $_" }
if ($global:Config.General["DebugMode"] -eq "1") { Write-Host "Debug mode enabled." }
$userSelection = Show-OnboardingForm -INIConfig $global:Config
if ($userSelection.Cancel) { Write-Warning "Onboarding cancelled."; return }
Write-Host "Onboarding completed. The window remains open until you click 'Close'."
