#requires -Version 7.0
<#
    ----------------------------------------------------------------------------
    EASYONBOARDING - V1.1.5 (OPTIMIZED)
    ----------------------------------------------------------------------------
    Dieses Skript dient dem Onboarding neuer Mitarbeiter in Active Directory.
    Es liest eine INI-Konfiguration, zeigt ein WPF-Dialog an und erzeugt AD-Benutzer.
    Zusätzlich werden Reports (HTML, TXT, ggf. PDF) generiert und Logs erstellt.

    Folgende Verbesserungen wurden umgesetzt:
      1. Flexibler Umgang mit Log-Pfad (Datei oder Ordner).
      2. Laden von System.Windows.Forms für den OpenFileDialog unter PowerShell 7.
      3. Fallback für DisplayName, wenn im GUI keine Eingabe erfolgt.
      4. Abfangen von möglichen Fehlerquellen (z. B. wkhtmltopdf.exe nicht gefunden).
      5. Erweitertes Logging & Kommentierung, mehr Robustheit in der Fehlerbehandlung.
      6. DebugMode in der INI (wenn gewünscht).
      7. Klartext-Passwörter werden im Log weiterhin geschrieben – dies kann bei Bedarf
         entfernt oder verschlüsselt werden, sofern mehr Sicherheit erforderlich ist.

    ----------------------------------------------------------------------------
#>

[CmdletBinding()]
param(
    [string]$FirstName,
    [string]$LastName,
    [string]$Location,
    [string]$Company,
    [string]$License = "",
    [switch]$External,
    [string]$ScriptINIPath = "easyOnboardingConfig.ini"
)

# ------------------------------------------------------------------------------
#  .NET-ASSEMBLIES FÜR WPF + Windows.Forms
# ------------------------------------------------------------------------------
# WPF-Bestandteile:
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase

# Bei PowerShell 7 muss System.Windows.Forms ggf. manuell geladen werden, wenn
# wir den Windows-Forms-OpenFileDialog verwenden (alternativ kann man den WPF-Dialog
# Microsoft.Win32.OpenFileDialog einsetzen, dann ist dieser Schritt nicht nötig).
Add-Type -AssemblyName System.Windows.Forms

# ------------------------------------------------------------------------------
#  BENUTZERDEFINIERTER KLASSETYP "CompanyOption" 
#  (für das Company-Auswahlmenü in der GUI)
# ------------------------------------------------------------------------------
Add-Type -TypeDefinition @"
public class CompanyOption {
    public string Display { get; set; }
    public string Section { get; set; }
    public override string ToString() {
        return Display;
    }
}
"@

# ------------------------------------------------------------------------------
#  GLOBALE VARIABLEN
# ------------------------------------------------------------------------------
$global:Config = $null        # Enthält den INI-Inhalt nach dem Einlesen
$global:guiFontSize = 8       # Default-Schriftgröße
$global:CustomReportLogo = $null  # Logo für Report
$global:adGroupChecks = @{}       # Dictionary für AD-Group-Checkbox-Steuerelemente

# ------------------------------------------------------------------------------
#  FUNKTION: Get-IniContent
#  Liest eine INI-Datei ein und gibt ein verschachteltes Hashtable zurück.
# ------------------------------------------------------------------------------
function Get-IniContent {
    param(
        [Parameter(Mandatory=$true)][string]$Path
    )
    if (-not (Test-Path $Path)) {
        Throw "INI file not found: $Path"
    }
    $ini = @{}
    $currentSection = "Global"

    foreach ($line in Get-Content -Path $Path) {
        $trimmed = $line.Trim()
        # Leere Zeile oder Kommentar überspringen
        if (($trimmed -eq "") -or ($trimmed -match '^\s*[;#]')) { continue }

        # Neue Sektion [SECTION]
        if ($trimmed -match '^\[(.+)\]$') {
            $currentSection = $matches[1].Trim()
            if (-not $ini.ContainsKey($currentSection)) {
                $ini[$currentSection] = @{}
            }
        }
        # Key=Value
        elseif ($trimmed -match '^(.*?)=(.*)$') {
            $key   = $matches[1].Trim()
            $value = $matches[2].Trim()
            $ini[$currentSection][$key] = $value
        }
    }
    return $ini
}

# ------------------------------------------------------------------------------
#  FUNKTION: Log-Message
#  Schreibt Einträge ins Log – kann mit Datei- oder Ordnerangabe umgehen.
# ------------------------------------------------------------------------------
function Log-Message {
    param(
        [Parameter(Mandatory=$true)][string]$Message,
        [ValidateSet("Info","Warning","Error")][string]$Level = "Info"
    )
    try {
        # Ermitteln, wo wir loggen sollen: INI->[Logging]->LogPath
        if ($global:Config.ContainsKey("Logging") -and $global:Config.Logging.ContainsKey("LogPath")) {
            $logPath = $global:Config.Logging["LogPath"]
        }
        else {
            # Fallback: [General]->LogFilePath
            $logPath = $global:Config.General["LogFilePath"]
        }

        # Prüfen, ob logPath eine *.log-Datei oder ein Ordner ist
        $isLogFile = $false
        if ($logPath -match "\.log$") {
            $isLogFile = $true
        }

        if ($isLogFile) {
            # Der User hat eine *.log-Datei definiert
            $fullLogFile = $logPath
            $logDir = Split-Path $fullLogFile
            if (-not (Test-Path $logDir)) {
                New-Item -ItemType Directory -Path $logDir -Force | Out-Null
            }
        }
        else {
            # Der User hat eher einen Ordner für Logs konfiguriert
            if (-not (Test-Path $logPath)) {
                New-Item -ItemType Directory -Path $logPath -Force | Out-Null
            }
            $dateStr = (Get-Date -Format "yyyyMMdd")
            $fileName = "Onboarding_{0}.log" -f $dateStr
            $fullLogFile = Join-Path $logPath $fileName
        }

        # Zeitstempel & Level
        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        $entry = "$timestamp [$Level] $Message"
        Add-Content -Path $fullLogFile -Value $entry
    }
    catch {
        Write-Warning "Error writing to log file: $($_.Exception.Message)"
    }
}

# ------------------------------------------------------------------------------
#  FUNKTION: Generate-AdvancedPassword
#  Erzeugt ein zufälliges Passwort mit diversen Parametern (z. B. Sonderzeichen).
# ------------------------------------------------------------------------------
function Generate-AdvancedPassword {
    param(
        [int]$Length = 12,
        [bool]$IncludeSpecial = $true,
        [bool]$AvoidAmbiguous = $true,
        [int]$MinUpperCase = 2,
        [int]$MinDigits = 2,
        [int]$MinNonAlpha = 2
    )

    $lower   = 'abcdefghijklmnopqrstuvwxyz'
    $upper   = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
    $digits  = '0123456789'
    $special = '!@#$%^&*()'

    # Ambigue Zeichen entfernen, wenn gewünscht (z. B. I, l, 1, O, 0 ...)
    if ($AvoidAmbiguous) {
        $ambiguous = 'Il1O0'
        $upper   = -join ($upper.ToCharArray()   | Where-Object { $ambiguous -notcontains $_ })
        $lower   = -join ($lower.ToCharArray()   | Where-Object { $ambiguous -notcontains $_ })
        $digits  = -join ($digits.ToCharArray()  | Where-Object { $ambiguous -notcontains $_ })
        $special = -join ($special.ToCharArray() | Where-Object { $ambiguous -notcontains $_ })
    }

    $all = $lower + $upper + $digits
    if ($IncludeSpecial) {
        $all += $special
    }

    do {
        $passwordChars = @()
        # Mind. x Großbuchstaben
        for ($i = 0; $i -lt $MinUpperCase; $i++) {
            $passwordChars += $upper[(Get-Random -Minimum 0 -Maximum $upper.Length)]
        }
        # Mind. x Ziffern
        for ($i = 0; $i -lt $MinDigits; $i++) {
            $passwordChars += $digits[(Get-Random -Minimum 0 -Maximum $digits.Length)]
        }
        # Auffüllen bis zur Gesamtlänge
        while ($passwordChars.Count -lt $Length) {
            $passwordChars += $all[(Get-Random -Minimum 0 -Maximum $all.Length)]
        }

        # Reihenfolge zufällig mischen
        $passwordChars = $passwordChars | Sort-Object { Get-Random }
        $generatedPassword = -join $passwordChars

        # Prüfen, ob genug Nicht-Buchstaben drin sind
        $nonAlphaCount = ($generatedPassword.ToCharArray() | Where-Object { $_ -notmatch '[A-Za-z]' }).Count
    }
    while ($nonAlphaCount -lt $MinNonAlpha)

    return $generatedPassword
}

# ------------------------------------------------------------------------------
#  FUNKTION: Process-Onboarding
#  Kernlogik zum Anlegen/Aktualisieren von AD-Benutzern sowie Reporting.
# ------------------------------------------------------------------------------
function Process-Onboarding {
    param(
        [Parameter(Mandatory=$true)] [pscustomobject]$userData,
        [Parameter(Mandatory=$true)] [hashtable]$Config
    )

    # 1) AD-Modul laden (ActiveDirectory)
    try {
        Import-Module ActiveDirectory -ErrorAction Stop
    }
    catch {
        Throw "Could not load AD module: $($_.Exception.Message)"
    }

    # 2) Basis-Checks: Vorname & Nachname
    if ([string]::IsNullOrWhiteSpace($userData.FirstName)) {
        Throw "First Name is required!"
    }
    if ([string]::IsNullOrWhiteSpace($userData.LastName)) {
        Throw "Last Name is required!"
    }

    # 3) Passwort generieren oder fix
    if ($userData.PasswordMode -eq 1) {
        # Prüfen, ob alle nötigen Keys in [PasswordFixGenerate] existieren
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
        if ([string]::IsNullOrWhiteSpace($UserPW)) {
            Throw "Error: Generated password is empty. Please check [PasswordFixGenerate]."
        }
    }
    else {
        # fixes Passwort
        if (-not $Config.PasswordFixGenerate.ContainsKey("fixPassword")) {
            Throw "Error: 'fixPassword' is missing in [PasswordFixGenerate]!"
        }
        $UserPW = $Config.PasswordFixGenerate["fixPassword"]
        if ([string]::IsNullOrWhiteSpace($UserPW)) {
            Throw "Error: No fixed password provided in the INI!"
        }
    }

    $SecurePW = ConvertTo-SecureString $UserPW -AsPlainText -Force

    # 4) Generierung von 5 Custom-Passwörtern (falls der Report sie nutzt)
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

    # 5) Ermittlung von SamAccountName & UPN
    $SamAccountName = ($userData.FirstName.Substring(0,1) + $userData.LastName).ToLower()
    if (-not [string]::IsNullOrWhiteSpace($userData.UPNEntered)) {
        $UPN = $userData.UPNEntered
    }
    else {
        # Fallback aus CompanySection
        $companySection = $userData.CompanySection.Section
        if (-not $Config.ContainsKey($companySection)) {
            Throw "Error: Section '$companySection' does not exist in the INI!"
        }
        $companyData = $Config[$companySection]
        $suffix = ($companySection -replace "\D", "")

        if (-not $companyData.ContainsKey("ActiveDirectoryDomain$suffix") -or -not $companyData["ActiveDirectoryDomain$suffix"]) {
            Throw "Error: 'ActiveDirectoryDomain$suffix' is missing in the INI!"
        }
        $adDomain = "@" + $companyData["ActiveDirectoryDomain$suffix"].Trim()

        if (-not [string]::IsNullOrWhiteSpace($userData.UPNFormat)) {
            $upnTemplate = $userData.UPNFormat.ToString().ToUpperInvariant()
            switch -Wildcard ($upnTemplate) {
                "FIRSTNAME.LASTNAME"    { $UPN = "$($userData.FirstName).$($userData.LastName)$adDomain" }
                "F.LASTNAME"            { $UPN = "$($userData.FirstName.Substring(0,1)).$($userData.LastName)$adDomain" }
                "FIRSTNAMELASTNAME"     { $UPN = "$($userData.FirstName)$($userData.LastName)$adDomain" }
                "FLASTNAME"             { $UPN = "$($userData.FirstName.Substring(0,1))$($userData.LastName)$adDomain" }
                Default                 { $UPN = "$SamAccountName$adDomain" }
            }
        }
        else {
            $UPN = "$SamAccountName$adDomain"
        }
    }

    # 6) Company/Location-Daten
    $companySection = $userData.CompanySection.Section
    if (-not $Config.ContainsKey($companySection)) {
        Throw "Error: Section '$companySection' is missing in the INI!"
    }
    $companyData = $Config[$companySection]
    $suffix = ($companySection -replace "\D", "")

    if (-not $companyData.ContainsKey("Strasse$suffix")) { Throw "Error: 'Strasse$suffix' is missing in the INI!" }
    if (-not $companyData.ContainsKey("PLZ$suffix"))     { Throw "Error: 'PLZ$suffix' is missing in the INI!" }
    if (-not $companyData.ContainsKey("Ort$suffix"))     { Throw "Error: 'Ort$suffix' is missing in the INI!" }
    $Street = $companyData["Strasse$suffix"]
    $Zip    = $companyData["PLZ$suffix"]
    $City   = $companyData["Ort$suffix"]

    # MailSuffix-Fallback
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

    # 7) Zusätzliche Defaults
    if ($Config.ContainsKey("ADUserDefaults")) { 
        $adDefaults = $Config.ADUserDefaults 
    }
    else {
        $adDefaults = @{}
    }

    if (-not $Config.General.ContainsKey("DefaultOU")) {
        Throw "Error: 'DefaultOU' is missing in [General]!"
    }
    $defaultOU = $Config.General["DefaultOU"]

    # Branding-Report prüfen
    if (-not $Config.ContainsKey("Branding-Report")) {
        Throw "Error: Section [Branding-Report] is missing!"
    }
    $reportBranding = $Config["Branding-Report"]
    if (-not $reportBranding.ContainsKey("ReportPath"))   { Throw "Error: 'ReportPath' is missing in [Branding-Report]!" }
    if (-not $reportBranding.ContainsKey("ReportTitle"))  { Throw "Error: 'ReportTitle' is missing in [Branding-Report]!" }
    if (-not $reportBranding.ContainsKey("ReportFooter")) { Throw "Error: 'ReportFooter' is missing in [Branding-Report]!" }
    $reportPath        = $reportBranding["ReportPath"]
    $reportTitle       = $reportBranding["ReportTitle"]
    $finalReportFooter = $reportBranding["ReportFooter"]

    # 8) Fallback für DisplayName
    if ([string]::IsNullOrWhiteSpace($userData.DisplayName)) {
        $userData.DisplayName = "$($userData.FirstName) $($userData.LastName)"
    }

    # 9) ADUser Parameter
    $userParams = @{
        Name                  = $userData.DisplayName
        DisplayName           = $userData.DisplayName
        GivenName             = $userData.FirstName
        Surname               = $userData.LastName
        SamAccountName        = $SamAccountName
        UserPrincipalName     = $UPN
        AccountPassword       = $SecurePW
        Enabled               = (-not $userData.AccountDisabled)

        ChangePasswordAtLogon = if ($adDefaults.ContainsKey("MustChangePasswordAtLogon")) {
                                    $adDefaults["MustChangePasswordAtLogon"] -eq "True"
                                } else {
                                    $userData.MustChangePassword
                                }

        PasswordNeverExpires  = if ($adDefaults.ContainsKey("PasswordNeverExpires")) {
                                    $adDefaults["PasswordNeverExpires"] -eq "True"
                                } else {
                                    $userData.PasswordNeverExpires
                                }

        Path                  = $defaultOU
        City                  = $City
        StreetAddress         = $Street
        Country               = $Country
        postalCode            = $Zip
        Company               = $companyDisplay
    }

    # HomeDirectory, ProfilePath, etc.
    if ($adDefaults.ContainsKey("HomeDirectory")) {
        $userParams["HomeDirectory"] = $adDefaults["HomeDirectory"] -replace "%username%", $SamAccountName
    }
    if ($adDefaults.ContainsKey("ProfilePath")) {
        $userParams["ProfilePath"] = $adDefaults["ProfilePath"] -replace "%username%", $SamAccountName
    }
    if ($adDefaults.ContainsKey("LogonScript")) {
        $userParams["ScriptPath"] = $adDefaults["LogonScript"]
    }

    # 10) Zusätzliche AD-Attribute
    $otherAttributes = @{}
    if ($userData.EmailAddress -and $userData.EmailAddress.Trim()) {
        if ($userData.EmailAddress -notmatch "@") {
            # E-Mail Domain anfügen
            if ($userData.MailSuffix -eq "MailSuffix | Company") {
                # ggf. CompanyDomain
                $companyDomain = ""
                if ($companyData.ContainsKey("MailDomain$suffix")) {
                    $companyDomain = $companyData["MailDomain$suffix"].Trim()
                }
                $otherAttributes["mail"] = "$($userData.EmailAddress)$companyDomain"
            }
            else {
                $otherAttributes["mail"] = "$($userData.EmailAddress)$($userData.MailSuffix)"
            }
        }
        else {
            $otherAttributes["mail"] = $userData.EmailAddress
        }
    }
    if ($userData.Description)       { $otherAttributes["description"] = $userData.Description }
    if ($userData.OfficeRoom)        { $otherAttributes["physicalDeliveryOfficeName"] = $userData.OfficeRoom }
    if ($userData.PhoneNumber)       { $otherAttributes["telephoneNumber"] = $userData.PhoneNumber }
    if ($userData.MobileNumber)      { $otherAttributes["mobile"] = $userData.MobileNumber }
    if ($userData.Position)          { $otherAttributes["title"] = $userData.Position }
    if ($userData.DepartmentField)   { $otherAttributes["department"] = $userData.DepartmentField }

    # IPPhone aus Company
    if ($companyData.ContainsKey("Telefon1") -and -not [string]::IsNullOrWhiteSpace($companyData["Telefon1"])) {
        $otherAttributes["ipPhone"] = $companyData["Telefon1"].Trim()
    }

    if ($otherAttributes.Count -gt 0) {
        $userParams["OtherAttributes"] = $otherAttributes
    }

    # AccountExpirationDate
    if (-not [string]::IsNullOrWhiteSpace($userData.Ablaufdatum)) {
        try {
            $expirationDate = [DateTime]::Parse($userData.Ablaufdatum)
            $userParams["AccountExpirationDate"] = $expirationDate
        }
        catch {
            Write-Warning "Invalid termination date format: $($_.Exception.Message)"
        }
    }

    # 11) Anlegen/Updaten
    try {
        $existingUser = Get-ADUser -Filter { SamAccountName -eq $SamAccountName } -ErrorAction SilentlyContinue
    }
    catch {
        $existingUser = $null
    }

    if (-not $existingUser) {
        Write-Host "Creating new user: $($userData.DisplayName)"
        try {
            New-ADUser @userParams -ErrorAction Stop
            Write-Host "AD user created."
        }
        catch {
            $line = $_.InvocationInfo.ScriptLineNumber
            $func = $_.InvocationInfo.FunctionName
            $script = $_.InvocationInfo.ScriptName
            $errorDetails = $_.Exception.ToString()
            $detailedMessage = "Error creating user in script '$script', function '$func', line ${line}: $errorDetails"
            Throw $detailedMessage
        }

        if ($userData.SmartcardLogonRequired) {
            try {
                Set-ADUser -Identity $SamAccountName -SmartcardLogonRequired $true
            }
            catch {
                Write-Warning "Error enabling SmartcardLogon: $($_.Exception.Message)"
            }
        }
        if ($userData.CannotChangePassword) {
            Write-Host "(Note: 'CannotChangePassword' via ACL still needs to be implemented.)"
        }
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

    # Passwort via Set-ADAccountPassword
    try {
        Set-ADAccountPassword -Identity $SamAccountName -Reset -NewPassword $SecurePW -ErrorAction SilentlyContinue
    }
    catch {
        Write-Warning "Error setting password: $($_.Exception.Message)"
    }

    # 12) AD-Gruppen
    if ($userData.External) {
        Write-Host "External user: Skipping default AD group assignment."
    }
    else {
        foreach ($groupKey in $userData.ADGroupsSelected) {
            $groupName = $Config.ADGroups[$groupKey]
            if ($groupName) {
                try {
                    Add-ADGroupMember -Identity $groupName -Members $SamAccountName -ErrorAction Stop
                    Write-Host "Group '$groupName' added."
                }
                catch {
                    Write-Warning "Error adding group '$groupName': $($_.Exception.Message)"
                }
            }
        }

        if ($Config.ContainsKey("UserCreationDefaults") -and $Config.UserCreationDefaults.ContainsKey("InitialGroupMembership")) {
            $defaultGroups = $Config.UserCreationDefaults["InitialGroupMembership"].Split(";") | ForEach-Object { $_.Trim() }
            foreach ($grp in $defaultGroups) {
                if ($grp) {
                    try {
                        Add-ADGroupMember -Identity $grp -Members $SamAccountName -ErrorAction Stop
                        Write-Host "Default group '$grp' added."
                    }
                    catch {
                        Write-Warning "Error adding default group '$grp': $($_.Exception.Message)"
                    }
                }
            }
        }

        # Lizenz-Gruppen
        if ($userData.License -and $userData.License -ne "NONE") {
            $licenseKey = "MS365_" + $userData.License
            if ($Config.ContainsKey("LicensesGroups") -and $Config.LicensesGroups.ContainsKey($licenseKey)) {
                $licenseGroup = $Config.LicensesGroups[$licenseKey]
                try {
                    Add-ADGroupMember -Identity $licenseGroup -Members $SamAccountName -ErrorAction Stop
                    Write-Host "License group '$licenseGroup' added."
                }
                catch {
                    Write-Warning "Error adding license group '$licenseGroup': $($_.Exception.Message)"
                }
            }
            else {
                Write-Warning "License key '$licenseKey' not found in [LicensesGroups]."
            }
        }

        # MS365 AD-Sync
        if ($Config.ContainsKey("ActivateUserMS365ADSync") -and $Config.ActivateUserMS365ADSync["ADSync"] -eq "1") {
            $syncGroup = $Config.ActivateUserMS365ADSync["ADSyncADGroup"]
            if ($syncGroup) {
                try {
                    Add-ADGroupMember -Identity $syncGroup -Members $SamAccountName -ErrorAction Stop
                    Write-Host "AD-Sync group '$syncGroup' added."
                }
                catch {
                    Write-Warning "Error adding AD-Sync group '$syncGroup': $($_.Exception.Message)"
                }
            }
        }
    }

    # 13) Logging
    try {
        Log-Message -Message ("ONBOARDING PERFORMED BY: " + ([System.Security.Principal.WindowsIdentity]::GetCurrent().Name) +
        ", SamAccountName: $SamAccountName, Display Name: '$($userData.DisplayName)', UPN: '$UPN', " +
        "Location: '$($userData.Location)', Company: '$companyDisplay', License: '$($userData.License)', " +
        "Password: '$UserPW', External: $($userData.External)") -Level Info
        Write-Host "Log written."
    }
    catch {
        Write-Warning "Error logging: $($_.Exception.Message)"
    }

    # 14) Reports (HTML, TXT)
    try {
        if (-not (Test-Path $reportPath)) {
            New-Item -ItemType Directory -Path $reportPath -Force | Out-Null
        }

        $htmlFile         = Join-Path $reportPath "$SamAccountName.html"
        $htmlTemplatePath = $reportBranding["TemplatePath"]
        if (-not $htmlTemplatePath -or -not (Test-Path $htmlTemplatePath)) {
            $htmlTemplate = "INI or HTML Template not provided!"
        }
        else {
            $htmlTemplate = Get-Content -Path $htmlTemplatePath -Raw
        }

        $logoTag = ""
        if ($userData.CustomLogo -and (Test-Path $userData.CustomLogo)) {
            $logoTag = "<img src='$($userData.CustomLogo)' alt='Logo' />"
        }

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

        # TXT-Erstellung optional
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
                        $line  = $Config.Websites[$key]
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
    catch {
        Write-Warning "Error creating reports: $($_.Exception.Message)"
    }
}

# ------------------------------------------------------------------------------
#  FUNKTION: Show-OnboardingForm
#  GUI-Erzeugung per WPF und Erfassung der Eingaben (FirstName, LastName, etc.)
# ------------------------------------------------------------------------------
function Show-OnboardingForm {
    param([hashtable]$INIConfig)

    # Branding-GUI lesen
    $guiBranding = if ($INIConfig.ContainsKey("Branding-GUI")) { $INIConfig["Branding-GUI"] } else { @{} }

    # Schriftgröße
    if ($guiBranding.ContainsKey("FontSize")) {
        try {
            $global:guiFontSize = [int]$guiBranding["FontSize"]
            if ($global:guiFontSize -lt 1) { $global:guiFontSize = 8 }
        }
        catch {
            $global:guiFontSize = 8
        }
    }

    # Hauptfenster
    $window = New-Object System.Windows.Window
    $window.Title = "easyONBOARDING - Create New Employee (AD User)"
    $window.Width = 1300
    $window.Height = 900
    $window.WindowStartupLocation = 'CenterScreen'

    # Hintergrund (Farbverlauf)
    $gradient = New-Object System.Windows.Media.LinearGradientBrush
    $gradient.StartPoint = [System.Windows.Point]::new(0,0)
    $gradient.EndPoint   = [System.Windows.Point]::new(1,1)
    $stop1 = [System.Windows.Media.GradientStop]::new([System.Windows.Media.Colors]::WhiteSmoke, 0)
    $stop2 = [System.Windows.Media.GradientStop]::new([System.Windows.Media.Colors]::LightGray, 1)
    $gradient.GradientStops.Add($stop1)
    $gradient.GradientStops.Add($stop2)
    $window.Background = $gradient

    # Hauptgrid (3 Spalten)
    $mainGrid = New-Object System.Windows.Controls.Grid
    $colLeft  = New-Object System.Windows.Controls.ColumnDefinition; $colLeft.Width = [System.Windows.GridLength]::new(420)
    $colMid   = New-Object System.Windows.Controls.ColumnDefinition; $colMid.Width  = [System.Windows.GridLength]::new(450)
    $colRight = New-Object System.Windows.Controls.ColumnDefinition; $colRight.Width= [System.Windows.GridLength]::new(1,[System.Windows.GridUnitType]::Star)
    [void]$mainGrid.ColumnDefinitions.Add($colLeft)
    [void]$mainGrid.ColumnDefinitions.Add($colMid)
    [void]$mainGrid.ColumnDefinitions.Add($colRight)

    # Kleine Hilfsfunktion für Panel (Rand + Stackpanel)
    function New-SectionPanel {
        $border = New-Object System.Windows.Controls.Border
        $border.BorderBrush     = [System.Windows.Media.Brushes]::DarkGray
        $border.BorderThickness = [System.Windows.Thickness]::new(1)
        $border.CornerRadius    = [System.Windows.CornerRadius]::new(5)
        $border.Margin          = [System.Windows.Thickness]::new(10)
        $border.Background      = [System.Windows.Media.Brushes]::WhiteSmoke

        $stack = New-Object System.Windows.Controls.StackPanel
        [System.Windows.Controls.Grid]::SetIsSharedSizeScope($stack, $true)
        $border.Child = $stack
        return $border, $stack
    }

    # ----------------------------------------------------------------------------
    #   1) Linkes Panel: Eingabefelder
    # ----------------------------------------------------------------------------
    $borderLeft, $panelLeft = New-SectionPanel
    [System.Windows.Controls.Grid]::SetColumn($borderLeft, 0)
    [void]$mainGrid.Children.Add($borderLeft)

    # Eingabe: FirstName, LastName, External?, DisplayName etc.
    function New-Label {
        param(
            [string]$Content,
            [bool]$Bold = $false
        )
        $lbl = New-Object System.Windows.Controls.Label
        $lbl.Content = $Content
        if ($Bold) { $lbl.FontWeight = [System.Windows.FontWeights]::Bold }
        $lbl.Margin = [System.Windows.Thickness]::new(5,5,5,0)
        $lbl.FontSize = $global:guiFontSize
        return $lbl
    }

    function New-TextBox {
        param([string]$Text = "")
        $tb = New-Object System.Windows.Controls.TextBox
        $tb.Text = $Text
        $tb.Margin = [System.Windows.Thickness]::new(5,5,5,0)
        $tb.FontSize = $global:guiFontSize
        $tb.MinWidth = 200
        return $tb
    }

    function New-CheckBox {
        param([string]$Content, [bool]$IsChecked = $false)
        $cb = New-Object System.Windows.Controls.CheckBox
        $cb.Content = $Content
        $cb.IsChecked = $IsChecked
        $cb.Margin = [System.Windows.Thickness]::new(5,5,5,0)
        $cb.FontSize = $global:guiFontSize
        return $cb
    }

    function New-ComboBox {
        param([string[]]$Items, [string]$Default = "")
        $cmb = New-Object System.Windows.Controls.ComboBox
        foreach ($i in $Items) { [void]$cmb.Items.Add($i) }
        if ($Default -and $cmb.Items.Contains($Default)) {
            $cmb.SelectedItem = $Default
        }
        elseif ($cmb.Items.Count -gt 0) {
            $cmb.SelectedIndex = 0
        }
        $cmb.Margin = [System.Windows.Thickness]::new(5,5,5,0)
        $cmb.FontSize = $global:guiFontSize
        $cmb.MinWidth = 200
        return $cmb
    }

    function New-Button {
        param([string]$Content)
        $btn = New-Object System.Windows.Controls.Button
        $btn.Content = $Content
        $btn.Margin  = [System.Windows.Thickness]::new(5,5,5,0)
        $btn.FontSize= $global:guiFontSize
        $btn.Height  = 40
        return $btn
    }

    function New-LabeledRow {
        param(
            [AllowEmptyString()][string]$LabelContent,
            [Parameter(Mandatory=$true)][System.Windows.FrameworkElement]$InputControl,
            [bool]$Bold = $false
        )
        $rowGrid = New-Object System.Windows.Controls.Grid
        $col1 = New-Object System.Windows.Controls.ColumnDefinition
        $col1.Width = "Auto"
        $col1.SharedSizeGroup = "LeftLabels"
        $col2 = New-Object System.Windows.Controls.ColumnDefinition
        $col2.Width = "*"
        $col2.SharedSizeGroup = "LeftInputs"

        [void]$rowGrid.ColumnDefinitions.Add($col1)
        [void]$rowGrid.ColumnDefinitions.Add($col2)

        $lbl = New-Label -Content $LabelContent -Bold:$Bold
        [System.Windows.Controls.Grid]::SetColumn($lbl,0)
        [System.Windows.Controls.Grid]::SetColumn($InputControl,1)
        [void]$rowGrid.Children.Add($lbl)
        [void]$rowGrid.Children.Add($InputControl)
        return $rowGrid
    }

    # Anlegen der Textboxen & Checkboxes
    $panelLeft.Children.Add((New-LabeledRow -LabelContent "First Name *:" -InputControl (New-TextBox) -Bold $true)) | Out-Null
    $panelLeft.Children.Add((New-LabeledRow -LabelContent "Last Name *:"  -InputControl (New-TextBox) -Bold $true)) | Out-Null
    $externalChk = New-CheckBox "External?" $false
    $panelLeft.Children.Add((New-LabeledRow -LabelContent "External Employee:" -InputControl $externalChk -Bold $true)) | Out-Null
    $panelLeft.Children.Add((New-LabeledRow -LabelContent "Display Name:" -InputControl (New-TextBox) -Bold $true)) | Out-Null

    # DisplayName-Templates
    $templates = @()
    if ($INIConfig.ContainsKey("DisplayNameTemplates")) {
        $templates = $INIConfig["DisplayNameTemplates"].Keys | ForEach-Object { $INIConfig["DisplayNameTemplates"][$_] }
    }
    $templateCmb = New-ComboBox -Items $templates
    $panelLeft.Children.Add((New-LabeledRow -LabelContent "Template:" -InputControl $templateCmb -Bold $true)) | Out-Null

    # Weitere Felder: Description, Office, Phone, ...
    $panelLeft.Children.Add((New-LabeledRow -LabelContent "Description:"    -InputControl (New-TextBox) -Bold $true)) | Out-Null
    $panelLeft.Children.Add((New-LabeledRow -LabelContent "Office:"         -InputControl (New-TextBox) -Bold $true)) | Out-Null
    $panelLeft.Children.Add((New-LabeledRow -LabelContent "Phone:"          -InputControl (New-TextBox) -Bold $true)) | Out-Null
    $panelLeft.Children.Add((New-LabeledRow -LabelContent "Mobile:"         -InputControl (New-TextBox) -Bold $true)) | Out-Null
    $panelLeft.Children.Add((New-LabeledRow -LabelContent "Position:"       -InputControl (New-TextBox) -Bold $true)) | Out-Null
    $panelLeft.Children.Add((New-LabeledRow -LabelContent "Department:"     -InputControl (New-TextBox) -Bold $true)) | Out-Null
    $panelLeft.Children.Add((New-LabeledRow -LabelContent "Termination Date:" -InputControl (New-TextBox) -Bold $true)) | Out-Null
    $panelLeft.Children.Add((New-Label "Not effective when disabled!")) | Out-Null

    # Location
    $locItems = $INIConfig.STANDORTE.Keys | Where-Object { $_ -match '_Bez$' } | ForEach-Object { $INIConfig.STANDORTE[$_] }
    $locCmb   = New-ComboBox -Items $locItems
    $panelLeft.Children.Add((New-LabeledRow -LabelContent "Location *:" -InputControl $locCmb -Bold $true)) | Out-Null

    # Company
    $companyOptions = New-Object System.Collections.ArrayList
    foreach ($section in $INIConfig.Keys | Where-Object { $_ -like "Company*" }) {
        $suffix = ($section -replace "\D","")
        $visibleKey = "$section`_Visible"
        if ($INIConfig[$section].ContainsKey($visibleKey) -and $INIConfig[$section][$visibleKey] -ne "1") {
            continue
        }
        if ($INIConfig[$section].ContainsKey("NameFirma$suffix") -and $INIConfig[$section]["NameFirma$suffix"]) {
            $display = $INIConfig[$section]["NameFirma$suffix"].Trim()
            $co = New-Object CompanyOption
            $co.Display = $display
            $co.Section= $section
            [void]$companyOptions.Add($co)
        }
    }

    $cmbCompany = New-ComboBox -Items @()
    foreach ($co in $companyOptions) {
        [void]$cmbCompany.Items.Add($co)
    }
    if ($cmbCompany.Items.Count -gt 0) {
        $cmbCompany.SelectedIndex = 0
    }
    $panelLeft.Children.Add((New-LabeledRow -LabelContent "Company *:" -InputControl $cmbCompany -Bold $true)) | Out-Null

    # License
    $licenses = @("NONE")
    if ($INIConfig.ContainsKey("LicensesGroups")) {
        $licenses += ($INIConfig.LicensesGroups.Keys | ForEach-Object { $_ -replace '^MS365_','' })
    }
    $licCmb = New-ComboBox -Items $licenses
    $panelLeft.Children.Add((New-LabeledRow -LabelContent "MS365 License:" -InputControl $licCmb -Bold $true)) | Out-Null
    $panelLeft.Children.Add((New-Label "* REQUIRED FIELDS" $true)) | Out-Null

    # ----------------------------------------------------------------------------
    #   2) Mittleres Panel: UPN, E-Mail, Flags, Random Password, AD Groups
    # ----------------------------------------------------------------------------
    $borderMid, $panelMiddle = New-SectionPanel
    [System.Windows.Controls.Grid]::SetColumn($borderMid, 1)
    [void]$mainGrid.Children.Add($borderMid)

    $midGrid = New-Object System.Windows.Controls.Grid
    $rowTop    = New-Object System.Windows.Controls.RowDefinition; $rowTop.Height    = "Auto"
    $rowBottom = New-Object System.Windows.Controls.RowDefinition; $rowBottom.Height = "*"
    [void]$midGrid.RowDefinitions.Add($rowTop)
    [void]$midGrid.RowDefinitions.Add($rowBottom)
    [void]$panelMiddle.Children.Add($midGrid)

    # UPN / E-Mail / Flags / Random PW
    $topStack = New-Object System.Windows.Controls.StackPanel
    [System.Windows.Controls.Grid]::SetRow($topStack, 0)
    [void]$midGrid.Children.Add($topStack)

    $lblUPNHead = New-Label "UPN + Template" $true
    $topStack.Children.Add($lblUPNHead) | Out-Null
    $rowUPN     = New-LabeledRow -LabelContent "User Name (UPN):" -InputControl (New-TextBox)
    $topStack.Children.Add($rowUPN) | Out-Null
    $rowTemplate= New-LabeledRow -LabelContent "Template:" -InputControl (New-ComboBox -Items @("FIRSTNAME.LASTNAME","F.LASTNAME","FIRSTNAMELASTNAME","FLASTNAME"))
    $topStack.Children.Add($rowTemplate) | Out-Null

    $lblMail    = New-Label "E-Mail + Mail Suffix" $true
    $topStack.Children.Add($lblMail) | Out-Null
    $rowEmail      = New-LabeledRow -LabelContent "Email Address:" -InputControl (New-TextBox)
    $rowMailSuffix = New-LabeledRow -LabelContent "Mail Suffix:" -InputControl (New-ComboBox -Items @("MailSuffix | Company"))
    $topStack.Children.Add($rowEmail) | Out-Null
    $topStack.Children.Add($rowMailSuffix) | Out-Null

    $lblUserSettings = New-Label "AD-USER SETTINGS" $true
    $topStack.Children.Add($lblUserSettings) | Out-Null
    $chkAccountDisabled   = New-CheckBox "Account Disabled"         $false
    $chkPWNeverExpires    = New-CheckBox "Password Never Expires"   $false
    $chkSmartcard         = New-CheckBox "Smartcard Logon"          $false
    $chkCannotChange      = New-CheckBox "Prevent Password Change"  $false
    $chkMustChange        = New-CheckBox "Change Password at Logon" $false
    $topStack.Children.Add($chkAccountDisabled)   | Out-Null
    $topStack.Children.Add($chkPWNeverExpires)    | Out-Null
    $topStack.Children.Add($chkSmartcard)         | Out-Null
    $topStack.Children.Add($chkCannotChange)      | Out-Null
    $topStack.Children.Add($chkMustChange)        | Out-Null

    # RANDOM PASSWORD
    $lblRandPW = New-Label "RANDOM PASSWORD" $true
    $topStack.Children.Add($lblRandPW) | Out-Null

    $spRandPW = New-Object System.Windows.Controls.StackPanel
    $spRandPW.Orientation = 'Vertical'
    $topStack.Children.Add($spRandPW) | Out-Null

    $spRandRow1 = New-Object System.Windows.Controls.StackPanel
    $spRandRow1.Orientation = 'Horizontal'
    $lblPWLen = New-Label "Number of Characters:" $true
    $spRandRow1.Children.Add($lblPWLen) | Out-Null
    $txtPWLen = New-TextBox "12"; $txtPWLen.Width = 50
    $spRandRow1.Children.Add($txtPWLen) | Out-Null
    $spRandPW.Children.Add($spRandRow1) | Out-Null

    $spRandRow2 = New-Object System.Windows.Controls.StackPanel
    $spRandRow2.Orientation = 'Vertical'
    $chkIncludeSpecial = New-CheckBox "Special Characters" $true
    $chkAvoidAmbig     = New-CheckBox "Avoid Ambiguous Characters" $true
    $spRandRow2.Children.Add($chkIncludeSpecial) | Out-Null
    $spRandRow2.Children.Add($chkAvoidAmbig)     | Out-Null
    $spRandPW.Children.Add($spRandRow2) | Out-Null

    # AD Groups
    $bottomStack = New-Object System.Windows.Controls.StackPanel
    [System.Windows.Controls.Grid]::SetRow($bottomStack, 1)
    [void]$midGrid.Children.Add($bottomStack)
    $lblADGroups = New-Label "AD Groups" $true
    $bottomStack.Children.Add($lblADGroups) | Out-Null

    $svADGroups = New-Object System.Windows.Controls.ScrollViewer
    $svADGroups.Height = 300; $svADGroups.MaxWidth = 400
    $svADGroups.HorizontalScrollBarVisibility = 'Auto'
    $svADGroups.VerticalScrollBarVisibility   = 'Auto'

    $ugADGroups = New-Object System.Windows.Controls.Primitives.UniformGrid
    $ugADGroups.Rows = 5
    $ugADGroups.Columns = 3
    $svADGroups.Content = $ugADGroups
    $bottomStack.Children.Add($svADGroups) | Out-Null

    if ($INIConfig.ContainsKey("ADGroups")) {
        $adGroupKeys = $INIConfig.ADGroups.Keys | Where-Object { $_ -match '^ADGroup\d+$' }
        foreach ($key in $adGroupKeys) {
            $visibleKey = "${key}_Visible"
            if ($INIConfig.ADGroups.ContainsKey($visibleKey) -and $INIConfig.ADGroups[$visibleKey] -eq "1") {
                $labelKey    = "${key}_Label"
                $displayText = if ($INIConfig.ADGroups.ContainsKey($labelKey)) { $INIConfig.ADGroups[$labelKey] } else { $key }
                $cbGroup = New-CheckBox $displayText $false
                [void]$ugADGroups.Children.Add($cbGroup)
                $global:adGroupChecks[$key] = $cbGroup
            }
        }
    }

    # ----------------------------------------------------------------------------
    #   3) Rechtes Panel: Branding, Logo, Buttons
    # ----------------------------------------------------------------------------
    $borderRight, $panelRight = New-SectionPanel
    [System.Windows.Controls.Grid]::SetColumn($borderRight, 2)
    [void]$mainGrid.Children.Add($borderRight)

    # Branding-Header
    $guiHeaderRaw = $guiBranding["GUI_Header"]
    if ($guiHeaderRaw) {
        # Ersetzen von {ScriptVersion}, {LastUpdate}, {Author} aus [ScriptInfo]
        $guiHeaderRaw = $guiHeaderRaw -replace "\{ScriptVersion\}", ($INIConfig.ScriptInfo["ScriptVersion"] ?? "1.0.0")
        $guiHeaderRaw = $guiHeaderRaw -replace "\{LastUpdate\}",    ($INIConfig.ScriptInfo["LastUpdate"] ?? "01.01.2025")
        $guiHeaderRaw = $guiHeaderRaw -replace "\{Author\}",        ($INIConfig.ScriptInfo["Author"]     ?? "PSscripts.de")
    }
    else {
        $guiHeaderRaw = "HEADER TEXT"
    }

    $lblRightHeader = New-Label $guiHeaderRaw $true
    $panelRight.Children.Add($lblRightHeader) | Out-Null

    # Header-Logo
    $picLogo = New-Object System.Windows.Controls.Image
    $picLogo.Width  = 260
    $picLogo.Height = 60
    $picLogo.Margin = [System.Windows.Thickness]::new(5,5,5,5)
    if ($guiBranding.ContainsKey("HeaderLogo") -and (Test-Path $guiBranding["HeaderLogo"])) {
        $imgSource = New-Object System.Windows.Media.Imaging.BitmapImage
        $imgSource.BeginInit()
        $imgSource.UriSource = [System.Uri]::new($guiBranding["HeaderLogo"], [System.UriKind]::Absolute)
        $imgSource.EndInit()
        $picLogo.Source = $imgSource
    }
    $panelRight.Children.Add($picLogo) | Out-Null

    # adminAPPGUIx_*
    for ($i = 1; $i -le 3; $i++) {
        $labelKey = "adminAPPGUI${i}_Label"
        $pathKey  = "adminAPPGUI${i}_Path"
        if ($guiBranding.ContainsKey($labelKey) -and $guiBranding.ContainsKey($pathKey) -and
            -not [string]::IsNullOrWhiteSpace($guiBranding[$labelKey]) -and
            -not [string]::IsNullOrWhiteSpace($guiBranding[$pathKey])) {
            $localLabel = $guiBranding[$labelKey]
            $localPath  = $guiBranding[$pathKey]
            $btnTool = New-Button $localLabel
            $btnTool.Background = [System.Windows.Media.Brushes]::DarkGray
            $btnTool.Foreground = [System.Windows.Media.Brushes]::White
            $btnTool.FontWeight = [System.Windows.FontWeights]::Bold
            $btnTool.Add_Click({ Start-Process $localPath })
            [void]$panelRight.Children.Add($btnTool)
        }
    }

    $lblInfoTitle = New-Label "easyONBOARDING" $true
    $panelRight.Children.Add($lblInfoTitle) | Out-Null

    if ($guiBranding.ContainsKey("GUI_ExtraText") -and $guiBranding["GUI_ExtraText"]) {
        $rawText = $guiBranding["GUI_ExtraText"] -replace '\\n','`n'
        $lblInfoBody = New-Object System.Windows.Controls.TextBlock
        $lblInfoBody.Text = $rawText
        $lblInfoBody.TextWrapping = 'Wrap'
        $lblInfoBody.Margin = [System.Windows.Thickness]::new(5,5,5,5)
        $lblInfoBody.FontSize = $global:guiFontSize
        $panelRight.Children.Add($lblInfoBody) | Out-Null
    }

    # Onboarding-Dokument -> Logo
    $lblDoc = New-Label "Onboarding Document > Logo" $true
    $panelRight.Children.Add($lblDoc) | Out-Null

    $btnBrowseLogo = New-Button "Browse"
    $btnBrowseLogo.Background = [System.Windows.Media.Brushes]::DarkGray
    $btnBrowseLogo.Foreground = [System.Windows.Media.Brushes]::White
    $btnBrowseLogo.FontWeight = [System.Windows.FontWeights]::Bold
    $btnBrowseLogo.Add_Click({
        $ofd = New-Object System.Windows.Forms.OpenFileDialog
        $ofd.Filter = "Image Files|*.jpg;*.jpeg;*.png;*.gif;*.bmp"
        if ($ofd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $btnBrowseLogo.Content = "Logo Selected"
            $btnBrowseLogo.Background = [System.Windows.Media.Brushes]::LightGreen
            $global:CustomReportLogo = $ofd.FileName
        }
    })
    $panelRight.Children.Add($btnBrowseLogo) | Out-Null

    # ONBOARDING Documents (HTML/TXT)
    $lblOnbDocs = New-Label "ONBOARDING Documents" $true
    $lblOnbDocs.Margin = [System.Windows.Thickness]::new(5,20,5,5)
    $panelRight.Children.Add($lblOnbDocs) | Out-Null

    $spDocs = New-Object System.Windows.Controls.StackPanel
    $spDocs.Orientation = 'Horizontal'
    $lblUPNHTML = New-Label "upn.HTML" $true
    $lblUPNTXT  = New-Label "upn.TXT"  $true
    $chkTXT_Left= New-CheckBox "" $true
    [void]$spDocs.Children.Add($lblUPNHTML)
    [void]$spDocs.Children.Add($lblUPNTXT)
    [void]$spDocs.Children.Add($chkTXT_Left)
    $panelRight.Children.Add($spDocs) | Out-Null

    # Fortschrittsanzeige & Status
    $progressBar = New-Object System.Windows.Controls.ProgressBar
    $progressBar.Margin  = [System.Windows.Thickness]::new(5,20,5,5)
    $progressBar.Minimum = 0
    $progressBar.Maximum = 100
    $progressBar.Value   = 0
    $progressBar.Height  = 20
    $panelRight.Children.Add($progressBar) | Out-Null

    $lblStatus = New-Object System.Windows.Controls.Label
    $lblStatus.Content = "Ready..."
    $lblStatus.Margin  = [System.Windows.Thickness]::new(5,0,5,10)
    $lblStatus.FontSize= $global:guiFontSize
    $panelRight.Children.Add($lblStatus) | Out-Null

    # Button-Stack (Onboarding, PDF, Info, Close)
    $btnStack = New-Object System.Windows.Controls.StackPanel
    $btnStack.Margin = [System.Windows.Thickness]::new(5,40,5,5)
    $btnStack.Orientation = 'Vertical'
    $panelRight.Children.Add($btnStack) | Out-Null

    # ONBOARDING Button
    $btnOnboard = New-Button "ONBOARDING"
    $btnOnboard.Background = [System.Windows.Media.Brushes]::LightGreen
    $btnOnboard.Add_Click({
        $lblStatus.Content = "Onboarding in progress..."
        $progressBar.Value = 30
        try {
            # Erzeuge $result-Objekt
            $result = [PSCustomObject]@{
                FirstName         = $panelLeft.Children[0].Children[1].Text
                LastName          = $panelLeft.Children[1].Children[1].Text
                DisplayName       = $panelLeft.Children[3].Children[1].Text
                Description       = $panelLeft.Children[4].Children[1].Text
                OfficeRoom        = $panelLeft.Children[5].Children[1].Text
                PhoneNumber       = $panelLeft.Children[6].Children[1].Text
                MobileNumber      = $panelLeft.Children[7].Children[1].Text
                Position          = $panelLeft.Children[8].Children[1].Text
                DepartmentField   = $panelLeft.Children[9].Children[1].Text
                Ablaufdatum       = $panelLeft.Children[10].Children[1].Text
                External          = $externalChk.IsChecked
                Location          = $locCmb.SelectedItem
                CompanySection    = $cmbCompany.SelectedItem
                License           = $licCmb.SelectedItem
                PasswordMode      = 1  # 1 => Zufallspasswort generieren
                OutputTXT         = $chkTXT_Left.IsChecked
                EmailAddress      = $rowEmail.Children[1].Text
                MailSuffix        = $rowMailSuffix.Children[1].SelectedItem
                AccountDisabled   = $chkAccountDisabled.IsChecked
                PasswordNeverExpires = $chkPWNeverExpires.IsChecked
                SmartcardLogonRequired = $chkSmartcard.IsChecked
                CannotChangePassword   = $chkCannotChange.IsChecked
                MustChangePassword     = $chkMustChange.IsChecked

                CustomLogo        = $null
                ADGroupsSelected  = @()
                UPNEntered        = $rowUPN.Children[1].Text
                UPNFormat         = $rowTemplate.Children[1].SelectedItem
            }

            # CustomLogo:
            $result.CustomLogo = if ($global:CustomReportLogo) {
                $global:CustomReportLogo
            }
            else {
                if ($INIConfig.ContainsKey("Branding-Report") -and $INIConfig["Branding-Report"].ContainsKey("ReportLogo")) {
                    $INIConfig["Branding-Report"]["ReportLogo"]
                }
                else { $null }
            }

            # AD-Gruppen
            $selectedGroups = @()
            foreach ($key in $global:adGroupChecks.Keys) {
                if ($global:adGroupChecks[$key].IsChecked) {
                    $selectedGroups += $key
                }
            }
            $result.ADGroupsSelected = $selectedGroups

            # AD-Logik:
            Process-Onboarding -userData $result -Config $INIConfig

            $lblStatus.Content = "Onboarding successfully completed."
            $progressBar.Value = 100
        }
        catch {
            $lblStatus.Content = "Error: $($_.Exception.Message)"
            [System.Windows.MessageBox]::Show("Error during onboarding: $($_.Exception.Message)","Error",'OK','Error')
        }
    })
    $btnStack.Children.Add($btnOnboard) | Out-Null

    # CREATE PDF Button (Optionale PDF-Erzeugung)
    $btnPDF = New-Button "CREATE PDF"
    $btnPDF.Background = [System.Windows.Media.Brushes]::LightYellow
    $btnPDF.Add_Click({
        try {
            # Pfad für PDF-Script
            $scriptPath = "easyOnboarding_PDFCreator.ps1"
            if (Test-Path $scriptPath) {
                # Pfad aus [Branding-Report]->wkhtmltopdfPath laden
                $pdfExePath = $INIConfig["Branding-Report"]["wkhtmltopdfPath"]
                if (-not (Test-Path $pdfExePath)) {
                    Write-Error "wkhtmltopdf.exe not found at $pdfExePath"
                    return
                }
                # Beispiel: Start-Process pwsh ...
                Start-Process pwsh -ArgumentList "-File `"$scriptPath`" -htmlFile 'C:\temp\test.html' -pdfFile 'C:\temp\test.pdf' -wkhtmltopdfPath `"$pdfExePath`"" -NoNewWindow -Wait
            }
            else {
                [System.Windows.MessageBox]::Show("PDF-Script not found: $scriptPath","Info",'OK','Information')
            }
        }
        catch {
            [System.Windows.MessageBox]::Show("Error creating PDF: $($_.Exception.Message)","Error",'OK','Error')
        }
    })
    $btnStack.Children.Add($btnPDF) | Out-Null

    # INFO Button
    $btnInfo = New-Button "INFO"
    $btnInfo.Background = [System.Windows.Media.Brushes]::LightBlue
    $btnInfo.Add_Click({
        $infoFile = Join-Path (Split-Path $PSCommandPath) "info.txt"
        if (Test-Path $infoFile) {
            Start-Process notepad $infoFile
        }
        else {
            [System.Windows.MessageBox]::Show("info.txt not found.","Info",'OK','Information')
        }
    })
    $btnStack.Children.Add($btnInfo) | Out-Null

    # CLOSE Button
    $btnClose = New-Button "CLOSE"
    $btnClose.Background = [System.Windows.Media.Brushes]::LightCoral
    $btnClose.Add_Click({ $window.Close() })
    $btnStack.Children.Add($btnClose) | Out-Null

    # Footer (Website)
    $footerText = $guiBranding.ContainsKey("FooterWebseite") ? $guiBranding["FooterWebseite"] : "www.PSscripts.de"
    $lblFooterRight = New-Label $footerText
    $lblFooterRight.Margin = [System.Windows.Thickness]::new(5,20,5,5)
    $panelRight.Children.Add($lblFooterRight) | Out-Null

    # Anzeige:
    $window.Content = $mainGrid
    $null = $window.ShowDialog()
}

# ------------------------------------------------------------------------------
#  HAUPTTEIL - INI LADEN & GUI STARTEN
# ------------------------------------------------------------------------------
try {
    Write-Host "Loading INI: $ScriptINIPath"
    $global:Config = Get-IniContent -Path $ScriptINIPath
    Log-Message -Message "INI file loaded successfully." -Level Info
}
catch {PROJ_easyONBOARDING
    Throw "Error loading INI file: $_"
}

# DebugMode?
if ($global:Config.General["DebugMode"] -eq "1") {
    Write-Host "Debug mode enabled."
    # ggf. zusätzliche Debug-Ausgaben
}

# GUI starten
Show-OnboardingForm -INIConfig $global:Config

# SIG # Begin signature block
# MIIbywYJKoZIhvcNAQcCoIIbvDCCG7gCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDQr2SkjGyOivNa
# i7OpC5ziiQCHfs0BO6Kvs1bZ9j7slKCCFhcwggMQMIIB+KADAgECAhB3jzsyX9Cg
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
# DQEJBDEiBCCgXIIF2Qu4sRe9XkV6rYJt/U0uocKCGa9EzBY8zjdw/jANBgkqhkiG
# 9w0BAQEFAASCAQBJzpiynWHV/PBmVF/9y5kEQlAaSXQN/SX6prwLIHihM2pZXFzz
# p13N1Wv4yZZdPR1r2GTLnsd6S2lavay20kSSE8rG4t2iGkIVw9vPcEDmPodvXzkO
# kq/nJUtmXhOTTtRLqkJzisuMfOlHV9i7//ceYg+YQI+DvHqAsDd+Z3IMjrORKv+1
# glPRKan6KHjTlLMkQ6cZ4HZMHmGRnuNs3NRxrO6DRVndtss9PW9hYtknYs9VZRmn
# BQLyZ/Pk121A28JpPhb/JHqA9bs4MqoaF1nWZAcf2OUskyvQT+kyC9ILzAuGezCb
# a7MisvGiRFz9RkL9PoWFkPUSaKSKCwegAn37oYIDIDCCAxwGCSqGSIb3DQEJBjGC
# Aw0wggMJAgEBMHcwYzELMAkGA1UEBhMCVVMxFzAVBgNVBAoTDkRpZ2lDZXJ0LCBJ
# bmMuMTswOQYDVQQDEzJEaWdpQ2VydCBUcnVzdGVkIEc0IFJTQTQwOTYgU0hBMjU2
# IFRpbWVTdGFtcGluZyBDQQIQC65mvFq6f5WHxvnpBOMzBDANBglghkgBZQMEAgEF
# AKBpMBgGCSqGSIb3DQEJAzELBgkqhkiG9w0BBwEwHAYJKoZIhvcNAQkFMQ8XDTI1
# MDcwNTEwMTE1N1owLwYJKoZIhvcNAQkEMSIEIDWGbKbnWtbhz4vRA6vgWEflRadf
# Yz84/s5iwRSmavCIMA0GCSqGSIb3DQEBAQUABIICAD7UKBMKflfZMrm/ujMoTQQQ
# 4NnQCodrugf6hNmUQOSeGa1i50MhJqv6qMSypuWFW3Y1uxqoLeuQYnpLoGpW6laH
# YyI2BTVyt2X9puA08VfC7/t1gDdKeKHsTmkqgL7Xp8tCJOJtil8PQWqrs20Eo3eR
# o5OWLOZmSn/jGx93luBAVRs0MjMl0eOtEt41YXDF0cbmSS655UvbvMg9xyqJ7TuB
# hja6oenTskhvxf+1mjNPcKn5J8BNImm0JnsmoYimCngq9xJvJfvzrRQKVZZdcf90
# /ppGKvsnVu6s6H2d0SGdoAE9BFJoPmSoDGN2DpdCKONQZAV1ts5dWS+dIOYzwc4M
# 2UHi4rtCk1a7pxOshD/ZDxjos2CZ1FRz8HiDcQp8rcjcTj456MGjpV3mecXdsxQG
# YGx+Q9u8B16IGi45yEg3FnnxhbEbpcZj17gUfwWXUh3sOhhTUjSWx9EyjTdaVjYy
# Zh2I9m/ZNcJgpqzszoeGbLa1K157fgNS/RlIbcRT87+eQSL+woTBqCeF/rGqTwnd
# quxkSi80UKgKy5SGaa/QLtYxOW7xqWiwzvuF06X2LGkPwh7crCszgA40SD3dk8mU
# 1Q1wHN+lUKnyQXYDCZLX1zFwZRdRXGdiL8icoRIRMAoIKdpJSRCm1eO5WxHJxj1B
# e7sIkdSLgfNu5F7GlmoX
# SIG # End signature block
