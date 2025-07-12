#----------------------------------------------------
# 1 | INI-Datei einlesen
# Liest die INI-Datei und gibt ein OrderedDictionary zurück.
#----------------------------------------------------
#requires -Version 5.1  
[CmdletBinding()]
param(
    [string]$Vorname,
    [string]$Nachname,
    [string]$Standort,
    [string]$Company,         # z.B. "1" für Company1, "2" für Company2 etc.
    [string]$License = "",
    [switch]$Extern,
    [string]$ScriptINIPath = "easyONBOARDING*Config.ini"
)

#----------------------------------------------------
# 2 | Vorbereitung
# Initialisiert Assemblys für die Passwortgenerierung.
#----------------------------------------------------
Add-Type -AssemblyName System.Web

#----------------------------------------------------
# FUNKTION | INI-Datei einlesen
# Liest eine INI-Datei und gibt ein OrderedDictionary zurück.
#----------------------------------------------------
function Read-INIFile {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Path
    )
    if (-not (Test-Path $Path)) {
        Throw "INI-Datei nicht gefunden: $Path"
    }
    try {
        $iniContent = Get-Content -Path $Path | Where-Object { $_ -notmatch '^\s*[;#]' -and $_.Trim() -ne "" }
    }
    catch {
        Throw "Fehler beim Lesen der INI-Datei: $_"
    }
    $section = $null
    $result  = New-Object 'System.Collections.Specialized.OrderedDictionary'
    foreach ($line in $iniContent) {
        if ($line -match '^\[(.+)\]$') {
            $section = $matches[1].Trim()
            if (-not $result.Contains($section)) {
                $result[$section] = New-Object System.Collections.Specialized.OrderedDictionary
            }
        }
        elseif ($line -match '^(.*?)=(.*)$') {
            $key   = $matches[1].Trim()
            $value = $matches[2].Trim()
            if ($section -and $result[$section]) {
                $result[$section][$key] = $value
            }
        }
    }
    return $result
}

#----------------------------------------------------
# FUNKTION | GUI-Hilfsfunktionen
# Erzeugt Labels, Textboxen, CheckBoxen, ComboBoxen und horizontale Linien.
#----------------------------------------------------
function AddLabel {
    param(
        [Parameter(Mandatory=$true)] $parent,
        [string]$text,
        [int]$x,
        [int]$y,
        [switch]$Bold
    )
    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = $text
    $lbl.Location = New-Object System.Drawing.Point($x, $y)
    if ($Bold) {
        $lbl.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 8, [System.Drawing.FontStyle]::Bold)
    }
    else {
        $lbl.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 8)
    }
    $lbl.AutoSize = $true
    $parent.Controls.Add($lbl)
    return $lbl
}

function AddTextBox {
    param(
        [Parameter(Mandatory=$true)] $parent,
        [string]$default,
        [int]$x,
        [int]$y,
        [int]$width = 200
    )
    $tb = New-Object System.Windows.Forms.TextBox
    $tb.Text = $default
    $tb.Location = New-Object System.Drawing.Point($x, $y)
    $tb.Width = $width
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
    $cb = New-Object System.Windows.Forms.CheckBox
    $cb.Text = $text
    $cb.Location = New-Object System.Drawing.Point($x, $y)
    $cb.Checked = $checked
    $cb.AutoSize = $true
    $parent.Controls.Add($cb)
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
    if ($default -ne "" -and $cmb.Items.Contains($default)) {
        $cmb.SelectedItem = $default
    }
    elseif ($cmb.Items.Count -gt 0) {
        $cmb.SelectedIndex = 0
    }
    $parent.Controls.Add($cmb)
    return $cmb
}

function Build-EmailAddress {
    param(
        [string]$inputEmail,
        [string]$companyMailDomain,
        [object]$mailSuffixItem
    )
    $email = $inputEmail.Trim()
    if ($email -ne "" -and $email -notmatch "@") {
        if (-not [string]::IsNullOrWhiteSpace($companyMailDomain)) {
            $email = "$email$companyMailDomain"
        }
        elseif (-not [string]::IsNullOrWhiteSpace($mailSuffixItem)) {
            $email = "$email$($mailSuffixItem)"
        }
    }
    return $email
}

function Add-HorizontalLine {
    param(
        [Parameter(Mandatory=$true)]
        [System.Windows.Forms.Panel]$Parent,
        [Parameter(Mandatory=$true)]
        [int]$y,                      # Y-Position der Linie
        [int]$x = 10,                 # X-Startposition (Standard: 10)
        [int]$width = 0,              # Breite der Linie (Standard: $Parent.Width - 20)
        [int]$height = 1,             # Höhe der Linie (Standard: 1 Pixel)
        [System.Drawing.Color]$color = [System.Drawing.Color]::Gray  # Farbe der Linie
    )
    if ($width -eq 0) {
        $width = $Parent.Width - 20
    }
    $line = New-Object System.Windows.Forms.Panel
    $line.Size = New-Object System.Drawing.Size($width, $height)
    $line.Location = New-Object System.Drawing.Point($x, $y)
    $line.BackColor = $color
    $Parent.Controls.Add($line)
    return $line
}

#----------------------------------------------------
# FUNKTION | Erweiterte Passwortgenerierung
# Erzeugt ein Passwort anhand der übergebenen Parameter.
#----------------------------------------------------
function Generate-AdvancedPassword {
    param(
        [int]$Laenge = 12,
        [bool]$IncludeSpecial = $true,
        [bool]$AvoidAmbiguous = $true,
        [int]$MinUpperCase = 2,
        [int]$MinDigits = 2
    )
    # Beispiel: Erzeugt zunächst ein Passwort mittels der Membership-Methode.
    $pw = [System.Web.Security.Membership]::GeneratePassword($Laenge, 2)
    if ($AvoidAmbiguous) {
        $pw = $pw -replace '[{}()\[\]\/\\~,;:.<>\"]','X'
    }
    # Zusätzliche Prüfungen (z. B. Mindestanzahl Großbuchstaben oder Ziffern) können hier ergänzt werden.
    return $pw
}

#----------------------------------------------------
# 3 | Onboarding verarbeiten
# Führt AD-Erstellung, Logging und Report-Erzeugung durch.
#----------------------------------------------------
function Process-Onboarding {
    param(
        [Parameter(Mandatory=$true)] [pscustomobject]$userData,
        [Parameter(Mandatory=$true)] [hashtable]$Config
    )
    # AD-Modul laden
    try {
        Import-Module ActiveDirectory -ErrorAction Stop
    }
    catch {
        Throw "AD-Modul konnte nicht geladen werden: $($_.Exception.Message)"
    }
    if ([string]::IsNullOrWhiteSpace($userData.Vorname)) {
        Throw "Vorname muss eingegeben werden!"
    }

    ## Passwortgenerierung – Werte ausschließlich aus der INI (ohne Fallback)
    if ($userData.PasswordMode -eq 1) {
        if (-not $Config.PasswordFixGenerate.Contains("PasswordLaenge")) { 
            Throw "Fehler: 'PasswordLaenge' fehlt im Abschnitt [PasswordFixGenerate] der INI!" 
        }
        if (-not $Config.PasswordFixGenerate.Contains("IncludeSpecialChars")) { 
            Throw "Fehler: 'IncludeSpecialChars' fehlt im Abschnitt [PasswordFixGenerate] der INI!" 
        }
        if (-not $Config.PasswordFixGenerate.Contains("AvoidAmbiguousChars")) { 
            Throw "Fehler: 'AvoidAmbiguousChars' fehlt im Abschnitt [PasswordFixGenerate] der INI!" 
        }
        if (-not $Config.PasswordFixGenerate.Contains("MinNonAlpha")) { 
            Throw "Fehler: 'MinNonAlpha' fehlt im Abschnitt [PasswordFixGenerate] der INI!" 
        }
        if (-not $Config.PasswordFixGenerate.Contains("MinUpperCase")) { 
            Throw "Fehler: 'MinUpperCase' fehlt im Abschnitt [PasswordFixGenerate] der INI!" 
        }
        if (-not $Config.PasswordFixGenerate.Contains("MinDigits")) { 
            Throw "Fehler: 'MinDigits' fehlt im Abschnitt [PasswordFixGenerate] der INI!" 
        }

        $UserPW = Generate-AdvancedPassword -Laenge ([int]$Config.PasswordFixGenerate["PasswordLaenge"]) `
                                              -IncludeSpecial ([bool]$Config.PasswordFixGenerate["IncludeSpecialChars"]) `
                                              -AvoidAmbiguous ([bool]$Config.PasswordFixGenerate["AvoidAmbiguousChars"]) `
                                              -MinUpperCase ([int]$Config.PasswordFixGenerate["MinUpperCase"]) `
                                              -MinDigits ([int]$Config.PasswordFixGenerate["MinDigits"])
        if ([string]::IsNullOrWhiteSpace($UserPW)) {
            Throw "Fehler: Das generierte Passwort ist leer. Bitte überprüfen Sie die INI-Einstellungen im Abschnitt [PasswordFixGenerate]."
        }
    }
    else {
        if (-not $Config.PasswordFixGenerate.Contains("fixPassword")) { 
            Throw "Fehler: 'fixPassword' fehlt im Abschnitt [PasswordFixGenerate] der INI!" 
        }
        $UserPW = $Config.PasswordFixGenerate["fixPassword"]
        if ([string]::IsNullOrWhiteSpace($UserPW)) {
            Throw "Fehler: Kein Fix-Passwort in der INI angegeben!"
        }
    }
    $SecurePW = ConvertTo-SecureString $UserPW -AsPlainText -Force

    $SamAccountName = ($userData.Vorname.Substring(0,1) + $userData.Nachname).ToLower()
    if ($userData.UPNEntered) {
        $UPN = $userData.UPNEntered
    }
    else {
        # Unternehmensdaten laden
        $companySection = $userData.CompanySection.Section
        if (-not $Config.Contains($companySection)) {
            Throw "Fehler: Die Sektion '$companySection' existiert nicht in der INI!"
        }
        $companyData = $Config[$companySection]
        $suffix = ($companySection -replace "\D", "")
        if (-not $companyData.Contains("ActiveDirectoryDomain$suffix") -or -not $companyData["ActiveDirectoryDomain$suffix"]) {
            Throw "Fehler: 'ActiveDirectoryDomain$suffix' fehlt oder ist leer in der INI!"
        }
        $adDomain = "@" + $companyData["ActiveDirectoryDomain$suffix"].Trim()
        switch -Wildcard ($userData.UPNFormat) {
            "VORNAME.NACHNAME"    { $UPN = "$($userData.Vorname).$($userData.Nachname)$adDomain" }
            "V.NACHNAME"          { $UPN = "$($userData.Vorname.Substring(0,1)).$($userData.Nachname)$adDomain" }
            "VORNAMENACHNAME"     { $UPN = "$($userData.Vorname)$($userData.Nachname)$adDomain" }
            "VNACHNAME"           { $UPN = "$($userData.Vorname.Substring(0,1))$($userData.Nachname)$adDomain" }
            Default               { $UPN = "$SamAccountName$adDomain" }
        }
    }

    Write-Host "DisplayName: $($userData.DisplayName)"
    Write-Host "SamAccountName: $SamAccountName"
    Write-Host "UPN: $UPN"

    try {
        $existingUser = Get-ADUser -Filter { SamAccountName -eq $SamAccountName } -ErrorAction SilentlyContinue
    }
    catch {
        $existingUser = $null
    }
    
    # Laden weiterer Firmen- und Logindaten
    $companySection = $userData.CompanySection.Section
    $companyData = $Config[$companySection]
    $suffix = ($companySection -replace "\D", "")
    if (-not $companyData.Contains("Strasse$suffix")) { Throw "Fehler: 'Strasse$suffix' fehlt in der INI!" }
    $Strasse = $companyData["Strasse$suffix"]
    if (-not $companyData.Contains("PLZ$suffix")) { Throw "Fehler: 'PLZ$suffix' fehlt in der INI!" }
    $PLZ     = $companyData["PLZ$suffix"]
    if (-not $companyData.Contains("Ort$suffix")) { Throw "Fehler: 'Ort$suffix' fehlt in der INI!" }
    $Ort     = $companyData["Ort$suffix"]
    if (-not $userData.MailSuffix -or $userData.MailSuffix -eq "") {
        if (-not $Config.Contains("MailEndungen") -or -not $Config.MailEndungen.Contains("Domain1")) {
            Throw "Fehler: Mail-Endungen sind in der INI nicht vollständig definiert!"
        }
        $userData.MailSuffix = $Config.MailEndungen["Domain1"]
    }
    if (-not $companyData.Contains("Country$suffix") -or -not $companyData["Country$suffix"]) {
        Throw "Fehler: 'Country$suffix' fehlt in der INI!"
    }
    $Country = $companyData["Country$suffix"]
    if (-not $companyData.Contains("NameFirma$suffix") -or -not $companyData["NameFirma$suffix"]) {
        Throw "Fehler: 'NameFirma$suffix' fehlt in der INI!"
    }
    $companyDisplay = $companyData["NameFirma$suffix"]

    if (-not $Config.General.Contains("DefaultOU") -or -not $Config.General["DefaultOU"]) {
        Throw "Fehler: 'DefaultOU' fehlt in der INI!"
    }
    $defaultOU = $Config.General["DefaultOU"]

    if (-not $Config.General.Contains("LogFilePath") -or -not $Config.General["LogFilePath"]) {
        Throw "Fehler: 'LogFilePath' fehlt in der INI!"
    }
    $logFilePath = $Config.General["LogFilePath"]

    if (-not $Config.General.Contains("ReportPath") -or -not $Config.General["ReportPath"]) {
        Throw "Fehler: 'ReportPath' fehlt in der INI!"
    }
    $reportPath = $Config.General["ReportPath"]

    if (-not $Config.General.Contains("ReportTitle") -or -not $Config.General["ReportTitle"]) {
        Throw "Fehler: 'ReportTitle' fehlt in der INI!"
    }
    $reportTitle = $Config.General["ReportTitle"]

    if (-not $Config.General.Contains("ReportFooter") -or -not $Config.General["ReportFooter"]) {
        Throw "Fehler: 'ReportFooter' fehlt in der INI!"
    }
    $reportFooter = $Config.General["ReportFooter"]

    if (-not $Config.Contains("Branding-Report") -or -not $Config["Branding-Report"].Contains("ReportFooter") -or -not $Config["Branding-Report"]["ReportFooter"]) {
        Throw "Fehler: 'ReportFooter' fehlt im Abschnitt [Branding-Report] der INI!"
    }
    $reportBranding = $Config["Branding-Report"]
    $finalReportFooter = $reportBranding["ReportFooter"]

    #----------------------------------------------------
    # 6 | AD-Benutzer anlegen/aktualisieren
    # Legt einen neuen AD-Benutzer an oder aktualisiert einen bestehenden.
    #----------------------------------------------------
    $userParams = @{
        Name                  = $userData.DisplayName
        DisplayName           = $userData.DisplayName
        GivenName             = $userData.Vorname
        Surname               = $userData.Nachname
        SamAccountName        = $SamAccountName
        UserPrincipalName     = $UPN
        AccountPassword       = $SecurePW
        Enabled               = (-not $userData.AccountDisabled)
        ChangePasswordAtLogon = $userData.MustChangePassword
        PasswordNeverExpires  = $userData.PasswordNeverExpires
        Path                  = $defaultOU
        City                  = $Ort
        StreetAddress         = $Strasse
        Country               = $Country
    }
    if ($Config.ADUserDefaults.Contains("LogonScript")) {
        $userParams["ScriptPath"] = $Config.ADUserDefaults["LogonScript"]
    }

    Write-Host "Benutzerparameter (Debug):"
    $userParams.GetEnumerator() | ForEach-Object { Write-Host "$($_.Key) = $($_.Value)" }

    if (-not $existingUser) {
        Write-Host "Erstelle neuen Benutzer: $($userData.DisplayName)"
        $otherAttributes = @{}
        if ($userData.EmailAddress -and $userData.EmailAddress.Trim() -ne "") {
            if ($userData.EmailAddress -notmatch "@") {
                $otherAttributes["mail"] = "$($userData.EmailAddress)$($userData.MailSuffix)"
            }
            else {
                $otherAttributes["mail"] = $userData.EmailAddress
            }
        }
        if ($userData.Description) { $otherAttributes["description"] = $userData.Description }
        if ($userData.OfficeRoom)  { $otherAttributes["physicalDeliveryOfficeName"] = $userData.OfficeRoom }
        if ($userData.PhoneNumber) { $otherAttributes["telephoneNumber"] = $userData.PhoneNumber }
        if ($userData.MobileNumber){ $otherAttributes["mobile"] = $userData.MobileNumber }
        if ($userData.Position)    { $otherAttributes["title"] = $userData.Position }
        if ($userData.DepartmentField) { $otherAttributes["department"] = $userData.DepartmentField }
        $filteredAttrs = @{}
        foreach ($k in $otherAttributes.Keys) {
            if ($otherAttributes[$k].Trim() -ne "") {
                $filteredAttrs[$k] = $otherAttributes[$k]
            }
        }
        if ($filteredAttrs.Count -gt 0) {
            $userParams["OtherAttributes"] = $filteredAttrs
        }
        try {
            try {
                New-ADUser @userParams -ErrorAction Stop
                Write-Host "AD-Benutzer erstellt."
            }
            catch {
                $line = $_.InvocationInfo.ScriptLineNumber
                $function = $_.InvocationInfo.FunctionName
                $script = $_.InvocationInfo.ScriptName
                $errorDetails = $_.Exception.ToString()
                $detailedMessage = "Fehler beim Erstellen des Benutzers im Skript '$script', Funktion '$function', Zeile ${line}: $errorDetails"
                Throw $detailedMessage
            }
        }
        catch {
            Throw $_.Exception.Message
        }
        if ($userData.SmartcardLogonRequired) {
            try { Set-ADUser -Identity $SamAccountName -SmartcardLogonRequired $true } catch {}
        }
        if ($userData.CannotChangePassword) {
            Write-Host "(Hinweis: 'CannotChangePassword' via ACL müsste hier umgesetzt werden.)"
        }
    }
    else {
        Write-Host "Benutzer '$SamAccountName' existiert bereits - Update erfolgt."
        try {
            Set-ADUser -Identity $existingUser.DistinguishedName `
                -GivenName $userData.Vorname `
                -Surname $userData.Nachname `
                -City $Ort `
                -StreetAddress $Strasse `
                -Country $Country `
                -Enabled (-not $userData.AccountDisabled) `
                -ErrorAction SilentlyContinue
            Set-ADUser -Identity $existingUser.DistinguishedName -ChangePasswordAtLogon:$userData.MustChangePassword -PasswordNeverExpires:$userData.PasswordNeverExpires
        }
        catch {
            Write-Warning "Fehler beim Aktualisieren: $($_.Exception.Message)"
        }
    }
    try {
        Set-ADAccountPassword -Identity $SamAccountName -Reset -NewPassword $SecurePW -ErrorAction SilentlyContinue
    }
    catch {
        Write-Warning "Fehler beim Setzen des Passworts: $($_.Exception.Message)"
    }

    #----------------------------------------------------
    # 7 | AD-Gruppen zuweisen
    # Weist den Benutzer den AD-Gruppen zu, sofern nicht als externer Mitarbeiter markiert.
    #----------------------------------------------------
    if (-not $userData.Extern) {
        foreach ($gKey in $userData.ADGroupsSelected) {
            $gName = $Config.ADGroups[$gKey]
            if ($gName) {
                try { Add-ADGroupMember -Identity $gName -Members $SamAccountName } catch {
                    Write-Warning "Fehler bei AD-Gruppe '$gName': $($_.Exception.Message)"
                }
            }
        }
    }
    else {
        Write-Host "Externer Mitarbeiter: Standardmäßige AD-Gruppen-Zuweisung wird übersprungen."
    }
    if ($userData.Location) {
        $signaturKey = $Config.STANDORTE[$userData.Location]
        if ($signaturKey) {
            $signaturGroup = $Config.SignaturGruppe_Optional[$signaturKey]
            if ($signaturGroup) {
                try { Add-ADGroupMember -Identity $signaturGroup -Members $SamAccountName } catch {}
            }
        }
    }
    if ($License) {
        $licKey = "MS365_" + $License
        $licGroup = $Config.LicensesGroups[$licKey]
        if ($licGroup) {
            try { Add-ADGroupMember -Identity $licGroup -Members $SamAccountName } catch {}
        }
    }
    if ($Config.ActivateUserMS365ADSync["ADSync"] -eq '1') {
        $adSyncGroup = $Config.ActivateUserMS365ADSync["ADSyncADGroup"]
        try { Add-ADGroupMember -Identity $adSyncGroup -Members $SamAccountName } catch {}
    }
    if ($userData.Extern) {
        Write-Host "Externer Mitarbeiter: Bitte weisen Sie alle AD-Gruppen händisch zu."
    }

    #----------------------------------------------------
    # 8 | Logging
    # Schreibt einen Logeintrag in die definierte Logdatei.
    #----------------------------------------------------
    try {
        if (-not (Test-Path $logFilePath)) {
            New-Item -ItemType Directory -Path $logFilePath -Force | Out-Null
        }
        $logDate = (Get-Date -Format 'yyyyMMdd')
        $logFile = Join-Path $logFilePath "Onboarding_$logDate.log"
        $currentUser = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
        $logEntry = "ONBOARDING DURCHGEFÜHRT VON: $currentUser`r`n" +
                    ("[{0}] Sam={1}, Anzeigename='{2}', UPN='{3}', Standort='{4}', Firma='{5}', MS365 Lizenz='{6}', ADGruppen=({7}), Passwort='{8}', Extern={9}" -f (Get-Date), $SamAccountName, $userData.DisplayName, $UPN, $userData.Location, $companyDisplay, $userData.MS365License, ($userData.ADGroupsSelected -join ','), $UserPW, $userData.Extern)
        Add-Content -Path $logFile -Value $logEntry
        Write-Host "Log geschrieben: $logFile"
    }
    catch {
        Write-Warning "Fehler beim Logging: $($_.Exception.Message)"
    }

    #----------------------------------------------------
    # 9 | Reports erzeugen
    # Erzeugt HTML- und TXT-Reports; PDF wird über einen separaten Button erstellt.
    #----------------------------------------------------
    try {
        if (-not (Test-Path $reportPath)) {
            New-Item -ItemType Directory -Path $reportPath -Force | Out-Null
        }
        $htmlFile = Join-Path $reportPath "$SamAccountName.html"
        $htmlTemplatePath = $reportBranding["TemplatePath"]
        if (-not $htmlTemplatePath -or -not (Test-Path $htmlTemplatePath)) {
            $htmlTemplate = "INI BZW. HTMLTemplate nicht angegeben!"
        }
        else {
            $htmlTemplate = Get-Content -Path $htmlTemplatePath -Raw
        }
        $htmlContent = $htmlTemplate `
            -replace "{{ReportTitle}}", ([string]$reportTitle) `
            -replace "{{Admin}}", $env:USERNAME `
            -replace "{{ReportDate}}", (Get-Date -Format "yyyy-MM-dd") `
            -replace "{{Vorname}}", $userData.Vorname `
            -replace "{{Nachname}}", $userData.Nachname `
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
            -replace "{{LogoTag}}", ""
        Set-Content -Path $htmlFile -Value $htmlContent -Encoding UTF8
        Write-Host "HTML-Report erstellt: $htmlFile"

        if ($userData.OutputTXT) {
            $txtFile = Join-Path $reportPath "$SamAccountName.txt"
            $txtContent = @"
Onboarding Report für neue Mitarbeiter
Erstellt von: $($env:USERNAME)
Datum: $(Get-Date -Format 'yyyy-MM-dd')

Benutzerdetails:
----------------
Vorname:      $($userData.Vorname)
Nachname:     $($userData.Nachname)
Anzeigename:  $($userData.DisplayName)
Externer MA:  $($userData.Extern)
Beschreibung: $($userData.Description)
Büro:         $($userData.OfficeRoom)
Rufnummer:    $($userData.PhoneNumber)
Mobil:        $($userData.MobileNumber)
Position:     $($userData.Position)
Abteilung:    $($userData.DepartmentField)
LoginName:    $SamAccountName
Passwort:     $UserPW

Weiterführende Links:
---------------------
"@
            if ($Config.Contains("Websites")) {
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
            Write-Host "TXT-Report erstellt: $txtFile"
        }
    }
    catch {
        Write-Warning "Fehler beim Erstellen der Reports: $($_.Exception.Message)"
    }
}

#----------------------------------------------------
# 4 | GUI-Erstellung: Onboarding-Form
# Erstellt das GUI-Formular für das Onboarding.
#----------------------------------------------------
function Show-OnboardingForm {
    param(
        [hashtable]$INIConfig
    )
    $result = [PSCustomObject]@{
        Vorname               = ""
        Nachname              = ""
        DisplayName           = ""
        Description           = ""
        OfficeRoom            = ""
        PhoneNumber           = ""
        MobileNumber          = ""
        Position              = ""
        DepartmentField       = ""
        Location              = ""
        CompanySection        = $null
        MS365License          = ""
        PasswordNeverExpires  = $false
        MustChangePassword    = $false
        AccountDisabled       = $false
        CannotChangePassword  = $false
        PasswordMode          = 1
        FixPassword           = ""
        PasswordLaenge        = 12
        IncludeSpecialChars   = $true
        AvoidAmbiguousChars   = $true
        OutputHTML            = $true
        OutputPDF             = $false
        OutputTXT             = $false
        UPNEntered            = ""
        UPNFormat             = ""
        EmailAddress          = ""
        MailSuffix            = ""
        Ablaufdatum           = ""
        Cancel                = $false
        ADGroupsSelected      = @()
        Extern                = $false
        SmartcardLogonRequired= $false
    }

    $currentUser = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name

    # Branding aus der INI
    $guiBranding    = if ($INIConfig.Contains("Branding-GUI")) { $INIConfig["Branding-GUI"] } else { @{} }
    $reportBranding = if ($INIConfig.Contains("Branding-Report")) { $INIConfig["Branding-Report"] } else { @{} }

    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    #----------------------------------------------------
    # Hauptfenster erstellen
    #----------------------------------------------------
    $form = New-Object System.Windows.Forms.Form
    $form.Text = if ($guiBranding.APPName) { $guiBranding.APPName } else { "easyONBOARDING" }
    $form.StartPosition = "CenterScreen"
    $form.Size = New-Object System.Drawing.Size(1085,1025)
    $form.AutoScroll = $true
    if ($guiBranding.BackgroundImage -and (Test-Path $guiBranding.BackgroundImage)) {
        $form.BackgroundImage = [System.Drawing.Image]::FromFile($guiBranding.BackgroundImage)
        $form.BackgroundImageLayout = 'Stretch'
    }

    #----------------------------------------------------
    # Überschrift/Info anzeigen
    #----------------------------------------------------
    $lblInfo = New-Object System.Windows.Forms.Label
    $lblInfo.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 10, [System.Drawing.FontStyle]::Bold)
    $lblInfo.Location = New-Object System.Drawing.Point(10,10)
    $lblInfo.AutoSize = $true
    $scriptInfo = if ($INIConfig.Contains("ScriptInfo")) { $INIConfig["ScriptInfo"] } else { @{} }
    $general     = if ($INIConfig.Contains("General")) { $INIConfig["General"] } else { @{} }
    $lblInfo.Text = "ScriptVersion=$($scriptInfo.ScriptVersion) | LastUpdate=$($scriptInfo.LastUpdate) | Author=$($scriptInfo.Author)`r`n" +
                    "ONBOARDING DURCHGEFÜHRT VON: $currentUser`r`n" +
                    "DOMAIN: $($general.DomainName1) | OU: $($general.DefaultOU) | REPORT: $($general.ReportPath)"
    $form.Controls.Add($lblInfo)

    #----------------------------------------------------
    # Header-Logo anzeigen
    #----------------------------------------------------
    $picHeaderLogo = New-Object System.Windows.Forms.PictureBox
    $picHeaderLogo.Size = New-Object System.Drawing.Size(125,50)
    $picHeaderLogo.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::StretchImage
    $picHeaderLogo.Location = New-Object System.Drawing.Point(($form.ClientSize.Width - 125 - 10), 10)
    if ($guiBranding.HeaderLogo -and (Test-Path $guiBranding.HeaderLogo)) {
        $picHeaderLogo.Image = [System.Drawing.Image]::FromFile($guiBranding.HeaderLogo)
    }
    if ($guiBranding.HeaderLogoURL) {
        $picHeaderLogo.Add_Click({ Start-Process $guiBranding.HeaderLogoURL })
    }
    $form.Controls.Add($picHeaderLogo)

    #----------------------------------------------------
    # Linkes Panel (Datenformular) erstellen
    #----------------------------------------------------
    $panelLeft = New-Object System.Windows.Forms.Panel
    $panelLeft.Location = New-Object System.Drawing.Point(10,80)
    $panelLeft.Size = New-Object System.Drawing.Size(520,700)
    $panelLeft.AutoScroll = $true
    $panelLeft.BorderStyle = 'FixedSingle'
    $panelLeft.BackColor = [System.Drawing.Color]::WhiteSmoke
    $form.Controls.Add($panelLeft)

    #----------------------------------------------------
    # Rechtes Panel (weitere Einstellungen) erstellen
    #----------------------------------------------------
    $panelRight = New-Object System.Windows.Forms.Panel
    $panelRight.Location = New-Object System.Drawing.Point(540,80)
    $panelRight.Size = New-Object System.Drawing.Size(520,700)
    $panelRight.AutoScroll = $true
    $panelRight.BorderStyle = 'FixedSingle'
    $panelRight.BackColor = [System.Drawing.Color]::WhiteSmoke
    $form.Controls.Add($panelRight)

    #----------------------------------------------------
    # Unteres Panel (Buttons & Status) erstellen
    #----------------------------------------------------
    $panelBottom = New-Object System.Windows.Forms.Panel
    $panelBottom.Dock = 'Bottom'
    $panelBottom.Height = 150
    $panelBottom.BorderStyle = 'None'
    $form.Controls.Add($panelBottom)

    $panelFooter = New-Object System.Windows.Forms.Panel
    $panelFooter.Dock = 'Bottom'
    $panelFooter.Height = 40
    $panelFooter.BorderStyle = 'FixedSingle'
    $lblFooter = New-Object System.Windows.Forms.Label
    $lblFooter.Location = New-Object System.Drawing.Point(10,10)
    $lblFooter.AutoSize = $true
    $lblFooter.Text = if ($guiBranding.FooterWebseite) { $guiBranding.FooterWebseite } else { "www.easyONBOARDING.com" }
    $panelFooter.Controls.Add($lblFooter)
    $form.Controls.Add($panelFooter)

    #----------------------------------------------------
    # Fortschrittsbalken hinzufügen
    #----------------------------------------------------
    $progressBar = New-Object System.Windows.Forms.ProgressBar
    $progressBar.Location = New-Object System.Drawing.Point(10, 30)
    $progressBar.Size = New-Object System.Drawing.Size(1050, 20)
    $progressBar.Minimum = 0
    $progressBar.Maximum = 100
    $progressBar.Value = 0
    $panelBottom.Controls.Add($progressBar)

    #----------------------------------------------------
    # (Optional) Status-Label erstellen
    #----------------------------------------------------
    $lblStatus = New-Object System.Windows.Forms.Label
    $lblStatus.Location = New-Object System.Drawing.Point(10, 5)
    $lblStatus.Size = New-Object System.Drawing.Size(500, 20)
    $lblStatus.Text = "Bereit..."
    $panelBottom.Controls.Add($lblStatus)

    #----------------------------------------------------
    # Elemente im linken Panel: Formulareingaben
    #----------------------------------------------------
    $yLeft = 10
    AddLabel $panelLeft "Vorname:" 10 $yLeft -Bold | Out-Null
    $txtVorname = AddTextBox $panelLeft "" 150 $yLeft; $yLeft += 30

    AddLabel $panelLeft "Nachname:" 10 $yLeft -Bold | Out-Null
    $txtNachname = AddTextBox $panelLeft "" 150 $yLeft; $yLeft += 40

    AddLabel $panelLeft "Externer Mitarbeiter:" 10 $yLeft -Bold | Out-Null
    $chkExternal = AddCheckBox $panelLeft "" $false 150 $yLeft; $yLeft += 35

    AddLabel $panelLeft "Anzeigename:" 10 $yLeft -Bold | Out-Null
    $txtDisplayName = AddTextBox $panelLeft "" 150 $yLeft; $yLeft += 30

    AddLabel $panelLeft "Anzeigename Vorlage:" 10 $yLeft -Bold | Out-Null
    $templates = @()
    if ($INIConfig.Contains("DisplayNameTemplates")) {
        $templates = $INIConfig["DisplayNameTemplates"].Keys | ForEach-Object { $INIConfig["DisplayNameTemplates"][$_] }
    }
    $cmbDisplayNameTemplate = AddComboBox $panelLeft $templates 150 $yLeft 250 ""; $yLeft += 40

    AddLabel $panelLeft "Beschreibung:" 10 $yLeft -Bold | Out-Null
    $txtDescription = AddTextBox $panelLeft "" 150 $yLeft; $yLeft += 30

    AddLabel $panelLeft "Buero:" 10 $yLeft -Bold | Out-Null
    $txtOffice = AddTextBox $panelLeft "" 150 $yLeft; $yLeft += 40

    AddLabel $panelLeft "Rufnummer:" 10 $yLeft -Bold | Out-Null
    $txtPhone = AddTextBox $panelLeft "" 150 $yLeft; $yLeft += 30

    AddLabel $panelLeft "Mobil:" 10 $yLeft -Bold | Out-Null
    $txtMobile = AddTextBox $panelLeft "" 150 $yLeft; $yLeft += 40

    AddLabel $panelLeft "Position:" 10 $yLeft -Bold | Out-Null
    $txtPosition = AddTextBox $panelLeft "" 150 $yLeft; $yLeft += 30

    AddLabel $panelLeft "Abteilung (manuell):" 10 $yLeft -Bold | Out-Null
    $txtDeptField = AddTextBox $panelLeft "" 150 $yLeft; $yLeft += 40

    AddLabel $panelLeft "Austrittsdatum:" 10 $yLeft -Bold | Out-Null
    $txtAblaufdatum = AddTextBox $panelLeft "" 150 $yLeft; $yLeft += 50

    $lineAbove = Add-HorizontalLine -Parent $panelLeft -y $yLeft
    $yLeft += 20

    AddLabel $panelLeft "STANDORT*:" 10 $yLeft -Bold | Out-Null
    $locationDisplayList = $INIConfig.STANDORTE.Keys | Where-Object { $_ -match '_Bez$' } | ForEach-Object { $INIConfig.STANDORTE[$_] }
    $cmbLocation = AddComboBox $panelLeft $locationDisplayList 150 $yLeft 250 ""; $yLeft += 30

    AddLabel $panelLeft "FIRMA:" 10 $yLeft -Bold | Out-Null
    $companyOptions = @()
    foreach ($section in $INIConfig.Keys | Where-Object { $_ -like "Company*" }) {
        $suffix = ($section -replace "\D", "")
        $visibleKey = "$section_Visible"
        if ($INIConfig[$section].Contains($visibleKey)) {
            if ($INIConfig[$section][$visibleKey] -ne "1") { continue }
        }
        if ($INIConfig[$section].Contains("NameFirma$suffix") -and -not [string]::IsNullOrWhiteSpace($INIConfig[$section]["NameFirma$suffix"])) {
            $display = $INIConfig[$section]["NameFirma$suffix"].Trim()
            $companyOptions += [PSCustomObject]@{ Display = $display; Section = $section }
        }
    }
    $cmbCompany = New-Object System.Windows.Forms.ComboBox
    $cmbCompany.DropDownStyle = 'DropDownList'
    $cmbCompany.FormattingEnabled = $true
    $cmbCompany.Location = New-Object System.Drawing.Point(150, $yLeft)
    $cmbCompany.Width = 250
    $cmbCompany.DataSource = $companyOptions
    $cmbCompany.DisplayMember = "Display"
    $cmbCompany.ValueMember = "Section"
    $panelLeft.Controls.Add($cmbCompany)
    $yLeft += 40

    $lineAbove = Add-HorizontalLine -Parent $panelLeft -y $yLeft
    $yLeft += 20

    AddLabel $panelLeft "MS365 Lizenz*:" 10 $yLeft -Bold | Out-Null
    $cmbMS365License = AddComboBox $panelLeft ( @("KEINE") + ($INIConfig.LicensesGroups.Keys | ForEach-Object { $_ -replace '^MS365_','' } ) ) 150 $yLeft 200 ""
    $yLeft += 55

    $lineAbove = Add-HorizontalLine -Parent $panelLeft -y $yLeft
    $yLeft += 20

    #----------------------------------------------------
    # Abschnitt | ONBOARDING DOKUMENT ERSTELLEN
    # Zwei Zeilen: oberhalb X (HTML) und darunter TXT
    #----------------------------------------------------
    AddLabel $panelLeft "ONBOARDING DOKUMENT ERSTELLEN:" 10 $yLeft -Bold | Out-Null
    $yLeft += 20
    AddLabel $panelLeft "upn.HTML" 150 $yLeft -Bold | Out-Null
    $lblHTML = AddLabel $panelLeft "X" 250 $yLeft
    $lblHTML.ForeColor = [System.Drawing.Color]::Gray
    $yLeft += 20
    AddLabel $panelLeft "upn.TXT" 150 $yLeft -Bold | Out-Null
    $chkTXT_Left = AddCheckBox $panelLeft "" $true 250 $yLeft
    $yLeft += 35

    #----------------------------------------------------
    # Elemente im rechten Panel: Weitere Einstellungen
    #----------------------------------------------------
    $yRight = 10
    AddLabel $panelRight "Benutzer Name (UPN):" 10 $yRight -Bold | Out-Null
    $txtUPN = AddTextBox $panelRight "" 150 $yRight 200; $yRight += 35

    AddLabel $panelRight "UPN-Format-Vorlagen:" 10 $yRight -Bold | Out-Null
    $cmbUPNFormat = AddComboBox $panelRight @("VORNAME.NACHNAME","V.NACHNAME","VORNAMENACHNAME","VNACHNAME") 150 $yRight 200; $yRight += 50

    AddLabel $panelRight "E-Mail-Adresse:" 10 $yRight -Bold | Out-Null
    $txtEmail = AddTextBox $panelRight "" 150 $yRight 200; $yRight += 35

    AddLabel $panelRight "Mail-Endung:" 10 $yRight -Bold | Out-Null
    $cmbMailSuffix = AddComboBox $panelRight @() 150 $yRight 250 ""
    if ($INIConfig.Contains("MailEndungen")) {
        foreach ($key in $INIConfig.MailEndungen.Keys) {
            [void]$cmbMailSuffix.Items.Add($INIConfig.MailEndungen[$key])
        }
        if ($cmbMailSuffix.Items.Count -gt 0) {
            $cmbMailSuffix.SelectedIndex = 0
        }
    }
    $yRight += 40

    $lineAbove = Add-HorizontalLine -Parent $panelRight -y $yRight
    $yRight += 10

    #----------------------------------------------------
    # Abschnitt | AD-Benutzer-Flags
    # Zwei Zeilen, je eine Option pro Zeile.
    #----------------------------------------------------
    AddLabel $panelRight "AD-Benutzer-Flags:" 10 $yRight -Bold | Out-Null
    $yRight += 20
    $chkAccountDisabled = AddCheckBox $panelRight "Konto deaktiviert" $true 150 $yRight
    $yRight += 20
    $chkPWNeverExpires  = AddCheckBox $panelRight "Passwort laeuft nicht ab" $false 150 $yRight
    $yRight += 20
    $chkMustChange      = AddCheckBox $panelRight "Passwortaenderung beim Login" $false 150 $yRight
    $yRight += 20
    $chkCannotChangePW  = AddCheckBox $panelRight "Passwortaenderung verhindern" $false 150 $yRight
    $yRight += 20
    $chkSmartcardLogonRequired = AddCheckBox $panelRight "Smartcard-Anmeldung erforderlich" $false 150 $yRight
    $yRight += 40
    $lineAbove = Add-HorizontalLine -Parent $panelRight -y $yRight
    $yRight += 10

    #----------------------------------------------------
    # Abschnitt | PASSWORT-OPTIONEN
    # Zwei Zeilen: Erste Zeile Radio-Buttons, zweite Zeile festes Passwort.
    #----------------------------------------------------
    AddLabel $panelRight "PASSWORT?" 10 $yRight -Bold | Out-Null
    $yRight += 20
    $rbFix = New-Object System.Windows.Forms.RadioButton
    $rbFix.Text = "FEST"
    $rbFix.Location = New-Object System.Drawing.Point(150, $yRight)
    $panelRight.Controls.Add($rbFix)
    $rbRand = New-Object System.Windows.Forms.RadioButton
    $rbRand.Text = "GENERIERT"
    $rbRand.Location = New-Object System.Drawing.Point(260, $yRight)
    $panelRight.Controls.Add($rbRand)
    $yRight += 35
    AddLabel $panelRight "Festes Passwort:" 10 $yRight -Bold | Out-Null
    $txtFixPW = AddTextBox $panelRight "" 150 $yRight 150
    $yRight += 35
    $rbFix.Add_CheckedChanged({
        if ($rbFix.Checked) {
            $txtFixPW.Enabled = $true
        }
    })
    $rbRand.Add_CheckedChanged({
        if ($rbRand.Checked) {
            $txtFixPW.Enabled = $false
            $txtFixPW.Text = ""
        }
    })
    $rbRand.Checked = $true
    $txtFixPW.Enabled = $false
    AddLabel $panelRight "GENERIERT:" 10 $yRight -Bold | Out-Null
    $chkIncludeSpecial = AddCheckBox $panelRight "Sonderzeichen einbeziehen" $true 150 $yRight
    $chkAvoidAmbig     = AddCheckBox $panelRight "Aehnliche Zeichen vermeiden" $true 330 $yRight
    $yRight += 20
    AddLabel $panelRight "Passwortlaenge:" 10 $yRight -Bold | Out-Null
    $txtPWLen = AddTextBox $panelRight "12" 150 $yRight 50
    $yRight += 50
    $lineAbove = Add-HorizontalLine -Parent $panelRight -y $yRight
    $yRight += 10

    #----------------------------------------------------
    # Abschnitt | AD-Gruppen
    # Zeigt eine Auflistung der verfügbaren AD-Gruppen.
    #----------------------------------------------------
    AddLabel $panelRight "AD-Gruppen:" 10 $yRight -Bold | Out-Null; $yRight += 25
    $panelADGroups = New-Object System.Windows.Forms.Panel
    $panelADGroups.Location = New-Object System.Drawing.Point(10, $yRight)
    $panelADGroups.Size = New-Object System.Drawing.Size(480,150)
    $panelADGroups.AutoScroll = $true
    $panelADGroups.BorderStyle = 'None'
    $panelRight.Controls.Add($panelADGroups)
    $yRight += ($panelADGroups.Height + 10)
    $adGroupChecks = @{}
    if ($INIConfig.Contains("ADGroups")) {
        $adGroupKeys = $INIConfig.ADGroups.Keys | Where-Object { $_ -notmatch '^(DefaultADGroup|.*_(Visible|Label))$' }
        $groupCount = 0
        foreach ($g in $adGroupKeys) {
            $visibleKey = $g + "_Visible"
            $isVisible = $true
            if ($INIConfig.ADGroups.Contains($visibleKey) -and $INIConfig.ADGroups[$visibleKey] -eq '0') {
                $isVisible = $false
            }
            if ($isVisible) {
                $labelKey = $g + "_Label"
                $displayText = $g
                if ($INIConfig.ADGroups.Contains($labelKey) -and $INIConfig.ADGroups[$labelKey]) {
                    $displayText = $INIConfig.ADGroups[$labelKey]
                }
                $col = $groupCount % 3
                $row = [math]::Floor($groupCount / 3)
                $x = 10 + ($col * 150)
                $yPos = 10 + ($row * 30)
                $cbGroup = AddCheckBox $panelADGroups $displayText $false $x $yPos
                $adGroupChecks[$g] = $cbGroup
                $groupCount++
            }
        }
    }
    else {
        AddLabel $panelRight "Keine [ADGroups] Sektion gefunden." 10 $yRight -Bold | Out-Null
        $yRight += 25
    }

    #----------------------------------------------------
    # 5 | Unteres Panel: Buttons
    # Enthält Schaltflächen für Onboarding, PDF-Erstellung, Info und Schließen.
    # Platzierung der Buttons näher am Footer (weniger Abstand).
    #----------------------------------------------------
    $btnWidth = 175
    $btnHeight = 35
    $btnSpacing = 20
    $clientWidth = [int]$form.ClientSize.Width
    $totalButtonsWidth = (4 * $btnWidth) + (3 * $btnSpacing)
    $startX = [int](($clientWidth - $totalButtonsWidth) / 2)
    # Berechne Y-Position der Buttons: 5 Pixel Abstand vom unteren Rand des panelBottom
    $btnY = $panelBottom.Height - $btnHeight - 5

    $btnOnboard = New-Object System.Windows.Forms.Button
    $btnOnboard.Text = "ONBOARDING"
    $btnOnboard.Size = New-Object System.Drawing.Size($btnWidth, $btnHeight)
    $btnOnboard.Location = New-Object System.Drawing.Point($startX, $btnY)
    $btnOnboard.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnOnboard.BackColor = [System.Drawing.Color]::LightGreen
    $btnOnboard.Add_Click({
        # Sammle alle Eingaben in $result
        $result.Vorname         = $txtVorname.Text
        $result.Nachname        = $txtNachname.Text
        $result.Description     = $txtDescription.Text
        $result.OfficeRoom      = $txtOffice.Text
        $result.PhoneNumber     = $txtPhone.Text
        $result.MobileNumber    = $txtMobile.Text
        $result.Position        = $txtPosition.Text
        $result.DepartmentField = $txtDeptField.Text
        $result.Ablaufdatum     = $txtAblaufdatum.Text
        $result.Location        = $cmbLocation.SelectedItem
        $result.MS365License    = $cmbMS365License.SelectedItem
        $result.CompanySection  = $cmbCompany.SelectedItem
        $result.UPNFormat       = $cmbUPNFormat.SelectedItem
        $result.EmailAddress    = $txtEmail.Text.Trim()
        $result.MailSuffix      = $cmbMailSuffix.SelectedItem
        $result.OutputPDF       = $chkPDF_Left.Checked
        $result.OutputTXT       = $chkTXT_Left.Checked
        $result.UPNEntered      = $txtUPN.Text
        $result.PasswordNeverExpires  = $chkPWNeverExpires.Checked
        $result.MustChangePassword    = $chkMustChange.Checked
        $result.AccountDisabled       = $chkAccountDisabled.Checked
        $result.CannotChangePassword  = $chkCannotChangePW.Checked
        $result.SmartcardLogonRequired= $chkSmartcardLogonRequired.Checked
        if ($rbFix.Checked) {
            $result.PasswordMode = 0
            $result.FixPassword  = $txtFixPW.Text
        }
        else {
            $result.PasswordMode = 1
            $result.PasswordLaenge      = [int]$txtPWLen.Text
            $result.IncludeSpecialChars = $chkIncludeSpecial.Checked
            $result.AvoidAmbiguousChars = $chkAvoidAmbig.Checked
        }
        if ($chkExternal.Checked) {
            $result.DisplayName = "EXTERN | $($txtVorname.Text) $($txtNachname.Text)"
            $result.ADGroupsSelected = @()
            $result.Extern = $true
        }
        else {
            if (-not [string]::IsNullOrWhiteSpace($txtDisplayName.Text)) {
                $result.DisplayName = $txtDisplayName.Text
            }
            elseif ($cmbDisplayNameTemplate.SelectedItem -and $cmbDisplayNameTemplate.SelectedItem -match '{first}') {
                $template = $cmbDisplayNameTemplate.SelectedItem
                $result.DisplayName = $template -replace '{first}', $txtVorname.Text -replace '{last}', $txtNachname.Text
            }
            else {
                $result.DisplayName = "$($cmbCompany.SelectedItem.Display) | $($txtVorname.Text) $($txtNachname.Text)"
            }
            $groupSel = @()
            foreach ($key in $adGroupChecks.Keys) {
                if ($adGroupChecks[$key].Checked) { $groupSel += $key }
            }
            $result.ADGroupsSelected = $groupSel
            $result.Extern = $false
        }
        # Aufruf der AD-Erstellungs- und Logik-Funktion
        try {
            Process-Onboarding -userData $result -Config $INIConfig
            $lblStatus.Text = "Onboarding erfolgreich abgeschlossen."
            $progressBar.Value = 100
        }
        catch {
            $lblStatus.Text = "Fehler: $($_.Exception.Message)"
            [System.Windows.Forms.MessageBox]::Show("Fehler im Onboarding: $($_.Exception.Message)", "Fehler", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    })
    $panelBottom.Controls.Add($btnOnboard)
    
    # Button: PDF erstellen – ruft externes PDF-Skript auf
    $btnPDF = New-Object System.Windows.Forms.Button
    $btnPDF.Text = "ERSTELLE PDF"
    $btnPDF.Size = New-Object System.Drawing.Size($btnWidth, $btnHeight)
    $btnPDF.Location = New-Object System.Drawing.Point(($startX + $btnWidth + $btnSpacing), $btnY)
    $btnPDF.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnPDF.BackColor = [System.Drawing.Color]::LightYellow
    $btnPDF.Add_Click({
        if (-not $result.Vorname -or -not $result.Nachname) {
            [System.Windows.Forms.MessageBox]::Show("Bitte erst Onboarding durchführen, bevor das PDF erstellt werden kann.", "Fehler", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            return
        }
        $SamAccountName = ($result.Vorname.Substring(0,1) + $result.Nachname).ToLower()
        $htmlReportPath = Join-Path $INIConfig.General["ReportPath"] "$SamAccountName.html"
        $pdfReportPath  = Join-Path $INIConfig.General["ReportPath"] "$SamAccountName.pdf"
        $wkhtmltopdfPath = $reportBranding["wkhtmltopdfPath"]
        if (-not $wkhtmltopdfPath -or -not (Test-Path $wkhtmltopdfPath)) {
            $wkhtmltopdfPath = "C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe"
        }
        $pdfScript = Join-Path $PSScriptRoot "easyOnboarding_PDFCreator.ps1"
        Start-Process -FilePath "powershell.exe" -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$pdfScript`" -htmlFile `"$htmlReportPath`" -pdfFile `"$pdfReportPath`" -wkhtmltopdfPath `"$wkhtmltopdfPath`"" -NoNewWindow -Wait
    })
    $panelBottom.Controls.Add($btnPDF)
    
    # Button: Info
    $btnInfo = New-Object System.Windows.Forms.Button
    $btnInfo.Text = "INFO"
    $btnInfo.Size = New-Object System.Drawing.Size($btnWidth, $btnHeight)
    $btnInfo.Location = New-Object System.Drawing.Point(($startX + 2 * ($btnWidth + $btnSpacing)), $btnY)
    $btnInfo.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnInfo.BackColor = [System.Drawing.Color]::LightBlue
    $panelBottom.Controls.Add($btnInfo)
    $infoFilePath = ""
    if ($INIConfig.Contains("ScriptInfo") -and $INIConfig["ScriptInfo"].Contains("InfoFile")) {
        $infoFilePath = $INIConfig["ScriptInfo"]["InfoFile"]
    }
    $btnInfo.Add_Click({
        try {
            if ((-not [string]::IsNullOrWhiteSpace($infoFilePath)) -and (Test-Path $infoFilePath)) {
                Start-Process notepad.exe $infoFilePath
            }
            else {
                [System.Windows.Forms.MessageBox]::Show("Info-Datei nicht gefunden!", "Fehler", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            }
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show("Fehler beim Öffnen der Info-Datei: $($_.Exception.Message)", "Fehler", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    })

    # Button: Close – schließt das Formular
    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "CLOSE"
    $btnCancel.Size = New-Object System.Drawing.Size($btnWidth, $btnHeight)
    $btnCancel.Location = New-Object System.Drawing.Point(($startX + 3 * ($btnWidth + $btnSpacing)), $btnY)
    $btnCancel.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnCancel.BackColor = [System.Drawing.Color]::LightCoral
    $btnCancel.Add_Click({
        $result.Cancel = $true
        $form.Close()
    })
    $panelBottom.Controls.Add($btnCancel)

    [void]$form.ShowDialog()
    return $result
}

#----------------------------------------------------
# 5 | Hauptablauf
# Lädt die INI, zeigt das GUI an und verarbeitet die Eingaben.
#----------------------------------------------------
try {
    Write-Host "Lade INI: $ScriptINIPath"
    $Config = Read-INIFile $ScriptINIPath
}
catch {
    Throw "Fehler beim Laden der INI-Datei: $_"
}

if ($Config.General["DebugMode"] -eq "1") {
    Write-Host "DebugMode aktiviert."
}
$Language = $Config.General["Language"]

# GUI anzeigen und Eingaben erfassen
$userSelection = Show-OnboardingForm -INIConfig $Config
if ($userSelection.Cancel) {
    Write-Warning "Onboarding abgebrochen."
    return
}

Write-Host "`nOnboarding abgeschlossen. Das Fenster bleibt geöffnet, bis Sie auf 'Close' klicken."

# SIG # Begin signature block
# MIIbywYJKoZIhvcNAQcCoIIbvDCCG7gCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDrJ/JeBEL+d3SS
# Ve5iU/Pk2h7Mvnc26vgMg//wnV+/CqCCFhcwggMQMIIB+KADAgECAhB3jzsyX9Cg
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
# DQEJBDEiBCCRpWHfAmL/POvXs81sdcHCrEeN94R08lDhXirJ10r7fTANBgkqhkiG
# 9w0BAQEFAASCAQA6InYtY0q98dgBXnoriZ9aLKNJhpQXMlmo9wJkoPK2TOfRYeou
# O6+UrIL6q7RVuB46hym3kR8ymLvp+vZ5NrRbSGF4eiHi8oDU638bzdn0AscrujFJ
# BjARxLPOG5mrTT7DlhgQZsfjxYfBUtyElUfIWNU8kg2I8XVzV6bpRER86vRZktXy
# Hva/8aBp6oLE3Z7115yIHKte5Chp/YBBp7x1EhfhiM/Fs4yqcMSay0DDC8c6+7ON
# 7aDAUvdyPA0SHZCcAMdv40jW+GQXMl/jHzGGouIzSvp4kxNZgMlVT7NMdj8jTow+
# ogqJ74utUIONSralZ4aqww/kqd17l+/06xH7oYIDIDCCAxwGCSqGSIb3DQEJBjGC
# Aw0wggMJAgEBMHcwYzELMAkGA1UEBhMCVVMxFzAVBgNVBAoTDkRpZ2lDZXJ0LCBJ
# bmMuMTswOQYDVQQDEzJEaWdpQ2VydCBUcnVzdGVkIEc0IFJTQTQwOTYgU0hBMjU2
# IFRpbWVTdGFtcGluZyBDQQIQC65mvFq6f5WHxvnpBOMzBDANBglghkgBZQMEAgEF
# AKBpMBgGCSqGSIb3DQEJAzELBgkqhkiG9w0BBwEwHAYJKoZIhvcNAQkFMQ8XDTI1
# MDcwNTEwMTE1OVowLwYJKoZIhvcNAQkEMSIEIG8K7ZwF2tl8cBO/UY0a7+Sjm/NH
# 7hPNKxNouHcqURBhMA0GCSqGSIb3DQEBAQUABIICAJWR+E6fD4GZtFbpt6d4qiBR
# lkLQC9jeTfVYqCV2VbJtqY1qmeMCqGWGIftK+Pp91EMWDZ0WgxtDNYgq+T4uzP+5
# XG94AW96BM8fZrMH7SuroucFM+aWVH97VNogVJfc5kxcybRq9uOaQW/7RGOkDwG3
# 1elrcWXhjvovTq+XeJN4G1UTksz9pL7fC29tiJfcxdvKYJy7KLuxWJaPQVTuXl7Y
# 1jRYDXCUF0vPi6Fft11rpn35avoT3fLrPQvVBlHcnCwmkCMEnJew6R57na+pFXLe
# VMtiv8KVzgMQzsrRS+kvV9AGeM6g9GtSAxZVo51OeADjlKLej3QCN5imi6ZgxQYY
# vK0e+95P2eAE+ZN7lCK2dC97gS2NuQnLCopnDIZQdq2jHt2Y6fQ21i20vuqI4HI1
# JeWaZZwdg8cLtk25HroVs+CpzGr3IKkKxYDQSADd56UD6ID8hCmNd0j2tDWoti5h
# NQe2rJgQAz8wC0MYUUoUk8hGHq/tSbLxBVQu41qeNdEDGNmyWKo6PsNYBqO1n9wm
# IJ5FVB5D9f7q9B9F7Pjqn8xg//ThSOCLowEqG6XvuttLNxbe1FJ9bDxIzZoPORzY
# bZ23FNrbUhQX7ho4YpizKIn0VYp5xBfrEmRWZ0z2wlUyMLPSV6farkiFJNns6Sr9
# IJ1c+0Nb7+iWb6AJgKx8
# SIG # End signature block
