#requires -Version 7.0
[CmdletBinding()]
param(
    [string]$Vorname,
    [string]$Nachname,
    [string]$Standort,
    [string]$Company,         # z. B. "1" für Company1, "2" für Company2 etc.
    [string]$License = "",
    [switch]$Extern,
    [string]$ScriptINIPath = "easyONBOARDING*Config.ini"
)

<#
  Dieses Skript wurde von PowerShell 5.x in PowerShell 7 konvertiert,
  auf Syntaxfehler geprüft und notwendige Elemente ergänzt.
  Es führt einen Check durch, ob das Betriebssystem Windows 10 oder Windows Server 2016 (bzw. höher) ist.
#>

#----------------------------------------------------
# OS-Kompatibilitätsprüfung: Nur Windows 10 / Server 2016+ zulassen
#----------------------------------------------------
try {
    $os = Get-CimInstance -ClassName Win32_OperatingSystem
    if (-not $os.Version.StartsWith("10.")) {
        Throw "Dieses Skript läuft nur unter Windows 10 oder Windows Server 2016 bzw. höher. Aktuelle OS-Version: $($os.Version)"
    }
}
catch {
    Write-Error "Fehler beim Überprüfen der OS-Version: $_"
    exit
}

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
        [int]$width = 250
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
# FUNKTION | Passwortgenerierung
#----------------------------------------------------
function Generate-AdvancedPassword {
    param(
        [int]$Laenge = 12,
        [bool]$IncludeSpecial = $true,
        [bool]$AvoidAmbiguous = $true,
        [int]$MinUpperCase = 2,
        [int]$MinDigits = 2,
        [int]$MinNonAlpha = 2
    )
    # Zeichensatzdefinition
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
        # Füge Mindestanzahl Großbuchstaben ein
        for ($i = 0; $i -lt $MinUpperCase; $i++) {
             $passwordChars += $upper[(Get-Random -Minimum 0 -Maximum $upper.Length)]
        }
        # Füge Mindestanzahl Ziffern ein
        for ($i = 0; $i -lt $MinDigits; $i++) {
             $passwordChars += $digits[(Get-Random -Minimum 0 -Maximum $digits.Length)]
        }
        # Fülle restliche Zeichen bis zur gewünschten Länge auf
        while ($passwordChars.Count -lt $Laenge) {
             $passwordChars += $all[(Get-Random -Minimum 0 -Maximum $all.Length)]
        }
        # Mische die Zeichen zufällig
        $passwordChars = $passwordChars | Sort-Object { Get-Random }
        $generatedPassword = -join $passwordChars
        # Zähle nicht-alphabetische Zeichen (Ziffern + Sonderzeichen)
        $nonAlphaCount = ($generatedPassword.ToCharArray() | Where-Object { $_ -notmatch '[A-Za-z]' }).Count
    } while ($nonAlphaCount -lt $MinNonAlpha)
    
    return $generatedPassword
}

#----------------------------------------------------
# FUNKTION | Onboarding verarbeiten
# Funktionen: AD-Erstellung, Logging und Report-Erstellung
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

        # INI-Werte explizit in Boolean umwandeln
        $includeSpecial = ($Config.PasswordFixGenerate["IncludeSpecialChars"] -match "^(?i:true|1)$")
        $avoidAmbiguous = ($Config.PasswordFixGenerate["AvoidAmbiguousChars"] -match "^(?i:true|1)$")
        
        $UserPW = Generate-AdvancedPassword -Laenge ([int]$Config.PasswordFixGenerate["PasswordLaenge"]) `
                                              -IncludeSpecial $includeSpecial `
                                              -AvoidAmbiguous $avoidAmbiguous `
                                              -MinUpperCase ([int]$Config.PasswordFixGenerate["MinUpperCase"]) `
                                              -MinDigits ([int]$Config.PasswordFixGenerate["MinDigits"]) `
                                              -MinNonAlpha ([int]$Config.PasswordFixGenerate["MinNonAlpha"])
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
    # AD-Benutzer anlegen/aktualisieren
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
    # AD-Gruppen zuweisen
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
    # Logging
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
    # Reports erzeugen
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
        # Prüfe, ob ein benutzerdefiniertes Logo gewählt wurde:
        if ($userData.CustomLogo -and (Test-Path $userData.CustomLogo)) {
            $logoTag = "<img src='$($userData.CustomLogo)' alt='Logo' />"
        }
        else {
            $logoTag = ""
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
            -replace "{{LogoTag}}", $logoTag
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
# FUNKTION | GUI-Erstellung: Onboarding-Form
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
        CustomLogo            = ""   # Neue Eigenschaft für benutzerdefiniertes Logo
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
    $form.Size = New-Object System.Drawing.Size(1085,1075)
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
    $panelLeft.Size = New-Object System.Drawing.Size(520,750)
    $panelLeft.AutoScroll = $true
    $panelLeft.BorderStyle = 'FixedSingle'
    $panelLeft.BackColor = [System.Drawing.Color]::WhiteSmoke
    $form.Controls.Add($panelLeft)

    #----------------------------------------------------
    # Rechtes Panel (weitere Einstellungen) erstellen
    #----------------------------------------------------
    $panelRight = New-Object System.Windows.Forms.Panel
    $panelRight.Location = New-Object System.Drawing.Point(540,80)
    $panelRight.Size = New-Object System.Drawing.Size(520,750)
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
    $cmbLocation = AddComboBox $panelLeft $locationDisplayList 150 $yLeft 250 ""; 
    $yLeft += 25

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
    $cmbMS365License = AddComboBox $panelLeft ( @("KEINE") + ($INIConfig.LicensesGroups.Keys | ForEach-Object { $_ -replace '^MS365_','' } ) ) 150 $yLeft 250 ""
    $yLeft += 40

    $lineAbove = Add-HorizontalLine -Parent $panelLeft -y $yLeft
    $yLeft += 20

    #----------------------------------------------------
    # Neuer Abschnitt: Logo Auswählen und Onboarding Dokument
    #----------------------------------------------------
    AddLabel $panelLeft "LOGO Auswählen ..." 10 $yLeft -Bold | Out-Null
    $btnBrowseLogo = New-Object System.Windows.Forms.Button
    $btnBrowseLogo.Text = "Durchsuchen"
    $btnBrowseLogo.Location = New-Object System.Drawing.Point(150, $yLeft)
    $btnBrowseLogo.Size = New-Object System.Drawing.Size(250, 30)
    $btnBrowseLogo.BackColor = [System.Drawing.Color]::DarkGray
    $btnBrowseLogo.ForeColor = [System.Drawing.Color]::White
    $btnBrowseLogo.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 8, [System.Drawing.FontStyle]::Bold)
    $panelLeft.Controls.Add($btnBrowseLogo)
    $btnBrowseLogo.Add_Click({
        $ofd = New-Object System.Windows.Forms.OpenFileDialog
        $ofd.Filter = "Bilddateien|*.jpg;*.jpeg;*.png;*.gif;*.bmp"
        if ($ofd.ShowDialog() -eq "OK") {
            $result.CustomLogo = $ofd.FileName
            $btnBrowseLogo.Text = "Logo gewählt"
            $btnBrowseLogo.BackColor = [System.Drawing.Color]::LightGreen  # Farbe wechselt zu Grün nach Auswahl
        }
    })
    $yLeft += 40
    $lineAbove = Add-HorizontalLine -Parent $panelLeft -y $yLeft
    $yLeft += 10

    AddLabel $panelLeft "ONBOARDING DOKUMENT" 10 $yLeft -Bold | Out-Null
    $yLeft += 25
    AddLabel $panelLeft "upn.HTML" 150 $yLeft -Bold | Out-Null
    $lblHTML = AddLabel $panelLeft "X" 250 $yLeft
    $lblHTML.ForeColor = [System.Drawing.Color]::Gray
    $yLeft += 20
    AddLabel $panelLeft "upn.TXT" 150 $yLeft -Bold | Out-Null
    $chkTXT_Left = AddCheckBox $panelLeft "" $true 250 $yLeft
    $yLeft += 10

    #----------------------------------------------------
    # Elemente im rechten Panel: Weitere Einstellungen
    #----------------------------------------------------
    $yRight = 10
    AddLabel $panelRight "Benutzer Name (UPN):" 10 $yRight -Bold | Out-Null
    $txtUPN = AddTextBox $panelRight "" 150 $yRight 250; $yRight += 25

    AddLabel $panelRight "UPN-Format-Vorlagen:" 10 $yRight -Bold | Out-Null
    $cmbUPNFormat = AddComboBox $panelRight @("VORNAME.NACHNAME","V.NACHNAME","VORNAMENACHNAME","VNACHNAME") 150 $yRight 250; $yRight += 35

    $lineAbove = Add-HorizontalLine -Parent $panelRight -y $yRight
    $yRight += 20

    AddLabel $panelRight "E-Mail-Adresse:" 10 $yRight -Bold | Out-Null
    $txtEmail = AddTextBox $panelRight "" 150 $yRight 250; $yRight += 35

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
    $yRight += 35

    $lineAbove = Add-HorizontalLine -Parent $panelRight -y $yRight
    $yRight += 10

    #----------------------------------------------------
    # Abschnitt | AD-Benutzer-Flags
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
    $yRight += 35
    $lineAbove = Add-HorizontalLine -Parent $panelRight -y $yRight
    $yRight += 10

    #----------------------------------------------------
    # Abschnitt | PASSWORT-OPTIONEN
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
    AddLabel $panelRight "RANDOM PASSWORD:" 10 $yRight -Bold | Out-Null
    $yRight += 25
    AddLabel $panelRight "Anzahl Zeichen:" 10 $yRight -Bold | Out-Null
    $txtPWLen = AddTextBox $panelRight "12" 150 $yRight 50
    $chkIncludeSpecial = AddCheckBox $panelRight "Sonderzeichen" $true 215 $yRight
    $chkAvoidAmbig     = AddCheckBox $panelRight "Aehnliche Zeichen vermeiden" $true 330 $yRight
    $yRight += 35
    $lineAbove = Add-HorizontalLine -Parent $panelRight -y $yRight
    $yRight += 10

    #----------------------------------------------------
    # Abschnitt | AD-Gruppen
    #----------------------------------------------------
    AddLabel $panelRight "AD-Gruppen:" 10 $yRight -Bold | Out-Null; $yRight += 25
    $panelADGroups = New-Object System.Windows.Forms.Panel
    $panelADGroups.Location = New-Object System.Drawing.Point(10, $yRight)
    $panelADGroups.Size = New-Object System.Drawing.Size(480,200)
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
    # Unteres Panel: Buttons
    #----------------------------------------------------
    $btnWidth = 175
    $btnHeight = 35
    $btnSpacing = 20
    $clientWidth = [int]$form.ClientSize.Width
    $totalButtonsWidth = (4 * $btnWidth) + (3 * $btnSpacing)
    $startX = [int](($clientWidth - $totalButtonsWidth) / 2)
    $btnY = $panelBottom.Height - $btnHeight - 5

    $btnOnboard = New-Object System.Windows.Forms.Button
    $btnOnboard.Text = "ONBOARDING"
    $btnOnboard.Size = New-Object System.Drawing.Size($btnWidth, $btnHeight)
    $btnOnboard.Location = New-Object System.Drawing.Point($startX, $btnY)
    $btnOnboard.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnOnboard.BackColor = [System.Drawing.Color]::LightGreen
    $btnOnboard.Add_Click({
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
# Hauptablauf: INI laden, GUI anzeigen und Onboarding verarbeiten
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

$userSelection = Show-OnboardingForm -INIConfig $Config
if ($userSelection.Cancel) {
    Write-Warning "Onboarding abgebrochen."
    return
}

Write-Host "Onboarding abgeschlossen. Das Fenster bleibt geöffnet, bis Sie auf 'Close' klicken."
