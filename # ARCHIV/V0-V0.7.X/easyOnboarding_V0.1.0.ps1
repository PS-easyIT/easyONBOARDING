#requires -Version 5.1
[CmdletBinding()]
param(
    [string]$Vorname,
    [string]$Nachname,
    [string]$Abteilung,
    [string]$Standort,
    [string]$Company,
    [string]$License = "",
    [switch]$Extern,
    [string]$ScriptINIPath = "easyOnboarding_V0.1.0_Config.ini"
)

#########################################################################
# Funktion: INI-Datei einlesen und in ein Hash-Objekt umwandeln
#########################################################################
function Read-INIFile {
    param (
        [Parameter(Mandatory=$true)]
        [string]$Path
    )
    
    if (-not (Test-Path $Path)) {
        Throw "INI-Datei wurde nicht gefunden: $Path"
    }
    
    # Lese alle Zeilen und entferne Kommentare/Leerzeilen
    $iniContent = Get-Content -Path $Path | Where-Object { $_ -notmatch '^\s*[;#]' -and $_ -ne '' }
    
    # Erstelle ein leeres OrderedDictionary
    $iniHash = New-Object 'System.Collections.Specialized.OrderedDictionary'
    $section = $null
    
    foreach ($line in $iniContent) {
        if ($line -match '^\[(.+)\]$') {
            # Neue Sektion gefunden
            $section = $matches[1]
            if (-not $iniHash.Contains($section)) {
                # Füge eine neue Sektion als weiteres OrderedDictionary hinzu
                $iniHash.Add($section, (New-Object 'System.Collections.Specialized.OrderedDictionary'))
            }
        }
        elseif ($line -match '^(.*?)=(.*)$') {
            # Key=Value-Zeile gefunden
            $key = $matches[1].Trim()
            $value = $matches[2].Trim()
            if ($section -and $iniHash[$section] -ne $null) {
                $iniHash[$section].Add($key, $value)
            }
        }
    }
    return $iniHash
}

#########################################################################
# Funktion: GUI zur Abfrage fehlender Parameter via Windows Forms
# Mit Dropdown-Menüs für Standort, Firma und Lizenz
#########################################################################
function Show-UserInputForm {
    param(
        [Parameter(Mandatory)]
        [array]$StandortOptions,
        [Parameter(Mandatory)]
        [array]$CompanyOptions
    )
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "PHscripts.de | easyONBOARDING"
    $form.Size = New-Object System.Drawing.Size(400,380)
    $form.StartPosition = "CenterScreen"

    # Vorname
    $labelVorname = New-Object System.Windows.Forms.Label
    $labelVorname.Location = New-Object System.Drawing.Point(10,20)
    $labelVorname.Size = New-Object System.Drawing.Size(100,20)
    $labelVorname.Text = "FIRST NAME:"
    $form.Controls.Add($labelVorname)

    $txtVorname = New-Object System.Windows.Forms.TextBox
    $txtVorname.Location = New-Object System.Drawing.Point(120,20)
    $txtVorname.Size = New-Object System.Drawing.Size(250,20)
    if ($Vorname) { $txtVorname.Text = $Vorname }
    $form.Controls.Add($txtVorname)

    # Nachname
    $labelNachname = New-Object System.Windows.Forms.Label
    $labelNachname.Location = New-Object System.Drawing.Point(10,50)
    $labelNachname.Size = New-Object System.Drawing.Size(100,20)
    $labelNachname.Text = "LAST NAME:"
    $form.Controls.Add($labelNachname)

    $txtNachname = New-Object System.Windows.Forms.TextBox
    $txtNachname.Location = New-Object System.Drawing.Point(120,50)
    $txtNachname.Size = New-Object System.Drawing.Size(250,20)
    if ($Nachname) { $txtNachname.Text = $Nachname }
    $form.Controls.Add($txtNachname)

    # Abteilung
    $labelAbteilung = New-Object System.Windows.Forms.Label
    $labelAbteilung.Location = New-Object System.Drawing.Point(10,80)
    $labelAbteilung.Size = New-Object System.Drawing.Size(100,20)
    $labelAbteilung.Text = "DEPARTMENT:"
    $form.Controls.Add($labelAbteilung)

    $txtAbteilung = New-Object System.Windows.Forms.TextBox
    $txtAbteilung.Location = New-Object System.Drawing.Point(120,80)
    $txtAbteilung.Size = New-Object System.Drawing.Size(250,20)
    if ($Abteilung) { $txtAbteilung.Text = $Abteilung }
    $form.Controls.Add($txtAbteilung)

    # Standort (Dropdown)
    $labelStandort = New-Object System.Windows.Forms.Label
    $labelStandort.Location = New-Object System.Drawing.Point(10,110)
    $labelStandort.Size = New-Object System.Drawing.Size(100,20)
    $labelStandort.Text = "LOCATION:"
    $form.Controls.Add($labelStandort)

    $cmbStandort = New-Object System.Windows.Forms.ComboBox
    $cmbStandort.Location = New-Object System.Drawing.Point(120,110)
    $cmbStandort.Size = New-Object System.Drawing.Size(250,20)
    $cmbStandort.Items.AddRange($StandortOptions)
    if ($Standort) { $cmbStandort.SelectedItem = $Standort } else { $cmbStandort.SelectedIndex = 0 }
    $form.Controls.Add($cmbStandort)

    # Firma (Dropdown)
    $labelCompany = New-Object System.Windows.Forms.Label
    $labelCompany.Location = New-Object System.Drawing.Point(10,140)
    $labelCompany.Size = New-Object System.Drawing.Size(100,20)
    $labelCompany.Text = "COMPANY:"
    $form.Controls.Add($labelCompany)

    $cmbCompany = New-Object System.Windows.Forms.ComboBox
    $cmbCompany.Location = New-Object System.Drawing.Point(120,140)
    $cmbCompany.Size = New-Object System.Drawing.Size(250,20)
    $cmbCompany.Items.AddRange($CompanyOptions)
    if ($Company) { $cmbCompany.SelectedItem = $Company } else { $cmbCompany.SelectedIndex = 0 }
    $form.Controls.Add($cmbCompany)

    # Lizenz (Dropdown)
    $labelLicense = New-Object System.Windows.Forms.Label
    $labelLicense.Location = New-Object System.Drawing.Point(10,170)
    $labelLicense.Size = New-Object System.Drawing.Size(100,20)
    $labelLicense.Text = "MS365 LIZENZ:"
    $form.Controls.Add($labelLicense)

    $cmbLicense = New-Object System.Windows.Forms.ComboBox
    $cmbLicense.Location = New-Object System.Drawing.Point(120,170)
    $cmbLicense.Size = New-Object System.Drawing.Size(250,20)
    $cmbLicense.Items.AddRange(@("BUSINESS-STD", "BUSINESS-PREM", "E3", "E5"))
    if ($License -ne "") { $cmbLicense.SelectedItem = $License } else { $cmbLicense.SelectedIndex = 0 }
    $form.Controls.Add($cmbLicense)

    # Extern (Checkbox)
    $chkExtern = New-Object System.Windows.Forms.CheckBox
    $chkExtern.Location = New-Object System.Drawing.Point(120,200)
    $chkExtern.Size = New-Object System.Drawing.Size(250,20)
    $chkExtern.Text = "EXTERNAL?"
    $chkExtern.Checked = $Extern.IsPresent
    $form.Controls.Add($chkExtern)

    # OK-Button
    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Location = New-Object System.Drawing.Point(120,240)
    $okButton.Size = New-Object System.Drawing.Size(100,30)
    $okButton.Text = "ANLEGEN"
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.AcceptButton = $okButton
    $form.Controls.Add($okButton)

    # Abbrechen-Button
    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Location = New-Object System.Drawing.Point(230,240)
    $cancelButton.Size = New-Object System.Drawing.Size(100,30)
    $cancelButton.Text = "Abbrechen"
    $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.CancelButton = $cancelButton
    $form.Controls.Add($cancelButton)

    $dialogResult = $form.ShowDialog()
    if ($dialogResult -eq [System.Windows.Forms.DialogResult]::OK) {
        return @{
            Vorname   = $txtVorname.Text
            Nachname  = $txtNachname.Text
            Abteilung = $txtAbteilung.Text
            Standort  = $cmbStandort.SelectedItem
            Company   = $cmbCompany.SelectedItem
            License   = $cmbLicense.SelectedItem
            Extern    = $chkExtern.Checked
        }
    } else {
        Write-Host "Abbruch durch den Benutzer."
        exit
    }
}

#########################################################################
# Wenn nicht alle Pflichtparameter übergeben wurden, wird die INI geladen
# und das GUI angezeigt. Dabei werden die Dropdown-Optionen für Standort und
# Firma aus der INI ermittelt.
#########################################################################
if (-not $Vorname -or -not $Nachname -or -not $Abteilung -or -not $Standort -or -not $Company) {
    $ConfigForGUI = Read-INIFile -Path $ScriptINIPath
    # Standort-Optionen: alle Keys aus [STANDORTE] außer DefaultStandort
    $StandortOptions = $ConfigForGUI.STANDORTE.Keys | Where-Object { $_ -ne "DefaultStandort" }
    # Company-Optionen: alle Sektionen, die mit "DomainName" beginnen
    $CompanyOptions = $ConfigForGUI.Keys | Where-Object { $_ -like "DomainName*" }
    $userInput = Show-UserInputForm -StandortOptions $StandortOptions -CompanyOptions $CompanyOptions
    $Vorname   = $userInput.Vorname
    $Nachname  = $userInput.Nachname
    $Abteilung = $userInput.Abteilung
    $Standort  = $userInput.Standort
    $Company   = $userInput.Company
    $License   = $userInput.License
    if ($userInput.Extern) { $Extern = $true } else { $Extern = $false }
}

#########################################################################
# Ab hier folgt der restliche Code (INI-Parsing, AD-Benutzererstellung, Logging,
# HTML-/PDF-Report, etc.) – im Prinzip unverändert aus dem ursprünglichen Skript.
#########################################################################

Write-Host "Starte Onboarding für: $Vorname $Nachname"

######################################################################
# SamAccountName & UPN generieren
######################################################################
$SamAccountName = ($Vorname.Substring(0,1) + $Nachname).ToLower()
$UPN = ($SamAccountName + $Mailendung).ToLower()

######################################################################
# Passwort festlegen (fix oder generiert)
######################################################################
if ($passwordMode -eq '1') {
    $RandomPW = [System.Web.Security.Membership]::GeneratePassword($passwordLaenge,2)
    $UserPW   = $RandomPW
} else {
    $UserPW = $fixPassword
}
$SecurePW = ConvertTo-SecureString $UserPW -AsPlainText -Force

$Anzeigename = "$Vorname $Nachname"

######################################################################
# AD-Benutzer erstellen oder aktualisieren
######################################################################
try {
    $existingUser = Get-ADUser -Filter { SamAccountName -eq $SamAccountName } -ErrorAction SilentlyContinue
    if (-not $existingUser) {
        New-ADUser -Name $Anzeigename `
            -GivenName $Vorname `
            -Surname $Nachname `
            -SamAccountName $SamAccountName `
            -UserPrincipalName $UPN `
            -AccountPassword $SecurePW `
            -Enabled $true `
            -ChangePasswordAtLogon $mustChangePW `
            -PasswordNeverExpires $passwordNeverExpires `
            -Path $defaultOU `
            -StreetAddress $Strasse `
            -PostalCode $PLZ `
            -City $Ort `
            -State $State `
            -Country $Country `
            -Office $Office `
            -Title $Title -ErrorAction Stop
        Write-Host "AD-Benutzer '$Anzeigename' erstellt."
    } else {
        Write-Host "AD-Benutzer '$SamAccountName' existiert bereits. Aktualisiere relevante Felder..."
        Set-ADUser -Identity $existingUser.DistinguishedName `
            -StreetAddress $Strasse `
            -PostalCode $PLZ `
            -City $Ort `
            -State $State `
            -Country $Country `
            -Office $Office `
            -Title $Title -ErrorAction SilentlyContinue
    }
} catch {
    Write-Warning "Fehler beim Erstellen/Aktualisieren des AD-Benutzers: $_"
    return
}

######################################################################
# Gruppenmitgliedschaften (Abteilung, Signatur, Lizenz, ADSync)
######################################################################
try {
    if ($deptGroups) {
        $deptGroupsSplit = $deptGroups -split ';'
        foreach ($grp in $deptGroupsSplit) {
            if ($grp) {
                Add-ADGroupMember -Identity $grp -Members $SamAccountName -ErrorAction Stop
                Write-Host "Gruppe '$grp' zugewiesen."
            }
        }
    }
    if ($signaturGroup) {
        Add-ADGroupMember -Identity $signaturGroup -Members $SamAccountName -ErrorAction SilentlyContinue
        Write-Host "Signaturgruppe '$signaturGroup' zugewiesen."
    }
    if ($licenseGroup) {
        Add-ADGroupMember -Identity $licenseGroup -Members $SamAccountName -ErrorAction SilentlyContinue
        Write-Host "Lizenzgruppe '$licenseGroup' zugewiesen."
    }
    if ($adSyncEnabled -and $adSyncGroup) {
        Add-ADGroupMember -Identity $adSyncGroup -Members $SamAccountName -ErrorAction SilentlyContinue
        Write-Host "ADSync-Gruppe '$adSyncGroup' zugewiesen."
    }
} catch {
    Write-Warning "Fehler bei der Gruppen-Zuweisung: $_"
}

if ($Extern) {
    Write-Host "Externer Mitarbeiter: Hier kann eine Sonderbehandlung implementiert werden."
}

######################################################################
# Logging: Logfile mit den wichtigsten Informationen erstellen
######################################################################
try {
    if (-not (Test-Path $logFilePath)) {
        New-Item -ItemType Directory -Path $logFilePath -Force | Out-Null
    }
    $logDate  = (Get-Date -Format 'yyyyMMdd')
    $logFile  = Join-Path $logFilePath "Onboarding_$logDate.log"
    $logEntry = "[{0}] SamAccountName={1}, Anzeigename='{2}', UPN='{3}', Abteilung={4}, Lizenz={5}, Standort={6}, Firma={7}, Passwort='{8}', Extern={9}" -f (Get-Date), $SamAccountName, $Anzeigename, $UPN, $Abteilung, $License, $Standort, $Company, $UserPW, $Extern
    Add-Content -Path $logFile -Value $logEntry
    Write-Host "Logfile geschrieben: $logFile"
} catch {
    Write-Warning "Fehler beim Schreiben des Logfiles: $_"
}

######################################################################
# HTML-Report (und optional PDF) erstellen
######################################################################
try {
    if (-not (Test-Path $reportPath)) {
        New-Item -ItemType Directory -Path $reportPath -Force | Out-Null
    }
    # Firmenlogo einbinden
    $logoTag = ""
    if ($firmaLogoPath -and (Test-Path $firmaLogoPath)) {
        $logoTag = "<img src='file:///$firmaLogoPath' alt='Firmenlogo' style='float:right; max-width:120px; margin:10px;'/>"
    }
    # Links als Tabelle (2 Spalten) generieren
    $linksHtml = ""
    if ($employeeLinks.Count -gt 0) {
        $linksHtml += "<h3>Wichtige Links:</h3><table border='1' cellpadding='5'><tr><th>Name</th><th>URL</th></tr>"
        foreach ($link in $employeeLinks) {
            $parts = $link -split ';'
            $linkName = $parts[0]
            $linkURL  = $parts[1]
            $linksHtml += "<tr><td>$linkName</td><td><a href='$linkURL' target='_blank'>$linkURL</a></td></tr>"
        }
        $linksHtml += "</table>"
    }
    $htmlPath = Join-Path $reportPath "$SamAccountName.html"
    $htmlContent = @"
<html>
<head>
    <meta charset='UTF-8'>
    <title>$reportTitle</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; }
        .header { overflow: auto; }
        .footer { margin-top: 20px; font-size: 0.9em; color: #666; }
        .info { margin-bottom: 10px; }
    </style>
</head>
<body>
<div class="header">
    $logoTag
    <h1>$headerText</h1>
    <h2>$reportTitle</h2>
</div>
<div class="info">
    <p><b>Benutzer:</b> $Anzeigename</p>
    <p><b>Login (SamAccountName):</b> $SamAccountName</p>
    <p><b>UPN (E-Mail):</b> $UPN</p>
    <p><b>Abteilung:</b> $Abteilung</p>
    <p><b>Standort:</b> $Standort</p>
    <p><b>Lizenz:</b> $License</p>
    <p><b>Extern:</b> $Extern</p>
    <p><b>Erstelltes Passwort:</b> $UserPW</p>
</div>
$linksHtml
<div class="footer">
    <p>$reportFooter</p>
    <p>$footerText</p>
    <p>Erstellt am: $(Get-Date)</p>
</div>
</body>
</html>
"@
    Set-Content -Path $htmlPath -Value $htmlContent -Encoding UTF8
    Write-Host "HTML-Report erstellt: $htmlPath"

    # Optional: PDF-Erstellung (z.B. mit wkhtmltopdf) – Code hier auskommentiert:
    # $pdfPath = [System.IO.Path]::ChangeExtension($htmlPath, ".pdf")
    # & "C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe" $htmlPath $pdfPath
    # Write-Host "PDF-Report erstellt: $pdfPath"

} catch {
    Write-Warning "Fehler beim Erstellen des Reports: $_"
}

Write-Host "Onboarding abgeschlossen."
