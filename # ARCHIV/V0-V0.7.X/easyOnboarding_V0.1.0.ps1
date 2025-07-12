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

# SIG # Begin signature block
# MIIbywYJKoZIhvcNAQcCoIIbvDCCG7gCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCASMz45vv+d5GbL
# xXCxk7WjlHwplpN8og2/Rr3fzqHBc6CCFhcwggMQMIIB+KADAgECAhB3jzsyX9Cg
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
# DQEJBDEiBCDMXiaBfS41PaPGgcX3sK3yJN0c4jMaeULbL+NtxLcUxDANBgkqhkiG
# 9w0BAQEFAASCAQA/lTeLz8DLCDsTi7w9CYyEOLqgBP7HqxYGWRVFt5+Y0SE3eSKu
# /DJftxzU/nvcU+90abU5NDGKEc7bdKu0hncelDLW3S7Ryg5QMvPzIL1fYMA1C2d+
# giYi69CcL0w6hk/RgyG7/+nnZ6d2ObBK5D52WDFK43PHSUwTE/7ntsnAC1COIhLO
# 0JMwpP0dlhA7Rkq1gwpJU5x2UOKvaA7lNHVA6AvDrhTxmIMtT7PMuw0JQ2mF3X19
# cIR4FvvXpCnkBB8wPhJOtkmIqVXdP4Fp2C4ablC4jI/OoEkzYHVUndfW18rZTckA
# 8rJvI6YqwzD0CrdSYfawShUjKwgbRoDneCXjoYIDIDCCAxwGCSqGSIb3DQEJBjGC
# Aw0wggMJAgEBMHcwYzELMAkGA1UEBhMCVVMxFzAVBgNVBAoTDkRpZ2lDZXJ0LCBJ
# bmMuMTswOQYDVQQDEzJEaWdpQ2VydCBUcnVzdGVkIEc0IFJTQTQwOTYgU0hBMjU2
# IFRpbWVTdGFtcGluZyBDQQIQC65mvFq6f5WHxvnpBOMzBDANBglghkgBZQMEAgEF
# AKBpMBgGCSqGSIb3DQEJAzELBgkqhkiG9w0BBwEwHAYJKoZIhvcNAQkFMQ8XDTI1
# MDcwNTEwMTE1NFowLwYJKoZIhvcNAQkEMSIEICpdTXvvZZw93EqdmDOLVNZ9c1Cl
# +jOKISSA6XIA2q6yMA0GCSqGSIb3DQEBAQUABIICACnRyvm9oElgTcTost1Gta/j
# 3mdXjykRCh1/6D6H/nIbujolm79vqIV+dGePEI922q1CKYr/wvv9NDclUVZLW+5G
# PExWpNiNXz/h5qp3Nrwem41wRKeUtr44eJkMqZ6bVnC1FTNuR0HyNw4BQjWQOaKt
# t4/DA+z1o3WAlGXNxnNyVQMnsl5GuuyvfhW6XQ0afUMlOGprbCrgknfjMX4oW9at
# gwk0C6WXZ3BdXsX+8nWkgNCpcNiNWi0ZyQV6R8qEjM9jBqGz0aEw2HKIl+OOVx/4
# 7PWGxWmAQEIifbz54eyU2bIzjgxVSncb6kEiiwGZ61RdD4fT3aBsOJbHIQRaGbAe
# JXAIK7AXnCIIFNkAuu6h7ryngrzl79o6n80QAMfwrpwI8o2vPy37ZLif+Fei6gxV
# 4MNHaPmCzJyK38PlXOMT61KbQH5sm5Voa4Stx9qNEsHfK+INJYcXM9B9IiVTChI3
# UTQDBzosaAMuB9TgBA2DOpjhTd7XZ4JKCjzexNemYxjADc2nNzvc0EPyyI4ZcfqI
# 1PiRMjCO0ZiglijkNdNbl1UJXgWTU4DzTH2D51OUjDO89njqMprF4bmlxOp8w1S8
# mWAM8Yq/yrIQqbSiLSvVwf1WqPhV8kfoZBA7/+Smp7UuKIKY0ddOg42NW4QqUQc/
# IDqaTQ9fawVtrQkhxK0w
# SIG # End signature block
