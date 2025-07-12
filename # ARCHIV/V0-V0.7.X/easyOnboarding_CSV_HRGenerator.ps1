# OnboardingDataCollector_Minimal_GUI_Ablaufdatum.ps1
# Dieses Skript oeffnet eine GUI zur Eingabe der Onboarding-Daten.
# Es speichert nur die Felder, die spaeter automatisch importiert werden sollen.
# Zusätzliche Felder:
# - Auswahl, ob der Nutzer Mitarbeiter (intern) oder Externer ist.
#   Bei intern: Auswahl der Company via DropDown (Company1 bis Company5).
#   Bei extern: Textfeld zur Angabe der Firma des Externen.
#
# Im unteren Bereich der GUI werden links die CSVGenerator-Einstellungen (CSVFolder und CSVFile)
# aus der INI angezeigt (Überschriften fett), rechts davon wird das Bild (HeaderLogo aus Branding-GUI)
# in der Größe 125x50 dargestellt.
#
# Eingabefelder:
# - Vorname, Nachname, Beschreibung, Buero, Telefonnummer, Mobilnummer, Position, Abteilung, Ablaufdatum

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

#############################################
# Funktion: INI-Datei einfach einlesen (ohne Kommentare)
#############################################
function Read-IniFile {
    param([string]$Path)
    $ini = @{}
    $currentSection = ""
    foreach ($line in Get-Content $Path) {
        $line = $line.Trim()
        if ($line -match "^\[(.+?)\]") {
            $currentSection = $matches[1]
            $ini[$currentSection] = @{}
        }
        elseif ($line -match "^(.*?)\s*=\s*(.*)$") {
            if ($currentSection) {
                $key = $matches[1].Trim()
                $value = $matches[2].Trim()
                $ini[$currentSection][$key] = $value
            }
        }
    }
    return $ini
}

#############################################
# INI-Datei finden (Muster: easyOnboarding*Config.ini)
#############################################
$iniFileCandidate = Get-ChildItem -Path . -Filter "easyOnboarding*Config.ini" | Select-Object -First 1
if ($null -eq $iniFileCandidate) {
    Throw "Keine INI-Datei gefunden, die dem Muster 'easyOnboarding*Config.ini' entspricht."
}
$iniFilePath = $iniFileCandidate.FullName

# INI einlesen
$iniData = Read-IniFile -Path $iniFilePath

#############################################
# CSVGenerator-Einstellungen aus der INI
#############################################
if ($iniData.Keys -contains "CSVGenerator") {
    $csvConfig = $iniData["CSVGenerator"]
} else {
    $csvConfig = @{ CSVFolder = "C:\temp\OnboardingData"; CSVFile = "OnboardingData.csv"; FontSize = "10"; FormBackColor = "#F0F0F0" }
}
$csvFolder    = $csvConfig.CSVFolder
$csvFileName  = $csvConfig.CSVFile
$fontSize     = $csvConfig.FontSize
$formBackColor = $csvConfig.FormBackColor
if (-not (Test-Path $csvFolder)) {
    New-Item -ItemType Directory -Path $csvFolder -Force | Out-Null
}
$csvFile = Join-Path $csvFolder $csvFileName

#############################################
# Branding-GUI: HeaderLogo aus der INI lesen
#############################################
if ($iniData.Keys -contains "Branding-GUI") {
    $brandingGUI = $iniData["Branding-GUI"]
    if ($brandingGUI.Keys -contains "HeaderLogo") {
        $headerLogo = $brandingGUI["HeaderLogo"]
    }
    else {
        $headerLogo = "APPICON.PNG"
    }
}
#############################################
# GUI-Erstellung: Funktionen für Label und TextBox
#############################################
function Add-Label {
    param(
        [System.Windows.Forms.Control]$parent,
        [string]$text,
        [int]$x,
        [int]$y
    )
    $label = New-Object System.Windows.Forms.Label
    $label.Text = $text
    $label.Location = New-Object System.Drawing.Point -ArgumentList $x, $y
    $label.AutoSize = $true
    $parent.Controls.Add($label)
    return $label
}

function Add-TextBox {
    param(
        [System.Windows.Forms.Control]$parent,
        [string]$default,
        [int]$x,
        [int]$y,
        [int]$width = 200
    )
    $tb = New-Object System.Windows.Forms.TextBox
    $tb.Text = $default
    $tb.Location = New-Object System.Drawing.Point -ArgumentList $x, $y
    $tb.Width = $width
    $parent.Controls.Add($tb)
    return $tb
}

#############################################
# Erstelle das Formular
#############################################
$form = New-Object System.Windows.Forms.Form
$form.Text = "Onboarding Daten erfassen"
$form.Size = New-Object System.Drawing.Size(500,700)
$form.StartPosition = "CenterScreen"
$form.BackColor = [System.Drawing.ColorTranslator]::FromHtml($formBackColor)
$form.Font = New-Object System.Drawing.Font($form.Font.FontFamily, [int]$fontSize)

$yPos = 20
$gap  = 30

# Eingabefelder
$txtVorname = Add-TextBox -parent $form -default "" -x 150 -y $yPos -width 315
Add-Label -parent $form -text "Vorname:" -x 20 -y $yPos
$yPos += $gap

$txtNachname = Add-TextBox -parent $form -default "" -x 150 -y $yPos -width 315
Add-Label -parent $form -text "Nachname:" -x 20 -y $yPos
$yPos += $gap

$yPos += 15

$txtDescription = Add-TextBox -parent $form -default "" -x 150 -y $yPos -width 315
Add-Label -parent $form -text "Beschreibung:" -x 20 -y $yPos
$yPos += $gap

$txtOffice = Add-TextBox -parent $form -default "" -x 150 -y $yPos -width 315
Add-Label -parent $form -text "Buero (OfficeRoom):" -x 20 -y $yPos
$yPos += $gap

$yPos += 15

$txtPhone = Add-TextBox -parent $form -default "" -x 150 -y $yPos -width 175
Add-Label -parent $form -text "Telefonnummer:" -x 20 -y $yPos
$yPos += $gap

$txtMobile = Add-TextBox -parent $form -default "" -x 150 -y $yPos -width 175
Add-Label -parent $form -text "Mobilnummer:" -x 20 -y $yPos
$yPos += $gap

$yPos += 15

$txtPosition = Add-TextBox -parent $form -default "" -x 150 -y $yPos -width 315
Add-Label -parent $form -text "Position:" -x 20 -y $yPos
$yPos += $gap

$txtDeptField = Add-TextBox -parent $form -default "" -x 150 -y $yPos -width 315
Add-Label -parent $form -text "Abteilung:" -x 20 -y $yPos
$yPos += $gap

$yPos += 15

$txtAblaufdatum = Add-TextBox -parent $form -default "" -x 150 -y $yPos -width 175
Add-Label -parent $form -text "Ablaufdatum:" -x 20 -y $yPos
$yPos += $gap

$yPos += 30

#############################################
# Neuer Abschnitt: Auswahl intern/extern
#############################################
$grpMitarbeiter = New-Object System.Windows.Forms.GroupBox
$grpMitarbeiter.Text = "Mitarbeitertyp"
$grpMitarbeiter.Location = New-Object System.Drawing.Point -ArgumentList 20, $yPos
$grpMitarbeiter.Size = New-Object System.Drawing.Size -ArgumentList 330, 60
$form.Controls.Add($grpMitarbeiter)

$rbIntern = New-Object System.Windows.Forms.RadioButton
$rbIntern.Text = "Mitarbeiter"
$rbIntern.Location = New-Object System.Drawing.Point -ArgumentList 10, 20
$rbIntern.AutoSize = $true
$rbIntern.Checked = $true
$grpMitarbeiter.Controls.Add($rbIntern)

$rbExtern = New-Object System.Windows.Forms.RadioButton
$rbExtern.Text = "Externer"
$rbExtern.Location = New-Object System.Drawing.Point -ArgumentList 120, 20
$rbExtern.AutoSize = $true
$grpMitarbeiter.Controls.Add($rbExtern)

$yPos += 70

# ComboBox für interne Company-Auswahl
$comboCompany = New-Object System.Windows.Forms.ComboBox
$comboCompany.Location = New-Object System.Drawing.Point -ArgumentList 150, $yPos
$comboCompany.Width = 200
$comboCompany.DropDownStyle = 'DropDownList'
$comboCompany.Items.AddRange(@("Company1","Company2","Company3","Company4","Company5"))
$comboCompany.SelectedIndex = 0
$form.Controls.Add($comboCompany)

$lblCompany = New-Object System.Windows.Forms.Label
$lblCompany.Text = "Firma (intern):"
$lblCompany.Location = New-Object System.Drawing.Point -ArgumentList 20, ($yPos+5)
$lblCompany.AutoSize = $true
$form.Controls.Add($lblCompany)

# TextBox für externe Firma
$txtExternFirma = Add-TextBox -parent $form -default "" -x 150 -y ($yPos + 40) -width 200
$lblExternFirma = New-Object System.Windows.Forms.Label
$lblExternFirma.Text = "Firma (extern):"
$lblExternFirma.Location = New-Object System.Drawing.Point -ArgumentList 20, ($yPos+40)
$lblExternFirma.AutoSize = $true
$form.Controls.Add($lblExternFirma)
$txtExternFirma.Enabled = $false
$lblExternFirma.Enabled = $false

# RadioButton-Events zur Steuerung
$rbIntern.Add_CheckedChanged({
    if ($rbIntern.Checked) {
        $comboCompany.Enabled = $true
        $lblCompany.Enabled = $true
        $txtExternFirma.Enabled = $false
        $lblExternFirma.Enabled = $false
    }
})
$rbExtern.Add_CheckedChanged({
    if ($rbExtern.Checked) {
        $comboCompany.Enabled = $false
        $lblCompany.Enabled = $false
        $txtExternFirma.Enabled = $true
        $lblExternFirma.Enabled = $true
    }
})

$yPos += 90

#############################################
# Neuer Informationsbereich: Links CSV-Einstellungen, Rechts Logo
#############################################
# Linke Seite: Überschriften (fett) und Werte für CSVFolder und CSVFile
$lblCSVFolderHeading = New-Object System.Windows.Forms.Label
$lblCSVFolderHeading.Text = "CSVFolder: "
$lblCSVFolderHeading.Location = New-Object System.Drawing.Point -ArgumentList 20, $yPos
$lblCSVFolderHeading.AutoSize = $true
$lblCSVFolderHeading.Font = New-Object System.Drawing.Font($lblCSVFolderHeading.Font, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($lblCSVFolderHeading)

$lblCSVFolderValue = New-Object System.Windows.Forms.Label
$lblCSVFolderValue.Text = $csvFolder
$lblCSVFolderValue.Location = New-Object System.Drawing.Point -ArgumentList ($lblCSVFolderHeading.Right + 5), $yPos
$lblCSVFolderValue.AutoSize = $true
$form.Controls.Add($lblCSVFolderValue)

$lblCSVFileHeading = New-Object System.Windows.Forms.Label
$lblCSVFileHeading.Text = "CSVFile: "
$lblCSVFileHeading.Location = New-Object System.Drawing.Point -ArgumentList 20, ($yPos + 20)
$lblCSVFileHeading.AutoSize = $true
$lblCSVFileHeading.Font = New-Object System.Drawing.Font($lblCSVFileHeading.Font, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($lblCSVFileHeading)

$lblCSVFileValue = New-Object System.Windows.Forms.Label
$lblCSVFileValue.Text = $csvFileName
$lblCSVFileValue.Location = New-Object System.Drawing.Point -ArgumentList ($lblCSVFileHeading.Right + 5), ($yPos + 20)
$lblCSVFileValue.AutoSize = $true
$form.Controls.Add($lblCSVFileValue)

# Rechte Seite: Logo anzeigen (verwende HeaderLogo aus Branding-GUI)
$picLogo = New-Object System.Windows.Forms.PictureBox
$picLogo.Location = New-Object System.Drawing.Point -ArgumentList 350, $yPos
$picLogo.Size = New-Object System.Drawing.Size -ArgumentList 125, 50
$picLogo.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::StretchImage
$picLogo.ImageLocation = $headerLogo
$form.Controls.Add($picLogo)

$yPos += 50

#############################################
# Buttons: "Speichern" (hellgrün) und "Schliessen" (hellrot)
#############################################
$btnY = $form.ClientSize.Height - 50

$btnSave = New-Object System.Windows.Forms.Button
$btnSave.Text = "Save CSV File"
$btnSave.Location = New-Object System.Drawing.Point -ArgumentList 80, $btnY
$btnSave.Size = New-Object System.Drawing.Size -ArgumentList 175,30
$btnSave.BackColor = [System.Drawing.Color]::LightGreen
$btnSave.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$form.Controls.Add($btnSave)

$btnClose = New-Object System.Windows.Forms.Button
$btnClose.Text = "Close"
$btnClose.Location = New-Object System.Drawing.Point -ArgumentList 290, $btnY
$btnClose.Size = New-Object System.Drawing.Size -ArgumentList 100,30
$btnClose.BackColor = [System.Drawing.Color]::LightCoral
$btnClose.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$form.Controls.Add($btnClose)

# Aktion: Beim Klick auf "Speichern" werden alle Daten gesammelt und in die CSV geschrieben;
# Fenster bleibt offen (kein $form.Close())
$btnSave.Add_Click({
    if ($rbIntern.Checked) {
        $company = $comboCompany.SelectedItem
    }
    elseif ($rbExtern.Checked) {
        $company = $txtExternFirma.Text.Trim()
    }
    else {
        $company = ""
    }
    
    $data = [PSCustomObject]@{
        Vorname         = $txtVorname.Text.Trim()
        Nachname        = $txtNachname.Text.Trim()
        Description     = $txtDescription.Text.Trim()
        OfficeRoom      = $txtOffice.Text.Trim()
        PhoneNumber     = $txtPhone.Text.Trim()
        MobileNumber    = $txtMobile.Text.Trim()
        Position        = $txtPosition.Text.Trim()
        DepartmentField = $txtDeptField.Text.Trim()
        Ablaufdatum     = $txtAblaufdatum.Text.Trim()
        Company         = $company
        LoginName       = $env:USERNAME
    }
    
    if (-not (Test-Path $csvFile)) {
        $data | Export-Csv -Path $csvFile -NoTypeInformation -Encoding UTF8
    }
    else {
        $data | Export-Csv -Path $csvFile -NoTypeInformation -Append -Encoding UTF8
    }
    [System.Windows.Forms.MessageBox]::Show("Daten wurden gespeichert in:`n$csvFile", "Erfolg", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    # Fenster bleibt offen
})

$btnClose.Add_Click({
    $form.Close()
})

[void] $form.ShowDialog()

# SIG # Begin signature block
# MIIcCAYJKoZIhvcNAQcCoIIb+TCCG/UCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCpIQ9+8oQHPfkg
# 3g8LTOQnZQrmZyP0k+qlFtFOYjI9eKCCFk4wggMQMIIB+KADAgECAhB3jzsyX9Cg
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
# LSd+2DrZ8LaHlv1b0VysGMNNn3O3AamfV6peKOK5lDCCBrQwggScoAMCAQICEA3H
# rFcF/yGZLkBDIgw6SYYwDQYJKoZIhvcNAQELBQAwYjELMAkGA1UEBhMCVVMxFTAT
# BgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNvbTEh
# MB8GA1UEAxMYRGlnaUNlcnQgVHJ1c3RlZCBSb290IEc0MB4XDTI1MDUwNzAwMDAw
# MFoXDTM4MDExNDIzNTk1OVowaTELMAkGA1UEBhMCVVMxFzAVBgNVBAoTDkRpZ2lD
# ZXJ0LCBJbmMuMUEwPwYDVQQDEzhEaWdpQ2VydCBUcnVzdGVkIEc0IFRpbWVTdGFt
# cGluZyBSU0E0MDk2IFNIQTI1NiAyMDI1IENBMTCCAiIwDQYJKoZIhvcNAQEBBQAD
# ggIPADCCAgoCggIBALR4MdMKmEFyvjxGwBysddujRmh0tFEXnU2tjQ2UtZmWgyxU
# 7UNqEY81FzJsQqr5G7A6c+Gh/qm8Xi4aPCOo2N8S9SLrC6Kbltqn7SWCWgzbNfiR
# +2fkHUiljNOqnIVD/gG3SYDEAd4dg2dDGpeZGKe+42DFUF0mR/vtLa4+gKPsYfwE
# u7EEbkC9+0F2w4QJLVSTEG8yAR2CQWIM1iI5PHg62IVwxKSpO0XaF9DPfNBKS7Za
# zch8NF5vp7eaZ2CVNxpqumzTCNSOxm+SAWSuIr21Qomb+zzQWKhxKTVVgtmUPAW3
# 5xUUFREmDrMxSNlr/NsJyUXzdtFUUt4aS4CEeIY8y9IaaGBpPNXKFifinT7zL2gd
# FpBP9qh8SdLnEut/GcalNeJQ55IuwnKCgs+nrpuQNfVmUB5KlCX3ZA4x5HHKS+rq
# BvKWxdCyQEEGcbLe1b8Aw4wJkhU1JrPsFfxW1gaou30yZ46t4Y9F20HHfIY4/6vH
# espYMQmUiote8ladjS/nJ0+k6MvqzfpzPDOy5y6gqztiT96Fv/9bH7mQyogxG9QE
# PHrPV6/7umw052AkyiLA6tQbZl1KhBtTasySkuJDpsZGKdlsjg4u70EwgWbVRSX1
# Wd4+zoFpp4Ra+MlKM2baoD6x0VR4RjSpWM8o5a6D8bpfm4CLKczsG7ZrIGNTAgMB
# AAGjggFdMIIBWTASBgNVHRMBAf8ECDAGAQH/AgEAMB0GA1UdDgQWBBTvb1NK6eQG
# fHrK4pBW9i/USezLTjAfBgNVHSMEGDAWgBTs1+OC0nFdZEzfLmc/57qYrhwPTzAO
# BgNVHQ8BAf8EBAMCAYYwEwYDVR0lBAwwCgYIKwYBBQUHAwgwdwYIKwYBBQUHAQEE
# azBpMCQGCCsGAQUFBzABhhhodHRwOi8vb2NzcC5kaWdpY2VydC5jb20wQQYIKwYB
# BQUHMAKGNWh0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0
# ZWRSb290RzQuY3J0MEMGA1UdHwQ8MDowOKA2oDSGMmh0dHA6Ly9jcmwzLmRpZ2lj
# ZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRSb290RzQuY3JsMCAGA1UdIAQZMBcwCAYG
# Z4EMAQQCMAsGCWCGSAGG/WwHATANBgkqhkiG9w0BAQsFAAOCAgEAF877FoAc/gc9
# EXZxML2+C8i1NKZ/zdCHxYgaMH9Pw5tcBnPw6O6FTGNpoV2V4wzSUGvI9NAzaoQk
# 97frPBtIj+ZLzdp+yXdhOP4hCFATuNT+ReOPK0mCefSG+tXqGpYZ3essBS3q8nL2
# UwM+NMvEuBd/2vmdYxDCvwzJv2sRUoKEfJ+nN57mQfQXwcAEGCvRR2qKtntujB71
# WPYAgwPyWLKu6RnaID/B0ba2H3LUiwDRAXx1Neq9ydOal95CHfmTnM4I+ZI2rVQf
# jXQA1WSjjf4J2a7jLzWGNqNX+DF0SQzHU0pTi4dBwp9nEC8EAqoxW6q17r0z0noD
# js6+BFo+z7bKSBwZXTRNivYuve3L2oiKNqetRHdqfMTCW/NmKLJ9M+MtucVGyOxi
# Df06VXxyKkOirv6o02OoXN4bFzK0vlNMsvhlqgF2puE6FndlENSmE+9JGYxOGLS/
# D284NHNboDGcmWXfwXRy4kbu4QFhOm0xJuF2EZAOk5eCkhSxZON3rGlHqhpB/8Ml
# uDezooIs8CVnrpHMiD2wL40mm53+/j7tFaxYKIqL0Q4ssd8xHZnIn/7GELH3IdvG
# 2XlM9q7WP/UwgOkw/HQtyRN62JK4S1C8uw3PdBunvAZapsiI5YKdvlarEvf8EA+8
# hcpSM9LHJmyrxaFtoza2zNaQ9k+5t1wwggbtMIIE1aADAgECAhAKgO8YS43xBYLR
# xHanlXRoMA0GCSqGSIb3DQEBCwUAMGkxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5E
# aWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1l
# U3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYgMjAyNSBDQTEwHhcNMjUwNjA0MDAwMDAw
# WhcNMzYwOTAzMjM1OTU5WjBjMQswCQYDVQQGEwJVUzEXMBUGA1UEChMORGlnaUNl
# cnQsIEluYy4xOzA5BgNVBAMTMkRpZ2lDZXJ0IFNIQTI1NiBSU0E0MDk2IFRpbWVz
# dGFtcCBSZXNwb25kZXIgMjAyNSAxMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIIC
# CgKCAgEA0EasLRLGntDqrmBWsytXum9R/4ZwCgHfyjfMGUIwYzKomd8U1nH7C8Dr
# 0cVMF3BsfAFI54um8+dnxk36+jx0Tb+k+87H9WPxNyFPJIDZHhAqlUPt281mHrBb
# ZHqRK71Em3/hCGC5KyyneqiZ7syvFXJ9A72wzHpkBaMUNg7MOLxI6E9RaUueHTQK
# WXymOtRwJXcrcTTPPT2V1D/+cFllESviH8YjoPFvZSjKs3SKO1QNUdFd2adw44wD
# cKgH+JRJE5Qg0NP3yiSyi5MxgU6cehGHr7zou1znOM8odbkqoK+lJ25LCHBSai25
# CFyD23DZgPfDrJJJK77epTwMP6eKA0kWa3osAe8fcpK40uhktzUd/Yk0xUvhDU6l
# vJukx7jphx40DQt82yepyekl4i0r8OEps/FNO4ahfvAk12hE5FVs9HVVWcO5J4dV
# mVzix4A77p3awLbr89A90/nWGjXMGn7FQhmSlIUDy9Z2hSgctaepZTd0ILIUbWuh
# KuAeNIeWrzHKYueMJtItnj2Q+aTyLLKLM0MheP/9w6CtjuuVHJOVoIJ/DtpJRE7C
# e7vMRHoRon4CWIvuiNN1Lk9Y+xZ66lazs2kKFSTnnkrT3pXWETTJkhd76CIDBbTR
# ofOsNyEhzZtCGmnQigpFHti58CSmvEyJcAlDVcKacJ+A9/z7eacCAwEAAaOCAZUw
# ggGRMAwGA1UdEwEB/wQCMAAwHQYDVR0OBBYEFOQ7/PIx7f391/ORcWMZUEPPYYzo
# MB8GA1UdIwQYMBaAFO9vU0rp5AZ8esrikFb2L9RJ7MtOMA4GA1UdDwEB/wQEAwIH
# gDAWBgNVHSUBAf8EDDAKBggrBgEFBQcDCDCBlQYIKwYBBQUHAQEEgYgwgYUwJAYI
# KwYBBQUHMAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBdBggrBgEFBQcwAoZR
# aHR0cDovL2NhY2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0VHJ1c3RlZEc0VGlt
# ZVN0YW1waW5nUlNBNDA5NlNIQTI1NjIwMjVDQTEuY3J0MF8GA1UdHwRYMFYwVKBS
# oFCGTmh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRHNFRp
# bWVTdGFtcGluZ1JTQTQwOTZTSEEyNTYyMDI1Q0ExLmNybDAgBgNVHSAEGTAXMAgG
# BmeBDAEEAjALBglghkgBhv1sBwEwDQYJKoZIhvcNAQELBQADggIBAGUqrfEcJwS5
# rmBB7NEIRJ5jQHIh+OT2Ik/bNYulCrVvhREafBYF0RkP2AGr181o2YWPoSHz9iZE
# N/FPsLSTwVQWo2H62yGBvg7ouCODwrx6ULj6hYKqdT8wv2UV+Kbz/3ImZlJ7YXwB
# D9R0oU62PtgxOao872bOySCILdBghQ/ZLcdC8cbUUO75ZSpbh1oipOhcUT8lD8QA
# GB9lctZTTOJM3pHfKBAEcxQFoHlt2s9sXoxFizTeHihsQyfFg5fxUFEp7W42fNBV
# N4ueLaceRf9Cq9ec1v5iQMWTFQa0xNqItH3CPFTG7aEQJmmrJTV3Qhtfparz+BW6
# 0OiMEgV5GWoBy4RVPRwqxv7Mk0Sy4QHs7v9y69NBqycz0BZwhB9WOfOu/CIJnzkQ
# TwtSSpGGhLdjnQ4eBpjtP+XB3pQCtv4E5UCSDag6+iX8MmB10nfldPF9SVD7weCC
# 3yXZi/uuhqdwkgVxuiMFzGVFwYbQsiGnoa9F5AaAyBjFBtXVLcKtapnMG3VH3EmA
# p/jsJ3FVF3+d1SVDTmjFjLbNFZUWMXuZyvgLfgyPehwJVxwC+UpX2MSey2ueIu9T
# HFVkT+um1vshETaWyQo8gmBto/m3acaP9QsuLj3FNwFlTxq25+T4QwX9xa6ILs84
# ZPvmpovq90K8eWyG2N01c4IhSOxqt81nMYIFEDCCBQwCAQEwNDAgMR4wHAYDVQQD
# DBVQaGluSVQtUFNzY3JpcHRzX1NpZ24CEHePOzJf0KCMSL6wELasExMwDQYJYIZI
# AWUDBAIBBQCggYQwGAYKKwYBBAGCNwIBDDEKMAigAoAAoQKAADAZBgkqhkiG9w0B
# CQMxDAYKKwYBBAGCNwIBBDAcBgorBgEEAYI3AgELMQ4wDAYKKwYBBAGCNwIBFTAv
# BgkqhkiG9w0BCQQxIgQgaTnqbE0j7OijGGdexQCEJWe1v/F0aNDpQeIaSf4Wooow
# DQYJKoZIhvcNAQEBBQAEggEADpFvDHWP5K3HX9wTRdVUv+UER8GgNLMHqFoSb7dw
# psbZL+Z7tRSBgC6Eq26sZTdEH0pSbJz8osHiH9kwUhbh4pnt8rkRwJctkez3Tcds
# /EsnWdS+vaBNw5J2yljIg4jFghAq24a3F4u6nCQ8vuVOnLTi5fHDfbqVmRJTdMKa
# 1GNf4215OXq0lxnvKjDFpKytfRfCkJ49y83HLcufk4iD72kKMR3IZIBiw22TyXLS
# sALJRofxD5EBfEXf9ILuO/W/QAIwfdTEWbTWAuDdTrD2073vUdua8z2mMlNced1U
# U2ECl2t2hq7EaiZKaCS01ngtFokJSNqx1myxx/KCduvRA6GCAyYwggMiBgkqhkiG
# 9w0BCQYxggMTMIIDDwIBATB9MGkxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdp
# Q2VydCwgSW5jLjFBMD8GA1UEAxM4RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3Rh
# bXBpbmcgUlNBNDA5NiBTSEEyNTYgMjAyNSBDQTECEAqA7xhLjfEFgtHEdqeVdGgw
# DQYJYIZIAWUDBAIBBQCgaTAYBgkqhkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqG
# SIb3DQEJBTEPFw0yNTA3MTIwODAwMzJaMC8GCSqGSIb3DQEJBDEiBCDqJLYW6xq6
# IjbSzKIMN73s7/zMRjFwnfGJlK56hn87rzANBgkqhkiG9w0BAQEFAASCAgBn8to/
# NVnfHVbrm12g6Qnov9sesh/mrbWatBR80mP33ZfJPmWZvZ52OPRQ+h37E6VmeXBr
# iwhmPheX0AtRluyDmQ5E1JjZeOunukuViVAoUlonrxEWBDJWeHmSeXzBftiqZYlm
# ddDqCSaaEm9pMCa0QefOdvzKXxP/NoKIXxNf8wFMA4OjsUa3j4fYHnhKPZa2pOQP
# j47Gj9qYCXfjtdgrYk+2xjbsAne4GfPdIFzJhOab6awsXaP/LtMQ5IQj6rt1D9Ov
# Tx3ak3ZY5HylKnsXr91K5JrFFV8MdRBrHbdFVOtPsyGfIJikTBP7Q4r/LydH31p/
# z7MGRZwx3k+D6UkwRtyJm1akOFQ+ZZvxr/K44YmiPSa/jaYSBk0emb32WOPLF/hQ
# coqn+/q24L4QxGBdxJ4kL4m7T0Ffo80Hx/6W6U9d6NOu2e/oJ+ciKjx659jTb5Iv
# 3B4w59PKlS0UjDLbcLAmUyO0NJghxnB1jkdkyxzguOtsnJeE/PTDExTT8AARji7i
# xiT0URWzQgnpxQFrTPSqZPSDk4yBhTlqtgpWDGCJHTHPpeJQa1bzjrkWux8KQAPe
# vksQxpv7sTP9muWawJNTL6df28owk4OwOlLaVSCPcMoQ2VcizYSbFJjOd2qpq+O4
# +4+Wr93WBqflCdq5twW/9A6AhBcIq3jIBjA9xg==
# SIG # End signature block
