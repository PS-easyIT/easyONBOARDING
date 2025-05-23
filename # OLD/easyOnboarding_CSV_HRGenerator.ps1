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
