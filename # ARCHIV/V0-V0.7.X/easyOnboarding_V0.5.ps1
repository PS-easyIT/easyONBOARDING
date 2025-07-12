#requires -Version 5.1
[CmdletBinding()]
param(
    [string]$Vorname,
    [string]$Nachname,
    [string]$Standort,
    [string]$Company,         # Erwartet z. B. "1" für Company1, "2" für Company2 etc.
    [string]$License = "",
    [switch]$Extern,
    [string]$ScriptINIPath = "easyONBOARDING_V0.5_Config.ini"
)

# Für Passwort-Generierung
Add-Type -AssemblyName System.Web

###############################################################################
# 1) INI-Datei einlesen
###############################################################################
function Read-INIFile {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Path
    )
    if (-not (Test-Path $Path)) {
        Throw "INI-Datei nicht gefunden: $Path"
    }
    $iniContent = Get-Content -Path $Path | Where-Object { $_ -notmatch '^\s*[;#]' -and $_.Trim() -ne "" }
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

###############################################################################
# 2) Hilfsfunktionen (GUI-Elemente)
###############################################################################
function AddLabel {
    param(
        [Parameter(Mandatory=$true)] $parent,
        [AllowEmptyString()][string]$text,
        [int]$x,
        [int]$y,
        [switch]$Bold
    )
    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = $text
    $lbl.Location = New-Object System.Drawing.Point($x, $y)
    if ($Bold) {
        $lbl.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 8, [System.Drawing.FontStyle]::Bold)
    } else {
        $lbl.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 8)
    }
    $lbl.AutoSize = $true
    $parent.Controls.Add($lbl)
    return $lbl
}
function AddTextBox {
    param(
        [Parameter(Mandatory=$true)] $parent,
        [AllowEmptyString()][string]$default,
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
        [AllowEmptyString()][string]$text,
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
        [AllowEmptyString()][string]$default = ""
    )
    $cmb = New-Object System.Windows.Forms.ComboBox
    $cmb.DropDownStyle = 'DropDownList'
    $cmb.Location = New-Object System.Drawing.Point($x, $y)
    $cmb.Width = $width
    foreach ($i in $items) { [void]$cmb.Items.Add($i) }
    if ($default -ne "" -and $cmb.Items.Contains($default)) {
        $cmb.SelectedItem = $default
    } elseif ($cmb.Items.Count -gt 0) {
        $cmb.SelectedIndex = 0
    }
    $parent.Controls.Add($cmb)
    return $cmb
}

###############################################################################
# 3) GUI-Erstellung
###############################################################################
function Show-OnboardingForm {
    param(
        [hashtable]$INIConfig
    )
    $currentUser = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name

    $guiBranding = if ($INIConfig.Contains("Branding-GUI")) { $INIConfig["Branding-GUI"] } else { @{} }
    $reportBranding = if ($INIConfig.Contains("Branding-Report")) { $INIConfig["Branding-Report"] } else { @{} }

    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    $form = New-Object System.Windows.Forms.Form
    $form.Text = if ($guiBranding.APPName) { $guiBranding.APPName } else { "easyONBOARDING" }
    $form.StartPosition = "CenterScreen"
    $form.Size = New-Object System.Drawing.Size(1085,900)
    $form.AutoScroll = $true

    if ($guiBranding.BackgroundImage -and (Test-Path $guiBranding.BackgroundImage)) {
        $form.BackgroundImage = [System.Drawing.Image]::FromFile($guiBranding.BackgroundImage)
        $form.BackgroundImageLayout = 'Stretch'
    }

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

    $panelLeft = New-Object System.Windows.Forms.Panel
    $panelLeft.Location = New-Object System.Drawing.Point(10,80)
    $panelLeft.Size = New-Object System.Drawing.Size(520,650)
    $panelLeft.AutoScroll = $true
    $panelLeft.BorderStyle = 'FixedSingle'
    $form.Controls.Add($panelLeft)

    $panelRight = New-Object System.Windows.Forms.Panel
    $panelRight.Location = New-Object System.Drawing.Point(540,80)
    $panelRight.Size = New-Object System.Drawing.Size(520,650)
    $panelRight.AutoScroll = $false
    $panelRight.BorderStyle = 'FixedSingle'
    $form.Controls.Add($panelRight)

    $panelBottom = New-Object System.Windows.Forms.Panel
    $panelBottom.Dock = 'Bottom'
    $panelBottom.Height = 60
    $panelBottom.BorderStyle = 'None'
    $form.Controls.Add($panelBottom)

    $panelFooter = New-Object System.Windows.Forms.Panel
    $panelFooter.Dock = 'Bottom'
    $panelFooter.Height = 40
    $panelFooter.BorderStyle = 'FixedSingle'
    $lblFooter = AddLabel $panelFooter " " 10 10
    $lblFooter.Text = if ($guiBranding.FooterWebseite) { $guiBranding.FooterWebseite } else { "www.easyONBOARDING.com" }
    $form.Controls.Add($panelFooter)

    ############################################################################
    # Elemente im PanelLeft (Eingabe)
    ############################################################################
    $yLeft = 10
    AddLabel $panelLeft "Vorname:" 10 $yLeft -Bold | Out-Null
    $txtVorname = AddTextBox $panelLeft "" 150 $yLeft; $yLeft += 30

    AddLabel $panelLeft "Nachname:" 10 $yLeft -Bold | Out-Null
    $txtNachname = AddTextBox $panelLeft "" 150 $yLeft; $yLeft += 40

    AddLabel $panelLeft "Externer Mitarbeiter:" 10 $yLeft -Bold | Out-Null
    $chkExternal = AddCheckBox $panelLeft "" $false 150 $yLeft; $yLeft += 40

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
    $txtOffice = AddTextBox $panelLeft "" 150 $yLeft; $yLeft += 30

    AddLabel $panelLeft "Rufnummer:" 10 $yLeft -Bold | Out-Null
    $txtPhone = AddTextBox $panelLeft "" 150 $yLeft; $yLeft += 30

    AddLabel $panelLeft "Mobil:" 10 $yLeft -Bold | Out-Null
    $txtMobile = AddTextBox $panelLeft "" 150 $yLeft; $yLeft += 30

    AddLabel $panelLeft "Position:" 10 $yLeft -Bold | Out-Null
    $txtPosition = AddTextBox $panelLeft "" 150 $yLeft; $yLeft += 30

    AddLabel $panelLeft "Abteilung (manuell):" 10 $yLeft -Bold | Out-Null
    $txtDeptField = AddTextBox $panelLeft "" 150 $yLeft; $yLeft += 55

    AddLabel $panelLeft "Location*:" 10 $yLeft -Bold | Out-Null
    $locationDisplayList = $INIConfig.STANDORTE.Keys | Where-Object { $_ -match '_Bez$' } | ForEach-Object { $INIConfig.STANDORTE[$_] }
    $cmbLocation = AddComboBox $panelLeft $locationDisplayList 150 $yLeft 250 ""; $yLeft += 30

    # --- DropDown für Company: nur anzeigen, wenn der Visible-Schalter auf "1" steht (oder nicht definiert ist) ---
    AddLabel $panelLeft "Firma:" 10 $yLeft -Bold | Out-Null
    $companyOptions = @()
    foreach ($section in $INIConfig.Keys | Where-Object { $_ -like "Company*" }) {
        $suffix = ($section -replace "\D", "")
        $visibleKey = "$section`_Visible"
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

    AddLabel $panelLeft "MS365 Lizenz*:" 10 $yLeft -Bold | Out-Null
    $cmbMS365License = AddComboBox $panelLeft ( @("KEINE") + ($INIConfig.LicensesGroups.Keys | ForEach-Object { $_ -replace '^MS365_','' } ) ) 150 $yLeft 200 ""; $yLeft += 55

    AddLabel $panelLeft "ONBOARDING DOKUMENT ERZEUGEN?" 10 $yLeft -Bold | Out-Null
    $yLeft += 20
    $chkHTML_Left = AddCheckBox $panelLeft "HTML erzeugen" $true 10 $yLeft
    $chkPDF_Left  = AddCheckBox $panelLeft "PDF erzeugen" $true 150 $yLeft
    $chkTXT_Left  = AddCheckBox $panelLeft "TXT erzeugen" $true 290 $yLeft; $yLeft += 35

    ############################################################################
    # Elemente im PanelRight (E-Mail, UPN, etc.)
    ############################################################################
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
    $yRight += 55

    AddLabel $panelRight "AD-Benutzer-Flags:" 10 $yRight -Bold | Out-Null; $yRight += 20
    $chkPWNeverExpires = AddCheckBox $panelRight "PasswordNeverExpires" $false 10 $yRight
    $chkMustChange     = AddCheckBox $panelRight "MustChangePasswordAtLogon" $false 150 $yRight
    $chkAccountDisabled = AddCheckBox $panelRight "AccountDisabled" $false 10 ($yRight + 35)
    $chkCannotChangePW  = AddCheckBox $panelRight "CannotChangePassword" $false 150 ($yRight + 35)
    $chkSmartcardLogonRequired = AddCheckBox $panelRight "SmartcardLogonRequired" $false 300 ($yRight + 35); $yRight += 85

    AddLabel $panelRight "PASSWORT-OPTIONEN:" 10 $yRight -Bold | Out-Null; $yRight += 25
    $rbFix = New-Object System.Windows.Forms.RadioButton
    $rbFix.Text = "FEST"
    $rbFix.Location = New-Object System.Drawing.Point(10, $yRight)
    $panelRight.Controls.Add($rbFix)
    $rbRand = New-Object System.Windows.Forms.RadioButton
    $rbRand.Text = "GENERIERT"
    $rbRand.Location = New-Object System.Drawing.Point(150, $yRight)
    $panelRight.Controls.Add($rbRand); $yRight += 35
    AddLabel $panelRight "Festes Passwort:" 10 $yRight -Bold | Out-Null
    $txtFixPW = AddTextBox $panelRight "" 150 $yRight 150; $yRight += 35
    AddLabel $panelRight "Passwortlaenge:" 10 $yRight -Bold | Out-Null
    $txtPWLen = AddTextBox $panelRight "12" 150 $yRight 50; $yRight += 35
    $chkIncludeSpecial = AddCheckBox $panelRight "IncludeSpecialChars" $true 10 $yRight
    $chkAvoidAmbig     = AddCheckBox $panelRight "AvoidAmbiguousChars" $true 150 $yRight; $yRight += 50

    AddLabel $panelRight "AD-Gruppen:" 10 $yRight -Bold | Out-Null; $yRight += 25
    $panelADGroups = New-Object System.Windows.Forms.Panel
    $panelADGroups.Location = New-Object System.Drawing.Point(10, $yRight)
    $panelADGroups.Size = New-Object System.Drawing.Size(480,150)
    $panelADGroups.AutoScroll = $true
    $panelRight.Controls.Add($panelADGroups); $yRight += $panelADGroups.Height + 10
    $adGroupChecks = @{}
    if ($INIConfig.Contains("ADGroups")) {
        $adGroupKeys = $INIConfig.ADGroups.Keys | Where-Object { $_ -notmatch '^(DefaultADGroup|.*_(Visible|Label))$' }
        $groupCount = 0
        foreach ($g in $adGroupKeys) {
            $visibleKey = $g + "_Visible"
            $isVisible = $true
            if ($INIConfig.ADGroups.Contains($visibleKey) -and $INIConfig.ADGroups[$visibleKey] -eq '0') { $isVisible = $false }
            if ($isVisible) {
                $labelKey = $g + "_Label"
                $displayText = $g
                if ($INIConfig.ADGroups.Contains($labelKey) -and $INIConfig.ADGroups[$labelKey]) {
                    $displayText = $INIConfig.ADGroups[$labelKey]
                }
                $col = $groupCount % 3
                $row = [math]::Floor($groupCount / 3)
                $x = 10 + ($col * 170)
                $yPos = 10 + ($row * 30)
                $cbGroup = AddCheckBox $panelADGroups $displayText $false $x $yPos
                $adGroupChecks[$g] = $cbGroup
                $groupCount++
            }
        }
    }
    else {
        AddLabel $panelRight "Keine [ADGroups] Sektion gefunden." 10 $yRight -Bold | Out-Null; $yRight += 25
    }

    ############################################################################
    # PanelBottom: Buttons
    ############################################################################
    $btnWidth = 175
    $btnHeight = 30
    $btnSpacing = 20
    $clientWidth = [int]$form.ClientSize.Width
    $totalButtonsWidth = (3 * $btnWidth) + (2 * $btnSpacing)
    $startX = [int](($clientWidth - $totalButtonsWidth) / 2)
    
    $btnOnboard = New-Object System.Windows.Forms.Button
    $btnOnboard.Text = "New-ADUser"
    $btnOnboard.Size = New-Object System.Drawing.Size($btnWidth, $btnHeight)
    $btnOnboard.Location = New-Object System.Drawing.Point($startX, 15)
    $btnOnboard.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnOnboard.BackColor = [System.Drawing.Color]::LightGreen
    $panelBottom.Controls.Add($btnOnboard)
    
    $btnInfo = New-Object System.Windows.Forms.Button
    $btnInfo.Text = "Info"
    $btnInfo.Size = New-Object System.Drawing.Size($btnWidth, $btnHeight)
    $btnInfo.Location = New-Object System.Drawing.Point([int]($startX + $btnWidth + $btnSpacing), 15)
    $btnInfo.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnInfo.BackColor = [System.Drawing.Color]::LightBlue
    $panelBottom.Controls.Add($btnInfo)
    
    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "Close"
    $btnCancel.Size = New-Object System.Drawing.Size($btnWidth, $btnHeight)
    $btnCancel.Location = New-Object System.Drawing.Point([int]($startX + 2 * ($btnWidth + $btnSpacing)), 15)
    $btnCancel.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnCancel.BackColor = [System.Drawing.Color]::LightCoral
    $panelBottom.Controls.Add($btnCancel)
    
    $infoFilePath = ""
    if ($INIConfig.Contains("ScriptInfo") -and $INIConfig.ScriptInfo.Contains("InfoFile")) {
        $infoFilePath = $INIConfig.ScriptInfo["InfoFile"]
    }
    $btnInfo.Add_Click({
        if ((-not [string]::IsNullOrWhiteSpace($infoFilePath)) -and (Test-Path $infoFilePath)) {
            Start-Process notepad.exe $infoFilePath
        }
        else {
            [System.Windows.Forms.MessageBox]::Show("Info-Datei nicht gefunden!", "Fehler", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    })

    ############################################################################
    # Ergebnis-Objekt (GUI-Ergebnis)
    ############################################################################
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
        CompanySection        = $null      # Speichert das komplette Company-Objekt
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
        OutputHTML            = $chkHTML_Left.Checked
        OutputPDF             = $chkPDF_Left.Checked
        OutputTXT             = $chkTXT_Left.Checked
        UPNEntered            = ""
        UPNFormat             = ""    # Wird nun aus der ComboBox übernommen
        EmailAddress          = ""
        MailSuffix            = ""    # Wird später gesetzt
        Cancel                = $false
        ADGroupsSelected      = @()
        Extern                = $false
        SmartcardLogonRequired= $false
    }

    function UpdatePWFields {
        if ($rbFix.Checked) {
            $txtFixPW.Enabled = $true
            $txtPWLen.Enabled = $false
            $chkIncludeSpecial.Enabled = $false
            $chkAvoidAmbig.Enabled = $false
        } else {
            $txtFixPW.Enabled = $false
            $txtPWLen.Enabled = $true
            $chkIncludeSpecial.Enabled = $true
            $chkAvoidAmbig.Enabled = $true
        }
    }
    $rbFix.Add_CheckedChanged({ UpdatePWFields })
    $rbRand.Add_CheckedChanged({ UpdatePWFields })
    $rbRand.Checked = $true
    UpdatePWFields

    $btnOnboard.Add_Click({
        # UPN-Format übernehmen
        $result.UPNFormat = $cmbUPNFormat.SelectedItem
        # Prüfe, ob eine Company ausgewählt wurde
        if (-not $cmbCompany.SelectedItem) {
            Throw "Fehler: Es wurde keine Company ausgewählt. Bitte wählen Sie einen Bereich aus."
        }
        # Werte aus der GUI übernehmen
        $result.Vorname         = $txtVorname.Text
        $result.Nachname        = $txtNachname.Text
        $result.Description     = $txtDescription.Text
        $result.OfficeRoom      = $txtOffice.Text
        $result.PhoneNumber     = $txtPhone.Text
        $result.MobileNumber    = $txtMobile.Text
        $result.Position        = $txtPosition.Text
        $result.DepartmentField = $txtDeptField.Text
        $result.Location        = $cmbLocation.SelectedItem
        $result.MS365License    = $cmbMS365License.SelectedItem
        $result.CompanySection  = $cmbCompany.SelectedItem
        $selectedCompany = $cmbCompany.SelectedItem
        $suffix = ($selectedCompany.Section -replace "\D", "")
        if (($INIConfig[$selectedCompany.Section].Keys -contains ("NameFirma" + $suffix)) -and -not [string]::IsNullOrWhiteSpace($INIConfig[$selectedCompany.Section]["NameFirma" + $suffix])) {
            $prefix = $INIConfig[$selectedCompany.Section]["NameFirma" + $suffix].Trim()
        } else {
            $prefix = $selectedCompany.Display
        }
        if (-not [string]::IsNullOrWhiteSpace($prefix)) {
            $DisplayName = "$prefix | $($txtVorname.Text) $($txtNachname.Text)"
        } else {
            $DisplayName = "$($txtVorname.Text) $($txtNachname.Text)"
        }
        Write-Host "DisplayName wird gesetzt: $DisplayName"
        $result.DisplayName = $DisplayName

        $emailInput = $txtEmail.Text.Trim()
        if (($INIConfig.Keys -contains $selectedCompany.Section) -and 
            ($INIConfig[$selectedCompany.Section].Keys -contains ("MailDomain" + $suffix)) -and 
            -not [string]::IsNullOrWhiteSpace($INIConfig[$selectedCompany.Section]["MailDomain" + $suffix])) {
            $companyMailDomain = $INIConfig[$selectedCompany.Section]["MailDomain" + $suffix].Trim()
        } else {
            $companyMailDomain = ""
        }
        if ($emailInput -ne "") {
            if ($emailInput -notmatch "@") {
                if (-not [string]::IsNullOrWhiteSpace($companyMailDomain)) {
                    $emailInput = "$emailInput$companyMailDomain"
                } elseif (-not [string]::IsNullOrWhiteSpace($cmbMailSuffix.SelectedItem)) {
                    $emailInput = "$emailInput$cmbMailSuffix.SelectedItem"
                }
            }
        }
        $result.EmailAddress = $emailInput
        $result.MailSuffix = $cmbMailSuffix.SelectedItem

        if ($chkExternal.Checked) {
            if ([string]::IsNullOrWhiteSpace($txtDisplayName.Text)) {
                $result.DisplayName = "EXTERN | $($txtVorname.Text) $($txtNachname.Text)"
            }
            $result.ADGroupsSelected = @()
        } else {
            $groupSel = @()
            foreach ($key in $adGroupChecks.Keys) {
                if ($adGroupChecks[$key].Checked) { $groupSel += $key }
            }
            $result.ADGroupsSelected = $groupSel
        }
        $form.Close()
    })

    $btnCancel.Add_Click({
        $result.Cancel = $true
        $form.Close()
    })

    $null = $form.ShowDialog()
    return $result
}

###############################################################################
# 4) Hauptablauf: INI laden, GUI anzeigen, Ergebnisse verarbeiten
###############################################################################
Write-Host "Lade INI: $ScriptINIPath"
$Config = Read-INIFile $ScriptINIPath

if ($Config.General.DebugMode -eq "1") {
    Write-Host "DebugMode aktiviert."
}
$Language = $Config.General.Language

$userSelection = Show-OnboardingForm -INIConfig $Config

if ($userSelection.Cancel) {
    Write-Warning "Onboarding abgebrochen."
    return
}

$Vorname             = $userSelection.Vorname
$Nachname            = $userSelection.Nachname
$DisplayName         = $userSelection.DisplayName
$Description         = $userSelection.Description
$OfficeRoom          = $userSelection.OfficeRoom
$PhoneNumber         = $userSelection.PhoneNumber
$MobileNumber        = $userSelection.MobileNumber
$Position            = $userSelection.Position
$DepartmentField     = $userSelection.DepartmentField
$Location            = $userSelection.Location
$MS365License        = $userSelection.MS365License
$passwordNeverExpires = $userSelection.PasswordNeverExpires
$mustChangePW         = $userSelection.MustChangePassword
$accountDisabled      = $userSelection.AccountDisabled
$cannotChangePW       = $userSelection.CannotChangePassword
$passwordMode         = $userSelection.PasswordMode
$fixPassword          = $userSelection.FixPassword
$passwordLaenge       = $userSelection.PasswordLaenge
$includeSpecial       = $userSelection.IncludeSpecialChars
$avoidAmbiguous       = $userSelection.AvoidAmbiguousChars
$createHTML           = $userSelection.OutputHTML
$createPDF            = $userSelection.OutputPDF
$createTXT            = $userSelection.OutputTXT
$UPNManual            = $userSelection.UPNEntered
$UPNTemplate          = $userSelection.UPNFormat
$EmailAddress         = $userSelection.EmailAddress
$mailSuffix           = $userSelection.MailSuffix

Write-Host "`nStarte Onboarding für: $Vorname $Nachname"

###############################################################################
# 5) Domain-/Branding-/Logging-Daten aus INI
###############################################################################
$companySection = $userSelection.CompanySection.Section
if (-not $Config.Contains($companySection)) { 
    Throw "Fehler: Die Sektion '$companySection' existiert nicht in der INI!" 
}
$companyData = $Config[$companySection]
$suffix = ($companySection -replace "\D", "")

$Strasse = $companyData["Strasse$suffix"]
$PLZ     = $companyData["PLZ$suffix"]
$Ort     = $companyData["Ort$suffix"]

if ($Config.Contains("MailEndungen")) {
    if (-not $mailSuffix -or $mailSuffix -eq "") { $mailSuffix = $Config.MailEndungen.Domain1 }
}

if ($companyData.Contains("Country$suffix")) { 
    $Country = $companyData["Country$suffix"]
} else { 
    $Country = "DE" 
}
if ($companyData.Contains("NameFirma$suffix")) {
    $companyDisplay = $companyData["NameFirma$suffix"]
} else {
    $companyDisplay = $Company
}
$defaultOU   = $Config.General["DefaultOU"]
$logFilePath = $Config.General["LogFilePath"]
$reportPath  = $Config.General["ReportPath"]
$reportTitle = $Config.General["ReportTitle"]
$reportFooter = $Config.General["ReportFooter"]

$firmaLogoPath = $Config.General["ReportLogo"]
$headerText   = $Config.General["ReportHeader"]
$footerText   = $Config.General["ReportFooter"]

$employeeLinks = @()
if ($Config.Contains("Websites")) {
    foreach ($key in $Config.Websites.Keys) {
        if ($key -match '^EmployeeLink\d+$') {
            $employeeLinks += $Config.Websites[$key]
        }
    }
}

if ($Standort -and ($Config.STANDORTE.Keys -contains "${Standort}_Bez")) {
    $standortDisplay = $Config.STANDORTE["${Standort}_Bez"]
} else {
    $standortDisplay = $Standort
}

$adSyncEnabled = $Config.ActivateUserMS365ADSync["ADSync"]
$adSyncGroup   = $Config.ActivateUserMS365ADSync["ADSyncADGroup"]
if ($adSyncEnabled -eq '1') { $adSyncEnabled = $true } else { $adSyncEnabled = $false }

###############################################################################
# 6) AD-Benutzer anlegen / aktualisieren
###############################################################################
try { Import-Module ActiveDirectory -ErrorAction Stop } catch { Write-Warning "AD-Modul konnte nicht geladen werden: $($_.Exception.Message)"; return }
if ([string]::IsNullOrWhiteSpace($Vorname)) { Throw "Vorname muss eingegeben werden!" }
function Generate-RandomPassword {
    param(
        [int]$Length,
        [bool]$IncludeSpecial,
        [bool]$AvoidAmbiguous
    )
    $minNonAlpha = 2
    if ($Config.PasswordFixGenerate.Contains("MinNonAlpha")) {
        $minNonAlpha = [int]$Config.PasswordFixGenerate["MinNonAlpha"]
    }
    $pw = [System.Web.Security.Membership]::GeneratePassword($Length, $minNonAlpha)
    if ($AvoidAmbiguous) { $pw = $pw -replace '[{}()\[\]\/\\`~,;:.<>\"]','X' }
    return $pw
}
if ($passwordMode -eq 1) {
    $UserPW = Generate-RandomPassword -Length $passwordLaenge -IncludeSpecial $includeSpecial -AvoidAmbiguous $avoidAmbiguous
    if ([string]::IsNullOrWhiteSpace($UserPW)) { $UserPW = "Standard123!" }
} else { $UserPW = $fixPassword }
$SecurePW = ConvertTo-SecureString $UserPW -AsPlainText -Force
$SamAccountName = ($Vorname.Substring(0,1) + $Nachname).ToLower()
if ($UPNManual) { 
    $UPN = $UPNManual 
} else {
    if ( ($Config.Keys -contains $companySection) -and 
         ($Config[$companySection].Keys -contains ("ActiveDirectoryDomain" + $suffix)) -and 
         -not [string]::IsNullOrWhiteSpace($companyData["ActiveDirectoryDomain" + $suffix]) ) {
        $adDomain = "@" + $companyData["ActiveDirectoryDomain" + $suffix].Trim()
    } else {
        $adDomain = $mailSuffix
    }
    switch -Wildcard ($UPNTemplate) {
        "VORNAME.NACHNAME"    { $UPN = "$Vorname.$Nachname$adDomain" }
        "V.NACHNAME"          { $UPN = "$($Vorname.Substring(0,1)).$Nachname$adDomain" }
        "VORNAMENACHNAME"     { $UPN = "$Vorname$Nachname$adDomain" }
        "VNACHNAME"           { $UPN = "$($Vorname.Substring(0,1))$Nachname$adDomain" }
        Default               { $UPN = "$SamAccountName$adDomain" }
    }
}

if ( ($Config.Keys -contains $companySection) -and 
     ($companyData.Keys -contains ("NameFirma" + $suffix)) -and 
     -not [string]::IsNullOrWhiteSpace($companyData["NameFirma" + $suffix]) ) {
    $prefix = $companyData["NameFirma" + $suffix].Trim()
} else {
    $prefix = $userSelection.CompanySection.Display
}

$DisplayName = "$prefix | $Vorname $Nachname"
Write-Host "DisplayName wird gesetzt: $DisplayName"
Write-Host "SamAccountName : $SamAccountName"
Write-Host "UPN            : $UPN"
Write-Host "Passwort       : $UserPW"
try { 
    $existingUser = Get-ADUser -Filter { SamAccountName -eq $SamAccountName } -ErrorAction SilentlyContinue 
} catch { 
    $existingUser = $null 
}

$userParams = @{
    Name                  = $DisplayName
    DisplayName           = $DisplayName
    GivenName             = $Vorname
    Surname               = $Nachname
    SamAccountName        = $SamAccountName
    UserPrincipalName     = $UPN
    AccountPassword       = $SecurePW
    Enabled               = (-not $userSelection.AccountDisabled)
    ChangePasswordAtLogon = $userSelection.MustChangePassword
    PasswordNeverExpires  = $userSelection.PasswordNeverExpires
    Path                  = $defaultOU
    City                  = $Ort
    StreetAddress         = $Strasse
    Country               = $Country
}
if ($Config.ADUserDefaults.Contains("ProfilePath")) {
    $userParams["ProfilePath"] = $Config.ADUserDefaults["ProfilePath"]
}
if ($Config.ADUserDefaults.Contains("LogonScript")) {
    $userParams["ScriptPath"] = $Config.ADUserDefaults["LogonScript"]
}
if ($Config.ADUserDefaults.Contains("LocalProfilePath")) {
    $userParams["HomeDirectory"] = $Config.ADUserDefaults["LocalProfilePath"]
}

if (-not $existingUser) {
    Write-Host "Erstelle neuen Benutzer: $DisplayName"
    $otherAttributes = @{}
    if ($EmailAddress -and $EmailAddress.Trim() -ne "") {
        if ($EmailAddress -notmatch "@") {
            $otherAttributes["mail"] = "$EmailAddress$mailSuffix"
        }
        else {
            $otherAttributes["mail"] = $EmailAddress
        }
    }
    if ($Description -and $Description.Trim() -ne "") { $otherAttributes["description"] = $Description }
    if ($OfficeRoom -and $OfficeRoom.Trim() -ne "") { $otherAttributes["physicalDeliveryOfficeName"] = $OfficeRoom }
    if ($PhoneNumber -and $PhoneNumber.Trim() -ne "") { $otherAttributes["telephoneNumber"] = $PhoneNumber }
    if ($MobileNumber -and $MobileNumber.Trim() -ne "") { $otherAttributes["mobile"] = $MobileNumber }
    if ($Position -and $Position.Trim() -ne "") { $otherAttributes["title"] = $Position }
    if ($DepartmentField -and $DepartmentField.Trim() -ne "") { $otherAttributes["department"] = $DepartmentField }
   
    $filteredAttrs = @{}
    foreach ($key in $otherAttributes.Keys) {
        if ($otherAttributes[$key] -and $otherAttributes[$key].Trim() -ne "") {
            $filteredAttrs[$key] = $otherAttributes[$key]
        }
    }
    if ($filteredAttrs.Count -gt 0) {
        $userParams["OtherAttributes"] = $filteredAttrs
    }
    
    try {
        New-ADUser @userParams -ErrorAction Stop
        Write-Host "AD-Benutzer erstellt."
    } catch {
        Write-Warning "Fehler beim Erstellen des Benutzers: $($_.Exception.Message)"
        Write-Host "Weitere Details:"
        $_ | Format-List * -Force
        return
    }
    
    if ($userSelection.SmartcardLogonRequired) {
        try { Set-ADUser -Identity $SamAccountName -SmartcardLogonRequired $true -ErrorAction Stop } catch { Write-Warning "Fehler bei SmartcardLogonRequired: $($_.Exception.Message)" }
    }
    if ($userSelection.CannotChangePassword) {
        Write-Host "(Hinweis: 'CannotChangePassword' via ACL müsste hier umgesetzt werden.)"
    }
} else {
    Write-Host "Benutzer '$SamAccountName' existiert bereits - Update erfolgt."
    try {
        Set-ADUser -Identity $existingUser.DistinguishedName `
            -GivenName $Vorname `
            -Surname $Nachname `
            -City $Ort `
            -StreetAddress $Strasse `
            -Country $Country `
            -Enabled (-not $userSelection.AccountDisabled) `
            -ErrorAction SilentlyContinue
        Set-ADUser -Identity $existingUser.DistinguishedName -ChangePasswordAtLogon:$userSelection.MustChangePassword -PasswordNeverExpires:$userSelection.PasswordNeverExpires
    } catch { Write-Warning "Fehler beim Aktualisieren: $($_.Exception.Message)" }
}
try { 
    Set-ADAccountPassword -Identity $SamAccountName -Reset -NewPassword $SecurePW -ErrorAction SilentlyContinue 
} catch { Write-Warning "Fehler beim Setzen des Passworts: $($_.Exception.Message)" }
if ($cannotChangePW) { Write-Host "(Hinweis: 'CannotChangePassword' via ACL wäre hier umzusetzen.)" }

###############################################################################
# 7) AD-Gruppen zuweisen
###############################################################################
if (-not $userSelection.Extern) {
    foreach ($groupKey in $userSelection.ADGroupsSelected) {
        $groupName = $Config.ADGroups[$groupKey]
        if ($groupName) {
            try { Add-ADGroupMember -Identity $groupName -Members $SamAccountName -ErrorAction Stop } catch { Write-Warning "Fehler bei AD-Gruppe '$groupName': $($_.Exception.Message)" }
        }
    }
} else {
    Write-Host "Externer Mitarbeiter: Standardmäßige AD-Gruppen-Zuweisung wird übersprungen."
}
if ($Standort) {
    $signaturKey = $Config.STANDORTE[$Standort]
    if ($signaturKey) {
        $signaturGroup = $Config.SignaturGruppe_Optional[$signaturKey]
        if ($signaturGroup) {
            try { Add-ADGroupMember -Identity $signaturGroup -Members $SamAccountName -ErrorAction SilentlyContinue } catch { Write-Warning "Fehler bei Signatur-Gruppe: $($_.Exception.Message)" }
        }
    }
}
if ($License) {
    $licenseKey = "MS365_" + $License
    $licenseGroup = $Config.LicensesGroups[$licenseKey]
    if ($licenseGroup) {
        try { Add-ADGroupMember -Identity $licenseGroup -Members $SamAccountName -ErrorAction SilentlyContinue } catch { Write-Warning "Fehler bei Lizenz-Gruppe: $($_.Exception.Message)" }
    }
}
if ($adSyncEnabled -and $adSyncGroup) {
    try { Add-ADGroupMember -Identity $adSyncGroup -Members $SamAccountName -ErrorAction SilentlyContinue } catch { Write-Warning "Fehler bei ADSync-Gruppe: $($_.Exception.Message)" }
}
if ($userSelection.Extern) { Write-Host "Externer Mitarbeiter: Bitte weisen Sie alle AD-Gruppen händisch zu." }

###############################################################################
# 8) Logging
###############################################################################
try {
    if (-not (Test-Path $logFilePath)) { New-Item -ItemType Directory -Path $logFilePath -Force | Out-Null }
    $logDate = (Get-Date -Format 'yyyyMMdd')
    $logFile = Join-Path $logFilePath "Onboarding_$logDate.log"
    $currentUser = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
    $logEntry = "ONBOARDING DURCHGEFÜHRT VON: $currentUser`r`n" +
                "[{0}] Sam={1}, Anzeigename='{2}', UPN='{3}', Standort='{4}', Company='{5}', Location='{6}', MS365 Lizenz='{7}', ADGruppen=({8}), Passwort='{9}', Extern={10}" -f (Get-Date), $SamAccountName, $DisplayName, $UPN, $standortDisplay, $companyDisplay, $standortDisplay, $MS365License, ($userSelection.ADGroupsSelected -join ','), $UserPW, $Extern
    Add-Content -Path $logFile -Value $logEntry
    Write-Host "Log geschrieben: $logFile"
} catch { Write-Warning "Fehler beim Logging: $($_.Exception.Message)" }

###############################################################################
# 9) Reports erzeugen (HTML, PDF, TXT)
###############################################################################
try {
    if (-not (Test-Path $reportPath)) {
        New-Item -ItemType Directory -Path $reportPath -Force | Out-Null
    }

    if ($createHTML) {
        # HTML-Report-Pfad
        $htmlFile = Join-Path $reportPath "$SamAccountName.html"

        # Logo oben rechts (ReportLogo)
        $logoTag = ""
        if ($firmaLogoPath -and (Test-Path $firmaLogoPath)) {
            $logoTag = "<div class='logo' style='float:right;'><img src='file:///$firmaLogoPath' style='max-width:150px; margin:10px;'/></div>"
        }

        # Relevante Benutzerdaten in einer Tabelle
        $userDetails = [ordered]@{
            "Vorname"         = $Vorname
            "Nachname"        = $Nachname
            "Description"     = $Description
            "Buero"           = $OfficeRoom
            "Rufnummer"       = $PhoneNumber
            "Mobil"           = $MobileNumber
            "Position"        = $Position
            "Abteilung"       = $DepartmentField
            "Ablaufdatum"     = $Ablaufdatum
            "Company"         = $companyDisplay
            "LoginName"       = $SamAccountName
        }
        $userDetailsHTML = ""
        foreach ($key in $userDetails.Keys) {
            $userDetailsHTML += "<tr><td><strong>$key</strong></td><td>$($userDetails[$key])</td></tr>`r`n"
        }

        # Webseiten aus der INI – z. B. EmployeeLinks
        $websitesHTML = ""
        if ($INIConfig.Contains("Websites")) {
            foreach ($key in $INIConfig.Websites.Keys) {
                if ($key -match '^EmployeeLink\d+$') {
                    $websitesHTML += "<li><a href='$($INIConfig.Websites[$key])' target='_blank'>$($INIConfig.Websites[$key])</a></li>`r`n"
                }
            }
            if ($websitesHTML -ne "") {
                $websitesHTML = "<ul>" + $websitesHTML + "</ul>"
            }
        }

        # Lese den HTML-Report-Template-Pfad aus der INI (Sektion "Branding-Report", Key "TemplatePath")
        if ($reportBranding -and $reportBranding.TemplatePath -and (Test-Path $reportBranding.TemplatePath)) {
            $templatePath = $reportBranding.TemplatePath
        } else {
            $templatePath = Join-Path -Path $PSScriptRoot -ChildPath "REPORTTemplate.txt"
        }
        $htmlTemplate = Get-Content -Path $templatePath -Raw

        # Ersetze Platzhalter im Template
        $htmlContent = $htmlTemplate -replace "{{ReportTitle}}", $reportTitle `
                                      -replace "{{LogoTag}}", $logoTag `
                                      -replace "{{UserDetailsHTML}}", $userDetailsHTML `
                                      -replace "{{WebsitesHTML}}", $websitesHTML `
                                      -replace "{{ReportFooter}}", ($reportBranding.ReportFooter -or $reportFooter)

        # Speichere den HTML-Inhalt
        Set-Content -Path $htmlFile -Value $htmlContent -Encoding UTF8
        Write-Host "HTML-Report erstellt: $htmlFile"

        # Weitere Report-Formate (PDF, TXT) können hier ergänzt werden
    }
    else {
        Write-Host "HTML-Report: Deaktiviert / nicht gewünscht."
    }
}
catch {
    Write-Warning "Fehler beim Erstellen der Reports: $($_.Exception.Message)"
}
Write-Host "`nOnboarding abgeschlossen."

# SIG # Begin signature block
# MIIbywYJKoZIhvcNAQcCoIIbvDCCG7gCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDxwgnT2sm8jv10
# puROKcTFLjNlUlX2TcuKtkajYOXYAqCCFhcwggMQMIIB+KADAgECAhB3jzsyX9Cg
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
# DQEJBDEiBCDb0EqW0W8FBRGpPjUCANtx2hPwacKQ4GpT6B1chYedgjANBgkqhkiG
# 9w0BAQEFAASCAQAZbN5DNP9wVkvWHqQAEEQHuosKj/bSVt1jAY9Cl5HUO6i5BdiA
# lSKaUQuWPjUzxS4Bwi4jSZYSQh7AmVHbMckYRZixYXg/v3thJEZhMhds3+uaqDca
# 0bd+LGL+kAtGnMi10C83RqHc83Ud7DkRwuKhWueMmRHJFq6wVnFuzbXcm9DMJ8wE
# sC9QKDXcwOSx/3J6BRqmadlPhsuKRI3Y7yK81ZQxKNE6ZJsn7AFoPa5kqZflFY4F
# r6mj3YoyGmPNWtv2uCBVjk0+jEbFAkWMvWgawbuYj/6NCgciU1gwKVx7SYShwH4t
# CfERHJPwHLeKidMx+ukvREc3pfdoA/vsW5eUoYIDIDCCAxwGCSqGSIb3DQEJBjGC
# Aw0wggMJAgEBMHcwYzELMAkGA1UEBhMCVVMxFzAVBgNVBAoTDkRpZ2lDZXJ0LCBJ
# bmMuMTswOQYDVQQDEzJEaWdpQ2VydCBUcnVzdGVkIEc0IFJTQTQwOTYgU0hBMjU2
# IFRpbWVTdGFtcGluZyBDQQIQC65mvFq6f5WHxvnpBOMzBDANBglghkgBZQMEAgEF
# AKBpMBgGCSqGSIb3DQEJAzELBgkqhkiG9w0BBwEwHAYJKoZIhvcNAQkFMQ8XDTI1
# MDcwNTEwMTE1NlowLwYJKoZIhvcNAQkEMSIEIHCkXlzjfSUDovyHa7pwB7PfZwVk
# Au8W6cNczBabkuxpMA0GCSqGSIb3DQEBAQUABIICAI7EDnnpxP+1aoZ+aKASh/gk
# kzZHU/DbtM7hTx6BxW5OvAYd2IuiTDNo4xLMHmxyx/kBKE0lyGYud2Ru5a1I+Tlv
# 4YEV3dF12qb5Hm0TQV8PsSryi779tSQv7VNkIZ6f2zYHIOIqVH4xPhAjc4QWkoe4
# UEc9masI4DMgzsEDoRMIQztViB+yton5edRwusE9ujD8enofUSibLm3FqMAr+0cC
# qpyfnxwxabWxqcCWwZiqejUit8lgboevt0szz17sg+I+Os1FfBAlGEO47s9yDfPU
# Zq3G/7aqb9Qp8gfpbgQJ7xr7AZqHpuYfgfhKz1WYGCOGfjPVeiuWU+vv2+k6nfzL
# cfNQXO0OZUVhQGxvj/qkf9cdl7u0KDjaMhL0nfZfsiyVVu+hYbn1+sxi/61cVHT6
# 7QXu9dajttcMIPFdMcnHx4qYRBpVbE0SgFnCd/JRcfAq9fV+p/M6881S5SJV7LwS
# MmxmHr9jpY3au2S3JgUft7MAwXb3xIgI3sd/34jfwjBhkIFLyqgBxh2y26bqu0me
# LyyZATV0ysfLayiLEoyxBL9u1MaFAf8kcRdpG+DQKL8IFyJLq0TqY2w/HwqzKMQi
# qhZDW1QKouHmPD8cK8H4/tI+V9vyztSkvtAtdwqwVIwl6rh4Ou3Crw70kl23FTcx
# J5OG63UECYW6j6y86JvP
# SIG # End signature block
