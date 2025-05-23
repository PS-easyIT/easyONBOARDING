#requires -Version 5.1
[CmdletBinding()]
param(
    [string]$Vorname,
    [string]$Nachname,
    [string]$Standort,
    [string]$Company,         # Erwartet hier z. B. "1" für Company1, "2" für Company2 etc.
    [string]$License = "",
    [switch]$Extern,
    [string]$ScriptINIPath = "C:\SCRIPT\easyOnboardingConfig.ini"
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
# 2) Hilfsfunktionen (erlauben leere Strings)
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

    # --- Beschriftung und DropDown für Company ---
    AddLabel $panelLeft "Firma:" 10 $yLeft -Bold | Out-Null
    $companyOptions = @()
    foreach ($section in $INIConfig.Keys | Where-Object { $_ -like "Company*" }) {
        $suffix = ($section -replace "\D", "")
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
    # --------------------------------------------------------

    AddLabel $panelLeft "MS365 Lizenz*:" 10 $yLeft -Bold | Out-Null
    $cmbMS365License = AddComboBox $panelLeft ( @("KEINE") + ($INIConfig.LicensesGroups.Keys | ForEach-Object { $_ -replace '^MS365_','' } ) ) 150 $yLeft 200 ""; $yLeft += 55

    AddLabel $panelLeft "ONBOARDING DOKUMENT ERZEUGEN?" 10 $yLeft -Bold | Out-Null
    $yLeft += 20
    $chkHTML_Left = AddCheckBox $panelLeft "HTML erzeugen" $true 10 $yLeft
    $chkPDF_Left  = AddCheckBox $panelLeft "PDF erzeugen" $true 150 $yLeft
    $chkTXT_Left  = AddCheckBox $panelLeft "TXT erzeugen" $true 290 $yLeft; $yLeft += 35

    ############################################################################
    # Elemente im PanelRight (Eingabe: E-Mail, Mail-Endung, UPN, Flags, etc.)
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
                $y = 10 + ($row * 30)
                $cbGroup = AddCheckBox $panelADGroups $displayText $false $x $y
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
    $btnOnboard.Text = "ONBOARDEN"
    $btnOnboard.Size = New-Object System.Drawing.Size($btnWidth, $btnHeight)
    $btnOnboard.Location = New-Object System.Drawing.Point($startX, 15)
    $btnOnboard.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnOnboard.BackColor = [System.Drawing.Color]::LightGreen
    $panelBottom.Controls.Add($btnOnboard)
    
    $btnInfo = New-Object System.Windows.Forms.Button
    $btnInfo.Text = "INFO"
    $btnInfo.Size = New-Object System.Drawing.Size($btnWidth, $btnHeight)
    $btnInfo.Location = New-Object System.Drawing.Point([int]($startX + $btnWidth + $btnSpacing), 15)
    $btnInfo.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnInfo.BackColor = [System.Drawing.Color]::LightBlue
    $panelBottom.Controls.Add($btnInfo)
    
    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "ABBRECHEN"
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
        CompanySection        = ""
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
        UPNFormat             = "VORNAME.NACHNAME"
        EmailAddress          = ""
        MailSuffix            = ""
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
        # Vor dem Zugriff: Prüfe, ob eine Company ausgewählt wurde
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

        # Speichere den ausgewählten Company-Abschnitt direkt in das Ergebnis
        $result.CompanySection = $cmbCompany.SelectedItem.Section

        $selectedCompany = $cmbCompany.SelectedItem
        if (-not $selectedCompany) {
            Throw "Fehler: Es wurde keine Company ausgewählt. Bitte wählen Sie einen Bereich aus."
        }
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
                } elseif (-not [string]::IsNullOrWhiteSpace($mailSuffix)) {
                    $emailInput = "$emailInput$mailSuffix"
                }
            }
        }
        $result.EmailAddress = $emailInput

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
$companySection = $userSelection.CompanySection
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
        if ($key -match '^EmployeeLink\d+$') { $employeeLinks += $Config.Websites[$key] }
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
    if (-not (Test-Path $reportPath)) { New-Item -ItemType Directory -Path $reportPath -Force | Out-Null }
    if ($createHTML) {
        $htmlFile = Join-Path $reportPath "$SamAccountName.html"
        $logoTag = ""
        if ($firmaLogoPath -and (Test-Path $firmaLogoPath)) {
            $logoTag = "<img src='file:///$firmaLogoPath' style='float:right; max-width:120px; margin:10px;'/>"
        }
        $reportData = [ordered]@{
            Vorname             = $Vorname
            Nachname            = $Nachname
            Anzeigename         = $DisplayName
            Beschreibung        = $Description
            Buero               = $OfficeRoom
            Rufnummer           = $PhoneNumber
            Mobil               = $MobileNumber
            Position            = $Position
            Abteilung           = $DepartmentField
            Location            = $standortDisplay
            Company             = $companyDisplay
            MS365Lizenz         = $MS365License
            UPNEntered          = $UPNManual
            UPNFormat           = $cmbUPNFormat.SelectedItem
            EmailAddress        = $EmailAddress
            FixPassword         = $fixPassword
            Passwortlaenge      = $passwordLaenge
            IncludeSpecialChars = $includeSpecial
            AvoidAmbiguousChars = $avoidAmbiguous
        }
        $htmlContent = @"
<html>
<head>
  <meta charset='UTF-8'>
  <title>$reportTitle</title>
  <style>
    body { font-family: $($reportBranding.ReportFontFamily -or "Arial"); font-size: $($reportBranding.ReportFontSize -or "10")pt; background-color: $($reportBranding.ReportThemeColor -or "#FFFFFF"); margin:20px; }
    table { border-collapse: collapse; width: 100%; }
    th, td { border: 1px solid #ccc; padding: 5px; text-align: left; }
    th { background-color: #f0f0f0; }
  </style>
</head>
<body>
  <h1>$($reportBranding.ReportHeader -or "Report Header")</h1>
  <h2>$reportTitle</h2>
  $logoTag
  <table>
    <tr><th>Feld</th><th>Wert</th></tr>
"@
        foreach ($prop in $reportData.PSObject.Properties) {
            $htmlContent += "<tr><td>$($prop.Name)</td><td>$($prop.Value)</td></tr>`r`n"
        }
        $htmlContent += @"
  </table>
  <footer><p>$($reportBranding.ReportFooter -or $reportFooter)</p></footer>
</body>
</html>
"@
        Set-Content -Path $htmlFile -Value $htmlContent -Encoding UTF8
        Write-Host "HTML-Report erstellt: $htmlFile"

        if ($createPDF) {
            $pdfFile = [System.IO.Path]::ChangeExtension($htmlFile, ".pdf")
            $wkhtml = "C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe"
            if (Test-Path $wkhtml) {
                & $wkhtml $htmlFile $pdfFile
                Write-Host "PDF-Report erstellt: $pdfFile"
            } else { Write-Warning "wkhtmltopdf.exe nicht gefunden -> kein PDF erstellt." }
        }
        if ($createTXT) {
            $txtFile = [System.IO.Path]::ChangeExtension($htmlFile, ".txt")
            $txtContent = "Onboarding Report`r`n=================`r`n"
            foreach ($prop in $reportData.PSObject.Properties) {
                $txtContent += "$($prop.Name) : $($prop.Value)`r`n"
            }
            Set-Content -Path $txtFile -Value $txtContent -Encoding UTF8
            Write-Host "TXT-Report erstellt: $txtFile"
        }
    }
    else { Write-Host "HTML-Report: Deaktiviert / nicht gewünscht." }
}
catch { Write-Warning "Fehler beim Erstellen der Reports: $($_.Exception.Message)" }

Write-Host "`nOnboarding abgeschlossen."