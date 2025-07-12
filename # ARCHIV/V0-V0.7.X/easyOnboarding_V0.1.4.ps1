#requires -Version 5.1
[CmdletBinding()]
param(
    [string]$Vorname,
    [string]$Nachname,
    [string]$Standort,
    [string]$Company,
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
# 2) GUI-Erstellung
###############################################################################
function Show-OnboardingForm {
    param(
        [hashtable]$INIConfig
    )
    # Branding-Bereiche getrennt laden
    $guiBranding = if ($INIConfig.Contains("Branding-GUI")) { $INIConfig["Branding-GUI"] } else { @{} }
    $reportBranding = if ($INIConfig.Contains("Branding-Report")) { $INIConfig["Branding-Report"] } else { @{} }

    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    # Hauptfenster
    $form = New-Object System.Windows.Forms.Form
    $form.Text = if ($guiBranding.APPName) { $guiBranding.APPName } else { "easyONBOARDING" }
    $form.StartPosition = "CenterScreen"
    $form.Size = New-Object System.Drawing.Size(1085,900)
    $form.AutoScroll = $true

    # Hintergrundbild (GUI)
    if ($guiBranding.BackgroundImage -and (Test-Path $guiBranding.BackgroundImage)) {
        $form.BackgroundImage = [System.Drawing.Image]::FromFile($guiBranding.BackgroundImage)
        $form.BackgroundImageLayout = 'Stretch'
    }

    # Info-Label oben (links im Header)
    $lblInfo = New-Object System.Windows.Forms.Label
    $lblInfo.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 10, [System.Drawing.FontStyle]::Bold)
    $lblInfo.Location = New-Object System.Drawing.Point(10,10)
    $lblInfo.AutoSize = $true
    $scriptInfo = if ($INIConfig.Contains("ScriptInfo")) { $INIConfig["ScriptInfo"] } else { @{} }
    $general     = if ($INIConfig.Contains("General")) { $INIConfig["General"] } else { @{} }
    $lblInfo.Text = "ScriptVersion=$($scriptInfo.ScriptVersion) | LastUpdate=$($scriptInfo.LastUpdate) | Author=$($scriptInfo.Author)`r`nDOMAIN: $($general.DomainName1) | OU: $($general.DefaultOU) | REPORT: $($general.ReportPath)"
    $form.Controls.Add($lblInfo)

    # Header-Logo (GUI)
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

    # Panels für Eingabe und Ausgabe
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

    # Footer-Panel (GUI)
    $panelFooter = New-Object System.Windows.Forms.Panel
    $panelFooter.Dock = 'Bottom'
    $panelFooter.Height = 40
    $panelFooter.BorderStyle = 'FixedSingle'
    $lblFooter = New-Object System.Windows.Forms.Label
    $lblFooter.AutoSize = $true
    $lblFooter.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 8)
    $lblFooter.Location = New-Object System.Drawing.Point(10,10)
    $lblFooter.Text = if ($guiBranding.FooterWebseite) { $guiBranding.FooterWebseite } else { "www.easyONBOARDING.com" }
    $panelFooter.Controls.Add($lblFooter)
    $form.Controls.Add($panelFooter)

    ############################################################################
    # Hilfsfunktionen: AddLabel, AddTextBox, AddCheckBox, AddComboBox
    ############################################################################
    function AddLabel {
        param(
            [Parameter(Mandatory=$true)] $parent,
            [Parameter(Mandatory=$true)] [string]$text,
            [Parameter(Mandatory=$true)] [int]$x,
            [Parameter(Mandatory=$true)] [int]$y,
            [switch]$Bold
        )
        $lbl = New-Object System.Windows.Forms.Label
        $lbl.Text = $text
        $lbl.Location = New-Object System.Drawing.Point($x, $y)
        if ($Bold) {
            $lbl.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 8, [System.Drawing.FontStyle]::Bold)
        }
        $lbl.AutoSize = $true
        $parent.Controls.Add($lbl)
        return $lbl
    }
    function AddTextBox($parent, [string]$default, [int]$x, [int]$y, [int]$width=200) {
        $tb = New-Object System.Windows.Forms.TextBox
        $tb.Text = $default
        $tb.Location = New-Object System.Drawing.Point($x, $y)
        $tb.Width = $width
        $parent.Controls.Add($tb)
        return $tb
    }
    function AddCheckBox($parent, [string]$text, [bool]$checked, [int]$x, [int]$y) {
        $cb = New-Object System.Windows.Forms.CheckBox
        $cb.Text = $text
        $cb.Location = New-Object System.Drawing.Point($x, $y)
        $cb.Checked = $checked
        $cb.AutoSize = $true
        $parent.Controls.Add($cb)
        return $cb
    }
    function AddComboBox($parent, [string[]]$items, [int]$x, [int]$y, [int]$width=150, [string]$default='') {
        $cmb = New-Object System.Windows.Forms.ComboBox
        $cmb.DropDownStyle = 'DropDownList'
        $cmb.Location = New-Object System.Drawing.Point($x, $y)
        $cmb.Width = $width
        foreach ($i in $items) { [void]$cmb.Items.Add($i) }
        if ($default -and $cmb.Items.Contains($default)) {
            $cmb.SelectedItem = $default
        } elseif ($cmb.Items.Count -gt 0) {
            $cmb.SelectedIndex = 0
        }
        $parent.Controls.Add($cmb)
        return $cmb
    }

    ############################################################################
    # Elemente im PanelLeft (z. B. Vorname, Nachname, Anzeigename, etc.)
    ############################################################################
    $yLeft = 10
    AddLabel $panelLeft "Vorname:" 10 $yLeft -Bold | Out-Null
    $txtVorname = AddTextBox $panelLeft "" 150 $yLeft; $yLeft += 30

    AddLabel $panelLeft "Nachname:" 10 $yLeft -Bold | Out-Null
    $txtNachname = AddTextBox $panelLeft "" 150 $yLeft; $yLeft += 30

    $chkExternal = AddCheckBox $panelLeft "Externer Mitarbeiter" $false 10 $yLeft
    $yLeft += 35

    AddLabel $panelLeft "Anzeigename:" 10 $yLeft -Bold | Out-Null
    $txtDisplayName = AddTextBox $panelLeft "" 150 $yLeft; $yLeft += 30

    AddLabel $panelLeft "Anzeigename Vorlage:" 10 $yLeft -Bold | Out-Null
    $templates = @()
    if ($INIConfig.Contains("DisplayNameTemplates")) {
        $templates = $INIConfig["DisplayNameTemplates"].Keys | ForEach-Object { $INIConfig["DisplayNameTemplates"][$_] }
    }
    $cmbDisplayNameTemplate = AddComboBox $panelLeft $templates 150 $yLeft 250 ""
    $yLeft += 30

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
    $txtDeptField = AddTextBox $panelLeft "" 150 $yLeft; $yLeft += 30

    AddLabel $panelLeft "CustomAttrib1:" 10 $yLeft -Bold | Out-Null
    $txtCA1 = AddTextBox $panelLeft "" 150 $yLeft; $yLeft += 30

    AddLabel $panelLeft "CustomAttrib2:" 10 $yLeft -Bold | Out-Null
    $txtCA2 = AddTextBox $panelLeft "" 150 $yLeft; $yLeft += 30

    AddLabel $panelLeft "CustomAttrib3:" 10 $yLeft -Bold | Out-Null
    $txtCA3 = AddTextBox $panelLeft "" 150 $yLeft; $yLeft += 30

    AddLabel $panelLeft "CustomAttrib4:" 10 $yLeft -Bold | Out-Null
    $txtCA4 = AddTextBox $panelLeft "" 150 $yLeft; $yLeft += 30

    AddLabel $panelLeft "CustomAttrib5:" 10 $yLeft -Bold | Out-Null
    $txtCA5 = AddTextBox $panelLeft "" 150 $yLeft; $yLeft += 30

    AddLabel $panelLeft "Location*:" 10 $yLeft -Bold | Out-Null
    # Sammle alle _Bez-Schlüssel aus [STANDORTE] und verwende deren Werte
    $locationDisplayList = $INIConfig.STANDORTE.Keys |
        Where-Object { $_ -match '_Bez$' } |
        ForEach-Object { $INIConfig.STANDORTE[$_] }
    $cmbLocation = AddComboBox $panelLeft $locationDisplayList 150 $yLeft 250 ""
    $yLeft += 30

    AddLabel $panelLeft "Company*:" 10 $yLeft -Bold | Out-Null
    $cmbCompany = AddComboBox $panelLeft (($INIConfig.Keys | Where-Object { $_ -like "DomainName*" })) 150 $yLeft 250 ""
    $yLeft += 30

    AddLabel $panelLeft "MS365 Lizenz*:" 10 $yLeft -Bold | Out-Null
    $cmbMS365License = AddComboBox $panelLeft ( @("KEINE") + ($INIConfig.LicensesGroups.Keys | ForEach-Object { $_ -replace '^MS365_','' } ) ) 150 $yLeft 200 ""
    $yLeft += 50

    AddLabel $panelLeft "DOKUMENT ERZEUGEN?" 10 $yLeft -Bold | Out-Null
    $yLeft += 20
    $chkHTML_Left = AddCheckBox $panelLeft "HTML erzeugen" $true 10 $yLeft
    $chkPDF_Left  = AddCheckBox $panelLeft "PDF erzeugen" $true 150 $yLeft
    $chkTXT_Left  = AddCheckBox $panelLeft "TXT erzeugen" $true 290 $yLeft
    $yLeft += 35

    ############################################################################
    # Elemente im PanelRight
    ############################################################################
    $yRight = 10

    AddLabel $panelRight "E-Mail-Adresse:" 10 $yRight -Bold | Out-Null
    $txtEmail = AddTextBox $panelRight "" 150 $yRight 200
    $yRight += 35

    # Neues Element: Dropdown für Mail-Endungen
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

    AddLabel $panelRight "Benutzer Name (UPN):" 10 $yRight -Bold | Out-Null
    $txtUPN = AddTextBox $panelRight "" 150 $yRight 200
    $yRight += 35

    AddLabel $panelRight "UPN-Format-Vorlagen:" 10 $yRight -Bold | Out-Null
    $cmbUPNFormat = AddComboBox $panelRight @("VORNAME.NACHNAME","V.NACHNAME","VORNAMENACHNAME","VNACHNAME") 150 $yRight 200
    $yRight += 50

    AddLabel $panelRight "AD-Benutzer-Flags:" 10 $yRight -Bold | Out-Null
    $yRight += 20
    $chkPWNeverExpires = AddCheckBox $panelRight "PasswordNeverExpires" $false 10 $yRight
    $chkMustChange     = AddCheckBox $panelRight "MustChangePasswordAtLogon" $false 150 $yRight
    $chkAccountDisabled = AddCheckBox $panelRight "AccountDisabled" $false 10 ($yRight + 35)
    $chkCannotChangePW  = AddCheckBox $panelRight "CannotChangePassword" $false 150 ($yRight + 35)
    $chkSmartcardLogonRequired = AddCheckBox $panelRight "SmartcardLogonRequired" $false 300 ($yRight + 35)
    $yRight += 75

    AddLabel $panelRight "PASSWORT-OPTIONEN:" 10 $yRight -Bold | Out-Null
    $yRight += 25
    $rbFix = New-Object System.Windows.Forms.RadioButton
    $rbFix.Text = "FEST"
    $rbFix.Location = New-Object System.Drawing.Point(10, $yRight)
    $panelRight.Controls.Add($rbFix)
    $rbRand = New-Object System.Windows.Forms.RadioButton
    $rbRand.Text = "GENERIERT"
    $rbRand.Location = New-Object System.Drawing.Point(150, $yRight)
    $panelRight.Controls.Add($rbRand)
    $yRight += 35
    AddLabel $panelRight "Festes Passwort:" 10 $yRight -Bold | Out-Null
    $txtFixPW = AddTextBox $panelRight "" 150 $yRight 150
    $yRight += 35
    AddLabel $panelRight "Passwortlaenge:" 10 $yRight -Bold | Out-Null
    $txtPWLen = AddTextBox $panelRight "12" 150 $yRight 50
    $yRight += 35
    $chkIncludeSpecial = AddCheckBox $panelRight "IncludeSpecialChars" $true 10 $yRight
    $chkAvoidAmbig     = AddCheckBox $panelRight "AvoidAmbiguousChars" $true 150 $yRight
    $yRight += 45

    AddLabel $panelRight "AD-Gruppen:" 10 $yRight -Bold | Out-Null
    $yRight += 25
    $panelADGroups = New-Object System.Windows.Forms.Panel
    $panelADGroups.Location = New-Object System.Drawing.Point(10, $yRight)
    $panelADGroups.Size = New-Object System.Drawing.Size(480,150)
    $panelADGroups.AutoScroll = $true
    $panelRight.Controls.Add($panelADGroups)
    $yRight += $panelADGroups.Height + 10
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
        AddLabel $panelRight "Keine [ADGroups] Sektion gefunden." 10 $yRight -Bold | Out-Null
        $yRight += 25
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
    # Ergebnis-Objekt
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
        CustomAttribute1      = ""
        CustomAttribute2      = ""
        CustomAttribute3      = ""
        CustomAttribute4      = ""
        CustomAttribute5      = ""
        Location              = ""
        CompanySelection      = ""
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
        $result.Vorname         = $txtVorname.Text
        $result.Nachname        = $txtNachname.Text
        $result.DisplayName     = $txtDisplayName.Text
        $result.Description     = $txtDescription.Text
        $result.OfficeRoom      = $txtOffice.Text
        $result.PhoneNumber     = $txtPhone.Text
        $result.MobileNumber    = $txtMobile.Text
        $result.Position        = $txtPosition.Text
        $result.DepartmentField = $txtDeptField.Text
        $result.CustomAttribute1= $txtCA1.Text
        $result.CustomAttribute2= $txtCA2.Text
        $result.CustomAttribute3= $txtCA3.Text
        $result.CustomAttribute4= $txtCA4.Text
        $result.CustomAttribute5= $txtCA5.Text

        $result.Location = $cmbLocation.SelectedItem
        $result.CompanySelection = $cmbCompany.SelectedItem
        $result.MS365License = $cmbMS365License.SelectedItem

        if ($chkExternal.Checked) {
            if ([string]::IsNullOrWhiteSpace($txtDisplayName.Text)) {
                $template = $cmbDisplayNameTemplate.SelectedItem
                $displayNameTemplate = $template -replace "\{first\}", $txtVorname.Text -replace "\{last\}", $txtNachname.Text
                $result.DisplayName = "EXTERN | " + $displayNameTemplate
            }
            else {
                $result.DisplayName = $txtDisplayName.Text
            }
            $result.ADGroupsSelected = @()
        }
        else {
            if ([string]::IsNullOrWhiteSpace($txtDisplayName.Text)) {
                $template = $cmbDisplayNameTemplate.SelectedItem
                $displayNameTemplate = $template -replace "\{first\}", $txtVorname.Text -replace "\{last\}", $txtNachname.Text
                $result.DisplayName = $displayNameTemplate
            }
            else {
                $result.DisplayName = $txtDisplayName.Text
            }
            $groupSel = @()
            foreach ($key in $adGroupChecks.Keys) {
                if ($adGroupChecks[$key].Checked) { $groupSel += $key }
            }
            $result.ADGroupsSelected = $groupSel
        }
        $result.Extern = $chkExternal.Checked

        $result.PasswordNeverExpires = $chkPWNeverExpires.Checked
        $result.MustChangePassword   = $chkMustChange.Checked
        $result.AccountDisabled      = $chkAccountDisabled.Checked
        $result.CannotChangePassword = $chkCannotChangePW.Checked
        $result.SmartcardLogonRequired = $chkSmartcardLogonRequired.Checked

        $result.OutputHTML = $chkHTML_Left.Checked
        $result.OutputPDF  = $chkPDF_Left.Checked
        $result.OutputTXT  = $chkTXT_Left.Checked

        $result.UPNEntered = $txtUPN.Text.Trim()
        $result.UPNFormat  = $cmbUPNFormat.SelectedItem

        $result.EmailAddress = $txtEmail.Text.Trim()

        if (-not $result.Vorname) {
            [System.Windows.Forms.MessageBox]::Show("Vorname darf nicht leer sein!","Fehler",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error)
            return
        }
        if (-not $result.Nachname) {
            [System.Windows.Forms.MessageBox]::Show("Nachname darf nicht leer sein!","Fehler",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error)
            return
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
# 3) Hauptablauf: INI laden, GUI anzeigen, Werte verarbeiten
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

# Übergabe der GUI-Ergebnisse in Variablen
$Vorname             = $userSelection.Vorname
$Nachname            = $userSelection.Nachname
$DisplayName         = $userSelection.DisplayName
$Description         = $userSelection.Description
$OfficeRoom          = $userSelection.OfficeRoom
$PhoneNumber         = $userSelection.PhoneNumber
$MobileNumber        = $userSelection.MobileNumber
$Position            = $userSelection.Position
$DepartmentField     = $userSelection.DepartmentField
$CustomAttribute1    = $userSelection.CustomAttribute1
$CustomAttribute2    = $userSelection.CustomAttribute2
$CustomAttribute3    = $userSelection.CustomAttribute3
$CustomAttribute4    = $userSelection.CustomAttribute4
$CustomAttribute5    = $userSelection.CustomAttribute5
$Location            = $userSelection.Location
$CompanySelection    = $userSelection.CompanySelection
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

Write-Host "`nStarte Onboarding für: $Vorname $Nachname"

###############################################################################
# 4) Domain-/Branding-/Logging-Daten aus INI
###############################################################################
if (-not $Company -or $Company -eq "") { $Company = $Config.General.DomainName1 }
if (-not $Config.Contains($Company)) { Throw "Fehler: Die Sektion '$Company' existiert nicht in der INI!" }
$companyData = $Config[$Company]
$Strasse     = $companyData.Strasse
$PLZ         = $companyData.PLZ
$Ort         = $companyData.Ort
if ($Config.Contains("MailEndungen")) {
    $mailSuffix = $cmbMailSuffix.SelectedItem
}
if ($companyData.Contains("Country")) { 
    $Country = $companyData["Country"]
} else { 
    $Country = "DE" 
}
if ($companyData.Contains("Bezeichnung")) {
    $companyDisplay = $companyData["Bezeichnung"]
} else {
    $companyDisplay = $Company
}
$defaultOU   = $Config.General["DefaultOU"]
$logFilePath = $Config.General["LogFilePath"]
$reportPath  = $Config.General["ReportPath"]
$reportTitle = $Config.General["ReportTitle"]
$reportFooter = $Config.General["ReportFooter"]

$firmaLogoPath = $Config.Branding["FirmenLogo"]
# Für Reports nutzen wir das Report-Branding:
$reportBranding = if ($Config.Contains("Branding-Report")) { $Config["Branding-Report"] } else { @{} }
$headerText = if ($reportBranding.ReportHeader) { $reportBranding.ReportHeader } else { $Config.Branding["Header"] }
$footerText = if ($reportBranding.ReportFooter) { $reportBranding.ReportFooter } else { $Config.Branding["Footer"] }

$employeeLinks = @()
if ($Config.Contains("Websites")) {
    foreach ($key in $Config.Websites.Keys) {
        if ($key -match '^EmployeeLink\d+$') { $employeeLinks += $Config.Websites[$key] }
    }
}

if ($Standort -and $Config.STANDORTE.ContainsKey("${Standort}_Bez")) {
    $standortDisplay = $Config.STANDORTE["${Standort}_Bez"]
} else {
    $standortDisplay = $Standort
}

$adSyncEnabled = $Config.ActivateUserMS365ADSync["ADSync"]
$adSyncGroup   = $Config.ActivateUserMS365ADSync["ADSyncADGroup"]
if ($adSyncEnabled -eq '1') { $adSyncEnabled = $true } else { $adSyncEnabled = $false }

###############################################################################
# 5) AD-Benutzer anlegen / aktualisieren
###############################################################################
try { Import-Module ActiveDirectory -ErrorAction Stop } catch { Write-Warning "AD-Modul konnte nicht geladen werden: $_"; return }
if ([string]::IsNullOrWhiteSpace($Vorname)) { Throw "Vorname muss eingegeben werden!" }
function Generate-RandomPassword {
    param(
        [int]$Length,
        [bool]$IncludeSpecial,
        [bool]$AvoidAmbiguous
    )
    $minNonAlpha = 2
    if ($Config.PasswordFixGenerate.ContainsKey("MinNonAlpha")) {
        $minNonAlpha = [int]$Config.PasswordFixGenerate.MinNonAlpha
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
if ($UPNManual) { $UPN = $UPNManual } else {
    switch -Wildcard ($UPNTemplate) {
        "VORNAME.NACHNAME"    { $UPN = "$Vorname.$Nachname$mailSuffix" }
        "V.NACHNAME"          { $UPN = "$($Vorname.Substring(0,1)).$Nachname$mailSuffix" }
        "VORNAMENACHNAME"     { $UPN = "$Vorname$Nachname$mailSuffix" }
        "VNACHNAME"           { $UPN = "$($Vorname.Substring(0,1))$Nachname$mailSuffix" }
        Default               { $UPN = "$SamAccountName$mailSuffix" }
    }
}
if ([string]::IsNullOrWhiteSpace($DisplayName)) { $DisplayName = "$Vorname $Nachname" }
Write-Host "SamAccountName : $SamAccountName"
Write-Host "UPN            : $UPN"
Write-Host "Passwort       : $UserPW"
try { $existingUser = Get-ADUser -Filter { SamAccountName -eq $SamAccountName } -ErrorAction SilentlyContinue } catch { $existingUser = $null }
if (-not $existingUser) {
    Write-Host "Erstelle neuen Benutzer: $DisplayName"
    try {
        New-ADUser `
            -Name $DisplayName `
            -GivenName $Vorname `
            -Surname $Nachname `
            -SamAccountName $SamAccountName `
            -UserPrincipalName $UPN `
            -AccountPassword $SecurePW `
            -Enabled (-not $result.AccountDisabled) `
            -ChangePasswordAtLogon $result.MustChangePassword `
            -PasswordNeverExpires $result.PasswordNeverExpires `
            -Path $defaultOU `
            -City $Ort `
            -StreetAddress $Strasse `
            -Country $Country `
            -OtherAttributes @{ mail = $EmailAddress; description = $Description; physicalDeliveryOfficeName = $OfficeRoom; telephoneNumber = $PhoneNumber; mobile = $MobileNumber; title = $Position; department = $DepartmentField } `
            -ErrorAction Stop
        Write-Host "AD-Benutzer erstellt."
    } catch { Write-Warning "Fehler beim Erstellen des Benutzers: $_"; return }
    
    if ($result.SmartcardLogonRequired) {
        try { Set-ADUser -Identity $SamAccountName -SmartcardLogonRequired $true -ErrorAction Stop } catch { Write-Warning "Fehler bei SmartcardLogonRequired: $_" }
    }
    if ($result.CannotChangePassword) {
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
            -Enabled (-not $result.AccountDisabled) `
            -ErrorAction SilentlyContinue
        Set-ADUser -Identity $existingUser.DistinguishedName -ChangePasswordAtLogon:$result.MustChangePassword -PasswordNeverExpires:$result.PasswordNeverExpires
    } catch { Write-Warning "Fehler beim Aktualisieren: $_" }
}
try { Set-ADAccountPassword -Identity $SamAccountName -Reset -NewPassword $SecurePW -ErrorAction SilentlyContinue } catch { Write-Warning "Fehler beim Setzen des Passworts: $_" }
if ($cannotChangePW) { Write-Host "(Hinweis: 'CannotChangePassword' via ACL wäre hier umzusetzen.)" }

###############################################################################
# 6) AD-Gruppen zuweisen
###############################################################################
if (-not $userSelection.Extern) {
    foreach ($groupKey in $userSelection.ADGroupsSelected) {
        $groupName = $Config.ADGroups[$groupKey]
        if ($groupName) {
            try { Add-ADGroupMember -Identity $groupName -Members $SamAccountName -ErrorAction Stop } catch { Write-Warning "Fehler bei AD-Gruppe '$groupName': $_" }
        }
    }
} else { Write-Host "Externer Mitarbeiter: Standardmäßige AD-Gruppen-Zuweisung wird übersprungen." }
if ($Standort) {
    $signaturKey = $Config.STANDORTE[$Standort]
    if ($signaturKey) {
        $signaturGroup = $Config.SignaturGruppe_Optional[$signaturKey]
        if ($signaturGroup) {
            try { Add-ADGroupMember -Identity $signaturGroup -Members $SamAccountName -ErrorAction SilentlyContinue } catch { Write-Warning "Fehler bei Signatur-Gruppe: $_" }
        }
    }
}
if ($License) {
    $licenseKey = "MS365_" + $License
    $licenseGroup = $Config.LicensesGroups[$licenseKey]
    if ($licenseGroup) {
        try { Add-ADGroupMember -Identity $licenseGroup -Members $SamAccountName -ErrorAction SilentlyContinue } catch { Write-Warning "Fehler bei Lizenz-Gruppe: $_" }
    }
}
if ($adSyncEnabled -and $adSyncGroup) {
    try { Add-ADGroupMember -Identity $adSyncGroup -Members $SamAccountName -ErrorAction SilentlyContinue } catch { Write-Warning "Fehler bei ADSync-Gruppe: $_" }
}
if ($userSelection.Extern) { Write-Host "Externer Mitarbeiter: Bitte weisen Sie alle AD-Gruppen händisch zu." }

###############################################################################
# 7) Logging
###############################################################################
try {
    if (-not (Test-Path $logFilePath)) { New-Item -ItemType Directory -Path $logFilePath -Force | Out-Null }
    $logDate = (Get-Date -Format 'yyyyMMdd')
    $logFile = Join-Path $logFilePath "Onboarding_$logDate.log"
    $logEntry = "[{0}] Sam={1}, Anzeigename='{2}', UPN='{3}', Standort='{4}', Company='{5}', Location='{6}', MS365 Lizenz='{7}', ADGruppen=({8}), Passwort='{9}', Extern={10}" -f (Get-Date), $SamAccountName, $DisplayName, $UPN, $standortDisplay, $companyDisplay, $standortDisplay, $MS365License, ($userSelection.ADGroupsSelected -join ','), $UserPW, $Extern
    Add-Content -Path $logFile -Value $logEntry
    Write-Host "Log geschrieben: $logFile"
} catch { Write-Warning "Fehler beim Logging: $_" }

###############################################################################
# 8) Reports erzeugen (HTML, PDF, TXT)
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
            CustomAttrib1       = $CustomAttribute1
            CustomAttrib2       = $CustomAttribute2
            CustomAttrib3       = $CustomAttribute3
            CustomAttrib4       = $CustomAttribute4
            CustomAttrib5       = $CustomAttribute5
            Location            = $standortDisplay
            Company             = $companyDisplay
            MS365Lizenz         = $cmbMS365License.SelectedItem
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
  <h1>$($reportBranding.ReportHeader -or $headerText)</h1>
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
  <footer><p>$($reportBranding.ReportFooter -or $footerText)</p></footer>
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
catch { Write-Warning "Fehler beim Erstellen der Reports: $_" }

###############################################################################
# 9) Ausgabe der neuen INI-Konfiguration (als Text)
###############################################################################
Write-Host "`nNeue INI-Konfiguration:"
foreach ($section in $Config.Keys) {
    Write-Host "[$section]"
    foreach ($key in $Config[$section].Keys) { Write-Host "$key = $($Config[$section][$key])" }
    Write-Host ""
}
Write-Host "`nOnboarding abgeschlossen."

# SIG # Begin signature block
# MIIbywYJKoZIhvcNAQcCoIIbvDCCG7gCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCALe5cbhonuRD+K
# HDxsiQViCtSfbR9QRF2znAQJA5FngaCCFhcwggMQMIIB+KADAgECAhB3jzsyX9Cg
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
# DQEJBDEiBCCpp4N7/3fSrgPLB1LS1bv/thd/oGtuu6Jffb+HWTyyhDANBgkqhkiG
# 9w0BAQEFAASCAQCb4S1XfqVgdYOPBjMdxyRdyii0wfS5Zgbz3CWKcqfAX2Z6VJMt
# 65dgrHbLEuv0mnHWqzcLx5Cnmx5jDPz/jPt7jr/VLIRRiv+mZ0MNTyMsvmw9XgpK
# TF1Ioci5yHPwWDg/YVxHRO3ChKco3VefP/c8DJ62dLogvrq8w614Rch4QOYozVCY
# Tn28seyvgEImGN3lWNKwsrVB/FYDaRzgpBVhOk+rCsBaXhTn03cRLnR/rTOQnsOB
# 16khmoS9Cv1udCOHwQLQbm1mcDFbE+IdgbHHuA3fXHDGldTDXy/3psTA7osIOSNn
# UNpx15wL2vt5pemO45I7k5wQMd6OkvV+LkSJoYIDIDCCAxwGCSqGSIb3DQEJBjGC
# Aw0wggMJAgEBMHcwYzELMAkGA1UEBhMCVVMxFzAVBgNVBAoTDkRpZ2lDZXJ0LCBJ
# bmMuMTswOQYDVQQDEzJEaWdpQ2VydCBUcnVzdGVkIEc0IFJTQTQwOTYgU0hBMjU2
# IFRpbWVTdGFtcGluZyBDQQIQC65mvFq6f5WHxvnpBOMzBDANBglghkgBZQMEAgEF
# AKBpMBgGCSqGSIb3DQEJAzELBgkqhkiG9w0BBwEwHAYJKoZIhvcNAQkFMQ8XDTI1
# MDcwNTEwMTE1NlowLwYJKoZIhvcNAQkEMSIEIH35Q0QhZxDYWFJ3JL14xCwJbUm6
# ufvGzG9+eENoG24rMA0GCSqGSIb3DQEBAQUABIICAER4BRjw6jJZHnJ1/1SSKuK/
# rgrfEF1Y5o9v6XOPbSDVnjrXlwy1uqpjYrcgTEM2+29m+toCqZEvLO9+nnKpoxk1
# LHv8ARw7ZpeMOEWTy5FKK5S7OoZsyQ6StjmyDv67BNVLfey9H3r1i+bv6Oj11y+w
# tTmH4HK1fWVoiv1Bz0WyxZYYlHRIY3FzKrRoQK3F9zjwnmmccRoqO1MRoP1P/AAU
# ly0hJUG/RX6cN/g5tYbMii2Bux/+1R88HbswgQqJ93rjRRUSODp5EXtNk5NbrNxa
# 9XHe+a+lb0I+TAVNzaVdv/UdbDuadgyTf5hgrq6YwRzpeelr4SgZ9ma9dsa53204
# hzQb4CaQFT1Lh6vhUv6mPfJLfqgrNcDLRAeQgM2nYP3+fI6xGPsSj5ZTqMrvD9Om
# A43TaTqk9kjW6JpROZfHJ3Ep3A1uAdwvXJlMI1cEQEEVwtlGyAdd7sF5aLFSNXnk
# QYYwu44gdoXobrwSEMQZ85SggSSQT1hJ6vf9V2jXXmREk0esRUbv5j+xdIH4XZ6J
# 08skDdOE0ZxDO/cyo1S+rXVEVKguRCBaPZKt0j77NbESDKbsOO0D9aqwvM9uNVHF
# 0OUwiG4Az+Qh4HU6oen34Urbpy4kPhKeSfF+oPfQPLDe0dqOzfD8zjGIQBYr0wMC
# GlVfwj+0fnEhzlQAri72
# SIG # End signature block
