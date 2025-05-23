#requires -Version 5.1
[CmdletBinding()]
param(
    [string]$Vorname,
    [string]$Nachname,
    [string]$Standort,
    [string]$Company,
    [string]$License = "",
    [switch]$Extern,
    [string]$ScriptINIPath = "easyOnboarding_V0.1.2_Config.ini"
)

# Für Passwort-Generierung (Membership.GeneratePassword)
Add-Type -AssemblyName System.Web

###############################################################################
# 1) INI-Datei einlesen (OrderedDictionary – Verwende .Contains statt .ContainsKey)
###############################################################################
function Read-INIFile {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Path
    )
    if (-not (Test-Path $Path)) {
        Throw "INI-Datei nicht gefunden: $Path"
    }
    $iniContent = Get-Content -Path $Path | Where-Object {
        $_ -notmatch '^\s*[;#]' -and $_.Trim() -ne ""
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

###############################################################################
# 2) GUI-Erstellung (Panels: Links, Rechts, Unten)
###############################################################################
function Show-OnboardingForm {
    param(
        [hashtable]$INIConfig
    )

    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    # Hauptfenster
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "easyOnboarding - GUI"
    $form.StartPosition = "CenterScreen"
    $form.Size = New-Object System.Drawing.Size(1100,750)
    $form.AutoScroll = $true

    # Info-Label oben
    $lblInfo = New-Object System.Windows.Forms.Label
    $lblInfo.AutoSize = $true
    $lblInfo.Font     = 'Microsoft Sans Serif,10,style=Bold'
    $lblInfo.Location = '10,10'
    $lblInfo.Size     = New-Object System.Drawing.Size(1000,40)
    $form.Controls.Add($lblInfo)

    # Panel links (manuelle Eingaben + Location, Company)
    $panelLeft = New-Object System.Windows.Forms.Panel
    $panelLeft.Location = New-Object System.Drawing.Point(10,60)
    $panelLeft.Size     = New-Object System.Drawing.Size(520,600)
    $panelLeft.AutoScroll = $true
    $panelLeft.BorderStyle = 'FixedSingle'
    $form.Controls.Add($panelLeft)

    # Panel rechts (AD-Gruppen, MS365 Lizenz, AD-Flags, PW-Optionen, UPN)
    $panelRight = New-Object System.Windows.Forms.Panel
    $panelRight.Location = New-Object System.Drawing.Point(550,60)
    $panelRight.Size     = New-Object System.Drawing.Size(520,600)
    $panelRight.AutoScroll = $true
    $panelRight.BorderStyle = 'FixedSingle'
    $form.Controls.Add($panelRight)

    # Panel unten für Buttons (angedockt)
    $panelBottom = New-Object System.Windows.Forms.Panel
    $panelBottom.Dock = 'Bottom'
    $panelBottom.Height = 60
    $panelBottom.BorderStyle = 'None'
    $form.Controls.Add($panelBottom)

    ############################################################################
    # Hilfsfunktionen
    ############################################################################
    function AddLabel($parent, [string]$text, [int]$x, [int]$y) {
        $lbl = New-Object System.Windows.Forms.Label
        $lbl.Text = $text
        $lbl.Location = New-Object System.Drawing.Point($x, $y)
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
    # ScriptInfo / General für Info-Label
    ############################################################################
    $ScriptVersion = ""
    $LastUpdate    = ""
    $Author        = ""
    $domainName1   = ""
    $defaultOU     = ""
    $reportPath    = ""
    if ($INIConfig.Contains('ScriptInfo')) {
        $ScriptVersion = $INIConfig['ScriptInfo'].ScriptVersion
        $LastUpdate    = $INIConfig['ScriptInfo'].LastUpdate
        $Author        = $INIConfig['ScriptInfo'].Author
    }
    if ($INIConfig.Contains('General')) {
        $domainName1 = $INIConfig['General'].DomainName1
        $defaultOU   = $INIConfig['General'].DefaultOU
        $reportPath  = $INIConfig['General'].ReportPath
    }
    $lblInfo.Text = "ScriptVersion=$ScriptVersion | LastUpdate=$LastUpdate | Author=$Author`r`nDOMAIN: $domainName1 | OU: $defaultOU | REPORT: $reportPath"

    ############################################################################
    # PanelLeft: Manuelle Eingaben + Location und Company (Pflichtfelder)
    ############################################################################
    $yLeft = 10
    AddLabel $panelLeft "Vorname:" 10 $yLeft | Out-Null
    $txtVorname = AddTextBox $panelLeft "" 150 $yLeft; $yLeft += 30

    AddLabel $panelLeft "Nachname:" 10 $yLeft | Out-Null
    $txtNachname = AddTextBox $panelLeft "" 150 $yLeft; $yLeft += 30

    AddLabel $panelLeft "Anzeigename:" 10 $yLeft | Out-Null
    $txtDisplayName = AddTextBox $panelLeft "" 150 $yLeft; $yLeft += 30

    AddLabel $panelLeft "Beschreibung:" 10 $yLeft | Out-Null
    $txtDescription = AddTextBox $panelLeft "" 150 $yLeft; $yLeft += 30

    AddLabel $panelLeft "Büro:" 10 $yLeft | Out-Null
    $txtOffice = AddTextBox $panelLeft "" 150 $yLeft; $yLeft += 30

    AddLabel $panelLeft "Rufnummer:" 10 $yLeft | Out-Null
    $txtPhone = AddTextBox $panelLeft "" 150 $yLeft; $yLeft += 30

    AddLabel $panelLeft "Mobil:" 10 $yLeft | Out-Null
    $txtMobile = AddTextBox $panelLeft "" 150 $yLeft; $yLeft += 30

    AddLabel $panelLeft "Position:" 10 $yLeft | Out-Null
    $txtPosition = AddTextBox $panelLeft "" 150 $yLeft; $yLeft += 30

    AddLabel $panelLeft "Abteilung (manuell):" 10 $yLeft | Out-Null
    $txtDeptField = AddTextBox $panelLeft "" 150 $yLeft; $yLeft += 30

    AddLabel $panelLeft "CustomAttrib1:" 10 $yLeft | Out-Null
    $txtCA1 = AddTextBox $panelLeft "" 150 $yLeft; $yLeft += 30

    AddLabel $panelLeft "CustomAttrib2:" 10 $yLeft | Out-Null
    $txtCA2 = AddTextBox $panelLeft "" 150 $yLeft; $yLeft += 30

    AddLabel $panelLeft "CustomAttrib3:" 10 $yLeft | Out-Null
    $txtCA3 = AddTextBox $panelLeft "" 150 $yLeft; $yLeft += 30

    AddLabel $panelLeft "CustomAttrib4:" 10 $yLeft | Out-Null
    $txtCA4 = AddTextBox $panelLeft "" 150 $yLeft; $yLeft += 30

    AddLabel $panelLeft "CustomAttrib5:" 10 $yLeft | Out-Null
    $txtCA5 = AddTextBox $panelLeft "" 150 $yLeft; $yLeft += 30

    # DropDown: Location (Pflichtfeld)
    AddLabel $panelLeft "Location*:" 10 $yLeft | Out-Null
    # Alle Keys aus [STANDORTE] außer "DefaultStandort" als Auswahlmöglichkeiten
    $cmbLocation = AddComboBox $panelLeft (($INIConfig.STANDORTE.Keys | Where-Object { $_ -ne "DefaultStandort" })) 150 $yLeft 250 ""
    $yLeft += 30

    # DropDown: Company (Pflichtfeld)
    AddLabel $panelLeft "Company*:" 10 $yLeft | Out-Null
    $cmbCompany = AddComboBox $panelLeft (($INIConfig.Keys | Where-Object { $_ -like "DomainName*" })) 150 $yLeft 250 ""
    $yLeft += 30

    ############################################################################
    # PanelRight: AD-Gruppen, MS365 Lizenz, AD-Flags, Passwortoptionen, UPN, Output
    ############################################################################
    $yRight = 10

    # DropDown: MS365 Lizenz (Pflichtfeld)
    AddLabel $panelRight "MS365 Lizenz*:" 10 $yRight | Out-Null
    $cmbMS365License = AddComboBox $panelRight (($INIConfig.LicensesGroups.Keys | ForEach-Object { $_ -replace '^MS365_','' })) 150 $yRight 200 ""
    $yRight += 35

    # AD-Gruppen (Checkboxen)
    AddLabel $panelRight "AD-Gruppen:" 10 $yRight | Out-Null
    $yRight += 25

    $adGroupChecks = @{}
    if ($INIConfig.Contains('ADGroups')) {
        $adGroupKeys = $INIConfig.ADGroups.Keys | Where-Object { $_ -notmatch '^(DefaultADGroup|.*_(Visible|Label))$' }
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
                $cbGroup = AddCheckBox $panelRight $displayText $false 10 $yRight
                $adGroupChecks[$g] = $cbGroup
                $yRight += 25
            }
        }
    }
    else {
        AddLabel $panelRight "Keine [ADGroups] Sektion gefunden." 10 $yRight | Out-Null
        $yRight += 25
    }

    # Externer Mitarbeiter
    $chkExtern = AddCheckBox $panelRight "Externer Mitarbeiter?" $false 10 $yRight
    $yRight += 35

    # AD-Benutzer-Flags
    AddLabel $panelRight "AD-Benutzer-Flags:" 10 $yRight | Out-Null
    $yRight += 20

    $chkPWNeverExpires = AddCheckBox $panelRight "PasswordNeverExpires" $false 10 $yRight
    $chkMustChange     = AddCheckBox $panelRight "MustChangePasswordAtLogon" $false 250 $yRight
    $yRight += 25

    $chkAccountDisabled = AddCheckBox $panelRight "AccountDisabled" $false 10 $yRight
    $chkCannotChangePW  = AddCheckBox $panelRight "CannotChangePassword" $false 250 $yRight
    $yRight += 35

    # Passwortoptionen
    AddLabel $panelRight "PASSWORT-OPTIONEN:" 10 $yRight | Out-Null
    $yRight += 25

    $rbFix = New-Object System.Windows.Forms.RadioButton
    $rbFix.Text = "Fixes Passwort"
    $rbFix.Location = New-Object System.Drawing.Point(10, $yRight)
    $panelRight.Controls.Add($rbFix)

    $rbRand = New-Object System.Windows.Forms.RadioButton
    $rbRand.Text = "Generiertes Passwort"
    $rbRand.Location = New-Object System.Drawing.Point(150, $yRight)
    $panelRight.Controls.Add($rbRand)
    $yRight += 25

    AddLabel $panelRight "FixPasswort:" 10 $yRight | Out-Null
    $txtFixPW = AddTextBox $panelRight "" 130 $yRight 150
    $yRight += 30

    AddLabel $panelRight "Passwortlänge:" 10 $yRight | Out-Null
    $txtPWLen = AddTextBox $panelRight "12" 130 $yRight 50
    $yRight += 30

    $chkIncludeSpecial = AddCheckBox $panelRight "IncludeSpecialChars" $true 10 $yRight
    $chkAvoidAmbig     = AddCheckBox $panelRight "AvoidAmbiguousChars" $true 250 $yRight
    $yRight += 35

    # Output-Optionen
    AddLabel $panelRight "OUTPUT - ONBOARDING?" 10 $yRight | Out-Null
    $yRight += 20

    $chkHTML = AddCheckBox $panelRight "HTML erzeugen" $true 10 $yRight
    $chkPDF  = AddCheckBox $panelRight "PDF erzeugen"  $true 150 $yRight
    $chkTXT  = AddCheckBox $panelRight "TXT erzeugen"  $true 290 $yRight
    $yRight += 35

    # UPN
    AddLabel $panelRight "BenutzerPrincipalName (UPN):" 10 $yRight | Out-Null
    $txtUPN = AddTextBox $panelRight "" 200 $yRight 200
    $yRight += 30

    AddLabel $panelRight "UPN-Format-Vorlagen:" 10 $yRight | Out-Null
    $cmbUPNFormat = AddComboBox $panelRight @("VORNAME.NACHNAME","V.NACHNAME","VORNAMENACHNAME","VNACHNAME") 200 $yRight 200
    $yRight += 40

    ############################################################################
    # PanelBottom: Buttons
    ############################################################################
    $btnOK = New-Object System.Windows.Forms.Button
    $btnOK.Text = "Onboarding Starten"
    $btnOK.Size = New-Object System.Drawing.Size(150,30)
    $btnOK.Location = New-Object System.Drawing.Point(400, 10)
    $panelBottom.Controls.Add($btnOK)

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "Abbrechen"
    $btnCancel.Size = New-Object System.Drawing.Size(150,30)
    $btnCancel.Location = New-Object System.Drawing.Point(600, 10)
    $panelBottom.Controls.Add($btnCancel)

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
        Location              = ""   # Neue Eigenschaft
        CompanySelection      = ""   # Neue Eigenschaft
        MS365License          = ""   # Neue Eigenschaft (aus dem DropDown in PanelRight)
        PasswordNeverExpires  = $false
        MustChangePassword    = $false
        AccountDisabled       = $false
        CannotChangePassword  = $false
        PasswordMode          = 1   # 0=fix, 1=generiert
        FixPassword           = ""
        PasswordLaenge        = 12
        IncludeSpecialChars   = $true
        AvoidAmbiguousChars   = $true
        OutputHTML            = $true
        OutputPDF             = $true
        OutputTXT             = $true
        UPNEntered            = ""
        UPNFormat             = "VORNAME.NACHNAME"
        Cancel                = $false
    }

    # Umschalten fix/random
    function UpdatePWFields {
        if ($rbFix.Checked) {
            $txtFixPW.Enabled         = $true
            $txtPWLen.Enabled         = $false
            $chkIncludeSpecial.Enabled= $false
            $chkAvoidAmbig.Enabled    = $false
        } else {
            $txtFixPW.Enabled         = $false
            $txtPWLen.Enabled         = $true
            $chkIncludeSpecial.Enabled= $true
            $chkAvoidAmbig.Enabled    = $true
        }
    }
    $rbFix.Add_CheckedChanged({ UpdatePWFields })
    $rbRand.Add_CheckedChanged({ UpdatePWFields })
    $rbRand.Checked = $true
    UpdatePWFields

    # Klick "Onboarding Starten"
    $btnOK.Add_Click({
        # Linke Felder
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

        # Neue DropDowns in PanelLeft
        $result.Location = $cmbLocation.SelectedItem
        $result.CompanySelection = $cmbCompany.SelectedItem

        # DropDown in PanelRight für MS365 Lizenz
        $result.MS365License = $cmbMS365License.SelectedItem

        # Rechte Felder
        $result.PasswordNeverExpires = $chkPWNeverExpires.Checked
        $result.MustChangePassword   = $chkMustChange.Checked
        $result.AccountDisabled      = $chkAccountDisabled.Checked
        $result.CannotChangePassword = $chkCannotChangePW.Checked

        if ($rbFix.Checked) { $result.PasswordMode = 0 } else { $result.PasswordMode = 1 }
        $result.FixPassword         = $txtFixPW.Text
        $result.PasswordLaenge      = [int]$txtPWLen.Text
        $result.IncludeSpecialChars = $chkIncludeSpecial.Checked
        $result.AvoidAmbiguousChars = $chkAvoidAmbig.Checked

        $result.OutputHTML = $chkHTML.Checked
        $result.OutputPDF  = $chkPDF.Checked
        $result.OutputTXT  = $chkTXT.Checked

        $result.UPNEntered = $txtUPN.Text.Trim()
        $result.UPNFormat  = $cmbUPNFormat.SelectedItem

        # Minimale Validierung
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

    # Klick "Abbrechen"
    $btnCancel.Add_Click({
        $result.Cancel = $true
        $form.Close()
    })

    # Aktualisiere Info-Label
    $ScriptVersion = ""
    $LastUpdate    = ""
    $Author        = ""
    if ($INIConfig.Contains('ScriptInfo')) {
        $ScriptVersion = $INIConfig['ScriptInfo'].ScriptVersion
        $LastUpdate    = $INIConfig['ScriptInfo'].LastUpdate
        $Author        = $INIConfig['ScriptInfo'].Author
    }
    if ($INIConfig.Contains('General')) {
        $domainName1 = $INIConfig['General'].DomainName1
        $defaultOU   = $INIConfig['General'].DefaultOU
        $reportPath  = $INIConfig['General'].ReportPath
    }
    $lblInfo.Text = "ScriptVersion=$ScriptVersion | LastUpdate=$LastUpdate | Author=$Author`r`nDOMAIN: $domainName1 | OU: $defaultOU | REPORT: $reportPath"

    $null = $form.ShowDialog()
    return $result
}

###############################################################################
# 3) Hauptablauf: INI laden, GUI anzeigen, Werte verarbeiten
###############################################################################
Write-Host "Lade INI: $ScriptINIPath"
$Config = Read-INIFile $ScriptINIPath

$userSelection = Show-OnboardingForm -INIConfig $Config
if ($userSelection.Cancel) {
    Write-Warning "Onboarding abgebrochen."
    return
}

# Übernehme die GUI-Ergebnisse
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

Write-Host "`nStarte Onboarding für: $Vorname $Nachname"

###############################################################################
# 4) Domain-/Branding-/Logging-Daten aus INI
###############################################################################
# Falls $Company leer ist, Standard: "DomainName1" verwenden
if (-not $Company -or $Company -eq "") {
    $Company = "DomainName1"
}

if (-not $Config.Contains($Company)) {
    Throw "Fehler: Die Sektion '$Company' existiert nicht in der INI!"
}
$companyData = $Config[$Company]
$Strasse     = $companyData.Strasse
$PLZ         = $companyData.PLZ
$Ort         = $companyData.Ort
$Mailendung  = $companyData.Mailendung
if ($companyData.Contains("Country")) {
    $Country = $companyData["Country"]
} else {
    $Country = "DE"
}

$defaultOU    = $Config.General["DefaultOU"]
$logFilePath  = $Config.General["LogFilePath"]
$reportPath   = $Config.General["ReportPath"]
$reportTitle  = $Config.General["ReportTitle"]
$reportFooter = $Config.General["ReportFooter"]

$firmaLogoPath = $Config.Branding["FirmaLogo"]
$headerText    = $Config.Branding["Header"]
$footerText    = $Config.Branding["Footer"]

# Websites
$employeeLinks = @()
if ($Config.Contains("Websites")) {
    foreach ($key in $Config.Websites.Keys) {
        if ($key -match '^EmployeeLink\d+$') {
            $employeeLinks += $Config.Websites[$key]
        }
    }
}

# ADSync-Einstellungen
$adSyncEnabled = $Config.ActivateUserMS365ADSync["ADSync"]
$adSyncGroup   = $Config.ActivateUserMS365ADSync["ADSyncADGroup"]
if ($adSyncEnabled -eq '1') { $adSyncEnabled = $true } else { $adSyncEnabled = $false }

###############################################################################
# 5) AD-Benutzer anlegen / aktualisieren
###############################################################################
try {
    Import-Module ActiveDirectory -ErrorAction Stop
} catch {
    Write-Warning "AD-Modul konnte nicht geladen werden: $_"
    return
}

if ([string]::IsNullOrWhiteSpace($Vorname)) {
    Throw "Vorname muss eingegeben werden!"
}

function Generate-RandomPassword {
    param(
        [int]$Length,
        [bool]$IncludeSpecial,
        [bool]$AvoidAmbiguous
    )
    $minNonAlpha = 2
    $pw = [System.Web.Security.Membership]::GeneratePassword($Length, $minNonAlpha)
    if ($AvoidAmbiguous) {
        $pw = $pw -replace '[{}()\[\]\/\\`~,;:.<>\"]','X'
    }
    return $pw
}

if ($passwordMode -eq 1) {
    $UserPW = Generate-RandomPassword -Length $passwordLaenge -IncludeSpecial $includeSpecial -AvoidAmbiguous $avoidAmbiguous
    if ([string]::IsNullOrWhiteSpace($UserPW)) { $UserPW = "Standard123!" }
} else {
    $UserPW = $fixPassword
}
$SecurePW = ConvertTo-SecureString $UserPW -AsPlainText -Force

$SamAccountName = ($Vorname.Substring(0,1) + $Nachname).ToLower()

if ($UPNManual) {
    $UPN = $UPNManual
} else {
    switch -Wildcard ($UPNTemplate) {
        "VORNAME.NACHNAME"    { $UPN = "$Vorname.$Nachname$Mailendung" }
        "V.NACHNAME"          { $UPN = "$($Vorname.Substring(0,1)).$Nachname$Mailendung" }
        "VORNAMENACHNAME"     { $UPN = "$Vorname$Nachname$Mailendung" }
        "VNACHNAME"           { $UPN = "$($Vorname.Substring(0,1))$Nachname$Mailendung" }
        Default               { $UPN = "$SamAccountName$Mailendung" }
    }
}
if ([string]::IsNullOrWhiteSpace($DisplayName)) {
    $DisplayName = "$Vorname $Nachname"
}

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
    try {
        New-ADUser `
            -Name $DisplayName `
            -GivenName $Vorname `
            -Surname $Nachname `
            -SamAccountName $SamAccountName `
            -UserPrincipalName $UPN `
            -AccountPassword $SecurePW `
            -Enabled ($accountDisabled -notlike 'True') `
            -ChangePasswordAtLogon $mustChangePW `
            -PasswordNeverExpires $passwordNeverExpires `
            -Path $defaultOU `
            -City $Ort `
            -StreetAddress $Strasse `
            -Country $Country `
            -ErrorAction Stop
        Write-Host "AD-Benutzer erstellt."
    } catch {
        Write-Warning "Fehler beim Erstellen des Benutzers: $_"
        return
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
            -Enabled ($accountDisabled -notlike 'True') `
            -ErrorAction SilentlyContinue
        Set-ADUser -Identity $existingUser.DistinguishedName -ChangePasswordAtLogon:$mustChangePW -PasswordNeverExpires:$passwordNeverExpires
    } catch {
        Write-Warning "Fehler beim Aktualisieren: $_"
    }
}

try {
    Set-ADAccountPassword -Identity $SamAccountName -Reset -NewPassword $SecurePW -ErrorAction SilentlyContinue
} catch {
    Write-Warning "Fehler beim Setzen des Passworts: $_"
}

if ($cannotChangePW) {
    Write-Host "(Hinweis: 'CannotChangePassword' via ACL wäre hier umzusetzen.)"
}

###############################################################################
# 6) AD-Gruppen zuweisen
###############################################################################
foreach ($groupKey in $userSelection.ADGroupsSelected) {
    $groupName = $Config.ADGroups[$groupKey]
    if ($groupName) {
        try {
            Add-ADGroupMember -Identity $groupName -Members $SamAccountName -ErrorAction Stop
            Write-Host "AD-Gruppe '$groupName' zugewiesen (Gruppe=$groupKey)."
        } catch {
            Write-Warning "Fehler bei AD-Gruppe '$groupName': $_"
        }
    }
}

if ($Standort) {
    $signaturKey = $Config.STANDORTE[$Standort]
    if ($signaturKey) {
        $signaturGroup = $Config.SignaturGruppe_Optional[$signaturKey]
        if ($signaturGroup) {
            try {
                Add-ADGroupMember -Identity $signaturGroup -Members $SamAccountName -ErrorAction SilentlyContinue
                Write-Host "Signatur-Gruppe '$signaturGroup' zugewiesen."
            } catch {
                Write-Warning "Fehler bei Signatur-Gruppe: $_"
            }
        }
    }
}

if ($License) {
    $licenseKey = "MS365_" + $License
    $licenseGroup = $Config.LicensesGroups[$licenseKey]
    if ($licenseGroup) {
        try {
            Add-ADGroupMember -Identity $licenseGroup -Members $SamAccountName -ErrorAction SilentlyContinue
            Write-Host "Lizenz-Gruppe '$licenseGroup' zugewiesen."
        } catch {
            Write-Warning "Fehler bei Lizenz-Gruppe: $_"
        }
    }
}

if ($adSyncEnabled -and $adSyncGroup) {
    try {
        Add-ADGroupMember -Identity $adSyncGroup -Members $SamAccountName -ErrorAction SilentlyContinue
        Write-Host "ADSync-Gruppe '$adSyncGroup' zugewiesen."
    } catch {
        Write-Warning "Fehler bei ADSync-Gruppe: $_"
    }
}

if ($Extern) {
    Write-Host "Externer Mitarbeiter - ggf. weitere Spezialgruppen/Einschränkungen."
}

###############################################################################
# 7) Logging
###############################################################################
try {
    if (-not (Test-Path $logFilePath)) {
        New-Item -ItemType Directory -Path $logFilePath -Force | Out-Null
    }
    $logDate = (Get-Date -Format 'yyyyMMdd')
    $logFile = Join-Path $logFilePath "Onboarding_$logDate.log"
    $logEntry = "[{0}] Sam={1}, Anzeigename='{2}', UPN='{3}', Standort={4}, Company='{5}', Location='{6}', MS365 Lizenz='{7}', ADGruppen=({8}), Passwort='{9}', Extern={10}" -f (Get-Date), $SamAccountName, $DisplayName, $UPN, $Standort, $CompanySelection, $Location, $MS365License, ($userSelection.ADGroupsSelected -join ','), $UserPW, $Extern
    Add-Content -Path $logFile -Value $logEntry
    Write-Host "Log geschrieben: $logFile"
} catch {
    Write-Warning "Fehler beim Logging: $_"
}

###############################################################################
# 8) Reports erzeugen (HTML, PDF, TXT)
###############################################################################
try {
    if (-not (Test-Path $reportPath)) {
        New-Item -ItemType Directory -Path $reportPath -Force | Out-Null
    }
    if ($createHTML) {
        $htmlFile = Join-Path $reportPath "$SamAccountName.html"
        $logoTag = ""
        if ($firmaLogoPath -and (Test-Path $firmaLogoPath)) {
            $logoTag = "<img src='file:///$firmaLogoPath' style='float:right; max-width:120px; margin:10px;'/>"
        }
        $linksHtml = ""
        if ($employeeLinks.Count -gt 0) {
            $linksHtml += "<h3>Wichtige Links</h3><ul>"
            foreach ($linkLine in $employeeLinks) {
                $parts = $linkLine -split ';'
                if ($parts.Count -ge 2) {
                    $lkName = $parts[0]
                    $lkUrl  = $parts[1]
                    $desc   = ""
                    if ($parts.Count -eq 3) { $desc = " - " + $parts[2] }
                    $linksHtml += "<li><a href='$lkUrl' target='_blank'>$lkName</a>$desc</li>"
                }
            }
            $linksHtml += "</ul>"
        }
        $html = @"
<html>
<head>
  <meta charset='UTF-8'>
  <title>$reportTitle</title>
  <style>
    body { font-family: Arial; margin:20px; }
    .header { overflow:auto; }
    .footer { margin-top:20px; font-size:0.9em; color:#666; }
  </style>
</head>
<body>
<div class="header">
  $logoTag
  <h1>$headerText</h1>
  <h2>$reportTitle</h2>
</div>
<p><b>Benutzer:</b> $DisplayName</p>
<p><b>SamAccountName:</b> $SamAccountName</p>
<p><b>UPN:</b> $UPN</p>
<p><b>AD-Gruppen:</b> $($userSelection.ADGroupsSelected -join ', ')</p>
<p><b>Standort:</b> $Standort</p>
<p><b>Company:</b> $CompanySelection</p>
<p><b>Location:</b> $Location</p>
<p><b>MS365 Lizenz:</b> $MS365License</p>
<p><b>Extern:</b> $Extern</p>
<p><b>Passwort:</b> $UserPW</p>
<p><b>Anzeigename (individuell):</b> $DisplayName</p>
<p><b>Beschreibung:</b> $Description</p>
<p><b>Büro:</b> $OfficeRoom</p>
<p><b>Rufnummer:</b> $PhoneNumber</p>
<p><b>CustomAttrib1:</b> $CustomAttribute1</p>
<p><b>CustomAttrib2:</b> $CustomAttribute2</p>
<p><b>CustomAttrib3:</b> $CustomAttribute3</p>
<p><b>CustomAttrib4:</b> $CustomAttribute4</p>
<p><b>CustomAttrib5:</b> $CustomAttribute5</p>
<p><b>Mobil:</b> $MobileNumber</p>
<p><b>Position:</b> $Position</p>
<p><b>Abteilung (manuell):</b> $DepartmentField</p>
$linksHtml
<div class="footer">
  <p>$reportFooter</p>
  <p>$footerText</p>
  <p>Erstellt am: $(Get-Date)</p>
</div>
</body>
</html>
"@
        Set-Content -Path $htmlFile -Value $html -Encoding UTF8
        Write-Host "HTML-Report erstellt: $htmlFile"

        if ($createPDF) {
            $pdfFile = [System.IO.Path]::ChangeExtension($htmlFile, ".pdf")
            $wkhtml = "C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe"
            if (Test-Path $wkhtml) {
                & $wkhtml $htmlFile $pdfFile
                Write-Host "PDF-Report erstellt: $pdfFile"
            } else {
                Write-Warning "wkhtmltopdf.exe nicht gefunden -> kein PDF erstellt."
            }
        }
        if ($createTXT) {
            $txtFile = [System.IO.Path]::ChangeExtension($htmlFile, ".txt")
            $txtContent = @"
Onboarding Report
=================
Benutzer  : $DisplayName
SamAccount: $SamAccountName
UPN       : $UPN
AD-Gruppen: $($userSelection.ADGroupsSelected -join ', ')
Standort  : $Standort
Company   : $CompanySelection
Location  : $Location
MS365 Lizenz: $MS365License
Extern    : $Extern
Passwort  : $UserPW
Datum     : $(Get-Date)
"@
            Set-Content -Path $txtFile -Value $txtContent -Encoding UTF8
            Write-Host "TXT-Report erstellt: $txtFile"
        }
    } else {
        Write-Host "HTML-Report: Deaktiviert / nicht gewünscht."
    }
} catch {
    Write-Warning "Fehler beim Erstellen der Reports: $_"
}

Write-Host "`nOnboarding abgeschlossen."
