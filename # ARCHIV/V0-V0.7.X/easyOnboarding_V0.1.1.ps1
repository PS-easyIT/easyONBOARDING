#requires -Version 5.1
[CmdletBinding()]
param(
    [string]$Vorname,
    [string]$Nachname,
    [string]$Standort,
    [string]$Company,
    [string]$License = "",
    [switch]$Extern,
    [string]$ScriptINIPath = "easyOnboarding_V0.1.1_Config.ini"
)

##############################################################################
# 1) INI-Datei einlesen (OrderedDictionary – .Contains verwenden)
##############################################################################
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
                $result[$section] = New-Object 'System.Collections.Specialized.OrderedDictionary'
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

##############################################################################
# 2) GUI-Erstellung (Drei Panels: Links, Rechts, Unten)
##############################################################################
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
    $form.Size = New-Object System.Drawing.Size(1100,700)
    $form.Topmost = $false

    # Panel oben: Info-Label
    $lblInfo = New-Object System.Windows.Forms.Label
    $lblInfo.AutoSize = $true
    $lblInfo.Font = 'Microsoft Sans Serif,10,style=Bold'
    $lblInfo.Location = '10,10'
    $lblInfo.Size = New-Object System.Drawing.Size(1000,40)
    $form.Controls.Add($lblInfo)

    # Panel links: Manuelle Eingaben
    $panelLeft = New-Object System.Windows.Forms.Panel
    $panelLeft.Location = New-Object System.Drawing.Point(10,60)
    $panelLeft.Size = New-Object System.Drawing.Size(520,550)
    $panelLeft.AutoScroll = $true
    $panelLeft.BorderStyle = 'FixedSingle'
    $form.Controls.Add($panelLeft)

    # Panel rechts: AD-Gruppen, Flags, Passwortoptionen, UPN, Output
    $panelRight = New-Object System.Windows.Forms.Panel
    $panelRight.Location = New-Object System.Drawing.Point(550,60)
    $panelRight.Size = New-Object System.Drawing.Size(520,550)
    $panelRight.AutoScroll = $true
    $panelRight.BorderStyle = 'FixedSingle'
    $form.Controls.Add($panelRight)

    # Panel unten: Buttons – mit etwas Rand (10px Abstand links/rechts)
    $panelBottom = New-Object System.Windows.Forms.Panel
    $panelBottom.Location = New-Object System.Drawing.Point(10,630)
    $panelBottom.Size = New-Object System.Drawing.Size(1080,60)
    $panelBottom.BorderStyle = 'None'
    $panelBottom.Anchor = [System.Windows.Forms.AnchorStyles] "Left,Right,Bottom"
    $form.Controls.Add($panelBottom)

    # Hilfsfunktionen
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
        foreach ($i in $items) { $null = $cmb.Items.Add($i) }
        if ($default -and $cmb.Items.Contains($default)) {
            $cmb.SelectedItem = $default
        } elseif ($cmb.Items.Count -gt 0) {
            $cmb.SelectedIndex = 0
        }
        $parent.Controls.Add($cmb)
        return $cmb
    }

    # --- Aus INI auslesen: ScriptInfo / General ---
    $ScriptVersion = ""
    $LastUpdate = ""
    $Author = ""
    $domainName1 = ""
    $defaultOU = ""
    $reportPath = ""

    if ($INIConfig.Contains('ScriptInfo')) {
        $ScriptVersion = $INIConfig['ScriptInfo'].ScriptVersion
        $LastUpdate = $INIConfig['ScriptInfo'].LastUpdate
        $Author = $INIConfig['ScriptInfo'].Author
    }
    if ($INIConfig.Contains('General')) {
        $domainName1 = $INIConfig['General'].DomainName1
        $defaultOU = $INIConfig['General'].DefaultOU
        $reportPath = $INIConfig['General'].ReportPath
    }
    $lblInfo.Text = "ScriptVersion=$ScriptVersion | LastUpdate=$LastUpdate | Author=$Author`r`nDOMAIN: $domainName1 | OU: $defaultOU | REPORT: $reportPath"

    ############################################################################
    # PanelLeft => Manuelle Eingaben
    ############################################################################
    $yLeft = 10
    AddLabel $panelLeft "Vorname:" 10 $yLeft | Out-Null
    $txtVorname = AddTextBox $panelLeft "" 150 $yLeft; $yLeft += 30

    AddLabel $panelLeft "Nachname:" 10 $yLeft | Out-Null
    $txtNachname = AddTextBox $panelLeft "" 150 $yLeft; $yLeft += 30

    AddLabel $panelLeft "Anzeigename:" 10 $yLeft | Out-Null
    $txtDisplayName = AddTextBox $panelLeft "" 150 $yLeft 300; $yLeft += 30

    AddLabel $panelLeft "Beschreibung:" 10 $yLeft | Out-Null
    $txtDescription = AddTextBox $panelLeft "" 150 $yLeft 300; $yLeft += 30

    AddLabel $panelLeft "Büro (OfficeRoom):" 10 $yLeft | Out-Null
    $txtOfficeRoom = AddTextBox $panelLeft "" 150 $yLeft 300; $yLeft += 30

    AddLabel $panelLeft "Telefon (Phone):" 10 $yLeft | Out-Null
    $txtPhone = AddTextBox $panelLeft "" 150 $yLeft 300; $yLeft += 30

    AddLabel $panelLeft "Mobil (Mobile):" 10 $yLeft | Out-Null
    $txtMobile = AddTextBox $panelLeft "" 150 $yLeft 300; $yLeft += 30

    AddLabel $panelLeft "Position:" 10 $yLeft | Out-Null
    $txtPosition = AddTextBox $panelLeft "" 150 $yLeft 300; $yLeft += 30

    AddLabel $panelLeft "Abteilung (manuell):" 10 $yLeft | Out-Null
    $txtDepartmentField = AddTextBox $panelLeft "" 150 $yLeft 300; $yLeft += 30

    AddLabel $panelLeft "CustomAttrib1:" 10 $yLeft | Out-Null
    $txtCA1 = AddTextBox $panelLeft "" 150 $yLeft 300; $yLeft += 30

    AddLabel $panelLeft "CustomAttrib2:" 10 $yLeft | Out-Null
    $txtCA2 = AddTextBox $panelLeft "" 150 $yLeft 300; $yLeft += 30

    AddLabel $panelLeft "CustomAttrib3:" 10 $yLeft | Out-Null
    $txtCA3 = AddTextBox $panelLeft "" 150 $yLeft 300; $yLeft += 30

    AddLabel $panelLeft "CustomAttrib4:" 10 $yLeft | Out-Null
    $txtCA4 = AddTextBox $panelLeft "" 150 $yLeft 300; $yLeft += 30

    AddLabel $panelLeft "CustomAttrib5:" 10 $yLeft | Out-Null
    $txtCA5 = AddTextBox $panelLeft "" 150 $yLeft 300; $yLeft += 30

    ############################################################################
    # PanelRight => AD-Gruppen, AD-Flags, Passwortoptionen, UPN, Output
    ############################################################################
    $yRight = 10

    # Überschrift für AD-Gruppen
    AddLabel $panelRight "AD-Gruppen (Mehrfachauswahl):" 10 $yRight | Out-Null
    $yRight += 25

    $groupChecks = @{}
    if ($INIConfig.Contains("ADGroups")) {
        $allKeys = $INIConfig["ADGroups"].Keys
        $defaultADGroups = @()
        if ($INIConfig["ADGroups"].Contains("DefaultADGroups")) {
            $defaultADGroups = $INIConfig["ADGroups"].DefaultADGroups -split ';'
        }
        $groupKeys = $allKeys | Where-Object {
            $_ -notmatch "_Label$" -and $_ -notmatch "_Visible$" -and $_ -notmatch "^DefaultADGroups$"
        }
        foreach ($gk in $groupKeys) {
            $visibleKey = $gk + "_Visible"
            $isVisible = $true
            if ($INIConfig["ADGroups"].Contains($visibleKey)) {
                if ($INIConfig["ADGroups"][$visibleKey] -eq "0") {
                    $isVisible = $false
                }
            }
            if ($isVisible) {
                $labelKey = $gk + "_Label"
                $labelTxt = $gk
                if ($INIConfig["ADGroups"].Contains($labelKey)) {
                    $labelTxt = $INIConfig["ADGroups"][$labelKey]
                }
                $cbGroup = AddCheckBox $panelRight $labelTxt $false 10 $yRight
                if ($defaultADGroups -contains $INIConfig["ADGroups"][$gk]) {
                    $cbGroup.Checked = $true
                }
                $groupChecks[$gk] = $cbGroup
                $yRight += 25
            }
        }
    }

    # AD-Flags mit Zeilenumbruch im Label
    AddLabel $panelRight "AD-Benutzer-Flags:`r`n" 10 $yRight | Out-Null
    $yRight += 20

    $chkPWNeverExpires = AddCheckBox $panelRight "PasswordNeverExpires" $false 10 $yRight
    $chkMustChange     = AddCheckBox $panelRight "MustChangePasswordAtLogon" $false 250 $yRight
    $yRight += 25

    $chkAccountDisabled = AddCheckBox $panelRight "AccountDisabled" $false 10 $yRight
    $chkCannotChangePW  = AddCheckBox $panelRight "CannotChangePassword" $false 250 $yRight
    $yRight += 35

    # Passwortoptionen
    AddLabel $panelRight "Passwort-Optionen:" 10 $yRight | Out-Null
    $yRight += 20

    $rbFix = New-Object System.Windows.Forms.RadioButton
    $rbFix.Text = "FEST"
    $rbFix.Location = New-Object System.Drawing.Point(10, $yRight)
    $panelRight.Controls.Add($rbFix)
    $rbRand = New-Object System.Windows.Forms.RadioButton
    $rbRand.Text = "GENERIERT"
    $rbRand.Location = New-Object System.Drawing.Point(150, $yRight)
    $panelRight.Controls.Add($rbRand)
    $yRight += 25

    $lblFixPW = AddLabel $panelRight "Festes Passwort:" 10 $yRight
    $txtFixPW = AddTextBox $panelRight "" 130 $yRight 150
    $yRight += 30

    $lblPWLen = AddLabel $panelRight "PW-Länge:" 10 $yRight
    $txtPWLen = AddTextBox $panelRight "12" 130 $yRight 50
    $yRight += 30

    $chkIncludeSpecial = AddCheckBox $panelRight "IncludeSpecialChars" $true 10 $yRight
    $chkAvoidAmbig     = AddCheckBox $panelRight "AvoidAmbiguousChars" $true 250 $yRight
    $yRight += 35

    function UpdatePWFields {
        if ($rbFix.Checked) {
            $txtFixPW.Enabled = $true
            $lblFixPW.Enabled = $true
            $txtPWLen.Enabled = $false
            $lblPWLen.Enabled = $false
            $chkIncludeSpecial.Enabled = $false
            $chkAvoidAmbig.Enabled = $false
        }
        else {
            $txtFixPW.Enabled = $false
            $lblFixPW.Enabled = $false
            $txtPWLen.Enabled = $true
            $lblPWLen.Enabled = $true
            $chkIncludeSpecial.Enabled = $true
            $chkAvoidAmbig.Enabled = $true
        }
    }
    $rbFix.Add_CheckedChanged({ UpdatePWFields })
    $rbRand.Add_CheckedChanged({ UpdatePWFields })
    $rbRand.Checked = $true
    UpdatePWFields
    $yRight += 10

    # Output-Optionen
    AddLabel $panelRight "OUTPUT - Onboarding?" 10 $yRight | Out-Null
    $yRight += 20
    $chkHTML = AddCheckBox $panelRight "HTML erzeugen" $true 10 $yRight
    $chkPDF  = AddCheckBox $panelRight "PDF erzeugen"  $true 150 $yRight
    $chkTXT  = AddCheckBox $panelRight "TXT erzeugen"  $true 290 $yRight
    $yRight += 35

    # UPN
    AddLabel $panelRight "UPN (manuell):" 10 $yRight | Out-Null
    $txtUPN = AddTextBox $panelRight "" 130 $yRight 200
    $yRight += 30
    AddLabel $panelRight "UPN-Format-Vorlagen:" 10 $yRight | Out-Null
    $cmbUPNFormat = AddComboBox $panelRight @("VORNAME.NACHNAME","V.NACHNAME","VORNAMENACHNAME","VNACHNAME") 130 $yRight 200
    $yRight += 40

    ############################################################################
    # PanelBottom => Buttons (mit Abstand zum Rand)
    ############################################################################
    $btnOK = New-Object System.Windows.Forms.Button
    $btnOK.Text = "Onboarding Starten"
    $btnOK.Size = New-Object System.Drawing.Size(150,30)
    $btnOK.Location = New-Object System.Drawing.Point(350,10)
    $panelBottom.Controls.Add($btnOK)

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "Abbrechen"
    $btnCancel.Size = New-Object System.Drawing.Size(150,30)
    $btnCancel.Location = New-Object System.Drawing.Point(530,10)
    $panelBottom.Controls.Add($btnCancel)

    # Ergebnis-Objekt
    $result = [PSCustomObject]@{
        Vorname             = ""
        Nachname            = ""
        DisplayName         = ""
        Description         = ""
        OfficeRoom          = ""
        PhoneNumber         = ""
        MobileNumber        = ""
        Position            = ""
        DepartmentField     = ""
        CustomAttribute1    = ""
        CustomAttribute2    = ""
        CustomAttribute3    = ""
        CustomAttribute4    = ""
        CustomAttribute5    = ""
        PasswordNeverExpires= $false
        MustChangePassword  = $false
        AccountDisabled     = $false
        CannotChangePassword= $false
        PasswordMode        = 1
        FixPassword         = ""
        PasswordLaenge      = 12
        IncludeSpecialChars = $true
        AvoidAmbiguousChars = $true
        OutputHTML          = $true
        OutputPDF           = $true
        OutputTXT           = $true
        UPNEntered          = ""
        UPNFormat           = "VORNAME.NACHNAME"
        Cancel              = $false
        ADGroupsSelected    = @()
    }

    # Klick "Onboarding Starten"
    $btnOK.Add_Click({
        $result.Vorname         = $txtVorname.Text
        $result.Nachname        = $txtNachname.Text
        $result.DisplayName     = $txtDisplayName.Text
        $result.Description     = $txtDescription.Text
        $result.OfficeRoom      = $txtOfficeRoom.Text
        $result.PhoneNumber     = $txtPhone.Text
        $result.MobileNumber    = $txtMobile.Text
        $result.Position        = $txtPosition.Text
        $result.DepartmentField = $txtDepartmentField.Text
        $result.CustomAttribute1= $txtCA1.Text
        $result.CustomAttribute2= $txtCA2.Text
        $result.CustomAttribute3= $txtCA3.Text
        $result.CustomAttribute4= $txtCA4.Text
        $result.CustomAttribute5= $txtCA5.Text

        # AD-Gruppen-Checkboxen
        $selGroups = @()
        foreach ($gk in $groupChecks.Keys) {
            if ($groupChecks[$gk].Checked) {
                $selGroups += $gk
            }
        }
        $result.ADGroupsSelected = $selGroups

        # AD-Flags
        $result.PasswordNeverExpires = $chkPWNeverExpires.Checked
        $result.MustChangePassword   = $chkMustChange.Checked
        $result.AccountDisabled      = $chkAccountDisabled.Checked
        $result.CannotChangePassword = $chkCannotChangePW.Checked

        # Passwortoptionen
        if ($rbFix.Checked) {
            $result.PasswordMode = 0
        }
        else {
            $result.PasswordMode = 1
        }
        $result.FixPassword = $txtFixPW.Text
        $result.PasswordLaenge = [int]$txtPWLen.Text
        $result.IncludeSpecialChars = $chkIncludeSpecial.Checked
        $result.AvoidAmbiguousChars = $chkAvoidAmbig.Checked

        # Output
        $result.OutputHTML = $chkHTML.Checked
        $result.OutputPDF = $chkPDF.Checked
        $result.OutputTXT = $chkTXT.Checked

        # UPN
        $result.UPNEntered = $txtUPN.Text.Trim()
        $result.UPNFormat = $cmbUPNFormat.SelectedItem

        # Validierung
        if (-not $result.Vorname) {
            [System.Windows.Forms.MessageBox]::Show("Vorname darf nicht leer sein!", "Fehler", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            return
        }
        if (-not $result.Nachname) {
            [System.Windows.Forms.MessageBox]::Show("Nachname darf nicht leer sein!", "Fehler", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
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

##############################################################################
# 3) Hauptablauf: INI laden, GUI anzeigen, Werte verarbeiten
##############################################################################
Write-Host "Lade INI: $ScriptINIPath"
$Config = Read-INIFile $ScriptINIPath

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
$DeptField           = $userSelection.DepartmentField
$CA1                 = $userSelection.CustomAttribute1
$CA2                 = $userSelection.CustomAttribute2
$CA3                 = $userSelection.CustomAttribute3
$CA4                 = $userSelection.CustomAttribute4
$CA5                 = $userSelection.CustomAttribute5
$passwordNeverExpires= $userSelection.PasswordNeverExpires
$mustChangePW        = $userSelection.MustChangePassword
$accountDisabled     = $userSelection.AccountDisabled
$cannotChangePW      = $userSelection.CannotChangePassword
$passwordMode        = $userSelection.PasswordMode
$fixPassword         = $userSelection.FixPassword
$passwordLaenge      = $userSelection.PasswordLaenge
$includeSpecial      = $userSelection.IncludeSpecialChars
$avoidAmbiguous      = $userSelection.AvoidAmbiguousChars
$createHTML          = $userSelection.OutputHTML
$createPDF           = $userSelection.OutputPDF
$createTXT           = $userSelection.OutputTXT
$UPNManual           = $userSelection.UPNEntered
$UPNTemplate         = $userSelection.UPNFormat
$selectedGroups      = $userSelection.ADGroupsSelected

Write-Host "`nStarte Onboarding für: $Vorname $Nachname ..."

if (-not $Vorname) {
    Write-Warning "Vorname ist leer. Abbruch."
    return
}
if (-not $Nachname) {
    Write-Warning "Nachname ist leer. Abbruch."
    return
}

Write-Host "DisplayName=$DisplayName, Description=$Description, Phone=$PhoneNumber"
Write-Host "Fertig - Hier folgt dann die AD-Logik, Logging, Reports, etc."

##############################################################################
# 4) Active Directory-Erstellung (Beispielcode)
##############################################################################
try {
    Import-Module ActiveDirectory -ErrorAction Stop
} catch {
    Write-Warning "Das ActiveDirectory-Modul konnte nicht geladen werden: $_"
    return
}

# Beispiel: Generierung von SamAccountName
# Falls Vorname/Nachname gefüllt sind:
if ($Vorname.Length -ge 1 -and $Nachname.Length -ge 1) {
    $SamAccountName = ($Vorname.Substring(0,1) + $Nachname).ToLower()
} else {
    $SamAccountName = "user_" + (Get-Random)
}
Write-Host "SamAccountName: $SamAccountName"

# UPN-Logik (falls manuell leer ist)
if (-not [string]::IsNullOrEmpty($UPNManual)) {
    $UPN = $UPNManual
} else {
    switch -Wildcard ($UPNTemplate) {
        "VORNAME.NACHNAME"    { $UPN = "$Vorname.$Nachname@phinit.de" }
        "V.NACHNAME"          { $UPN = "$($Vorname.Substring(0,1)).$Nachname@phinit.de" }
        "VORNAMENACHNAME"     { $UPN = "$Vorname$Nachname@phinit.de" }
        "VNACHNAME"           { $UPN = "$($Vorname.Substring(0,1))$Nachname@phinit.de" }
        default               { $UPN = "$SamAccountName@phinit.de" }
    }
}
Write-Host "UPN: $UPN"

# Passwort
if ($passwordMode -eq 0) {
    # Fixes Passwort
    $UserPW = $fixPassword
} else {
    # Generiertes Passwort
    function Generate-RandomPassword {
        param(
            [int]$Length,
            [bool]$IncludeSpecial,
            [bool]$AvoidAmbiguous
        )
        $minNonAlpha = 2
        $pw = [System.Web.Security.Membership]::GeneratePassword($Length, $minNonAlpha)
        if ($AvoidAmbiguous) {
            # Ersetze Zeichen, die als "ambiguous" gelten könnten
            $pw = $pw -replace '[{}()\[\]\/\\`~,;:.<>\"]','X'
        }
        return $pw
    }
    $UserPW = Generate-RandomPassword -Length $passwordLaenge -IncludeSpecial $includeSpecial -AvoidAmbiguous $avoidAmbiguous
}
Write-Host "Passwort (Klartext): $UserPW"

# SecureString
$SecurePW = ConvertTo-SecureString $UserPW -AsPlainText -Force

# AD-Benutzer anlegen oder aktualisieren
try {
    $existingUser = Get-ADUser -Filter { SamAccountName -eq $SamAccountName } -ErrorAction SilentlyContinue
} catch {
    $existingUser = $null
}

if (-not $existingUser) {
    Write-Host "Lege neuen AD-Benutzer an: $SamAccountName"
    try {
        New-ADUser `
            -Name $DisplayName `
            -GivenName $Vorname `
            -Surname $Nachname `
            -SamAccountName $SamAccountName `
            -UserPrincipalName $UPN `
            -AccountPassword $SecurePW `
            -Enabled ([bool] -not $accountDisabled) `
            -ChangePasswordAtLogon $mustChangePW `
            -PasswordNeverExpires $passwordNeverExpires `
            -Path "OU=Mitarbeiter,DC=phinit,DC=de" `
            -Description $Description `
            -Office $OfficeRoom `
            -Title $Position `
            -ErrorAction Stop
        Write-Host "AD-Benutzer erstellt."
    } catch {
        Write-Warning "Fehler beim Erstellen des Benutzers: $_"
        return
    }
} else {
    Write-Host "Benutzer '$SamAccountName' existiert bereits – führe Update durch."
    try {
        Set-ADUser -Identity $existingUser.DistinguishedName `
            -GivenName $Vorname `
            -Surname $Nachname `
            -Name $DisplayName `
            -Description $Description `
            -Office $OfficeRoom `
            -Title $Position `
            -Enabled ([bool] -not $accountDisabled) `
            -ErrorAction SilentlyContinue

        # Kennwort-Flags
        Set-ADUser -Identity $existingUser.DistinguishedName -ChangePasswordAtLogon:$mustChangePW -PasswordNeverExpires:$passwordNeverExpires
    } catch {
        Write-Warning "Fehler beim Aktualisieren: $_"
    }
}

# Passwort zurücksetzen
try {
    Set-ADAccountPassword -Identity $SamAccountName -Reset -NewPassword $SecurePW -ErrorAction SilentlyContinue
    Write-Host "Passwort zurückgesetzt."
} catch {
    Write-Warning "Fehler beim Setzen des Passworts: $_"
}

if ($cannotChangePW) {
    Write-Host "(Hinweis) 'CannotChangePassword' => Per ACL in AD-Berechtigungen umsetzbar."
}

# Zusätzliche Felder (z.B. PhoneNumber, MobileNumber, DepartmentField, CustomAttributeX)
try {
    $updateProps = @{
        telephoneNumber      = $PhoneNumber
        mobile              = $MobileNumber
        department          = $DeptField
        extensionAttribute1 = $CA1
        extensionAttribute2 = $CA2
        extensionAttribute3 = $CA3
        extensionAttribute4 = $CA4
        extensionAttribute5 = $CA5
    }
    $updateProps.GetEnumerator() | Where-Object { $_.Value -and $_.Value -ne "" } | ForEach-Object {
        Set-ADUser -Identity $SamAccountName -Add @{ ($_.Key) = $_.Value } -ErrorAction SilentlyContinue
    }
    Write-Host "Zusätzliche Felder aktualisiert."
} catch {
    Write-Warning "Fehler bei zusätzlichen Feldern: $_"
}

##############################################################################
# 5) Gruppenmitgliedschaften (ausgewählte ADGroups)
##############################################################################
if ($userSelection.ADGroupsSelected.Count -gt 0) {
    foreach ($grpKey in $userSelection.ADGroupsSelected) {
        $grpValue = $Config.ADGroups[$grpKey]
        if ($grpValue) {
            try {
                Add-ADGroupMember -Identity $grpValue -Members $SamAccountName -ErrorAction SilentlyContinue
                Write-Host "Gruppe '$grpValue' zugewiesen (Key=$grpKey)."
            } catch {
                Write-Warning "Fehler bei Gruppe '$grpValue': $_"
            }
        }
    }
}

##############################################################################
# 6) Logging (Beispiel)
##############################################################################
if ($Config.General.LogFilePath) {
    try {
        $logPath = $Config.General.LogFilePath
        if (-not (Test-Path $logPath)) {
            New-Item -ItemType Directory -Path $logPath -Force | Out-Null
        }
        $logFile = Join-Path $logPath "Onboarding_$(Get-Date -Format 'yyyyMMdd').log"
        $logEntry = "[{0}] Sam={1}, Anzeigename={2}, UPN={3}, PW={4}, Groups={5}" -f (Get-Date), $SamAccountName, $DisplayName, $UPN, $UserPW, ($userSelection.ADGroupsSelected -join ',')
        Add-Content -Path $logFile -Value $logEntry
        Write-Host "Log geschrieben: $logFile"
    } catch {
        Write-Warning "Fehler beim Logging: $_"
    }
}

##############################################################################
# 7) Reports (HTML, PDF, TXT) je nach Checkbox
##############################################################################
try {
    $reportPath = $Config.General.ReportPath
    if ($reportPath -and (Test-Path $reportPath)) {
        $samHtml = Join-Path $reportPath "$SamAccountName.html"
        if ($createHTML) {
            $htmlContent = @"
<html>
<head><title>Onboarding-Report</title></head>
<body>
<h1>Onboarding Report</h1>
<p><b>Benutzer:</b> $DisplayName ($Vorname $Nachname)</p>
<p><b>SamAccountName:</b> $SamAccountName</p>
<p><b>UPN:</b> $UPN</p>
<p><b>Abteilungen/Groups:</b> $($userSelection.ADGroupsSelected -join ', ')</p>
<p><b>Passwort (Klartext):</b> $UserPW</p>
<p><b>Telefon:</b> $PhoneNumber, <b>Mobil:</b> $MobileNumber</p>
<p>Erstellt am: $(Get-Date)</p>
</body>
</html>
"@
            Set-Content -Path $samHtml -Value $htmlContent -Encoding UTF8
            Write-Host "HTML-Report erstellt: $samHtml"

            if ($createPDF) {
                $pdfPath = [System.IO.Path]::ChangeExtension($samHtml, ".pdf")
                $wkhtml = "C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe"
                if (Test-Path $wkhtml) {
                    & $wkhtml $samHtml $pdfPath
                    Write-Host "PDF-Report erstellt: $pdfPath"
                } else {
                    Write-Warning "wkhtmltopdf.exe nicht gefunden – kein PDF erstellt."
                }
            }
            if ($createTXT) {
                $txtPath = [System.IO.Path]::ChangeExtension($samHtml, ".txt")
                $txtContent = @"
Onboarding Report
=================
Benutzer : $DisplayName ($Vorname $Nachname)
SamAccount: $SamAccountName
UPN       : $UPN
Passwort  : $UserPW
Datum     : $(Get-Date)
"@
                Set-Content -Path $txtPath -Value $txtContent -Encoding UTF8
                Write-Host "TXT-Report erstellt: $txtPath"
            }
        }
    }
} catch {
    Write-Warning "Fehler beim Erstellen des Reports: $_"
}

Write-Host "`nOnboarding-Prozess abgeschlossen."


# SIG # Begin signature block
# MIIbywYJKoZIhvcNAQcCoIIbvDCCG7gCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCZpDhdFAlUvkxb
# rUBU2kfcDtrcbkmbMs5J7/ygs7LN/aCCFhcwggMQMIIB+KADAgECAhB3jzsyX9Cg
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
# DQEJBDEiBCBtcKcEjnbVNERvE1z9tMkf3OaHv+mPg6X0lJqxVcMQFjANBgkqhkiG
# 9w0BAQEFAASCAQAyKCyD13jNGxroW6Gz2VuHzJBKtTafk9HMAOZqAvPlV4GnHvhR
# J3Cmu94RTU7BN9pMfoUHBKD9t/ihy1KL8tY3Z8uzSQQjX/eWYKZQ/fTw2Lj6ZLLw
# dOXxgjKRfR5I2j9IxytyZ+6NlYXLt6cZEZ1SzdiBv1Zo3l/raDxZVfCRn/H7RzM7
# bvfokAFcop4qIp9PxPtHKDbdJSVlYvmKfassRcgSwnWw3PFTNwEXAzoDxIUDq2nw
# HecrKXbh3FtibpZmHJZY+YuV3VZvN5phu7OSuli46ay5AcwjIQYBzKwHOdthIjrn
# JRQd7AnMdFNL3KjFWPjhdy2orl8pHp82zaA6oYIDIDCCAxwGCSqGSIb3DQEJBjGC
# Aw0wggMJAgEBMHcwYzELMAkGA1UEBhMCVVMxFzAVBgNVBAoTDkRpZ2lDZXJ0LCBJ
# bmMuMTswOQYDVQQDEzJEaWdpQ2VydCBUcnVzdGVkIEc0IFJTQTQwOTYgU0hBMjU2
# IFRpbWVTdGFtcGluZyBDQQIQC65mvFq6f5WHxvnpBOMzBDANBglghkgBZQMEAgEF
# AKBpMBgGCSqGSIb3DQEJAzELBgkqhkiG9w0BBwEwHAYJKoZIhvcNAQkFMQ8XDTI1
# MDcwNTEwMTE1NVowLwYJKoZIhvcNAQkEMSIEIPG3A6Gk5RCJqgQWY/PxnF6X2b0B
# WrcVoC01e3nrY595MA0GCSqGSIb3DQEBAQUABIICABL15ueXHGEYb4kjr5mg7jmd
# hrINMA3YRRVMYAJ1p9EhCrSQv2oou83u2Lv2TFFPLRey8QOWtuh8b7WLzb6TORBd
# cOapJKBRmhmgA2D5eefjS5o7mXb/FncqIbqW7EE3kQSDOeLIqecH2LM089cfUVuj
# kE9fifY0axdxF8ZKpi9LalDpMYhEkqIHzbU6lAO9rdEvhLzXKUEckZI7HhOIL4oK
# wEHF58P8fiNj1InOo75K2HsbMIY/CjvR4p4rf3h61No6KTsnWjZk3BluRl9zukT4
# QOygNXIfjImJbvBAcIBqoVX1P/rbu38xjYjiNul6Qjw3AjSH8mdHTy7ugqeXI59c
# kAcIMt/UbuCyZwWT/U8N3snrHGZPpaEgOpdwfXp2rmMhsI+RQJSzlLnVYqCuGvIs
# xH/wIQBRDdSG9iUaj8wgCi2WGJ2FL9c+97Nnqbd+b61MZE+kUE00MSHcbTsSRRpT
# HcPeqhd55dHZRvOaRVWy/41qQ4QV9L0/4FG0GwXN9y1wI9Wjy4qVtY/O0LqldSau
# ulkuSinzduWZIQik7TDY/2ESMc14PwGOluFElcpupELVe4RAu+uYHkCMgR7IY0Y0
# pW14V63gpqdrtxTv136yXV6sDgC90GKhuzN/aysNxJ6UH1k7mAr8TxcQvA+zAg9M
# LcwDsY0tuOC5EyyQXJPh
# SIG # End signature block
