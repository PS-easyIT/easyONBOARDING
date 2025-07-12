#region [UI Manager Module]
# Manages WPF user interface, navigation, and panel switching
# Version: 1.4.XX
# Author: IT Team

#region [Module Variables]
$global:UIManager = @{
    Window = $null
    CurrentPanel = "dashboard"
    Panels = @{
        dashboard = @{ Element = $null; IsVisible = $true }
        onboarding = @{ Element = $null; IsVisible = $false }
        userupdate = @{ Element = $null; IsVisible = $false }
        csvimport = @{ Element = $null; IsVisible = $false }
        reports = @{ Element = $null; IsVisible = $false }
        tools = @{ Element = $null; IsVisible = $false }
        settings = @{ Element = $null; IsVisible = $false }
    }
    NavigationButtons = @{}
    StatusElements = @{}
    IsInitialized = $false
}
#endregion

#region [Initialization Functions]

function Initialize-UIManager {
    <#
    .SYNOPSIS
    Initialisiert den UI Manager mit der WPF-Window-Referenz
    
    .PARAMETER Window
    Die WPF-Window-Instanz
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [System.Windows.Window]$Window
    )
    
    try {
        Write-LogMessage "Initialisiere UI Manager..." "INFO"
        
        $global:UIManager.Window = $Window
        
        # Sammle Panel-Referenzen
        Initialize-PanelReferences
        
        # Sammle Navigation-Button-Referenzen
        Initialize-NavigationButtons
        
        # Sammle Status-Element-Referenzen
        Initialize-StatusElements
        
        # Setze Event-Handler für Navigation
        Set-NavigationEventHandlers
        
        # Initialer Panel-Zustand
        Show-Panel -PanelName "dashboard"
        
        $global:UIManager.IsInitialized = $true
        Write-LogMessage "UI Manager erfolgreich initialisiert" "INFO"
        
        return $true
    }
    catch {
        Write-LogMessage "Fehler beim Initialisieren des UI Manager: $($_.Exception.Message)" "ERROR"
        return $false
    }
}

function Initialize-PanelReferences {
    <#
    .SYNOPSIS
    Sammelt Referenzen zu allen Panel-Elementen
    #>
    
    foreach ($panelName in $global:UIManager.Panels.Keys) {
        $panelElementName = $panelName + "Panel"
        $element = $global:UIManager.Window.FindName($panelElementName)
        
        if ($element) {
            $global:UIManager.Panels[$panelName].Element = $element
            Write-LogMessage "Panel-Referenz gefunden: $panelElementName" "DEBUG"
        }
        else {
            Write-LogMessage "Warnung: Panel-Element nicht gefunden: $panelElementName" "WARN"
        }
    }
}

function Initialize-NavigationButtons {
    <#
    .SYNOPSIS
    Sammelt Referenzen zu allen Navigation-Buttons
    #>
    
    $navButtonNames = @("btnDashboard", "btnOnboarding", "btnUserUpdate", "btnCSVImport", "btnReports", "btnTools", "btnSettings", "btnHeaderSettings")
    
    foreach ($buttonName in $navButtonNames) {
        $button = $global:UIManager.Window.FindName($buttonName)
        
        if ($button) {
            $global:UIManager.NavigationButtons[$buttonName] = $button
            Write-LogMessage "Navigation-Button gefunden: $buttonName" "DEBUG"
        }
        else {
            Write-LogMessage "Warnung: Navigation-Button nicht gefunden: $buttonName" "WARN"
        }
    }
}

function Initialize-StatusElements {
    <#
    .SYNOPSIS
    Sammelt Referenzen zu Status-Anzeige-Elementen
    #>
    
    $statusElementNames = @{
        "lblConnectionStatus" = "AD-Verbindungsstatus"
        "lblDC" = "Domain Controller"
        "lblDomain" = "Domain"
        "lblADStatus" = "AD-Status"
        "lblLastAction" = "Letzte Aktion"
        "lblStatus" = "Allgemeiner Status"
        "lblLastCreatedUser" = "Letzter erstellter Benutzer"
        "lblTotalUsers" = "Benutzer heute"
        "txtOutput" = "Ausgabe-Konsole"
        "txtUpdateLog" = "Update-Log"
    }
    
    foreach ($elementName in $statusElementNames.Keys) {
        $element = $global:UIManager.Window.FindName($elementName)
        
        if ($element) {
            $global:UIManager.StatusElements[$elementName] = $element
            Write-LogMessage "Status-Element gefunden: $elementName" "DEBUG"
        }
        else {
            Write-LogMessage "Warnung: Status-Element nicht gefunden: $elementName" "WARN"
        }
    }
}

#endregion

#region [Navigation Functions]

function Show-Panel {
    <#
    .SYNOPSIS
    Zeigt ein spezifisches Panel an und versteckt alle anderen
    
    .PARAMETER PanelName
    Name des anzuzeigenden Panels
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet("dashboard", "onboarding", "userupdate", "csvimport", "reports", "tools", "settings")]
        [string]$PanelName
    )
    
    try {
        if (-not $global:UIManager.IsInitialized) {
            throw "UI Manager ist nicht initialisiert"
        }
        
        Write-LogMessage "Wechsle zu Panel: $PanelName" "INFO"
        
        # Verstecke alle Panels
        foreach ($panel in $global:UIManager.Panels.Keys) {
            if ($global:UIManager.Panels[$panel].Element) {
                $global:UIManager.Panels[$panel].Element.Visibility = "Collapsed"
                $global:UIManager.Panels[$panel].IsVisible = $false
            }
        }
        
        # Zeige gewähltes Panel
        if ($global:UIManager.Panels.ContainsKey($PanelName) -and $global:UIManager.Panels[$PanelName].Element) {
            $global:UIManager.Panels[$PanelName].Element.Visibility = "Visible"
            $global:UIManager.Panels[$PanelName].IsVisible = $true
            $global:UIManager.CurrentPanel = $PanelName
            
            # Update Navigation-Button-Status
            Update-NavigationButtonStates -ActivePanel $PanelName
            
            # Trigger Panel-spezifische Initialisierung
            Invoke-PanelInitialization -PanelName $PanelName
        }
        else {
            throw "Panel '$PanelName' nicht gefunden oder nicht initialisiert"
        }
        
        return $true
    }
    catch {
        Write-LogMessage "Fehler beim Anzeigen des Panels '$PanelName': $($_.Exception.Message)" "ERROR"
        return $false
    }
}

function Update-NavigationButtonStates {
    <#
    .SYNOPSIS
    Aktualisiert die visuellen Zustände der Navigation-Buttons
    
    .PARAMETER ActivePanel
    Der Name des aktuell aktiven Panels
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ActivePanel
    )
    
    # Mapping zwischen Panel-Namen und Button-Namen
    $panelButtonMapping = @{
        "dashboard" = "btnDashboard"
        "onboarding" = "btnOnboarding"
        "userupdate" = "btnUserUpdate"
        "csvimport" = "btnCSVImport"
        "reports" = "btnReports"
        "tools" = "btnTools"
        "settings" = "btnSettings"
    }
    
    foreach ($panel in $panelButtonMapping.Keys) {
        $buttonName = $panelButtonMapping[$panel]
        
        if ($global:UIManager.NavigationButtons.ContainsKey($buttonName)) {
            $button = $global:UIManager.NavigationButtons[$buttonName]
            
            if ($panel -eq $ActivePanel) {
                # Aktiver Button-Style
                $button.Background = "#0078D4"
                $button.Foreground = "White"
            }
            else {
                # Inaktiver Button-Style
                $button.Background = "#4A90E2"
                $button.Foreground = "White"
            }
        }
    }
}

function Set-NavigationEventHandlers {
    <#
    .SYNOPSIS
    Setzt Event-Handler für alle Navigation-Buttons
    #>
    
    $navigationMapping = @{
        "btnDashboard" = "dashboard"
        "btnOnboarding" = "onboarding"
        "btnUserUpdate" = "userupdate"
        "btnCSVImport" = "csvimport"
        "btnReports" = "reports"
        "btnTools" = "tools"
        "btnSettings" = "settings"
        "btnHeaderSettings" = "settings"
    }
    
    foreach ($buttonName in $navigationMapping.Keys) {
        if ($global:UIManager.NavigationButtons.ContainsKey($buttonName)) {
            $button = $global:UIManager.NavigationButtons[$buttonName]
            $targetPanel = $navigationMapping[$buttonName]
            
            # Remove existing event handlers to prevent duplicates
            $button.remove_Click.Invoke($button.Tag)
            
            # Create new event handler
            $clickHandler = {
                param($sender, $e)
                try {
                    Show-Panel -PanelName $targetPanel
                    Update-StatusText -Element "lblLastAction" -Text "Navigation zu $targetPanel"
                }
                catch {
                    Write-LogMessage "Fehler bei Navigation zu $targetPanel`: $($_.Exception.Message)" "ERROR"
                }
            }.GetNewClosure()
            
            $button.add_Click($clickHandler)
            $button.Tag = $clickHandler
            
            Write-LogMessage "Event-Handler gesetzt für: $buttonName -> $targetPanel" "DEBUG"
        }
    }
}

function Invoke-PanelInitialization {
    <#
    .SYNOPSIS
    Führt Panel-spezifische Initialisierungslogik aus
    
    .PARAMETER PanelName
    Name des zu initialisierenden Panels
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$PanelName
    )
    
    try {
        switch ($PanelName) {
            "dashboard" {
                Update-DashboardInfo
            }
            "onboarding" {
                Initialize-OnboardingPanel
            }
            "userupdate" {
                Initialize-UserUpdatePanel
            }
            "settings" {
                Initialize-SettingsPanel
            }
            default {
                Write-LogMessage "Keine spezifische Initialisierung für Panel '$PanelName' definiert" "DEBUG"
            }
        }
    }
    catch {
        Write-LogMessage "Fehler bei der Panel-Initialisierung für '$PanelName': $($_.Exception.Message)" "ERROR"
    }
}

#endregion

#region [Status Update Functions]

function Update-StatusText {
    <#
    .SYNOPSIS
    Aktualisiert Text in einem Status-Element
    
    .PARAMETER Element
    Name des Status-Elements
    
    .PARAMETER Text
    Der neue Text
    
    .PARAMETER Color
    Optionale Textfarbe
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Element,
        
        [Parameter(Mandatory = $true)]
        [string]$Text,
        
        [Parameter(Mandatory = $false)]
        [string]$Color = ""
    )
    
    try {
        if ($global:UIManager.StatusElements.ContainsKey($Element)) {
            $statusElement = $global:UIManager.StatusElements[$Element]
            
            if ($statusElement -is [System.Windows.Controls.TextBlock]) {
                $statusElement.Text = $Text
                
                if (-not [string]::IsNullOrWhiteSpace($Color)) {
                    $statusElement.Foreground = $Color
                }
            }
            elseif ($statusElement -is [System.Windows.Controls.TextBox]) {
                if ($Element -eq "txtOutput" -or $Element -eq "txtUpdateLog") {
                    # Für Log-TextBoxen: Append mit Timestamp
                    $timestamp = Get-Date -Format "HH:mm:ss"
                    $logEntry = "[$timestamp] $Text`r`n"
                    $statusElement.AppendText($logEntry)
                    $statusElement.ScrollToEnd()
                }
                else {
                    $statusElement.Text = $Text
                }
            }
            
            Write-LogMessage "Status aktualisiert: $Element = $Text" "DEBUG"
        }
        else {
            Write-LogMessage "Warnung: Status-Element '$Element' nicht gefunden" "WARN"
        }
    }
    catch {
        Write-LogMessage "Fehler beim Aktualisieren des Status-Elements '$Element': $($_.Exception.Message)" "ERROR"
    }
}

function Update-ADConnectionStatus {
    <#
    .SYNOPSIS
    Aktualisiert die AD-Verbindungsstatus-Anzeigen
    
    .PARAMETER Status
    Hashtable mit AD-Verbindungsinformationen
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$Status
    )
    
    try {
        if ($Status.IsConnected) {
            Update-StatusText -Element "lblConnectionStatus" -Text "Status: Verbunden" -Color "#107C10"
            Update-StatusText -Element "lblDC" -Text $Status.DomainController
            Update-StatusText -Element "lblDomain" -Text $Status.Domain
            Update-StatusText -Element "lblADStatus" -Text "Verbunden" -Color "#107C10"
        }
        else {
            Update-StatusText -Element "lblConnectionStatus" -Text "Status: Nicht verbunden" -Color "#E81123"
            Update-StatusText -Element "lblDC" -Text "Nicht verfügbar"
            Update-StatusText -Element "lblDomain" -Text "Nicht verfügbar"
            Update-StatusText -Element "lblADStatus" -Text "Nicht verbunden" -Color "#E81123"
            
            if (-not [string]::IsNullOrWhiteSpace($Status.LastError)) {
                Update-StatusText -Element "txtOutput" -Text "AD-Verbindungsfehler: $($Status.LastError)"
            }
        }
    }
    catch {
        Write-LogMessage "Fehler beim Aktualisieren des AD-Verbindungsstatus: $($_.Exception.Message)" "ERROR"
    }
}

function Update-DashboardInfo {
    <#
    .SYNOPSIS
    Aktualisiert die Dashboard-Informationen
    #>
    
    try {
        # AD-Verbindungsstatus abrufen und anzeigen
        if (Get-Module -Name "ADManager" -ErrorAction SilentlyContinue) {
            $adStatus = Get-ADConnectionStatus
            Update-ADConnectionStatus -Status $adStatus
        }
        
        # Weitere Dashboard-Updates können hier hinzugefügt werden
        Update-StatusText -Element "lblLastAction" -Text "Dashboard aktualisiert"
        
        Write-LogMessage "Dashboard-Informationen aktualisiert" "INFO"
    }
    catch {
        Write-LogMessage "Fehler beim Aktualisieren der Dashboard-Informationen: $($_.Exception.Message)" "ERROR"
    }
}

#endregion

#region [Panel-Specific Initialization]

function Initialize-OnboardingPanel {
    <#
    .SYNOPSIS
    Initialisiert das Onboarding-Panel
    #>
    
    try {
        Write-LogMessage "Initialisiere Onboarding-Panel..." "INFO"
        
        # Lade OUs in ComboBox
        if (Get-Module -Name "ADManager" -ErrorAction SilentlyContinue) {
            $cmbOU = $global:UIManager.Window.FindName("cmbOU")
            if ($cmbOU) {
                $ous = Get-AvailableOUs
                $cmbOU.ItemsSource = $ous
                $cmbOU.DisplayMemberPath = "DisplayName"
                $cmbOU.SelectedValuePath = "DistinguishedName"
            }
            
            # Lade Groups in DataGrid
            $dgGroups = $global:UIManager.Window.FindName("dgGroups")
            if ($dgGroups) {
                $groups = Get-AvailableGroups | ForEach-Object {
                    [PSCustomObject]@{
                        IsSelected = $false
                        Name = $_.Name
                        Description = $_.Description
                        Type = $_.TypeDisplay
                    }
                }
                $dgGroups.ItemsSource = $groups
            }
        }
        
        Update-StatusText -Element "lblStatus" -Text "Bereit für Benutzer-Onboarding" -Color "#107C10"
    }
    catch {
        Write-LogMessage "Fehler bei der Initialisierung des Onboarding-Panels: $($_.Exception.Message)" "ERROR"
    }
}

function Initialize-UserUpdatePanel {
    <#
    .SYNOPSIS
    Initialisiert das User-Update-Panel
    #>
    
    try {
        Write-LogMessage "Initialisiere User-Update-Panel..." "INFO"
        
        # Panel-spezifische Initialisierung hier
        Update-StatusText -Element "lblSelectedUser" -Text "Kein Benutzer ausgewählt"
        Update-StatusText -Element "txtUpdateLog" -Text "User-Update-Panel bereit"
    }
    catch {
        Write-LogMessage "Fehler bei der Initialisierung des User-Update-Panels: $($_.Exception.Message)" "ERROR"
    }
}

function Initialize-SettingsPanel {
    <#
    .SYNOPSIS
    Initialisiert das Settings-Panel
    #>
    
    try {
        Write-LogMessage "Initialisiere Settings-Panel..." "INFO"
        
        # Settings-Panel wird in einem separaten Modul implementiert
        # Hier nur grundlegende UI-Initialisierung
        Update-StatusText -Element "txtOutput" -Text "Settings-Panel wird geladen..."
    }
    catch {
        Write-LogMessage "Fehler bei der Initialisierung des Settings-Panels: $($_.Exception.Message)" "ERROR"
    }
}

#endregion

#region [Utility Functions]

function Clear-OutputConsole {
    <#
    .SYNOPSIS
    Leert die Ausgabe-Konsole
    #>
    
    try {
        if ($global:UIManager.StatusElements.ContainsKey("txtOutput")) {
            $global:UIManager.StatusElements["txtOutput"].Clear()
        }
    }
    catch {
        Write-LogMessage "Fehler beim Leeren der Ausgabe-Konsole: $($_.Exception.Message)" "ERROR"
    }
}

function Get-CurrentPanel {
    <#
    .SYNOPSIS
    Gibt den Namen des aktuell angezeigten Panels zurück
    #>
    
    return $global:UIManager.CurrentPanel
}

function Test-UIManagerInitialized {
    <#
    .SYNOPSIS
    Prüft ob der UI Manager initialisiert ist
    #>
    
    return $global:UIManager.IsInitialized
}

#endregion

#region [Export Functions]
Export-ModuleMember -Function @(
    'Initialize-UIManager',
    'Show-Panel',
    'Update-StatusText',
    'Update-ADConnectionStatus',
    'Update-DashboardInfo',
    'Initialize-OnboardingPanel',
    'Initialize-UserUpdatePanel',
    'Initialize-SettingsPanel',
    'Clear-OutputConsole',
    'Get-CurrentPanel',
    'Test-UIManagerInitialized'
)
#endregion 
# SIG # Begin signature block
# MIIcCAYJKoZIhvcNAQcCoIIb+TCCG/UCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCYoYow1FCFNrFh
# R3EdlbrlXy9x7DCxMZ0YfeP6YZOcM6CCFk4wggMQMIIB+KADAgECAhB3jzsyX9Cg
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
# BgkqhkiG9w0BCQQxIgQgtC7H59HwQf5+TiIvik5C96EoMGiVsKHo/ZPB6C1wRDQw
# DQYJKoZIhvcNAQEBBQAEggEAnxdocpeVUwfpH/s+RQV367VCSpbN8DqnwyaaYQa6
# CXGI0fM0+zY3m5u0uEzsVBwfLeDlvZNXulnFHeeAZWSTaoradSsMZ55kZ6SaiAZb
# xqF6bujTXuRaIbEmeirR7ZSw081qNLHV+3grAP+1djQe6tyjIL2EP1nu97L2Ph1R
# gtKA9bHAdrpHy+qoOLXFWt8dpcZaIPPiKKWSBajvNhUBn902La1MGF7+ttMnwSAD
# ZL0fE5e6dWScZ/Gh8BYIqIMYkUT7qnCkW/JRShFgQ+fwTSK8Z8tr5KFVVac5OrVn
# dkIEJXhbKSSDlnulTJKLPhLGjPNT90CauScdh7FALMhoUaGCAyYwggMiBgkqhkiG
# 9w0BCQYxggMTMIIDDwIBATB9MGkxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdp
# Q2VydCwgSW5jLjFBMD8GA1UEAxM4RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3Rh
# bXBpbmcgUlNBNDA5NiBTSEEyNTYgMjAyNSBDQTECEAqA7xhLjfEFgtHEdqeVdGgw
# DQYJYIZIAWUDBAIBBQCgaTAYBgkqhkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqG
# SIb3DQEJBTEPFw0yNTA3MTExOTMyMzNaMC8GCSqGSIb3DQEJBDEiBCDNH3iCYV8u
# WXWQlF1jDagOhHimfOj7PjIJJsTlFtYCjzANBgkqhkiG9w0BAQEFAASCAgBYNbTU
# yhNVZpXjfdMUKKc91pP3/89UrY37yvYtRRE+qOkQcBMiy3KehCxphKy+ZC/+9Gwt
# SyE9Dh5RT9kAD7kiTy0pjEGAjOIYjk4aDsLPhX11XFw0Brfc6GbMcmbQ66uINSa/
# e/UBQOY2FU+WlgBGTgrZQhDyUXF89XnGZhBwhLmlxfh6b1ReeWPY9BdORQpPuTHe
# 1nX+GDgTHGP19gSUu/lz6EWTbxlCk4tLPCzeRnQEi5KwN1dvcoYiW2TcgipWkylV
# 7Z+zOo+5df3NCBmRqv5zrsOuYZ1rHSJjFqDnhJmuEe7cw/p4Eg8DyWrRSi90oFCj
# 9r8Zzg41QikolXs2du4ReF6sIOGnAAS+uuWNf8RbQN85BZ7Q3NhDjwEvgSV0L9Io
# bBEnYr6f7zUyZ5+0NsGyxRyJu9SKva40RfcX1gTEmB/ISpaBP9Bi98OEauNffks9
# 6x3Zxu0li7Z3SSXIjIxRcwlJOG+TF3j0fRq12+mSnFXMl6ISzu0Um8mOu4TDomfw
# DJfes3p+22+x1wYhqP4etdmfdFC+GTfz2ph8RxJw0FHhtpVBN7c7a0ht5G/WkAtY
# nOJvq/gBjcp4Udt/EfVSjsk7oZTKdFiGc0Kf0xo/mK4Ff61WuCtTf9THlrntjdrG
# zrd3TH8VGUfl3fGQpM9UOcxLv9CIAqiMvaSeLA==
# SIG # End signature block
