<#
.SYNOPSIS
    Basismodule mit allgemeinen Hilfsfunktionen für easyONBOARDING.
.DESCRIPTION
    Enthält zentrale Hilfsfunktionen für Debugging, Logging, Fehlerbehandlung und
    GUI-Interaktionen, die von anderen Modulen gemeinsam genutzt werden können.
.NOTES
    Version: 1.0
    Author: Generated based on code analysis
    Creation Date: [Aktuelles Datum]
#>

#region Common Helper Functions
function Write-DebugOutput {
    <#
    .SYNOPSIS
        Verbesserte Debug-Ausgabe-Funktion mit Logging.
    .DESCRIPTION
        Gibt Debugging-Informationen in einer einheitlichen Form aus und schreibt sie optional ins Log.
    .PARAMETER Message
        Die auszugebende Nachricht.
    .PARAMETER Level
        Der Schweregrad der Nachricht (INFO, WARNING, ERROR, DEBUG).
    .PARAMETER LogToFile
        Gibt an, ob die Nachricht auch in eine Logdatei geschrieben werden soll.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, Position = 0)]
        [string]$Message,
        
        [Parameter(Mandatory = $false, Position = 1)]
        [ValidateSet("INFO", "WARNING", "ERROR", "DEBUG")]
        [string]$Level = "INFO",
        
        [Parameter(Mandatory = $false)]
        [switch]$LogToFile
    )
    
    # Definiere Farbcodes entsprechend dem Level
    $colors = @{
        "INFO" = "Cyan"
        "WARNING" = "Yellow"
        "ERROR" = "Red"
        "DEBUG" = "Gray"
    }
    
    # Formatiere Zeitstempel und Ausgabe
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $outputMessage = "[$timestamp] [$Level] $Message"
    
    # Ausgabe in der Konsole
    Write-Host $outputMessage -ForegroundColor $colors[$Level]
    
    # Optionales Logging in Datei
    if ($LogToFile -and (Get-Command -Name "Write-LogMessage" -ErrorAction SilentlyContinue)) {
        Write-LogMessage -Message $Message -LogLevel $Level
    }
    
    # Wenn Level ERROR ist, zusätzlich in den Error-Stream schreiben
    if ($Level -eq "ERROR") {
        Write-Error $Message -ErrorAction Continue
    }
}

function Test-IsAdmin {
    <#
    .SYNOPSIS
        Prüft, ob das Skript mit Administratorrechten ausgeführt wird.
    .DESCRIPTION
        Überprüft, ob der aktuelle PowerShell-Prozess mit Administratorrechten ausgeführt wird.
    .OUTPUTS
        [bool] True, wenn mit Administratorrechten ausgeführt, sonst False.
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param()
    
    try {
        $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
        $principal = New-Object Security.Principal.WindowsPrincipal($identity)
        return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    }
    catch {
        Write-DebugOutput "Fehler bei der Überprüfung der Administratorrechte: $($_.Exception.Message)" -Level ERROR
        return $false
    }
}

function Start-ElevatedProcess {
    <#
    .SYNOPSIS
        Startet den aktuellen Prozess mit erhöhten Rechten neu.
    .DESCRIPTION
        Startet das aktuelle Skript erneut als Administrator, wenn es nicht bereits mit 
        Administratorrechten läuft.
    .OUTPUTS
        Startet einen neuen Prozess und beendet den aktuellen, wenn erfolgreich.
    #>
    [CmdletBinding()]
    param()
    
    try {
        $scriptPath = $MyInvocation.PSCommandPath
        $arguments = "-File `"$scriptPath`" $($MyInvocation.BoundParameters.Keys | ForEach-Object { "-$_ `"$($MyInvocation.BoundParameters[$_])`"" })"
        
        Write-DebugOutput "Starte Prozess mit erhöhten Rechten: powershell.exe $arguments" -Level INFO
        
        Start-Process -FilePath powershell.exe -ArgumentList $arguments -Verb RunAs -Wait
        exit
    }
    catch {
        Write-DebugOutput "Fehler beim Starten mit erhöhten Rechten: $($_.Exception.Message)" -Level ERROR
        throw
    }
}

function Import-RequiredModule {
    <#
    .SYNOPSIS
        Importiert ein Modul mit verbesserter Fehlerbehandlung und Abhängigkeitsmanagement.
    .DESCRIPTION
        Importiert ein PowerShell-Modul aus dem Modulverzeichnis, prüft dabei Abhängigkeiten
        und bietet umfangreiche Fehlerbehandlung.
    .PARAMETER ModuleName
        Der Name des zu importierenden Moduls (ohne .psm1-Erweiterung).
    .PARAMETER Critical
        Gibt an, ob das Modul für die Anwendung kritisch ist.
    .PARAMETER Dependencies
        Eine Liste der Module, von denen dieses Modul abhängt.
    .OUTPUTS
        [bool] True, wenn das Modul erfolgreich geladen wurde, sonst False.
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ModuleName,
        
        [Parameter(Mandatory = $false)]
        [switch]$Critical,
        
        [Parameter(Mandatory = $false)]
        [string[]]$Dependencies = @()
    )
    
    begin {
        if (-not $global:ModulesDir) {
            Write-DebugOutput "FEHLER: ModulesDir ist nicht definiert!" -Level ERROR
            return $false
        }
        
        $modulePath = Join-Path -Path $global:ModulesDir -ChildPath "$ModuleName.psm1"
        Write-DebugOutput "Versuche Modul zu importieren: $ModuleName von $modulePath" -Level DEBUG
    }
    
    process {
        # Erst Abhängigkeiten prüfen und laden
        foreach ($dependency in $Dependencies) {
            $dependencyLoaded = $false
            
            # Prüfen, ob das Abhängigkeitsmodul bereits geladen ist
            if (Get-Module -Name $dependency -ErrorAction SilentlyContinue) {
                Write-DebugOutput "Abhängigkeitsmodul $dependency ist bereits geladen" -Level DEBUG
                $dependencyLoaded = $true
            }
            else {
                # Versuchen, die Abhängigkeit zu laden
                $dependencyPath = Join-Path -Path $global:ModulesDir -ChildPath "$dependency.psm1"
                if (Test-Path $dependencyPath) {
                    try {
                        Import-Module $dependencyPath -Force -Global -ErrorAction Stop
                        Write-DebugOutput "Abhängigkeitsmodul geladen: $dependency" -Level INFO
                        $dependencyLoaded = $true
                    }
                    catch {
                        $errorMsg = "Fehler beim Laden des Abhängigkeitsmoduls '$dependency': $($_.Exception.Message)"
                        Write-DebugOutput $errorMsg -Level ERROR
                        
                        if ($Critical) {
                            throw "Kritische Abhängigkeit '$dependency' konnte nicht geladen werden. Das Skript kann nicht fortgesetzt werden."
                        }
                    }
                }
                else {
                    $errorMsg = "Abhängigkeitsmodul nicht gefunden: $dependency"
                    Write-DebugOutput $errorMsg -Level ERROR
                    
                    if ($Critical) {
                        throw "Kritische Abhängigkeit '$dependency' nicht gefunden. Das Skript kann nicht fortgesetzt werden."
                    }
                }
            }
            
            if (-not $dependencyLoaded -and $Critical) {
                return $false
            }
        }
        
        # Prüfen, ob das Modul bereits geladen ist
        if (Get-Module -Name $ModuleName -ErrorAction SilentlyContinue) {
            Write-DebugOutput "Modul $ModuleName ist bereits geladen" -Level DEBUG
            return $true
        }
        
        # Dann das eigentliche Modul laden
        try {
            if (Test-Path $modulePath) {
                Import-Module $modulePath -Force -Global -ErrorAction Stop
                Write-DebugOutput "Modul geladen: $ModuleName" -Level INFO
                
                # Spezialisierte Initialisierung je nach Modul
                switch ($ModuleName) {
                    "CheckModulesAD" {
                        if (Get-Command -Name "Initialize-ADModule" -ErrorAction SilentlyContinue) {
                            Initialize-ADModule
                        }
                    }
                    "CheckModulesAssemblies" {
                        if (Get-Command -Name "Initialize-RequiredAssemblies" -ErrorAction SilentlyContinue) {
                            Initialize-RequiredAssemblies
                        }
                    }
                    "GUIUPNTemplateInfo" {
                        if (Get-Command -Name "Update-UPNTemplateDisplay" -ErrorAction SilentlyContinue) {
                            Update-UPNTemplateDisplay
                        }
                    }
                }
                
                return $true
            }
            else {
                throw "Moduldatei nicht gefunden: $modulePath"
            }
        }
        catch {
            $errorMsg = "Fehler beim Laden des Moduls '$ModuleName': $($_.Exception.Message)"
            Write-DebugOutput $errorMsg -Level ERROR
            
            if ($Critical) {
                $criticalMsg = "Kritisches Modul '$ModuleName' konnte nicht geladen werden. Das Skript kann nicht fortgesetzt werden."
                Write-DebugOutput $criticalMsg -Level ERROR
                
                throw $criticalMsg
            }
            return $false
        }
    }
}

function Get-UIElement {
    <#
    .SYNOPSIS
        Findet ein UI-Element im WPF-Fenster.
    .DESCRIPTION
        Sucht rekursiv nach einem UI-Element mit dem angegebenen Namen im WPF-Fenster
        oder einem anderen übergeordneten Element.
    .PARAMETER Name
        Der Name des zu suchenden UI-Elements.
    .PARAMETER Parent
        Das übergeordnete Element, in dem gesucht werden soll. Standard ist das Hauptfenster.
    .OUTPUTS
        Das gefundene UI-Element oder $null, wenn es nicht gefunden wurde.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name,
        
        [Parameter(Mandatory = $false)]
        [System.Windows.DependencyObject]$Parent = $global:window
    )
    
    # Zuerst über FindName versuchen (schnell)
    if ($Parent -and $Parent.FindName) {
        $element = $Parent.FindName($Name)
        if ($null -ne $element) {
            return $element
        }
    }
    
    # Wenn das nicht funktioniert und FindVisualChild verfügbar ist
    if (Get-Command -Name "Find-VisualChild" -ErrorAction SilentlyContinue) {
        $element = Find-VisualChild -Parent $Parent -ChildName $Name
        if ($null -ne $element) {
            return $element
        }
    }
    
    # Wenn GUIEventHandler Modul geladen ist, dessen Funktionen nutzen
    if (Get-Command -Name "Find-UIElement" -ErrorAction SilentlyContinue) {
        return Find-UIElement -Name $Name -Parent $Parent
    }
    
    # Rekursive Suche als Fallback
    if ($Parent) {
        try {
            # Visuelle Kinderzahl ermitteln
            $childCount = [System.Windows.Media.VisualTreeHelper]::GetChildrenCount($Parent)
            
            for ($i = 0; $i -lt $childCount; $i++) {
                $child = [System.Windows.Media.VisualTreeHelper]::GetChild($Parent, $i)
                
                # Prüfen, ob das aktuelle Kind den gesuchten Namen hat
                if ($child.Name -eq $Name) {
                    return $child
                }
                
                # Rekursiv in Kindern suchen
                $result = Get-UIElement -Name $Name -Parent $child
                if ($null -ne $result) {
                    return $result
                }
            }
        }
        catch {
            Write-DebugOutput "Fehler bei der Suche nach UI-Element '$Name': $($_.Exception.Message)" -Level ERROR
        }
    }
    
    # Wenn nichts funktioniert hat
    Write-DebugOutput "Element '$Name' nicht gefunden" -Level WARNING
    return $null
}

function Register-EventHandler {
    <#
    .SYNOPSIS
        Registriert einen Event-Handler für ein UI-Element mit verbesserter Fehlerbehandlung.
    .DESCRIPTION
        Fügt einem WPF-Element einen Event-Handler mit umfassender Fehlerbehandlung hinzu und loggt die Registrierung.
    .PARAMETER Element
        Das UI-Element, für das der Event-Handler registriert werden soll.
    .PARAMETER EventName
        Der Name des Events, z.B. "Click".
    .PARAMETER Handler
        Der Scriptblock, der als Handler ausgeführt werden soll.
    .PARAMETER ElementDescription
        Eine Beschreibung des Elements für bessere Logging-Meldungen.
    .OUTPUTS
        [bool] True, wenn die Registrierung erfolgreich war, sonst False.
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param(
        [Parameter(Mandatory = $true)]
        [System.Windows.UIElement]$Element,
        
        [Parameter(Mandatory = $true)]
        [string]$EventName,
        
        [Parameter(Mandatory = $true)]
        [scriptblock]$Handler,
        
        [Parameter(Mandatory = $false)]
        [string]$ElementDescription = ""
    )
    
    if ($null -eq $Element) {
        Write-DebugOutput "Event kann nicht registriert werden: Element ist null ($ElementDescription)" -Level WARNING
        return $false
    }
    
    try {
        # GUIEventHandler Modul verwenden, wenn verfügbar
        if (Get-Command -Name "Register-GUIEvent" -ErrorAction SilentlyContinue) {
            $result = Register-GUIEvent -Control $Element -EventAction $Handler -EventName $EventName -ErrorMessagePrefix "Fehler in $ElementDescription"
            Write-DebugOutput "Event $EventName registriert für $ElementDescription mit Register-GUIEvent" -Level DEBUG
            return $result
        }
        
        # Alternativ direkt registrieren
        $eventMethod = "Add_$EventName"
        $Element.$eventMethod($Handler)
        Write-DebugOutput "Event $EventName direkt registriert für $ElementDescription" -Level DEBUG
        return $true
    }
    catch {
        Write-DebugOutput "Fehler bei der Registrierung des Events $EventName für $ElementDescription : $($_.Exception.Message)" -Level ERROR
        return $false
    }
}

function Import-XamlWindow {
    <#
    .SYNOPSIS
        Importiert eine XAML-Datei als WPF-Fenster mit erweiterter Fehlerbehandlung.
    .DESCRIPTION
        Lädt eine XAML-Datei und erstellt daraus ein WPF-Fenster-Objekt. Bietet umfassende 
        Fehlerbehandlung und Fallback-Mechanismen.
    .PARAMETER XamlPath
        Der vollständige Pfad zur XAML-Datei.
    .OUTPUTS
        Das geladene WPF-Fenster oder eine Ausnahme, wenn es nicht geladen werden konnte.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$XamlPath
    )
    
    Write-DebugOutput "Starte XAML-Import mit erweiterter Fehlerbehandlung" -Level INFO
    
    try {
        # Zuerst prüfen ob die Spezialfunktion verfügbar ist
        if (Get-Command -Name "Import-XAMLGUI" -ErrorAction SilentlyContinue) {
            Write-DebugOutput "Verwende Import-XAMLGUI aus LoadConfigXAML-Modul" -Level INFO
            $window = Import-XAMLGUI -XAMLPath $XamlPath
            
            if ($null -eq $window) {
                throw "Import-XAMLGUI hat ein Null-Fenster-Objekt zurückgegeben"
            }
            
            return $window
        }
        
        # Wenn nicht, schauen nach anderen Funktionen
        if (Get-Command -Name "Load-XamlInterface" -ErrorAction SilentlyContinue) {
            Write-DebugOutput "Verwende Load-XamlInterface als Fallback" -Level INFO
            $window = Load-XamlInterface -XamlPath $XamlPath
            
            if ($null -eq $window) {
                throw "Load-XamlInterface hat ein Null-Fenster-Objekt zurückgegeben"
            }
            
            return $window
        }
        
        # Manueller Import ohne Module als letzte Möglichkeit
        Write-DebugOutput "Manueller XAML-Import als letzter Ausweg" -Level WARNING
        
        if (-not (Test-Path $XamlPath)) {
            throw "XAML-Datei nicht gefunden: $XamlPath"
        }
        
        [xml]$xaml = Get-Content -Path $XamlPath -ErrorAction Stop
        $reader = New-Object System.Xml.XmlNodeReader $xaml
        $window = [Windows.Markup.XamlReader]::Load($reader)
        
        if ($null -eq $window) {
            throw "XamlReader.Load hat ein Null-Fenster-Objekt zurückgegeben"
        }
        
        Write-DebugOutput "Manueller XAML-Import erfolgreich" -Level INFO
        return $window
    }
    catch {
        $errorMsg = "XAML-Import fehlgeschlagen: $($_.Exception.Message)"
        Write-DebugOutput $errorMsg -Level ERROR
        
        # Wenn Debug-Helfer verfügbar, versuche zusätzliche Diagnose
        if ($script:DebugMode -gt 1 -and (Get-Command -Name "Debug-XAMLFile" -ErrorAction SilentlyContinue)) {
            Write-DebugOutput "Versuche XAML-Datei-Diagnose..." -Level INFO
            try {
                $diagResult = Debug-XAMLFile -XamlPath $XamlPath
                Write-DebugOutput "XAML-Diagnose-Ergebnis: $diagResult" -Level INFO
            }
            catch {
                Write-DebugOutput "XAML-Diagnose fehlgeschlagen: $($_.Exception.Message)" -Level ERROR
            }
        }
        
        # Eine einfache Fehleranzeige erstellen
        try {
            $errorWindow = New-Object System.Windows.Window
            $errorWindow.Title = "XAML-Ladefehler"
            $errorWindow.Width = 500
            $errorWindow.Height = 300
            
            $stackPanel = New-Object System.Windows.Controls.StackPanel
            $errorWindow.Content = $stackPanel
            
            $textBlock = New-Object System.Windows.Controls.TextBlock
            $textBlock.Text = "Fehler beim Laden der XAML-Datei: $XamlPath"
            $textBlock.Margin = New-Object System.Windows.Thickness(10)
            $textBlock.TextWrapping = [System.Windows.TextWrapping]::Wrap
            $stackPanel.Children.Add($textBlock)
            
            $errorTextBox = New-Object System.Windows.Controls.TextBox
            $errorTextBox.Text = $_.Exception.ToString()
            $errorTextBox.IsReadOnly = $true
            $errorTextBox.AcceptsReturn = $true
            $errorTextBox.TextWrapping = [System.Windows.TextWrapping]::Wrap
            $errorTextBox.VerticalScrollBarVisibility = [System.Windows.Controls.ScrollBarVisibility]::Auto
            $errorTextBox.Height = 180
            $errorTextBox.Margin = New-Object System.Windows.Thickness(10)
            $stackPanel.Children.Add($errorTextBox)
            
            $closeButton = New-Object System.Windows.Controls.Button
            $closeButton.Content = "Schließen"
            $closeButton.Width = 100
            $closeButton.Margin = New-Object System.Windows.Thickness(10)
            $closeButton.Add_Click({ $errorWindow.Close() })
            $stackPanel.Children.Add($closeButton)
            
            $errorWindow.ShowDialog() | Out-Null
            
            Write-DebugOutput "Fehlerdialog dem Benutzer angezeigt" -Level INFO
        }
        catch {
            Write-DebugOutput "Fehlerdialog konnte nicht erstellt werden: $($_.Exception.Message)" -Level ERROR
        }
        
        throw $errorMsg
    }
}

function Initialize-UIElementsFromFile {
    <#
    .SYNOPSIS
        Findet und initialisiert alle UI-Elemente eines bestimmten Typs aus einer Liste.
    .DESCRIPTION
        Durchsucht eine Textdatei mit Element-Namen und initialisiert alle UI-Elemente, die in der Datei aufgeführt sind.
    .PARAMETER ElementListPath
        Der Pfad zur Textdatei mit den Element-Namen.
    .PARAMETER ElementType
        Der Typ der zu initialisierenden Elemente (z.B. Button, TextBox).
    .PARAMETER ParentElement
        Das übergeordnete Element, in dem die UI-Elemente gesucht werden sollen.
    .OUTPUTS
        [System.Collections.Hashtable] Eine Hashtable mit allen gefundenen Elementen.
    #>
    [CmdletBinding()]
    [OutputType([System.Collections.Hashtable])]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ElementListPath,
        
        [Parameter(Mandatory = $false)]
        [string]$ElementType = "",
        
        [Parameter(Mandatory = $false)]
        [System.Windows.DependencyObject]$ParentElement = $global:window
    )
    
    $elements = @{}
    
    try {
        if (-not (Test-Path $ElementListPath)) {
            Write-DebugOutput "Element-Listendatei nicht gefunden: $ElementListPath" -Level WARNING
            return $elements
        }
        
        $elementNames = Get-Content -Path $ElementListPath -ErrorAction Stop
        
        foreach ($name in $elementNames) {
            # Leere Zeilen und Kommentare überspringen
            if ([string]::IsNullOrWhiteSpace($name) -or $name.Trim().StartsWith("#")) {
                continue
            }
            
            $element = Get-UIElement -Name $name.Trim() -Parent $ParentElement
            
            if ($null -ne $element) {
                # Wenn ein Elementtyp angegeben wurde, nur Elemente dieses Typs hinzufügen
                if ([string]::IsNullOrEmpty($ElementType) -or $element.GetType().Name -eq $ElementType) {
                    $elements[$name] = $element
                    Write-DebugOutput "Element gefunden und initialisiert: $name" -Level DEBUG
                }
            }
            else {
                Write-DebugOutput "Element nicht gefunden: $name" -Level WARNING
            }
        }
        
        Write-DebugOutput "$($elements.Count) Elemente initialisiert" -Level INFO
    }
    catch {
        Write-DebugOutput "Fehler bei der Initialisierung der UI-Elemente: $($_.Exception.Message)" -Level ERROR
    }
    
    return $elements
}
#endregion

#region Configuration Functions
function Get-AppConfiguration {
    <#
    .SYNOPSIS
        Lädt die Anwendungskonfiguration aus der INI-Datei.
    .DESCRIPTION
        Lädt die Anwendungskonfiguration aus der INI-Datei und gibt sie als Hashtable zurück.
    .PARAMETER IniPath
        Der Pfad zur INI-Datei.
    .OUTPUTS
        [hashtable] Die geladene Konfiguration.
    #>
    [CmdletBinding()]
    [OutputType([hashtable])]
    param(
        [Parameter(Mandatory = $true)]
        [string]$IniPath
    )
    
    try {
        if (-not (Test-Path $IniPath)) {
            Write-DebugOutput "INI-Datei nicht gefunden: $IniPath" -Level WARNING
            return @{}
        }
        
        if (Get-Command -Name "Get-IniContent" -ErrorAction SilentlyContinue) {
            $config = Get-IniContent -Path $IniPath
            Write-DebugOutput "Konfiguration aus $IniPath geladen" -Level INFO
            return $config
        }
        else {
            Write-DebugOutput "Get-IniContent-Funktion nicht gefunden. Manuelle INI-Verarbeitung wird verwendet." -Level WARNING
            
            # Einfache Implementierung der INI-Datei-Verarbeitung
            $ini = @{}
            $section = "Default"
            $content = Get-Content -Path $IniPath -ErrorAction Stop
            
            foreach ($line in $content) {
                $line = $line.Trim()
                
                if ([string]::IsNullOrWhiteSpace($line) -or $line.StartsWith(";") -or $line.StartsWith("#")) {
                    continue
                }
                
                if ($line -match "^\[(.+)\]$") {
                    $section = $matches[1]
                    if (-not $ini.ContainsKey($section)) {
                        $ini[$section] = @{}
                    }
                }
                elseif ($line -match "^([^=]+)=(.*)$") {
                    $key = $matches[1].Trim()
                    $value = $matches[2].Trim()
                    
                    if (-not $ini.ContainsKey($section)) {
                        $ini[$section] = @{}
                    }
                    
                    $ini[$section][$key] = $value
                }
            }
            
            Write-DebugOutput "Konfiguration manuell aus $IniPath geladen" -Level INFO
            return $ini
        }
    }
    catch {
        Write-DebugOutput "Fehler beim Laden der Konfiguration: $($_.Exception.Message)" -Level ERROR
        return @{}
    }
}

function Save-AppConfiguration {
    <#
    .SYNOPSIS
        Speichert die Anwendungskonfiguration in der INI-Datei.
    .DESCRIPTION
        Speichert die Anwendungskonfiguration in der INI-Datei.
    .PARAMETER Config
        Die zu speichernde Konfiguration.
    .PARAMETER IniPath
        Der Pfad zur INI-Datei.
    .OUTPUTS
        [bool] True, wenn das Speichern erfolgreich war, sonst False.
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$Config,
        
        [Parameter(Mandatory = $true)]
        [string]$IniPath
    )
    
    try {
        # Sicherstellen, dass das Verzeichnis existiert
        $iniDir = Split-Path -Parent $IniPath
        if (-not (Test-Path $iniDir)) {
            New-Item -Path $iniDir -ItemType Directory -Force | Out-Null
            Write-DebugOutput "Verzeichnis für INI-Datei erstellt: $iniDir" -Level INFO
        }
        
        # Wenn die Out-IniFile-Funktion verfügbar ist, nutze sie
        if (Get-Command -Name "Out-IniFile" -ErrorAction SilentlyContinue) {
            $Config | Out-IniFile -FilePath $IniPath -Force
            Write-DebugOutput "Konfiguration mit Out-IniFile in $IniPath gespeichert" -Level INFO
            return $true
        }
        else {
            Write-DebugOutput "Out-IniFile-Funktion nicht gefunden. Manuelle INI-Erstellung wird verwendet." -Level WARNING
            
            # Einfache INI-Datei-Erstellung
            $content = ""
            
            foreach ($section in $Config.Keys) {
                $content += "[$section]`r`n"
                
                foreach ($key in $Config[$section].Keys) {
                    $value = $Config[$section][$key]
                    $content += "$key=$value`r`n"
                }
                
                $content += "`r`n"
            }
            
            $content | Set-Content -Path $IniPath -Force -Encoding UTF8
            Write-DebugOutput "Konfiguration manuell in $IniPath gespeichert" -Level INFO
            return $true
        }
    }
    catch {
        Write-DebugOutput "Fehler beim Speichern der Konfiguration: $($_.Exception.Message)" -Level ERROR
        return $false
    }
}

function Merge-Configurations {
    <#
    .SYNOPSIS
        Führt zwei Konfigurationen zusammen.
    .DESCRIPTION
        Führt zwei Konfigurationen zusammen, wobei die zweite Konfiguration Vorrang hat.
    .PARAMETER BaseConfig
        Die Basiskonfiguration.
    .PARAMETER OverrideConfig
        Die Konfiguration, die Vorrang hat.
    .OUTPUTS
        [hashtable] Die zusammengeführte Konfiguration.
    #>
    [CmdletBinding()]
    [OutputType([hashtable])]
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$BaseConfig,
        
        [Parameter(Mandatory = $true)]
        [hashtable]$OverrideConfig
    )
    
    $result = $BaseConfig.Clone()
    
    foreach ($section in $OverrideConfig.Keys) {
        if (-not $result.ContainsKey($section)) {
            $result[$section] = @{}
        }
        
        foreach ($key in $OverrideConfig[$section].Keys) {
            $result[$section][$key] = $OverrideConfig[$section][$key]
        }
    }
    
    return $result
}
#endregion

#region Export Module Members
# Public-Funktionen exportieren
Export-ModuleMember -Function Write-DebugOutput
Export-ModuleMember -Function Test-IsAdmin
Export-ModuleMember -Function Start-ElevatedProcess
Export-ModuleMember -Function Import-RequiredModule
Export-ModuleMember -Function Get-UIElement
Export-ModuleMember -Function Register-EventHandler
Export-ModuleMember -Function Import-XamlWindow
Export-ModuleMember -Function Initialize-UIElementsFromFile
Export-ModuleMember -Function Get-AppConfiguration
Export-ModuleMember -Function Save-AppConfiguration
Export-ModuleMember -Function Merge-Configurations
#endregion
