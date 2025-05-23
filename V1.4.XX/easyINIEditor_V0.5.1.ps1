# Erweiterter INI Editor mit festen GUI-Einstellungen und Änderungsprotokollierung
# (ohne dynamische Buttons zum Hinzufügen/Löschen von Abschnitten und Schlüsseln)
# WPF Version

# Globale Debug-Variable - Standardwert (wird später aus [General] überschrieben, falls vorhanden)
$script:Debug = 0

#############################################
# Debug-Ausgabe Funktion
#############################################
function Write-DebugMessage {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Message,
        [int]$Level = 1
    )
    
    # Wenn Debug-Level ausreichend hoch ist oder Force benutzt wird
    if ($script:Debug -ge $Level) {
        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        $caller = (Get-PSCallStack)[1].Command
        $debugMessage = "[$timestamp] [DEBUG] [$caller] $Message"
        
        # Ausgabe auf der Konsole
        Write-Host $debugMessage -ForegroundColor Cyan
        
        # Optional auch in Log-Datei schreiben
        if ($null -ne $logFilePath -and $logFilePath -ne "") {
            try {
                Add-Content -Path $logFilePath -Value $debugMessage -ErrorAction SilentlyContinue
            } catch {
                # Silent fail für Log-Fehler im Debug-Modus
            }
        }
    }
}

# Benötigte Assemblies für WPF laden
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase
Write-DebugMessage "WPF-Assemblies geladen"

# Klasse für Grid-Items mit PropertyChanged-Ereignissen für WPF-Binding
Add-Type -TypeDefinition @"
    using System;
    using System.ComponentModel;
    using System.Runtime.CompilerServices;
    using System.Collections.ObjectModel;

    public class IniDataItem : INotifyPropertyChanged
    {
        private string _key;
        private string _value;
        private string _comment;

        public event PropertyChangedEventHandler PropertyChanged;

        public string Key 
        { 
            get { return _key; } 
            set 
            { 
                if (_key != value) 
                {
                    _key = value;
                    NotifyPropertyChanged();
                }
            } 
        }

        public string Value 
        { 
            get { return _value; } 
            set 
            { 
                if (_value != value) 
                {
                    _value = value;
                    NotifyPropertyChanged();
                }
            } 
        }

        public string Comment 
        { 
            get { return _comment; } 
            set 
            { 
                if (_comment != value) 
                {
                    _comment = value;
                    NotifyPropertyChanged();
                }
            } 
        }

        protected void NotifyPropertyChanged([CallerMemberName] string propertyName = "")
        {
            if (PropertyChanged != null)
                PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
        }
    }

    public class IniDataCollection : ObservableCollection<IniDataItem> { }
"@
Write-DebugMessage "IniDataItem-Klasse definiert"

#############################################
# Zusätzliche Funktion: Loggen
#############################################
function Log-Change {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Message,
        [string]$LogPath = "LOG_easyINIEditor-Changes.log",
        [switch]$IsDebug
    )
    try {
        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        $prefix = if ($IsDebug) { "DEBUG" } else { "CHANGE" }
        $logEntry = "[$timestamp] [$prefix] $Message"
        
        # In Logdatei schreiben
        Add-Content -Path $LogPath -Value $logEntry -ErrorAction Stop
        
        # Bei Debug-Meldungen auch auf Konsole ausgeben
        if ($IsDebug -and $script:Debug -gt 0) {
            Write-Host $logEntry -ForegroundColor Cyan
        }
    }
    catch {
        Write-Warning "Fehler beim Schreiben in die Log-Datei ($LogPath): $_"
    }
}

#############################################
# INI einlesen (mit Kommentarunterstützung)
#############################################
function Read-IniFile {
    param(
        [Parameter(Mandatory=$true)]
        [ValidateScript({Test-Path $_ -PathType Leaf})]
        [string]$Path
    )
    
    Write-DebugMessage "Lese INI-Datei: $Path"
    $ini = [ordered]@{}
    $currentSection = ""
    $pendingComment = ""

    try {
        foreach ($line in Get-Content $Path -Encoding UTF8 -ErrorAction Stop) {
            $trimmed = $line.Trim()

            # Kommentarzeile?
            if ($trimmed -match "^\s*;") {
                # Kommentartext ohne führendes ";"
                $commentText = $trimmed.Substring(1).Trim()
                if ([string]::IsNullOrEmpty($pendingComment)) {
                    $pendingComment = $commentText
                }
                else {
                    $pendingComment += "`n" + $commentText
                }
                continue
            }

            # Abschnittszeile [Abschnitt]?
            if ($trimmed -match "^\[(.+?)\]") {
                $currentSection = $matches[1]
                Write-DebugMessage "Abschnitt gefunden: [$currentSection]" -Level 2

                # Falls Abschnitt noch nicht existiert => anlegen
                if (-not $ini.Contains($currentSection)) {
                    $ini[$currentSection] = [ordered]@{}
                }
                # Kommentar gilt nicht mehr für nächsten Abschnitt
                $pendingComment = ""
            }

            # key = value ? (abgesehen von reinen Leer-/Kommentarzeilen)
            elseif ($trimmed -match "^(.*?)\s*=\s*(.*)$") {
                $key = $matches[1].Trim()
                $value = $matches[2].Trim()

                if (-not [string]::IsNullOrEmpty($currentSection)) {
                    if (-not $ini.Contains($currentSection)) {
                        $ini[$currentSection] = [ordered]@{}
                    }
                    $ini[$currentSection][$key] = [PSCustomObject]@{
                        Value   = $value
                        Comment = $pendingComment
                    }
                    Write-DebugMessage "[$currentSection] Key: '$key', Value: '$value'" -Level 3
                }
                $pendingComment = ""
            }
            else {
                # Unbekannte Zeile -> Kommentar zurücksetzen
                $pendingComment = ""
            }
        }
        Write-DebugMessage "INI-Datei erfolgreich eingelesen: $Path mit $(($ini.Keys).Count) Abschnitten"
        return $ini
    }
    catch {
        Write-Error "Fehler beim Lesen der INI-Datei '$Path': $_"
        return [ordered]@{}
    }
}

#############################################
# INI schreiben (Kommentare wieder einfügen)
#############################################
function Write-IniFile {
    param(
        [Parameter(Mandatory=$true)]
        [hashtable]$IniData,
        [Parameter(Mandatory=$true)]
        [string]$Path,
        [System.Collections.ArrayList]$DesiredOrder
    )
    
    Write-DebugMessage "Schreibe INI-Datei: $Path"
    try {
        # Abschnitte sortieren: zuerst DesiredOrder, dann Rest
        $remaining = [ordered]@{}
        foreach ($section in $IniData.Keys) {
            $remaining[$section] = $IniData[$section]
        }
        $orderedIni = [ordered]@{}

        if ($null -ne $DesiredOrder) {
            foreach ($section in $DesiredOrder) {
                # Contains statt ContainsKey verwenden
                if ($remaining.Contains($section)) {
                    $orderedIni[$section] = $remaining[$section]
                    $remaining.Remove($section) | Out-Null
                }
            }
        }
        
        foreach ($left in $remaining.Keys) {
            $orderedIni[$left] = $remaining[$left]
        }
        
        $lines = @()
        foreach ($sec in $orderedIni.Keys) {
            $lines += "[$sec]"
            if ($orderedIni[$sec].Count -gt 0) {
                foreach ($key in $orderedIni[$sec].Keys) {
                    $entry = $orderedIni[$sec][$key]
                    if (-not [string]::IsNullOrEmpty($entry.Comment)) {
                        foreach ($cl in $entry.Comment -split "`n") {
                            $lines += "; " + $cl
                        }
                    }
                    $lines += "$key=$($entry.Value)"
                }
            }
            $lines += ""  # Leere Zeile als Trenner
        }
        
        # UTF8 ohne BOM verwenden
        $utf8NoBom = New-Object System.Text.UTF8Encoding $false
        [System.IO.File]::WriteAllLines($Path, $lines, $utf8NoBom)
        Write-DebugMessage "INI-Datei erfolgreich geschrieben: $Path"
        
        return $true
    }
    catch {
        Write-Error "Fehler beim Schreiben der INI-Datei '$Path': $_"
        return $false
    }
}

#############################################
# Gewünschte Reihenfolge festlegen
#############################################
$desiredOrder = New-Object System.Collections.ArrayList
[void]$desiredOrder.AddRange(@(
    "ScriptInfo",
    "Logging",
    "RemoteExecution",
    "WPFGUI",
    "Report",
    "ADUserDefaults",
    "ActivateUserMS365ADSync",
    "MailEndungen",
    "Company",
    "MailEndungen",
    "ADGroups",
    "Jira"
))

# Hierarchie für die Anzeige in der GUI
$sectionHierarchy = @{
    "WPFGUI" = @("WPFGUILogos")
    "Report" = @("CustomPWLabels", "Websites")
    "ADUserDefaults" = @("UserCreationDefaults", "DisplayNameUPNTemplates", "PasswordFixGenerate")
    "MailEndungen" = @("ProxyMailAddress")
    "Company" = @("CompanyHelpdesk", "CompanyHR", "CompanyWLAN", "CompanyVPN")
    "ADGroups" = @("LicensesGroups", "TLGroups", "ALGroups")
}

Write-DebugMessage "Reihenfolge der Abschnitte definiert"

#############################################
# Pfad zur INI-Datei (alle Dateien mit .ini laden)
#############################################
$iniPath = "*.ini"
Write-DebugMessage "Suche nach INI-Dateien mit Pfad: $iniPath"

#############################################
# Alle INI-Dateien einlesen und in $AllIniData speichern
#############################################
$AllIniData = @{}
try {
    $files = Get-ChildItem -Path $iniPath -File -ErrorAction Stop
    Write-DebugMessage "Gefundene INI-Dateien: $($files.Count)"
    
    if ($files.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Keine INI-Dateien gefunden im aktuellen Verzeichnis.", "Hinweis", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        Write-DebugMessage "Keine INI-Dateien gefunden. Programm wird beendet."
        exit
    }

    foreach ($f in $files) {
        Write-DebugMessage "Lade INI-Datei: $($f.FullName)"
        $AllIniData[$f.FullName] = Read-IniFile -Path $f.FullName
    }
}
catch {
    $errorMsg = "Fehler beim Laden der INI-Dateien: $_"
    Write-DebugMessage $errorMsg
    [System.Windows.Forms.MessageBox]::Show($errorMsg, "Fehler", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    exit
}

#############################################
# Editor-Einstellungen aus [CONFIGEDITOR] aus der ersten INI-Datei laden (falls vorhanden)
#############################################
$editorSettings = @{
    FormBackColor     = "#F0F0F0"
    ListViewBackColor = "#FFFFFF"
    DataGridBackColor = "#FAFAFA"
    FontName          = "Segoe UI"
    FontSize          = 12
    LogFilePath       = ".\Logs\ini_changes.log"
}
$firstFile = $files[0].FullName
Write-DebugMessage "Lade Editor-Einstellungen aus erster INI-Datei: $firstFile"

# Debug-Einstellung aus [General] laden, falls vorhanden
if ($AllIniData[$firstFile].Keys -contains "General") {
    if ($AllIniData[$firstFile]["General"].Keys -contains "Debug") {
        $script:Debug = [int]($AllIniData[$firstFile]["General"]["Debug"].Value)
        Write-DebugMessage "Debug-Einstellung aus INI-Datei geladen: $($script:Debug)"
    }
}

if ($AllIniData[$firstFile].Keys -contains "CONFIGEDITOR") {
    foreach ($key in $AllIniData[$firstFile]["CONFIGEDITOR"].Keys) {
        $valObj = $AllIniData[$firstFile]["CONFIGEDITOR"][$key]
        if ($valObj -and $editorSettings.Contains($key)) {
            $editorSettings[$key] = $valObj.Value
            Write-DebugMessage "Editor-Einstellung geladen: $key = $($valObj.Value)" -Level 2
        }
    }
}
$logFilePath = $editorSettings["LogFilePath"]
Write-DebugMessage "Log-Datei-Pfad: $logFilePath"

# Verzeichnis für Log-Datei erstellen, falls noch nicht vorhanden
$logDir = Split-Path -Parent $logFilePath
if (-not (Test-Path $logDir)) {
    Write-DebugMessage "Log-Verzeichnis existiert nicht, wird erstellt: $logDir"
    try {
        New-Item -Path $logDir -ItemType Directory -Force | Out-Null
    }
    catch {
        Write-Warning "Fehler beim Erstellen des Log-Verzeichnisses: $_"
    }
}

# Erstes Log schreiben
Log-Change -Message "easyINIEditor gestartet. Debug-Level: $script:Debug" -LogPath $logFilePath -IsDebug

#############################################
# GUI erstellen mit WPF
#############################################

# XAML-Definition der Benutzeroberfläche
[xml]$xaml = @"
<Window 
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="INI Editor" Height="825" Width="1250" 
    WindowStartupLocation="CenterScreen">
    <Window.Resources>
        <Style TargetType="Button">
            <Setter Property="Margin" Value="5"/>
            <Setter Property="Padding" Value="10,5"/>
            <Setter Property="FontSize" Value="12"/>
            <Setter Property="Height" Value="40"/>
        </Style>
    </Window.Resources>
    <Grid Margin="0,10,0,5" Background="#FFB7CFDC">
        <!-- Angepasste RowDefinitions -->
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <!-- Header -->
            <RowDefinition Height="*"/>
            <!-- Hauptinhalt -->
            <RowDefinition Height="Auto"/>
            <!-- Buttons -->
            <RowDefinition Height="Auto"/>
            <!-- Footer -->
        </Grid.RowDefinitions>

        <!-- Neuer Header -->
        <StackPanel Orientation="Horizontal" Grid.Row="0" HorizontalAlignment="Center" Margin="0,0,0,10" Background="#FF88A0AD" Width="1250">
            <Image Source="$PSScriptRoot\assets\appiconinieditor.png" Width="150" Height="60"/>
            <TextBlock Text="easyONBOARDING - INIEditor" VerticalAlignment="Center" FontSize="24" Margin="10,0,0,0"/>
        </StackPanel>

        <!-- Hauptinhalt verschoben in Grid.Row="1" -->
        <Grid Grid.Row="1">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="300"/>
                <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>
            <!-- ListView für Sections -->
            <ListView Name="SectionListView" Grid.Column="0" Margin="5,0,14,133" 
                      SelectionMode="Single" 
                      IsSynchronizedWithCurrentItem="False"/>
            <!-- DataGrid für Keys und Values -->
            <DataGrid Name="KeyValueGrid" Grid.Column="1" AutoGenerateColumns="False" 
                      CanUserAddRows="False" 
                      SelectionUnit="Cell" SelectionMode="Single" Height="646" VerticalAlignment="Top" Margin="0,0,5,0">
                <DataGrid.Columns>
                    <DataGridTextColumn Header="Key" Binding="{Binding Key, Mode=TwoWay, UpdateSourceTrigger=PropertyChanged}" Width="150" IsReadOnly="False"/>
                    <DataGridTextColumn Header="Value" Binding="{Binding Value, Mode=TwoWay, UpdateSourceTrigger=PropertyChanged}" Width="275" IsReadOnly="False"/>
                    <DataGridTextColumn Header="Comment" Binding="{Binding Comment}" Width="450" IsReadOnly="True"/>
                </DataGrid.Columns>
            </DataGrid>
        </Grid>

        <!-- Buttons verschoben in Grid.Row="2" -->
        <DockPanel Grid.Row="1" HorizontalAlignment="Left" Margin="0,541,0,18">
            <Button Name="SaveButton" Width="87" Height="100" Background="LightGreen" DockPanel.Dock="Left">
                <StackPanel Orientation="Vertical">
                    <Image Source="$PSScriptRoot\assets\okbutton.png" Width="64" Height="64"/>
                    <TextBlock Text="Save" HorizontalAlignment="Center"/>
                </StackPanel>
            </Button>
            <Button Name="InfoButton" Width="87" Height="100" Background="LightBlue" DockPanel.Dock="Left">
                <StackPanel Orientation="Vertical">
                    <Image Source="$PSScriptRoot\assets\infofaq.png" Width="64" Height="64"/>
                    <TextBlock Text="Info" HorizontalAlignment="Center"/>
                </StackPanel>
            </Button>
            <Button Name="CloseButton" Width="87" Height="100" Background="LightCoral" DockPanel.Dock="Left">
                <StackPanel Orientation="Vertical">
                    <Image Source="$PSScriptRoot\assets\close.png" Width="64" Height="64"/>
                    <TextBlock Text="Close" HorizontalAlignment="Center"/>
                </StackPanel>
            </Button>
        </DockPanel>

        <!-- Neuer Footer in Grid.Row="3" -->
        <Border Grid.Row="3" BorderThickness="1,0,0,0" BorderBrush="Gray" Margin="0,10,0,-5" Background="#FF88A0AD">
            <StackPanel Orientation="Horizontal" HorizontalAlignment="Center" Margin="5">
                <TextBlock Text="Version: 0.6.2 " FontSize="10" Margin="0,0,25,0"/>
                <TextBlock Text="Andreas Hepp | PhinIT.de  -  PSscripts.de" FontSize="10"/>
            </StackPanel>
        </Border>
    </Grid>
</Window>
"@

# XAML in WPF-Objekte umwandeln
try {
    $reader = [System.Xml.XmlNodeReader]::new($xaml)
    $window = [System.Windows.Markup.XamlReader]::Load($reader)
}
catch {
    Write-Error "Fehler beim Laden des XAML: $_"
    exit 1
}

# WPF-Steuerelemente abrufen
$sectionListView = $window.FindName("SectionListView")
$keyValueGrid = $window.FindName("KeyValueGrid")
$saveButton = $window.FindName("SaveButton")
$infoButton = $window.FindName("InfoButton")
$closeButton = $window.FindName("CloseButton")

# Prüfen, ob alle Steuerelemente gefunden wurden
if (-not ($sectionListView -and $keyValueGrid -and $saveButton -and $infoButton -and $closeButton)) {
    Write-Error "Konnte nicht alle UI-Elemente finden. UI könnte fehlerhaft sein."
    exit 1
}

# Farben aus den Einstellungen anwenden
$window.Background = [System.Windows.Media.BrushConverter]::new().ConvertFromString($editorSettings.FormBackColor)
$sectionListView.Background = [System.Windows.Media.BrushConverter]::new().ConvertFromString($editorSettings.ListViewBackColor)
$keyValueGrid.Background = [System.Windows.Media.BrushConverter]::new().ConvertFromString($editorSettings.DataGridBackColor)

# Schriftart aus Einstellungen anwenden
$fontFamily = [System.Windows.Media.FontFamily]::new($editorSettings.FontName)
$fontSize = [double]::Parse($editorSettings.FontSize)
$window.FontFamily = $fontFamily
$window.FontSize = $fontSize

#############################################
# Linke Ansicht befüllen (ListView)
#############################################
$script:currentFile = $null
$script:currentSection = $null
$script:originalValues = @{}
$script:gridItems = New-Object IniDataCollection

function Populate-SectionListView {
    # Items-Objekt leeren statt direkt mit der Collection zu arbeiten
    $sectionListView.Items.Clear()
    
    # Eine Liste für die ListViewItems erstellen
    $listItems = New-Object System.Collections.ObjectModel.ObservableCollection[System.Windows.Controls.ListViewItem]

    foreach ($filePath in $AllIniData.Keys) {
        # Fetter Header für Dateiname
        $fileHeaderItem = [System.Windows.Controls.ListViewItem]::new()
        $fileHeaderTextBlock = [System.Windows.Controls.TextBlock]::new()
        $fileHeaderTextBlock.Text = "INI File: " + (Split-Path $filePath -Leaf)
        $fileHeaderTextBlock.FontWeight = "Bold"
        $fileHeaderItem.Content = $fileHeaderTextBlock
        
        # Tag mit Metadaten
        $fileHeaderItem.Tag = @{
            Type    = "FileHeader"
            File    = $filePath
            Section = $null
        }
        $listItems.Add($fileHeaderItem)

        # Reihenfolge der Sections festlegen
        $currentIni = $AllIniData[$filePath]
        $allSections = $currentIni.Keys
        $processedSections = New-Object System.Collections.ArrayList
        
        # Zuerst die Hauptabschnitte in der definierten Reihenfolge hinzufügen
        foreach ($sec in $desiredOrder) {
            if ($allSections -contains $sec) {
                $sectionItem = [System.Windows.Controls.ListViewItem]::new()
                $sectionTextBlock = [System.Windows.Controls.TextBlock]::new()
                $sectionTextBlock.Text = $sec
                $sectionItem.Content = $sectionTextBlock
                $sectionItem.Tag = @{
                    Type    = "Section"
                    File    = $filePath
                    Section = $sec
                    IsChild = $false
                }
                $listItems.Add($sectionItem)
                [void]$processedSections.Add($sec)
                
                # Unterabschnitte hinzufügen, falls vorhanden
                if ($sectionHierarchy.ContainsKey($sec)) {
                    foreach ($childSec in $sectionHierarchy[$sec]) {
                        if ($allSections -contains $childSec) {
                            $childItem = [System.Windows.Controls.ListViewItem]::new()
                            $childTextBlock = [System.Windows.Controls.TextBlock]::new()
                            $childTextBlock.Text = "   -- " + $childSec
                            $childItem.Content = $childTextBlock
                            $childItem.Tag = @{
                                Type    = "Section"
                                File    = $filePath
                                Section = $childSec
                                IsChild = $true
                                Parent  = $sec
                            }
                            $listItems.Add($childItem)
                            [void]$processedSections.Add($childSec)
                        }
                    }
                }
            }
        }
        
        # Dann alle restlichen Abschnitte hinzufügen, die nicht in der Liste waren
        foreach ($sec in $allSections) {
            if (-not $processedSections.Contains($sec)) {
                $sectionItem = [System.Windows.Controls.ListViewItem]::new()
                $sectionItem.Content = $sec
                $sectionItem.Tag = @{
                    Type    = "Section"
                    File    = $filePath
                    Section = $sec
                    IsChild = $false
                }
                $listItems.Add($sectionItem)
            }
        }
    }
    
    # Alle Items auf einmal zum ListView hinzufügen
    foreach ($item in $listItems) {
        $sectionListView.Items.Add($item)
    }
}

Populate-SectionListView

#############################################
# Event-Handler für Steuerelemente
#############################################

# ListView-Event für Sectionauswahl
$sectionListView.Add_SelectionChanged({
    $selectedItem = $sectionListView.SelectedItem
    if ($selectedItem -eq $null) { return }
    
    $tag = $selectedItem.Tag
    if ($tag -and $tag.Type -eq "Section") {
        $script:currentFile = $tag.File
        $script:currentSection = $tag.Section

        # Neue Collection für DataGrid erstellen
        $script:gridItems = New-Object IniDataCollection
        $keyValueGrid.ItemsSource = $script:gridItems
        
        $thisIniData = $AllIniData[$script:currentFile]
        
        # Originaldaten für Vergleiche speichern
        $script:originalValues = @{}
        
        if ($thisIniData.Contains($script:currentSection)) {
            foreach ($key in $thisIniData[$script:currentSection].Keys) {
                $entry = $thisIniData[$script:currentSection][$key]
                
                # DataGrid-Item erstellen mit der definierten Klasse
                $item = New-Object IniDataItem
                $item.Key = $key
                $item.Value = $entry.Value
                $item.Comment = $entry.Comment
                
                $script:gridItems.Add($item)
                
                # Originalwert speichern
                $script:originalValues[$key] = $entry.Value
            }
        }
    }
    else {
        $script:currentFile = $null
        $script:currentSection = $null
        $script:originalValues = @{}
        $script:gridItems = New-Object IniDataCollection
        $keyValueGrid.ItemsSource = $script:gridItems
    }
})

# DataGrid CellEditEnding Event Handler
$keyValueGrid.Add_CellEditEnding({
    param($sender, $e)
    
    try {
        # Stellen Sie sicher, dass die Zelle bearbeitet wurde
        if ($e.EditAction -eq [System.Windows.Controls.DataGridEditAction]::Commit) {
            # Die Änderungen werden durch die INotifyPropertyChanged-Implementierung automatisch übernommen
            # Wir müssen hier nichts weiter tun
        }
    }
    catch {
        Write-Warning "Fehler beim Bearbeiten der Zelle: $_"
        # Verhindert Absturz bei Bearbeitungsfehlern
    }
})

# Save-Button Event
$saveButton.Add_Click({
    if (-not $script:currentFile -or -not $script:currentSection) {
        [System.Windows.MessageBox]::Show("Bitte wählen Sie einen Abschnitt aus.", "Hinweis")
        return
    }

    try {
        # Neue Section-Daten aufbauen
        $sectionData = [ordered]@{}
        $changedKeys = @()
        
        foreach ($item in $script:gridItems) {
            $key = $item.Key
            $value = $item.Value
            $comment = $item.Comment

            if (-not [string]::IsNullOrEmpty($key)) {
                $sectionData[$key] = [PSCustomObject]@{
                    Value   = $value
                    Comment = $comment
                }
                
                # Prüfen ob Wert geändert wurde
                if ($script:originalValues.Contains($key) -and $script:originalValues[$key] -ne $value) {
                    $changedKeys += "Key '$key' geändert: '$($script:originalValues[$key])' -> '$value'"
                }
                elseif (-not $script:originalValues.Contains($key)) {
                    $changedKeys += "Key '$key' hinzugefügt mit Wert: '$value'"
                }
            }
        }
        
        # Prüfen ob Keys entfernt wurden
        foreach ($oldKey in $script:originalValues.Keys) {
            if (-not $sectionData.Contains($oldKey)) {
                $changedKeys += "Key '$oldKey' entfernt (alter Wert: '$($script:originalValues[$oldKey])')"
            }
        }

        try {
            # Abschnitt ersetzen
            $AllIniData[$script:currentFile][$script:currentSection] = $sectionData

            # INI-Datei schreiben
            $success = Write-IniFile -IniData $AllIniData[$script:currentFile] -Path $script:currentFile -DesiredOrder $desiredOrder
            
            if ($success) {
                [System.Windows.MessageBox]::Show("INI-Datei erfolgreich gespeichert.", "Gespeichert")
                
                # Änderungen loggen
                $fileName = Split-Path $script:currentFile -Leaf
                Log-Change -Message "Abschnitt '$script:currentSection' in Datei '$fileName' wurde gespeichert." -LogPath $logFilePath
                
                foreach ($change in $changedKeys) {
                    Log-Change -Message "  - $change" -LogPath $logFilePath
                }
                
                # Originalwerte aktualisieren
                $script:originalValues = @{}
                foreach ($key in $sectionData.Keys) {
                    $script:originalValues[$key] = $sectionData[$key].Value
                }
            }
            else {
                [System.Windows.MessageBox]::Show("Fehler beim Speichern der INI-Datei.", "Fehler")
            }
        }
        catch {
            [System.Windows.MessageBox]::Show("Fehler beim Speichern: $_", "Fehler")
        }
    }
    catch {
        [System.Windows.MessageBox]::Show("Fehler beim Speichern: $_", "Fehler")
    }
})

# Info-Button Event
$infoButton.Add_Click({
    if (Test-Path ".\assets\info_editor.txt") {
        Start-Process "notepad.exe" -ArgumentList "info_editor.txt"
    }
    else {
        [System.Windows.MessageBox]::Show("Die Datei 'info_editor.txt' wurde nicht gefunden.", "Info")
    }
})

# Close-Button Event
$closeButton.Add_Click({
    $window.Close()
})

#############################################
# Fenster anzeigen
#############################################
# Sicherstellen, dass alles initialisiert ist, bevor wir das Fenster anzeigen
try {
    # DataGrid mit einer leeren Collection initialisieren
    $script:gridItems = New-Object IniDataCollection
    $keyValueGrid.ItemsSource = $script:gridItems
    
    Populate-SectionListView
    [void]$window.ShowDialog()
}
catch {
    Write-Error "Fehler beim Anzeigen des Fensters: $_"
}
