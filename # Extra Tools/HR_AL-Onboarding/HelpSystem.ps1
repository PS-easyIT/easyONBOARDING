<#
.SYNOPSIS
    Hilfesystem für easyONBOARDING HR-AL Tool
.DESCRIPTION
    Dieses Modul stellt Hilfefunktionen für das HR-AL Onboarding Tool bereit
.NOTES
    Version: 1.0
#>

# Laden der benötigten Assemblies
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase

# Hilfethemen als Hashtable
$global:helpTopics = @{
    "General" = @{
        "Title" = "Allgemeine Hilfe"
        "Content" = "Das easyONBOARDING HR-AL Tool unterstützt den Onboarding-Prozess neuer Mitarbeiter.`n`nJe nach Ihrer Rolle im System haben Sie unterschiedliche Berechtigungen und Zugriff auf verschiedene Funktionen."
    }
    "HR" = @{
        "Title" = "HR-Funktionen"
        "Content" = "Im HR-Bereich können Sie neue Onboarding-Anfragen erstellen.`n`nErforderliche Felder: Vorname, Nachname, Startdatum und zugewiesener Manager.`n`nNach dem Erstellen wird der zugewiesene Manager benachrichtigt."
    }
    "Manager" = @{
        "Title" = "Manager-Funktionen"
        "Content" = "Als Manager ergänzen Sie die Informationen für Ihre neuen Mitarbeiter.`n`nFüllen Sie die notwendigen Positionsdetails, Abteilung und andere erforderliche Informationen aus.`n`nNach dem Absenden werden Ihre Angaben zur Überprüfung an die HR-Abteilung weitergeleitet."
    }
    "Verification" = @{
        "Title" = "Verifikations-Hilfe"
        "Content" = "Die HR-Verifikation bestätigt die Vollständigkeit und Korrektheit aller Angaben.`n`nNach der Verifikation wird die Anfrage an die IT-Abteilung zur Bearbeitung weitergeleitet."
    }
    "IT" = @{
        "Title" = "IT-Funktionen"
        "Content" = "Der IT-Bereich zeigt alle für die technische Einrichtung vorbereiteten Anfragen.`n`nHier können Sie die Erstellung von Benutzerkonten und die Vorbereitung der Ausstattung bestätigen.`n`nNach Abschluss wird der Onboarding-Prozess als vollständig markiert."
    }
}

# Funktion zum Anzeigen des Hilfefensters
function Show-HelpWindow {
    param(
        [string]$Topic = "General"
    )

    # Wenn das Thema nicht existiert, verwende "General"
    if (-not $global:helpTopics.ContainsKey($Topic)) {
        $Topic = "General"
    }

    # Hilfedaten abrufen
    $helpData = $global:helpTopics[$Topic]
    $title = $helpData.Title
    $content = $helpData.Content

    # WPF-Fenster erstellen
    $helpWindow = New-Object System.Windows.Window
    $helpWindow.Title = "easyONBOARDING Hilfe - $title"
    $helpWindow.Width = 600
    $helpWindow.Height = 400
    $helpWindow.WindowStartupLocation = "CenterScreen"
    $helpWindow.ResizeMode = "CanResize"
    $helpWindow.Background = "#F0F0F0"

    # Layout-Grid
    $grid = New-Object System.Windows.Controls.Grid
    $helpWindow.Content = $grid

    # Zeilen definieren
    $grid.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition -Property @{ Height = "Auto" }))
    $grid.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition -Property @{ Height = "*" }))
    $grid.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition -Property @{ Height = "Auto" }))

    # Spalten definieren
    $grid.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{ Width = "200" }))
    $grid.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{ Width = "*" }))

    # Überschrift
    $header = New-Object System.Windows.Controls.TextBlock
    $header.Text = $title
    $header.FontSize = 20
    $header.FontWeight = "Bold"
    $header.Margin = "10,10,10,10"
    $header.Grid.SetRow($header, 0)
    $header.Grid.SetColumn($header, 0)
    $header.Grid.SetColumnSpan($header, 2)
    $grid.Children.Add($header)

    # Themenbaum
    $topicList = New-Object System.Windows.Controls.ListBox
    $topicList.Margin = "10,5,5,10"
    $topicList.Grid.SetRow($topicList, 1)
    $topicList.Grid.SetColumn($topicList, 0)
    $grid.Children.Add($topicList)

    # Themen hinzufügen
    foreach ($key in $global:helpTopics.Keys) {
        $item = New-Object System.Windows.Controls.ListBoxItem
        $item.Content = $global:helpTopics[$key].Title
        $item.Tag = $key
        $topicList.Items.Add($item)

        # Wenn es das aktuelle Thema ist, auswählen
        if ($key -eq $Topic) {
            $topicList.SelectedItem = $item
        }
    }

    # Inhalt
    $contentBox = New-Object System.Windows.Controls.TextBox
    $contentBox.Text = $content
    $contentBox.IsReadOnly = $true
    $contentBox.TextWrapping = "Wrap"
    $contentBox.VerticalScrollBarVisibility = "Auto"
    $contentBox.Margin = "5,5,10,10"
    $contentBox.Padding = "10"
    $contentBox.Background = "White"
    $contentBox.BorderThickness = "1"
    $contentBox.BorderBrush = "#D0D0D0"
    $contentBox.Grid.SetRow($contentBox, 1)
    $contentBox.Grid.SetColumn($contentBox, 1)
    $grid.Children.Add($contentBox)

    # Schließen-Button
    $closeButton = New-Object System.Windows.Controls.Button
    $closeButton.Content = "Schließen"
    $closeButton.Width = 120
    $closeButton.Height = 30
    $closeButton.Margin = "0,10,10,10"
    $closeButton.HorizontalAlignment = "Right"
    $closeButton.Grid.SetRow($closeButton, 2)
    $closeButton.Grid.SetColumn($closeButton, 1)
    $grid.Children.Add($closeButton)

    # Event-Handler für Themenwechsel
    $topicList.Add_SelectionChanged({
        $selectedItem = $topicList.SelectedItem
        if ($selectedItem -ne $null) {
            $selectedTopic = $selectedItem.Tag
            $helpData = $global:helpTopics[$selectedTopic]
            $contentBox.Text = $helpData.Content
        }
    })

    # Event-Handler für Schließen-Button
    $closeButton.Add_Click({
        $helpWindow.Close()
    })

    # Fenster anzeigen
    $helpWindow.ShowDialog()
}

function Get-HelpContent {
    param(
        [string]$Topic
    )
    
    switch ($Topic) {
        "Allgemein" {
            return @"
HR-AL-Onboarding Tool - Allgemeine Hilfe

Das HR-AL-Onboarding Tool unterstützt den strukturierten Onboarding-Prozess neuer Mitarbeiter.

Workflow-Übersicht:
1. HR erstellt einen neuen Onboarding-Eintrag mit Basisdaten
2. Der Manager des neuen Mitarbeiters ergänzt abteilungsspezifische Informationen
3. HR verifiziert die Vollständigkeit aller Angaben
4. IT richtet das Benutzerkonto und die benötigte Ausstattung ein

Jeder Schritt wird im System dokumentiert und die beteiligten Personen werden automatisch benachrichtigt.

Hilfestellung bei Problemen:
- Stellen Sie sicher, dass Sie mit den korrekten Berechtigungen angemeldet sind
- Bei Fehlermeldungen notieren Sie den genauen Wortlaut
- Wenden Sie sich bei technischen Problemen an die IT-Abteilung
"@
        }
        "HR" {
            return @"
HR-Modul - Hilfe für HR-Mitarbeiter

In diesem Bereich erstellen Sie neue Onboarding-Einträge für Mitarbeiter:

Erforderliche Felder:
- Vorname und Nachname
- Beschreibung der Stelle
- Zugeordneter Manager (wichtig für den Workflow!)

Bei externen Mitarbeitern:
- Aktivieren Sie das Kontrollkästchen "Extern"
- Geben Sie den Namen der Firma an

Workflow:
1. Füllen Sie alle relevanten Felder aus
2. Klicken Sie auf "Speichern", um den Eintrag zu erstellen
3. Der ausgewählte Manager wird automatisch benachrichtigt

Nach Abschluss der Manager-Phase erhalten Sie den Eintrag zur Überprüfung zurück.

Tipps:
- Stellen Sie sicher, dass alle erforderlichen Felder ausgefüllt sind
- Weisen Sie den richtigen Manager zu, da diese Zuordnung später nicht mehr geändert werden kann
- Externe Mitarbeiter werden im System speziell gekennzeichnet
"@
        }
        "Manager" {
            return @"
Manager-Modul - Hilfe für Vorgesetzte

In diesem Bereich ergänzen Sie als Manager die Informationen für den neuen Mitarbeiter:

Erforderliche Felder:
- Position im Unternehmen
- Geschäftsbereich/Abteilung
- Benötigte Software und Zugriffsrechte

Workflow:
1. Sie erhalten eine Benachrichtigung über neue Onboarding-Anfragen
2. Wählen Sie den entsprechenden Eintrag aus der Liste
3. Füllen Sie alle erforderlichen Felder aus
4. Geben Sie besondere Anforderungen im Notizfeld an
5. Klicken Sie auf "Speichern", um den Eintrag zur HR-Verifizierung weiterzuleiten

Zugriffsrechte:
- Sie können nur Einträge bearbeiten, die Ihnen zugewiesen wurden
- Nach dem Absenden sind die Daten für die HR-Verifizierung gesperrt

Tipps:
- Je detaillierter Ihre Angaben, desto schneller kann die IT-Abteilung die Einrichtung vornehmen
- Vergessen Sie nicht, spezielle Software oder Zugriffe anzugeben
"@
        }
        "Verifikation" {
            return @"
Verifikations-Modul - Hilfe für HR-Verifizierung

In diesem Bereich überprüfen und bestätigen Sie als HR-Mitarbeiter die Vollständigkeit der Onboarding-Daten:

Workflow:
1. Prüfen Sie alle eingegebenen Daten auf Vollständigkeit und Korrektheit
2. Setzen Sie das Häkchen bei "Von HR verifiziert", wenn alle Angaben korrekt sind
3. Fügen Sie bei Bedarf Anmerkungen hinzu
4. Klicken Sie auf "Bestätigen", um den Eintrag zur IT-Bearbeitung freizugeben

Besondere Hinweise:
- Die Verifizierung ist ein kritischer Schritt im Workflow
- Ohne Verifizierung kann die IT-Abteilung nicht mit der Einrichtung beginnen
- Bei fehlenden Informationen kontaktieren Sie den zuständigen Manager

Tipps:
- Prüfen Sie besonders die Konsistenz der Angaben (z.B. Abteilung und Manager)
- Stellen Sie sicher, dass alle speziellen Anforderungen klar formuliert sind
"@
        }
        "IT" {
            return @"
IT-Modul - Hilfe für IT-Mitarbeiter

In diesem Bereich führen Sie die technische Einrichtung für neue Mitarbeiter durch:

Checkliste:
- AD-Konto erstellen: Legen Sie das Benutzerkonto nach Namenskonvention an
- Mailbox konfigurieren: Erstellen und konfigurieren Sie die E-Mail-Adresse
- Berechtigungen einrichten: Setzen Sie die angeforderten Zugriffsrechte
- Hardware vorbereiten: Stellen Sie die benötigte Ausstattung bereit
- Software installieren: Installieren Sie die angeforderte Software

Workflow:
1. Wählen Sie einen für IT freigegebenen Eintrag aus der Liste
2. Bearbeiten Sie die Punkte der Checkliste
3. Markieren Sie erledigte Aufgaben als abgeschlossen
4. Geben Sie relevante Hinweise im Notizfeld ein
5. Setzen Sie die Häkchen bei "Konto erstellt" und "Ausstattung bereit"
6. Klicken Sie auf "Abschließen", um den Onboarding-Prozess zu beenden

Tipps:
- Nutzen Sie die Exportfunktion, um Daten für andere IT-Systeme zu exportieren
- Dokumentieren Sie besondere Konfigurationen im Notizfeld
- Bei Unklarheiten wenden Sie sich an den zuständigen Manager oder HR
"@
        }
        "Fehlerbehebung" {
            return @"
Fehlerbehebung - Häufige Probleme und Lösungen

Problem: Validation Error - Assigned manager is required
Lösung: Stellen Sie sicher, dass ein Manager aus der Dropdown-Liste ausgewählt ist.
        Die Manager-Liste sollte automatisch geladen werden. Falls keine Manager
        angezeigt werden, prüfen Sie Ihre Netzwerkverbindung oder wenden Sie sich
        an die IT-Abteilung.

Problem: Property 'HRVerified' cannot be found
Lösung: Dieses Problem betrifft ältere Datensätze. Verwenden Sie die neueste
        Version des Tools oder erstellen Sie einen neuen Datensatz.

Problem: IT checklist control reference is null
Lösung: Dies ist ein Anzeigeproblm und beeinträchtigt die Funktionalität nicht.
        Die IT-Checkliste wird in der nächsten Version korrigiert.

Problem: Login oder Berechtigungsprobleme
Lösung: Stellen Sie sicher, dass Sie die richtigen Berechtigungen für Ihre Rolle haben.
        Wenden Sie sich an Ihren IT-Administrator, um die Berechtigungen zu überprüfen.

Problem: CSV-Export funktioniert nicht
Lösung: Prüfen Sie, ob Sie Schreibrechte im Zielverzeichnis haben.
        Schließen Sie alle anderen Anwendungen, die möglicherweise die CSV-Datei geöffnet haben.

Bei anhaltenden Problemen kontaktieren Sie bitte den IT-Support und geben Sie eine
möglichst genaue Beschreibung des Problems sowie alle Fehlermeldungen an.
"@
        }
        default {
            return @"
HR-AL-Onboarding Tool - Allgemeine Übersicht

Willkommen in der Hilfefunktion des HR-AL-Onboarding Tools!

Dieses Tool unterstützt den strukturierten Onboarding-Prozess für neue Mitarbeiter
durch verschiedene Abteilungen hinweg:

1. HR startet den Prozess mit Basisinformationen
2. Der Manager ergänzt abteilungsspezifische Details
3. HR verifiziert die Vollständigkeit aller Angaben
4. IT richtet das Benutzerkonto und die Ausstattung ein

Wählen Sie ein Thema aus der linken Navigation, um detaillierte Hilfe zu erhalten.

Bei technischen Problemen konsultieren Sie bitte den Abschnitt "Fehlerbehebung"
oder wenden Sie sich an den IT-Support.

Version: 0.1.1
"@
        }
    }
}

# Export der Funktionen
Export-ModuleMember -Function Show-HelpWindow, Get-HelpContent
