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

# SIG # Begin signature block
# MIIbywYJKoZIhvcNAQcCoIIbvDCCG7gCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAZdcAeOJEjt7VE
# 7YzmGd5+IyJ8IoOr2QQMESYLr+x+2KCCFhcwggMQMIIB+KADAgECAhB3jzsyX9Cg
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
# DQEJBDEiBCB/7pR/y/2KPWfEzGcmKjHIJ0OctCFUmMMiB1ZVSF2fJTANBgkqhkiG
# 9w0BAQEFAASCAQCnK8L/mFXAH0np+Zh8PzkmUjpN/Bo21zlj0M+CNcbIAWXG1e56
# SU2glF+Yc+hEgHEUd38YXEsfSJUl+tfiUcmdBrFSbPJ7j0lguI0NKnAh2ukC3+b9
# DumB5WUc9YegqGYQpdnzorhopWhmZFXxThpgN5U80uQvpk+U3R+3I/BzODBJBQwW
# HbEMd0Nih7iW9aibUC+YH61B1UXPYxtyvK7sySf21OpC4N+QdHTbEHsUvGIPwGK2
# kaEVvHvMbaH7sxNEiPPjkzg06y0Nahh9d3EaoZjgOtp4PhFlckdR2BTFpAJqwVqT
# xJFKDFq+wx3aNXJax2JErm3tG2U9IAb4PvOroYIDIDCCAxwGCSqGSIb3DQEJBjGC
# Aw0wggMJAgEBMHcwYzELMAkGA1UEBhMCVVMxFzAVBgNVBAoTDkRpZ2lDZXJ0LCBJ
# bmMuMTswOQYDVQQDEzJEaWdpQ2VydCBUcnVzdGVkIEc0IFJTQTQwOTYgU0hBMjU2
# IFRpbWVTdGFtcGluZyBDQQIQC65mvFq6f5WHxvnpBOMzBDANBglghkgBZQMEAgEF
# AKBpMBgGCSqGSIb3DQEJAzELBgkqhkiG9w0BBwEwHAYJKoZIhvcNAQkFMQ8XDTI1
# MDcwNTEwMTIwM1owLwYJKoZIhvcNAQkEMSIEIN+lUPsNn5uB0o8cZDhmNrCInZGg
# oAIJeahrNsFCpYIFMA0GCSqGSIb3DQEBAQUABIICAD4loznFvLrHtSKeAvuPuDPG
# O18GYjArascPdWkg/E2DziwHIfjKMUkALalp/2tgOnbZAsuiRjPZr583l5BFbhaR
# Sz1AFWz/luKzsexBZAdjmt34A7EDadvISof83qMEbBORTBQF+P3c7KfsdHFx8ZVe
# VLkzwwyyWmF0jdF8zPoKOWpou0w61+jKQOMCEPIhNLkDIYGnqRJLXR/SiqC7isKi
# qhFtuazbJe3c5FpkFviW5YUaiwMI1pmMZWx/YQzwmC2QBhODZV80duOOiPzIvdYu
# ew42N9ewp5orfcHb+pHhIt6UljQ8hoFIDCILcPV6DySrasuiyqh/mE6yE0QRq9rf
# 9pXGx5W6EXXd83Bsz9CvPlBIZoFgDhvpyMRS5MK+cidpH6cNq1XhQdR8E1gIP9Pc
# 2jwVbBaqJ+Wl26xrM+YJubR9SUx+D9TiQIwm1kiG1UmXWyYbVJroFazLBOw55xUH
# god0XhvW6gtBz4fm9tVGERN+BnFG5bWGeIFy5Joy5L+zU4rPjskuTJMWnUy4Tx0O
# 7HPdJrU1sNrQeR9QgrORUhDmVKBKWKyQYPuSL3j3ivCiertg3K0S02avADVh8ZK+
# AjV+qV7NvZF5BXow2ROQW9s0kcSZLYQj+sqIfwVcwRFWVGUzkEU8bzWX5BeY5dRu
# Xs5Rk4lgdXJ/gHkhnULD
# SIG # End signature block
