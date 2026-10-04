# Änderungsprotokoll

Alle nennenswerten Änderungen an easyONBOARDING. Format angelehnt an
[Keep a Changelog](https://keepachangelog.com/de/1.1.0/), Versionierung nach
[Semantic Versioning](https://semver.org/lang/de/). Einzige Versionsquelle ist die Datei `VERSION`.

## [Unreleased]

### Hinzugefügt

- Installer `Install-easyONBOARDING.ps1` mit eigener WPF-Oberfläche (Modul `easyONB.Setup`): erstellt
  `Config\easyONB.ini` aus der Vorlage. Auf einem Domänencontroller werden Domäne, UPN-Suffixe,
  Domänencontroller, OUs sowie Entra Connect und Exchange aus dem AD vorbelegt; auf Clients werden die
  Werte manuell eingetragen. Jedes Feld und jeder Bereich hat einen Hinweistext (Anzeige nach 250 ms).
  Die erzeugte Datei wird wie in der Anwendung geprüft; vorhandene Dateien werden gesichert.

- Exchange Hybrid: Verbindung zu Exchange Server **und** Exchange Online. Lokale Befehle werden mit dem
  Präfix `EobOnPrem` importiert; Postfachaktionen ermitteln, ob das Postfach online oder lokal liegt.
  Die Umwandlung in ein freigegebenes Postfach erfolgt bei Cloud-Postfächern nach dem
  Microsoft-Verfahren (Exchange Online, danach `Set-RemoteMailbox -Type Shared` lokal).
- Zertifikatsanmeldung (App-only) für Exchange Online (`[Exchange] AppId`, `CertificateThumbprint`,
  `Organization`) und Microsoft Graph (`[Graph] ClientId`, `CertificateThumbprint`). Die geplante
  Aufgabe verbindet damit Exchange Online und Graph unbeaufsichtigt; ohne Zertifikat bleiben diese
  Phasen wie bisher manuell. Neue Schematypen `Guid` und `Thumbprint`, Befund
  `CFG_CERTIFICATE_AUTH_INCOMPLETE` bei unvollständigen Angaben.
- Entra Connect: Serverprüfung über `Get-ADSyncScheduler` (Erreichbarkeit, laufender Zyklus,
  Stagingmodus), in der Oberfläche unter Tools › Entra Connect prüfen.

### Geändert

- Repository neu gegliedert: Die Versionen 1.x und die separaten Zusatzwerkzeuge sind entfernt.
  `README.md`, `CHANGELOG.md`, `SECURITY.md`, `CONTRIBUTING.md` und `VERSION` liegen im
  Repository-Stamm. Qualitätsprüfung, `Get-EobVersion` und CI berücksichtigen den Stamm; findet
  `Get-EobVersion` keine `VERSION`-Datei, gilt die Modulversion. README vollständig überarbeitet.
- `Install-PS7_PDF.ps1` neu geschrieben, läuft auch unter Windows PowerShell 5.1. Das Skript prüft
  zuerst nur: PowerShell 7, ActiveDirectory-Modul, PDF-Erzeugung und Ausführungsrichtlinie.
  Fehlendes wird nur auf Auswahl eingerichtet, über eine WPF-Oberfläche mit Hinweistexten oder
  `-CheckOnly`/`-Install`.
  - PowerShell 7 kommt per winget, sonst als MSI; das MSI nur über https und nur mit gültiger
    Microsoft-Signatur.
  - RSAT wird auf Clients als Windows-Funktion, auf Servern als Windows-Feature installiert.
    wkhtmltopdf nur per winget.
  - Installationen nur in einer Administrator-PowerShell; das Skript erhöht sich nie selbst.
  - Die Ausführungsrichtlinie wird nur nach Bestätigung und nur für den aktuellen Benutzer geändert.
  - Das optionale Kopieren der Anwendung überschreibt keine vorhandene Konfiguration. Tests:
    `Tests/Unit/Prerequisites.Tests.ps1`.
- Oberfläche auf 1400 x 900 ausgelegt: Navigation, Kopf, Abstände, Schriftgrößen und
  Tabellenhöhen so angepasst, dass Assistenten und Ansichten ohne unnötiges Scrollen passen.
  Das Dashboard zeigt die vier Kennzahlen in einer Zeile; in „Benutzer aktualisieren“ stehen die
  Attribute zweispaltig. Hinweistexte in Ansichten, Dialogen und Banner sind gekürzt.
- Auf kleineren Arbeitsflächen (z. B. 1366 x 768) wird das Fenster beim Start auf die
  Arbeitsfläche begrenzt.

### Behoben

- Exchange Hybrid: Postfächer in Exchange Online wurden über die lokale Sitzung nicht gefunden; alle
  Postfachaktionen des Offboardings entfielen dadurch („Kein Postfach gefunden“).
- Microsoft Graph: Der Status meldete „Verbunden“, obwohl `Microsoft.Graph.Users.Actions`
  (`Revoke-MgUserSignInSession`) fehlte; der Schritt scheiterte erst bei der Ausführung.

- Bestätigungsdialog: Lange Detaillisten (z. B. vor der endgültigen Löschung) konnten das Feld für
  die Tippbestätigung aus dem sichtbaren Bereich schieben; lange Zeilen wurden abgeschnitten. Die
  Details scrollen jetzt in einem eigenen Bereich und brechen um.

- Massenverarbeitung: Ein Fehler außerhalb der Zeilenverarbeitung (z. B. beim Aktualisieren der
  Ergebnistabelle) ließ die Ausführungssperre bestehen; Ansichtswechsel und Schließen des Fensters
  waren danach blockiert. Der Stapel wird jetzt geordnet beendet; die Moduswahl ist während der
  Ausführung gesperrt.
- Massenverarbeitung: Kennwörter nicht ausgeführter Zeilen (Fehler, Abbruch, verworfener oder neu
  importierter Stapel) blieben samt Klartext-Redaktionswert bis zum Programmende im Speicher. Sie
  werden jetzt beim Abschluss, beim Moduswechsel, beim Neuimport und beim Beenden verworfen
  (`Invoke-EobOnboardingBatch`: Fehlerzeilen sofort).
- Audit: `Get-EobAuditEntry` liest nur noch die Monatsdateien des angefragten Zeitraums (plus ein
  Monat Puffer) und filtert eine Operation-ID vor dem JSON-Parsen. Das Dashboard las bisher bei
  jedem Aufruf das gesamte Audit-Archiv.
- Audit-Ansicht: Filtertexte mit `[` oder `]` führten zu einem Fehler (Platzhaltermuster); gefiltert
  wird jetzt als Teilzeichenfolge ohne Beachtung der Groß-/Kleinschreibung.
- Offboarding-Warteschlange: Eine leere oder unvollständige JSON-Datei brach das Lesen der gesamten
  Warteschlange ab (Dashboard, Offboarding-Ansicht, geplante Aufgabe). Defekte Einträge werden jetzt
  protokolliert und übersprungen.

## [3.0.0] - 2026-09-27

Vollständige Neuentwicklung im Ordner `V3.0.0`. Die Versionen 1.3.10 und 1.4.x bleiben unverändert
in `V1.3.10` bzw. `V1.4.XX` erhalten und sind veraltet. Die Befund-IDs (C-xx, H-xx, M-xx) beziehen
sich auf die Bestandsaufnahme in [docs/ANALYSIS.md](V3.0.0/docs/ANALYSIS.md).

### Hinzugefügt

- Modulare Architektur mit zehn Modulen (Präfix `Eob`, Manifeste mit expliziten Exporten) und
  Einstieg `Start-easyONBOARDING.ps1` inklusive `-CheckOnly` für Prüfungen ohne Oberfläche.
- Plan-Engine: Jeder Vorgang wird als Plan mit Schritten, Risiken und Befunden erzeugt, in der
  Vorschau angezeigt, simuliert (`-WhatIf`) oder nach Bestätigung ausgeführt; Handler nur aus
  `easyONB.*`-Modulen, Kennwörter nur als Laufzeitverweis.
- Neue WPF-Oberfläche: linke Navigation (Dashboard, Onboarding, Offboarding, Benutzer aktualisieren,
  Massenverarbeitung, Reports und Audit, Tools, Einstellungen, Info), Onboarding-Assistent mit 8
  Schritten, helles und dunkles Farbschema, Akzentfarbe, zentrale Dialoge, Tastenkürzel Strg+1…9,
  schrittweise Ausführung mit Abbruch zwischen zwei Schritten.
- Mehrstufiges Offboarding mit 7 Vorlagen, Phasen (Sofort, Zum Austritt, Aufbewahrung, Endgültige
  Löschung), Warteschlange, Zustandsberichten vorher/nachher, Dokumentation von
  Postfachberechtigungen, Archivierung von Home-Verzeichnissen und abgesicherter Löschung.
- `Scripts/Invoke-DueOffboardingPhases.ps1` für fällige Phasen per geplanter Aufgabe (nie Löschung,
  nie privilegierte Konten).
- Massenverarbeitung per CSV für Onboarding und Offboarding (alle Zeilen, Vorschau je Zeile,
  Tippbestätigung ab `[Bulk] ConfirmationThreshold`).
- Konfigurationsschema mit Typen, Standardwerten, Validierung und Hinweisen; kommentarerhaltendes
  Schreiben; Migration von 1.x-Konfigurationen (Oberfläche und `Scripts/Convert-LegacyConfiguration.ps1`).
- Strukturiertes Logging mit Operation-ID, Akteur, Aktion, Ziel, Ergebnis und Dauer; Audit-Log als
  JSON Lines; Redaktion sensibler Werte vor jedem Schreibvorgang; Ausweichverzeichnis.
- Berichte als HTML, JSON, CSV, TXT und PDF (Microsoft Edge headless, alternativ wkhtmltopdf,
  sonst HTML); Willkommensdokument aus den bisherigen HTML-Vorlagen mit eingebetteten Bildern.
- Integrationsstatus im Dashboard für AD, Exchange, Microsoft Graph, Entra Connect Sync, SMTP,
  Dateiserver und PDF-Erzeugung (Exchange, Graph und Entra Connect: vorbereitet, nicht gegen reale
  Umgebungen verifiziert).
- Pester-Tests (Unit und Integration, externe Systeme ausschließlich über Stubs/Mocks),
  `PSScriptAnalyzerSettings.psd1`, `Scripts/Invoke-EobQualityCheck.ps1` (Parser, Kodierung, XAML,
  PSScriptAnalyzer, Secret-/Platzhalter-Scan, Doku, Version) und GitHub-Workflow
  `.github/workflows/v3-ci.yml` (Windows und Linux, Actions per Commit-SHA, nur Lesezugriff).
- Dokumentation: README, SECURITY, CONTRIBUTING, Installation, Konfiguration (inkl. aus dem Schema
  erzeugter Referenz), Onboarding, Offboarding, Migration, Tests, Architektur.

### Geändert

- Laufzeit: PowerShell 7.2 oder höher. Die Oberfläche benötigt Windows; `-CheckOnly`, Tests und
  Skripte laufen auch unter Linux/macOS.
- Onboarding überschreibt nie ein bestehendes Konto: Kollisionen werden erkannt und je nach
  `[Identity] CollisionStrategy` nummeriert oder als Fehler gemeldet (C-05).
- UPN- und Anzeigenamen-Vorlagen werden korrekt ausgewertet; Legacy-Schlüsselwörter wie
  `FIRSTNAME.LASTNAME` bleiben gültig (F-09).
- "Benutzer aktualisieren" ändert nur tatsächlich geänderte Werte und setzt nie nebenbei ein
  Kennwort (C-04); der Kennwort-Reset ist eine eigene, bestätigte Aktion. Deaktiviert angelegte
  Konten werden dort bei der Übergabe aktiviert.
- AD-Suchen und Filter werden maskiert (H-05); Führungskräfte werden eindeutig aufgelöst (H-12).
- Welcome-Mail über `System.Net.Mail` statt `Send-MailMessage`, Absender einheitlich
  `[EmailSettings] FromAddress` (M-09).
- Berichte sind HTML-kodiert, CSV-Exporte gegen Formel-Injection geschützt (H-06).

### Sicherheit

- Keine Kennwörter mehr in Logs, Debug-Ausgaben, Berichten, CSV-/JSON-Dateien, Willkommensdokumenten
  oder E-Mails (C-01, C-02, C-03, C-07, C-08). Kennwort-Platzhalter in Vorlagen werden nie befüllt.
- Kennwörter mit kryptografischem Zufallsgenerator (H-01); `New-EobPassword` liefert ausschließlich
  `SecureString`, `Test-EobPasswordPolicy` akzeptiert keinen Klartext.
- Privilegierte Gruppen werden nie zugewiesen; Erkennung über SIDs/RIDs, Namen (deutsch/englisch),
  `adminCount` und verschachtelte Mitgliedschaft (H-03). Geschützte Konten (eigenes Konto,
  eingebaute Konten, Notfall-, Dienst- und konfigurierte Konten/OUs) werden nie bearbeitet.
- Live-Ausführungen verlangen eine Bestätigung (`ConfirmImpact High`), destruktive Aktionen eine
  Tippbestätigung; Simulation ist Standard.
- Keine Selbst-Elevation und kein `-ExecutionPolicy Bypass` mehr (H-09); keine Downloads ohne
  Prüfung (H-10, Installer nicht Bestandteil von 3.0).
- Graph mit minimalen Scopes `User.Read.All` und `User.RevokeSessions.All` statt
  `Directory.ReadWrite.All` (M-08).
- PDF-Erzeugung ohne `--enable-local-file-access` (wkhtmltopdf); Microsoft Edge läuft headless mit
  temporärem Profil, ohne Erweiterungen und ohne Hintergrundnetzwerkzugriffe.

### Entfernt

- Fester Standardkennwortwert `fixPassword` (C-06) und Klartext-Secrets in der Konfiguration
  (`CompanyVPNPassword`, `JiraToken`, SMTP-Kennwort, Webhooks; H-11): werden erkannt, gemeldet
  und ignoriert.
- Frei konfigurierbarer `SyncCommand` (Codeausführung über die Konfiguration, H-04); Entra Connect
  Sync startet nur noch `Start-ADSyncSyncCycle`.
- Zusatzkennwörter `{{CustomPW1..5}}` und die Graph-Lizenzverwaltung (gruppenbasierte Lizenzierung
  über `[LicensesGroups]`).

### Veraltet

- Versionen 1.3.10 und 1.4.x (siehe `DEPRECATED.md` in den jeweiligen Ordnern).
- Legacy-Konfigurationsschlüssel werden weiter gelesen und im Prüfbericht mit Ersatz genannt
  (Übersicht: [docs/CONFIGURATION-REFERENCE.md](V3.0.0/docs/CONFIGURATION-REFERENCE.md)).

### Bekannte Einschränkungen

- Die Oberfläche wurde automatisiert nur statisch (XAML-Prüfung, Abgleich Code/XAML) und in der CI
  unter Windows ohne Anzeige geladen; ein manueller Test nach [docs/TESTING.md](V3.0.0/docs/TESTING.md) ist
  vor dem produktiven Einsatz erforderlich.
- Exchange-, Graph- und Entra-Connect-Funktionen sind nicht gegen reale Umgebungen verifiziert.
- Die Dateien der Version 3 sind nicht Authenticode-signiert (Zertifikat des Maintainers erforderlich).

## Frühere Versionen

- **1.4.x** (Entwicklungsstand 1.4.23, Ordner `V1.4.XX`): nicht als final freigegeben; der
  GUI-Onboarding-Pfad ist defekt (siehe Bestandsaufnahme).
- **1.3.10** (2025-03-30, Ordner `V1.3.10`): letzte als final markierte Version der Reihe 1.x.
