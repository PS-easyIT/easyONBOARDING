# easyONBOARDING 3

Version **3.0.0** · PowerShell 7.2+ · Windows (Oberfläche) · Autor: Andreas Hepp (PHINIT.DE)

easyONBOARDING legt neue Mitarbeitende im Active Directory an, pflegt bestehende Konten und
führt Austritte in mehreren Phasen durch – mit Vorschau, Simulation, Bestätigung, Audit und
Berichten. Version 3 ist eine vollständige Neuentwicklung auf Basis der Versionen 1.3.10 und
1.4.x. Diese liegen unverändert in den Ordnern `V1.3.10` und `V1.4.XX` des Repositorys und
gelten als veraltet.

## Funktionen

| Bereich | Umfang | Status |
|---|---|---|
| Onboarding | Assistent in 8 Schritten, Rollenvorlagen, Kontonamen mit Transliteration und Kollisionsprüfung, Gruppen und Lizenzgruppen, Vorschau, Simulation, einmalige Anzeige des Startkennworts | produktiv |
| Offboarding | 7 Vorlagen (Standardaustritt, Sofortige Sperrung, Befristeter Mitarbeiter, Externer Dienstleister, Ruhestand, Interner Wechsel, Testmodus), Phasen Sofort / Zum Austritt / Aufbewahrung / Endgültige Löschung, Warteschlange, Zustandsberichte vorher/nachher | produktiv (AD, Dateiserver) |
| Benutzer aktualisieren | Attribute, Führungskraft, Gruppen mit Vorher/Nachher-Vorschau; Kennwort-Reset nur als eigene, bestätigte Aktion | produktiv |
| Massenverarbeitung | CSV-Import für Onboarding und Offboarding, Tippbestätigung ab einstellbarer Anzahl | produktiv |
| Berichte und Audit | HTML, JSON, CSV, TXT, PDF (Microsoft Edge, alternativ wkhtmltopdf); Audit-Log als JSON Lines | produktiv |
| Exchange (Online, Server, Hybrid) | Postfach anlegen (Server/Hybrid), Abwesenheitsnotiz, Weiterleitung, Ausblenden, Umwandlung in freigegebenes Postfach, Berechtigungsbericht; Hybrid mit Verbindung zu Exchange Server und Exchange Online (Postfach wird online oder lokal gefunden); Zertifikatsanmeldung für die geplante Aufgabe | vorbereitet¹ |
| Microsoft Graph | Anmeldesitzungen widerrufen (delegiert `User.Read.All`, `User.RevokeSessions.All` oder als Anwendung per Zertifikat) | vorbereitet¹ |
| Entra Connect Sync | Delta-/Initial-Synchronisation per WinRM; Serverprüfung (Erreichbarkeit, laufender Zyklus, Stagingmodus) | vorbereitet¹ |

¹ *Vorbereitet* heißt: implementiert, mit Mocks getestet, per Konfiguration abschaltbar (Standard: aus)
und im Dashboard mit Status angezeigt, aber nicht gegen reale Exchange-, Graph- oder
Entra-Connect-Umgebungen verifiziert. Vor dem
produktiven Einsatz in einer Testumgebung prüfen (siehe [docs/TESTING.md](docs/TESTING.md)).

## Sicherheitsgrundsätze

* **Simulation zuerst:** Neue Vorgänge starten standardmäßig im Simulationsmodus
  (`[Security] SimulationByDefault=1`). Der Live-Modus wird bewusst und mit Bestätigung aktiviert.
* **Keine Kennwörter in Logs, Berichten, CSV-/JSON-Dateien oder E-Mails.** Startkennwörter entstehen
  mit einem kryptografischen Zufallsgenerator, liegen nur als `SecureString` im Speicher und werden
  genau einmal angezeigt (Kopieren mit automatischem Leeren der Zwischenablage, Druck aus dem Speicher).
* **Privilegierte Gruppen und Konten sind gesperrt:** Domain Admins, Enterprise Admins, Schema Admins
  und vergleichbare Gruppen werden nie zugewiesen (Erkennung über SID/RID, Namen, `adminCount` und
  verschachtelte Mitgliedschaft). Geschützte Konten werden nie bearbeitet.
* **Keine automatische endgültige Löschung:** Löschen setzt eine Aufbewahrungsfrist, einen gesicherten
  Zustandsbericht, eine Vorschau und die Eingabe `LÖSCHEN <Konto>` voraus.
* **Keine Klartext-Secrets in der Konfiguration**, kein Ändern der Ausführungsrichtlinie, keine
  pauschalen Administratorrechte, keine automatische Installation von Abhängigkeiten.

Details: [SECURITY.md](SECURITY.md).

## Schnellstart

```powershell
# 1. Voraussetzungen prüfen: PowerShell 7.2+, RSAT-ActiveDirectory
$PSVersionTable.PSVersion
Get-Module -ListAvailable ActiveDirectory

# 2. Konfiguration aus der Vorlage anlegen und anpassen
Copy-Item .\Config\easyONB.ini.template .\Config\easyONB.ini

# 3. Konfiguration und Umgebung prüfen (ohne Oberfläche)
pwsh -NoProfile -File .\Start-easyONBOARDING.ps1 -CheckOnly

# 4. Oberfläche starten (Simulation ist voreingestellt)
pwsh -NoProfile -File .\Start-easyONBOARDING.ps1
```

Heruntergeladene ZIP-Dateien vorher mit `Unblock-File` freigeben. Die Ausführungsrichtlinie wird
nicht verändert; empfohlen ist `RemoteSigned` bzw. `AllSigned` mit signierten Skripten. Ausführlich:
[docs/INSTALLATION.md](docs/INSTALLATION.md).

## Aufbau

```
V3.0.0/
├── Start-easyONBOARDING.ps1     Einstieg: Oberfläche oder -CheckOnly
├── VERSION                      einzige Versionsquelle
├── Config/                      easyONB.ini.template, templates/ (Rollen, Offboarding), companies/, schemas/
├── GUI/                         MainWindow, Views/, Dialogs/, Styles/ (Light/Dark, Controls)
├── Modules/                     Core, Configuration, Security, ActiveDirectory, Exchange, Entra,
│                                Reporting, Onboarding, Offboarding, UI
├── ReportTemplates/             Vorlagen des Willkommensdokuments
├── Assets/                      Bilder für Berichte
├── Scripts/                     geplante Offboarding-Phasen, Migration, Qualitätsprüfung, Doku-Generator
├── Tests/                       Pester (Unit, Integration, Fixtures)
└── docs/                        Installation, Konfiguration, Onboarding, Offboarding, Migration, Tests, Architektur
```

Laufzeitdaten entstehen in `Logs/`, `Reports/` und `Data/` (per `.gitignore` ausgeschlossen, da sie
personenbezogene Daten enthalten).

## Dokumentation

| Dokument | Inhalt |
|---|---|
| [docs/INSTALLATION.md](docs/INSTALLATION.md) | Voraussetzungen, Installation, Rechte, geplante Aufgabe |
| [docs/CONFIGURATION.md](docs/CONFIGURATION.md) | Aufbau der INI, Zusammenführung, Vorlagen, Beispiele |
| [docs/CONFIGURATION-REFERENCE.md](docs/CONFIGURATION-REFERENCE.md) | alle Schlüssel (aus dem Schema erzeugt) |
| [docs/ONBOARDING.md](docs/ONBOARDING.md) | Assistent, Kontonamen, Gruppen, Kennwörter, CSV |
| [docs/OFFBOARDING.md](docs/OFFBOARDING.md) | Vorlagen, Phasen, Warteschlange, Löschschutz |
| [docs/MIGRATION.md](docs/MIGRATION.md) | Umstieg von 1.3.10 / 1.4.x |
| [docs/TESTING.md](docs/TESTING.md) | automatische Tests, CI, manuelle Prüfliste |
| [docs/ARCHITECTURE.md](docs/ARCHITECTURE.md) | Architektur und Entscheidungen |
| [CHANGELOG.md](CHANGELOG.md) | Änderungen |
| [SECURITY.md](SECURITY.md) | Sicherheitsrichtlinie und Meldung von Schwachstellen |
| [CONTRIBUTING.md](CONTRIBUTING.md) | Entwicklung, Konventionen, Qualitätsprüfung |

## Autor und Lizenz

Autor und Maintainer: **Andreas Hepp** (PHINIT.DE, [www.PSscripts.de](https://www.PSscripts.de)).

Im Repository liegt derzeit keine Lizenzdatei. Die Dokumentation der Version 1.4.x nennt die
MIT-Lizenz; die verbindliche Festlegung für Version 3 trifft der Autor.
