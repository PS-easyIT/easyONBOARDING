# Installation und Betrieb

## Voraussetzungen

| Komponente | Anforderung | Hinweis |
|---|---|---|
| Betriebssystem | Windows 10/11 oder Windows Server 2019+ | für die Oberfläche (WPF); `-CheckOnly` und Tests auch unter Linux/macOS |
| PowerShell | 7.2 oder höher | [Installation](https://learn.microsoft.com/powershell/scripting/install/installing-powershell-on-windows) durch die IT |
| Active Directory | Modul `ActiveDirectory` (RSAT) | Windows 10/11: `Add-WindowsCapability -Online -Name Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0`; Server: `Install-WindowsFeature RSAT-AD-PowerShell` |
| PDF (optional) | Microsoft Edge, alternativ wkhtmltopdf | ohne PDF-Engine werden HTML-Berichte erzeugt |
| Exchange Online (optional) | Modul `ExchangeOnlineManagement` | interaktive Anmeldung |
| Exchange Server/Hybrid (optional) | Remote-PowerShell-Endpunkt, Kerberos | `[Exchange] OnPremisesUri` |
| Microsoft Graph (optional) | `Microsoft.Graph.Authentication`, `Microsoft.Graph.Users.Actions` | Scopes `User.Read.All`, `User.RevokeSessions.All` |
| Entra Connect Sync (optional) | WinRM zum Synchronisationsserver | Mitgliedschaft in `ADSyncOperators` auf dem Server |

easyONBOARDING installiert keine Module und ändert keine Systemeinstellungen. Fehlende Komponenten
zeigt das Dashboard bzw. `-CheckOnly` mit Hinweis an. Optionale Module installiert die IT bei
Bedarf gezielt, z. B. mit `Install-PSResource <Name> -Scope CurrentUser`.

## Installation

1. Den Ordner `V3.0.0` aus dem Repository bzw. Release in ein Verzeichnis kopieren, auf das nur die
   IT-Administration Schreibzugriff hat, z. B. `D:\Tools\easyONBOARDING\V3.0.0`.
2. Nach dem Download einer ZIP-Datei die Dateien freigeben:
   `Get-ChildItem -Path D:\Tools\easyONBOARDING -Recurse | Unblock-File`.
3. Ausführungsrichtlinie: easyONBOARDING ändert sie nicht und benötigt kein `Bypass`. Empfohlen ist
   `RemoteSigned` (lokale, freigegebene Skripte) bzw. `AllSigned`, wenn die Skripte mit dem
   Codesignaturzertifikat der Organisation signiert werden.
4. Konfiguration anlegen: `Copy-Item .\Config\easyONB.ini.template .\Config\easyONB.ini` und anpassen
   ([CONFIGURATION.md](CONFIGURATION.md)). Alternativ eine Konfiguration der Version 1.x übernehmen
   ([MIGRATION.md](MIGRATION.md)).
5. Prüfen: `pwsh -NoProfile -File .\Start-easyONBOARDING.ps1 -CheckOnly` (Exitcode 0 = keine Fehler,
   2 = Fehler in der Konfiguration). Warnungen und Hinweise beachten.
6. Starten: `pwsh -NoProfile -File .\Start-easyONBOARDING.ps1`. Eine Verknüpfung kann auf
   `pwsh.exe -NoProfile -File "D:\Tools\easyONBOARDING\V3.0.0\Start-easyONBOARDING.ps1"` zeigen.

Parameter von `Start-easyONBOARDING.ps1`: `-ConfigPath` (sonst `Config\easyONB.ini` bzw. Umgebungsvariable
`EASYONB_CONFIG`), `-CheckOnly`, `-Simulation` (erzwingt den Simulationsmodus), `-Theme Light|Dark`.

## Berechtigungen

easyONBOARDING benötigt **keine lokalen Administratorrechte** und **keine Domain-Admin-Rechte**. Es
arbeitet mit den Rechten des angemeldeten Kontos. Empfohlen ist ein eigenes IT-Konto mit delegierten
Rechten (Assistent "Objektverwaltung zuweisen" bzw. `dsacls`), beschränkt auf die betroffenen OUs:

| Aufgabe | Benötigte Rechte |
|---|---|
| Onboarding | Benutzerobjekte in den Ziel-OUs erstellen; Eigenschaften der angelegten Benutzer schreiben; Erweitertes Recht *Kennwort zurücksetzen*; `pwdLastSet` schreiben (Kennwortänderung bei Anmeldung) |
| Gruppen zuweisen/entfernen | *Mitglieder schreiben* (`member`) nur auf den verwendeten Gruppen |
| Benutzer aktualisieren | Eigenschaften schreiben (Attribute, `manager`); *Kennwort zurücksetzen* nur für die Reset-Funktion |
| Offboarding | `userAccountControl`, `accountExpires`, `logonHours`, `description`, `manager` schreiben; *Kennwort zurücksetzen*; Verschieben: Löschen in der Quell-OU und Erstellen in `[Offboarding] DisabledUsersOU` |
| Endgültige Löschung | Benutzerobjekte in der Ausgeschieden-OU löschen – nur für die Personen, die Löschungen freigeben |
| Home-Verzeichnisse | NTFS-Rechte unter `[FileServer] AllowedRoots` bzw. `ArchiveRoot` (Ordner anlegen, verschieben, Berechtigungen setzen) |
| Exchange | Server/Hybrid: Rollen *Mail Recipient Creation* und *Mail Recipients* (z. B. über die Rollengruppe *Recipient Management*); Online: Entra-Rolle *Exchange Recipient Administrator* oder eine gleichwertige RBAC-Rolle |
| Microsoft Graph | delegierte Berechtigungen `User.Read.All` und `User.RevokeSessions.All` mit Administratorzustimmung |
| Entra Connect Sync | WinRM-Zugriff und Mitgliedschaft in der lokalen Gruppe `ADSyncOperators` auf dem Synchronisationsserver |

Privilegierte Gruppen und geschützte Konten bearbeitet easyONBOARDING unabhängig von den Rechten nie
([SECURITY.md](../SECURITY.md)).

## Datenablage und Schutz

| Verzeichnis | Inhalt | Empfehlung |
|---|---|---|
| `Logs\` | Textlogs (täglich, `easyONB_yyyyMMdd.log`), Rotation nach `[Logging] MaxFileSizeMB` | nur IT lesend/schreibend |
| `Logs\audit\` | Audit-Log (monatlich, `easyONB_audit_yyyyMM.jsonl`) | nur IT; Aufbewahrung nach `[Logging] AuditRetentionDays` |
| `Reports\` | Vorgangsberichte, Willkommensdokumente, Zustandsberichte (`Reports\Offboarding`) | nur IT; enthält personenbezogene Daten |
| `Data\OffboardingQueue\` | Warteschlange des Offboardings (JSON je Vorgang, keine Secrets) | nur IT |

Pfade lassen sich in `[Logging]`, `[Report]` und `[Paths]` ändern. Ist ein Logverzeichnis nicht
beschreibbar (oder das konfigurierte Laufwerk nicht vorhanden), verwendet easyONBOARDING
`%LOCALAPPDATA%\easyONBOARDING\Logs` und zeigt dies in der Statusleiste an. Live-Ausführungen setzen
ein beschreibbares Audit-Log voraus.

## Geplante Aufgabe für Folgephasen des Offboardings

Phasen wie "Zum Austritt" oder "Aufbewahrung" werden zum Fälligkeitstag entweder in der Oberfläche
(Offboarding › Warteschlange) oder automatisch ausgeführt:

```powershell
pwsh.exe -NoProfile -NonInteractive -File "D:\Tools\easyONBOARDING\V3.0.0\Scripts\Invoke-DueOffboardingPhases.ps1"
```

* Konto: am besten ein gruppenverwaltetes Dienstkonto (gMSA) mit den delegierten Offboarding-Rechten
  und dem Recht "Als Batchauftrag anmelden"; keine Domain-Admin-Rechte.
* Zeitplan: täglich, z. B. 06:00 Uhr. Zuerst mit `-WhatIf` testen.
* Das Skript löscht **nie** Konten, bearbeitet **keine** privilegierten Konten und führt Phasen mit
  Exchange-/Graph-Aktionen nur mit `-AllowIntegration` und bestehender Verbindung aus (Exchange Server
  per Kerberos; Exchange Online und Graph erfordern eine interaktive Anmeldung und werden daher als
  "manuell erforderlich" gemeldet).
* Exitcodes: 0 = in Ordnung, 1 = mindestens ein Vorgang blockiert oder fehlgeschlagen,
  2 = Konfiguration oder AD-Verbindung fehlerhaft. Details stehen im Log und je Vorgang im Bericht.

## Aktualisieren und Entfernen

* **Aktualisieren:** neuen Versionsordner neben den alten legen, `Config\easyONB.ini`, `Config\companies`
  und eigene Vorlagen übernehmen, `-CheckOnly` ausführen. `Logs`, `Reports` und `Data` können
  übernommen oder über die Konfiguration an einen festen Ort gelegt werden.
* **Entfernen:** geplante Aufgabe löschen, Ordner entfernen. Logs, Berichte und Warteschlange gemäß
  den Aufbewahrungsvorgaben der Organisation archivieren oder löschen.
