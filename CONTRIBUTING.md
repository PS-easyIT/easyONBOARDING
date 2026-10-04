# Mitwirken

Beiträge sind willkommen. Bitte vor größeren Änderungen ein Issue anlegen und das Vorgehen abstimmen.
Sicherheitsrelevante Funde bitte vertraulich melden ([SECURITY.md](SECURITY.md)).

## Entwicklungsumgebung

* PowerShell 7.2 oder höher (Entwicklung und CI mit 7.4). Für die Oberfläche Windows 10/11 oder
  Windows Server 2019 oder höher.
* Pester 5.7.1 und PSScriptAnalyzer 1.24.0 im Benutzerbereich (keine systemweite Installation nötig):

  ```powershell
  Install-PSResource -Name Pester -Version 5.7.1 -Scope CurrentUser
  Install-PSResource -Name PSScriptAnalyzer -Version 1.24.0 -Scope CurrentUser
  ```

* Für Tests wird **kein** Active Directory, Exchange oder Microsoft Graph benötigt: externe Cmdlets
  sind in `Tests/Fixtures/ExternalCommandStubs.ps1` als Stubs definiert, die ohne Mock eine Ausnahme
  auslösen. Tests dürfen keine produktiven Benutzer, Gruppen oder Postfächer verändern.

## Prüfen vor jedem Commit

```powershell
# statische Prüfungen: Parser, Kodierung, XAML, PSScriptAnalyzer, Secrets, Doku, Version
pwsh -NoProfile -File ./Scripts/Invoke-EobQualityCheck.ps1

# Tests (Unit und Integration)
pwsh -NoProfile -Command "Invoke-Pester -Path ./Tests/Unit, ./Tests/Integration -Output Detailed"
```

Die CI (`.github/workflows/v3-ci.yml` im Repository-Stamm) führt dieselben Prüfungen unter Linux und
Windows aus; unter Windows zusätzlich das echte Laden der WPF-Oberfläche ohne Anzeige.

## Konventionen

* **Module:** öffentliche Funktionen mit genehmigtem Verb und Präfix `Eob`
  (`Get-Verb`), Export ausschließlich über das Manifest (`FunctionsToExport`), Hilfe per
  Kommentarblock mit mindestens `.SYNOPSIS`.
* **Strenge:** `Set-StrictMode -Version 3.0` in allen Fachmodulen und Skripten (UI-Modul: 1.0, siehe
  ADR-13), `$ErrorActionPreference = 'Stop'` in Skripten.
* **Systemänderungen** nur in Funktionen mit `[CmdletBinding(SupportsShouldProcess)]`; destruktive
  Aktionen mit `ConfirmImpact = 'High'`. Fachlogik erzeugt Pläne, die Ausführung läuft über die
  Plan-Engine.
* **Keine Secrets:** keine Kennwörter, Tokens oder Verbindungszeichenfolgen im Code, in Beispielen,
  Tests (außer über `New-EobTestSecureString`) oder Berichten. Kennwörter nur als `SecureString`.
* **Beispieldaten:** nur reservierte Domains (`example.com`, `example.local` usw.) und erfundene
  Personen, keine realen Domänen oder personenbezogenen Daten.
* **Oberfläche:** XAML ohne `x:Class` und ohne Ereignisattribute, Farben nur über Ressourcen
  (`DynamicResource`), alle im Code verwendeten Steuerelemente müssen in der XAML existieren
  (geprüft durch `Tests/Unit/UI.Tests.ps1`).
* **Dateien:** UTF-8 mit BOM für `.ps1/.psm1/.psd1`, Einrückung mit 4 Leerzeichen, keine Leerzeichen
  am Zeilenende. Analyzer-Ausnahmen nur per `SuppressMessageAttribute` mit Begründung.
* **Konfiguration:** neue Schlüssel immer im Schema (`Config/schemas/easyONB.schema.psd1`) und in der
  Vorlage anlegen, danach die Referenz neu erzeugen:
  `pwsh -File ./Scripts/Export-EobConfigurationReference.ps1`.

## Versionierung

Semantic Versioning. Die Version steht ausschließlich in `VERSION`; `ModuleVersion` aller Manifeste,
der oberste Eintrag in `CHANGELOG.md` und die README müssen übereinstimmen (geprüft durch
`Invoke-EobQualityCheck.ps1 -Check Version`). Jede nutzerrelevante Änderung erhält einen Eintrag im
CHANGELOG.

## Commits und Pull Requests

* kleine, in sich geschlossene Commits mit aussagekräftiger Nachricht,
* Tests für neues Verhalten und für behobene Fehler,
* im Pull Request beschreiben: Ziel, Änderungen, Auswirkungen auf Sicherheit und Kompatibilität,
  durchgeführte Tests (automatisch und manuell).

## Code-Signatur

Geänderte Skripte verlieren eine vorhandene Authenticode-Signatur. Signiert wird ausschließlich durch
den Maintainer bzw. die betreibende Organisation mit deren Zertifikat.
