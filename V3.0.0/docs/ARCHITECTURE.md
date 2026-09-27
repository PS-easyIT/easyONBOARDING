# Architektur und Architekturentscheidungen

Alle Pfade in diesem Dokument sind relativ zum Versionsordner `V3.0.0`.

## Überblick

```
Start-easyONBOARDING.ps1          Einstieg (GUI, -CheckOnly)
│
├── Modules/easyONB.Modules.ps1   Ladereihenfolge der Module
├── Modules/Core                  Version, Pfade, Logging/Audit, Redaktion, Plan-Engine
├── Modules/Configuration         INI-Parser, Schema, Validierung, Companies, Vorlagen
├── Modules/Security              Kennwörter, Richtlinien, geschützte Gruppen/Konten, Escaping
├── Modules/ActiveDirectory       AD-Adapter (Lesen, Schreiben mit ShouldProcess)
├── Modules/Exchange              optionale Postfachaktionen (Online/On-Premises)
├── Modules/Entra                 optionale Graph-/Entra-Connect-Funktionen
├── Modules/Reporting             Report-Modell, HTML/CSV/JSON/TXT/PDF, Legacy-Templates, Mail
├── Modules/Onboarding            Identitäten, Anfrage, Plan, Ausführung, CSV, Benutzer-Update
├── Modules/Offboarding           Vorlagen, Schutzprüfung, Snapshot, Plan, Phasen, Löschschutz, CSV
└── Modules/UI                    WPF: XAML-Laden, Themes, Dialoge, Views, kooperative Ausführung
```

Abhängigkeiten zeigen nur "nach unten" (Core ist die Basis). Fachmodule rufen AD, Exchange und
Entra ausschließlich über die Adapterfunktionen auf, damit sie in Pester vollständig gemockt
werden können.

## Datenfluss (Onboarding und Offboarding)

```
Formular/CSV → Request → Test-*Request (Feldfehler)
            → New-*Plan  (nur lesende AD-Abfragen, Kollisions- und Schutzprüfung)
            → Vorschau   (Schritte, Risiken, Warnungen, Vorher-Zustand)
            → Bestätigung (Checkbox bzw. Tippbestätigung bei destruktiven Aktionen)
            → Ausführung (Invoke-EobPlanStep je Schritt, -WhatIf in Simulation)
            → Report + Audit
```

## Entscheidungen (ADR)

**ADR-01 – Neue Anwendung im Versionsordner `V3.0.0`, Legacy unverändert.** Die Anwendung folgt der
bestehenden Konvention (ein Ordner je Version). Die signierten Legacy-Ordner bleiben als Rückfallebene
erhalten und werden als veraltet markiert. Begründung: Monolith in Kernpfaden defekt, kritische
Sicherheitsmängel, Signaturen. Nur `.github/workflows` liegt technisch bedingt im Repository-Root.

**ADR-02 – Module mit Präfix `Eob`.** Alle öffentlichen Funktionen verwenden genehmigte Verben und
das Präfix `Eob`, um Kollisionen mit AD-/Exchange-Cmdlets und Legacy-Funktionen zu vermeiden.
Jedes Modul besitzt ein Manifest (`.psd1`) mit expliziten Exporten.

**ADR-03 – Eine Versionsquelle.** `V3.0.0/VERSION`. Manifeste und `CHANGELOG.md` werden per Test
gegen diese Datei geprüft.

**ADR-04 – INI bleibt Primärformat.** Bestehende INI-Dateien werden weiter gelesen. Ein Schema
(`Config/schemas/easyONB.schema.psd1`) liefert Typen, Standardwerte, Pflichtangaben und Hinweise.
Zusammenführung: `Config/templates/*.ini` → `Config/companies/*.ini` → Haupt-INI (höchste Priorität).
Unbekannte Schlüssel werden gemeldet, aber nie verworfen. Secrets in der INI werden erkannt und
ignoriert. Hot-Reload wird nicht angeboten; "Neu laden" validiert vor der Übernahme.

**ADR-05 – Plan-Engine mit Handler-Allowlist.** Pläne bestehen aus Schritten mit Handlerfunktion,
Parametern, Risiko, Abhängigkeiten und Phase. Die Engine ruft nur Funktionen aus
`easyONB.*`-Modulen auf. Secrets werden nicht als Parameter gespeichert, sondern über
`@Secret:<Name>`-Verweise zur Laufzeit aus dem Speicher aufgelöst.

**ADR-06 – Kennwörter.** Erzeugung mit `RandomNumberGenerator` (CSPRNG). Im Speicher als
`SecureString`. Klartext nur für die optionale einmalige Anzeige bzw. den Druck aus dem Speicher.
Generierte Werte werden für die Laufzeit in die Log-Redaktion eingetragen. Ein fester
Standardkennwortwert aus der INI wird nicht mehr unterstützt.

**ADR-07 – Logging.** Textlog pro Tag plus Audit-Log (JSON Lines, monatlich). Pflichtfelder:
Zeitstempel, Operation-ID, Akteur, Aktion, Ziel, Ergebnis, Dauer, Level. Redaktion vor jedem
Schreibvorgang. Prozessübergreifende Sperre per benanntem Mutex. Schreibfehler werfen nicht, werden
aber gezählt und in GUI/Konsole angezeigt. Vor Live-Ausführungen wird die Schreibbarkeit des
Audit-Logs geprüft.

**ADR-08 – WPF ohne Fremdbibliotheken.** Themes als ResourceDictionaries in `Application.Resources`,
Farben ausschließlich über `DynamicResource`. Views als eigene XAML-Dateien (UserControl), Dialoge als
eigene Fenster. Controls werden pro View deklariert und beim Laden vollständig aufgelöst (fehlende
Controls → sichtbarer Fehler statt stiller `$null`-Werte).

**ADR-09 – Kooperative Ausführung statt Hintergrund-Runspaces.** Pläne werden Schritt für Schritt über
einen `DispatcherTimer` im UI-Runspace ausgeführt. Zwischen den Schritten aktualisiert sich die
Oberfläche, Abbrechen ist nur dort möglich (sichere Stellen). Begründung: Exchange-/Graph-Sitzungen
sind runspace-gebunden, Modulimporte in Hintergrund-Runspaces sind langsam, und die Fehleranfälligkeit
von Threadübergängen in PowerShell-WPF ist hoch.

**ADR-10 – Schutzmechanismen SID-basiert.** Privilegierte Gruppen werden über bekannte SIDs/RIDs
erkannt (sprachunabhängig), ergänzt um Namensmuster, `adminCount` und verschachtelte Mitgliedschaft
(`LDAP_MATCHING_RULE_IN_CHAIN`). Geschützte Konten: ausführendes Konto, eingebaute Konten
(RID 500–504), Break-Glass-, Service- und konfigurierte Konten bzw. OUs.

**ADR-11 – Mehrstufiges Offboarding über eine Warteschlange ohne ausführbaren Inhalt.** Gespeichert
werden nur Anfrage-Daten (ObjectGUID, Vorlage, Termine, Optionen). Fällige Phasen werden neu geplant,
neu validiert und erneut in der Vorschau bestätigt. Die endgültige Löschung verlangt Frist, Snapshot
mit Hash-Prüfung, Tippbestätigung und läuft nie unbeaufsichtigt.

**ADR-12 – PDF über vorhandene Komponenten.** Bevorzugt Microsoft Edge (headless), alternativ
wkhtmltopdf (Legacy, konfigurierbar), sonst HTML-Fallback mit Hinweis.

**ADR-13 – StrictMode.** `Set-StrictMode -Version 3.0` in allen Fachmodulen. Im UI-Modul
`-Version 1.0`, weil WPF-Objekte lose typisiert sind und die GUI nicht lokal automatisiert getestet
werden kann; UI-Fehler sollen nicht durch Property-Zugriffe auf optionale WPF-Eigenschaften entstehen.

**ADR-14 – Laufzeit.** PowerShell 7.2 oder höher (primäres und einziges getestetes Ziel), Windows
für GUI und AD-Funktionen. Keine lokalen Administratorrechte erforderlich; AD-Rechte werden über
Delegation vergeben.
