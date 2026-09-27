# Bestandsaufnahme easyONBOARDING (Phase A)

Stand der Analyse: 27.09.2026, Basis-Commit `e395922` ("Code Signing").
Analysiert wurde der tatsächliche Repository-Inhalt. Aussagen aus README-Dateien wurden gegen den Code geprüft.

Klassifizierung: **Critical**, **High**, **Medium**, **Low**, **Recommendation**.

---

## 1. Ermittelte Version

| Fundstelle | Angabe |
|---|---|
| `README.md` (Root) | 1.3.10, "Last Update 2025-03-30" |
| `V1.3.10/easyONBOARDING_V1.3.10.ps1` | 1.3.10 (Dateiname), Commit bd6cef8 "Final Version 1.3.10" |
| `V1.4.XX/easyONBOARDING_V1.4.23.ps1` | 1.4.23 (nur Dateiname) |
| `V1.4.XX/README.md`, `commingFeatures.md`, `IMPLEMENTATION_SUMMARY.md` | 1.4.1 |
| `V1.4.XX/assets/easyONB.ini` `[ScriptInfo] ScriptVersion` | 1.3.10 |
| `V1.4.XX/assets/easyONB.ini` `[WPFGUI] FooterText` | 1.3.9 / 22.03.2025 |
| Modulköpfe `V1.4.XX/Modules/*` | "1.4.XX" bzw. "1.0.0" |
| Git-Tags / Releases | keine |

**Ergebnis:** Letzte als final markierte Version ist **1.3.10**. `V1.4.XX` ist ein
Entwicklungsstand (1.4.23), dessen GUI-Workflow nicht funktionsfähig ist (siehe F-02).
Es gab keine zentrale Versionsquelle. Neue Source of Truth: Datei `V3.0.0/VERSION` (3.0.0),
geprüft durch einen Pester-Test gegen Modulmanifeste und `V3.0.0/CHANGELOG.md`.

## 2. Tatsächliche Projektstruktur (vor der Modernisierung)

```
/                       README.md (Version 1.3.10), keine LICENSE, keine .gitignore, keine Tests, keine CI
├── V1.3.10/            letzte finale Version: Monolith-Skript (≈5.500 Zeilen), MainGUI.xaml, easyONB.ini,
│                       easyINIEditor.ps1, PDFCreator.ps1, INSTALL-PS7_PDF.ps1, CSVGenerator/, ReportTemplates/
├── V1.4.XX/            Entwicklungsstand 1.4.23
│   ├── easyONBOARDING_V1.4.23.ps1   Monolith (6.349 Zeilen inkl. Signaturblock)
│   ├── MainGUI.xaml, MainGUI_optimized.xaml   leere Dateien (0 Byte)
│   ├── assets/MainGUI.xaml          tatsächlich geladene GUI (1.362 Zeilen)
│   ├── assets/MainGUI_backup.xaml   kein gültiges XML (Zeile 190)
│   ├── assets/easyONB.ini           Konfiguration (347 Zeilen)
│   ├── Modules/                     11 .psm1 + 3 .ps1
│   ├── ReportTemplates/             HTMLTemplate(.txt/ENG), rep.html (Duplikat), txtTemplate.txt
│   ├── TestImportCSVFiles/          6 CSV-Dateien mit je 1 Datensatz
│   └── Logs/easyOnboarding.log      eingecheckte Laufzeitdatei
├── # ARCHIV/           Versionen 0.1 – 1.3.x inkl. .exe/.zip-Binärdateien
├── # Extra Tools/      CSVGenerator, HR_AL-Onboarding (2 Dateien mit Syntaxfehlern), CheckScripts
├── # TEST/WEB Edition/ IIS-Prototyp (2 Dateien mit Syntaxfehlern)
├── # INSTALL/, # PDFCreator/, # Screenshots/
```

Alle `.ps1/.psm1` in `V1.3.10`, `V1.4.XX` und den Extra-Tools sind **Authenticode-signiert**
(Zertifikat "PhinIT-PSscripts_Sign"). Jede inhaltliche Änderung macht die Signatur ungültig.

## 3. Einstiegspunkt und Laufzeitverhalten (1.4.23)

* Einstieg: `V1.4.XX/easyONBOARDING_V1.4.23.ps1` (`#requires -Version 7.0`).
* Fordert beim Start Administratorrechte an und startet sich mit `-ExecutionPolicy Bypass` neu.
* Liest `assets/easyONB.ini`, lädt Module, lädt `assets/MainGUI.xaml`, registriert Handler, `ShowDialog()`.
* Die gesamte Logik (UPN, Onboarding, Reports, AD-Update, CSV, Navigation) liegt in einer Datei.

## 4. Module in `V1.4.XX/Modules` – Ist-Zustand

| Modul | Status laut Code | Ursache |
|---|---|---|
| SettingsManager, ADManager, UIManager | **nicht funktionsfähig** | rufen `Write-LogMessage` auf, das nur im Skript-Scope des Hauptskripts existiert; `Initialize-SettingsManager` wird mit nicht existierendem Parameter `-LogFunction` aufgerufen, `Initialize-ADManager` existiert nicht |
| ExtendedFunctions | lädt | `Send-WelcomeEmail` versendet Klartextkennwort; `-contains` statt `-match` bei E-Mail-Prüfung |
| ExtendedFunctions2 | **nicht funktionsfähig** | nutzt `Write-ModuleLog`, das dort nicht definiert ist; `Set-UserPhoto` ruft sich rekursiv selbst auf; `"$Identity_$timestamp"` erzeugt falschen Pfad |
| ReportingModule, ValidationModule, NotificationModule, LicenseModule | **werden nie geladen** | `$scriptPath` ist undefiniert → `Join-Path` wirft → gesamter Modul-Ladeblock bricht ab (im eingecheckten Log belegt) |
| SettingsGUI / SettingsGUI2 | lädt teilweise | `Save-IniContent` verwirft alle Kommentare der INI |
| easyINIEditor.ps1, PDFCreator.ps1 | nicht erreichbar | Hauptskript sucht sie in `$PSScriptRoot` statt in `Modules/` |

Namenskollisionen: `Send-WelcomeEmail` (2×), `Test-UserInput` (2×), `New-UserReport` (2×),
`Get-LicenseUsageReport` (2×), `Get-IniContent` (2×), `Update-StatusText` (2×), `Get-ADConnectionStatus` (2×).

## 5. Funktionen: vorhanden vs. nur dokumentiert

| Funktion laut Doku | Tatsächlicher Stand 1.4.23 |
|---|---|
| AD-Benutzeranlage | Code vorhanden, **GUI-Pfad defekt** (F-02) |
| UPN-Templates aus INI | **wirkungslos**: `New-UPN` liest `DefaultDisplayNameFormat` statt `DefaultUserPrincipalNameFormat`; Vorlagen mit `_` erzeugen falsche Werte (F-09) |
| CSV-Massenimport | **nur erster Datensatz** wird in das Formular übernommen; keine Massenverarbeitung |
| Benutzer aktualisieren | Handler registrieren sich nicht (Control-Namen fehlen in XAML); Code setzt bei jedem Update zwangsweise ein neues Kennwort (F-04) |
| PDF-Reports | wkhtmltopdf-Aufruf über `powershell.exe -ExecutionPolicy Bypass`, Skriptpfad falsch |
| Reports-/Tools-Panel | nur Onboarding-Summary (liest Logdatei) und Lizenz-Report; Rest "In Entwicklung" |
| Settings-GUI mit Hot-Reload, Live-Validierung | nicht vorhanden / nicht angebunden |
| Welcome-Mail, Postfach, Home-Verzeichnis, Profilbild, Backup, AD-Sync | Code vorhanden, Module teils nie geladen, teils defekt |
| Lizenzverwaltung per Graph | Modul wird nie geladen; fordert `Directory.ReadWrite.All` |
| Offboarding | **nicht vorhanden** |
| Tests / CI | **nicht vorhanden** |
| LICENSE (MIT laut README) | **keine Lizenzdatei vorhanden** |

## 6. Befunde

### Critical

| ID | Befund | Fundstelle |
|---|---|---|
| C-01 | Generiertes Kennwort im Debug-Log im Klartext (`Write-DebugMessage "Generated password: …"`) | 1.4.23 Z. 834; 1.3.10 Z. 595 |
| C-02 | Kennwort wird mit `Write-LogMessage … Password: '$mainPW'` in die Logdatei geschrieben (Parameter `-Level` wird bei der einfachen Funktion stillschweigend ignoriert, Eintrag landet als INFO im Log) | 1.4.23 Z. 3086; 1.3.10 Z. 2940 |
| C-03 | Kennwörter (`{{Passwort}}`, `{{CustomPW1..5}}`, `{{CompanyVPNPassword}}`) werden in HTML-/TXT-Reports auf Platte geschrieben; PDF-Export übernimmt sie | 1.4.23 Z. 3223–3268, 3315–3363, 5372 |
| C-04 | "Benutzer aktualisieren" setzt bei **jeder** Attributänderung ein neues Kennwort und schreibt es in einen HTML-Report | 1.4.23 Z. 5222–5419 |
| C-05 | Onboarding überschreibt einen **bestehenden** Benutzer mit gleichem SamAccountName (Attribute, Aktivierungsstatus, Kennwort-Reset, Gruppen) ohne Rückfrage | 1.4.23 Z. 2816–2847 |
| C-06 | Fester Standard-Kennwortwert `fixPassword=P@ssw0AHrd!` im Repository und als Betriebsmodus vorgesehen | `easyONB.ini` (beide Versionen) |
| C-07 | Welcome-Mail versendet das Initialkennwort im Klartext per SMTP | ExtendedFunctions / NotificationModule |
| C-08 | Kompletter `userData`-Inhalt inkl. manuell eingegebenem Kennwort wird per `ConvertTo-Json`/`Out-String` ins Debug-Log geschrieben | 1.4.23 Z. 2277, 4177 |

### High

| ID | Befund |
|---|---|
| H-01 | Kennwortgenerator nutzt `System.Random` bzw. `Get-Random` (nicht kryptografisch); Indexfehler (`$random.Next(26,50)` usw.) und fehlerhafte Array-Bereiche (`0..-1`) verfälschen Zusammensetzung |
| H-02 (F-02) | 57 im Code referenzierte Controls existieren nicht in `assets/MainGUI.xaml` (z. B. `cmbDisplayTemplate`, `cmbSuffix`, `txtPhone`, `lstUsersADUpdate`); `Get-DisplayNameFormat` erhält `$null` für einen Pflichtparameter → **Onboarding-Button schlägt immer fehl** |
| H-03 | Keine Sperre für privilegierte Gruppen: Gruppen aus INI, CSV oder Freitext (`InputBox`) werden ungeprüft per `Add-ADGroupMember` zugewiesen (auch "Domain Admins") |
| H-04 | `SyncCommand` aus der INI wird per `[ScriptBlock]::Create()` remote ausgeführt → Codeausführung über Konfigurationsdatei |
| H-05 | AD-Filter-Injection: Suchbegriffe/Namen werden unescaped in `-Filter`-Strings eingesetzt (`'*$searchTerm*'`) |
| H-06 | Reports: Benutzerwerte werden ohne HTML-Encoding eingesetzt; wkhtmltopdf läuft mit `--enable-local-file-access` (HTML-Injection → lokaler Dateizugriff); wkhtmltopdf ist seit 2023 archiviert |
| H-07 | `-replace` mit Benutzerwerten als Ersetzungsstring: `$`-Sequenzen (z. B. in Kennwörtern) werden als Regex-Substitution interpretiert |
| H-08 | Modul-Ladeblock bricht wegen undefiniertem `$scriptPath` ab; 4 Module werden nie geladen |
| H-09 | Selbst-Elevation mit `-ExecutionPolicy Bypass`; Admin-Rechte werden pauschal verlangt, obwohl AD-Operationen keine lokalen Adminrechte benötigen |
| H-10 | Installer lädt PowerShell-MSI und wkhtmltopdf ohne Hash-/Signaturprüfung herunter und führt sie aus |
| H-11 | Klartext-Secrets in der INI vorgesehen: `CompanyVPNPassword`, `JiraToken`, SMTP `Password` (NotificationModule) |
| H-12 | Manager-Zuweisung per Scriptblock-Filter mit Property-/Methodenaufruf (`$UserData.Manager`, `.Text.Trim()`) → funktioniert im AD-Provider nicht; Mehrdeutigkeit nicht behandelt |

### Medium

| ID | Befund |
|---|---|
| M-01 | Fehlende Mail-Adresse erzeugt `mail=placeholder@<domain>` |
| M-02 | `Get-ADUser -Filter *` lädt alle Benutzer (Performance/Last in großen Domänen) |
| M-03 | Stille Fehler: `-ErrorAction SilentlyContinue` bei Kennwort-Reset und Attributupdates, leere `catch`-Blöcke, Fehler nur als Debug-Meldung |
| M-04 | `Register-GUIEvent` verliert `$EventAction` (keine Closure) |
| M-05 | Report-Platzhalter wie `{{CompanyWebsite}}`, `{{CompanyWikiURL}}` werden nie ersetzt und bleiben sichtbar |
| M-06 | INI-Speichern verwirft Kommentare und schreibt nicht atomar |
| M-07 | `New-HomeDirectory` nutzt englischen Gruppennamen "Domain Admins" (scheitert in deutschsprachigen Domänen), entfernt vorher alle ACEs |
| M-08 | Graph-Scope `Directory.ReadWrite.All` statt minimal notwendiger Berechtigungen |
| M-09 | `Send-MailMessage` ist veraltet; SMTP-Konfigurationsschlüssel inkonsistent (`From` vs. `FromAddress`) |
| M-10 | Kein strukturiertes Logging (keine Operation-ID, kein Akteur, keine Dauer, kein Audit-Level, keine Rotation) |
| M-11 | Hart codierte Pfade `C:\easyIT\easyONBOARDING\…` in der INI |
| M-12 | XAML enthält `Text="$($env:USERNAME)"` als Literal; 85 hart codierte Farbwerte; kein Theme-Konzept |
| M-13 | Synchrone AD-Aufrufe im UI-Thread ohne Statusrückmeldung |

### Low

| ID | Befund |
|---|---|
| L-01 | Toter Code: Top-Level-Block mit `$userData.UPNFormat` (Z. 1264), `Process-ADGroups` (Platzhalter), doppelte Zweige in `Get-SelectedADGroups` |
| L-02 | Leere XAML-Dateien, ungültige Backup-XAML, `rep.html` Duplikat |
| L-03 | Eingecheckte Logdatei, Binärdateien (`.exe`, `.zip`, `.7z`) im Archiv |
| L-04 | Syntaxfehler in `# Extra Tools/HR_AL-Onboarding` (2 Dateien) und `# TEST/WEB Edition` (2 Dateien) |
| L-05 | Nicht genehmigte Verben (`Load-ADGroups`, `Process-ADGroups`, `Calculate-PasswordStrength`, …) |
| L-06 | PSScriptAnalyzer (1.24.0) auf 1.4.23 + Module: **422 Befunde** (5 Error, 417 Warning; u. a. 276× globale Variablen, 27× fehlendes ShouldProcess) |

### Recommendation

| ID | Empfehlung |
|---|---|
| R-01 | Modularisierung mit klarer Trennung GUI / Logik / Konfiguration / Infrastruktur |
| R-02 | Plan-/Vorschau-/Ausführungsmodell mit `SupportsShouldProcess` für alle schreibenden Operationen |
| R-03 | Pester-Tests mit AD-Stubs, CI mit Parser-, Analyzer-, Test- und Secret-Prüfung |
| R-04 | Einheitliche Versionierung, CHANGELOG, SECURITY/CONTRIBUTING |
| R-05 | Lizenzfrage durch den Maintainer klären (README nennt MIT, Lizenzdatei fehlt) |

## 7. XAML / WPF

* `assets/MainGUI.xaml` ist gültiges XML und nutzt `Name=` statt `x:Name=`; 151 benannte Elemente,
  109 davon werden vom Code nie verwendet, 57 vom Code verwendete Namen fehlen.
* Moderne Navigation (Dashboard, Onboarding, Update, CSV, Reports, Tools, Settings) ist als Layout
  vorhanden, aber überwiegend nicht an Logik angebunden.
* Keine Ressourcentrennung (Styles inline im Window), keine Themes, keine Tastaturführung, keine
  feldbezogene Validierung.

## 8. PowerShell-Kompatibilität

* 1.4.23 nutzt PS-7-Syntax (`?.`, `??`, `? :`) und verlangt PS 7; Kommentare behaupten 5.1-Kompatibilität.
* `PDFCreator` wird mit `powershell.exe` (5.1) gestartet, das Hauptskript mit `pwsh.exe`.
* Module werden mit `-Force` aus relativen Pfaden geladen; keine Manifeste.

## 9. Risiken für bestehende Installationen

* Produktive INI-Dateien enthalten ggf. `fixPassword`, `CompanyVPNPassword`, `JiraToken`.
* Vorhandene Logdateien und Reports (`Reports/*.html|txt|pdf`) können Klartextkennwörter enthalten
  und sollten geprüft und bereinigt werden (siehe `SECURITY.md`).
* Signierte Skripte: Umgebungen mit `AllSigned` benötigen eine Neusignierung geänderter Dateien.

## 10. Abweichungen Code ↔ Dokumentation (Auszug)

* README-Beispiel `CompanyADDomain`, `DefaultOU` unter `[Company]`, `LicensesGroups E3=` weichen von den tatsächlichen Schlüsseln ab.
* README nennt Module/Funktionen (`New-ADUserFromData`, `Validate-Settings`, `Initialize-GUI`), die nicht existieren.
* "Hot-Reload", "Live-Validierung", "Audit Trail", "Asset Management", "Schulungsplan", "Zugangskarten" sind als implementiert markiert, existieren aber nicht.
* Platzhalter `yourusername`, `yourdomain.com`, `support@yourdomain.com` in `V1.4.XX/README.md`.
* Beispiel-INI enthält reale Domänen des Autors (`phinit.de`, `psscripts.de` …) statt neutraler Beispieldaten.
