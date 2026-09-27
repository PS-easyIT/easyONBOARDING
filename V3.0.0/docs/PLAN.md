# Umsetzungsplan (Phase B)

Abgeleitet aus [ANALYSIS.md](ANALYSIS.md). Reihenfolge nach Priorität:
Sicherheit/Datenverlust → Fehlerbehebung → Stabilität → Onboarding → Offboarding → GUI → Tests → Doku → Optionales.

## Grundsatzentscheidung

Der Legacy-Code (1.3.10 / 1.4.23) ist signiert, monolithisch und in zentralen Pfaden defekt
(GUI-Onboarding, Module, Update-Tab). Eine schrittweise Reparatur *im* Monolithen würde jede
Signatur brechen und die kritischen Befunde (Kennwörter in Logs/Reports, Überschreiben
bestehender Konten) nur punktuell beheben. Deshalb:

1. **Neue modulare Anwendung 3.0.0 im Unterordner `V3.0.0`** gemäß Zielarchitektur (entspricht der
   bestehenden Konvention mit Versionsordnern `V1.3.10`, `V1.4.XX`).
2. **Legacy-Ordner bleiben unverändert** (Rollback, Nachvollziehbarkeit, gültige Signaturen) und werden
   mit einem Hinweis `DEPRECATED.md` versehen.
3. **INI-Kompatibilität:** Die 3.0-Konfiguration liest das bestehende INI-Format. Unsichere Schlüssel
   (`fixPassword`, Klartext-Secrets, `SyncCommand`) werden erkannt, gemeldet und ignoriert.
   Migrationsskript: `Scripts/Convert-LegacyConfiguration.ps1`.

Details der Architekturentscheidungen: [ARCHITECTURE.md](ARCHITECTURE.md).

## Arbeitspakete

| # | Aufgabe | Prio | Betroffene Dateien | Risiko | Akzeptanzkriterium |
|---|---|---|---|---|---|
| WP1 | Core: Version, Pfade, Ergebnisobjekte, strukturiertes Logging/Audit mit Redaktion, Plan-Engine | Sicherheit | `VERSION`, `Modules/Core/*` | gering | Logs enthalten keine Secrets (Test), Logfehler brechen den Prozess nicht ab |
| WP2 | Konfiguration: kompatibler INI-Parser, Schema, Validierung, Company-Auflösung, Include-Verzeichnisse, kommentarerhaltendes Schreiben, Migration | Sicherheit/Stabilität | `Modules/Configuration/*`, `Config/*`, `Scripts/Convert-LegacyConfiguration.ps1` | mittel (Kompatibilität) | Legacy-INI 1.3.10/1.4.23 wird gelesen; unsichere Schlüssel werden gemeldet; unbekannte Schlüssel bleiben erhalten |
| WP3 | Security: CSPRNG-Kennwörter, Richtlinienprüfung, SecureString, privilegierte Gruppen (SID-basiert), geschützte Konten, LDAP-/Filter-Escaping, Eingabeformate | Sicherheit | `Modules/Security/*` | gering | Privilegierte Gruppen werden nie zugewiesen; Tests ohne Ausgabe realer Kennwörter |
| WP4 | AD-Adapter mit `SupportsShouldProcess` für alle Schreiboperationen, sichere Suche | Fehlerbehebung | `Modules/ActiveDirectory/*` | mittel | Alle Schreibfunktionen unterstützen `-WhatIf/-Confirm`; Suche escaped |
| WP5 | Onboarding: Transliteration, Templates, Kollisionen, Anfrage-Validierung, Plan/Vorschau/Ausführung, CSV-Bulk | Onboarding | `Modules/Onboarding/*` | mittel | Kein bestehendes Konto wird überschrieben; Plan zeigt alle Gruppenänderungen |
| WP6 | Offboarding: Vorlagen aus Konfiguration, Schutzprüfungen, Snapshot vorher/nachher, Aktionskatalog, Mehrstufigkeit mit Warteschlange, abgesicherte Löschung, CSV | Offboarding | `Modules/Offboarding/*`, `Config/templates/*` | hoch (destruktiv) | Keine Aktion ohne Vorschau/Bestätigung; Löschung nur mit Frist, Backup, Tippbestätigung |
| WP7 | Reporting (HTML/CSV/JSON/TXT, PDF mit Fallback), Legacy-Templates sicher, Exchange/Entra optional, Integrationsstatus | Stabilität | `Modules/Reporting/*`, `Modules/Exchange/*`, `Modules/Entra/*`, `ReportTemplates/*` | mittel | Reports ohne Secrets (Test), HTML-Encoding (Test), fehlende Integrationen verständlich angezeigt |
| WP8 | WPF-GUI: linke Navigation, Dashboard, Wizards, Themes (Light/Dark), zentrale Dialoge, kooperative Ausführung | GUI | `GUI/*`, `Modules/UI/*`, `Start-easyONBOARDING.ps1` | hoch (nicht lokal testbar) | XAML ist gültig, alle referenzierten Controls existieren (statischer Test), XAML-Ladetest in CI (Windows) |
| WP9 | Tests & CI: Pester (Unit), Parser, PSScriptAnalyzer, Secret-Scan, Doku-Prüfung | Tests | `Tests/*`, `.github/workflows/v3-ci.yml` (Repository-Root, technisch erforderlich), `PSScriptAnalyzerSettings.psd1`, `Scripts/*` | gering | Alle Tests grün; Workflow mit gepinnten SHAs und `contents: read` |
| WP10 | Dokumentation: README, CHANGELOG, SECURITY, CONTRIBUTING, docs/* | Doku | `*.md`, `docs/*` | gering | Doku entspricht Implementierungsstand; Status (produktiv/simuliert/vorbereitet) markiert |

## Nicht-Ziele / bewusst ausgelassen

* Keine automatische Installation von Abhängigkeiten (nur Voraussetzungsanzeige mit Hinweisen).
* Kein automatischer Hot-Reload der Konfiguration (nur explizites "Neu laden" mit Validierung).
* Keine Neusignierung (erfordert das Zertifikat des Maintainers).
* Keine Änderungen an `# ARCHIV`, `# Extra Tools`, `# TEST` (Befunde dokumentiert).

## Umsetzungsstand (27.09.2026)

Alle Arbeitspakete WP1–WP10 sind umgesetzt (Details: [CHANGELOG.md](../CHANGELOG.md)). Offen und nur
außerhalb dieser Entwicklungsumgebung möglich: manuelle Prüfung der Oberfläche unter Windows,
Verifikation der Exchange-, Graph- und Entra-Connect-Funktionen gegen reale Testumgebungen und die
Code-Signatur durch den Maintainer (siehe [TESTING.md](TESTING.md#manuelle-prüfung)).
