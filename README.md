# easyONBOARDING

*Active Directory on- and offboarding with PowerShell 7 and WPF. Current version: **3.0.0** in
[`V3.0.0`](V3.0.0/README.md) (documentation in German).*

easyONBOARDING legt neue Mitarbeitende im Active Directory an, pflegt bestehende Konten und führt
Austritte in mehreren Phasen durch – mit Vorschau, Simulation, Bestätigung, Audit und Berichten.

**Aktuelle Version: 3.0.0** · PowerShell 7.2+ · Autor: Andreas Hepp (PHINIT.DE)

| | |
|---|---|
| Überblick und Schnellstart | [V3.0.0/README.md](V3.0.0/README.md) |
| Installation | [V3.0.0/docs/INSTALLATION.md](V3.0.0/docs/INSTALLATION.md) |
| Umstieg von 1.3.10 / 1.4.x | [V3.0.0/docs/MIGRATION.md](V3.0.0/docs/MIGRATION.md) |
| Änderungen | [V3.0.0/CHANGELOG.md](V3.0.0/CHANGELOG.md) |
| Sicherheit | [V3.0.0/SECURITY.md](V3.0.0/SECURITY.md) |
| Mitwirken | [V3.0.0/CONTRIBUTING.md](V3.0.0/CONTRIBUTING.md) |

## Aufbau des Repositorys

| Ordner | Inhalt | Status |
|---|---|---|
| `V3.0.0` | easyONBOARDING 3: Anwendung, Konfiguration, Tests, Dokumentation | **aktuell** |
| `V1.4.XX` | Entwicklungsstand 1.4.23 | veraltet ([DEPRECATED.md](V1.4.XX/DEPRECATED.md)) |
| `V1.3.10` | letzte finale Version der Reihe 1.x | veraltet ([DEPRECATED.md](V1.3.10/DEPRECATED.md)) |
| `# ARCHIV`, `# Extra Tools`, `# INSTALL`, `# PDFCreator`, `# Screenshots`, `# TEST` | ältere Versionen, Hilfswerkzeuge und Prototypen | historisch, nicht Bestandteil von 3.0 |
| `.github/workflows` | CI für Version 3 (`v3-ci.yml`) | aktuell |

Die Ordner der Versionen 1.x bleiben unverändert erhalten (Nachvollziehbarkeit, gültige Signaturen).
Sie enthalten bekannte kritische Sicherheitsmängel und sollten nicht mehr produktiv eingesetzt werden.

## Schnellstart (Version 3)

```powershell
cd .\V3.0.0
pwsh -STA -NoProfile -File .\Install-easyONBOARDING.ps1                  # Konfiguration erstellen
pwsh -NoProfile -File .\Start-easyONBOARDING.ps1 -CheckOnly    # prüfen
pwsh -NoProfile -File .\Start-easyONBOARDING.ps1               # Oberfläche (Simulation voreingestellt)
```

## Autor und Kontakt

Andreas Hepp (PHINIT.DE) – Fragen und Feedback: info@phinit.de – Informationen und Updates:
[www.PSscripts.de](https://www.PSscripts.de).

## Lizenz

Im Repository liegt derzeit keine Lizenzdatei. Die Dokumentation der Version 1.4.x nennt die
MIT-Lizenz; die verbindliche Festlegung trifft der Autor.
