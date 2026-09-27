# Veraltet: easyONBOARDING 1.4.x

Dieser Ordner enthält den Entwicklungsstand **1.4.23** (Dokumentation: 1.4.1). Er wurde nie als
final freigegeben und wird nicht mehr gepflegt. Der Inhalt bleibt unverändert erhalten
(Nachvollziehbarkeit, gültige Authenticode-Signaturen).

**Bitte nicht mehr produktiv einsetzen.** Bekannte kritische Mängel dieser Version (Auszug):

- Kennwörter gelangen in Logdateien, Berichte und Welcome-Mails; ein fester Standardkennwortwert ist
  vorgesehen.
- Das Onboarding kann bestehende Konten überschreiben; "Benutzer aktualisieren" setzt bei jeder
  Änderung ein neues Kennwort.
- Privilegierte Gruppen werden nicht gesperrt; `SyncCommand` aus der INI wird als Code ausgeführt.
- Der Onboarding-Pfad der Oberfläche ist defekt, mehrere Module werden nie geladen.

Vollständige Bestandsaufnahme: [V3.0.0/docs/ANALYSIS.md](../V3.0.0/docs/ANALYSIS.md).

**Nachfolger:** [easyONBOARDING 3](../V3.0.0/README.md) – Umstieg einschließlich Übernahme der INI:
[V3.0.0/docs/MIGRATION.md](../V3.0.0/docs/MIGRATION.md). Hinweise zur Bereinigung alter Logs und
Berichte: [V3.0.0/SECURITY.md](../V3.0.0/SECURITY.md).
