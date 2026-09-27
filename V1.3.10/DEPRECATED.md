# Veraltet: easyONBOARDING 1.3.10

Dieser Ordner enthält die letzte als final markierte Version der Reihe 1.x (**1.3.10**, 30.03.2025).
Sie wird nicht mehr gepflegt; der Inhalt bleibt unverändert erhalten (Nachvollziehbarkeit, gültige
Authenticode-Signaturen, Rückfallebene während des Umstiegs).

**Bitte nicht mehr produktiv einsetzen.** Bekannte kritische Mängel dieser Version (Auszug):

- Generierte Kennwörter werden im Debug- und Anwendungslog sowie in HTML-/TXT-Berichten im Klartext
  gespeichert; ein fester Standardkennwortwert (`fixPassword`) ist vorgesehen.
- Das Onboarding kann ein bestehendes Konto mit gleichem Anmeldenamen überschreiben.
- Gruppen werden ohne Sperre privilegierter Gruppen zugewiesen; der Kennwortgenerator ist nicht
  kryptografisch sicher.

Vollständige Bestandsaufnahme: [V3.0.0/docs/ANALYSIS.md](../V3.0.0/docs/ANALYSIS.md).

**Nachfolger:** [easyONBOARDING 3](../V3.0.0/README.md) – Umstieg einschließlich Übernahme der INI:
[V3.0.0/docs/MIGRATION.md](../V3.0.0/docs/MIGRATION.md). Hinweise zur Bereinigung alter Logs und
Berichte: [V3.0.0/SECURITY.md](../V3.0.0/SECURITY.md).
