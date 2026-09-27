# Berichtsvorlagen

`HTMLTemplate.txt` (Deutsch) und `HTMLTemplateENG.txt` (Englisch) sind die Vorlagen des
**Willkommensdokuments**, das beim Onboarding optional erzeugt wird (`[Report] CreateWelcomeDocument`).
Sie stammen aus Version 1.4.x und wurden für 3.0 bereinigt:

* Kennwort-Platzhalter (`{{Passwort}}`, `{{CustomPW1..5}}`, `{{CompanyVPNPassword}}`) wurden entfernt.
  Werden sie in eigenen Vorlagen weiterverwendet, ersetzt easyONBOARDING sie durch einen Hinweistext –
  Kennwörter gelangen nie in eine Datei.
* Die Kennwortkriterien werden aus der konfigurierten Richtlinie erzeugt (`{{PasswordPolicyList}}`).
* Alle Werte werden HTML-kodiert eingesetzt. Nicht befüllte Platzhalter werden entfernt und im
  Ergebnis als Warnung gemeldet.
* Zusätzliche Platzhalter lassen sich im Abschnitt `[ReportPlaceholders]` der INI definieren.

Die technischen Berichte (Vorgangsbericht, Offboarding-Zustandsbericht) werden unabhängig von diesen
Vorlagen erzeugt (HTML, JSON, CSV, TXT, optional PDF).

Eine Übersicht aller Platzhalter enthält `docs/CONFIGURATION.md`.
