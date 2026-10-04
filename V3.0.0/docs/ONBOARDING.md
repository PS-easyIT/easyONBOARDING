# Onboarding

## Ablauf

Jeder Vorgang folgt demselben Muster: **Eingabe → Prüfung → Plan → Vorschau → Bestätigung →
Ausführung → Bericht und Audit.** Nur lesende AD-Abfragen finden vor der Bestätigung statt.

Der Assistent (Navigation *Onboarding*) hat acht Schritte. Pflichtfelder sind mit `*` markiert,
Fehler werden direkt am Feld angezeigt; "Weiter" prüft den aktuellen Schritt.

| Schritt | Inhalt |
|---|---|
| 1 Unternehmen | Unternehmen, Rollenvorlage, Anzeigenamen-Vorlage, Art der Beschäftigung (intern/extern), optional Referenzbenutzer für die Gruppenübernahme |
| 2 Person | Vor- und Nachname, Anzeigename (leer = aus Vorlage), Personalnummer, Beschäftigungsart, Eintrittsdatum, Befristung |
| 3 Konto | Vorschläge für Anmeldename, UPN und E-Mail (änderbar) mit Kollisionsprüfung |
| 4 Organisation | Ziel-OU, Position, Abteilung, Führungskraft, Büro, Telefon, Mobil, Beschreibung |
| 5 Gruppen und Lizenz | Lizenzgruppe, Teamleitung/Abteilungsleitung, weitere Gruppen |
| 6 Optionen | Kennwort generieren oder manuell vergeben, Kennwortänderung bei Anmeldung, sofortige Aktivierung, Proxyadressen, Home-Verzeichnis, Postfach, Willkommensdokument, Welcome-Mail, Entra Connect Sync, Ticket, Bemerkung |
| 7 Vorschau | Zusammenfassung, alle geplanten Schritte mit Risiko und Details, Prüfergebnisse |
| 8 Ausführung | schrittweise Ausführung mit Fortschritt, Abbruch zwischen zwei Schritten, Ergebnis und Bericht |

Der Kopfbereich zeigt jederzeit, ob die Anwendung im **Simulations-** oder **Live-Modus** arbeitet.
In der Simulation werden alle Schritte mit `-WhatIf` durchlaufen; es entsteht ein vollständiger
Bericht, aber keine Änderung.

## Kontonamen

Anmeldename (`sAMAccountName`), UPN-Präfix und lokaler Teil der E-Mail-Adresse werden aus Vorlagen
erzeugt (`[Identity] SamAccountNameFormat`, `[DisplayNameUPNTemplates] DefaultUserPrincipalNameFormat`,
`[Identity] MailFormat`):

| Platzhalter | Bedeutung | Beispiel für "Jürgen Peter Größe" |
|---|---|---|
| `{first}` | erster Vorname | `juergen` |
| `{firstfull}` | alle Vornamen | `juergenpeter` |
| `{last}` | Nachname | `groesse` |
| `{f}`, `{l}` | Initialen | `j`, `g` |
| `{first:N}`, `{last:N}` | erste N Zeichen | `{last:3}` → `gro` |
| `{employeeid}` | Personalnummer | |

* Legacy-Schlüsselwörter wie `FIRSTNAME.LASTNAME`, `F.LASTNAME` oder `FLASTNAME` werden weiter
  unterstützt und in die Platzhalter übersetzt.
* Umlaute und Sonderzeichen werden transliteriert (`ä` → `ae`, `ß` → `ss`, `é` → `e` …); eigene Regeln
  in `[NameNormalization] Transliteration` (z. B. `ø=oe;å=aa`).
* Kleinschreibung (`[Identity] LowerCaseIdentifiers`), maximale Länge des Anmeldenamens (20) und
  Bindestriche (`[NameNormalization] KeepHyphen`) sind konfigurierbar.
* **Kollisionen:** Jeder Vorschlag wird gegen Anmeldenamen, UPNs, E-Mail- und Proxyadressen aller
  AD-Objekte geprüft. Mit `CollisionStrategy=AppendNumber` wird eine Zahl angehängt (ab
  `CollisionStartNumber`), mit `Fail` wird der Vorgang mit einer Meldung angehalten. **Ein bestehendes
  Konto wird nie überschrieben.**

## Gruppen und Lizenzen

Die Gruppen eines neuen Kontos setzen sich zusammen aus: Standardgruppen
(`[UserCreationDefaults] InitialGroupMembership`), Rollenvorlage, Lizenzgruppe (`[LicensesGroups]`),
Teamleitungs-/Abteilungsleitungsgruppe, Synchronisationsgruppe (`[ActivateUserMS365ADSync]`), manueller
Auswahl und – optional – den Gruppen eines Referenzbenutzers. Die Vorschau zeigt jede Gruppe mit ihrer
Herkunft.

**Privilegierte Gruppen werden nie zugewiesen** – auch nicht über Referenzbenutzer, Vorlagen oder CSV.
Sie erscheinen in der Vorschau als blockiert mit Begründung. Welche Gruppen als privilegiert gelten,
beschreibt [SECURITY.md](../../SECURITY.md#sicherheitskonzept-version-3).

## Kennwörter

* **Generiert (empfohlen):** kryptografisch sicher nach `[PasswordFixGenerate]` und – mit
  `[Security] UseDomainPasswordPolicy=1` – mindestens nach der Domänenrichtlinie.
* **Manuell:** nur wenn `[Security] AllowManualPassword=1`; das Kennwort wird gegen Mindestlänge,
  Komplexität, Kontoname, Namensbestandteile und häufige Kennwörter geprüft.
* **Anzeige:** Nach einer erfolgreichen Live-Ausführung erscheint das Kennwort genau einmal im Dialog
  *Zugangsdaten* (`[Security] ShowPasswordOnce`). Kopieren leert die Zwischenablage nach
  `ClipboardClearSeconds` Sekunden; Drucken erfolgt aus dem Speicher (`AllowCredentialPrint`). Danach
  ist das Kennwort nicht mehr abrufbar – bei Verlust hilft nur ein Reset (*Benutzer aktualisieren*).
* Kennwörter erscheinen **nie** in Logs, Berichten, Willkommensdokumenten oder E-Mails.

## Willkommensdokument und Welcome-Mail

Das Willkommensdokument wird aus `[Report] TemplatePathHTML` erzeugt (Platzhalter:
[ReportTemplates/README.md](../ReportTemplates/README.md)). Die Welcome-Mail
(`[EmailSettings] SendWelcomeEmail=1`) geht an die neue E-Mail-Adresse des Kontos, optional in Kopie
an `[EmailSettings] CopyAddress` (z. B. Personalabteilung). Beide enthalten **kein Kennwort**;
Kennwort-Platzhalter werden durch den Hinweis ersetzt, dass das Kennwort separat und persönlich
übergeben wird.

## Berichte

Nach jeder Ausführung (auch in der Simulation) entsteht ein Vorgangsbericht in `[Report] ReportPath`
(Formate nach `[Report] Formats`, Dateiname `<Zeitstempel>_<Konto>_<ID>`), dazu ein Eintrag im
Audit-Log. Berichte enthalten die Schritte mit Status, Befunde und die Zusammenfassung, aber keine
Kennwörter. Die Ansicht *Reports und Audit* listet Berichte und Audit-Einträge und exportiert sie.

## Benutzer aktualisieren

Suche nach Name, Anmeldename, UPN oder E-Mail; geänderte Felder, Führungskraft sowie hinzuzufügende
und zu entfernende Gruppen erscheinen in der Vorschau als Vorher/Nachher. Nur tatsächlich geänderte
Werte werden geschrieben. Deaktiviert angelegte Konten (`[ADUserDefaults] AccountDisabled=True`,
"Aktivierung bei Übergabe") werden hier mit *Konto aktivieren* freigeschaltet; liegt das Konto in der
OU für ausgeschiedene Benutzer, warnt die Vorschau. Ein **Kennwort-Reset** ist eine eigene Aktion mit
Bestätigung und einmaliger Anzeige des neuen Kennworts. Geschützte Konten lassen sich nicht bearbeiten.

## Massenverarbeitung (CSV)

Ansicht *Massenverarbeitung* › Onboarding. Trennzeichen (`;` oder `,`) und Kodierung (UTF-8 oder
Windows-1252) werden erkannt. Jede Zeile wird wie im Assistenten geprüft und geplant; fehlerhafte
Zeilen werden markiert und nicht ausgeführt. Ab `[Bulk] ConfirmationThreshold` Zeilen ist die
Eingabe `AUSFÜHREN <Anzahl>` erforderlich. Höchstens `[Bulk] MaxRows` Zeilen je Datei.

Unterstützte Spalten (Groß-/Kleinschreibung egal, deutsche und Legacy-Bezeichnungen möglich):

| Feld | Spaltennamen |
|---|---|
| Vorname* | `FirstName`, `GivenName`, `Vorname` |
| Nachname* | `LastName`, `Surname`, `Nachname` |
| Anzeigename | `DisplayName`, `Anzeigename` |
| Unternehmen | `Company`, `CompanyId` |
| Position | `Position`, `Title`, `JobTitle` |
| Abteilung | `Department`, `DepartmentField`, `Abteilung` |
| Telefon / Mobil | `PhoneNumber`, `OfficePhone`, `Telefon` / `MobileNumber`, `MobilePhone`, `Mobil` |
| Büro | `Office`, `OfficeRoom`, `Büro` |
| E-Mail (lokaler Teil), Domäne | `EmailAddress`, `Email`, `Mail`; `MailDomain` |
| Eintritt, Befristung | `StartDate`, `Eintrittsdatum`; `ExpirationDate`, `EndDate`, `Ablaufdatum` |
| Führungskraft | `Manager`, `Vorgesetzter`, `Fuehrungskraft` |
| Personalnummer, -art | `EmployeeId`, `Personalnummer`; `EmployeeNumber`; `EmployeeType` |
| Rolle, Lizenz, Gruppen | `RoleTemplate`, `Rolle`; `License`, `Lizenz`; `Groups`, `ADGroups`, `Gruppen` (durch `;` getrennt) |
| Ziel-OU | `OU`, `TargetOU` |
| Extern, Leitung | `External`, `Extern`; `TL`, `TeamLead`; `AL`, `DepartmentHead` |
| Sonstiges | `Description`, `Beschreibung`, `Ticket`, `Notes`, `Bemerkung`, `SamAccountName`, `UpnPrefix` |

Datumswerte im Format `TT.MM.JJJJ` oder `JJJJ-MM-TT`. Kennwörter werden für Massenvorgänge immer
generiert; `[Bulk] PasswordDelivery=Discard` verwirft sie (Reset bei der Übergabe),
`Print` druckt Zugangsdatenblätter aus dem Speicher. Kennwörter werden nie in Ergebnisdateien
geschrieben; Exporte sind gegen Formel-Injection geschützt.
