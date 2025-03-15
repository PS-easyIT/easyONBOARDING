# PSscripts.de  -  ANDREAS HEPP  –  easyOnboarding

Version         : 1.0.0

Letztes Update  : 23.02.2025

Autor           : ANDREAS HEPP

Website         : www.psscripts.de


# Beschreibung:
Dieses Skript dient dem Onboarding neuer Mitarbeiter in Active Directory.  
Es liest eine Konfigurationsdatei (INI) ein, zeigt eine grafische Benutzeroberfläche (GUI) zur Dateneingabe an 
und legt anschließend AD‑Benutzer an oder aktualisiert diese.  

Zudem werden diverse Reports (HTML, PDF, TXT) generiert und ein Logfile erstellt.


# UPN-Erzeugung:
Für den UPN wird zunächst geprüft, ob in der Company‑Sektion der INI der Schlüssel „ActiveDirectoryDomain“ definiert ist.  
Falls vorhanden, wird dessen Wert (ggf. mit führendem „@“) als UPN‑Suffix verwendet.  
Ist dieser Schlüssel nicht vorhanden, wird der im GUI ausgewählte Mail‑Suffix (oder alternativ der Schlüssel „MailDomain“) angehängt.


# DisplayName-Erzeugung:
Wird in der jeweiligen Company‑Sektion der INI der Schlüssel „NameFirma“ (oder alternativ „UserDisplayNameFirma“) gefunden,  
so wird dessen Wert als Präfix verwendet und der DisplayName wird im Format  "NameFirma | VORNAME NACHNAME"  gesetzt. 
Andernfalls wird lediglich Vorname und Nachname verwendet. 

Bei externen Mitarbeitern wird, wenn ausgewählt, ein entsprechender Hinweis (z. B. "EXTERN | …") hinterlegt.


# E-Mail-Feld:
Beim Befüllen der E‑Mail-Adresse wird geprüft, ob der eingegebene Wert bereits ein „@“ enthält.  
Falls nicht, wird – sofern in der Company‑Sektion ein MailDomain-Wert definiert ist – dieser angehängt,  
ansonsten der im GUI ausgewählte Mail‑Suffix.


# Report-Erstellung:
Das Skript erzeugt HTML-Reports, die dynamisch mit den folgenden Platzhaltern befüllt werden:

- **{{ReportTitle}}**: Der Titel des Reports (z. B. „Onboarding Report für neue Mitarbeiter“)
- **{{LogoTag}}**: HTML-Code für das Firmenlogo, basierend auf dem in der INI hinterlegten Pfad
- **{{UserDetailsHTML}}**: Dynamisch generierter HTML-Code, der weitere Benutzerdetails in Tabellenform enthält
- **{{WebsitesHTML}}**: HTML-Code für weiterführende Links (basierend auf den EmployeeLink-Einträgen in der INI)
- **{{ReportFooter}}**: Der Footer-Text des Reports (z. B. Copyright-Hinweis)
- **{{Vorname}}, {{Nachname}}, {{DisplayName}}**: Angaben zum Benutzer
- **{{Extern}}**: Zeigt „Ja“ oder „Nein“ an, abhängig davon, ob der Benutzer als externer Mitarbeiter gekennzeichnet ist
- **{{Description}}, {{Buero}}, {{Rufnummer}}, {{Mobil}}**: Weitere Kontaktdaten und Beschreibung
- **{{Position}}, {{Abteilung}}**: Berufliche Informationen
- **{{Ablaufdatum}}**: Das Ablaufdatum des Benutzerkontos (oder „nie“)
- **{{Company}}**: Firmenname bzw. Domain, der als Präfix im DisplayName genutzt wird
- **{{LoginName}}**: Der SamAccountName des Benutzers
- **{{Passwort}}**: Das generierte bzw. eingetragene Passwort
- **{{Admin}}**: Der Name des Administrators, der das Skript ausführt
- **{{ReportDate}}**: Das Erstellungsdatum des Reports


# Wichtige Hinweise:
- Stellen Sie sicher, dass alle benötigten Module (z. B. ActiveDirectory) vor der Ausführung verfügbar sind.
- Dazu bitte easyONBOARDING_INSTALL.exe als Administrator ausführen
