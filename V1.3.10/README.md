Version         : 1.3.10
Letztes Update  : 30.03.2025  
Autor           : ANDREAS HEPP  

Beschreibung:
---------------
Dieses Skript dient dem Onboarding neuer Mitarbeiter in Active Directory.  
Es liest eine externe INI-Konfiguration, zeigt eine moderne grafische Benutzeroberfläche (GUI) zur Dateneingabe an und legt anschließend AD‑Benutzer an oder aktualisiert diese.  
Neu integriert:
  • Erweiterte UPN‑Erzeugung basierend auf individuellen Templates  
  • PDF‑Report-Erstellung zusätzlich zu HTML und TXT  
  • CSV‑Import für die Massenverarbeitung  
  • Dynamische Logo‑Verwaltung & Upload  
  • Erweiterte Passwort‑Generierung (New‑AdvancedPassword) mit konfigurierbaren Sicherheitsoptionen  
  • AD‑User‑Update inklusive erneuter Reporterstellung  

## Regionen & Funktionen

- **00 – SCRIPT INITIALIZATION**  
  - Globale Fehlerbehandlung per trap  
  - Debug-Ausgaben abhängig vom INI‑DebugMode

- **01 – ADMIN AND VERSION CHECK**  
  - Test-Admin und Neustart mit UAC falls nötig

- **02 – INI FILE LOADING**  
  - Laden und Parsen der Konfigurationsdatei (INI)

- **03 – FUNCTION DEFINITIONS**  
  - Diverse Hilfsfunktionen inkl. Logging und Dateioperationen

- **03.1 – LOGGING FUNCTIONS**  
  - Einheitliches Logging mit Zeitstempel und verschiedenen Loglevels

- **03.2 – LOG FILE WRITER**  
  - Schreiben in Logdateien und Fehlerarchivierung

- **04 – ACTIVE DIRECTORY MODULE**  
  - Import und Prüfung des AD‐Moduls

- **05 – WPF ASSEMBLIES**  
  - Integration der WPF-Komponenten (PresentationFramework, etc.)

- **06 – XAML LOADING**  
  - Laden der XAML-Datei und initiales GUI-Binding

- **08 – UPN GENERATION**  
  - Automatische und kontextabhängige UPN‑Erzeugung  
  - Anpassbar per INI‑Templates

- **09 – AD USER CREATION**  
  - Erstellung und Aktualisierung von AD‑Benutzern  
  - Zuweisung von Attributen, Gruppen und Profilpfaden

- **10 – MAIL ACTIVATION (EXO)**  
  - Konfiguration von Exchange Online Attributen

- **11 – CUSTOM ATTRIBUTES**  
  - Zusätzliche AD-Attribute basierend auf GUI-Eingaben

- **12 – GUI EVENT BINDINGS**  
  - Verknüpfung von GUI-Komponenten mit PowerShell-Funktionen

- **13 – MULTI-TAB NAVIGATION**  
  - Bereichsübergreifende Navigation und dynamische Logo-Anpassung

- **14 – STYLE & DESIGN (GUI)**  
  - Anpassbare Themes, Farben und Schriftarten via INI

- **15 – TOOL TAB (Extras)**  
  - Werkzeuge wie AD-Benutzersuche, Reports und Reset‑Funktionen

- **16 – LICENSE / INFO TAB**  
  - Anzeige von Versionsinfos, Lizenz- und Autoreninformationen

- **17 – ZENTRALES LOGO & TAB-LOGOS**  
  - Dynamisches Laden des zentralen Logos und tab-spezifischer Logos

- **18 – EXIT / CLOSE LOGIC**  
  - Beenden des Skripts inkl. Cleanup und Speichern offener Änderungen

## Hinweise zur Konfiguration und Nutzung

- Die INI-Datei steuert u. a. GUI-Design, Texte, Logos sowie Log- und Berichtseinstellungen.  
- Anpassungen am UPN- und DisplayName-Format erfolgen über die entsprechenden INI-Templates.  
- Führen Sie das Skript als Administrator aus, um alle Funktionen uneingeschränkt nutzen zu können.

Wichtige Hinweise:
-------------------
- Für Änderungswünsche wenden Sie sich bitte an: info@phinit.de  
- Stellen Sie sicher, dass alle benötigten Module (z. B. ActiveDirectory) 
  und eine aktuelle PowerShell-Version (7.x) vor der Ausführung verfügbar sind.  
- Installieren Sie easyONBOARDING als Administrator (INSTALL-XXX.exe).

Weitere Informationen:
-----------------------
Für Updates, Best Practices und weitere Details besuchen Sie:  
    https://www.PSscripts.de