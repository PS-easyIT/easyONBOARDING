# easyONBOARDING v1.4.1

Ein professionelles PowerShell-Tool für IT-Administratoren zur Automatisierung und Vereinfachung des User-Onboarding-Prozesses in Active Directory-Umgebungen.

## 📋 Inhaltsverzeichnis

- [Überblick](#überblick)
- [Hauptfunktionen](#hauptfunktionen)
- [Voraussetzungen](#voraussetzungen)
- [Installation](#installation)
- [Konfiguration](#konfiguration)
- [Verwendung](#verwendung)
- [Architektur](#architektur)
- [Module](#module)
- [Sicherheit](#sicherheit)
- [Fehlerbehebung](#fehlerbehebung)
- [Contributing](#contributing)
- [Lizenz](#lizenz)
- [🆕 Erweiterte Funktionen (v1.4.1+)](#🆕-erweiterte-funktionen-v141)
- [📧 E-Mail-Konfiguration](#📧-e-mail-konfiguration)
- [🔄 AD Synchronisation](#��-ad-synchronisation)

## 🚀 Überblick

easyONBOARDING ist ein umfassendes PowerShell-basiertes Tool mit moderner WPF-GUI, das den gesamten Prozess der Benutzeranlage in Active Directory automatisiert. Es wurde speziell für IT-Administratoren entwickelt, um Zeit zu sparen und Fehler zu minimieren.

### Warum easyONBOARDING?

- **Zeitersparnis**: Reduziert den Onboarding-Prozess von Stunden auf Minuten
- **Fehlerminimierung**: Standardisierte Prozesse verhindern manuelle Fehler
- **Compliance**: Einheitliche Benutzeranlage nach Unternehmensrichtlinien
- **Reporting**: Umfassende Dokumentation aller Aktionen

## ✨ Hauptfunktionen

### Benutzeranlage
- ✅ Automatische AD-Benutzeranlage mit allen erforderlichen Attributen
- ✅ Intelligente UPN-Generierung basierend auf konfigurierbaren Templates
- ✅ Flexible Anzeigenamen-Formate
- ✅ Automatische E-Mail-Adressgenerierung
- ✅ Manager-Zuweisung mit Suche

### Gruppenverwaltung
- ✅ Dynamische AD-Gruppenzuweisung aus INI-Konfiguration
- ✅ Multi-Select für mehrere Gruppenzuweisungen
- ✅ Teamleiter-Gruppenzuweisung
- ✅ Lizenzgruppenverwaltung

### Passwort-Management
- ✅ Erweiterte Passwortgenerierung mit konfigurierbaren Regeln
- ✅ Manuelle Passworteingabe mit Validierung
- ✅ Passwort-Komplexitätsprüfung
- ✅ "Change at next logon" Option

### CSV-Import
- ✅ Bulk-Import von Benutzern
- ✅ Validierung vor Import
- ✅ Fehlerbehandlung und Logging
- ✅ Export von Ergebnissen

### Reporting
- ✅ PDF-Report-Generierung für neue Benutzer
- ✅ HTML-Reports für Übersichten
- ✅ CSV-Export für weitere Verarbeitung
- ✅ Umfassende Aktivitätslogs

### Multi-Company Support
- ✅ Unterstützung mehrerer Unternehmen/Domains
- ✅ Separate Konfigurationen pro Company
- ✅ Company-spezifische Templates

## 📌 Voraussetzungen

### System
- Windows 10/11 oder Windows Server 2016+
- PowerShell 7.0 oder höher
- .NET Framework 4.7.2+
- Administrator-Rechte

### Active Directory
- Active Directory PowerShell-Modul
- Berechtigungen zum Erstellen/Ändern von AD-Objekten
- Netzwerkzugriff auf Domain Controller

### Optional
- Exchange Management Tools (für Mailbox-Erstellung)
- PDF-Reader für Report-Anzeige

## 📦 Installation

### 1. Repository klonen
```powershell
git clone https://github.com/yourusername/easyONBOARDING.git
cd easyONBOARDING/V1.4.XX
```

### 2. Verzeichnisstruktur erstellen
```powershell
# Erstelle erforderliche Verzeichnisse
New-Item -ItemType Directory -Path ".\Logs" -Force
New-Item -ItemType Directory -Path ".\Reports" -Force
New-Item -ItemType Directory -Path ".\ReportTemplates" -Force
New-Item -ItemType Directory -Path ".\assets" -Force
```

### 3. Konfiguration anpassen
```powershell
# Kopiere und passe die INI-Datei an
Copy-Item ".\easyONB.ini.template" ".\easyONB.ini"
notepad ".\easyONB.ini"
```

### 4. Script starten
```powershell
# Mit Administrator-Rechten ausführen
pwsh.exe -ExecutionPolicy Bypass -File ".\easyONBOARDING_V1.4.1.ps1"
```

## ⚙️ Konfiguration

### INI-Datei Struktur

Die `easyONB.ini` enthält alle Konfigurationseinstellungen:

```ini
[Company]
CompanyName=Ihre Firma GmbH
CompanyShortName=YF
CompanyADDomain=@yourdomain.com
CompanyMailDomain=yourdomain.com
DefaultOU=OU=Users,OU=Company,DC=yourdomain,DC=com

[DisplayNameUPNTemplates]
DefaultDisplayNameFormat={first} {last}
DisplayNameTemplate1=Nachname, Vorname|{last}, {first}
DisplayNameTemplate2=Vorname Nachname|{first} {last}
UPNTemplate1=firstname.lastname
UPNTemplate2=f.lastname

[ADGroups]
ADGroup1=GRP-AllUsers
ADGroup2=GRP-Office365
ADGroup3=GRP-VPN-Access

[LicensesGroups]
E3=GRP-License-E3
E5=GRP-License-E5
F3=GRP-License-F3

[Logging]
LogFile=.\Logs
DebugMode=1
```

### Wichtige Konfigurationsabschnitte

#### Company
- Definiert Unternehmensinformationen
- Mehrere Company-Sections für Multi-Tenant möglich

#### DisplayNameUPNTemplates
- Templates für Anzeigenamen und UPN
- Unterstützt Platzhalter: {first}, {last}

#### ADGroups
- Liste aller verfügbaren AD-Gruppen
- Werden als Checkboxen in der GUI angezeigt

#### Logging
- LogFile: Pfad für Log-Dateien
- DebugMode: 1 = aktiviert, 0 = deaktiviert

## 🖥️ Verwendung

### Grundlegende Benutzeranlage

1. **Script starten** mit Administrator-Rechten
2. **AD-Verbindung** herstellen (automatisch oder manuell)
3. **Navigation** zu "User Onboarding"
4. **Pflichtfelder** ausfüllen:
   - Vorname
   - Nachname
   - OU auswählen (Refresh-Button drücken)
5. **Optionale Felder**:
   - Anzeigename (wird automatisch generiert)
   - E-Mail-Adresse
   - Manager
   - AD-Gruppen
6. **Passwort** generieren oder manuell eingeben
7. **"Create User"** klicken

### CSV-Import

1. CSV-Datei vorbereiten mit Spalten:
   - FirstName
   - LastName
   - DisplayName (optional)
   - Email (optional)
   - Department (optional)
   - Manager (optional)

2. Navigation zu "CSV Import"
3. Datei auswählen und importieren
4. Validierung prüfen
5. Import starten

### Benutzer-Update

1. Navigation zu "User Update"
2. Benutzer suchen (Name oder SamAccountName)
3. Eigenschaften ändern
4. Änderungen speichern

## 🏗️ Architektur

### Hauptkomponenten

```
easyONBOARDING/
├── easyONBOARDING_V1.4.1.ps1    # Hauptscript
├── MainGUI.xaml                   # GUI-Definition
├── easyONB.ini                    # Konfiguration
├── Modules/                       # PowerShell-Module
│   ├── ADManager.psm1            # AD-Funktionen
│   ├── SettingsManager.psm1      # Einstellungsverwaltung
│   └── UIManager.psm1            # GUI-Funktionen
├── assets/                        # Icons und Bilder
├── Logs/                          # Log-Dateien
├── Reports/                       # Generierte Reports
└── ReportTemplates/               # Report-Vorlagen
```

### Datenfluss

1. **INI-Konfiguration** wird beim Start geladen
2. **Module** werden importiert
3. **GUI** wird aus XAML geladen
4. **Event-Handler** werden registriert
5. **AD-Verbindung** wird hergestellt
6. **Benutzerinteraktion** löst Funktionen aus
7. **Logging** dokumentiert alle Aktionen

## 📚 Module

### ADManager.psm1
- `New-ADUserFromData`: Erstellt AD-Benutzer
- `Get-ADUserDetails`: Lädt Benutzerinformationen
- `Update-ADUserProperties`: Aktualisiert Eigenschaften
- `Add-UserToGroups`: Fügt Benutzer zu Gruppen hinzu

### SettingsManager.psm1
- `Get-Configuration`: Lädt Konfiguration
- `Save-Configuration`: Speichert Einstellungen
- `Validate-Settings`: Prüft Konfiguration

### UIManager.psm1
- `Initialize-GUI`: Initialisiert die Oberfläche
- `Update-UIElements`: Aktualisiert GUI-Elemente
- `Show-MessageDialog`: Zeigt Dialoge an

## 🔒 Sicherheit

### Best Practices

1. **Berechtigungen**
   - Minimale AD-Rechte verwenden
   - Service-Account für Automation

2. **Passwörter**
   - Niemals im Klartext speichern
   - Komplexe Passwort-Richtlinien

3. **Logging**
   - Alle Aktionen protokollieren
   - Logs regelmäßig archivieren

4. **Validierung**
   - Eingaben immer validieren
   - SQL-Injection verhindern

### Sicherheitshinweise

- ⚠️ Script nur mit notwendigen Rechten ausführen
- ⚠️ INI-Datei vor unbefugtem Zugriff schützen
- ⚠️ Logs können sensible Daten enthalten
- ⚠️ Regelmäßige Sicherheitsupdates durchführen

## 🔧 Fehlerbehebung

### Häufige Probleme

#### AD-Modul kann nicht geladen werden
```powershell
# RSAT installieren
Add-WindowsCapability -Online -Name Rsat.ActiveDirectory.DS-LDS.Tools
```

#### Keine Verbindung zum AD
- Domain Controller erreichbar prüfen
- Berechtigungen überprüfen
- DNS-Auflösung testen

#### GUI wird nicht angezeigt
- PowerShell-Version prüfen (7.0+)
- XAML-Datei vorhanden?
- .NET Framework aktuell?

### Debug-Modus

Aktivieren Sie den Debug-Modus in der INI:
```ini
[Logging]
DebugMode=1
```

### Log-Analyse

Logs befinden sich im `Logs`-Verzeichnis:
- `easyOnboarding.log`: Hauptlog
- `easyOnboarding_error.log`: Fehlerlog

## 🤝 Contributing

Wir freuen uns über Beiträge! Bitte beachten Sie:

1. Fork des Repositories erstellen
2. Feature-Branch erstellen (`git checkout -b feature/AmazingFeature`)
3. Änderungen committen (`git commit -m 'Add AmazingFeature'`)
4. Branch pushen (`git push origin feature/AmazingFeature`)
5. Pull Request erstellen

### Coding Standards

- PowerShell Best Practices befolgen
- Funktionen dokumentieren
- Error Handling implementieren
- Tests schreiben

## 📄 Lizenz

Dieses Projekt ist lizenziert unter der MIT-Lizenz - siehe [LICENSE](LICENSE) für Details.

## 👥 Autoren

- **IT Team** - *Initial work* - [YourGitHub](https://github.com/yourusername)

## 🙏 Danksagungen

- PowerShell Community für Inspiration
- Alle Contributor und Tester
- Microsoft für PowerShell und Active Directory

---

📧 **Support**: support@yourdomain.com  
🌐 **Website**: https://easy-it.phinit.de  
📖 **Dokumentation**: [Wiki](https://github.com/yourusername/easyONBOARDING/wiki)

## 🆕 Erweiterte Funktionen (v1.4.1+)

### Fehlende Funktionen implementiert

Die folgenden Funktionen wurden als Module implementiert und können über die Konfiguration aktiviert werden:

#### ExtendedFunctions Modul
- **Export-UserData**: Exportiert Benutzerdaten in verschiedene Formate (CSV, JSON, XML, HTML)
- **Send-WelcomeEmail**: Optionaler E-Mail-Versand an neue Benutzer mit Anmeldeinformationen
- **Create-UserMailbox**: Automatische Exchange-Postfach-Erstellung (On-Premises/Online)
- **Set-UserPhoto**: Profilbild-Upload in AD und Exchange
- **Create-HomeDirectory**: Automatische Home-Verzeichnis-Erstellung mit korrekten Berechtigungen
- **Set-UserPermissions**: Erweiterte Berechtigungsverwaltung für Gruppen und Ordner
- **Backup-UserData**: Backup-Funktionalität für Benutzerdaten
- **Generate-UserReport**: Detaillierte Berichte (Summary, Detailed, Audit)
- **Validate-UserInput**: Eingabevalidierung mit konfigurierbaren Regeln
- **Test-Prerequisites**: Voraussetzungsprüfung für das System

### AD Synchronisation
- **ADSync Integration**: Automatische Synchronisation mit Azure AD Connect
- **Konfigurierbarer Sync-Server**: Angabe des AD Connect Servers in der Domäne
- **Auto-Sync Option**: Automatischer Sync nach Benutzeranlage
- **Test-Funktionen**: Verbindungstest zum Sync-Server

### Moderne Settings GUI
- **Übersichtliches Design**: Modernes Windows 11 Fluent Design
- **Kategorisierte Einstellungen**: 
  - General Settings
  - Company Information
  - Active Directory
  - Email Settings (mit optionaler Welcome Mail)
  - Password Policy
  - Groups & Licenses
  - Reporting
  - Advanced Settings
- **Live-Validierung**: Eingaben werden sofort validiert
- **Hot-Reload**: Änderungen werden ohne Neustart übernommen

### ProxyMailAddress Unterstützung
- **Automatische Proxy-Adressen**: Konfiguration von primären und sekundären E-Mail-Adressen
- **Exchange-Integration**: Unterstützung für Exchange Online und On-Premises
- **Multi-Domain Support**: Verschiedene E-Mail-Domänen für unterschiedliche Zwecke

### IT-Onboarding Optimierungen
- **Checklisten-Integration**: Vollständiger Onboarding-Prozess mit allen IT-Aufgaben
- **Asset Management Ready**: Vorbereitung für Geräte- und Lizenzverwaltung
- **Dokumentation**: Automatische Generierung von Onboarding-Dokumenten
- **Audit Trail**: Vollständige Nachvollziehbarkeit aller Aktionen

## 📧 E-Mail-Konfiguration

### Welcome E-Mail (Optional)
```ini
[EmailSettings]
SendWelcomeEmail=1  # 0=deaktiviert, 1=aktiviert
SMTPServer=mail.company.com
SMTPPort=25
UseSSL=0
FromAddress=it@company.com
```

Die Welcome E-Mail ist standardmäßig deaktiviert und kann bei Bedarf aktiviert werden.

## 🔄 AD Synchronisation

### Konfiguration
```ini
[ADSync]
EnableADSync=1
ADSyncServer=ADCONNECT01  # Name des AD Connect Servers
ADSyncGroup=ADSyncUsers
SyncCommand=Start-ADSyncSyncCycle -PolicyType Delta
AutoSyncNewUsers=1  # Automatischer Sync nach Benutzeranlage
```

### Verwendung
1. AD Sync Server in den Einstellungen konfigurieren
2. Verbindung testen mit der Test-Funktion
3. Optional: Auto-Sync aktivieren für automatische Synchronisation