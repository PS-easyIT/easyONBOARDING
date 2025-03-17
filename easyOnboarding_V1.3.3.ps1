<#
  Dieses Skript integriert Funktionen für Onboarding, AD-Gruppenerstellung und den INI-Editor in eine WPF-Oberfläche.
  This script integrates functions for onboarding, AD group creation and the INI editor in a WPF interface.
  
  Voraussetzungen: Administratorrechte und PowerShell 7 oder höher
  Requirements: Administrative rights and PowerShell 7 or higher
    - INI-Datei     ("easyONBOARDINGConfig.ini")
    - XAML-Datei    ("MainGUI.xaml")
#>

# [Region 00 | SCRIPT INITIALIZATION]
# [ENGLISH - Sets up error handling and basic script environment]
# [GERMAN - Richtet Fehlerbehandlung und Skriptumgebung ein]
#requires -Version 7.0

$ErrorActionPreference = "Stop"
trap {
    $errorMessage = "ERROR: Unbehandelter Fehler aufgetreten! " +
                    "Fehlermeldung: $($_.Exception.Message); " +
                    "Position: $($_.InvocationInfo.PositionMessage); " +
                    "StackTrace: $($_.ScriptStackTrace)"
    if ($null -ne $global:Config -and $null -ne $global:Config.General -and $global:Config.General.DebugMode -eq "1") {
        Write-Error $errorMessage
    } else {
        Write-Error "Ein schwerwiegender Fehler ist aufgetreten. Bitte prüfen Sie die Logdatei."
    }
    exit 1
}

#region [Region 01 | ADMIN AND VERSION CHECK]
# [ENGLISH - Verifies administrator rights and PowerShell version before proceeding]
# [GERMAN - Überprüft Administratorrechte und PowerShell-Version vor der Ausführung]
$principal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Error "Dieses Skript muss als Administrator ausgeführt werden."
    exit
}
if ($PSVersionTable.PSVersion.Major -lt 7) {
    Write-Error "Dieses Skript benötigt PowerShell 7 oder höher."
    exit
}
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
#endregion

#region [Region 02 | INI FILE LOADING]
# [ENGLISH - Loads configuration data from the external INI file]
# [GERMAN - Lädt Konfigurationsdaten aus der externen INI-Datei]
function Get-IniContent {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Path
    )
    
    # [2.1 | VALIDATE INI FILE]
    # [ENGLISH - Checks if the INI file exists and is accessible]
    # [GERMAN - Überprüft, ob die INI-Datei existiert und zugänglich ist]
    if (-not (Test-Path $Path)) {
        Throw "INI-Datei nicht gefunden: $Path"
    }
    
    $ini = [ordered]@{}
    $currentSection = "Global"
    $ini[$currentSection] = [ordered]@{}
    
    try {
        $lines = Get-Content -Path $Path -ErrorAction Stop
    } catch {
        Throw "Fehler beim Lesen der INI-Datei: $($_.Exception.Message)"
    }
    
    # [2.2 | PARSE INI CONTENT]
    # [ENGLISH - Processes the INI file line by line to extract sections and key-value pairs]
    # [GERMAN - Verarbeitet die INI-Datei Zeile für Zeile, um Abschnitte und Schlüssel-Wert-Paare zu extrahieren]
    foreach ($line in $lines) {
        $line = $line.Trim()
        # Überspringe leere Zeilen und Kommentarzeilen (beginnend mit ; oder #)
        if ([string]::IsNullOrWhiteSpace($line) -or $line -match '^\s*[;#]') { continue }
    
        if ($line -match '^\[(.+)\]$') {
            $currentSection = $matches[1].Trim()
            if (-not $ini.Contains($currentSection)) {
                $ini[$currentSection] = [ordered]@{}
            }
        } elseif ($line -match '^(.*?)=(.*)$') {
            $key = $matches[1].Trim()
            $value = $matches[2].Trim()
            $ini[$currentSection][$key] = $value
        }
    }
    
    return $ini
}

$INIPath = Join-Path $ScriptDir "easyONBOARDINGConfig.ini"
try {
    $global:Config = Get-IniContent -Path $INIPath
}
catch {
    Write-Host "Fehler beim Laden der INI-Datei: $_"
    exit
}
#endregion

#region [Region 03 | FUNCTION DEFINITIONS]
# [Contains all helper and core functionality functions]
# [Enthält alle Helfer- und Kernfunktionalitäten]

#region [Region 03.1 | LOGGING FUNCTIONS]
# [Defines logging capabilities for different message levels]
# [Definiert Protokollierungsfunktionen für verschiedene Nachrichtenebenen]
function Write-LogMessage {
    # [03.1.1 - Primary logging wrapper for consistent message formatting]
    # [Primärer Logging-Wrapper für konsistente Nachrichtenformatierung]
    param(
        [string]$message,
        [string]$logLevel = "INFO"
    )
    Write-Log -message $message -logLevel $logLevel
}

function Write-DebugMessage {
    # [03.1.2 - Debug-specific logging with conditional execution]
    # [Debug-spezifische Protokollierung mit bedingter Ausführung]
    param(
        [Parameter(Mandatory=$false)]
        [AllowEmptyString()]
        [string]$Message
    )
    
    # PowerShell 7 optimized null check using null-conditional operator
    if ($global:Config?.General?.DebugMode -eq "1") {
        # Call Write-Log without capturing output to avoid pipeline return
        Write-Log -Message $Message -LogLevel "DEBUG"
    }
}
#endregion

# Log-Level (Standard ist "INFO"): "WARN", "ERROR", "DEBUG".
Write-Host "Logging-Funktion..."

#region [Region 03.2 | LOG FILE WRITER] 
# [Core logging function that writes messages to log files]
# [Kernfunktion zur Protokollierung von Nachrichten in Logdateien]
function Write-Log {
    # [03.2.1 - Low-level file logging implementation]
    # [Implementierung der Protokollierung auf Dateisystem-Ebene]
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true, ValueFromPipeline=$true)]
        [AllowEmptyString()]
        [string]$Message,
        
        [Parameter(Mandatory=$false)]
        [ValidateSet("INFO", "WARN", "ERROR", "DEBUG")]
        [string]$LogLevel = "INFO"
    )
    process {
        # log path determination using null-conditional and coalescing operators
        $logFilePath = $global:Config?.Logging?.LogPath ?? 
                      $global:Config?.General?.LogFilePath ?? 
                      (Join-Path $ScriptDir "Logs")
        
        # Ensure the log directory exists
        if (-not (Test-Path $logFilePath)) {
            try {
                # Use -Force to create parent directories if needed
                [void](New-Item -ItemType Directory -Path $logFilePath -Force -ErrorAction Stop)
            } catch {
                # Handle directory creation errors gracefully
                Write-Warning "Fehler beim Erstellen des Log-Verzeichnisses: $($_.Exception.Message)"
                # Try writing to a fallback location if possible
                $logFilePath = $env:TEMP
            }
        }

        # Define log file paths
        $logFile = Join-Path -Path $logFilePath -ChildPath "easyOnboarding.log"
        $errorLogFile = Join-Path -Path $logFilePath -ChildPath "easyOnboarding_error.log"
        
        # Generate timestamp and log entry
        $timeStamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        $logEntry = "$timeStamp [$LogLevel] $Message"

        try {
            Add-Content -Path $logFile -Value $logEntry -ErrorAction Stop
        } catch {
            # Bei Schreibfehler in der Logdatei, zusätzlichen Fehler loggen
            try {
                Add-Content -Path $errorLogFile -Value "$timeStamp [ERROR] Fehler beim Schreiben in Logdatei: $($_.Exception.Message)" -ErrorAction SilentlyContinue
            } catch {
                # Letzter Versuch: Ausgabe auf der Konsole
                Write-Warning "Kritischer Fehler: Kann in keine Logdatei schreiben: $($_.Exception.Message)"
            }
        }

        # Nur bei aktivem Debug-Modus auf der Konsole ausgeben, aber keinen Output zurückgeben
        if ($null -ne $global:Config -and $null -ne $global:Config.General -and $global:Config.General.DebugMode -eq "1") {
            Write-Host "Debug: $logEntry" -ForegroundColor Cyan
        }
        
        # Keine Rückgabe, um unerwünschte Pipeline-Ausgaben zu vermeiden
    }
}
#endregion

Write-DebugMessage "INI, LOG, Module, etc. regions initialized."

#region [Region 04 | ACTIVE DIRECTORY MODULE]
# [Loads the Active Directory PowerShell module required for user management]
# [Lädt das Active Directory PowerShell-Modul, das für die Benutzerverwaltung erforderlich ist]
Write-DebugMessage "Loading Active Directory module."
try {
    Write-DebugMessage "AD-Module geladen"
    Import-Module ActiveDirectory -SkipEditionCheck -ErrorAction Stop
    Write-DebugMessage "Active Directory module loaded successfully."
} catch {
    Write-Log -Message "Modul ActiveDirectory konnte nicht geladen werden: $($_.Exception.Message)" -LogLevel "ERROR"
    Throw "Kritischer Fehler: ActiveDirectory-Modul fehlt!"
}
#endregion

#region [Region 05 | WPF ASSEMBLIES]
# [Loads required WPF assemblies for the GUI interface]
# [Lädt erforderliche WPF-Assemblies für die GUI-Schnittstelle]
Write-DebugMessage "Loading WPF assemblies."
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Windows.Forms
Write-DebugMessage "WPF assemblies loaded successfully."
#endregion

Write-DebugMessage "Determining XAML file path."

#region [Region 06 | XAML LOADING]
# [Determines XAML file path and loads the GUI definition]
# [Ermittelt den XAML-Dateipfad und lädt die GUI-Definition]
Write-DebugMessage "Loading XAML file: $xamlPath"
$xamlPath = Join-Path $ScriptDir "MainGUI.xaml"
if (-not (Test-Path $xamlPath)) {
    Write-Error "Die XAML-Datei wurde nicht gefunden: $xamlPath"
    exit
}
try {
    [xml]$xaml = Get-Content -Path $xamlPath
    $reader = New-Object System.Xml.XmlNodeReader $xaml
    $window = [Windows.Markup.XamlReader]::Load($reader)
    if (-not $window) {
        Throw "XAML konnte nicht geladen werden."
    }
    Write-DebugMessage "XAML file loaded successfully."
    if ($global:Config.General.DebugMode -eq "1") {
        Write-Log "XAML-Datei erfolgreich geladen." "DEBUG"
    }
}
catch {
    Write-Error "Fehler beim Laden der XAML-Datei. Bitte überprüfe den Inhalt der Datei. $_"
    exit
}
#endregion

Write-DebugMessage "XAML file loaded"

#region [Region 07 | PASSWORD MANAGEMENT]
# [Functions for setting and removing password change restrictions]
# [Funktionen zum Festlegen und Entfernen von Passwortänderungsbeschränkungen]

#region [Region 07.1 | PREVENT PASSWORD CHANGE]
# [Sets ACL restrictions to prevent users from changing their passwords]
# [Setzt ACL-Einschränkungen, um Benutzer daran zu hindern, ihre Passwörter zu ändern]
Write-DebugMessage "Defining Set-CannotChangePassword function."
function Set-CannotChangePassword {
    # [07.1.1 - Modifies ACL settings to deny password change permissions]
    # [Ändert ACL-Einstellungen, um Passwortänderungsberechtigungen zu verweigern]
    param(
        [Parameter(Mandatory=$true)]
        [string]$SamAccountName
    )

    try {
        $adUser = Get-ADUser -Identity $SamAccountName -Properties DistinguishedName -ErrorAction Stop
        if (-not $adUser) {
            Write-Warning "Benutzer $SamAccountName wurde nicht gefunden."
            return
        }

        $user = [ADSI]"LDAP://$($adUser.DistinguishedName)"
        $acl = $user.psbase.ObjectSecurity

        Write-DebugMessage "Set-CannotChangePassword: AccessRule definieren"
        # AccessRule definieren: SELF darf 'Change Password' nicht
        $denyRule = New-Object System.DirectoryServices.ActiveDirectoryAccessRule(
            [System.Security.Principal.NTAccount]"NT AUTHORITY\\SELF",
            [System.DirectoryServices.ActiveDirectoryRights]::WriteProperty,
            [System.Security.AccessControl.AccessControlType]::Deny,
            [GUID]"ab721a53-1e2f-11d0-9819-00aa0040529b"  # GUID für 'User-Change-Password'
        )
        $acl.AddAccessRule($denyRule)
        $user.psbase.ObjectSecurity = $acl
        $user.psbase.CommitChanges()
        Write-Log "Prevent Password Change wurde für $SamAccountName gesetzt." "DEBUG"
    }
    catch {
        Write-Warning "Fehler beim Setzen der Passwortänderungsbeschränkung"
    }
}
#endregion

Write-DebugMessage "Defining Remove-CannotChangePassword function."

#region [Region 07.2 | ALLOW PASSWORD CHANGE]
# [Removes ACL restrictions to allow users to change their passwords]
# [Entfernt ACL-Einschränkungen, um Benutzern das Ändern ihrer Passwörter zu ermöglichen]
Write-DebugMessage "Removing CannotChangePassword for $SamAccountName."
function Remove-CannotChangePassword {
    # [07.2.1 - Removes deny rules from user ACL for password change permission]
    # [Entfernt Deny-Regeln aus der Benutzer-ACL für Passwortänderungsberechtigungen]
    param(
        [Parameter(Mandatory=$true)]
        [string]$SamAccountName
    )

    $adUser = Get-ADUser -Identity $SamAccountName -Properties DistinguishedName
    if (-not $adUser) {
        Write-Warning "Benutzer $SamAccountName wurde nicht gefunden."
        return
    }

    $user = [ADSI]"LDAP://$($adUser.DistinguishedName)"
    $acl = $user.psbase.ObjectSecurity

    Write-DebugMessage "Remove-CannotChangePassword: Alle Deny-Regeln entfernen"
    # Alle Deny-Regeln entfernen, die SELF betreffen und die GUID ab721a53-1e2f-11d0-9819-00aa0040529b haben
    $rulesToRemove = $acl.Access | Where-Object {
        $_.IdentityReference -eq "NT AUTHORITY\\SELF" -and
        $_.AccessControlType -eq 'Deny' -and
        $_.ObjectType -eq "ab721a53-1e2f-11d0-9819-00aa0040529b"
    }
    foreach ($rule in $rulesToRemove) {
        $acl.RemoveAccessRule($rule) | Out-Null
    }
    $user.psbase.ObjectSecurity = $acl
    $user.psbase.CommitChanges()
    Write-Log "Prevent Password Change wurde für $SamAccountName aufgehoben." "DEBUG"
}
#endregion
#endregion

Write-DebugMessage "Defining UPN generation function."

#region [Region 08 | UPN GENERATION]
# [Creates user principal names based on templates and user data]
# [Erstellt Benutzerprinzipalnamen basierend auf Vorlagen und Benutzerdaten]
Write-DebugMessage "Generating UPN for user."
function New-UPN {
    param(
        [pscustomobject]$userData,
        [hashtable]$Config
    )

    # Prüfe, ob FirstName und LastName vorhanden sind
    if ([string]::IsNullOrWhiteSpace($userData.FirstName) -or [string]::IsNullOrWhiteSpace($userData.LastName)) {
        Throw "Error: FirstName und LastName müssen gesetzt sein!"
    }

    # 1) SamAccountName generieren (erster Buchstabe des Vornamens + ganzer Nachname, alles in Kleinbuchstaben)
    $SamAccountName = ''
    if ($userData.FirstName.Length -gt 0) {
        $SamAccountName = ($userData.FirstName.Substring(0,1) + $userData.LastName).ToLower()
    } else {
        $SamAccountName = $userData.LastName.ToLower()
    }
    Write-DebugMessage "SamAccountName= $SamAccountName"

    # 2) Falls ein manueller UPN eingegeben wurde, diesen sofort verwenden
    if (-not [string]::IsNullOrWhiteSpace($userData.UPNEntered)) {
        return @{
            SamAccountName = $SamAccountName
            UPN            = $userData.UPNEntered
            CompanySection = "Company"
        }
    }

    # 3) Ermitteln des Company‑Abschnitts (Standard: "Company")
    $companySection = "Company"  # Default value
    
    # Optimierte Prüfung auf CompanySection mit vereinfachter Null-Wert-Prüfung
    if ($null -ne $userData.CompanySection) {
        # Fall 1: CompanySection ist direkt ein String
        if ($userData.CompanySection -is [string] -and -not [string]::IsNullOrWhiteSpace($userData.CompanySection)) {
            $companySection = $userData.CompanySection
        }
        # Fall 2: CompanySection ist ein Objekt mit Section-Eigenschaft
        elseif ($userData.CompanySection -is [PSObject] -and 
                $null -ne ($userData.CompanySection.PSObject.Properties.Match('Section') | Select-Object -First 1) -and
                -not [string]::IsNullOrWhiteSpace($userData.CompanySection.Section)) {
            $companySection = $userData.CompanySection.Section
        }
    }
    Write-DebugMessage "Verwendeter Company-Abschnitt: '$companySection'"

    # 4) Prüfen, ob der gewünschte Abschnitt in der INI existiert mit verbesserter Fehlerbehandlung
    if ($null -eq $Config) {
        Throw "Error: Config-Objekt ist NULL! Keine Konfiguration verfügbar."
    }
    
    if (-not $Config.Contains($companySection)) {
        Write-DebugMessage "Abschnitt '$companySection' nicht in Config gefunden, verwende 'Company'"
        $companySection = "Company" # Fallback auf Standardwert
        if (-not $Config.Contains($companySection)) {
            Throw "Error: Weder der angeforderte Abschnitt noch der Standard-Abschnitt 'Company' existiert in der INI!"
        }
    }
    $companyData = $Config[$companySection]
    $suffix = ($companySection -replace "\D","")

    # 5) Domain-Key ermitteln und Domain auslesen - mit verbesserten Null-Prüfungen
    $domainKey = "CompanyActiveDirectoryDomain$suffix"
    $fallbackDomainKey = "CompanyActiveDirectoryDomain"
    
    # Prüfe erst den spezifischen Key, dann den Fallback - mit präziseren Checks
    if ($null -ne $companyData -and 
        $companyData.Contains($domainKey) -and 
        -not [string]::IsNullOrWhiteSpace($companyData[$domainKey])) {
        $adDomain = "@" + $companyData[$domainKey].Trim()
    }
    elseif ($null -ne $companyData -and 
            $companyData.Contains($fallbackDomainKey) -and 
            -not [string]::IsNullOrWhiteSpace($companyData[$fallbackDomainKey])) {
        $adDomain = "@" + $companyData[$fallbackDomainKey].Trim()
        Write-DebugMessage "Verwende Fallback-Domain: $adDomain"
    }
    else {
        # Erweiterte Fehlermeldung mit mehr Details zum Problem
        $availableKeys = $null -ne $companyData ? ($companyData.Keys -join ", ") : "Keine Keys vorhanden"
        Throw "Error: Domain-Information fehlt in der INI! Weder '$domainKey' noch '$fallbackDomainKey' gefunden oder gültig. Verfügbare Keys: $availableKeys"
    }

    # 6) UPN-Template immer aus der INI verwenden (kein GUI-Dropdown) - mit optimiertem NULL-Check
    $upnTemplate = "FIRSTNAME.LASTNAME" # Sicherer Default-Wert
    
    # PowerShell 7 ?:-Operator für NULL-Prüfung und verbesserte Lesbarkeit
    $displayNameTemplates = $null -ne $Config -and $Config.Contains("DisplayNameUPNTemplates") ? 
        $Config.DisplayNameUPNTemplates : $null
    
    if ($null -ne $displayNameTemplates -and 
        $displayNameTemplates.Contains("DefaultUserPrincipalNameFormat") -and 
        -not [string]::IsNullOrWhiteSpace($displayNameTemplates["DefaultUserPrincipalNameFormat"])) {
        $upnTemplate = $displayNameTemplates["DefaultUserPrincipalNameFormat"].ToUpper()
        Write-DebugMessage "UPN-Template aus INI geladen: $upnTemplate"
    }
    else {
        Write-DebugMessage "Kein gültiges UPN-Template in der INI gefunden, verwende Standard '$upnTemplate'"
    }
    
    Write-DebugMessage "UPN-Template (aus INI): $upnTemplate"

    # 7) UPN anhand des Templates erzeugen - mit sicherer Substring-Verarbeitung und Null-Prüfungen
    # Stellen wir sicher, dass FirstName und LastName nicht null sind
    $firstName = if ($null -eq $userData.FirstName) { "" } else { $userData.FirstName }
    $lastName = if ($null -eq $userData.LastName) { "" } else { $userData.LastName }
    
    $UPN = ""
    switch ($upnTemplate) {
        "FIRSTNAME.LASTNAME"    { $UPN = "$firstName.$lastName".ToLower() + $adDomain }
        "LASTNAME.FIRSTNAME"    { $UPN = "$lastName.$firstName".ToLower() + $adDomain }
        "FIRSTINITIAL.LASTNAME" { 
            if (-not [string]::IsNullOrEmpty($firstName) -and $firstName.Length -gt 0) {
                $UPN = "$($firstName.Substring(0,1)).$lastName".ToLower() + $adDomain
            } else {
                $UPN = "x.$lastName".ToLower() + $adDomain
            }
        }
        "FIRSTNAME.LASTINITIAL" { 
            if (-not [string]::IsNullOrEmpty($lastName)) {
                $UPN = "$firstName.$($lastName.Substring(0,1))".ToLower() + $adDomain
            } else {
                $UPN = "$firstName.x".ToLower() + $adDomain
            }
        }
        "FIRSTNAME_LASTNAME"    { $UPN = "$firstName_$lastName".ToLower() + $adDomain }
        "LASTNAME_FIRSTNAME"    { $UPN = "$lastName_$firstName".ToLower() + $adDomain }
        "FIRSTINITIAL_LASTNAME" { 
            if (-not [string]::IsNullOrEmpty($firstName)) {
                $UPN = "$($firstName.Substring(0,1))_$lastName".ToLower() + $adDomain
            } else {
                $UPN = "x_$lastName".ToLower() + $adDomain
            }
        }
        "FIRSTNAME_LASTINITIAL" { 
            if (-not [string]::IsNullOrEmpty($lastName)) {
                $UPN = "$firstName_$($lastName.Substring(0,1))".ToLower() + $adDomain
            } else {
                $UPN = "$firstName_x".ToLower() + $adDomain
            }
        }
        "FIRSTNAMELASTNAME"     { $UPN = "$firstName$lastName".ToLower() + $adDomain }
        "LASTNAMEFIRSTNAME"     { $UPN = "$lastName$firstName".ToLower() + $adDomain }
        "FIRSTINITIALLASTNAME"  { 
            if (-not [string]::IsNullOrEmpty($firstName)) {
                $UPN = "$($firstName.Substring(0,1))$lastName".ToLower() + $adDomain
            } else {
                $UPN = "x$lastName".ToLower() + $adDomain
            }
        }
        "FIRSTNAMELASTINITIAL"  { 
            if (-not [string]::IsNullOrWhiteSpace($lastName)) {
                $UPN = "$firstName$($lastName.Substring(0,1))".ToLower() + $adDomain
            } else {
                $UPN = "$firstName x".ToLower() + $adDomain
            }
        }
        default                 { $UPN = "$firstName.$lastName".ToLower() + $adDomain }
    }

    # 8) UPN von Umlauten und Sonderzeichen bereinigen
    if ($UPN) {
        $UPN = $UPN -replace "ä", "ae" -replace "ö", "oe" -replace "ü", "ue" -replace "ß", "ss" -replace "Ä", "ae" -replace "Ö", "oe" -replace "Ü", "ue"
        $UPN = $UPN -replace "[^a-zA-Z0-9._@-]", ""
    } else {
        # Fallback bei leerem UPN
        $UPN = "user@" + $companyData[$domainKey ?? $fallbackDomainKey].Trim()
    }

    # 9) Ergebnis zurückgeben
    return @{
        SamAccountName = $SamAccountName
        UPN            = $UPN
        CompanySection = $companySection
    }
}
#endregion

#region [Region 09 | AD USER CREATION]
# [Creates Active Directory user accounts based on input data]
# [Erstellt Active Directory-Benutzerkonten basierend auf Eingabedaten]
function New-ADUserAccount {
    param (
        [Parameter(Mandatory=$true)]
        [pscustomobject]$UserData,
        [Parameter(Mandatory=$true)]
        [hashtable]$Config
    )
    
    try {
        # 1) UPN-Daten generieren
        $upnData = New-UPN -userData $UserData -Config $Config
        $samAccountName = $upnData.SamAccountName
        $userPrincipalName = $upnData.UPN
        $companySection = $upnData.CompanySection
        
        Write-DebugMessage "Ermittelter UPN: $userPrincipalName"
        Write-DebugMessage "Verwendete Company-Section: $companySection"
        
        # 2) Prüfen, ob alle benötigten Informationen vorhanden sind
        if (-not $Config.Contains($companySection)) {
            throw "Keine Konfiguration für Company-Section '$companySection' gefunden."
        }
        
        $suffix = ($companySection -replace "\D","") 
        $adOUKey = "CompanyActiveDirectoryOU$suffix"
        $adPath = $Config[$companySection][$adOUKey]
        
        if (-not $adPath) {
            throw "Kein AD-Pfad (OU) für Company-Section '$companySection' gefunden."
        }
        
        # 3) Generieren des DisplayName nach konfiguriertem Template (aus INI)
        $displayNameFormat = ""
        if ($Config.Contains("DisplayNameUPNTemplates") -and 
            $Config.DisplayNameUPNTemplates.Contains("DefaultDisplayNameFormat")) {
            $displayNameFormat = $Config.DisplayNameUPNTemplates.DefaultDisplayNameFormat
        }
        
        # Standardmäßig "LastName, FirstName"
        if ([string]::IsNullOrWhiteSpace($displayNameFormat)) {
            $displayNameFormat = "LastName, FirstName"
        }
        
        $displayName = switch -Wildcard ($displayNameFormat) {
            "LastName, FirstName" { "$($UserData.LastName), $($UserData.FirstName)" }
            "FirstName LastName"  { "$($UserData.FirstName) $($UserData.LastName)" }
            "LastName FirstName"  { "$($UserData.LastName) $($UserData.FirstName)" }
            default               { "$($UserData.LastName), $($UserData.FirstName)" }
        }
        
        # 4) Prüfen, ob Benutzer bereits existiert
        $userExists = $false
        try {
            $existingUser = Get-ADUser -Identity $samAccountName
            if ($existingUser) {
                $userExists = $true
                Write-Log "Benutzer $samAccountName existiert bereits." -LogLevel "WARN"
            }
        } catch {
            # Benutzer existiert nicht - das ist gut so
            $userExists = $false
        }
        
        if ($userExists) {
            return @{
                Success = $false
                Message = "Benutzer $samAccountName existiert bereits."
                SamAccountName = $samAccountName
            }
        }
        
        # 5) Generieren eines sicheren Passworts, falls nicht gesetzt
        if ([string]::IsNullOrWhiteSpace($UserData.Password)) {
            # Passwortlänge aus Config oder Standard 12 Zeichen
            $passwordLength = 12
            if ($Config.Contains("General") -and $Config.General.Contains("DefaultPasswordLength")) {
                $tempLength = [int]::TryParse($Config.General.DefaultPasswordLength, [ref]$passwordLength)
                if (-not $tempLength) { $passwordLength = 12 }
            }
            
            # Zeichenpool für Passwort
            $charPool = "abcdefghijkmnopqrstuvwxyzABCDEFGHJKLMNPQRSTUVWXYZ23456789!@#$%^&*_=+-"
            $securePassword = ""
            $random = New-Object System.Random
            
            # Mindestens je 1 Großbuchstabe, Kleinbuchstabe, Zahl und Sonderzeichen
            $securePassword += $charPool.Substring($random.Next(0, 25), 1)  # Kleinbuchstabe
            $securePassword += $charPool.Substring($random.Next(26, 50), 1) # Großbuchstabe
            $securePassword += $charPool.Substring($random.Next(51, 59), 1) # Zahl
            $securePassword += $charPool.Substring($random.Next(60, $charPool.Length-1), 1) # Sonderzeichen
            
            # Auffüllen bis zur gewünschten Länge
            for ($i = 4; $i -lt $passwordLength; $i++) {
                $securePassword += $charPool.Substring($random.Next(0, $charPool.Length-1), 1)
            }
            
            # Randomisieren der Zeichenreihenfolge
            $securePasswordArray = $securePassword.ToCharArray()
            $randomizedPassword = ""
            for ($i = $securePasswordArray.Count; $i -gt 0; $i--) {
                $randomPosition = $random.Next(0, $i)
                $randomizedPassword += $securePasswordArray[$randomPosition]
                $securePasswordArray = $securePasswordArray[0..($randomPosition-1)] + $securePasswordArray[($randomPosition+1)..($securePasswordArray.Count-1)]
            }
            
            $UserData.Password = $randomizedPassword
            Write-DebugMessage "Generiertes Passwort: $randomizedPassword"
        }
        
        # 6) Weitere AD-Attribute sammeln
        $adUserParams = @{
            SamAccountName = $samAccountName
            UserPrincipalName = $userPrincipalName
            Name = $displayName
            DisplayName = $displayName
            GivenName = $UserData.FirstName
            Surname = $UserData.LastName
            Path = $adPath
            AccountPassword = (ConvertTo-SecureString -String $UserData.Password -AsPlainText -Force)
            Enabled = $true
            PasswordNeverExpires = $false
            ChangePasswordAtLogon = $true
        }
        
        # Optional: Weitere Attribute, falls vorhanden
        if (-not [string]::IsNullOrWhiteSpace($UserData.Email)) {
            $adUserParams.EmailAddress = $UserData.Email
        }
        
        if (-not [string]::IsNullOrWhiteSpace($UserData.Description)) {
            $adUserParams.Description = $UserData.Description
        }
        
        if (-not [string]::IsNullOrWhiteSpace($UserData.Phone)) {
            $adUserParams.OfficePhone = $UserData.Phone
        }
        
        # 7) Benutzer erstellen
        Write-DebugMessage "Erstelle AD-Benutzer: $samAccountName"
        $newUser = New-ADUser @adUserParams -PassThru
        
        # 8) Optionale Attribute setzen
        # Falls bestimmte Attribute separat gesetzt werden müssen...
        
        # 9) Passwortoptionen setzen (falls konfiguriert)
        if ($UserData.PreventPasswordChange -eq $true) {
            Write-DebugMessage "Setze 'Prevent Password Change' für $samAccountName"
            Set-CannotChangePassword -SamAccountName $samAccountName
        }
        
        Write-Log "Benutzer $samAccountName wurde erfolgreich angelegt." -LogLevel "INFO"
        return @{
            Success = $true
            Message = "Benutzer wurde erfolgreich angelegt."
            SamAccountName = $samAccountName
            UserPrincipalName = $userPrincipalName
            Password = $UserData.Password
        }
    }
    catch {
        Write-Log "Fehler beim Anlegen des Benutzers: $($_.Exception.Message)" -LogLevel "ERROR"
        return @{
            Success = $false
            Message = "Fehler beim Anlegen des Benutzers: $($_.Exception.Message)"
            SamAccountName = $samAccountName
        }
    }
}
#endregion

Write-DebugMessage "Registering GUI event handlers."

#region [Region 09 | GUI EVENT HANDLER]
# [Function to register and handle GUI button click events with error handling]
# [Funktion zur Registrierung und Behandlung von GUI-Button-Klick-Ereignissen mit Fehlerbehandlung]
Write-DebugMessage "Registering GUI event handler for $Control."
function Register-GUIEvent {
    param (
        [Parameter(Mandatory=$true)]
        $Control,
        [Parameter(Mandatory=$true)]
        [ScriptBlock]$EventAction,
        [string]$ErrorMessagePrefix = "Fehler beim Ausführen des Ereignisses"
    )
    if ($Control) {
        $Control.Add_Click({
            try {
                & $EventAction
            }
            catch {
                [System.Windows.MessageBox]::Show("${ErrorMessagePrefix}: $($_.Exception.Message)", "Fehler", "OK", "Error")
            }
        })
    }
}
#endregion

Write-DebugMessage "Defining AD command execution function."

#region [Region 10 | AD COMMAND EXECUTION]
# [Consolidates error handling for Active Directory commands]
# [Konsolidiert die Fehlerbehandlung für Active Directory-Befehle]
Write-DebugMessage "Executing AD command."
function Invoke-ADCommand {
    param (
        [Parameter(Mandatory=$true)]
        [ScriptBlock]$Command,
        
        [Parameter(Mandatory=$true)]
        [string]$ErrorContext
    )
    try {
        & $Command
    }
    catch {
        $errMsg = "Fehler in Invoke-ADCommand"
        Write-Error $errMsg
        Throw $errMsg
    }
}
#endregion

Write-DebugMessage "Defining configuration value access function."

#region [Region 11 | CONFIGURATION VALUE ACCESS]
# [Helper function for safely accessing configuration values]
# [Hilfsfunktion für den sicheren Zugriff auf Konfigurationswerte]
Write-DebugMessage "Accessing configuration value: $Key."
function Get-ConfigValue {
    param (
        [Parameter(Mandatory = $true)]
        [hashtable]$Section,
        
        [Parameter(Mandatory = $true)]
        [string]$Key,
        
        [bool]$Mandatory = $true
    )
    
    if ([string]::IsNullOrWhiteSpace($Key)) {
        Throw "Error: Der übergebene Schlüssel ist null oder leer."
    }
    
    if ($Section.Contains($Key)) {
        return $Section[$Key]
    }
    elseif ($Mandatory) {
        Throw "Error: Der Schlüssel fehlt in der Konfiguration!"
    }
    return $null
}
#endregion

Write-DebugMessage "Defining template processing functions."

#region [Region 12 | TEMPLATE PROCESSING]
# [Functions to replace placeholders in templates with user data]
# [Funktionen zum Ersetzen von Platzhaltern in Vorlagen mit Benutzerdaten]
Write-DebugMessage "Resolving template placeholders."
function Resolve-TemplatePlaceholders {
    # [12.1 - Generic placeholder replacement for string templates]
    # [Generischer Platzhalterersatz für String-Vorlagen]
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Template,

        [Parameter(Mandatory = $true)]
        [ValidateNotNull()]
        [pscustomobject]$userData
    )
    # Ersetzt Platzhalter {first} und {last} mit den entsprechenden Werten aus $userData
    $result = $Template -replace '{first}', $userData.FirstName `
                          -replace '{last}', $userData.LastName
    return $result
}

Write-DebugMessage "Funktion zum Ersetzen der Platzhalter - UPN"
# Im UPN-Teil:
if (-not [string]::IsNullOrWhiteSpace($userData.UPNFormat)) {
    # Normalisiere das Template: trimmen und in Kleinbuchstaben umwandeln
    $upnTemplate = $userData.UPNFormat.Trim().ToLower()
    Write-DebugMessage "Invoke-Onboarding: UPN Format from userData: $upnTemplate"
    if ($upnTemplate -like "*{first}*") {
        # Dynamische Ersetzung der Platzhalter aus dem Template
        $upnBase = Resolve-TemplatePlaceholders -Template $upnTemplate -userData $userData
        $UPN = "$upnBase$adDomain"
    }
    else {
        # Feste Fälle als Fallback – hier kannst du bei Bedarf weitere Fälle ergänzen
        switch ($upnTemplate) {
            "firstname.lastname"    { $UPN = "$($userData.FirstName).$($userData.LastName)$adDomain" }
            "f.lastname"            { $UPN = "$($userData.FirstName.Substring(0,1)).$($userData.LastName)$adDomain" }
            "firstnamelastname"     { $UPN = "$($userData.FirstName)$($userData.LastName)$adDomain" }
            "flastname"             { $UPN = "$($userData.FirstName.Substring(0,1))$($userData.LastName)$adDomain" }
            Default                 { $UPN = "$SamAccountName$adDomain" }
        }
    }
}
else {
    Write-DebugMessage "No UPNFormat given, fallback to SamAccountName + domain"
    $UPN = "$SamAccountName$adDomain"
}
#endregion

Write-DebugMessage "Loading AD groups."

#region [Region 13 | AD GROUP LOADING]
# [Loads AD groups from the INI file and binds them to the GUI]
# [Lädt AD-Gruppen aus der INI-Datei und bindet sie an die GUI]
Write-DebugMessage "Loading AD groups from INI."
function Load-ADGroups 
{
    try {
        Write-Log "Lade AD-Gruppen aus der INI..." "DEBUG"
        if (-not $global:Config.Contains("ADGroups")) {
            Write-Error "Die [ADGroups]-Sektion fehlt in der INI."
            return
        }
        $adGroupsIni = $global:Config["ADGroups"]
        $groupList = @()

        # Nur Schlüssel berücksichtigen, die mit "ADGroup" beginnen und keine Label-Schlüssel sind
        foreach ($key in $adGroupsIni.Keys) {
            if ($key -match '^ADGroup\d+$') {
                $labelKey = "${key}_Label"
                if ($adGroupsIni.Contains($labelKey) -and -not [string]::IsNullOrWhiteSpace($adGroupsIni[$labelKey])) {
                    $groupList += $adGroupsIni[$labelKey]
                }
                else {
                    $groupList += $adGroupsIni[$key]
                }
            }
        }
        
        $icADGroups = $window.FindName("icADGroups")
        if ($null -eq $icADGroups) {
            Write-Error "ItemsControl 'icADGroups' nicht gefunden. Überprüfe den XAML-Namen."
            return
        }
        
        # Binde die Liste der Gruppen aus der INI an die ItemsSource des ItemsControl
        $icADGroups.ItemsSource = $groupList
        Write-Log "Gruppen aus der INI erfolgreich an ItemsControl gebunden." "DEBUG"
    }
    catch {
        Write-Error "Fehler beim Laden der AD-Gruppen aus INI: $_"
    }
}
# Aufruf der Funktion nach dem Laden der XAML:
Load-ADGroups
#endregion

Write-DebugMessage "Populating dropdowns (OU, Location, License, MailSuffix, DisplayName Template)."

#region [Region 14 | DROPDOWN POPULATION]
# [Functions to populate various dropdown menus from configuration data]
# [Funktionen zum Befüllen verschiedener Dropdown-Menüs aus Konfigurationsdaten]

Write-DebugMessage "Dropdown: OU Refresh"

#region [Region 14.1 | OU DROPDOWN]
# [Populates the Organizational Unit dropdown with data from AD]
# [Befüllt das Dropdown für Organisationseinheiten mit Daten aus AD]
# --- OU Refresh --- 
# [14.1.1 - Fetches OUs from beneath the configured default OU]
# [Holt OUs unterhalb der konfigurierten Standard-OU]
Write-DebugMessage "Refreshing OU dropdown."
$btnRefreshOU = $window.FindName("btnRefreshOU")
if ($btnRefreshOU) {
    $btnRefreshOU.Add_Click({
        try {
            $defaultOUFromINI = $global:Config.General["DefaultOU"]
            $OUList = Get-ADOrganizationalUnit -Filter * -SearchBase $defaultOUFromINI |
                Select-Object -ExpandProperty DistinguishedName
                Write-DebugMessage "Gefundene OUs: $($OUList.Count)"
            $cmbOU = $window.FindName("cmbOU")
            if ($cmbOU) {
                # Zuerst ItemsSource auf null setzen
                $cmbOU.ItemsSource = $null
                # Dann die ItemsCollection leeren
                $cmbOU.Items.Clear()
                # Jetzt die neue ItemsSource zuweisen
                $cmbOU.ItemsSource = $OUList
                Write-DebugMessage "Dropdown: OU-DropDown erfolgreich befüllt."
            }
        }
        catch {
            [System.Windows.MessageBox]::Show("Fehler beim Laden der OUs: $($_.Exception.Message)", "Fehler")
        }
    })
}
#endregion

Write-DebugMessage "Dropdown: DisplayName Dropdown befüllen"

#region [Region 14.2 | DISPLAY NAME TEMPLATES]
# [Populates the display name template dropdown from INI configuration]
# [Befüllt das Dropdown für Anzeigenamen-Vorlagen aus der INI-Konfiguration]
Write-DebugMessage "Populating display name template dropdown."
$comboBoxDisplayTemplate = $window.FindName("cmbDisplayTemplate")

if ($comboBoxDisplayTemplate -and $global:Config.Contains("DisplayNameUPNTemplates")) {
    # Vorherige Items und Bindings löschen
    $comboBoxDisplayTemplate.Items.Clear()
    [System.Windows.Data.BindingOperations]::ClearBinding($comboBoxDisplayTemplate, [System.Windows.Controls.ComboBox]::ItemsSourceProperty)

    $DisplayNameTemplateList = @()
    $displayTemplates = $global:Config["DisplayNameUPNTemplates"]

    # Füge den Default-Eintrag hinzu (aus DefaultDisplayNameFormat)
    $defaultDisplayNameFormat = if ($displayTemplates.Contains("DefaultDisplayNameFormat")) {
        $displayTemplates["DefaultDisplayNameFormat"]
    } else {
        ""
    }
    
    Write-DebugMessage "Default DisplayName Format: $defaultDisplayNameFormat"

    if ($defaultDisplayNameFormat) {
        $DisplayNameTemplateList += [PSCustomObject]@{
            Name     = $defaultDisplayNameFormat
            Template = $defaultDisplayNameFormat
        }
    }

    # Füge alle weiteren DisplayName Templates hinzu (nur diejenigen, die mit "DisplayNameTemplate" beginnen)
    if ($displayTemplates -and $displayTemplates.Keys.Count -gt 0) {
        foreach ($key in $displayTemplates.Keys) {
            if ($key -like "DisplayNameTemplate*") {
                $pattern = $displayTemplates[$key]
                Write-DebugMessage "DisplayName Template gefunden: $pattern"
                $DisplayNameTemplateList += [PSCustomObject]@{
                    Name     = $pattern
                    Template = $pattern
                }
            }
        }
    }

    # Setze die ItemsSource, DisplayMemberPath und SelectedValuePath
    $comboBoxDisplayTemplate.ItemsSource = $DisplayNameTemplateList
    $comboBoxDisplayTemplate.DisplayMemberPath = "Name"
    $comboBoxDisplayTemplate.SelectedValuePath = "Template"

    # Setze den Default-Wert, falls vorhanden
    if ($defaultDisplayNameFormat) {
        $comboBoxDisplayTemplate.SelectedValue = $defaultDisplayNameFormat
    }
}
Write-DebugMessage "Display name template dropdown populated successfully."
#endregion

Write-DebugMessage "Dropdown: License Dropdown befüllen"

#region [Region 14.3 | LICENSE OPTIONS]
# [Populates the license dropdown with available license options]
# [Befüllt das Dropdown für Lizenzoptionen mit verfügbaren Lizenzen]
Write-DebugMessage "Populating license dropdown."
$comboBoxLicense = $window.FindName("cmbLicense")
if ($comboBoxLicense -and $global:Config.Contains("LicensesGroups")) {
    [System.Windows.Data.BindingOperations]::ClearBinding($comboBoxLicense, [System.Windows.Controls.ComboBox]::ItemsSourceProperty)
    $LicenseListFromINI = @()
    foreach ($licenseKey in $global:Config["LicensesGroups"].Keys) {
        $LicenseListFromINI += [PSCustomObject]@{
            Name  = $licenseKey -replace '^MS365_', ''
            Value = $licenseKey
        }
    }
    $comboBoxLicense.ItemsSource = $LicenseListFromINI
    if ($LicenseListFromINI.Count -gt 0) {
        $comboBoxLicense.SelectedValue = $LicenseListFromINI[0].Value
    }
}
Write-DebugMessage "License dropdown populated successfully."
#endregion

Write-DebugMessage "Dropdown: TLGroups Dropdown befüllen"

#region [Region 14.4 | TEAM LEADER GROUPS]
# [Populates the team leader group dropdown from configuration]
# [Befüllt das Dropdown für Teamleiter-Gruppen aus der Konfiguration]
Write-DebugMessage "Populating team leader group dropdown."
$comboBoxTLGroup = $window.FindName("cmbTLGroup")
if ($comboBoxTLGroup -and $global:Config.Contains("TLGroups")) {
    # Alte Bindings und Items löschen
    $comboBoxTLGroup.Items.Clear()
    [System.Windows.Data.BindingOperations]::ClearBinding($comboBoxTLGroup, [System.Windows.Controls.ComboBox]::ItemsSourceProperty)
    
    $TLGroupOptions = @()
    foreach ($key in $global:Config["TLGroups"].Keys) {
        # Hier wird der Schlüssel (z. B. DEV) als Anzeigetext genutzt,
        # während der zugehörige Wert (z. B. TEAM1TL) separat gespeichert wird.
        $TLGroupOptions += [PSCustomObject]@{
            Name  = $key
            Group = $global:Config["TLGroups"][$key]
        }
    }
    $comboBoxTLGroup.ItemsSource = $TLGroupOptions
    $comboBoxTLGroup.DisplayMemberPath = "Name"
    $comboBoxTLGroup.SelectedValuePath = "Group"
    if ($TLGroupOptions.Count -gt 0) {
        $comboBoxTLGroup.SelectedIndex = 0
    }
}
Write-DebugMessage "Team leader group dropdown populated successfully."
#endregion

Write-DebugMessage "Dropdown: MailSuffix Dropdown befüllen"

#region [Region 14.5 | EMAIL DOMAIN SUFFIXES]
# [Populates the email domain suffix dropdown from configuration]
# [Befüllt das Dropdown für E-Mail-Domain-Suffixe aus der Konfiguration]
Write-DebugMessage "Populating email domain suffix dropdown."
$comboBoxSuffix = $window.FindName("cmbSuffix")
if ($comboBoxSuffix -and $global:Config.Contains("MailEndungen")) {
    [System.Windows.Data.BindingOperations]::ClearBinding($comboBoxSuffix, [System.Windows.Controls.ComboBox]::ItemsSourceProperty)
    $MailSuffixListFromINI = @()
    foreach ($domainKey in $global:Config["MailEndungen"].Keys) {
        $domainValue = $global:Config["MailEndungen"][$domainKey]
        $MailSuffixListFromINI += [PSCustomObject]@{
            Key   = $domainValue
            Value = $domainValue
        }
    }
    $comboBoxSuffix.ItemsSource = $MailSuffixListFromINI
    # Standardwert aus dem [Company]-Abschnitt:
    $defaultSuffix = $global:Config.Company["CompanyMailDomain"]
    $comboBoxSuffix.SelectedValue = $defaultSuffix
}
Write-DebugMessage "Email domain suffix dropdown populated successfully."
#endregion

Write-DebugMessage "Dropdown: Location Dropdown befüllen"

#region [Region 14.6 | LOCATION OPTIONS]
# [Populates the location dropdown with available office locations]
# [Befüllt das Dropdown für Standorte mit verfügbaren Bürostandorten]
Write-DebugMessage "Populating location dropdown."
if ($global:Config.Contains("STANDORTE")) {
    $locationList = @()
    foreach ($key in $global:Config["STANDORTE"].Keys) {
        if ($key -match '^(STANDORTE_\d+)$') {
            $bezKey = $key + "_Bez"
            $locationObj = [PSCustomObject]@{
                Key = $global:Config["STANDORTE"][$key]
                Bez = if ($global:Config["STANDORTE"].Contains($bezKey)) { $global:Config["STANDORTE"][$bezKey] } else { $global:Config["STANDORTE"][$key] }
            }
            $locationList += $locationObj
        }
    }
    $global:LocationListFromINI = $locationList
    $cmbLocation = $window.FindName("cmbLocation")
    if ($cmbLocation -and $global:LocationListFromINI) {
        [System.Windows.Data.BindingOperations]::ClearBinding($cmbLocation, [System.Windows.Controls.ComboBox]::ItemsSourceProperty)
        $cmbLocation.ItemsSource = $global:LocationListFromINI
        if ($global:LocationListFromINI.Count -gt 0) {
            $cmbLocation.SelectedIndex = 0
        }
    }
    Write-DebugMessage "Location dropdown populated successfully."
}
#endregion
#endregion

Write-DebugMessage "Defining utility functions."

#region [Region 15 | UTILITY FUNCTIONS]
# [Miscellaneous utility functions for the application]
# [Verschiedene Hilfsfunktionen für die Anwendung]

Write-DebugMessage "Set-Logo"

#region [Region 15.1 | LOGO MANAGEMENT]
# [Function to handle logo uploads and management for reports]
# [Funktion zur Verwaltung von Logo-Uploads für Berichte]
Write-DebugMessage "Setting logo."
function Set-Logo {
    # [15.1.1 - Handles file selection and saving of logo images]
    # [Behandelt Dateiauswahl und Speichern von Logo-Bildern]
    param (
        [hashtable]$brandingConfig
    )
    try {
        Write-DebugMessage "GUI Öffne Datei-Auswahldialog für Logo-Upload"
        $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
        $openFileDialog.Filter = "Bilddateien (*.jpg;*.png;*.bmp)|*.jpg;*.jpeg;*.png;*.bmp"
        $openFileDialog.Title = "Wähle ein Logo für das Onboarding-Dokument"
        if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $selectedFilePath = $openFileDialog.FileName
            Write-Log "Ausgewählte Datei: $selectedFilePath" "DEBUG"

            if (-not $brandingConfig.Contains("TemplateLogo") -or [string]::IsNullOrWhiteSpace($brandingConfig["TemplateLogo"])) {
                Throw "Kein 'TemplateLogo' in der Branding-Report-Sektion definiert."
            }

            $TemplateLogo = $brandingConfig["TemplateLogo"]
            if (-not (Test-Path $TemplateLogo)) {
                try {
                    New-Item -ItemType Directory -Path $TemplateLogo -Force -ErrorAction Stop | Out-Null
                } catch {
                    Throw "Konnte das Zielverzeichnis für das Logo nicht erstellen: $($_.Exception.Message)"
                }
            }

            $targetLogoTemplate = Join-Path -Path $TemplateLogo -ChildPath "Onboarding_Logo.png"
            Copy-Item -Path $selectedFilePath -Destination $targetLogoTemplate -Force -ErrorAction Stop
            Write-Log "Logo erfolgreich gespeichert unter: $targetLogoTemplate" "DEBUG"
            [System.Windows.MessageBox]::Show("Das Logo wurde erfolgreich gespeichert!nSpeicherort: $targetLogoTemplate", "Erfolg", "OK", "Information")
        }
    } catch {
        [System.Windows.MessageBox]::Show("Fehler beim Hochladen des Logos: $($_.Exception.Message)", "Fehler", "OK", "Error")
    }
}
#endregion

Write-DebugMessage "Defining advanced password generation function."

#region [Region 15.2 | PASSWORD GENERATION]
# [Advanced password generation with security requirements]
# [Erweiterte Passwortgenerierung mit Sicherheitsanforderungen]
Write-DebugMessage "Generating advanced password."
function New-AdvancedPassword {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$false)]
        [int]$Length = 12,

        [Parameter(Mandatory=$false)]
        [int]$MinUpperCase = 2,

        [Parameter(Mandatory=$false)]
        [int]$MinDigits = 2,

        [Parameter(Mandatory=$false)]
        [bool]$AvoidAmbiguous = $false,

        [Parameter(Mandatory=$false)]
        [bool]$IncludeSpecial = $true,

        [Parameter(Mandatory=$false)]
        [ValidateRange(1,10)]
        [int]$MinNonAlpha = 2
    )

    if ($Length -lt ($MinUpperCase + $MinDigits)) {
        Throw "Error: Die Passwortlänge ($Length) ist zu kurz für die geforderten Mindestwerte (MinUpperCase + MinDigits = $($MinUpperCase + $MinDigits))."
    }

    Write-DebugMessage "New-AdvancedPassword: Zeichenpools definieren"
    $lower = 'abcdefghijklmnopqrstuvwxyz'
    $upper = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
    $digits = '0123456789'
    $special = '!@#$%^&*()'
    
    Write-DebugMessage "New-AdvancedPassword: Entferne ambigue Zeichen, falls gewünscht"
    if ($AvoidAmbiguous) {
        $ambiguous = 'Il1O0'
        $upper = -join ($upper.ToCharArray() | Where-Object { $ambiguous -notcontains $_ })
        $lower = -join ($lower.ToCharArray() | Where-Object { $ambiguous -notcontains $_ })
        $digits = -join ($digits.ToCharArray() | Where-Object { $ambiguous -notcontains $_ })
        $special = -join ($special.ToCharArray() | Where-Object { $ambiguous -notcontains $_ })
    }
    
    # Ensure that character pools are never empty
    if (-not $upper -or $upper.Length -eq 0) { $upper = 'ABCDEFGHJKLMNPQRSTUVWXYZ' }
    if (-not $lower -or $lower.Length -eq 0) { $lower = 'abcdefghijkmnopqrstuvwxyz' }
    if (-not $digits -or $digits.Length -eq 0) { $digits = '23456789' }
    if (-not $special -or $special.Length -eq 0) { $special = '!@#$%^&*()' }
    
    # Recalculate 'all' after ensuring pools are not empty
    $all = $lower + $upper + $digits
    if ($IncludeSpecial) { $all += $special }
    
    do {
        Write-DebugMessage "New-AdvancedPassword: Beginne Passwort-Generierung"
        # Initialize as a simple string array
        $passwordChars = [System.Collections.ArrayList]::new()
        
        Write-DebugMessage "New-AdvancedPassword: Füge Mindestanzahl an Großbuchstaben hinzu"
        for ($i = 0; $i -lt $MinUpperCase; $i++) {
            [void]$passwordChars.Add($upper[(Get-Random -Minimum 0 -Maximum $upper.Length)].ToString())
        }
        
        Write-DebugMessage "New-AdvancedPassword: Füge Mindestanzahl an Ziffern hinzu"
        for ($i = 0; $i -lt $MinDigits; $i++) {
            [void]$passwordChars.Add($digits[(Get-Random -Minimum 0 -Maximum $digits.Length)].ToString())
        }
        
        Write-DebugMessage "New-AdvancedPassword: Fülle bis zur gewünschten Länge auf"
        while ($passwordChars.Count -lt $Length) {
            [void]$passwordChars.Add($all[(Get-Random -Minimum 0 -Maximum $all.Length)].ToString())
        }
        
        Write-DebugMessage "New-AdvancedPassword: Zufällige Reihenfolge"
        # Get array of strings, then join them at the end
        $shuffledChars = $passwordChars | Get-Random -Count $passwordChars.Count
        $generatedPassword = -join $shuffledChars
        
        Write-DebugMessage "New-AdvancedPassword: Überprüfe Mindestanzahl nicht alphabetischer Zeichen"
        # Count characters that don't match letters
        $nonAlphaCount = ($generatedPassword.ToCharArray() | Where-Object { $_ -notmatch '[a-zA-Z]' }).Count
        
    } while ($nonAlphaCount -lt $MinNonAlpha)
    
    Write-DebugMessage "Advanced password generated successfully."
    return $generatedPassword
}
#endregion

#region [Region 16 | CORE ONBOARDING PROCESS]
# [Main function that handles the complete user onboarding process]
# [Hauptfunktion, die den gesamten Benutzer-Onboarding-Prozess verarbeitet]
Write-DebugMessage "Starting onboarding process."
function Invoke-Onboarding {
    # [16.1 - Processes AD user creation and configuration in a single workflow]
    # [Verarbeitet AD-Benutzererstellung und -konfiguration in einem einzigen Workflow]
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [pscustomobject]$userData,

        [Parameter(Mandatory = $true)]
        [hashtable]$Config
    )

    Write-DebugMessage "Invoke-Onboarding: Start"

    # Debugging information
    if ($global:Config.General.DebugMode -eq "1") {
        Write-DebugMessage "[DEBUG] UPNFormat aus INI geladen: $($Config.DisplayNameUPNTemplates.DefaultUserPrincipalNameFormat)"
        Write-DebugMessage "[DEBUG] userData: $($userData | ConvertTo-Json -Depth 1 -Compress)"
    }

    # [16.1.1 - ActiveDirectory Module Verification]
    # [Überprüft, ob das ActiveDirectory-Modul geladen ist, bevor mit Operationen fortgefahren wird]
    Write-DebugMessage "Invoke-Onboarding: Überprüfe, ob das ActiveDirectory-Modul bereits geladen ist"
    if (-not (Get-Module -Name ActiveDirectory)) {
        Throw "ActiveDirectory module is not loaded. Bitte sicherstellen, dass es installiert und vorab importiert wurde."
    }

    # [16.1.2 - Check mandatory fields: FirstName]
    Write-DebugMessage "Invoke-Onboarding: Checking mandatory fields"
    if ([string]::IsNullOrWhiteSpace($userData.FirstName)) {
        Throw "First Name is required!"
    }
    if ([string]::IsNullOrWhiteSpace($userData.LastName)) {
        Throw "Last Name is required!"
    }
    
    # Validierung der OU
    $selectedOU = if (-not [string]::IsNullOrWhiteSpace($userData.OU)) {
        $userData.OU
    } else {
        if (-not $Config.General.Contains("DefaultOU") -or [string]::IsNullOrWhiteSpace($Config.General["DefaultOU"])) {
            Throw "Error: Keine OU angegeben und keine DefaultOU in der Konfiguration gefunden!"
        }
        $Config.General["DefaultOU"]
    }
    Write-DebugMessage "Ausgewählte OU: $selectedOU"

    # Debug-Ausgaben mit Null-Prüfungen
    Write-DebugMessage "DEBUG ZEILE 898: Überprüfung aller wichtigen Variablen"
    $SamAccountName = $null
    $ADGroupsSelected = if ($userData.PSObject.Properties.Match("ADGroupsSelected").Count -gt 0) { $userData.ADGroupsSelected } else { "KEINE" }
    $License = if ($userData.PSObject.Properties.Match("License").Count -gt 0) { $userData.License } else { "Keine" }
    
    Write-DebugMessage "SamAccountName: Wird generiert"
    Write-DebugMessage "UPNFormat: $($userData.UPNFormat ?? 'Nicht definiert')"
    Write-DebugMessage "MailSuffix: $($userData.MailSuffix ?? 'Nicht definiert')"
    Write-DebugMessage "ADGroupsSelected: $ADGroupsSelected"
    Write-DebugMessage "License: $License"
    
    # [16.1.3 - Password generation or fixed password]
    Write-DebugMessage "Invoke-Onboarding: Determining main password"
    $mainPW = $null
    if ($userData.PasswordMode -eq 1) {
        # [16.1.3.1 - Use advanced generation]
        Write-DebugMessage "Invoke-Onboarding: Checking [PasswordFixGenerate] keys"
        foreach ($key in @("PasswordLaenge","IncludeSpecialChars","AvoidAmbiguousChars","MinNonAlpha","MinUpperCase","MinDigits")) {
            if (-not $Config.PasswordFixGenerate.Contains($key)) {
                Throw "Error: '$key' is missing in [PasswordFixGenerate]!"
            }
        }
        Write-DebugMessage "Invoke-Onboarding: Reading password settings from INI"
        $includeSpecial = ($Config.PasswordFixGenerate["IncludeSpecialChars"] -match "^(?i:true|1)$")
        $avoidAmbiguous = ($Config.PasswordFixGenerate["AvoidAmbiguousChars"] -match "^(?i:true|1)$")

        Write-DebugMessage "Invoke-Onboarding: Generating password with New-AdvancedPassword"
        $mainPW = New-AdvancedPassword `
            -Length ([int]$Config.PasswordFixGenerate["PasswordLaenge"]) `
            -IncludeSpecial $includeSpecial `
            -AvoidAmbiguous $avoidAmbiguous `
            -MinUpperCase ([int]$Config.PasswordFixGenerate["MinUpperCase"]) `
            -MinDigits ([int]$Config.PasswordFixGenerate["MinDigits"]) `
            -MinNonAlpha ([int]$Config.PasswordFixGenerate["MinNonAlpha"])
        
        if ([string]::IsNullOrWhiteSpace($mainPW)) {
            Throw "Error: Generated password is empty. Please check [PasswordFixGenerate]."
        }
    }
    else {
        # [16.1.3.2 - Use fixPassword from INI]
        Write-DebugMessage "Invoke-Onboarding: Using fixPassword from INI"
        if (-not $Config.PasswordFixGenerate.Contains("fixPassword")) {
            Throw "Error: fixPassword is missing in PasswordFixGenerate!"
        }
        $mainPW = $Config.PasswordFixGenerate["fixPassword"]
        if ([string]::IsNullOrWhiteSpace($mainPW)) {
            Throw "Error: No fixed password provided in the INI!"
        }
    }

    # [16.1.3.3 - Convert to secure]
    $SecurePW = ConvertTo-SecureString $mainPW -AsPlainText -Force

    # [16.1.4 - Generate additional (custom) passwords (5 total)]
    Write-DebugMessage "Invoke-Onboarding: Generating additional custom passwords"
    $customPWLabels = if ($Config.Contains("CustomPWLabels")) { $Config["CustomPWLabels"] } else { @{} }
    # PowerShell 7+ null-coalescing Operator (??) durch alternative Methode ersetzt für PowerShell 5 Kompatibilität
    $pwLabel1 = if ($customPWLabels.Contains("CustomPW1_Label")) { $customPWLabels["CustomPW1_Label"] } else { "Custom PW #1" }
    $pwLabel2 = if ($customPWLabels.Contains("CustomPW2_Label")) { $customPWLabels["CustomPW2_Label"] } else { "Custom PW #2" }
    $pwLabel3 = if ($customPWLabels.Contains("CustomPW3_Label")) { $customPWLabels["CustomPW3_Label"] } else { "Custom PW #3" }
    $pwLabel4 = if ($customPWLabels.Contains("CustomPW4_Label")) { $customPWLabels["CustomPW4_Label"] } else { "Custom PW #4" }
    $pwLabel5 = if ($customPWLabels.Contains("CustomPW5_Label")) { $customPWLabels["CustomPW5_Label"] } else { "Custom PW #5" }

    $CustomPW1 = New-AdvancedPassword
    $CustomPW2 = New-AdvancedPassword
    $CustomPW3 = New-AdvancedPassword
    $CustomPW4 = New-AdvancedPassword
    $CustomPW5 = New-AdvancedPassword

    # [16.1.5 - SamAccountName and UPN creation logic]
    Write-DebugMessage "Invoke-Onboarding: Creating SamAccountName and UPN"
    # Nutze ausschließlich den Rückgabewert der New-UPN-Funktion:
    $upnResult = New-UPN -userData $userData -Config $Config
    $SamAccountName = $upnResult.SamAccountName
    $UPN = $upnResult.UPN
    $companySection = $upnResult.CompanySection
    
    Write-DebugMessage "Invoke-Onboarding: Reading [Company] data from config using section '$companySection'"
    if (-not $Config.Contains($companySection)) {
        Throw "Error: Section '$companySection' is missing in the INI!"
    }
    $companyData = $Config[$companySection]   

    # [16.1.5.1 - Enhanced logging after UPN creation]
    Write-DebugMessage "Generated SamAccountName: $SamAccountName"
    Write-DebugMessage "Generated UPN: $UPN"
    Write-DebugMessage "Using company section: $companySection"
    
    # [16.1.5.2 - Additional validation based on configuration]
    if ($Config.Contains("ValidateUPN") -and $Config.ValidateUPN.Contains("CheckForDuplicates") -and 
        $Config.ValidateUPN["CheckForDuplicates"] -eq "1") {
        
        Write-DebugMessage "Checking for UPN duplicates: $UPN"
        try {
            $existingUser = Get-ADUser -Filter {UserPrincipalName -eq $UPN} -ErrorAction Stop
            if ($existingUser) {
                Write-DebugMessage "WARNING: UPN $UPN is already in use by $($existingUser.SamAccountName)!"
                
                if ($Config.ValidateUPN.Contains("AppendRandomOnConflict") -and 
                    $Config.ValidateUPN["AppendRandomOnConflict"] -eq "1") {
                    
                    $random = -join ((48..57) + (97..122) | Get-Random -Count 3 | ForEach-Object {[char]$_})
                    $UPNBase = $UPN.Split('@')[0]
                    $UPNDomain = $UPN.Split('@')[1]
                    $UPN = "$UPNBase$random@$UPNDomain"
                    Write-DebugMessage "UPN modified to avoid conflict: $UPN"
                }
                else {
                    Throw "Error: UPN $UPN is already in use and conflict resolution is disabled!"
                }
            }
        }
        catch {
            if ($_.Exception.Message -notmatch "No user matches filter criteria") {
                Throw "Error checking for duplicate UPN: $($_.Exception.Message)"
            }
        }
    }
    
    # [16.1.5.3 - Apply name normalization for special characters if configured]
    if ($Config.Contains("NameNormalization") -and $Config.NameNormalization.Contains("Enabled") -and 
        $Config.NameNormalization["Enabled"] -eq "1") {
        
        Write-DebugMessage "Applying name normalization rules"
        
        # Handle special characters in names based on configuration
        if ($Config.NameNormalization.Contains("ReplaceSpecialChars")) {
            $replacements = $Config.NameNormalization["ReplaceSpecialChars"].Split(';')
            foreach ($replacement in $replacements) {
                if ($replacement -match '^(.+?)=(.+?)$') {
                    $from = $matches[1]
                    $to = $matches[2]
                    $UPN = $UPN -replace [regex]::Escape($from), $to
                    $SamAccountName = $SamAccountName -replace [regex]::Escape($from), $to
                    Write-DebugMessage "Replaced '$from' with '$to' in identifiers"
                }
            }
        }
        
        # Additional character normalization
        if ($Config.NameNormalization.Contains("NormalizeCase") -and 
            $Config.NameNormalization["NormalizeCase"] -eq "1") {
            $SamAccountName = $SamAccountName.ToLower()
            $UPNPrefix = $UPN.Split('@')[0]
            $UPNDomain = $UPN.Split('@')[1]
            $UPN = "$($UPNPrefix.ToLower())@$UPNDomain"
            Write-DebugMessage "Normalized case: $SamAccountName, $UPN"
        }
    }
    
    # [16.1.5.4 - Initialize attributes collection for later use]
    $otherAttributes = @{}
    
    # [16.1.5.5 - Check for custom attribute mappings]
    if ($Config.Contains("CustomAttributeMappings")) {
        Write-DebugMessage "Processing custom attribute mappings"
        foreach ($key in $Config.CustomAttributeMappings.Keys) {
            $attributeName = $key
            $userDataProperty = $Config.CustomAttributeMappings[$key]
            
            if ($userData.PSObject.Properties.Match($userDataProperty).Count -gt 0 -and 
                -not [string]::IsNullOrWhiteSpace($userData.$userDataProperty)) {
                $otherAttributes[$attributeName] = $userData.$userDataProperty
                Write-DebugMessage "Mapped custom attribute: $attributeName = $($userData.$userDataProperty)"
            }
        }
    }
    
    # [16.1.5.6 - Record normalized values back to userData]
    $userData | Add-Member -MemberType NoteProperty -Name "NormalizedSamAccountName" -Value $SamAccountName -Force
    $userData | Add-Member -MemberType NoteProperty -Name "NormalizedUPN" -Value $UPN -Force
    
    Write-DebugMessage "UPN and SamAccountName processing complete"
    # [16.1.6 - Read company data from [Company*] in the INI]
    Write-DebugMessage "Invoke-Onboarding: Reading [Company] data from config"
    if (-not $Config.Contains($companySection)) {
        Throw "Error: Section '$companySection' is missing in the INI!"
    }
    $companyData = $Config[$companySection]
    $suffix = ($companySection -replace "\D","")

    # [16.1.6.1 - Basic address info from company]
    $streetKey = "CompanyStrasse"
    $plzKey    = "CompanyPLZ"
    $ortKey    = "CompanyOrt"
    $Street = Get-ConfigValue -Section $companyData -Key "CompanyStrasse"
    $Zip    = Get-ConfigValue -Section $companyData -Key "CompanyPLZ"
    $City   = Get-ConfigValue -Section $companyData -Key "CompanyOrt"

    # [16.1.6.2 - If user did not specify any mailSuffix, fallback from [MailEndungen]/Domain1]
    Write-DebugMessage "Invoke-Onboarding: Checking mail suffix from userData"
    # Falls der Benutzer einen MailSuffix eingegeben hat, verwende diesen – sonst den INI-Standardwert
    if (-not [string]::IsNullOrWhiteSpace($userData.MailSuffix)) {
        $mailSuffix = $userData.MailSuffix.Trim()
        Write-DebugMessage "userData.MailSuffix= $($userData.MailSuffix)"
    }
    else {
        Write-DebugMessage "No mail suffix specified, fallback to [MailEndungen].Domain1"
        if (-not $Config.Contains("MailEndungen") -and -not $Config["MailEndungen"]) {
            Throw "Error: Section 'MailEndungen' is not defined in the INI!"
        }
        if (-not $Config["MailEndungen"].Contains("Domain1")) {
            Throw "Error: Key 'Domain1' is missing in section 'MailEndungen'!"
        }
        $mailSuffix = $Config["MailEndungen"]["Domain1"]
        Write-DebugMessage "userData.MailSuffix= $($userData.MailSuffix)"
    }
    $otherAttributes["mail"] = "$($userData.EmailAddress)$mailSuffix"

    # [16.1.6.3 - Country]
    $countryKey = "CompanyCountry$suffix"
    if (-not $companyData.Contains($countryKey) -or -not $companyData[$countryKey]) {
        Throw "Error: '$countryKey' is missing in the INI!"
    }
    $Country = $companyData[$countryKey]
    Write-DebugMessage "Country = $Country"

    # [16.1.6.4 - Company Display]
    $companyNameKey = "CompanyNameFirma$suffix"
    if (-not $companyData.Contains($companyNameKey) -or -not $companyData[$companyNameKey]) {
        Throw "Error: '$companyNameKey' is missing in the INI!"
    }
    $companyDisplay = $companyData[$companyNameKey]
    Write-DebugMessage "Company Display = $companyDisplay"

    # [16.1.7 - Integrate ADUserDefaults from INI (like mustChange, disabled, etc.)]
    Write-DebugMessage "Invoke-Onboarding: Checking [ADUserDefaults]"
    $adDefaults = $null
    if ($Config.Contains("ADUserDefaults")) {
        $adDefaults = $Config.ADUserDefaults
    }
    else {
        $adDefaults = @{}
    }

    # [16.1.7.1 - OU from [General].DefaultOU]
    if (-not $Config.General.Contains("DefaultOU")) {
        Throw "Error: 'DefaultOU' is missing in [General]!"
    }
    $defaultOU = $Config.General["DefaultOU"]

    # [16.1.8 - Build AD user parameters]
    Write-DebugMessage "Invoke-Onboarding: Building userParams for AD"
    $userParams = [ordered]@{
        # Basisinformationen
        Name                = $userData.DisplayName
        DisplayName         = $userData.DisplayName
        GivenName           = $userData.FirstName
        Surname             = $userData.LastName
        SamAccountName      = $SamAccountName
        UserPrincipalName   = $UPN
        AccountPassword     = $SecurePW

        # Konto aktivieren/ deaktivieren
        Enabled             = (-not $userData.AccountDisabled)

        # AD User Defaults
        ChangePasswordAtLogon = if ($adDefaults.Contains("MustChangePasswordAtLogon")) {
                                    $adDefaults["MustChangePasswordAtLogon"] -eq "True"
                                } else {
                                    $userData.MustChangePassword
                                }
        PasswordNeverExpires  = if ($adDefaults.Contains("PasswordNeverExpires")) {
                                    $adDefaults["PasswordNeverExpires"] -eq "True"
                                } else {
                                    $userData.PasswordNeverExpires
                                }

        # OU (verwende den vom Dropdown ausgewählten Wert – achte darauf, dass dieser nicht leer ist)
        Path                = $selectedOU

        # Adressdaten
        City                = $City
        StreetAddress       = $Street
        Country             = $Country
        postalCode          = $Zip

        # Firmenname/Anzeige
        Company             = $companyDisplay
    }

    # [16.1.8.1 - Additional defaults from [ADUserDefaults] for home/profile/logonscript]
    if ($adDefaults.Contains("HomeDirectory")) {
        $userParams["HomeDirectory"] = $adDefaults["HomeDirectory"] -replace "%username%", $SamAccountName
    }
    if ($adDefaults.Contains("ProfilePath")) {
        $userParams["ProfilePath"] = $adDefaults["ProfilePath"] -replace "%username%", $SamAccountName
    }
    if ($adDefaults.Contains("LogonScript")) {
        $userParams["ScriptPath"] = $adDefaults["LogonScript"]
    }

    # [16.1.8.2 - Collect additional attributes from the form]
    Write-DebugMessage "Invoke-Onboarding: Building otherAttributes"
    $otherAttributes = @{}

    # [16.1.8.3 - Bestimme den Mailsuffix: Falls vorhanden, verwende den vom Formular, ansonsten aus der INI aus [MailEndungen]/Domain1]
    if (-not [string]::IsNullOrWhiteSpace($userData.MailSuffix)) {
        $mailSuffix = $userData.MailSuffix.Trim()
    } else {
        Write-DebugMessage "No mail suffix specified, fallback to [MailEndungen].Domain1"
        if (-not ($Config.Contains("MailEndungen") -and $Config["MailEndungen"])) {
            Throw "Error: Section 'MailEndungen' is not defined in the INI!"
        }
        if (-not $Config["MailEndungen"].Contains("Domain1")) {
            Throw "Error: Key 'Domain1' is missing in section 'MailEndungen'!"
        }
        $mailSuffix = $Config["MailEndungen"]["Domain1"]
    }

    # [16.1.8.4 - Stelle sicher, dass EmailAddress einen gültigen (nicht-null) Wert hat]
    $emailAddress = $userData.EmailAddress
    if ($emailAddress -eq $null) { 
        $emailAddress = "" 
    }

    # [16.1.8.5 - Setze zunächst das "mail"-Attribut aus EmailAddress + mailSuffix]
    $otherAttributes["mail"] = "$emailAddress$mailSuffix"

    # [16.1.8.6 - Falls EmailAddress vorhanden ist, aber kein "@" enthält, wird ein alternativer Suffix genutzt]
    if (-not [string]::IsNullOrWhiteSpace($userData.EmailAddress)) {
        if ($userData.EmailAddress -notmatch "@") {
            $otherAttributes["mail"] = $userData.EmailAddress + $mailSuffix
        }
        else {
            $otherAttributes["mail"] = $userData.EmailAddress
        }
    } else {
        # [16.1.8.7 - Falls Email leer ist]
        $otherAttributes["mail"] = "placeholder$mailSuffix"
    }

    if ($userData.Description)    { $otherAttributes["description"] = $userData.Description }
    if ($userData.OfficeRoom)     { $otherAttributes["physicalDeliveryOfficeName"] = $userData.OfficeRoom }
    if ($userData.PhoneNumber)    { $otherAttributes["telephoneNumber"] = $userData.PhoneNumber }
    if ($userData.MobileNumber)   { $otherAttributes["mobile"] = $userData.MobileNumber }
    if ($userData.Position)       { $otherAttributes["title"] = $userData.Position }
    if ($userData.DepartmentField){ $otherAttributes["department"] = $userData.DepartmentField }

    # [16.1.8.8 - If there's a CompanyTelefon in the config (company phone), add as ipPhone]
    if ($companyData.Contains("CompanyTelefon") -and -not [string]::IsNullOrWhiteSpace($companyData["CompanyTelefon"])) {
        $otherAttributes["ipPhone"] = $companyData["CompanyTelefon"].Trim()
    }

    if ($otherAttributes.Count -gt 0) {
        $userParams["OtherAttributes"] = $otherAttributes
    }

    # [16.1.8.9 - If there's an Ablaufdatum (termination), set accountExpirationDate]
    if (-not [string]::IsNullOrWhiteSpace($userData.Ablaufdatum)) {
        try {
            $expirationDate = [DateTime]::Parse($userData.Ablaufdatum)
            $userParams["AccountExpirationDate"] = $expirationDate
        }
        catch {
            Write-Warning "Invalid termination date format: $($_.Exception.Message)"
        }
    }

    # [16.1.9 - Create or update AD user]
    Write-DebugMessage "Invoke-Onboarding: Checking if user already exists in AD"
    $existingUser = $null
    try {
        $existingUser = Get-ADUser -Filter { SamAccountName -eq $SamAccountName } -ErrorAction SilentlyContinue
    }
    catch {
        $existingUser = $null
    }

    if (-not $existingUser) {
        Write-DebugMessage "Invoke-Onboarding: Creating new user: $($userData.DisplayName)"
        Invoke-ADCommand -Command { New-ADUser @userParams -ErrorAction Stop } -ErrorContext "Erstellen des AD-Benutzers für $($userData.DisplayName)"
        
        # [16.1.9.1 - Smartcard Logon-Einstellung]
        if ($userData.SmartcardLogonRequired) {
            Write-DebugMessage "Invoke-Onboarding: Setting SmartcardLogonRequired for $($userData.DisplayName)"
            try {
                Set-ADUser -Identity $SamAccountName -SmartcardLogonRequired $true -ErrorAction Stop
                Write-DebugMessage "Smartcard logon requirement set successfully."
            }
            catch {
                Write-Warning "Error setting Smartcard logon"
            }
        }
        
        # [16.1.9.2 - "CannotChangePassword" – Hinweis, dass diese Funktionalität noch nicht implementiert ist]
        if ($userData.CannotChangePassword) {
            Write-DebugMessage "Invoke-Onboarding: (Note) 'CannotChangePassword' via ACL not yet implemented."
        }
    }
    else {
        Write-DebugMessage "Invoke-Onboarding: User '$SamAccountName' already exists - updating attributes"
        try {
            Set-ADUser -Identity $existingUser.DistinguishedName `
                -GivenName $userData.FirstName `
                -Surname $userData.LastName `
                -City $City `
                -StreetAddress $Street `
                -Country $Country `
                -Enabled (-not $userData.AccountDisabled) -ErrorAction SilentlyContinue
    
            Set-ADUser -Identity $existingUser.DistinguishedName `
                -ChangePasswordAtLogon:$userData.MustChangePassword `
                -PasswordNeverExpires:$userData.PasswordNeverExpires
    
            if ($otherAttributes.Count -gt 0) {
                Set-ADUser -Identity $existingUser.DistinguishedName -Replace $otherAttributes -ErrorAction SilentlyContinue
            }
        }
        catch {
            Write-Warning "Error updating user: $($_.Exception.Message)"
        }
    }    

    # [16.1.9.3 - Setting AD account password (Reset)]
    Write-DebugMessage "Invoke-Onboarding: Setting AD account password (Reset)"
    try {
        Set-ADAccountPassword -Identity $SamAccountName -Reset -NewPassword $SecurePW -ErrorAction SilentlyContinue
    }
    catch {
        Write-Warning "Error setting password: $($_.Exception.Message)"
    }

    # [16.1.10 - Handling AD group assignments]
    Write-DebugMessage "Invoke-Onboarding: Handling AD group assignments"
    if ($userData.External) {
        Write-DebugMessage "External user: skipping default AD group assignment."
    }
    else {
        # [16.1.10.1 - Check any ADGroupsSelected from the GUI]
        Write-DebugMessage "Processing ADGroupsSelected: $($userData.ADGroupsSelected | ConvertTo-Json -Compress)"

        # Ensure $adGroupsToProcess is always a string array
        $adGroupsToProcess = @()
        if ($null -ne $userData.ADGroupsSelected) {
            if ($userData.ADGroupsSelected -is [string]) {
                if ($userData.ADGroupsSelected -eq "KEINE") {
                    Write-DebugMessage "No AD groups selected ('KEINE')"
                    $adGroupsToProcess = @()
                } else {
                    # Split the semicolon-separated string into an array
                    $adGroupsToProcess = $userData.ADGroupsSelected.Split(';', [StringSplitOptions]::RemoveEmptyEntries)
                    Write-DebugMessage "Split ADGroupsSelected string into array of $($adGroupsToProcess.Count) items"
                }
            }
            elseif ($userData.ADGroupsSelected -is [array]) {
                $adGroupsToProcess = $userData.ADGroupsSelected
                Write-DebugMessage "Using array of $($adGroupsToProcess.Count) items directly from ADGroupsSelected"
            }
            else {
                # Force conversion to string then split (handles other object types)
                $groupString = $userData.ADGroupsSelected.ToString()
                $adGroupsToProcess = $groupString.Split(';', [StringSplitOptions]::RemoveEmptyEntries)
                Write-DebugMessage "Converted object to string and split into array of $($adGroupsToProcess.Count) items"
            }
        }

        foreach ($grpKey in $adGroupsToProcess) {
            Write-DebugMessage "Processing group key: '$grpKey'"
            # First try direct lookup using the value as a key
            if ($grpKey -and $Config.ADGroups.Contains($grpKey)) {
                $groupName = $Config.ADGroups[$grpKey]
                Write-DebugMessage "Found group name '$groupName' for key '$grpKey'"
                if ($groupName) {
                    try {
                        Add-ADGroupMember -Identity $groupName -Members $SamAccountName -ErrorAction Stop
                        Write-DebugMessage "Added user to group '$groupName'."
                    }
                    catch {
                        Write-Warning "Error adding group '$groupName': $($_.Exception.Message)"
                    }
                }
            } else {
                # Try to find if this is a display name/label by searching through values
                $found = $false
                foreach ($adGroupKey in $Config.ADGroups.Keys) {
                    # Look for keys that match ADGroup pattern and have matching label or value
                    if ($adGroupKey -match '^ADGroup\d+$') {
                        $labelKey = "${adGroupKey}_Label"
                        # Check if display label matches
                        if (($Config.ADGroups.Contains($labelKey) -and $Config.ADGroups[$labelKey] -eq $grpKey) -or
                            $Config.ADGroups[$adGroupKey] -eq $grpKey) {
                            
                            $found = $true
                            $groupName = $Config.ADGroups[$adGroupKey]
                            Write-DebugMessage "Found group name '$groupName' via label/value lookup"
                            
                            try {
                                Add-ADGroupMember -Identity $groupName -Members $SamAccountName -ErrorAction Stop
                                Write-DebugMessage "Added user to group '$groupName'."
                            }
                            catch {
                                Write-Warning "Error adding group '$groupName': $($_.Exception.Message)"
                            }
                            break
                        }
                    }
                }
                
                if (-not $found) {
                    # Try direct use as an AD group name
                    try {
                        Get-ADGroup -Identity $grpKey -ErrorAction Stop | Out-Null
                        Add-ADGroupMember -Identity $grpKey -Members $SamAccountName -ErrorAction Stop
                        Write-DebugMessage "Added user to group '$grpKey' using direct group name."
                    }
                    catch {
                        Write-DebugMessage "Group key '$grpKey' not found in Config.ADGroups and is not a valid AD group name"
                    }
                }
            }
        }
    }

    # [16.1.10.2 - Team Leader (TL) or Department Head (AL)]
    # Fix for TL group assignment
    if ($userData.TL -eq $true) {
        Write-DebugMessage "Invoke-Onboarding: TL is true - adding user to selected TLGroup: '$($userData.TLGroup)'"
        # Check both direct key lookup and if it exists as a key in the TLGroups section
        if (-not [string]::IsNullOrWhiteSpace($userData.TLGroup)) {
            $tlGroupName = ""
            if ($Config.TLGroups.Contains($userData.TLGroup)) {
                $tlGroupName = $Config.TLGroups[$userData.TLGroup]
                Write-DebugMessage "Found TL group '$tlGroupName' via key lookup"
            } elseif ($Config.TLGroups.Values -contains $userData.TLGroup) {
                $tlGroupName = $userData.TLGroup
                Write-DebugMessage "Using direct TL group name: $tlGroupName"
            }
            
            if (-not [string]::IsNullOrWhiteSpace($tlGroupName)) {
                try {
                    Add-ADGroupMember -Identity $tlGroupName -Members $SamAccountName -ErrorAction Stop
                    Write-DebugMessage "Added user to Team-Leader group '$tlGroupName'."
                }
                catch {
                    Write-Warning "Error adding TL group '$tlGroupName': $($_.Exception.Message)"
                }
            } else {
                Write-DebugMessage "Could not resolve TL group name from '$($userData.TLGroup)'"
            }
        }
        else {
            Write-DebugMessage "No TLGroup value specified though TL is true"
        }
    }

    # Fix for AL group assignment
    if ($userData.AL -eq $true) {
        Write-DebugMessage "Invoke-Onboarding: AL is true - adding user to ALGroup"
        if ($Config.Contains("ALGroup") -and $Config.ALGroup.Contains("Group")) {
            $alGroupName = $Config.ALGroup["Group"]
            Write-DebugMessage "Found ALGroup: $alGroupName"
            try {
                Add-ADGroupMember -Identity $alGroupName -Members $SamAccountName -ErrorAction Stop
                Write-DebugMessage "Added user to Abteilungsleiter group '$alGroupName'."
            }
            catch {
                Write-Warning "Error adding AL group '$alGroupName': $($_.Exception.Message)"
            }
        }
        else {
            Write-DebugMessage "ALGroup configuration is missing or invalid."
        }
    }            

    # [16.1.10.3 - Add user to default groups from [UserCreationDefaults]]
    if ($Config.Contains("UserCreationDefaults") -and $Config.UserCreationDefaults.Contains("InitialGroupMembership")) {
        Write-DebugMessage "Invoke-Onboarding: InitialGroupMembership"
        $defaultGroups = $Config.UserCreationDefaults["InitialGroupMembership"].Split(";") | ForEach-Object { $_.Trim() }
        foreach ($g in $defaultGroups) {
            if ($g) {
                try {
                    # Check if user is already a member before adding
                    $isMember = $false
                    try {
                        $groupMembers = Get-ADGroupMember -Identity $g -ErrorAction Stop
                        $isMember = ($groupMembers | Where-Object { $_.SamAccountName -eq $SamAccountName }) -ne $null
                    } catch {
                        # Group might not exist or other error
                        Write-DebugMessage "Error checking membership in group '$g': $($_.Exception.Message)"
                    }
                    
                    if (-not $isMember) {
                        Add-ADGroupMember -Identity $g -Members $SamAccountName -ErrorAction Stop
                        Write-DebugMessage "Added user to default group '$g'."
                    } else {
                        Write-DebugMessage "User is already a member of group '$g' - skipping."
                    }
                }
                catch {
                    Write-Warning "Error adding default group '$g': $($_.Exception.Message)"
                }
            }
        }
    }

    # [16.1.10.4 - License]
    # Fix for License group assignment 
    if (-not [string]::IsNullOrWhiteSpace($userData.License) -and $userData.License -notmatch "^(?i:keine|none)$") {
        $licenseKey = "MS365_" + $userData.License.Trim().ToUpper()
        Write-DebugMessage "Processing license group with key: $licenseKey"
        
        if ($Config.Contains("LicensesGroups")) {
            # Try direct key lookup
            if ($Config.LicensesGroups.Contains($licenseKey)) {
                $licenseGroup = $Config.LicensesGroups[$licenseKey]
                Write-DebugMessage "Found license group '$licenseGroup' via key lookup"
            } 
            # Try to see if the selected value itself is a group name
            elseif ($Config.LicensesGroups.Values -contains $userData.License) {
                $licenseGroup = $userData.License
                Write-DebugMessage "Using direct license value as group name: $licenseGroup"
            }
            # Try without MS365_ prefix if key not found
            elseif ($Config.LicensesGroups.Contains($userData.License.Trim().ToUpper())) {
                $licenseGroup = $Config.LicensesGroups[$userData.License.Trim().ToUpper()]
                Write-DebugMessage "Found license group '$licenseGroup' via simplified key lookup"
            }
            else {
                Write-Warning "License key '$licenseKey' not found in [LicensesGroups]."
                $licenseGroup = $null
            }
            
            if ($licenseGroup) {
                try {
                    Add-ADGroupMember -Identity $licenseGroup -Members $SamAccountName -ErrorAction Stop
                    Write-DebugMessage "Added user to license group '$licenseGroup'."
                }
                catch {
                    Write-Warning "Error adding license group '$licenseGroup': $($_.Exception.Message)"
                }
            }
        }
        else {
            Write-Warning "[LicensesGroups] section not found in configuration."
        }
    }
   
    # [16.1.10.5 - MS365 AD Sync]
    if ($Config.Contains("ActivateUserMS365ADSync") -and $Config.ActivateUserMS365ADSync.Contains("ADSync") -and $Config.ActivateUserMS365ADSync["ADSync"] -eq "1") {
        $syncGroup = $Config.ActivateUserMS365ADSync["ADSyncADGroup"]
        if ($syncGroup) {
            try {
                Add-ADGroupMember -Identity $syncGroup -Members $SamAccountName -ErrorAction Stop
                Write-DebugMessage "Added user to AD-Sync group '$syncGroup'."
            }
            catch {
                Write-Warning "Error adding AD-Sync group '$syncGroup': $($_.Exception.Message)"
            }
        }
    }

    # [16.1.11 - Logging]
    Write-DebugMessage "Invoke-Onboarding: Logging to file"
    try {
        Write-LogMessage -Message ("ONBOARDING PERFORMED BY: " + ([System.Security.Principal.WindowsIdentity]::GetCurrent().Name) +
            ", SamAccountName: $SamAccountName, Display Name: '$($userData.DisplayName)', UPN: '$UPN', Location: '$($userData.Location)', Company: '$companyDisplay', License: '$($userData.License)', Password: '$mainPW', External: $($userData.External)") -Level Info
        Write-DebugMessage "Log written"
    }
    catch {
        Write-Warning "Error logging: $($_.Exception.Message)"
    }

    # [16.1.11.1 - Initialize variables for reports]
    $reportBranding = $global:Config["Branding-Report"]
    $reportPath = $reportBranding["ReportPath"]
    if (-not (Test-Path $reportPath)) {
        New-Item -ItemType Directory -Path $reportPath -Force | Out-Null
    }

    $finalReportFooter = $reportBranding["ReportFooter"]

    # [16.1.12 - Create Reports (HTML + TXT)]
    Write-DebugMessage "Invoke-Onboarding: Creating HTML and/or TXT reports"
    try {
        if (-not (Test-Path $reportPath)) {
            New-Item -ItemType Directory -Path $reportPath -Force | Out-Null
        }

        # [16.1.12.1 - Load the HTML template]
        $htmlFile = Join-Path $reportPath "$SamAccountName.html"
        $htmlTemplatePath = $reportBranding["TemplatePath"]
        $htmlTemplate = ""
        if (-not $htmlTemplatePath -or -not (Test-Path $htmlTemplatePath)) {
            Write-DebugMessage "No valid HTML template found. Using fallback text."
            $htmlTemplate = "INI or HTML Template not provided!"
        }
        else {
            Write-DebugMessage "Reading HTML template from: $htmlTemplatePath"
            $htmlTemplate = Get-Content -Path $htmlTemplatePath -Raw
        }

        # [16.1.12.2 - Possibly build Websites HTML block if you want to show links in HTML]
        $websitesHTML = ""
        if ($Config.Contains("Websites")) {
            $websitesHTML += "<ul>`r`n"
            foreach ($key in $Config.Websites.Keys) {
                if ($key -match '^EmployeeLink\d+$') {
                    $line = $Config.Websites[$key]
                    # The format is "Label | URL | Description", or "Label;URL;Description"
                    if ($line -match '\|') {
                        $parts = $line -split '\|'
                    }
                    else {
                        $parts = $line -split ';'
                    }
                    if ($parts.Count -eq 3) {
                        $lbl = $parts[0].Trim()
                        $url = $parts[1].Trim()
                        $desc= $parts[2].Trim()
                        $websitesHTML += "<li><a href='$url' target='_blank'>$lbl</a> - $desc</li>`r`n"
                    }
                }
            }
            $websitesHTML += "</ul>`r`n"
        }

        # [16.1.12.3 - Replace placeholders in the HTML]
        $logoTag = ""
        if ($userData.CustomLogo -and (Test-Path $userData.CustomLogo)) {
            $logoTag = "<img src='$($userData.CustomLogo)' alt='Logo' />"
        }

        # Fix inline if expressions by assigning to variables first
        $description = if ($null -ne $userData.Description) { $userData.Description } else { "" }
        $officeRoom = if ($null -ne $userData.OfficeRoom) { $userData.OfficeRoom } else { "" }
        $phoneNumber = if ($null -ne $userData.PhoneNumber) { $userData.PhoneNumber } else { "" }
        $mobileNumber = if ($null -ne $userData.MobileNumber) { $userData.MobileNumber } else { "" }
        $position = if ($null -ne $userData.Position) { $userData.Position } else { "" }
        $departmentField = if ($null -ne $userData.DepartmentField) { $userData.DepartmentField } else { "" }
        $ablaufdatum = if ($null -ne $userData.Ablaufdatum) { $userData.Ablaufdatum } else { "" }

        $htmlContent = $htmlTemplate `
            -replace "{{ReportTitle}}",    ([string]$reportTitle) `
            -replace "{{Admin}}",          $env:USERNAME `
            -replace "{{ReportDate}}",     (Get-Date -Format "yyyy-MM-dd") `
            -replace "{{Vorname}}",        $userData.FirstName `
            -replace "{{Nachname}}",       $userData.LastName `
            -replace "{{DisplayName}}",    $userData.DisplayName `
            -replace "{{Description}}",    $description `
            -replace "{{Buero}}",          $officeRoom `
            -replace "{{Rufnummer}}",      $phoneNumber `
            -replace "{{Mobil}}",          $mobileNumber `
            -replace "{{Position}}",       $position `
            -replace "{{Abteilung}}",      $departmentField `
            -replace "{{Ablaufdatum}}",    $ablaufdatum `
            -replace "{{LoginName}}",      $SamAccountName `
            -replace "{{Passwort}}",       $mainPW `
            -replace "{{WebsitesHTML}}",   $websitesHTML `
            -replace "{{ReportFooter}}",   $finalReportFooter `
            -replace "{{LogoTag}}",        $logoTag `
            -replace "{{CustomPWLabel1}}", $pwLabel1 `
            -replace "{{CustomPWLabel2}}", $pwLabel2 `
            -replace "{{CustomPWLabel3}}", $pwLabel3 `
            -replace "{{CustomPWLabel4}}", $pwLabel4 `
            -replace "{{CustomPWLabel5}}", $pwLabel5 `
            -replace "{{CustomPW1}}",      $CustomPW1 `
            -replace "{{CustomPW2}}",      $CustomPW2 `
            -replace "{{CustomPW3}}",      $CustomPW3 `
            -replace "{{CustomPW4}}",      $CustomPW4 `
            -replace "{{CustomPW5}}",      $CustomPW5

        Set-Content -Path $htmlFile -Value $htmlContent -Encoding UTF8
        Write-DebugMessage "HTML report created: $htmlFile"

        # [16.1.12.4 - Create the TXT if userData.OutputTXT or if the INI says to]
        $shouldCreateTXT = $userData.OutputTXT -or ($Config.General["UserOnboardingCreateTXT"] -eq '1')
        if ($shouldCreateTXT) {
            # [16.1.12.4.1 - If you have a separate txtTemplate]
            $txtTemplatePath = Join-Path $PSScriptRoot "txtTemplate.txt"  # oder aus der Config, falls gewünscht
            $txtFile = Join-Path $reportPath "$SamAccountName.txt"

            [string]$txtContent = ""
            if (Test-Path $txtTemplatePath) {
                Write-DebugMessage "Reading txtTemplate from: $txtTemplatePath"
                $txtContent = Get-Content -Path $txtTemplatePath -Raw

                # [16.1.12.4.2 - Replace placeholders in the template]
                # Fix placeholder replacements to avoid parsing issues
                $defaultOU = $Config.General["DefaultOU"]
                $txtContent = $txtContent -replace '\$Config\.General\["DefaultOU"\]', $defaultOU
                $txtContent = $txtContent -replace "\`$SamAccountName", $SamAccountName
                $txtContent = $txtContent -replace "\`$UserPW", $mainPW
                $txtContent = $txtContent -replace "\`$userData\.AccountDisabled", $userData.AccountDisabled
                # Safer pattern replacement approach
                $txtContent = [regex]::Replace($txtContent, "\$\(.*?\)", "")
                
                # ...remaining code...
            }
            else {
                # [16.1.12.4.4 - If no external txtTemplate file, build it inline]
                Write-DebugMessage "No external txtTemplate found - using inline approach"
                $txtFile = Join-Path $reportPath "$SamAccountName.txt"

                $txtContent = @"
Onboarding Report for New Employees
Created by: $($env:USERNAME)
Date: $(Get-Date -Format 'yyyy-MM-dd')

Active Directory OU Path:
----------------------------------------------------
$($Config.General["DefaultOU"])
====================================================

Credentials
====================================================
User Name:     $SamAccountName
Password:      $mainPW
----------------------------------------------------
Enabled:       $([bool](-not $userData.AccountDisabled))

User Details:
====================================================
First Name:    $($userData.FirstName)
Last  Name:    $($userData.LastName)
Display Name:  $($userData.DisplayName)
External:      $($userData.External)
Office:        $($userData.OfficeRoom)
Phone:         $($userData.PhoneNumber)
Mobile:        $($userData.MobileNumber)
Position:      $($userData.Position)
Department:    $($userData.DepartmentField)
====================================================

Additional Passwords:
----------------------------------------------------
$pwLabel1 : $CustomPW1
$pwLabel2 : $CustomPW2
$pwLabel3 : $CustomPW3
$pwLabel4 : $CustomPW4
$pwLabel5 : $CustomPW5

Useful Links:
----------------------------------------------------
"@

                # [16.1.12.4.5 - Add the websites from [Websites]]
                if ($Config.Contains("Websites")) {
                    foreach ($key in $Config.Websites.Keys) {
                        if ($key -match '^EmployeeLink\d+$') {
                            $line = $Config.Websites[$key]
                            $parts = $line -split '[\|\;]'
                            if ($parts.Count -eq 3) {
                                $title = $parts[0].Trim()
                                $url   = $parts[1].Trim()
                                $desc  = $parts[2].Trim()
                                $txtContent += "$title : $url  ($desc)`r`n"
                            }
                        }
                    }
                }
                $txtContent += "`r`n$finalReportFooter`r`n"
            }

            Write-DebugMessage "Writing final TXT to: $txtFile"
            Out-File -FilePath $txtFile -InputObject $txtContent -Encoding UTF8
            Write-DebugMessage "TXT report created: $txtFile"
        }
    }
    catch {
        Write-Warning "Error creating reports: $($_.Exception.Message)"
    }

    Write-DebugMessage "Invoke-Onboarding: Returning final result object"
    # [16.1.13 - Return a small object with final SamAccountName, UPN, and PW if needed]
    Write-DebugMessage "Onboarding process completed successfully."
    return [ordered]@{
        SamAccountName = $SamAccountName
        UPN            = $UPN
        Password       = $mainPW
    }
} 
Write-Debug ("userData: " + ( $userData | Out-String ))

#region [Region 17 | DROPDOWN UTILITY]
# [Function to set values and options for dropdown controls]
# [Funktion zum Setzen von Werten und Optionen für Dropdown-Steuerelemente]
Write-DebugMessage "Set-DropDownValues"
function Set-DropDownValues {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        $DropDownControl,    # Das Dropdown/ComboBox-Steuerelement (WinForms oder WPF)

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$DataType,   # Erforderlicher Typ (z. B. "OU", "Location", "License", etc.)

        [Parameter(Mandatory = $true)]
        [ValidateNotNull()]
        $ConfigData        # Konfigurationsdatenquelle (z. B. Hashtable)
    )
    
    try {
        Write-Verbose "Befülle Dropdown für Typ '$DataType'..."
        Write-DebugMessage "Set-DropDownValues: Entferne vorhandene Bindungen und Items"
        
        # Entferne bestehende Bindungen und Items (unabhängig vom verwendeten Framework)
        if ($DropDownControl -is [System.Windows.Forms.ComboBox]) {
            $DropDownControl.DataSource = $null
            $DropDownControl.Items.Clear()
        }
        elseif ($DropDownControl -is [System.Windows.Controls.ItemsControl]) {
            $DropDownControl.ItemsSource = $null
            $DropDownControl.Items.Clear()
        }
        else {
            try { $DropDownControl.Items.Clear() } catch { }
        }
        Write-Verbose "Bestehende Items entfernt."
        
        Write-DebugMessage "Set-DropDownValues: Items und Standardwert aus Config abrufen"
        $items = @()
        $defaultValue = $null
        switch ($DataType.ToLower()) {
            "ou" {
                $items = $ConfigData.OUList
                $defaultValue = $ConfigData.DefaultOU
            }
            "location" {
                $items = $ConfigData.LocationList
                $defaultValue = $ConfigData.DefaultLocation
            }
            "license" {
                $items = $ConfigData.LicenseList
                $defaultValue = $ConfigData.DefaultLicense
            }
            "mailsuffix" {
                $items = $ConfigData.MailSuffixList
                $defaultValue = $ConfigData.DefaultMailSuffix
            }
            "displaynametemplate" {
                $items = $ConfigData.DisplayNameTemplates
                $defaultValue = $ConfigData.DefaultDisplayNameTemplate
            }
            "tlgroups" {
                $items = $ConfigData.TLGroupsList
                $defaultValue = $ConfigData.DefaultTLGroup
            }
            default {
                Write-Warning "Unbekannter Dropdown-Typ '$DataType'. Abbruch."
                return
            }
        }
    
        Write-DebugMessage "Set-DropDownValues: Sicherstellen, dass Items eine Sammlung ist"
        if ($items) {
            if ($items -is [string] -or -not ($items -is [System.Collections.IEnumerable])) {
                $items = @($items)
            }
        }
        else {
            $items = @()
        }
        Write-Verbose "$($items.Count) Einträge für '$DataType' aus der Konfiguration abgerufen."
    
        Write-DebugMessage "Set-DropDownValues: Setze neue Datenbindung/Items"
        if ($DropDownControl -is [System.Windows.Forms.ComboBox]) {
            $DropDownControl.DataSource = $items
        }
        else {
            $DropDownControl.ItemsSource = $items
        }
        Write-Verbose "Datenbindung für '$DataType'-Dropdown gesetzt."
    
        Write-DebugMessage "Set-DropDownValues: Setze Standardwert, falls vorhanden"
        if ($defaultValue) {
            $DropDownControl.SelectedItem = $defaultValue
            Write-Verbose "Standardwert für '$DataType' auf '$defaultValue' gesetzt."
        }
        elseif ($items.Count -gt 0) {
            try { $DropDownControl.SelectedIndex = 0 } catch { }
        }
    }
    catch {
        Write-Error "Fehler beim Befüllen des Dropdowns '$DataType': $($_.Exception.Message)"
    }
}
#endregion

Write-DebugMessage "Process-ADGroups"

#region [Region 18 | AD GROUP CREATION]
# [Function for handling AD group creation (placeholder)]
# [Funktion zur Erstellung von AD-Gruppen (Platzhalter)]
Write-DebugMessage "Processing AD groups."
function Process-ADGroups {
    param(
        [Parameter(Mandatory=$true)] $GroupData,
        [Parameter(Mandatory=$true)] [hashtable]$Config
    )
    Write-Log "Process-ADGroups: Gruppenerstellung wird durchgeführt..." "DEBUG"
    # Hier kann der Code zur Gruppenerstellung ergänzt und optimiert werden.
    return $true
}
#endregion

Write-DebugMessage "Get-ADData"

#region [Region 19 | AD DATA RETRIEVAL]
# [Function to retrieve and process data from Active Directory]
# [Funktion zum Abrufen und Verarbeiten von Daten aus Active Directory]
Write-DebugMessage "Retrieving AD data."
function Get-ADData {
    param(
        [Parameter(Mandatory=$false)]
        $Window
    )
    
    try {
        # Alle OUs abrufen und in ein sortiertes Array konvertieren
        $allOUs = Get-ADOrganizationalUnit -Filter * -Properties Name, DistinguishedName |
            ForEach-Object {
                [PSCustomObject]@{
                    Name = $_.Name
                    DN   = $_.DistinguishedName
                }
            } | Sort-Object Name

        # Debug-Ausgabe für OUs, wenn DebugMode=1
        if ($global:Config.General.DebugMode -eq "1") {
            Write-Log " Alle OUs wurden geladen:" "DEBUG"
            $allOUs | Format-Table | Out-String | Write-Log
        }
        Write-DebugMessage "Get-ADData: Alle Benutzer abrufen"
        # Alle Benutzer abrufen (hier werden die ersten 200 Benutzer sortiert nach DisplayName angezeigt)
        $allUsers = Get-ADUser -Filter * -Properties DisplayName, SamAccountName |
            ForEach-Object {
                [PSCustomObject]@{
                    DisplayName    = if ($_.DisplayName) { $_.DisplayName } else { $_.SamAccountName }
                    SamAccountName = $_.SamAccountName
                }
            } | Sort-Object DisplayName | Select-Object

        # Debug-Ausgabe für Benutzer, wenn DebugMode=1
        if ($global:Config.General.DebugMode -eq "1") {
            Write-Log " Benutzerliste wurde geladen:" "DEBUG"
            $allUsers | Format-Table | Out-String | Write-Log
        }
        Write-DebugMessage "Get-ADData: comboBoxOU und listBoxUsers"
        
        # Finde die ComboBox und ListBox im Window, falls übergeben
        if ($Window) {
            $comboBoxOU = $Window.FindName("cmbOU") 
            $listBoxUsers = $Window.FindName("lstUsers")
            
            if ($comboBoxOU) {
                $comboBoxOU.ItemsSource = $null
                $comboBoxOU.Items.Clear()
                $comboBoxOU.ItemsSource = $allOUs
                $comboBoxOU.DisplayMemberPath = "Name"
                $comboBoxOU.SelectedValuePath = "DN"
                if ($allOUs.Count -gt 0) {
                    $comboBoxOU.SelectedIndex = 0
                }
            }

            if ($listBoxUsers) {
                $listBoxUsers.ItemsSource = $null
                $listBoxUsers.Items.Clear()
                $listBoxUsers.ItemsSource = $allUsers
                $listBoxUsers.DisplayMemberPath = "DisplayName"
                $listBoxUsers.SelectedValuePath = "SamAccountName"
            }
        }
        
        # Rückgabe der Daten für weitere Verwendung
        return @{
            OUs = $allOUs
            Users = $allUsers
        }
    }
    catch {
        Write-DebugMessage "Error in Get-ADData: $($_.Exception.Message)"
        [System.Windows.MessageBox]::Show("AD-Verbindungsfehler: $($_.Exception.Message)")
        return $null
    }
}
#endregion

Write-DebugMessage "Open-INIEditor"

#region [Region 20 | INI EDITOR]
# [Function to open the INI configuration editor]
# [Funktion zum Öffnen des INI-Konfigurationseditors]
Write-DebugMessage "Opening INI editor."
function Open-INIEditor {
    try {
        if ($global:Config.General.DebugMode -eq "1") {
            Write-Log " INI Editor wird gestartet." "DEBUG"
        }
        Write-DebugMessage "INI Editor wurde gestartet."
        Write-LogMessage -Message "INI Editor wurde gestartet." -Level "Info"
        [System.Windows.MessageBox]::Show("INI Editor (Settings) – Funktionalität hier einbinden.", "Settings")
    }
    catch {
        Throw "Fehler beim Öffnen des INI Editors: $_"
    }
}
#endregion
#endregion

Write-DebugMessage "Defining GUI implementation."

#region [Region 21 | GUI IMPLEMENTATION]
# [Handles the entire GUI implementation and event handlers]
# [Behandelt die gesamte GUI-Implementierung und Event-Handler]

Write-DebugMessage "GUI Sammle ausgewählte Gruppen aus dem AD-Gruppen-Panel"

#region [Region 21.1 | AD GROUPS SELECTION]
# [Collects selected groups from the AD groups panel]
# [Sammelt ausgewählte Gruppen aus dem AD-Gruppen-Panel]
Write-DebugMessage "Collecting selected AD groups from panel."
function Get-SelectedADGroups {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [System.Windows.Controls.ItemsControl]$Panel,
        
        [Parameter(Mandatory=$false)]
        [string]$Separator = ";"
    )
    
    # Use System.Collections.ArrayList to avoid pipeline output when adding items
    [System.Collections.ArrayList]$selectedGroups = @()
    
    if ($null -eq $Panel) {
        Write-DebugMessage "Error: Panel is null"
        return "KEINE"
    }
    
    if ($Panel.Items.Count -eq 0) {
        Write-DebugMessage "Panel contains no items to select from"
        return "KEINE"
    }
    
    Write-DebugMessage "Processing $($Panel.Items.Count) items in AD groups panel"
    
    for ($i = 0; $i -lt $Panel.Items.Count; $i++) {
        $item = $Panel.Items[$i]
        if ($null -eq $item) { continue }
        
        try {
            $container = $Panel.ItemContainerGenerator.ContainerFromItem($item)
            if ($null -eq $container) { 
                Write-DebugMessage "Container for item '$item' is null - skipping"
                continue 
            }
            
            # Find CheckBox without relying on a specific name
            $checkBox = $null
            # This finds the first CheckBox control in the visual tree of the container
            $checkBox = [System.Windows.Media.VisualTreeHelper]::GetChild($container, 0)
            
            if ($checkBox -is [System.Windows.Controls.CheckBox]) {
                if ($checkBox.IsChecked) {
                    Write-DebugMessage "Selected group: $item"
                    [void]$selectedGroups.Add($item)
                }
            } else {
                # The direct child might be a ContentPresenter, try to find CheckBox inside it
                $contentPresenter = $checkBox
                if ($contentPresenter -is [System.Windows.Controls.ContentPresenter]) {
                    # Wait for ContentPresenter to generate its content
                    if (!$contentPresenter.Content) {
                        $contentPresenter.UpdateLayout()
                    }
                    
                    # Try to find CheckBox in ContentPresenter
                    $checkBox = FindVisualChild -parent $contentPresenter -childType ([System.Windows.Controls.CheckBox])
                    
                    if ($checkBox -and $checkBox.IsChecked) {
                        Write-DebugMessage "Selected group through ContentPresenter: $item"
                        [void]$selectedGroups.Add($item)
                    }
                }
            }
        } catch {
            Write-DebugMessage "Error processing AD group item '$item': $($_.Exception.Message)"
            # Continue with the next item
            continue
        }
    }
    
    Write-DebugMessage "Total selected groups: $($selectedGroups.Count)"
    
    # Return appropriate result
    if ($selectedGroups.Count -eq 0) {
        return "KEINE"
    } else {
        return ($selectedGroups -join $Separator)
    }
}

# Helper function to find a specific type of child in the visual tree
function FindVisualChild {
    param(
        [Parameter(Mandatory=$true)]
        [System.Windows.DependencyObject]$parent,
        
        [Parameter(Mandatory=$true)]
        [type]$childType
    )
    
    for ($i = 0; $i -lt [System.Windows.Media.VisualTreeHelper]::GetChildrenCount($parent); $i++) {
        $child = [System.Windows.Media.VisualTreeHelper]::GetChild($parent, $i)
        
        if ($child -and $child -is $childType) {
            return $child
        } else {
            $result = FindVisualChild -parent $child -childType $childType
            if ($result) {
                return $result
            }
        }
    }
    
    return $null
}
#endregion

Write-DebugMessage "Setting up logo upload button."

#region [Region 21.2 | LOGO UPLOAD BUTTON]
# [Sets up the logo upload button and its functionality]
# [Richtet den Logo-Upload-Button und dessen Funktionalität ein]
$btnUploadLogo = $window.FindName("btnUploadLogo")
if ($btnUploadLogo) {
    Register-GUIEvent -Control $btnUploadLogo -EventAction {
        if (-not $global:Config -or -not $global:Config.Contains("Branding-Report")) {
            Throw "Die Sektion [Branding-Report] fehlt in der INI-Datei."
        }
        $brandingConfig = $global:Config["Branding-Report"]
        Set-Logo -brandingConfig $brandingConfig
    } -ErrorMessagePrefix "Logo Upload Fehler"
} else {
    Write-DebugMessage "HINWEIS: Button 'btnUploadLogo' wurde nicht in der XAML-Oberfläche gefunden."
}

#GUI: START "easyONBOARDING BUTTON"
Write-DebugMessage "GUI: TAB easyONBOARDING geladen"

#region [Region 21.3 | ONBOARDING TAB]
# [Implements the main onboarding tab functionality]
# [Implementiert die Hauptfunktionalität des Onboarding-Tabs]
Write-DebugMessage "Onboarding tab loaded."
$onboardingTab = $window.FindName("Tab_Onboarding")
# [21.3.1 - Validates the tab exists in the XAML interface]
# [Überprüft, ob der Tab in der XAML-Oberfläche existiert]
if (-not $onboardingTab) {
    Write-DebugMessage "ERROR: Onboarding-Tab fehlt!"
    Write-Error "Das Onboarding-Tab (x:Name='Tab_Onboarding') wurde in der XAML nicht gefunden."
    exit
}
#endregion

Write-DebugMessage "Setting up onboarding button handler."

#region [Region 21.4 | ONBOARDING BUTTON HANDLER]
# [Implements the start onboarding button click handler]
# [Implementiert den Klick-Handler für den Onboarding-Start-Button]
$btnStartOnboarding = $onboardingTab.FindName("btnOnboard")
# [21.4.1 - Sets up the primary function to execute when user initiates onboarding]
# [Richtet die Hauptfunktion ein, die ausgeführt wird, wenn der Benutzer das Onboarding startet]
if (-not $btnStartOnboarding) {
    Write-DebugMessage "ERROR: btnOnboard wurde NICHT gefunden!"
    exit
}

Write-DebugMessage "Event-Handler für btnOnboard wird registriert"
$btnStartOnboarding.Add_Click({
    Write-DebugMessage "Onboarding-Button wurde geklickt!"

    try {
        # Finde das AD-Gruppen-Panel für die Selektion
        $icADGroups = $onboardingTab.FindName("icADGroups")
        $ADGroupsSelected = if ($icADGroups) { 
            # Get selected groups as a string
            Get-SelectedADGroups -Panel $icADGroups 
        } else { 
            "KEINE" 
        }
        
        Write-DebugMessage "Lade GUI-Elemente..."
        Write-DebugMessage "Selected AD groups: $ADGroupsSelected"
        $txtFirstName = $onboardingTab.FindName("txtFirstName")
        $txtLastName = $onboardingTab.FindName("txtLastName")
        $txtDisplayName = $onboardingTab.FindName("txtDisplayName")
        $txtEmail = $onboardingTab.FindName("txtEmail")
        $txtOffice = $onboardingTab.FindName("txtOffice")
        $txtPhone = $onboardingTab.FindName("txtPhone")
        $txtMobile = $onboardingTab.FindName("txtMobile")
        $txtPosition = $onboardingTab.FindName("txtPosition")
        $txtDepartment = $onboardingTab.FindName("txtDepartment")
        $txtTermination = $onboardingTab.FindName("txtTermination")
        $chkExternal = $onboardingTab.FindName("chkExternal")
        $chkTL = $onboardingTab.FindName("chkTL")
        $chkAL = $onboardingTab.FindName("chkAL")
        $txtDescription = $onboardingTab.FindName("txtDescription")
        $chkAccountDisabled = $onboardingTab.FindName("chkAccountDisabled")
        $chkMustChangePassword = $onboardingTab.FindName("chkPWChangeLogon")
        $chkPasswordNeverExpires = $onboardingTab.FindName("chkPWNeverExpires")
        $comboBoxDisplayTemplate = $onboardingTab.FindName("cmbDisplayTemplate")
        $comboBoxMailSuffix = $onboardingTab.FindName("cmbMailSuffix")
        $comboBoxLicense = $onboardingTab.FindName("cmbLicense")
        $comboBoxOU = $onboardingTab.FindName("cmbOU")
        $comboBoxTLGroup = $onboardingTab.FindName("cmbTLGroup")

        Write-DebugMessage "GUI-Elemente erfolgreich geladen."

        # **Validierung der Pflichtfelder**
        if (-not $txtFirstName -or -not $txtFirstName.Text) {
            Write-DebugMessage "FEHLER: Vorname fehlt!"
            [System.Windows.MessageBox]::Show("Fehler: Vorname darf nicht leer sein!", "Fehler", 'OK', 'Error')
            return
        }
        if (-not $txtLastName -or -not $txtLastName.Text) {
            Write-DebugMessage "FEHLER: Nachname fehlt!"
            [System.Windows.MessageBox]::Show("Fehler: Nachname darf nicht leer sein!", "Fehler", 'OK', 'Error')
            return
        }

        # **Werte setzen, NULL-Werte vermeiden**
        $firstName = $txtFirstName.Text.Trim()
        $lastName = $txtLastName.Text.Trim()
        $displayName = if ($txtDisplayName -and $txtDisplayName.Text) { $txtDisplayName.Text.Trim() } else { "$firstName $lastName" }
        $email = if ($txtEmail -and $txtEmail.Text) { $txtEmail.Text.Trim() } else { "" }
        $office = if ($txtOffice -and $txtOffice.Text) { $txtOffice.Text.Trim() } else { "" }
        $phone = if ($txtPhone -and $txtPhone.Text) { $txtPhone.Text.Trim() } else { "" }
        $mobile = if ($txtMobile -and $txtMobile.Text) { $txtMobile.Text.Trim() } else { "" }
        $position = if ($txtPosition -and $txtPosition.Text) { $txtPosition.Text.Trim() } else { "" }
        $department = if ($txtDepartment -and $txtDepartment.Text) { $txtDepartment.Text.Trim() } else { "" }
        $terminationDate = if ($txtTermination -and $txtTermination.Text) { $txtTermination.Text.Trim() } else { "" }
        $external = if ($chkExternal) { $chkExternal.IsChecked } else { $false }
        $tl = if ($chkTL) { $chkTL.IsChecked } else { $false }
        $al = if ($chkAL) { $chkAL.IsChecked } else { $false }

        # **MailSuffix, Lizenz und OU**
        $mailSuffix = if ($comboBoxMailSuffix -and $comboBoxMailSuffix.SelectedValue) { 
            $comboBoxMailSuffix.SelectedValue.ToString() 
        } elseif ($global:Config.Contains("Company") -and $global:Config.Company.Contains("CompanyMailDomain")) { 
            $global:Config.Company["CompanyMailDomain"] 
        } else { "" }
        
        $selectedLicense = if ($comboBoxLicense -and $comboBoxLicense.SelectedValue) { 
            $comboBoxLicense.SelectedValue.ToString() 
        } else { "Standard" }
        
        $selectedOU = if ($comboBoxOU -and $comboBoxOU.SelectedValue) { 
            $comboBoxOU.SelectedValue.ToString() 
        } elseif ($global:Config.Contains("General") -and $global:Config.General.Contains("DefaultOU")) { 
            $global:Config.General["DefaultOU"] 
        } else { "" }

        $selectedTLGroup = if ($comboBoxTLGroup -and $comboBoxTLGroup.SelectedValue) {
            $comboBoxTLGroup.SelectedValue.ToString()
        } else { "" }

        # **Globales Objekt erstellen**
        $global:userData = [PSCustomObject]@{
            OU               = $selectedOU
            FirstName        = $firstName
            LastName         = $lastName
            DisplayName      = $displayName
            Description      = if ($txtDescription -and $txtDescription.Text) { $txtDescription.Text.Trim() } else { "" }
            EmailAddress     = $email
            OfficeRoom       = $office
            PhoneNumber      = $phone
            MobileNumber     = $mobile
            Position         = $position
            DepartmentField  = $department
            Ablaufdatum      = $terminationDate
            ADGroupsSelected = $ADGroupsSelected
            External         = $external
            TL               = $tl
            AL               = $al
            TLGroup          = $selectedTLGroup
            MailSuffix       = $mailSuffix
            License          = $selectedLicense
            UPNFormat        = if ($comboBoxDisplayTemplate -and $comboBoxDisplayTemplate.SelectedValue) { 
                $comboBoxDisplayTemplate.SelectedValue.ToString().Trim() 
            } elseif ($global:Config.Contains("DisplayNameUPNTemplates") -and $global:Config.DisplayNameUPNTemplates.Contains("DefaultUserPrincipalNameFormat")) { 
                $global:Config.DisplayNameUPNTemplates["DefaultUserPrincipalNameFormat"] 
            } else { "FIRSTNAME.LASTNAME" }
            AccountDisabled  = if ($chkAccountDisabled) { $chkAccountDisabled.IsChecked } else { $false }
            MustChangePassword   = if ($chkMustChangePassword) { $chkMustChangePassword.IsChecked } else { $false }
            PasswordNeverExpires = if ($chkPasswordNeverExpires) { $chkPasswordNeverExpires.IsChecked } else { $false }
            CompanyName      = if ($global:Config.Contains("Company") -and $global:Config.Company.Contains("CompanyNameFirma")) { 
                $global:Config.Company["CompanyNameFirma"] 
            } else { "" }
            CompanyStrasse   = if ($global:Config.Contains("Company") -and $global:Config.Company.Contains("CompanyStrasse")) { 
                $global:Config.Company["CompanyStrasse"] 
            } else { "" }
            CompanyPLZ       = if ($global:Config.Contains("Company") -and $global:Config.Company.Contains("CompanyPLZ")) { 
                $global:Config.Company["CompanyPLZ"] 
            } else { "" }
            CompanyOrt       = if ($global:Config.Contains("Company") -and $global:Config.Company.Contains("CompanyOrt")) { 
                $global:Config.Company["CompanyOrt"] 
            } else { "" }
            CompanyDomain    = if ($global:Config.Contains("Company") -and $global:Config.Company.Contains("CompanyActiveDirectoryDomain")) { 
                $global:Config.Company["CompanyActiveDirectoryDomain"] 
            } else { "" }
            CompanyTelefon   = if ($global:Config.Contains("Company") -and $global:Config.Company.Contains("CompanyTelefon")) { 
                $global:Config.Company["CompanyTelefon"] 
            } else { "" }
            CompanyCountry   = if ($global:Config.Contains("Company") -and $global:Config.Company.Contains("CompanyCountry")) { 
                $global:Config.Company["CompanyCountry"] 
            } else { "" }
            # Fallback für PasswordMode, falls später genutzt
            PasswordMode     = 1 # Standard: Passwort generieren
        }        

        Write-DebugMessage "userData-Objekt erstellt -> `n$($global:userData | Out-String)"
        Write-DebugMessage "Starte Invoke-Onboarding..."

        try {
            $global:result = Invoke-Onboarding -userData $global:userData -Config $global:Config
            if (-not $global:result) {
                throw "Invoke-Onboarding hat kein Ergebnis zurückgegeben!"
            }

            Write-DebugMessage "Invoke-Onboarding abgeschlossen."
            [System.Windows.MessageBox]::Show("Onboarding erfolgreich durchgeführt.`nSamAccountName: $($global:result.SamAccountName)`nUPN: $($global:result.UPN)", "Erfolg")
        } catch {
            $errorMsg = "Fehler in Invoke-Onboarding: $($_.Exception.Message)"
            if ($_.Exception.InnerException) {
                $errorMsg += "`nInnerException: $($_.Exception.InnerException.Message)"
            }
            Write-DebugMessage "ERROR: $errorMsg"
            [System.Windows.MessageBox]::Show($errorMsg, "Fehler", 'OK', 'Error')
        }

    } catch {
        Write-DebugMessage "ERROR: Unbehandelter Fehler: $($_.Exception.Message)"
        [System.Windows.MessageBox]::Show("Fehler: $($_.Exception.Message)", "Fehler", 'OK', 'Error')
    }
})
#endregion

# END "easyONBOARDING BUTTON"
Write-DebugMessage "TAB easyONBOARDING - BUTTON - CREATE PDF"

#region [Region 21.5 | PDF CREATION BUTTON]
# [Implements the PDF creation button functionality]
# [Implementiert die Funktionalität des PDF-Erstellungs-Buttons]
Write-DebugMessage "Setting up PDF creation button."
$btnPDF = $window.FindName("btnPDF")
if ($btnPDF) {
    Write-DebugMessage "TAB easyONBOARDING - BUTTON ausgewählt - CREATE PDF"
    $btnPDF.Add_Click({
        try {
            if (-not $global:result -or -not $global:result.SamAccountName) {
                [System.Windows.MessageBox]::Show("Bitte zuerst das Onboarding durchführen, bevor das PDF erstellt werden kann.", "Fehler", 'OK', 'Error')
                Write-Log "$global:result ist NULL oder leer" "DEBUG"
                return
            }
            
            Write-Log "$global:result erfolgreich geladen: $($global:result | Out-String)" "DEBUG"
            
            $SamAccountName = $global:result.SamAccountName

            if (-not ($global:Config.Keys -contains "Branding-Report")) {
                [System.Windows.MessageBox]::Show("Fehler: Die Sektion [Branding-Report] fehlt in der INI-Datei.", "Fehler", 'OK', 'Error')
                return
            }

            $reportBranding = $global:Config["Branding-Report"]
            if (-not $reportBranding -or -not $reportBranding.Contains("ReportPath") -or [string]::IsNullOrWhiteSpace($reportBranding["ReportPath"])) {
                [System.Windows.MessageBox]::Show("ReportPath fehlt oder ist leer in [Branding-Report]. PDF kann nicht erstellt werden.", "Fehler", 'OK', 'Error')
                return
            }         

            $htmlReportPath = Join-Path $reportBranding["ReportPath"] "$SamAccountName.html"
            $pdfReportPath  = Join-Path $reportBranding["ReportPath"] "$SamAccountName.pdf"

            $wkhtmltopdfPath = $reportBranding["wkhtmltopdfPath"]
            if (-not (Test-Path $wkhtmltopdfPath)) {
                [System.Windows.MessageBox]::Show("wkhtmltopdf.exe wurde nicht gefunden! Bitte prüfen: $wkhtmltopdfPath", "Fehler", 'OK', 'Error')
                return
            }          

            if (-not (Test-Path $htmlReportPath)) {
                [System.Windows.MessageBox]::Show("HTML-Report nicht gefunden: $htmlReportPath`nDas PDF kann nicht erstellt werden.", "Fehler", 'OK', 'Error')
                return
            }

            $pdfScript = Join-Path $PSScriptRoot "easyOnboarding_PDFCreator.ps1"
            if (-not (Test-Path $pdfScript)) {
                [System.Windows.MessageBox]::Show("PDF-Skript 'easyOnboarding_PDFCreator.ps1' nicht gefunden!", "Fehler", 'OK', 'Error')
                return
            }

            $arguments = @(
                "-NoProfile",
                "-ExecutionPolicy", "Bypass",
                "-File", "`"$pdfScript`"",
                "-htmlFile", "`"$htmlReportPath`"",
                "-pdfFile", "`"$pdfReportPath`"",
                "-wkhtmltopdfPath", "`"$wkhtmltopdfPath`""
            )
            Start-Process -FilePath "powershell.exe" -ArgumentList $arguments -NoNewWindow -Wait

            [System.Windows.MessageBox]::Show("PDF erfolgreich erstellt: $pdfReportPath", "PDF-Erstellung", 'OK', 'Information')

        } catch {
            [System.Windows.MessageBox]::Show("Fehler beim Erstellen des PDFs: $($_.Exception.Message)", "Fehler", 'OK', 'Error')
        }
    })
}
#endregion

Write-DebugMessage "TAB easyONBOARDING - BUTTON Info"

#region [Region 21.6 | INFO BUTTON]
# [Implements the info button functionality]
# [Implementiert die Funktionalität des Info-Buttons]
Write-DebugMessage "Setting up info button."
$btnInfo = $window.FindName("btnInfo")
if ($btnInfo) {
    Write-DebugMessage "TAB easyONBOARDING - BUTTON ausgewählt: Info"
    $btnInfo.Add_Click({
        $infoFilePath = Join-Path $PSScriptRoot "info_easyONBOARDING.txt"

        if (Test-Path $infoFilePath) {
            Start-Process -FilePath $infoFilePath
        } else {
            [System.Windows.MessageBox]::Show("Die Datei info_easyONBOARDING.txt wurde nicht gefunden!", "Fehler", 'OK', 'Error')
        }
    })
}
#endregion

Write-DebugMessage "TAB easyONBOARDING - BUTTON Close"

#region [Region 21.7 | CLOSE BUTTON]
# [Implements the close button functionality]
# [Implementiert die Funktionalität des Schließen-Buttons]
Write-DebugMessage "Setting up close button."
$btnClose = $window.FindName("btnClose")
if ($btnClose) {
    Write-DebugMessage "TAB easyONBOARDING - BUTTON ausgewählt: Close"
    $btnClose.Add_Click({
        $window.Close()
    })
}
#endregion

Write-DebugMessage "GUI: TAB easyADUpdate geladen"

#region [Region 21.8 | AD UPDATE TAB]
# [Implements the AD Update tab functionality]
# [Implementiert die Funktionalität des AD-Update-Tabs]
Write-DebugMessage "AD Update tab loaded."

# Find the AD Update tab
$adUpdateTab = $window.FindName("Tab_ADUpdate")
if ($adUpdateTab) {

    # Retrieve relevant controls with corrected variable names to match XAML
    # Verbesserte Null-Wert-Prüfung mit Try-Catch für den Fall, dass XAML-Elemente fehlen
    try {
        $btnSearchADUpdate             = $adUpdateTab.FindName("btnSearchADUpdate")
        $txtSearchADUpdate             = $adUpdateTab.FindName("txtSearchADUpdate")
        $lstUsersADUpdate              = $adUpdateTab.FindName("lstUsersADUpdate")
        $btnADUserUpdate               = $adUpdateTab.FindName("btnADUserUpdate")
        $btnADUserCancel               = $adUpdateTab.FindName("btnADUserCancel")
        $txtFirstNameUpdate            = $adUpdateTab.FindName("txtFirstNameUpdate")
        $txtLastNameUpdate             = $adUpdateTab.FindName("txtLastNameUpdate")
        $txtDisplayNameUpdate          = $adUpdateTab.FindName("txtDisplayNameUpdate")
        $txtEmailUpdate                = $adUpdateTab.FindName("txtEmailUpdate")
        $txtDepartmentUpdate           = $adUpdateTab.FindName("txtDepartmentUpdate")
        $txtPhoneUpdate                = $adUpdateTab.FindName("txtPhoneUpdate")
        $txtMobileUpdate               = $adUpdateTab.FindName("txtMobileUpdate")
        $txtOfficeUpdate               = $adUpdateTab.FindName("txtOfficeUpdate")
        $chkAccountEnabledUpdate       = $adUpdateTab.FindName("chkAccountEnabledUpdate")
        $chkPasswordNeverExpiresUpdate = $adUpdateTab.FindName("chkPasswordNeverExpiresUpdate") 
        $chkMustChangePasswordUpdate   = $adUpdateTab.FindName("chkMustChangePasswordUpdate")
        $lstGroupsUpdate               = $adUpdateTab.FindName("lstGroupsUpdate")

        # Extended fields
        $txtManagerUpdate              = $adUpdateTab.FindName("txtManagerUpdate")
        $txtJobTitleUpdate             = $adUpdateTab.FindName("txtJobTitleUpdate")
        $txtLocationUpdate             = $adUpdateTab.FindName("txtLocationUpdate")
        $txtEmployeeIDUpdate           = $adUpdateTab.FindName("txtEmployeeIDUpdate")

        # Group membership controls
        $btnAddGroupUpdate             = $adUpdateTab.FindName("btnAddGroupUpdate")
        $btnRemoveGroupUpdate          = $adUpdateTab.FindName("btnRemoveGroupUpdate")
    } 
    catch {
        Write-DebugMessage "Error loading AD Update tab controls: $($_.Exception.Message)"
        return
    }

    # Search button functionality
    if ($null -ne $btnSearchADUpdate -and $null -ne $txtSearchADUpdate -and $null -ne $lstUsersADUpdate) {
        Write-DebugMessage "Registering event for btnSearchADUpdate"
        Register-GUIEvent -Control $btnSearchADUpdate -EventAction {
            $searchTerm = $txtSearchADUpdate.Text.Trim()
            if ([string]::IsNullOrWhiteSpace($searchTerm)) {
                [System.Windows.MessageBox]::Show("Bitte geben Sie ein Suchwort ein.", "Info", "OK", "Information")
                return
            }
            Write-DebugMessage "Searching AD users for '$searchTerm'..."
            try {
                # Improved search to work with different attribute types
                $allMatches = Get-ADUser -Filter "(DisplayName -like '*$searchTerm*') -or (SamAccountName -like '*$searchTerm*') -or (mail -like '*$searchTerm*')" -Properties DisplayName, SamAccountName, EmailAddress, Department, Title, Office, OfficePhone, Mobile, Manager -ErrorAction Stop

                # Optimiertes Array-Handling mit @() Operator
                if ($null -eq $allMatches) {
                    Write-DebugMessage "No users found for search term: $searchTerm"
                    [System.Windows.MessageBox]::Show("Keine Benutzer gefunden für: $searchTerm", "Keine Ergebnisse", "OK", "Information")
                    return
                }

                # Immer als Array behandeln, egal ob ein oder mehrere Ergebnisse
                $results = @($allMatches) | Select-Object DisplayName, SamAccountName, @{Name='Email';Expression={$_.EmailAddress}}
                
                # Itemsource sichern leeren und neu setzen
                $lstUsersADUpdate.ItemsSource = $null
                $lstUsersADUpdate.Items.Clear()
                $lstUsersADUpdate.ItemsSource = $results
                
                Write-DebugMessage "Found $($results.Count) users matching '$searchTerm'"
                
                if ($results.Count -eq 0) {
                    [System.Windows.MessageBox]::Show("Keine Benutzer gefunden für: $searchTerm", "Keine Ergebnisse", "OK", "Information")
                }
            }
            catch {
                Write-DebugMessage "Fehler bei der Suche: $($_.Exception.Message)"
                [System.Windows.MessageBox]::Show("Fehler bei der Suche: $($_.Exception.Message)", "Fehler", "OK", "Error")
            }
        } -ErrorMessagePrefix "Fehler bei der Benutzersuche"
    }
    else {
        Write-DebugMessage "Required search controls missing in XAML"
    }

    # Populate fields when selection changes
    if ($null -ne $lstUsersADUpdate) {
        $lstUsersADUpdate.Add_SelectionChanged({
            if ($null -ne $lstUsersADUpdate.SelectedItem) {
                $selectedUser = $lstUsersADUpdate.SelectedItem
                $samAccountName = $selectedUser.SamAccountName
                
                if ([string]::IsNullOrWhiteSpace($samAccountName)) {
                    Write-DebugMessage "Selected user has no SamAccountName"
                    return
                }
                
                try {
                    # Get full user details from AD
                    $adUser = Get-ADUser -Identity $samAccountName -Properties DisplayName, GivenName, Surname, EmailAddress, 
                        Department, Title, OfficePhone, Mobile, physicalDeliveryOfficeName, Enabled, 
                        PasswordNeverExpires, Manager, employeeID, MemberOf -ErrorAction Stop
                    
                    if ($null -ne $adUser) {
                        # Optimiertes Attributsetzen mit Null-Prüfung
                        if ($null -ne $txtFirstNameUpdate) { $txtFirstNameUpdate.Text = $adUser.GivenName ?? "" }
                        if ($null -ne $txtLastNameUpdate) { $txtLastNameUpdate.Text = $adUser.Surname ?? "" }
                        if ($null -ne $txtDisplayNameUpdate) { $txtDisplayNameUpdate.Text = $adUser.DisplayName ?? "" }
                        if ($null -ne $txtEmailUpdate) { $txtEmailUpdate.Text = $adUser.EmailAddress ?? "" }
                        if ($null -ne $txtDepartmentUpdate) { $txtDepartmentUpdate.Text = $adUser.Department ?? "" }
                        if ($null -ne $txtPhoneUpdate) { $txtPhoneUpdate.Text = $adUser.OfficePhone ?? "" }
                        if ($null -ne $txtMobileUpdate) { $txtMobileUpdate.Text = $adUser.Mobile ?? "" }
                        if ($null -ne $txtOfficeUpdate) { $txtOfficeUpdate.Text = $adUser.physicalDeliveryOfficeName ?? "" }
                        if ($null -ne $txtJobTitleUpdate) { $txtJobTitleUpdate.Text = $adUser.Title ?? "" }
                        if ($null -ne $txtLocationUpdate) { $txtLocationUpdate.Text = $adUser.physicalDeliveryOfficeName ?? "" }
                        if ($null -ne $txtEmployeeIDUpdate) { $txtEmployeeIDUpdate.Text = $adUser.employeeID ?? "" }
                        
                        # Set checkboxes mit Null-Prüfung
                        if ($null -ne $chkAccountEnabledUpdate) { $chkAccountEnabledUpdate.IsChecked = $adUser.Enabled }
                        if ($null -ne $chkPasswordNeverExpiresUpdate) { $chkPasswordNeverExpiresUpdate.IsChecked = $adUser.PasswordNeverExpires }
                        if ($null -ne $chkMustChangePasswordUpdate) { $chkMustChangePasswordUpdate.IsChecked = $false } # Default, da nicht direkt aus AD auslesbar
                        
                        # Manager handling mit verbesserter Fehlerbehandlung
                        if ($null -ne $txtManagerUpdate) {
                            if (-not [string]::IsNullOrEmpty($adUser.Manager)) {
                                try {
                                    $manager = Get-ADUser -Identity $adUser.Manager -Properties DisplayName -ErrorAction Stop
                                    $txtManagerUpdate.Text = if ($manager.DisplayName) { $manager.DisplayName } else { $manager.SamAccountName }
                                } catch {
                                    $txtManagerUpdate.Text = ""
                                    Write-DebugMessage "Error retrieving manager: $($_.Exception.Message)"
                                }
                            } else {
                                $txtManagerUpdate.Text = ""
                            }
                        }
                        
                        # Populate group membership list mit verbesserten Null- und Array-Prüfungen
                        if ($null -ne $lstGroupsUpdate) {
                            $lstGroupsUpdate.Items.Clear()
                            
                            # MemberOf kann null, ein einzelnes Objekt oder ein Array sein
                            $memberOfGroups = @($adUser.MemberOf)
                            
                            if ($memberOfGroups.Count -gt 0) {
                                foreach ($group in $memberOfGroups) {
                                    try {
                                        if (-not [string]::IsNullOrEmpty($group)) {
                                            $groupObj = Get-ADGroup -Identity $group -Properties Name -ErrorAction Stop
                                            if ($null -ne $groupObj -and -not [string]::IsNullOrEmpty($groupObj.Name)) {
                                                [void]$lstGroupsUpdate.Items.Add($groupObj.Name)
                                            }
                                        }
                                    } catch {
                                        Write-DebugMessage "Error retrieving group details"
                                    }
                                }
                            }
                        }
                        
                        Write-DebugMessage "Fields populated for user: $samAccountName"
                    } else {
                        Write-DebugMessage "AD User object is null for user: $samAccountName"
                        [System.Windows.MessageBox]::Show("Der ausgewählte Benutzer konnte nicht gefunden werden.", "Fehler", "OK", "Error")
                    }
                } catch {
                    Write-DebugMessage "Error loading user details: $($_.Exception.Message)"
                    [System.Windows.MessageBox]::Show("Fehler beim Laden der Benutzerdetails: $($_.Exception.Message)", "Fehler", "OK", "Error")
                }
            }
        })
    }

    # Update User button
    if ($null -ne $btnADUserUpdate) {
        Write-DebugMessage "Registering event handler for btnADUserUpdate"
        Register-GUIEvent -Control $btnADUserUpdate -EventAction {
            try {
                if ($null -eq $lstUsersADUpdate -or $null -eq $lstUsersADUpdate.SelectedItem) {
                    [System.Windows.MessageBox]::Show("Bitte einen Benutzer aus der Liste auswählen.", "Fehlende Auswahl", "OK", "Warning")
                    return
                }
                
                $userToUpdate = $lstUsersADUpdate.SelectedItem.SamAccountName
                if ([string]::IsNullOrWhiteSpace($userToUpdate)) {
                    [System.Windows.MessageBox]::Show("Ungültiger Benutzername ausgewählt.", "Fehler", "OK", "Warning")
                    return
                }
                
                Write-DebugMessage "Updating user: $userToUpdate"

                # Validation
                if ($null -eq $txtDisplayNameUpdate -or [string]::IsNullOrWhiteSpace($txtDisplayNameUpdate.Text)) {
                    [System.Windows.MessageBox]::Show("Der Anzeigename darf nicht leer sein.", "Validierung", "OK", "Warning")
                    return
                }

                # Build parameter hashtable with null checks
                $paramUpdate = @{
                    Identity = $userToUpdate
                }
                
                # Add parameters only if controls exist and have values
                if ($null -ne $txtDisplayNameUpdate -and -not [string]::IsNullOrWhiteSpace($txtDisplayNameUpdate.Text)) {
                    $paramUpdate["DisplayName"] = $txtDisplayNameUpdate.Text
                }
                
                if ($null -ne $chkAccountEnabledUpdate) {
                    $paramUpdate["Enabled"] = $chkAccountEnabledUpdate.IsChecked
                }
                
                if ($null -ne $chkPasswordNeverExpiresUpdate) {
                    $paramUpdate["PasswordNeverExpires"] = $chkPasswordNeverExpiresUpdate.IsChecked
                }

                # Only add parameters if control exists and has values
                if ($null -ne $txtFirstNameUpdate -and -not [string]::IsNullOrWhiteSpace($txtFirstNameUpdate.Text)) {
                    $paramUpdate["GivenName"] = $txtFirstNameUpdate.Text
                }
                
                if ($null -ne $txtLastNameUpdate -and -not [string]::IsNullOrWhiteSpace($txtLastNameUpdate.Text)) {
                    $paramUpdate["Surname"] = $txtLastNameUpdate.Text
                }
                
                if ($null -ne $txtEmailUpdate -and -not [string]::IsNullOrWhiteSpace($txtEmailUpdate.Text)) {
                    $paramUpdate["EmailAddress"] = $txtEmailUpdate.Text
                }
                
                if ($null -ne $txtDepartmentUpdate -and -not [string]::IsNullOrWhiteSpace($txtDepartmentUpdate.Text)) {
                    $paramUpdate["Department"] = $txtDepartmentUpdate.Text
                }
                
                if ($null -ne $txtPhoneUpdate -and -not [string]::IsNullOrWhiteSpace($txtPhoneUpdate.Text)) {
                    $paramUpdate["OfficePhone"] = $txtPhoneUpdate.Text
                }
                
                if ($null -ne $txtMobileUpdate -and -not [string]::IsNullOrWhiteSpace($txtMobileUpdate.Text)) {
                    $paramUpdate["MobilePhone"] = $txtMobileUpdate.Text
                }
                
                if ($null -ne $txtJobTitleUpdate -and -not [string]::IsNullOrWhiteSpace($txtJobTitleUpdate.Text)) {
                    $paramUpdate["Title"] = $txtJobTitleUpdate.Text
                }

                # Resolve manager if given with improved error handling
                if ($null -ne $txtManagerUpdate -and -not [string]::IsNullOrWhiteSpace($txtManagerUpdate.Text)) {
                    try {
                        $managerObj = Get-ADUser -Filter {(SamAccountName -eq $txtManagerUpdate.Text) -or (DisplayName -eq $txtManagerUpdate.Text)} -ErrorAction Stop
                        
                        if ($null -ne $managerObj) {
                            # Handle if multiple users match (take first one)
                            if ($managerObj -is [array]) {
                                $managerObj = $managerObj[0]
                                Write-DebugMessage "Multiple managers found, using first match: $($managerObj.SamAccountName)"
                            }
                            
                            $paramUpdate["Manager"] = $managerObj.DistinguishedName
                        } else {
                            Write-DebugMessage "Manager object is null."
                            [System.Windows.MessageBox]::Show("Der angegebene Manager konnte nicht gefunden werden: $($txtManagerUpdate.Text)", "Warnung", "OK", "Warning")
                        }
                    }
                    catch {
                        Write-DebugMessage "Manager lookup error: $($_.Exception.Message)"
                        [System.Windows.MessageBox]::Show("Der angegebene Manager konnte nicht gefunden werden: $($txtManagerUpdate.Text)", "Warnung", "OK", "Warning")
                        # We'll continue without setting the manager
                    }
                }

                # Update the user if we have valid parameters
                if ($paramUpdate.Count -gt 1) { # More than just Identity
                    Set-ADUser @paramUpdate -ErrorAction Stop
                    Write-DebugMessage "Updated user attributes: $($paramUpdate.Keys -join ', ')"
                }

                # Handle attributes that require Replace (Office/Location and Employee ID)
                $replaceHash = @{}
                
                if ($null -ne $txtOfficeUpdate -and -not [string]::IsNullOrWhiteSpace($txtOfficeUpdate.Text)) {
                    $replaceHash["physicalDeliveryOfficeName"] = $txtOfficeUpdate.Text
                }
                
                if ($null -ne $txtLocationUpdate -and -not [string]::IsNullOrWhiteSpace($txtLocationUpdate.Text)) {
                    $replaceHash["l"] = $txtLocationUpdate.Text  # 'l' attribute for location
                }
                
                if ($null -ne $txtEmployeeIDUpdate -and -not [string]::IsNullOrWhiteSpace($txtEmployeeIDUpdate.Text)) {
                    $replaceHash["employeeID"] = $txtEmployeeIDUpdate.Text
                }
                
                if ($replaceHash.Count -gt 0) {
                    Set-ADUser -Identity $userToUpdate -Replace $replaceHash -ErrorAction Stop
                    Write-DebugMessage "Updated replace attributes: $($replaceHash.Keys -join ', ')"
                }

                # Handle password change requirement
                if ($null -ne $chkMustChangePasswordUpdate -and $null -ne $chkMustChangePasswordUpdate.IsChecked) {
                    Set-ADUser -Identity $userToUpdate -ChangePasswordAtLogon $chkMustChangePasswordUpdate.IsChecked -ErrorAction Stop
                    Write-DebugMessage "Set ChangePasswordAtLogon to $($chkMustChangePasswordUpdate.IsChecked)"
                }

                [System.Windows.MessageBox]::Show("Der Benutzer '$userToUpdate' wurde erfolgreich aktualisiert.", "Erfolg", "OK", "Information")
                Write-LogMessage -Message "AD-Update für $userToUpdate durch $($env:USERNAME)" -LogLevel "INFO"
            }
            catch {
                Write-DebugMessage "AD User update error: $($_.Exception.Message)"
                [System.Windows.MessageBox]::Show("Fehler beim Aktualisieren: $($_.Exception.Message)", "Fehler", "OK", "Error")
            }
        } -ErrorMessagePrefix "Fehler beim AD-Update"
    }

    # Cancel button
    if ($null -ne $btnADUserCancel) {
        Write-DebugMessage "Registering event handler for btnADUserCancel"
        Register-GUIEvent -Control $btnADUserCancel -EventAction {
            Write-DebugMessage "Resetting update form"
            
            # Array der Textfeld-Namen für bessere Wartbarkeit
            $textFields = @("txtFirstNameUpdate","txtLastNameUpdate","txtDisplayNameUpdate","txtEmailUpdate",
                            "txtPhoneUpdate","txtMobileUpdate","txtDepartmentUpdate","txtOfficeUpdate",
                            "txtManagerUpdate","txtJobTitleUpdate","txtLocationUpdate","txtEmployeeIDUpdate")
            
            # Clear all text fields with null checking
            foreach ($name in $textFields) {
                $ctrl = $adUpdateTab.FindName($name)
                if ($null -ne $ctrl) { $ctrl.Text = "" }
            }
            
            # Reset checkboxes with null checking
            if ($null -ne $chkAccountEnabledUpdate) { $chkAccountEnabledUpdate.IsChecked = $false }
            if ($null -ne $chkPasswordNeverExpiresUpdate) { $chkPasswordNeverExpiresUpdate.IsChecked = $false }
            if ($null -ne $chkMustChangePasswordUpdate) { $chkMustChangePasswordUpdate.IsChecked = $false }
            
            # Clear selection and lists with null checking
            if ($null -ne $lstUsersADUpdate) { 
                $lstUsersADUpdate.ItemsSource = $null
                $lstUsersADUpdate.Items.Clear()
                $lstUsersADUpdate.SelectedIndex = -1 
            }
            
            if ($null -ne $lstGroupsUpdate) { $lstGroupsUpdate.Items.Clear() }
            
            [System.Windows.MessageBox]::Show("Formular zurückgesetzt.", "Zurückgesetzt", "OK", "Information")
        } -ErrorMessagePrefix "Fehler beim Zurücksetzen"
    }

    # Add Group button functionality
    if ($null -ne $btnAddGroupUpdate) {
        Register-GUIEvent -Control $btnAddGroupUpdate -EventAction {
            if ($null -eq $lstUsersADUpdate -or $null -eq $lstUsersADUpdate.SelectedItem) {
                [System.Windows.MessageBox]::Show("Bitte wählen Sie zuerst einen Benutzer aus.", "Hinweis", "OK", "Information")
                return
            }
            
            $selectedUser = $lstUsersADUpdate.SelectedItem.SamAccountName
            if ([string]::IsNullOrWhiteSpace($selectedUser)) {
                [System.Windows.MessageBox]::Show("Ungültiger Benutzername ausgewählt.", "Fehler", "OK", "Error")
                return
            }
            
            # Create a simple input dialog for the group name
            $inputDialog = New-Object System.Windows.Forms.Form
            $inputDialog.Width = 400
            $inputDialog.Height = 150
            $inputDialog.Text = "Gruppe hinzufügen"
            $inputDialog.StartPosition = "CenterScreen"
            $inputDialog.FormBorderStyle = "FixedDialog"
            $inputDialog.MaximizeBox = $false
            $inputDialog.MinimizeBox = $false
            
            $label = New-Object System.Windows.Forms.Label
            $label.Text = "Geben Sie den Namen der AD-Gruppe ein:"
            $label.Width = 350
            $label.Location = New-Object System.Drawing.Point(10, 10)
            
            $textBox = New-Object System.Windows.Forms.TextBox
            $textBox.Width = 350
            $textBox.Location = New-Object System.Drawing.Point(10, 40)
            
            $okButton = New-Object System.Windows.Forms.Button
            $okButton.Text = "OK"
            $okButton.Location = New-Object System.Drawing.Point(190, 80)
            $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
            
            $cancelButton = New-Object System.Windows.Forms.Button
            $cancelButton.Text = "Abbrechen"
            $cancelButton.Location = New-Object System.Drawing.Point(270, 80)
            $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
            
            $inputDialog.Controls.Add($label)
            $inputDialog.Controls.Add($textBox)
            $inputDialog.Controls.Add($okButton)
            $inputDialog.Controls.Add($cancelButton)
            $inputDialog.AcceptButton = $okButton
            $inputDialog.CancelButton = $cancelButton
            
            $result = $inputDialog.ShowDialog()
            
            if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
                $groupName = $textBox.Text.Trim()
                
                if (-not [string]::IsNullOrEmpty($groupName)) {
                    try {
                        # Verify the group exists with improved error handling
                        $group = Get-ADGroup -Identity $groupName -ErrorAction Stop
                        if ($null -eq $group) {
                            throw "Gruppe '$groupName' konnte nicht gefunden werden."
                        }
                        
                        # Check if the user is already a member
                        $isMember = $false
                        try {
                            $members = Get-ADGroupMember -Identity $groupName -ErrorAction Stop
                            if ($members -is [array]) {
                                $isMember = ($members | Where-Object { $_.SamAccountName -eq $selectedUser }).Count -gt 0
                            } else {
                                $isMember = $members.SamAccountName -eq $selectedUser
                            }
                        } catch {
                            Write-DebugMessage "Error checking group membership: $($_.Exception.Message)"
                        }
                        
                        if ($isMember) {
                            [System.Windows.MessageBox]::Show("Der Benutzer ist bereits Mitglied der Gruppe '$groupName'.", "Information", "OK", "Information")
                            return
                        }
                        
                        # Add user to group
                        Add-ADGroupMember -Identity $groupName -Members $selectedUser -ErrorAction Stop
                        
                        # Refresh the group list if it exists
                        if ($null -ne $lstGroupsUpdate) {
                            [void]$lstGroupsUpdate.Items.Add($group.Name)
                        }
                        
                        [System.Windows.MessageBox]::Show("Benutzer wurde zur Gruppe '$groupName' hinzugefügt.", "Erfolg", "OK", "Information")
                        Write-LogMessage -Message "Benutzer $selectedUser wurde zur Gruppe $groupName hinzugefügt von $($env:USERNAME)" -LogLevel "INFO"
                    }
                    catch {
                        Write-DebugMessage "Error adding user to group: $($_.Exception.Message)"
                        [System.Windows.MessageBox]::Show("Fehler beim Hinzufügen zur Gruppe: $($_.Exception.Message)", "Fehler", "OK", "Error")
                    }
                }
            }
        } -ErrorMessagePrefix "Fehler beim Hinzufügen einer Gruppe"
    }

    # Remove Group button functionality
    if ($null -ne $btnRemoveGroupUpdate) {
        Register-GUIEvent -Control $btnRemoveGroupUpdate -EventAction {
            if ($null -eq $lstUsersADUpdate -or $null -eq $lstUsersADUpdate.SelectedItem) {
                [System.Windows.MessageBox]::Show("Bitte wählen Sie zuerst einen Benutzer aus.", "Hinweis", "OK", "Information")
                return
            }
            
            if ($null -eq $lstGroupsUpdate -or $null -eq $lstGroupsUpdate.SelectedItem) {
                [System.Windows.MessageBox]::Show("Bitte wählen Sie eine Gruppe aus der Liste aus.", "Hinweis", "OK", "Information")
                return
            }
            
            $selectedUser = $lstUsersADUpdate.SelectedItem.SamAccountName
            if ([string]::IsNullOrWhiteSpace($selectedUser)) {
                [System.Windows.MessageBox]::Show("Ungültiger Benutzername ausgewählt.", "Fehler", "OK", "Error")
                return
            }
            
            $selectedGroup = $lstGroupsUpdate.SelectedItem.ToString()
            if ([string]::IsNullOrWhiteSpace($selectedGroup)) {
                [System.Windows.MessageBox]::Show("Ungültiger Gruppenname ausgewählt.", "Fehler", "OK", "Error")
                return
            }
            
            try {
                # Confirm action to prevent accidental removal
                $confirmation = [System.Windows.MessageBox]::Show(
                    "Möchten Sie den Benutzer '$selectedUser' wirklich aus der Gruppe '$selectedGroup' entfernen?",
                    "Bestätigung", 
                    "YesNo", 
                    "Question"
                )
                
                if ($confirmation -eq "Yes") {
                    # Remove user from group
                    Remove-ADGroupMember -Identity $selectedGroup -Members $selectedUser -Confirm:$false -ErrorAction Stop
                    
                    # Remove item from ListBox
                    if ($null -ne $lstGroupsUpdate) {
                        $lstGroupsUpdate.Items.Remove($selectedGroup)
                    }
                    
                    [System.Windows.MessageBox]::Show("Benutzer wurde aus der Gruppe '$selectedGroup' entfernt.", "Erfolg", "OK", "Information")
                    Write-LogMessage -Message "Benutzer $selectedUser wurde aus der Gruppe $selectedGroup entfernt von $($env:USERNAME)" -LogLevel "INFO"
                }
            }
            catch {
                Write-DebugMessage "Error removing user from group: $($_.Exception.Message)"
                [System.Windows.MessageBox]::Show("Fehler beim Entfernen aus der Gruppe: $($_.Exception.Message)", "Fehler", "OK", "Error")
            }
        } -ErrorMessagePrefix "Fehler beim Entfernen aus einer Gruppe"
    }
} else {
    Write-DebugMessage "AD Update tab (Tab_ADUpdate) not found in XAML."
}
#endregion

#region [Region 21.9 | AD GROUPS TAB]
# [Implements the AD Groups tab functionality]
# [Implementiert die Funktionalität des AD-Gruppen-Tabs]
Write-DebugMessage "AD Groups tab loaded."
$adGroupsTab = $window.FindName("Tab_ADGroups")
if ($adGroupsTab) {
    $btnCreateGroup = $adGroupsTab.FindName("buttonCreate")
    if ($btnCreateGroup) {
        $btnCreateGroup.Add_Click({
            $groupData = @{
                Name        = $adGroupsTab.FindName("textBoxPrefix").Text
                Separator   = $adGroupsTab.FindName("textBoxSeparator").Text
                StartNumber = $adGroupsTab.FindName("textBoxStart").Text
                EndNumber   = $adGroupsTab.FindName("textBoxEnd").Text
                Description = $adGroupsTab.FindName("textBoxDescription").Text
                Email       = $adGroupsTab.FindName("textBoxEmail").Text
            }
            try {
                Process-ADGroups -GroupData $groupData -Config $global:Config
                [System.Windows.MessageBox]::Show("Gruppe erfolgreich erstellt.", "Erfolg")
            }
            catch {
                [System.Windows.MessageBox]::Show("Fehler: $($_.Exception.Message)", "Fehler", 'OK', 'Error')
            }
        })
    }
    $buttonADGroupPreview = $adGroupsTab.FindName("buttonADGroupPreview")
    if ($buttonADGroupPreview) {
        $buttonADGroupPreview.Add_Click({
            [System.Windows.MessageBox]::Show("Vorschau der Gruppennamen.", "Vorschau")
        })
    }
    $buttonADGroupCreate = $adGroupsTab.FindName("buttonADGroupCreate")
    if ($buttonADGroupCreate) {
        $buttonADGroupCreate.Add_Click({
            # Hier kann der Code für die Gruppenerstellung ergänzt werden.
            [System.Windows.MessageBox]::Show("AD Group wird erstellt.", "Erstellung")
        })
    }
    $buttonADGroupClose = $adGroupsTab.FindName("buttonADGroupClose")
    if ($buttonADGroupClose) {
        $buttonADGroupClose.Add_Click({
            $window.Close()
        })
    }
}
#endregion

Write-DebugMessage "GUI: TAB Tab Settings geladen"

#region [Region 21.10 | SETTINGS TAB]
# [Implements the Settings tab functionality]
# [Implementiert die Funktionalität des Einstellungen-Tabs]
Write-DebugMessage "Settings tab loaded."
$settingsTab = $window.FindName("Tab_Settings")
if ($settingsTab) {
    $btnOpenEditor = $settingsTab.FindName("btnOpenEditor")
    if ($btnOpenEditor) {
        $btnOpenEditor.Add_Click({ Open-INIEditor })
    }
    $btnSaveChanges = $settingsTab.FindName("btnSaveChanges")
    if ($btnSaveChanges) {
        $btnSaveChanges.Add_Click({
            [System.Windows.MessageBox]::Show("INI-Datei gespeichert.", "Save")
        })
    }
    $btnInfoSettings = $settingsTab.FindName("btnInfoSettings")
    if ($btnInfoSettings) {
        $btnInfoSettings.Add_Click({
            [System.Windows.MessageBox]::Show("Informationen zum INI-Editor.", "Info")
        })
    }
    $btnCloseSettings = $settingsTab.FindName("btnCloseSettings")
    if ($btnCloseSettings) {
        $btnCloseSettings.Add_Click({
            $window.Close()
        })
    }
}
#endregion
#endregion

Write-DebugMessage "Main GUI execution started."

#region [Region 22 | MAIN GUI EXECUTION]
# [Main code block that starts and handles the GUI dialog]
# [Hauptcodeblock, der den GUI-Dialog startet und verarbeitet]

# [99.22.0 | MAIN GUI INITIALIZATION]
# ENGLISH - Debug message indicating the start of the main GUI execution
# GERMAN - Debug-Nachricht, die den Beginn der Hauptausführung der GUI anzeigt
Write-DebugMessage "GUI: Hauptausführung der GUI"

# Apply branding settings from INI file before showing the window
try {
    Write-DebugMessage "Applying branding settings from INI file"
    
    # Set window properties
    if ($global:Config.Contains("Branding-GUI")) {
        # Set window title with dynamic replacements
        $window.Title = $global:Config["Branding-GUI"]["GUI_Header"] -replace "{ScriptVersion}", 
            $global:Config["ScriptInfo"]["ScriptVersion"] -replace "{LastUpdate}", 
            $global:Config["ScriptInfo"]["LastUpdate"] -replace "{Author}", 
            $global:Config["ScriptInfo"]["Author"]
            
        # Set application name if specified
        if ($global:Config["Branding-GUI"].Contains("APPName")) {
            $window.Title = $global:Config["Branding-GUI"]["APPName"]
        }
        
        # Set window theme color if specified and valid
        if ($global:Config["Branding-GUI"].Contains("ThemeColor")) {
            try {
                $color = $global:Config["Branding-GUI"]["ThemeColor"]
                # Try to create a brush from the color string
                $window.Background = [System.Windows.Media.Brushes]::$color
            } catch {
                Write-DebugMessage "Invalid ThemeColor: $color"
            }
        }
        
        # Set font properties
        if ($global:Config["Branding-GUI"].Contains("FontFamily")) {
            $window.FontFamily = New-Object System.Windows.Media.FontFamily($global:Config["Branding-GUI"]["FontFamily"])
        }
        
        if ($global:Config["Branding-GUI"].Contains("FontSize")) {
            try {
                $window.FontSize = [double]$global:Config["Branding-GUI"]["FontSize"]
            } catch {
                Write-DebugMessage "Invalid FontSize value"
            }
        }
        
        # Set background image if specified
        if ($global:Config["Branding-GUI"].Contains("BackgroundImage")) {
            try {
                $bgImagePath = $global:Config["Branding-GUI"]["BackgroundImage"]
                if (Test-Path $bgImagePath) {
                    $bgImage = New-Object System.Windows.Media.Imaging.BitmapImage
                    $bgImage.BeginInit()
                    $bgImage.UriSource = New-Object System.Uri($bgImagePath, [System.UriKind]::Absolute)
                    $bgImage.EndInit()
                    $window.Background = New-Object System.Windows.Media.ImageBrush($bgImage)
                }
            } catch {
                Write-DebugMessage "Error setting background image: $($_.Exception.Message)"
            }
        }
        
        # Set header logo
        $picLogo = $window.FindName("picLogo")
        if ($picLogo -and $global:Config["Branding-GUI"].Contains("HeaderLogo")) {
            try {
                $logoPath = $global:Config["Branding-GUI"]["HeaderLogo"]
                if (Test-Path $logoPath) {
                    $logo = New-Object System.Windows.Media.Imaging.BitmapImage
                    $logo.BeginInit()
                    $logo.UriSource = New-Object System.Uri($logoPath, [System.UriKind]::Absolute) 
                    $logo.EndInit()
                    $picLogo.Source = $logo
                    
                    # Add click event if URL is specified
                    if ($global:Config["Branding-GUI"].Contains("HeaderLogoURL")) {
                        $picLogo.Cursor = [System.Windows.Input.Cursors]::Hand
                        $picLogo.Add_MouseLeftButtonUp({
                            $url = $global:Config["Branding-GUI"]["HeaderLogoURL"]
                            if (-not [string]::IsNullOrEmpty($url)) {
                                Start-Process $url
                            }
                        })
                    }
                }
            } catch {
                Write-DebugMessage "Error setting header logo: $($_.Exception.Message)"
            }
        }
        
        # Set footer website hyperlink
        $linkLabel = $window.FindName("linkLabel")
        if ($linkLabel -and $global:Config["Branding-GUI"].Contains("FooterWebseite")) {
            try {
                $websiteUrl = $global:Config["Branding-GUI"]["FooterWebseite"]
                $linkLabel.Inlines.Clear()
                
                # Create a proper hyperlink with text
                $hyperlink = New-Object System.Windows.Documents.Hyperlink
                $hyperlink.NavigateUri = New-Object System.Uri($websiteUrl)
                $hyperlink.Inlines.Add($websiteUrl)
                $hyperlink.Add_RequestNavigate({
                    param($sender, $e)
                    Start-Process $e.Uri.AbsoluteUri
                    $e.Handled = $true
                })
                
                $linkLabel.Inlines.Add($hyperlink)
            } catch {
                Write-DebugMessage "Error setting footer website link: $($_.Exception.Message)"
            }
        }
        
        # Set footer info text
        $footerInfo = $window.FindName("footerInfo")
        if ($footerInfo -and $global:Config["Branding-GUI"].Contains("GUI_ExtraText")) {
            $footerInfo.Text = $global:Config["Branding-GUI"]["GUI_ExtraText"]
        }
    }
    
    Write-DebugMessage "Branding settings applied successfully"
} catch {
    Write-DebugMessage "Error applying branding settings: $($_.Exception.Message)"
    Write-DebugMessage "Stack trace: $($_.ScriptStackTrace)"
}

# [22.1 - Shows main window and handles any GUI initialization errors]
# [Zeigt das Hauptfenster an und behandelt alle Fehler bei der GUI-Initialisierung]
try {
    $result = $window.ShowDialog()
    Write-DebugMessage "GUI started successfully, result: $result"
} catch {
    Write-DebugMessage "ERROR: GUI could not be started!"
    Write-DebugMessage "Error message: $($_.Exception.Message)"
    Write-DebugMessage "Error details: $($_.InvocationInfo.PositionMessage)"
    exit 1
}
#endregion

Write-DebugMessage "Main GUI execution completed."