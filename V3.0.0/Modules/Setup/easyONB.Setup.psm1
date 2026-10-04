#Requires -Version 7.2
<#
    easyONB.Setup
    Ersteinrichtung: erzeugt Config\easyONB.ini aus der Vorlage (Install-easyONBOARDING.ps1).

    - Auf einem Domänencontroller werden Domäne, UPN-Suffixe, Domänencontroller, OUs sowie (falls
      vorhanden) Entra Connect und Exchange aus dem Active Directory ermittelt und vorbelegt.
    - Auf einem Client werden alle Werte manuell eingetragen.
    - Die Felder sind in $script:SetupFields beschrieben (Abschnitt, Schlüssel, Hinweistext). Die
      Oberfläche erzeugt Eingabefelder und Hinweistexte daraus.
    - Die erzeugte Datei wird mit derselben Konfigurationsprüfung geladen wie in der Anwendung.
      Kommentare der Vorlage bleiben erhalten; Kennwörter oder andere Secrets werden nie geschrieben.
#>
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Oberflächenfunktionen ändern nur den Zustand des Fensters; geschrieben wird über New-EobSetupConfiguration (ShouldProcess).')]
param()

Set-StrictMode -Version 3.0

#region Felder

$script:SetupPages = [ordered]@{
    Company      = 'Unternehmen'
    Directory    = 'Active Directory'
    Files        = 'Dateien und E-Mail'
    Integrations = 'Integrationen'
    Security     = 'Sicherheit und Darstellung'
}

# Kind: Text, Combo (feste Auswahl), EditCombo (Auswahl oder freie Eingabe), Check.
# BoolFormat: Number (1/0) oder Word (True/False), je nach Schreibweise der Vorlage.
$script:SetupFields = @(
    @{
        Name = 'CompanyName'; Page = 'Company'; Group = 'Firma'; Section = 'Company'; Key = 'CompanyNameFirma'; Label = 'Firmenname'; Kind = 'Text'; Required = $true
        Hint = 'Name des Unternehmens. Wird in das AD-Attribut "company", in Berichte und in die Welcome-Mail übernommen.'
    }
    @{
        Name = 'UpnSuffix'; Page = 'Company'; Group = 'Firma'; Section = 'Company'; Key = 'CompanyActiveDirectoryDomain'; Label = 'UPN-Suffix'; Kind = 'EditCombo'; Required = $true
        Hint = 'Domänenteil der Anmeldung (benutzer@suffix). Auf einem Domänencontroller stehen alle UPN-Suffixe der Gesamtstruktur zur Auswahl.'
    }
    @{
        Name = 'MailDomain'; Page = 'Company'; Group = 'Firma'; Section = 'Company'; Key = 'CompanyMailDomain'; Label = 'E-Mail-Domäne'; Kind = 'Text'; Required = $true
        Hint = 'Domäne der E-Mail-Adressen neuer Konten, z. B. @example.com.'
    }
    @{
        Name = 'Ms365Domain'; Page = 'Company'; Group = 'Firma'; Section = 'Company'; Key = 'CompanyMS365Domain'; Label = 'Microsoft-365-Domäne'; Kind = 'Text'
        Hint = 'Optional: @<mandant>.onmicrosoft.com. Wird als zusätzliche Proxyadresse gesetzt.'
    }
    @{
        Name = 'Website'; Page = 'Company'; Group = 'Firma'; Section = 'Company'; Key = 'CompanyDomain'; Label = 'Webseite'; Kind = 'Text'
        Hint = 'Webseite des Unternehmens (www.example.com). Erscheint im Willkommensdokument und hinter dem Logo.'
    }
    @{
        Name = 'Country'; Page = 'Company'; Group = 'Firma'; Section = 'Company'; Key = 'CompanyCountry'; Label = 'Land (ISO-Code)'; Kind = 'Text'; Default = 'DE'
        Hint = 'Zweistelliger Ländercode nach ISO 3166, z. B. DE, AT oder CH. Wird in das AD-Attribut "c" geschrieben.'
    }
    @{
        Name = 'Street'; Page = 'Company'; Group = 'Anschrift'; Section = 'Company'; Key = 'CompanyStrasse'; Label = 'Straße'; Kind = 'Text'
        Hint = 'Anschrift für das AD-Attribut "streetAddress" neuer Konten.'
    }
    @{
        Name = 'PostalCode'; Page = 'Company'; Group = 'Anschrift'; Section = 'Company'; Key = 'CompanyPLZ'; Label = 'Postleitzahl'; Kind = 'Text'
        Hint = 'Postleitzahl für das AD-Attribut "postalCode".'
    }
    @{
        Name = 'City'; Page = 'Company'; Group = 'Anschrift'; Section = 'Company'; Key = 'CompanyOrt'; Label = 'Ort'; Kind = 'Text'
        Hint = 'Ort für das AD-Attribut "l".'
    }
    @{
        Name = 'Phone'; Page = 'Company'; Group = 'Anschrift'; Section = 'Company'; Key = 'CompanyTelefon'; Label = 'Zentrale Rufnummer'; Kind = 'Text'
        Hint = 'Zentrale Rufnummer des Unternehmens (Willkommensdokument).'
    }
    @{
        Name = 'ItName'; Page = 'Company'; Group = 'IT-Service'; Section = 'CompanyHelpdesk'; Key = 'CompanyITMitarbeiter'; Label = 'Name der IT-Abteilung'; Kind = 'Text'; Default = 'IT-Service'
        Hint = 'Absender in Berichten und Willkommensdokumenten.'
    }
    @{
        Name = 'HelpdeskMail'; Page = 'Company'; Group = 'IT-Service'; Section = 'CompanyHelpdesk'; Key = 'CompanyHelpdeskMail'; Label = 'E-Mail des Helpdesks'; Kind = 'Text'
        Hint = 'Kontaktadresse für neue Mitarbeitende. Wird auch in der Abwesenheitsnotiz beim Offboarding genannt.'
    }
    @{
        Name = 'HelpdeskPhone'; Page = 'Company'; Group = 'IT-Service'; Section = 'CompanyHelpdesk'; Key = 'CompanyHelpdeskTel'; Label = 'Telefon des Helpdesks'; Kind = 'Text'
        Hint = 'Rufnummer des Helpdesks im Willkommensdokument.'
    }

    @{
        Name = 'DomainController'; Page = 'Directory'; Group = 'Verbindung'; Section = 'ActiveDirectory'; Key = 'PreferredDomainController'; Label = 'Domänencontroller'; Kind = 'EditCombo'
        Hint = 'Fester, beschreibbarer Domänencontroller für alle Änderungen. Leer = automatische Auswahl.'
    }
    @{
        Name = 'SearchBase'; Page = 'Directory'; Group = 'Verbindung'; Section = 'ActiveDirectory'; Key = 'SearchBase'; Label = 'Suchbasis (optional)'; Kind = 'EditCombo'
        Hint = 'Schränkt die Benutzersuche auf eine OU ein (Distinguished Name). Leer = gesamte Domäne.'
    }
    @{
        Name = 'DefaultOU'; Page = 'Directory'; Group = 'Organisationseinheiten'; Section = 'ADUserDefaults'; Key = 'DefaultOU'; Label = 'Ziel-OU neuer Konten'; Kind = 'EditCombo'; Required = $true
        Hint = 'OU, in der neue Konten angelegt werden (Distinguished Name, z. B. OU=Mitarbeiter,DC=example,DC=local).'
    }
    @{
        Name = 'DisabledUsersOU'; Page = 'Directory'; Group = 'Organisationseinheiten'; Section = 'Offboarding'; Key = 'DisabledUsersOU'; Label = 'OU ausgeschiedener Konten'; Kind = 'EditCombo'
        Hint = 'Ziel der Offboarding-Aktion "In OU verschieben". Leer = Konten bleiben in ihrer OU.'
    }
    @{
        Name = 'InitialGroups'; Page = 'Directory'; Group = 'Konten'; Section = 'UserCreationDefaults'; Key = 'InitialGroupMembership'; Label = 'Standardgruppen'; Kind = 'Text'
        Hint = 'Gruppen, die jedes neue Konto erhält (durch ; getrennt). Privilegierte Gruppen werden nie zugewiesen.'
    }
    @{
        Name = 'AccountDisabled'; Page = 'Directory'; Group = 'Konten'; Section = 'ADUserDefaults'; Key = 'AccountDisabled'; Label = 'Neue Konten deaktiviert anlegen'; Kind = 'Check'; BoolFormat = 'Word'; Default = $true
        Hint = 'Empfohlen: Das Konto wird erst bei der Übergabe aktiviert (Benutzer aktualisieren > Konto aktivieren).'
    }
    @{
        Name = 'SyncGroupEnabled'; Page = 'Directory'; Group = 'Konten'; Section = 'ActivateUserMS365ADSync'; Key = 'ADSync'; Label = 'M365-Synchronisationsgruppe nutzen'; Kind = 'Check'; BoolFormat = 'Number'; Default = $false
        Hint = 'Neue Konten werden der Gruppe hinzugefügt, über die Entra Connect synchronisiert. Diese Gruppe wird beim Offboarding nie entfernt.'
    }
    @{
        Name = 'SyncGroup'; Page = 'Directory'; Group = 'Konten'; Section = 'ActivateUserMS365ADSync'; Key = 'ADSyncADGroup'; Label = 'Synchronisationsgruppe'; Kind = 'Text'
        Hint = 'Name der AD-Gruppe für die Microsoft-365-Synchronisation.'
    }

    @{
        Name = 'LogDirectory'; Page = 'Files'; Group = 'Ablage'; Section = 'Logging'; Key = 'LogFile'; Label = 'Log-Verzeichnis'; Kind = 'Text'; Default = 'Logs'
        Hint = 'Text- und Audit-Logs. Relative Pfade beziehen sich auf den Anwendungsordner; %ProgramData% ist erlaubt.'
    }
    @{
        Name = 'ReportDirectory'; Page = 'Files'; Group = 'Ablage'; Section = 'Report'; Key = 'ReportPath'; Label = 'Berichtsverzeichnis'; Kind = 'Text'; Default = 'Reports'
        Hint = 'Ablage der Berichte (HTML, JSON, PDF). Enthält personenbezogene Daten, aber nie Kennwörter.'
    }
    @{
        Name = 'AllowedRoots'; Page = 'Files'; Group = 'Dateiserver'; Section = 'FileServer'; Key = 'AllowedRoots'; Label = 'Stammpfade Home-Verzeichnisse'; Kind = 'Text'
        Hint = 'Freigaben, unter denen Home-Verzeichnisse angelegt oder archiviert werden dürfen (durch ; getrennt). Leer = keine Dateiserver-Aktionen.'
    }
    @{
        Name = 'HomeDirectoryPath'; Page = 'Files'; Group = 'Dateiserver'; Section = 'OnboardingExtensions'; Key = 'HomeDirectoryPath'; Label = 'Pfad Home-Verzeichnis'; Kind = 'Text'
        Hint = 'Vorlage für neue Home-Verzeichnisse, %username% wird ersetzt. Muss unter einem Stammpfad liegen.'
    }
    @{
        Name = 'ArchiveRoot'; Page = 'Files'; Group = 'Dateiserver'; Section = 'FileServer'; Key = 'ArchiveRoot'; Label = 'Archiv für Home-Verzeichnisse'; Kind = 'Text'
        Hint = 'Ziel beim Offboarding (Home-Verzeichnis archivieren). Muss unter einem Stammpfad liegen.'
    }
    @{
        Name = 'SmtpServer'; Page = 'Files'; Group = 'E-Mail (SMTP)'; Section = 'EmailSettings'; Key = 'SMTPServer'; Label = 'SMTP-Server'; Kind = 'Text'
        Hint = 'Relay für die Welcome-Mail (ohne Anmeldung oder mit dem Windows-Konto). Leer = keine E-Mails.'
    }
    @{
        Name = 'SmtpPort'; Page = 'Files'; Group = 'E-Mail (SMTP)'; Section = 'EmailSettings'; Key = 'SMTPPort'; Label = 'SMTP-Port'; Kind = 'Text'; Default = '25'
        Hint = 'Üblich: 25 (Relay) oder 587 (STARTTLS).'
    }
    @{
        Name = 'SmtpSsl'; Page = 'Files'; Group = 'E-Mail (SMTP)'; Section = 'EmailSettings'; Key = 'UseSSL'; Label = 'Verschlüsselt senden (TLS)'; Kind = 'Check'; BoolFormat = 'Number'; Default = $true
        Hint = 'Verbindung zum SMTP-Server mit TLS absichern.'
    }
    @{
        Name = 'MailFrom'; Page = 'Files'; Group = 'E-Mail (SMTP)'; Section = 'EmailSettings'; Key = 'FromAddress'; Label = 'Absenderadresse'; Kind = 'Text'
        Hint = 'Absender der Welcome-Mail, z. B. it-service@example.com.'
    }

    @{
        Name = 'ExchangeMode'; Page = 'Integrations'; Group = 'Exchange'; Section = 'Exchange'; Key = 'Mode'; Label = 'Exchange-Betriebsart'; Kind = 'Combo'; Options = @('None', 'Online', 'OnPremises', 'Hybrid'); Default = 'None'
        Hint = 'None = keine Postfachaktionen. Online = Exchange Online, OnPremises = Exchange Server, Hybrid = beides.'
    }
    @{
        Name = 'ExchangeUri'; Page = 'Integrations'; Group = 'Exchange'; Section = 'Exchange'; Key = 'OnPremisesUri'; Label = 'Exchange-Server-Endpunkt'; Kind = 'Text'
        Hint = 'PowerShell-Endpunkt des Exchange Servers (Kerberos), z. B. http://exchange.example.local/PowerShell/.'
    }
    @{
        Name = 'RoutingDomain'; Page = 'Integrations'; Group = 'Exchange'; Section = 'Exchange'; Key = 'RemoteRoutingDomain'; Label = 'Routingdomäne (Hybrid)'; Kind = 'Text'
        Hint = 'Nur Hybrid: <mandant>.mail.onmicrosoft.com für neue Remote-Postfächer.'
    }
    @{
        Name = 'GraphEnabled'; Page = 'Integrations'; Group = 'Microsoft Graph'; Section = 'Graph'; Key = 'Enabled'; Label = 'Microsoft Graph nutzen'; Kind = 'Check'; BoolFormat = 'Number'; Default = $false
        Hint = 'Widerruft beim Offboarding die Anmeldesitzungen in Microsoft 365.'
    }
    @{
        Name = 'TenantId'; Page = 'Integrations'; Group = 'Microsoft Graph'; Section = 'Graph'; Key = 'TenantId'; Label = 'Mandant (Tenant)'; Kind = 'Text'
        Hint = 'Tenant-ID oder Domäne, z. B. example.onmicrosoft.com.'
    }
    @{
        Name = 'SyncEnabled'; Page = 'Integrations'; Group = 'Entra Connect Sync'; Section = 'ADSync'; Key = 'EnableADSync'; Label = 'Synchronisation anbieten'; Kind = 'Check'; BoolFormat = 'Number'; Default = $false
        Hint = 'Bietet nach dem Onboarding an, einen Synchronisationszyklus zu starten.'
    }
    @{
        Name = 'SyncServer'; Page = 'Integrations'; Group = 'Entra Connect Sync'; Section = 'ADSync'; Key = 'ADSyncServer'; Label = 'Entra-Connect-Server'; Kind = 'Text'
        Hint = 'Server mit Entra Connect (PowerShell-Remoting, Mitgliedschaft in ADSyncOperators).'
    }
    @{
        Name = 'SyncAuto'; Page = 'Integrations'; Group = 'Entra Connect Sync'; Section = 'ADSync'; Key = 'AutoSyncNewUsers'; Label = 'Nach Onboarding automatisch'; Kind = 'Check'; BoolFormat = 'Number'; Default = $false
        Hint = 'Startet nach jedem Onboarding automatisch einen Delta-Zyklus.'
    }

    @{
        Name = 'SimulationByDefault'; Page = 'Security'; Group = 'Sicherheit'; Section = 'Security'; Key = 'SimulationByDefault'; Label = 'Start im Simulationsmodus'; Kind = 'Check'; BoolFormat = 'Number'; Default = $true
        Hint = 'Empfohlen: Die Oberfläche startet ohne Änderungen; der Live-Modus wird bewusst aktiviert.'
    }
    @{
        Name = 'AllowManualPassword'; Page = 'Security'; Group = 'Sicherheit'; Section = 'Security'; Key = 'AllowManualPassword'; Label = 'Manuelle Kennwörter erlauben'; Kind = 'Check'; BoolFormat = 'Number'; Default = $true
        Hint = 'Erlaubt, beim Onboarding ein Kennwort vorzugeben. Sonst wird immer ein sicheres Kennwort erzeugt.'
    }
    @{
        Name = 'BreakGlassAccounts'; Page = 'Security'; Group = 'Sicherheit'; Section = 'Security'; Key = 'BreakGlassAccounts'; Label = 'Notfallkonten'; Kind = 'Text'
        Hint = 'Konten, die nie bearbeitet werden (Anmeldename oder SID, durch ; getrennt).'
    }
    @{
        Name = 'ProtectedGroups'; Page = 'Security'; Group = 'Sicherheit'; Section = 'Security'; Key = 'ProtectedGroups'; Label = 'Zusätzlich geschützte Gruppen'; Kind = 'Text'
        Hint = 'Gruppen, die nie automatisch zugewiesen werden (Domain Admins usw. sind immer geschützt).'
    }
    @{
        Name = 'AppName'; Page = 'Security'; Group = 'Darstellung'; Section = 'WPFGUI'; Key = 'APPName'; Label = 'Anwendungsname'; Kind = 'Text'; Default = 'easyONBOARDING'
        Hint = 'Titel im Kopf der Oberfläche.'
    }
    @{
        Name = 'Theme'; Page = 'Security'; Group = 'Darstellung'; Section = 'UI'; Key = 'Theme'; Label = 'Farbschema'; Kind = 'Combo'; Options = @('Light', 'Dark'); Default = 'Light'
        Hint = 'Helles oder dunkles Farbschema; jederzeit in der Oberfläche umschaltbar.'
    }
    @{
        Name = 'AccentColor'; Page = 'Security'; Group = 'Darstellung'; Section = 'UI'; Key = 'AccentColor'; Label = 'Akzentfarbe'; Kind = 'Text'; Default = '#0F6CBD'
        Hint = 'Farbe für Schaltflächen und Markierungen im Format #RRGGBB.'
    }
)

function Get-EobSetupPage {
    <#
    .SYNOPSIS
        Liefert die Seiten des Installers (Id und Titel) in Anzeigereihenfolge.
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param()

    foreach ($id in $script:SetupPages.Keys) { [pscustomobject]@{ Id = $id; Title = $script:SetupPages[$id] } }
}

function Get-EobSetupField {
    <#
    .SYNOPSIS
        Liefert die Felddefinitionen des Installers (optional für eine Seite).
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param([string]$Page)

    foreach ($field in $script:SetupFields) {
        if ($Page -and $field.Page -ne $Page) { continue }
        [pscustomobject]@{
            Name       = $field.Name
            Page       = $field.Page
            Group      = $field.Group
            Section    = $field.Section
            Key        = $field.Key
            Label      = $field.Label
            Kind       = $field.Kind
            Required   = [bool]$(if ($field.ContainsKey('Required')) { $field.Required } else { $false })
            Default    = $(if ($field.ContainsKey('Default')) { $field.Default } else { $null })
            Options    = @($(if ($field.ContainsKey('Options')) { $field.Options } else { @() }))
            BoolFormat = $(if ($field.ContainsKey('BoolFormat')) { $field.BoolFormat } else { 'Number' })
            Hint       = $field.Hint
        }
    }
}

function Get-EobSetupDefaultValue {
    <#
    .SYNOPSIS
        Standardwerte aller Felder (Name -> Wert); Auswahlfelder ohne Standard bleiben leer.
    #>
    [CmdletBinding()]
    [OutputType([hashtable])]
    param()

    $values = @{}
    foreach ($field in Get-EobSetupField) {
        $values[$field.Name] = if ($null -ne $field.Default) { $field.Default } elseif ($field.Kind -eq 'Check') { $false } else { '' }
    }
    return $values
}

#endregion

#region Umgebung und Active Directory

function Get-EobSetupEnvironment {
    <#
    .SYNOPSIS
        Ermittelt, ob der Installer auf einem Domänencontroller, einem Domänenmitglied oder außerhalb von Windows läuft.
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param()

    $result = [ordered]@{
        IsWindows          = [bool]$IsWindows
        IsDomainController = $false
        PartOfDomain       = $false
        Domain             = ''
        ComputerName       = [Environment]::MachineName
        AdModuleAvailable  = $false
        Detail             = ''
    }
    if ($IsWindows) {
        try {
            # ProductType 2 = Domänencontroller (1 = Arbeitsstation, 3 = Mitgliedsserver)
            $os = Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction Stop -Verbose:$false
            $result.IsDomainController = ([int]$os.ProductType -eq 2)
            $system = Get-CimInstance -ClassName Win32_ComputerSystem -ErrorAction Stop -Verbose:$false
            $result.PartOfDomain = [bool]$system.PartOfDomain
            $result.Domain = [string]$system.Domain
        }
        catch {
            $result.Detail = "Systemabfrage nicht möglich: $($_.Exception.Message)"
        }
        $result.AdModuleAvailable = $null -ne (Get-Module -ListAvailable -Name 'ActiveDirectory' -ErrorAction SilentlyContinue)
    }
    if (-not $result.Detail) {
        $result.Detail = if ($result.IsDomainController) { "Domänencontroller $($result.ComputerName) ($($result.Domain))" }
        elseif ($result.PartOfDomain) { "Domänenmitglied $($result.ComputerName) ($($result.Domain))" }
        elseif ($IsWindows) { "Arbeitsgruppenrechner $($result.ComputerName)" }
        else { 'Kein Windows: nur Erstellung ohne Oberfläche möglich' }
    }
    return [pscustomobject]$result
}

function ConvertFrom-EobEntraConnectAccountDescription {
    <#
    .SYNOPSIS
        Liest Server und Mandant aus der Beschreibung des Entra-Connect-Dienstkontos (MSOL_...).
    .EXAMPLE
        ConvertFrom-EobEntraConnectAccountDescription -Description '... running on computer SYNC01 configured to synchronize to tenant example.onmicrosoft.com. ...'
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param([AllowNull()][AllowEmptyString()][string]$Description)

    if ([string]::IsNullOrWhiteSpace($Description)) { return $null }
    $pattern = '(?i)running on computer (?<computer>[A-Za-z0-9][A-Za-z0-9\-\.]*) configured to synchronize to tenant (?<tenant>[A-Za-z0-9\-]+(?:\.[A-Za-z0-9\-]+)+?)\.?(?:\s|$)'
    if ($Description -notmatch $pattern) { return $null }
    [pscustomobject]@{ Computer = $Matches['computer']; Tenant = $Matches['tenant'] }
}

function Get-EobSetupExchangeServerName {
    <#
    .SYNOPSIS
        Liefert den vollqualifizierten Namen eines Exchange Servers aus dessen networkAddress-Werten.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param([AllowNull()][string[]]$NetworkAddress)

    foreach ($address in @($NetworkAddress)) {
        if ([string]$address -match '^(?i)ncacn_ip_tcp:(?<name>[A-Za-z0-9\-]+(\.[A-Za-z0-9\-]+)+)$') { return $Matches['name'] }
    }
    return ''
}

function Select-EobSetupOu {
    <#
    .SYNOPSIS
        Wählt aus einer OU-Liste die erste OU, deren Name zum Muster passt (sonst den Ersatzwert).
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [AllowEmptyCollection()][string[]]$DistinguishedName = @(),
        [Parameter(Mandatory)][string]$Pattern,
        [AllowEmptyString()][string]$Fallback = ''
    )

    foreach ($dn in $DistinguishedName) {
        if ($dn -match '^OU=(?<name>(?:[^,\\]|\\.)+),' -and $Matches['name'] -match $Pattern) { return $dn }
    }
    return $Fallback
}

function Get-EobSetupDirectoryDefault {
    <#
    .SYNOPSIS
        Ermittelt Vorschlagswerte aus dem Active Directory (für die Ausführung auf einem Domänencontroller).
    .OUTPUTS
        Objekt mit Values (Feldname -> Wert), Options (Feldname -> Auswahlliste) und Notes (Hinweise).
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param([string]$Server)

    Import-Module -Name ActiveDirectory -ErrorAction Stop -Verbose:$false
    $common = @{ ErrorAction = 'Stop' }
    if ($Server) { $common['Server'] = $Server }
    $values = @{}
    $options = @{}
    $notes = [System.Collections.Generic.List[string]]::new()

    $domain = Get-ADDomain @common
    $forest = Get-ADForest @common
    $dnsRoot = [string]$domain.DNSRoot
    $suffixes = @(@($dnsRoot) + @($forest.UPNSuffixes | ForEach-Object { [string]$_ }) | Where-Object { $_ } | Select-Object -Unique)
    $options['UpnSuffix'] = $suffixes
    # Öffentliche Suffixe bevorzugen; interne Namen (.local, .intern ...) eignen sich nicht für E-Mail und Microsoft 365.
    $public = @($suffixes | Where-Object { $_ -notmatch '(?i)\.(local|lan|intern|internal|corp|home|localdomain)$' }) | Select-Object -First 1
    $values['UpnSuffix'] = if ($public) { $public } else { $dnsRoot }
    $values['MailDomain'] = '@' + $values['UpnSuffix']
    if (-not $public) { $notes.Add("Kein öffentlicher UPN-Suffix gefunden: $dnsRoot ist für E-Mail und Microsoft 365 ungeeignet.") }

    $controllers = @(Get-ADDomainController -Filter * @common | ForEach-Object { [string]$_.HostName } | Sort-Object)
    $options['DomainController'] = $controllers
    $self = @($controllers | Where-Object { ($_ -split '\.')[0] -ieq [Environment]::MachineName }) | Select-Object -First 1
    $values['DomainController'] = if ($self) { $self } else { [string]$domain.PDCEmulator }

    $ous = @(Get-ADOrganizationalUnit -Filter * -Properties CanonicalName -ResultSetSize 2000 @common | Sort-Object -Property CanonicalName | ForEach-Object { [string]$_.DistinguishedName })
    $usersContainer = [string]$domain.UsersContainer
    $options['DefaultOU'] = @($ous + @($usersContainer) | Where-Object { $_ })
    $options['DisabledUsersOU'] = $ous
    $options['SearchBase'] = @(@([string]$domain.DistinguishedName) + $ous)
    $values['DefaultOU'] = Select-EobSetupOu -DistinguishedName $ous -Pattern '(?i)^(Mitarbeiter|Benutzer|Users|Employees|Staff)$' -Fallback $usersContainer
    $values['DisabledUsersOU'] = Select-EobSetupOu -DistinguishedName $ous -Pattern '(?i)^(Ausgeschieden|Deaktiviert|Deaktivierte Benutzer|Disabled( Users)?|Former( Employees)?)$'
    if ($values['DefaultOU'] -eq $usersContainer) { $notes.Add('Keine Mitarbeiter-OU gefunden: Standard-Container CN=Users vorgeschlagen.') }

    $tenant = ''
    try {
        $accounts = @(Get-ADUser -LDAPFilter '(|(sAMAccountName=MSOL_*)(sAMAccountName=AAD_*))' -Properties Description @common)
        foreach ($account in $accounts) {
            $info = ConvertFrom-EobEntraConnectAccountDescription -Description ([string]$account.Description)
            if ($null -eq $info) { continue }
            $computer = if ($info.Computer.Contains('.')) { $info.Computer } else { "$($info.Computer).$dnsRoot" }
            $tenant = $info.Tenant
            $values['SyncServer'] = $computer
            $values['SyncEnabled'] = $true
            $values['TenantId'] = $tenant
            if ($tenant -match '(?i)\.onmicrosoft\.com$') { $values['Ms365Domain'] = '@' + $tenant }
            $notes.Add("Entra Connect gefunden: $computer, Mandant $tenant.")
            break
        }
    }
    catch {
        $notes.Add("Entra Connect nicht ermittelbar: $($_.Exception.Message)")
    }

    try {
        $configurationContext = [string](Get-ADRootDSE @common).configurationNamingContext
        $servers = @(Get-ADObject -SearchBase "CN=Microsoft Exchange,CN=Services,$configurationContext" -LDAPFilter '(objectClass=msExchExchangeServer)' -Properties networkAddress @common)
        foreach ($exchangeServer in $servers) {
            $name = Get-EobSetupExchangeServerName -NetworkAddress @($exchangeServer.networkAddress)
            if (-not $name) { continue }
            $values['ExchangeMode'] = if ($tenant) { 'Hybrid' } else { 'OnPremises' }
            $values['ExchangeUri'] = "http://$name/PowerShell/"
            if ($tenant -match '(?i)^(?<prefix>[a-z0-9\-]+)\.onmicrosoft\.com$') { $values['RoutingDomain'] = "$($Matches['prefix']).mail.onmicrosoft.com" }
            $notes.Add("Exchange Server gefunden: $name (Betriebsart $($values['ExchangeMode']) vorgeschlagen).")
            break
        }
    }
    catch {
        # Ohne Exchange-Organisation existiert der Container nicht; das ist kein Fehler.
        Write-Verbose "Keine Exchange-Organisation gefunden: $($_.Exception.Message)"
    }

    [pscustomobject]@{
        Domain  = $dnsRoot
        Values  = $values
        Options = $options
        Notes   = $notes.ToArray()
    }
}

#endregion

#region INI erzeugen

function ConvertTo-EobSetupIniText {
    <#
    .SYNOPSIS
        Erzeugt den INI-Text aus der Vorlage: ersetzt Werte, entfernt Beispielschlüssel und ergänzt fehlende Schlüssel.
    .DESCRIPTION
        Kommentare, Leerzeilen und Reihenfolge der Vorlage bleiben erhalten.
    .PARAMETER Values
        Abschnitt -> (Schlüssel -> Wert).
    .PARAMETER RemoveSection
        Abschnitte, deren Beispielschlüssel entfernt werden (außer den in Values gesetzten).
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory)][AllowEmptyString()][string]$TemplateText,
        [Parameter(Mandatory)][System.Collections.IDictionary]$Values,
        [string[]]$RemoveSection = @()
    )

    foreach ($section in $Values.Keys) {
        foreach ($key in $Values[$section].Keys) {
            if ([string]$Values[$section][$key] -match '[\r\n]') { throw "[$section] $key enthält einen Zeilenumbruch." }
        }
    }
    $lookup = @{}
    foreach ($section in $Values.Keys) { $lookup[$section.ToLowerInvariant()] = $section }
    $written = @{}
    $output = [System.Collections.Generic.List[string]]::new()
    $current = $null

    $flush = {
        param([string]$SectionName)
        if (-not $SectionName -or -not $lookup.ContainsKey($SectionName.ToLowerInvariant())) { return }
        $name = $lookup[$SectionName.ToLowerInvariant()]
        $pending = @($Values[$name].Keys | Where-Object { -not $written.ContainsKey("$name|$_".ToLowerInvariant()) })
        if ($pending.Count -eq 0) { return }
        # Vor abschließenden Leerzeilen des Abschnitts einfügen
        $insertAt = $output.Count
        while ($insertAt -gt 0 -and [string]::IsNullOrWhiteSpace($output[$insertAt - 1])) { $insertAt-- }
        foreach ($key in $pending) {
            $output.Insert($insertAt, "$key=$($Values[$name][$key])")
            $insertAt++
            $written["$name|$key".ToLowerInvariant()] = $true
        }
    }

    foreach ($line in ($TemplateText -replace '^﻿', '' -split '\r?\n')) {
        $trimmed = $line.Trim()
        if ($trimmed -match '^\[(?<name>[^\]]+)\]$') {
            & $flush $current
            $current = $Matches['name'].Trim()
            $output.Add($line)
            continue
        }
        if ($null -ne $current -and $trimmed -and -not $trimmed.StartsWith(';') -and -not $trimmed.StartsWith('#') -and $trimmed.Contains('=')) {
            $key = $trimmed.Substring(0, $trimmed.IndexOf('=')).Trim()
            $sectionKey = $current.ToLowerInvariant()
            if ($lookup.ContainsKey($sectionKey)) {
                $name = $lookup[$sectionKey]
                $match = @($Values[$name].Keys | Where-Object { $_ -ieq $key }) | Select-Object -First 1
                if ($match) {
                    $output.Add("$key=$($Values[$name][$match])")
                    $written["$name|$match".ToLowerInvariant()] = $true
                    continue
                }
            }
            if (@($RemoveSection | Where-Object { $_ -ieq $current }).Count -gt 0) { continue }
        }
        $output.Add($line)
    }
    & $flush $current

    # Abschnitte, die in der Vorlage fehlen, am Ende anfügen
    $templateSections = @([regex]::Matches($TemplateText, '(?m)^\s*\[([^\]]+)\]') | ForEach-Object { $_.Groups[1].Value.Trim().ToLowerInvariant() })
    foreach ($section in $Values.Keys) {
        if ($section.ToLowerInvariant() -in $templateSections) { continue }
        while ($output.Count -gt 0 -and [string]::IsNullOrWhiteSpace($output[$output.Count - 1])) { $output.RemoveAt($output.Count - 1) }
        $output.Add('')
        $output.Add("[$section]")
        foreach ($key in $Values[$section].Keys) { $output.Add("$key=$($Values[$section][$key])") }
    }
    while ($output.Count -gt 0 -and [string]::IsNullOrWhiteSpace($output[$output.Count - 1])) { $output.RemoveAt($output.Count - 1) }
    return (($output -join "`r`n") + "`r`n")
}

function ConvertTo-EobSetupBoolText {
    [CmdletBinding()]
    [OutputType([string])]
    param([AllowNull()][object]$Value, [ValidateSet('Number', 'Word')][string]$Format = 'Number')

    $flag = ConvertTo-EobBoolean -Value $Value
    if ($Format -eq 'Word') { return $(if ($flag) { 'True' } else { 'False' }) }
    return $(if ($flag) { '1' } else { '0' })
}

function Get-EobSetupIniValue {
    <#
    .SYNOPSIS
        Übersetzt die Feldwerte in INI-Werte inklusive abgeleiteter Werte (Berichtstexte, Webseite, Abwesenheitsnotiz).
    .PARAMETER RemoveSamples
        Entfernt die Beispieldaten der Vorlage (Gruppen, Links, WLAN, VPN, Platzhalter).
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory)][hashtable]$FieldValues,
        [switch]$RemoveSamples
    )

    $values = [ordered]@{}
    $set = { param([string]$Section, [string]$Key, [AllowNull()][object]$Value)
        if (-not $values.Contains($Section)) { $values[$Section] = [ordered]@{} }
        $values[$Section][$Key] = if ($null -eq $Value) { '' } else { ([string]$Value).Trim() }
    }
    $read = { param([string]$Name) if ($FieldValues.ContainsKey($Name) -and $null -ne $FieldValues[$Name]) { ([string]$FieldValues[$Name]).Trim() } else { '' } }

    foreach ($field in Get-EobSetupField) {
        $value = if ($FieldValues.ContainsKey($field.Name)) { $FieldValues[$field.Name] } else { $field.Default }
        if ($field.Kind -eq 'Check') { $value = ConvertTo-EobSetupBoolText -Value $value -Format $field.BoolFormat }
        & $set $field.Section $field.Key $value
    }

    # Normalisierung
    foreach ($name in @('MailDomain', 'Ms365Domain')) {
        $domain = & $read $name
        if ($domain -and -not $domain.StartsWith('@')) { $domain = '@' + $domain }
        $field = Get-EobSetupField | Where-Object Name -EQ $name
        & $set $field.Section $field.Key $domain
    }
    $company = & $read 'CompanyName'
    $itName = & $read 'ItName'
    if (-not $itName) { $itName = 'IT-Service' }
    $website = ((& $read 'Website') -replace '^(?i)https?://', '').TrimEnd('/')
    & $set 'Company' 'CompanyDomain' $website
    $helpdesk = & $read 'HelpdeskMail'
    $mailDomain = $values['Company']['CompanyMailDomain']

    # Abgeleitete Werte
    & $set 'Report' 'ReportHeader' $(if ($company) { "Willkommen bei $company" } else { 'Willkommen' })
    & $set 'Report' 'ReportFooter' $(if ($company) { "$company - $itName" } else { $itName })
    & $set 'EmailSettings' 'WelcomeEmailSubject' $(if ($company) { "Willkommen bei $company" } else { 'Willkommen' })
    & $set 'WPFGUI' 'LogoURL' $(if ($website) { "https://$website" } else { '' })
    & $set 'WPFGUI' 'FooterWebseite' $website
    & $set 'ReportPlaceholders' 'CompanyWebsite' $(if ($website) { "https://$website" } else { '' })
    $autoReply = 'Vielen Dank für Ihre Nachricht. Die Person ist nicht mehr im Unternehmen tätig.'
    if ($helpdesk) { $autoReply += " Bitte wenden Sie sich an $helpdesk." }
    & $set 'Offboarding' 'AutoReplyMessage' $autoReply
    if ($mailDomain) { & $set 'MailEndungen' 'Domain1' $mailDomain }

    $remove = [System.Collections.Generic.List[string]]::new()
    $remove.Add('MailEndungen')
    if ($RemoveSamples) {
        foreach ($section in @('Websites', 'ADGroups', 'TLGroups', 'LicensesGroups')) { $remove.Add($section) }
        & $set 'LicensesGroups' 'KEINE' ''
        & $set 'ALGroup' 'Group' ''
        & $set 'UserCreationDefaults' 'DefaultTLGroup' ''
        & $set 'EmailSettings' 'CopyAddress' ''
        foreach ($key in @('CompanyWikiURL', 'CompanyIntranetURL', 'CompanyHREmail', 'CompanyHRPhone', 'CompanyLocations', 'FAQURL', 'ITTrainingURL')) {
            & $set 'ReportPlaceholders' $key ''
        }
        foreach ($key in @('CompanySSID', 'CompanySSIDbyod', 'CompanySSIDGuest')) { & $set 'CompanyWLAN' $key '' }
        & $set 'CompanyVPN' 'CompanyVPNDomain' ''
    }
    [pscustomobject]@{ Values = $values; RemoveSection = $remove.ToArray() }
}

function Test-EobSetupFieldValue {
    <#
    .SYNOPSIS
        Prüft Feldwerte vor dem Schreiben (Pflichtfelder und Schematypen der Konfiguration).
    .OUTPUTS
        Befunde (Severity, Code, Field, Message).
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param([Parameter(Mandatory)][hashtable]$FieldValues)

    $schema = Get-EobConfigSchema
    foreach ($field in Get-EobSetupField) {
        $value = if ($FieldValues.ContainsKey($field.Name) -and $null -ne $FieldValues[$field.Name]) { ([string]$FieldValues[$field.Name]).Trim() } else { '' }
        if ($field.Kind -eq 'Check') { continue }
        if ($field.Required -and -not $value) {
            New-EobFinding -Severity Error -Code 'SETUP_REQUIRED' -Field $field.Name -Message "$($field.Label) ist ein Pflichtfeld."
            continue
        }
        if (-not $value) { continue }
        if ($value -match '[\r\n]') {
            New-EobFinding -Severity Error -Code 'SETUP_MULTILINE' -Field $field.Name -Message "$($field.Label): Zeilenumbrüche sind nicht zulässig."
            continue
        }
        $sectionDefinition = Get-EobSchemaSectionDefinition -Schema $schema -Name $field.Section
        if ($null -eq $sectionDefinition) { continue }
        $keyDefinition = Get-EobSchemaKeyDefinition -SectionDefinition $sectionDefinition -Key $field.Key
        if ($null -eq $keyDefinition) { continue }
        $problem = Test-EobConfigValueType -Definition $keyDefinition -Value $value
        if ($problem) {
            New-EobFinding -Severity Error -Code 'SETUP_INVALID' -Field $field.Name -Message "$($field.Label): $problem"
        }
    }
    $mode = if ($FieldValues.ContainsKey('ExchangeMode')) { [string]$FieldValues['ExchangeMode'] } else { 'None' }
    if ($mode -in @('OnPremises', 'Hybrid') -and -not ([string]$FieldValues['ExchangeUri']).Trim()) {
        New-EobFinding -Severity Error -Code 'SETUP_EXCHANGE_URI' -Field 'ExchangeUri' -Message "Exchange-Betriebsart $mode benötigt den Exchange-Server-Endpunkt."
    }
    if ($mode -eq 'Hybrid' -and -not ([string]$FieldValues['RoutingDomain']).Trim()) {
        New-EobFinding -Severity Warning -Code 'SETUP_ROUTING_DOMAIN' -Field 'RoutingDomain' -Message 'Hybrid ohne Routingdomäne: neue Remote-Postfächer können nicht angelegt werden.'
    }
    if ((ConvertTo-EobBoolean -Value $FieldValues['SyncEnabled']) -and -not ([string]$FieldValues['SyncServer']).Trim()) {
        New-EobFinding -Severity Error -Code 'SETUP_SYNC_SERVER' -Field 'SyncServer' -Message 'Die Synchronisation benötigt den Entra-Connect-Server.'
    }
}

function New-EobSetupConfiguration {
    <#
    .SYNOPSIS
        Schreibt die Konfiguration aus den Feldwerten und lädt sie zur Prüfung.
    .DESCRIPTION
        Eine vorhandene Datei wird nur mit -Force ersetzt; vorher wird eine Sicherung angelegt.
        Geschrieben wird atomar (temporäre Datei + Umbenennen) in UTF-8 mit BOM.
    .OUTPUTS
        Objekt mit Path, Backup und Config (Ergebnis von Import-EobConfiguration inkl. Findings).
    #>
    [CmdletBinding(SupportsShouldProcess)]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory)][hashtable]$FieldValues,
        [Parameter(Mandatory)][string]$Path,
        [switch]$RemoveSamples,
        [switch]$Force,
        [string]$TemplatePath = (Join-Path -Path (Join-Path -Path (Get-EobAppRoot) -ChildPath 'Config') -ChildPath 'easyONB.ini.template')
    )

    $errors = @(Test-EobSetupFieldValue -FieldValues $FieldValues | Where-Object Severity -EQ 'Error')
    if ($errors.Count -gt 0) { throw ('Die Eingaben sind unvollständig oder ungültig: ' + (($errors | ForEach-Object Message) -join ' | ')) }
    if (-not (Test-Path -LiteralPath $TemplatePath -PathType Leaf)) { throw "Vorlage nicht gefunden: $TemplatePath" }
    $exists = Test-Path -LiteralPath $Path -PathType Leaf
    if ($exists -and -not $Force) { throw "Die Konfiguration existiert bereits: $Path" }

    $ini = Get-EobSetupIniValue -FieldValues $FieldValues -RemoveSamples:$RemoveSamples
    $text = ConvertTo-EobSetupIniText -TemplateText (Get-EobTextFileContent -Path $TemplatePath).Text -Values $ini.Values -RemoveSection $ini.RemoveSection
    if (-not $PSCmdlet.ShouldProcess($Path, 'Konfiguration schreiben')) {
        return [pscustomobject]@{ Path = $Path; Backup = ''; Text = $text; Config = $null }
    }

    $directory = Split-Path -Parent $Path
    if ($directory -and -not (Test-Path -LiteralPath $directory -PathType Container)) { $null = New-Item -ItemType Directory -Path $directory -Force }
    $backup = ''
    if ($exists) {
        $backup = '{0}.{1}.bak' -f $Path, (Get-Date).ToString('yyyyMMdd-HHmmss')
        Copy-Item -LiteralPath $Path -Destination $backup -Force
    }
    $temporary = "$Path.$([guid]::NewGuid().ToString('N')).tmp"
    try {
        [System.IO.File]::WriteAllText($temporary, $text, [System.Text.UTF8Encoding]::new($true))
        Move-Item -LiteralPath $temporary -Destination $Path -Force
    }
    finally {
        if (Test-Path -LiteralPath $temporary) { Remove-Item -LiteralPath $temporary -Force -ErrorAction SilentlyContinue }
    }
    Write-EobLog -Level Audit -Action 'ConfigCreated' -Target $Path -Message "Konfiguration mit dem Installer erstellt$(if ($backup) { " (Sicherung: $backup)" })."
    [pscustomobject]@{ Path = $Path; Backup = $backup; Text = $text; Config = (Import-EobConfiguration -Path $Path) }
}

#endregion

#region Oberfläche

$script:SetupUi = $null

function Test-EobSetupGuiEnvironment {
    [CmdletBinding()]
    [OutputType([string])]
    param()

    if (-not $IsWindows) { return 'Der Installer benötigt Windows (WPF).' }
    if ([System.Threading.Thread]::CurrentThread.GetApartmentState() -ne [System.Threading.ApartmentState]::STA) {
        return 'Der Installer benötigt einen STA-Thread: pwsh -STA -File .\Install-easyONBOARDING.ps1'
    }
    return ''
}

function Import-EobSetupXaml {
    [CmdletBinding()]
    [OutputType([object])]
    param([Parameter(Mandatory)][string]$RelativePath)

    $path = Join-Path -Path (Join-Path -Path (Get-EobAppRoot) -ChildPath 'GUI') -ChildPath $RelativePath
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { throw "XAML-Datei fehlt: $path" }
    return [System.Windows.Markup.XamlReader]::Parse([System.IO.File]::ReadAllText($path, [System.Text.Encoding]::UTF8))
}

function Invoke-EobSetupSafely {
    <#
        Führt eine Aktion der Oberfläche aus und zeigt Fehler an, statt das Fenster zu beenden.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][scriptblock]$Action
    )

    $cursor = [System.Windows.Input.Mouse]::OverrideCursor
    try {
        [System.Windows.Input.Mouse]::OverrideCursor = [System.Windows.Input.Cursors]::Wait
        $null = & $Action
    }
    catch {
        $message = Protect-EobSensitiveText -Text $_.Exception.Message
        Write-EobLog -Level Error -Action 'Setup' -Target $Name -Message "Fehler bei '$Name': $message"
        [System.Windows.Input.Mouse]::OverrideCursor = $cursor
        if ($null -ne $script:SetupUi -and $script:SetupUi.Headless) { $script:SetupUi.DialogLog.Add("[Error] $Name`: $message"); return }
        $null = [System.Windows.MessageBox]::Show($script:SetupUi.Window, "$Name`n`n$message", 'easyONBOARDING Installer', 'OK', 'Error')
    }
    finally {
        [System.Windows.Input.Mouse]::OverrideCursor = $cursor
    }
}

function Set-EobSetupToolTip {
    <#
        Setzt einen Hinweistext mit kurzer Verzögerung (auch für deaktivierte Elemente).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][object]$Element,
        [Parameter(Mandatory)][string]$Text
    )

    $Element.ToolTip = $Text
    [System.Windows.Controls.ToolTipService]::SetInitialShowDelay($Element, 250)
    [System.Windows.Controls.ToolTipService]::SetShowDuration($Element, 30000)
    [System.Windows.Controls.ToolTipService]::SetShowOnDisabled($Element, $true)
}

function New-EobSetupFieldControl {
    [CmdletBinding()]
    [OutputType([object])]
    param([Parameter(Mandatory)][pscustomobject]$Field)

    switch ($Field.Kind) {
        'Check' {
            $control = [System.Windows.Controls.CheckBox]::new()
            $control.Margin = [System.Windows.Thickness]::new(0, 6, 0, 0)
            $control.VerticalAlignment = 'Center'
        }
        { $_ -in @('Combo', 'EditCombo') } {
            $control = [System.Windows.Controls.ComboBox]::new()
            # EditCombo: Auswahl aus dem AD oder freie Eingabe (z. B. auf einem Client).
            $control.IsEditable = ($Field.Kind -eq 'EditCombo')
            foreach ($option in $Field.Options) { $null = $control.Items.Add([string]$option) }
        }
        default {
            $control = [System.Windows.Controls.TextBox]::new()
        }
    }
    [System.Windows.Automation.AutomationProperties]::SetName($control, $Field.Label)
    Set-EobSetupToolTip -Element $control -Text $Field.Hint
    return $control
}

function Set-EobSetupFieldValue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Name,
        [AllowNull()][object]$Value
    )

    $control = $script:SetupUi.Fields[$Name]
    if ($null -eq $control) { return }
    if ($control -is [System.Windows.Controls.CheckBox]) { $control.IsChecked = ConvertTo-EobBoolean -Value $Value; return }
    if ($control -is [System.Windows.Controls.ComboBox]) {
        if ($control.IsEditable) { $control.Text = [string]$Value; return }
        $control.SelectedItem = @($control.Items | Where-Object { [string]$_ -eq [string]$Value }) | Select-Object -First 1
        return
    }
    $control.Text = [string]$Value
}

function Get-EobSetupFormValue {
    [CmdletBinding()]
    [OutputType([hashtable])]
    param()

    $values = @{}
    foreach ($name in $script:SetupUi.Fields.Keys) {
        $control = $script:SetupUi.Fields[$name]
        $values[$name] = if ($control -is [System.Windows.Controls.CheckBox]) { [bool]$control.IsChecked }
        elseif ($control -is [System.Windows.Controls.ComboBox]) { if ($control.IsEditable) { [string]$control.Text } else { [string]$control.SelectedItem } }
        else { [string]$control.Text }
    }
    return $values
}

function Add-EobSetupPageContent {
    <#
        Erzeugt die Eingabefelder einer Seite: eine Karte je Gruppe, zwei Felder je Zeile.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][object]$Panel,
        [Parameter(Mandatory)][string]$Page
    )

    $fields = @(Get-EobSetupField -Page $Page)
    foreach ($group in @($fields | ForEach-Object Group | Select-Object -Unique)) {
        $card = [System.Windows.Controls.Border]::new()
        $card.Style = $script:SetupUi.Application.FindResource('CardStyle')
        Set-EobSetupToolTip -Element $card -Text "Bereich $group. Zu jedem Feld erscheint ein Hinweis, wenn die Maus darauf zeigt."
        $stack = [System.Windows.Controls.StackPanel]::new()
        $heading = [System.Windows.Controls.TextBlock]::new()
        $heading.Text = $group
        $heading.Style = $script:SetupUi.Application.FindResource('SubheadingTextStyle')
        $null = $stack.Children.Add($heading)
        $grid = [System.Windows.Controls.Grid]::new()
        foreach ($width in @(190, -1, 24, 190, -1)) {
            $column = [System.Windows.Controls.ColumnDefinition]::new()
            $column.Width = if ($width -lt 0) { [System.Windows.GridLength]::new(1, 'Star') } else { [System.Windows.GridLength]::new($width) }
            $grid.ColumnDefinitions.Add($column)
        }
        $groupFields = @($fields | Where-Object Group -EQ $group)
        for ($index = 0; $index -lt $groupFields.Count; $index++) {
            $field = $groupFields[$index]
            $row = [Math]::Floor($index / 2)
            if ($index % 2 -eq 0) {
                $definition = [System.Windows.Controls.RowDefinition]::new()
                $definition.Height = [System.Windows.GridLength]::Auto
                $grid.RowDefinitions.Add($definition)
            }
            $column = if ($index % 2 -eq 0) { 0 } else { 3 }
            $label = [System.Windows.Controls.TextBlock]::new()
            $label.Text = if ($field.Required) { "$($field.Label) *" } else { $field.Label }
            $label.Style = $script:SetupUi.Application.FindResource('BodyTextStyle')
            $label.VerticalAlignment = 'Center'
            $label.Margin = [System.Windows.Thickness]::new(0, 6, 12, 0)
            Set-EobSetupToolTip -Element $label -Text $field.Hint
            [System.Windows.Controls.Grid]::SetRow($label, $row)
            [System.Windows.Controls.Grid]::SetColumn($label, $column)
            $control = New-EobSetupFieldControl -Field $field
            $control.Margin = [System.Windows.Thickness]::new(0, 6, 0, 0)
            [System.Windows.Controls.Grid]::SetRow($control, $row)
            [System.Windows.Controls.Grid]::SetColumn($control, $column + 1)
            $null = $grid.Children.Add($label)
            $null = $grid.Children.Add($control)
            $script:SetupUi.Fields[$field.Name] = $control
        }
        $null = $stack.Children.Add($grid)
        $card.Child = $stack
        $null = $Panel.Children.Add($card)
    }
}

function Set-EobSetupBanner {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][object]$Border,
        [Parameter(Mandatory)][object]$TextBlock,
        [ValidateSet('Info', 'Warning', 'Error', 'Success')][string]$Kind = 'Info',
        [AllowEmptyString()][string]$Text = ''
    )

    $Border.Style = $script:SetupUi.Application.FindResource("$($Kind)BannerStyle")
    $TextBlock.Text = $Text
    $Border.Visibility = if ($Text) { 'Visible' } else { 'Collapsed' }
}

function Update-EobSetupDirectoryValue {
    <#
        Belegt die Felder auf einem Domänencontroller aus dem Active Directory vor.
    #>
    [CmdletBinding()]
    param()

    $c = $script:SetupUi.C
    $defaults = Get-EobSetupDirectoryDefault
    foreach ($name in $defaults.Options.Keys) {
        $control = $script:SetupUi.Fields[$name]
        if ($control -isnot [System.Windows.Controls.ComboBox]) { continue }
        $control.Items.Clear()
        foreach ($option in @($defaults.Options[$name])) { $null = $control.Items.Add([string]$option) }
    }
    foreach ($name in $defaults.Values.Keys) { Set-EobSetupFieldValue -Name $name -Value $defaults.Values[$name] }
    $lines = @("Domänencontroller erkannt ($($defaults.Domain)): Felder aus dem Active Directory vorbelegt. Bitte prüfen und ergänzen.") + @($defaults.Notes)
    Set-EobSetupBanner -Border $c.SetupEnvironmentBanner -TextBlock $c.SetupEnvironmentText -Kind Success -Text ($lines -join "`n")
}

function Show-EobSetupFinding {
    [CmdletBinding()]
    param([AllowEmptyCollection()][object[]]$Finding = @())

    $list = $script:SetupUi.C.SetupFindingsList
    $list.Items.Clear()
    $labels = @{}
    foreach ($field in Get-EobSetupField) { $labels[$field.Name] = "$($script:SetupPages[$field.Page]) > $($field.Label)" }
    $severityText = @{ Error = 'Fehler'; Warning = 'Warnung'; Information = 'Hinweis' }
    foreach ($item in @($Finding | Sort-Object -Property @{ Expression = { @('Error', 'Warning', 'Information').IndexOf([string]$_.Severity) } })) {
        $where = if ($item.Field -and $labels.ContainsKey([string]$item.Field)) { $labels[[string]$item.Field] } else { [string]$item.Field }
        $null = $list.Items.Add(('{0}: {1}{2}' -f $severityText[[string]$item.Severity], $(if ($where) { "$where - " } else { '' }), $item.Message))
    }
}

function Invoke-EobSetupCheck {
    [CmdletBinding()]
    param()

    $c = $script:SetupUi.C
    $findings = @(Test-EobSetupFieldValue -FieldValues (Get-EobSetupFormValue))
    Show-EobSetupFinding -Finding $findings
    $errors = @($findings | Where-Object Severity -EQ 'Error').Count
    $c.SetupTabs.SelectedItem = $c.SetupCreateTab
    if ($errors -gt 0) {
        Set-EobSetupBanner -Border $c.SetupResultBanner -TextBlock $c.SetupResultText -Kind Error -Text "$errors Eingabe(n) fehlen oder sind ungültig. Details siehe Liste."
        return $false
    }
    Set-EobSetupBanner -Border $c.SetupResultBanner -TextBlock $c.SetupResultText -Kind Info -Text 'Eingaben vollständig. Die Konfiguration kann erstellt werden.'
    return $true
}

function Invoke-EobSetupCreate {
    [CmdletBinding()]
    param()

    $c = $script:SetupUi.C
    if (-not (Invoke-EobSetupCheck)) { return }
    $path = ([string]$c.SetupTargetText.Text).Trim()
    if (-not $path) { throw 'Bitte den Zielpfad der Konfiguration angeben.' }
    $path = [System.IO.Path]::GetFullPath($path)
    $force = $false
    if (Test-Path -LiteralPath $path -PathType Leaf) {
        if ($script:SetupUi.Headless) { $script:SetupUi.DialogLog.Add("[Question] Überschreiben: $path"); return }
        $answer = [System.Windows.MessageBox]::Show($script:SetupUi.Window, "Die Datei existiert bereits:`n$path`n`nErsetzen? Vorher wird eine Sicherung angelegt.", 'easyONBOARDING Installer', 'YesNo', 'Warning')
        if ($answer -ne 'Yes') { return }
        $force = $true
    }
    $result = New-EobSetupConfiguration -FieldValues (Get-EobSetupFormValue) -Path $path -RemoveSamples:([bool]$c.SetupRemoveSamplesCheck.IsChecked) -Force:$force -Confirm:$false
    $findings = @($result.Config.Findings | Where-Object { $_.Severity -in @('Error', 'Warning') })
    Show-EobSetupFinding -Finding $findings
    $errors = @($findings | Where-Object Severity -EQ 'Error').Count
    $text = "Konfiguration erstellt: $($result.Path)$(if ($result.Backup) { "`nSicherung: $($result.Backup)" })`nNächster Schritt: pwsh -File .\Start-easyONBOARDING.ps1 -CheckOnly"
    $kind = if ($errors -gt 0) { 'Warning' } else { 'Success' }
    if ($errors -gt 0) { $text += "`nDie Prüfung meldet $errors Fehler (siehe Liste)." }
    Set-EobSetupBanner -Border $c.SetupResultBanner -TextBlock $c.SetupResultText -Kind $kind -Text $text
}

function Initialize-EobSetupWindow {
    <#
    .SYNOPSIS
        Lädt Fenster und Felder, belegt Standardwerte und (auf einem Domänencontroller) AD-Werte vor.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$ConfigPath,
        [AllowNull()][pscustomobject]$Environment
    )

    $application = [System.Windows.Application]::Current
    if ($null -eq $application) {
        $application = [System.Windows.Application]::new()
        $application.ShutdownMode = [System.Windows.ShutdownMode]::OnExplicitShutdown
    }
    $application.Resources.MergedDictionaries.Clear()
    $application.Resources.MergedDictionaries.Add((Import-EobSetupXaml -RelativePath 'Styles/Theme.Light.xaml'))
    $application.Resources.MergedDictionaries.Add((Import-EobSetupXaml -RelativePath 'Styles/Controls.xaml'))
    $script:SetupUi.Application = $application

    $window = Import-EobSetupXaml -RelativePath 'Installer/SetupWindow.xaml'
    $script:SetupUi.Window = $window
    $c = @{}
    foreach ($name in @('SetupTitleText', 'SetupSubtitleText', 'SetupEnvironmentBanner', 'SetupEnvironmentText', 'SetupTabs', 'SetupCreateTab',
            'SetupTargetText', 'SetupBrowseButton', 'SetupRemoveSamplesCheck', 'SetupFindingsList', 'SetupResultBanner', 'SetupResultText',
            'SetupReloadButton', 'SetupCheckButton', 'SetupCreateButton', 'SetupCloseButton')) {
        $control = $window.FindName($name)
        if ($null -eq $control) { throw "Steuerelement fehlt: $name" }
        $c[$name] = $control
    }
    $script:SetupUi.C = $c
    $workArea = [System.Windows.SystemParameters]::WorkArea
    if ($workArea.Width -gt 0 -and $workArea.Height -gt 0) {
        $window.Width = [Math]::Min($window.Width, $workArea.Width)
        $window.Height = [Math]::Min($window.Height, $workArea.Height)
    }

    foreach ($page in Get-EobSetupPage) {
        $tab = [System.Windows.Controls.TabItem]::new()
        $tab.Header = $page.Title
        $scroll = [System.Windows.Controls.ScrollViewer]::new()
        $scroll.VerticalScrollBarVisibility = 'Auto'
        $scroll.HorizontalScrollBarVisibility = 'Disabled'
        $panel = [System.Windows.Controls.StackPanel]::new()
        $panel.Margin = [System.Windows.Thickness]::new(0, 4, 4, 8)
        Add-EobSetupPageContent -Panel $panel -Page $page.Id
        $scroll.Content = $panel
        $tab.Content = $scroll
        $c.SetupTabs.Items.Insert($c.SetupTabs.Items.Count - 1, $tab)
    }
    $c.SetupTabs.SelectedIndex = 0

    $defaults = Get-EobSetupDefaultValue
    foreach ($name in $defaults.Keys) { Set-EobSetupFieldValue -Name $name -Value $defaults[$name] }
    $c.SetupTargetText.Text = $ConfigPath

    $c.SetupBrowseButton.Add_Click({
            Invoke-EobSetupSafely -Name 'Zielpfad wählen' -Action {
                $dialog = [Microsoft.Win32.SaveFileDialog]::new()
                $dialog.Filter = 'INI-Dateien (*.ini)|*.ini'
                $dialog.FileName = Split-Path -Leaf $script:SetupUi.C.SetupTargetText.Text
                $folder = Split-Path -Parent $script:SetupUi.C.SetupTargetText.Text
                if ($folder -and (Test-Path -LiteralPath $folder)) { $dialog.InitialDirectory = $folder }
                if ($dialog.ShowDialog($script:SetupUi.Window)) { $script:SetupUi.C.SetupTargetText.Text = $dialog.FileName }
            }
        })
    $c.SetupCheckButton.Add_Click({ Invoke-EobSetupSafely -Name 'Eingaben prüfen' -Action { $null = Invoke-EobSetupCheck } })
    $c.SetupCreateButton.Add_Click({ Invoke-EobSetupSafely -Name 'Konfiguration erstellen' -Action { Invoke-EobSetupCreate } })
    $c.SetupReloadButton.Add_Click({ Invoke-EobSetupSafely -Name 'Werte aus dem AD laden' -Action { Update-EobSetupDirectoryValue } })
    $c.SetupCloseButton.Add_Click({ $script:SetupUi.Window.Close() })

    if ($null -eq $Environment) { $Environment = Get-EobSetupEnvironment }
    $script:SetupUi.Environment = $Environment
    $c.SetupSubtitleText.Text = "Erstellt die Konfiguration für easyONBOARDING  ·  $($Environment.Detail)"
    $canQuery = $Environment.IsDomainController -and $Environment.AdModuleAvailable
    $c.SetupReloadButton.IsEnabled = $canQuery
    if ($canQuery) {
        try {
            Update-EobSetupDirectoryValue
        }
        catch {
            Set-EobSetupBanner -Border $c.SetupEnvironmentBanner -TextBlock $c.SetupEnvironmentText -Kind Warning `
                -Text "Domänencontroller erkannt, AD-Abfrage fehlgeschlagen: $(Protect-EobSensitiveText -Text $_.Exception.Message). Bitte manuell ausfüllen."
        }
    }
    elseif ($Environment.IsDomainController) {
        Set-EobSetupBanner -Border $c.SetupEnvironmentBanner -TextBlock $c.SetupEnvironmentText -Kind Warning -Text 'Domänencontroller erkannt, aber das ActiveDirectory-Modul fehlt. Bitte manuell ausfüllen.'
    }
    else {
        Set-EobSetupBanner -Border $c.SetupEnvironmentBanner -TextBlock $c.SetupEnvironmentText -Kind Info `
            -Text 'Kein Domänencontroller: Bitte alle Felder manuell ausfüllen. Pflichtfelder sind mit * markiert.'
    }
}

function Start-EobSetupGui {
    <#
    .SYNOPSIS
        Startet den Installer (blockiert bis zum Schließen des Fensters).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$ConfigPath)

    $reason = Test-EobSetupGuiEnvironment
    if ($reason) { throw $reason }
    Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Xaml
    $script:SetupUi = @{ Application = $null; Window = $null; C = @{}; Fields = @{}; Environment = $null; Headless = $false; DialogLog = [System.Collections.Generic.List[string]]::new() }
    Initialize-EobSetupWindow -ConfigPath $ConfigPath
    $null = $script:SetupUi.Window.ShowDialog()
}

#endregion
