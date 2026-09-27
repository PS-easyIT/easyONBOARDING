#Requires -Version 7.2
<#
    easyONB.ActiveDirectory
    Adapter für das ActiveDirectory-Modul. Alle Aufrufe verwenden denselben Domänencontroller,
    Benutzereingaben werden LDAP-konform maskiert und alle schreibenden Funktionen unterstützen
    -WhatIf/-Confirm. Schreibende Funktionen prüfen Schutzregeln zusätzlich zur Planung
    (zweite Sicherheitslinie).
#>

Set-StrictMode -Version 3.0

$script:AdState = @{
    Server          = $null
    Domain          = $null
    Connected       = $false
    ModuleAvailable = $null
    LastError       = $null
    Config          = $null
    SearchBase      = $null
    MaxResults      = 50
}

$script:UserProperties = @(
    'DisplayName', 'GivenName', 'Surname', 'UserPrincipalName', 'mail', 'proxyAddresses', 'Title', 'Department',
    'Company', 'physicalDeliveryOfficeName', 'telephoneNumber', 'mobile', 'Description', 'employeeID', 'employeeNumber',
    'manager', 'directReports', 'memberOf', 'primaryGroupID', 'Enabled', 'AccountExpirationDate', 'LastLogonDate',
    'whenCreated', 'whenChanged', 'HomeDirectory', 'HomeDrive', 'ProfilePath', 'ScriptPath', 'logonHours',
    'adminCount', 'isCriticalSystemObject', 'ProtectedFromAccidentalDeletion'
)

# Zulässige New-ADUser-Parameter für zusätzliche Attribute (Schutz vor Parameter-Injection)
$script:AllowedNewUserParameters = @(
    'Title', 'Department', 'Company', 'Office', 'OfficePhone', 'MobilePhone', 'EmailAddress', 'EmployeeID',
    'EmployeeNumber', 'StreetAddress', 'City', 'PostalCode', 'Country', 'State', 'HomeDirectory', 'HomeDrive',
    'ProfilePath', 'ScriptPath', 'AccountExpirationDate', 'Description', 'Initials', 'HomePage',
    'PasswordNeverExpires', 'SmartcardLogonRequired', 'Manager'
)

# Zulässige LDAP-Attribute für -Replace/-Clear (Schutz vor Änderungen sicherheitsrelevanter Attribute)
$script:AllowedLdapAttributes = @(
    'displayName', 'givenName', 'sn', 'initials', 'title', 'department', 'company', 'physicalDeliveryOfficeName',
    'telephoneNumber', 'mobile', 'facsimileTelephoneNumber', 'pager', 'ipPhone', 'mail', 'proxyAddresses',
    'description', 'employeeID', 'employeeNumber', 'employeeType', 'streetAddress', 'l', 'postalCode', 'c', 'co',
    'countryCode', 'st', 'wWWHomePage', 'homeDirectory', 'homeDrive', 'profilePath', 'scriptPath', 'logonHours',
    'thumbnailPhoto', 'manager', 'division', 'extensionAttribute1', 'extensionAttribute2', 'extensionAttribute3',
    'extensionAttribute4', 'extensionAttribute5', 'extensionAttribute6', 'extensionAttribute7', 'extensionAttribute8',
    'extensionAttribute9', 'extensionAttribute10', 'extensionAttribute11', 'extensionAttribute12',
    'extensionAttribute13', 'extensionAttribute14', 'extensionAttribute15', 'msExchHideFromAddressLists',
    'info', 'otherTelephone', 'otherMobile'
)

#region Verbindung und Status

function Get-EobAdServerParameter {
    [CmdletBinding()]
    [OutputType([hashtable])]
    param()

    if ($script:AdState.Server) {
        return @{ Server = $script:AdState.Server }
    }
    return @{}
}

function Test-EobAdModuleAvailable {
    <#
    .SYNOPSIS
        Prüft, ob das ActiveDirectory-Modul verfügbar ist (lädt es bei Bedarf einmalig).
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param()

    if ($null -ne (Get-Command -Name 'Get-ADUser' -ErrorAction SilentlyContinue)) {
        $script:AdState.ModuleAvailable = $true
        return $true
    }
    if ($script:AdState.ModuleAvailable -eq $false) {
        return $false
    }
    try {
        Import-Module -Name ActiveDirectory -ErrorAction Stop -WarningAction SilentlyContinue
        $script:AdState.ModuleAvailable = $true
    }
    catch {
        $script:AdState.ModuleAvailable = $false
        $script:AdState.LastError = 'ActiveDirectory-Modul nicht verfügbar.'
    }
    return [bool]$script:AdState.ModuleAvailable
}

function Initialize-EobAdConnection {
    <#
    .SYNOPSIS
        Stellt die AD-Verbindung her und legt den Domänencontroller für alle Operationen fest.
    .DESCRIPTION
        Reihenfolge der DC-Auswahl: -Server, [ActiveDirectory] PreferredDomainController,
        [RemoteExecution] DefaultDCServerName, automatisch (beschreibbarer DC). Wirft nicht,
        sondern liefert ein Statusobjekt.
    .PARAMETER Config
        Konfiguration.
    .PARAMETER Server
        Expliziter Domänencontroller.
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [AllowNull()][pscustomobject]$Config,
        [string]$Server
    )

    $script:AdState.Config = $Config
    $script:AdState.Connected = $false
    $script:AdState.Domain = $null
    $script:AdState.LastError = $null
    $script:AdState.SearchBase = Get-EobConfigValue -Config $Config -Section 'ActiveDirectory' -Key 'SearchBase'
    $script:AdState.MaxResults = Get-EobConfigValue -Config $Config -Section 'ActiveDirectory' -Key 'MaxSearchResults' -As Int

    if (-not (Test-EobAdModuleAvailable)) {
        return Get-EobAdConnectionInfo
    }

    try {
        $target = $Server
        if (-not $target) { $target = Get-EobConfigValue -Config $Config -Section 'ActiveDirectory' -Key 'PreferredDomainController' }
        if (-not $target) { $target = Get-EobConfigValue -Config $Config -Section 'RemoteExecution' -Key 'DefaultDCServerName' }
        if (-not $target) {
            $dc = Get-ADDomainController -Discover -Writable -ErrorAction Stop
            $target = @($dc.HostName)[0]
        }
        $domain = Get-ADDomain -Server $target -ErrorAction Stop
        $script:AdState.Server = [string]$target
        $script:AdState.Domain = [pscustomobject]@{
            DnsRoot           = [string]$domain.DNSRoot
            NetBiosName       = [string]$domain.NetBIOSName
            DistinguishedName = [string]$domain.DistinguishedName
            DomainSid         = [string](Get-EobPropertyValue -InputObject $domain -Name 'DomainSID' -Default '')
            PdcEmulator       = [string](Get-EobPropertyValue -InputObject $domain -Name 'PDCEmulator' -Default '')
        }
        $script:AdState.Connected = $true
        Write-EobLog -Level Information -Action 'AdConnect' -Target $script:AdState.Server -Result 'Succeeded' -Message "Verbunden mit $($script:AdState.Domain.DnsRoot) über $($script:AdState.Server)."
    }
    catch {
        $script:AdState.LastError = ConvertTo-EobAdErrorMessage -ErrorRecord $_
        Write-EobLog -Level Warning -Action 'AdConnect' -Result 'Failed' -Message "AD-Verbindung fehlgeschlagen: $($script:AdState.LastError)"
    }
    return Get-EobAdConnectionInfo
}

function Get-EobAdConnectionInfo {
    <#
    .SYNOPSIS
        Liefert den aktuellen Verbindungszustand.
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param()

    [pscustomobject]@{
        PSTypeName      = 'Eob.AdConnection'
        ModuleAvailable = [bool]$script:AdState.ModuleAvailable
        Connected       = [bool]$script:AdState.Connected
        Server          = $script:AdState.Server
        Domain          = $script:AdState.Domain
        LastError       = $script:AdState.LastError
    }
}

function Get-EobAdStatus {
    <#
    .SYNOPSIS
        Integrationsstatus Active Directory (Dashboard).
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param([AllowNull()][object]$Config)

    $hint = 'Windows 10/11: Add-WindowsCapability -Online -Name Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0; Windows Server: Install-WindowsFeature RSAT-AD-PowerShell'
    if ($null -eq $script:AdState.ModuleAvailable) {
        $null = Test-EobAdModuleAvailable
    }
    if (-not $script:AdState.ModuleAvailable) {
        return New-EobIntegrationStatus -Name 'Active Directory' -State NotInstalled -Detail 'ActiveDirectory-Modul nicht gefunden.' -Hint $hint
    }
    if ($script:AdState.Connected) {
        $version = [string](Get-Module -Name ActiveDirectory | Select-Object -First 1 -ExpandProperty Version -ErrorAction SilentlyContinue)
        return New-EobIntegrationStatus -Name 'Active Directory' -State Connected -Version $version -Detail "$($script:AdState.Domain.DnsRoot) über $($script:AdState.Server)"
    }
    $detail = if ($script:AdState.LastError) { $script:AdState.LastError } else { 'Nicht verbunden.' }
    return New-EobIntegrationStatus -Name 'Active Directory' -State NotConnected -Detail $detail -Hint 'Netzwerk, DNS und Active Directory Web Services (Port 9389) prüfen.'
}

function Assert-EobAdConnected {
    [CmdletBinding()]
    param()

    if (-not $script:AdState.Connected) {
        $reason = if ($script:AdState.LastError) { " ($($script:AdState.LastError))" } else { '' }
        throw "Keine Active-Directory-Verbindung$reason."
    }
}

function ConvertTo-EobAdErrorMessage {
    <#
    .SYNOPSIS
        Übersetzt häufige AD-Fehler in verständliche Meldungen (redigiert).
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param([Parameter(Mandatory)][System.Management.Automation.ErrorRecord]$ErrorRecord)

    $exception = $ErrorRecord.Exception
    $typeName = $exception.GetType().Name
    $message = Protect-EobSensitiveText -Text $exception.Message
    $prefix = switch -Regex ($typeName) {
        'ADIdentityNotFoundException' { 'Objekt nicht gefunden'; break }
        'ADServerDownException' { 'Domänencontroller nicht erreichbar'; break }
        'UnauthorizedAccessException' { 'Zugriff verweigert (fehlende AD-Berechtigung)'; break }
        'ADPasswordComplexityException' { 'Kennwort erfüllt die Domänenrichtlinie nicht'; break }
        'ADIdentityAlreadyExistsException' { 'Objekt existiert bereits'; break }
        default { $null }
    }
    if (-not $prefix -and $message -match '(?i)access is denied|zugriff verweigert') {
        $prefix = 'Zugriff verweigert (fehlende AD-Berechtigung)'
    }
    if ($prefix) { return "${prefix}: $message" }
    return $message
}

#endregion

#region Lesen

function ConvertTo-EobAdUserInfo {
    <#
    .SYNOPSIS
        Normalisiert ein AD-Benutzerobjekt (serialisierbar, StrictMode-sicher).
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param([Parameter(Mandatory)][object]$AdUser)

    $get = { param($name, $default = $null) Get-EobPropertyValue -InputObject $AdUser -Name $name -Default $default }
    $strings = { param($name) @(& $get $name @() | ForEach-Object { [string]$_ } | Where-Object { $_ -ne '' }) }
    $dn = [string](& $get 'DistinguishedName' '')
    $parent = if ($dn -match '^(?:[^,\\]|\\.)+,(.+)$') { $Matches[1] } else { '' }
    $guid = & $get 'ObjectGUID' ''
    $sid = & $get 'SID' (& $get 'objectSid' '')
    $logonHours = & $get 'logonHours' $null
    $restricted = $false
    if ($null -ne $logonHours) {
        $bytes = [byte[]]@($logonHours)
        $restricted = ($bytes.Length -eq 21 -and @($bytes | Where-Object { $_ -ne 255 }).Count -gt 0)
    }
    [pscustomobject]@{
        PSTypeName                      = 'Eob.AdUser'
        SamAccountName                  = [string](& $get 'SamAccountName' '')
        UserPrincipalName               = [string](& $get 'UserPrincipalName' '')
        DisplayName                     = [string](& $get 'DisplayName' '')
        GivenName                       = [string](& $get 'GivenName' '')
        Surname                         = [string](& $get 'Surname' '')
        Mail                            = [string](& $get 'mail' (& $get 'EmailAddress' ''))
        ProxyAddresses                  = & $strings 'proxyAddresses'
        Title                           = [string](& $get 'Title' '')
        Department                      = [string](& $get 'Department' '')
        Company                         = [string](& $get 'Company' '')
        Office                          = [string](& $get 'physicalDeliveryOfficeName' (& $get 'Office' ''))
        OfficePhone                     = [string](& $get 'telephoneNumber' (& $get 'OfficePhone' ''))
        MobilePhone                     = [string](& $get 'mobile' (& $get 'MobilePhone' ''))
        Description                     = [string](& $get 'Description' '')
        EmployeeId                      = [string](& $get 'employeeID' '')
        EmployeeNumber                  = [string](& $get 'employeeNumber' '')
        Manager                         = [string](& $get 'manager' '')
        DirectReports                   = & $strings 'directReports'
        MemberOf                        = & $strings 'memberOf'
        PrimaryGroupId                  = [int](& $get 'primaryGroupID' 513)
        Enabled                         = [bool](& $get 'Enabled' $false)
        AccountExpirationDate           = & $get 'AccountExpirationDate' $null
        LastLogonDate                   = & $get 'LastLogonDate' $null
        WhenCreated                     = & $get 'whenCreated' $null
        WhenChanged                     = & $get 'whenChanged' $null
        HomeDirectory                   = [string](& $get 'HomeDirectory' '')
        HomeDrive                       = [string](& $get 'HomeDrive' '')
        ProfilePath                     = [string](& $get 'ProfilePath' '')
        ScriptPath                      = [string](& $get 'ScriptPath' '')
        LogonHoursRestricted            = $restricted
        AdminCount                      = [int](& $get 'adminCount' 0)
        ObjectClass                     = @(& $get 'ObjectClass' 'user') | ForEach-Object { [string]$_ } | Select-Object -Last 1
        ObjectGuid                      = [string]$guid
        Sid                             = [string]$sid
        DistinguishedName               = $dn
        ParentContainer                 = $parent
        IsCriticalSystemObject          = [bool](& $get 'isCriticalSystemObject' $false)
        ProtectedFromAccidentalDeletion = [bool](& $get 'ProtectedFromAccidentalDeletion' $false)
    }
}

function Find-EobAdUser {
    <#
    .SYNOPSIS
        Sucht Benutzer nach Name, SamAccountName, UPN, E-Mail oder Personalnummer.
    .PARAMETER SearchText
        Suchbegriff (mindestens 2 Zeichen). Wird LDAP-konform maskiert.
    .PARAMETER MaxResults
        Maximale Trefferzahl.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$SearchText,
        [ValidateRange(0, 500)][int]$MaxResults = 0
    )

    Assert-EobAdConnected
    $text = $SearchText.Trim()
    if ($text.Length -lt 2) {
        throw 'Der Suchbegriff muss mindestens 2 Zeichen lang sein.'
    }
    $value = ConvertTo-EobLdapFilterValue -Value $text
    $filter = "(&(objectCategory=person)(objectClass=user)(|(sAMAccountName=*$value*)(userPrincipalName=*$value*)(mail=*$value*)(displayName=*$value*)(givenName=*$value*)(sn=*$value*)(employeeID=$value)))"
    $limit = if ($MaxResults -gt 0) { $MaxResults } else { [Math]::Max(1, [int]$script:AdState.MaxResults) }
    $parameters = Get-EobAdServerParameter
    $parameters['LDAPFilter'] = $filter
    $parameters['Properties'] = @('DisplayName', 'mail', 'Department', 'Title', 'Enabled', 'UserPrincipalName', 'employeeID', 'adminCount')
    $parameters['ResultSetSize'] = $limit
    if ($script:AdState.SearchBase) { $parameters['SearchBase'] = $script:AdState.SearchBase }

    foreach ($user in @(Get-ADUser @parameters -ErrorAction Stop)) {
        ConvertTo-EobAdUserInfo -AdUser $user
    }
}

function Resolve-EobAdUser {
    <#
    .SYNOPSIS
        Löst eine Benutzerangabe eindeutig auf (SamAccountName, UPN, E-Mail, DN, GUID oder SID).
    .DESCRIPTION
        Kein oder mehrdeutiger Treffer führt zu einer Exception mit verständlicher Meldung.
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param([Parameter(Mandatory)][string]$Identity)

    Assert-EobAdConnected
    $value = $Identity.Trim()
    if ($value -eq '') { throw 'Leere Benutzerangabe.' }
    $parameters = Get-EobAdServerParameter
    $parameters['Properties'] = $script:UserProperties

    $parsedGuid = [guid]::Empty
    $isGuid = [guid]::TryParse($value, [ref]$parsedGuid)
    if ($value -match '^S-1-' -or $isGuid -or (Test-EobDistinguishedName -Value $value)) {
        try {
            return ConvertTo-EobAdUserInfo -AdUser (Get-ADUser -Identity $value @parameters -ErrorAction Stop)
        }
        catch {
            throw "Benutzer '$value' nicht gefunden: $(ConvertTo-EobAdErrorMessage -ErrorRecord $_)"
        }
    }

    $escaped = ConvertTo-EobLdapFilterValue -Value $value
    $filter = if ($value.Contains('@')) {
        "(&(objectCategory=person)(objectClass=user)(|(userPrincipalName=$escaped)(mail=$escaped)(proxyAddresses=smtp:$escaped)))"
    }
    else {
        "(&(objectCategory=person)(objectClass=user)(sAMAccountName=$escaped))"
    }
    $found = @(Get-ADUser -LDAPFilter $filter -ResultSetSize 2 @parameters -ErrorAction Stop)
    if ($found.Count -eq 0) { throw "Benutzer '$value' wurde nicht gefunden." }
    if ($found.Count -gt 1) { throw "Die Angabe '$value' ist nicht eindeutig ($($found.Count) Treffer)." }
    return ConvertTo-EobAdUserInfo -AdUser $found[0]
}

function Get-EobAdUser {
    <#
    .SYNOPSIS
        Liest alle für easyONBOARDING relevanten Eigenschaften eines Benutzers.
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param([Parameter(Mandatory)][string]$Identity)

    return Resolve-EobAdUser -Identity $Identity
}

function Get-EobAdIdentityConflict {
    <#
    .SYNOPSIS
        Prüft, ob SamAccountName, UPN oder E-Mail-Adresse bereits von einem AD-Objekt belegt sind.
    .DESCRIPTION
        Durchsucht alle Objektklassen (Benutzer, Gruppen, Kontakte, Computer) inklusive
        proxyAddresses. Liefert eine Liste der Konflikte (leer = frei).
    #>
    [CmdletBinding()]
    param(
        [string]$SamAccountName,
        [string]$UserPrincipalName,
        [string]$Mail
    )

    Assert-EobAdConnected
    $clauses = [System.Collections.Generic.List[string]]::new()
    if ($SamAccountName) { $clauses.Add("(sAMAccountName=$(ConvertTo-EobLdapFilterValue -Value $SamAccountName))") }
    if ($UserPrincipalName) { $clauses.Add("(userPrincipalName=$(ConvertTo-EobLdapFilterValue -Value $UserPrincipalName))") }
    if ($Mail) {
        $escapedMail = ConvertTo-EobLdapFilterValue -Value $Mail
        $clauses.Add("(mail=$escapedMail)")
        $clauses.Add("(proxyAddresses=smtp:$escapedMail)")
    }
    if ($clauses.Count -eq 0) { return }

    $parameters = Get-EobAdServerParameter
    $objects = @(Get-ADObject -LDAPFilter ("(|{0})" -f ($clauses -join '')) -Properties sAMAccountName, userPrincipalName, mail, proxyAddresses @parameters -ErrorAction Stop)
    foreach ($object in $objects) {
        $dn = [string](Get-EobPropertyValue -InputObject $object -Name 'DistinguishedName' -Default '')
        $objSam = [string](Get-EobPropertyValue -InputObject $object -Name 'sAMAccountName' -Default '')
        $objUpn = [string](Get-EobPropertyValue -InputObject $object -Name 'userPrincipalName' -Default '')
        $objMail = [string](Get-EobPropertyValue -InputObject $object -Name 'mail' -Default '')
        $objProxies = @(Get-EobPropertyValue -InputObject $object -Name 'proxyAddresses' -Default @()) | ForEach-Object { ([string]$_) -replace '^(?i)smtp:', '' }
        if ($SamAccountName -and $objSam -ieq $SamAccountName) {
            [pscustomobject]@{ Attribute = 'SamAccountName'; Value = $SamAccountName; ConflictingObject = $dn }
        }
        if ($UserPrincipalName -and $objUpn -ieq $UserPrincipalName) {
            [pscustomobject]@{ Attribute = 'UserPrincipalName'; Value = $UserPrincipalName; ConflictingObject = $dn }
        }
        if ($Mail -and ($objMail -ieq $Mail -or @($objProxies | Where-Object { $_ -ieq $Mail }).Count -gt 0)) {
            [pscustomobject]@{ Attribute = 'Mail'; Value = $Mail; ConflictingObject = $dn }
        }
    }
}

function Test-EobAdNameInContainer {
    <#
    .SYNOPSIS
        Prüft, ob ein CN (Objektname) im Zielcontainer bereits vergeben ist.
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param(
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][string]$Path
    )

    Assert-EobAdConnected
    $parameters = Get-EobAdServerParameter
    $found = @(Get-ADObject -LDAPFilter "(cn=$(ConvertTo-EobLdapFilterValue -Value $Name))" -SearchBase $Path -SearchScope OneLevel @parameters -ErrorAction Stop)
    return ($found.Count -gt 0)
}

function Test-EobAdOrganizationalUnit {
    <#
    .SYNOPSIS
        Prüft, ob ein Zielcontainer (OU oder Container) existiert.
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param([Parameter(Mandatory)][string]$DistinguishedName)

    Assert-EobAdConnected
    if (-not (Test-EobDistinguishedName -Value $DistinguishedName)) { return $false }
    try {
        $server = Get-EobAdServerParameter
        $object = Get-ADObject -Identity $DistinguishedName -Properties objectClass @server -ErrorAction Stop
        $classes = @(Get-EobPropertyValue -InputObject $object -Name 'ObjectClass' -Default @()) | ForEach-Object { [string]$_ }
        return (@($classes | Where-Object { $_ -in @('organizationalUnit', 'container', 'builtinDomain') }).Count -gt 0)
    }
    catch {
        return $false
    }
}

function Get-EobAdOrganizationalUnitList {
    <#
    .SYNOPSIS
        Liefert Organisationseinheiten für die Auswahl in der Oberfläche.
    #>
    [CmdletBinding()]
    param([string]$SearchBase)

    Assert-EobAdConnected
    $parameters = Get-EobAdServerParameter
    $parameters['Filter'] = '*'
    $parameters['Properties'] = @('CanonicalName')
    $parameters['ResultSetSize'] = 5000
    if ($SearchBase) { $parameters['SearchBase'] = $SearchBase }
    foreach ($ou in @(Get-ADOrganizationalUnit @parameters -ErrorAction Stop | Sort-Object -Property CanonicalName)) {
        $canonical = [string](Get-EobPropertyValue -InputObject $ou -Name 'CanonicalName' -Default '')
        [pscustomobject]@{
            Name              = [string](Get-EobPropertyValue -InputObject $ou -Name 'Name' -Default '')
            DistinguishedName = [string](Get-EobPropertyValue -InputObject $ou -Name 'DistinguishedName' -Default '')
            CanonicalName     = $canonical
            Depth             = [Math]::Max(0, @($canonical -split '/').Count - 2)
        }
    }
}

function Get-EobAdGroup {
    <#
    .SYNOPSIS
        Liest eine Gruppe (Name, DN, SID, adminCount, Kategorie).
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param([Parameter(Mandatory)][string]$Identity)

    Assert-EobAdConnected
    $parameters = Get-EobAdServerParameter
    $parameters['Properties'] = @('adminCount', 'Description', 'GroupCategory', 'GroupScope', 'mail')
    try {
        $group = Get-ADGroup -Identity $Identity @parameters -ErrorAction Stop
    }
    catch {
        $escaped = ConvertTo-EobLdapFilterValue -Value $Identity
        $candidates = @(Get-ADGroup -LDAPFilter "(|(cn=$escaped)(name=$escaped)(displayName=$escaped))" -ResultSetSize 2 @parameters -ErrorAction Stop)
        if ($candidates.Count -ne 1) {
            throw "Gruppe '$Identity' wurde nicht eindeutig gefunden."
        }
        $group = $candidates[0]
    }
    [pscustomobject]@{
        PSTypeName        = 'Eob.AdGroup'
        Name              = [string](Get-EobPropertyValue -InputObject $group -Name 'Name' -Default '')
        SamAccountName    = [string](Get-EobPropertyValue -InputObject $group -Name 'SamAccountName' -Default '')
        DistinguishedName = [string](Get-EobPropertyValue -InputObject $group -Name 'DistinguishedName' -Default '')
        SID               = [string](Get-EobPropertyValue -InputObject $group -Name 'SID' -Default '')
        AdminCount        = [int](Get-EobPropertyValue -InputObject $group -Name 'adminCount' -Default 0)
        GroupCategory     = [string](Get-EobPropertyValue -InputObject $group -Name 'GroupCategory' -Default '')
        GroupScope        = [string](Get-EobPropertyValue -InputObject $group -Name 'GroupScope' -Default '')
        Description       = [string](Get-EobPropertyValue -InputObject $group -Name 'Description' -Default '')
    }
}

function Get-EobAdTransitiveGroup {
    <#
    .SYNOPSIS
        Liefert alle Gruppen, in denen ein Objekt (Benutzer oder Gruppe) transitiv Mitglied ist.
    .DESCRIPTION
        Verwendet LDAP_MATCHING_RULE_IN_CHAIN (1.2.840.113556.1.4.1941) - eine Abfrage für die gesamte Kette.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$DistinguishedName)

    Assert-EobAdConnected
    $parameters = Get-EobAdServerParameter
    $filter = "(&(objectClass=group)(member:1.2.840.113556.1.4.1941:=$(ConvertTo-EobLdapFilterValue -Value $DistinguishedName)))"
    foreach ($group in @(Get-ADGroup -LDAPFilter $filter -Properties adminCount @parameters -ErrorAction Stop)) {
        [pscustomobject]@{
            PSTypeName        = 'Eob.AdGroup'
            Name              = [string](Get-EobPropertyValue -InputObject $group -Name 'Name' -Default '')
            SamAccountName    = [string](Get-EobPropertyValue -InputObject $group -Name 'SamAccountName' -Default '')
            DistinguishedName = [string](Get-EobPropertyValue -InputObject $group -Name 'DistinguishedName' -Default '')
            SID               = [string](Get-EobPropertyValue -InputObject $group -Name 'SID' -Default '')
            AdminCount        = [int](Get-EobPropertyValue -InputObject $group -Name 'adminCount' -Default 0)
        }
    }
}

function Get-EobAdPrivilegedMembership {
    <#
    .SYNOPSIS
        Ermittelt privilegierte Gruppen, in denen ein Objekt transitiv Mitglied ist.
    .PARAMETER DistinguishedName
        DN des Benutzers bzw. der Gruppe.
    .PARAMETER PrimaryGroupId
        Optional: primäre Gruppe (RID) eines Benutzers.
    #>
    [CmdletBinding()]
    [OutputType([string[]])]
    param(
        [Parameter(Mandatory)][string]$DistinguishedName,
        [int]$PrimaryGroupId = 0
    )

    $result = [System.Collections.Generic.List[string]]::new()
    foreach ($group in @(Get-EobAdTransitiveGroup -DistinguishedName $DistinguishedName)) {
        if ((Test-EobPrivilegedGroup -Group $group -Config $script:AdState.Config).IsProtected) {
            $result.Add($group.Name)
        }
    }
    if ($PrimaryGroupId -gt 0 -and $PrimaryGroupId -ne 513 -and $null -ne $script:AdState.Domain -and $script:AdState.Domain.DomainSid) {
        $primarySid = '{0}-{1}' -f $script:AdState.Domain.DomainSid, $PrimaryGroupId
        if ((Test-EobPrivilegedGroup -Group $primarySid -Config $script:AdState.Config).IsProtected) {
            $result.Add("Primäre Gruppe (RID $PrimaryGroupId)")
        }
    }
    return @($result | Select-Object -Unique)
}

function Get-EobAdGroupProtection {
    <#
    .SYNOPSIS
        Prüft eine Gruppe inklusive verschachtelter Mitgliedschaft in privilegierten Gruppen.
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param([Parameter(Mandatory)][string]$Identity)

    $group = Get-EobAdGroup -Identity $Identity
    $nested = @(Get-EobAdPrivilegedMembership -DistinguishedName $group.DistinguishedName)
    $protection = Test-EobPrivilegedGroup -Group $group -Config $script:AdState.Config -NestedIn $nested
    $protection | Add-Member -NotePropertyName 'GroupObject' -NotePropertyValue $group -Force
    return $protection
}

function Get-EobAdUserGroupMembership {
    <#
    .SYNOPSIS
        Liefert die direkten Gruppenmitgliedschaften eines Benutzers inkl. primärer Gruppe.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][pscustomobject]$User)

    Assert-EobAdConnected
    $parameters = Get-EobAdServerParameter
    foreach ($groupDn in @($User.MemberOf)) {
        try {
            $group = Get-ADGroup -Identity $groupDn -Properties adminCount, GroupCategory @parameters -ErrorAction Stop
            [pscustomobject]@{
                Name              = [string](Get-EobPropertyValue -InputObject $group -Name 'Name' -Default '')
                SamAccountName    = [string](Get-EobPropertyValue -InputObject $group -Name 'SamAccountName' -Default '')
                DistinguishedName = [string]$groupDn
                SID               = [string](Get-EobPropertyValue -InputObject $group -Name 'SID' -Default '')
                AdminCount        = [int](Get-EobPropertyValue -InputObject $group -Name 'adminCount' -Default 0)
                GroupCategory     = [string](Get-EobPropertyValue -InputObject $group -Name 'GroupCategory' -Default '')
                IsPrimary         = $false
            }
        }
        catch {
            [pscustomobject]@{ Name = $groupDn; SamAccountName = ''; DistinguishedName = [string]$groupDn; SID = ''; AdminCount = 0; GroupCategory = ''; IsPrimary = $false }
        }
    }
    if ($null -ne $script:AdState.Domain -and $script:AdState.Domain.DomainSid -and $User.PrimaryGroupId -gt 0) {
        $primarySid = '{0}-{1}' -f $script:AdState.Domain.DomainSid, $User.PrimaryGroupId
        try {
            $primary = Get-ADGroup -Identity $primarySid @parameters -ErrorAction Stop
            [pscustomobject]@{
                Name              = [string](Get-EobPropertyValue -InputObject $primary -Name 'Name' -Default '')
                SamAccountName    = [string](Get-EobPropertyValue -InputObject $primary -Name 'SamAccountName' -Default '')
                DistinguishedName = [string](Get-EobPropertyValue -InputObject $primary -Name 'DistinguishedName' -Default '')
                SID               = $primarySid
                AdminCount        = 0
                GroupCategory     = 'Security'
                IsPrimary         = $true
            }
        }
        catch {
            Write-EobLog -Level Debug -Message "Primäre Gruppe $primarySid konnte nicht gelesen werden."
        }
    }
}

function Resolve-EobAdUserReference {
    <#
    .SYNOPSIS
        Löst eine Führungskraft bzw. einen Referenzbenutzer eindeutig auf.
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param([Parameter(Mandatory)][string]$Identity)

    return Resolve-EobAdUser -Identity $Identity
}

function Get-EobAdDirectReport {
    <#
    .SYNOPSIS
        Liefert die direkt unterstellten Mitarbeitenden (max. 200).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][pscustomobject]$User)

    Assert-EobAdConnected
    $parameters = Get-EobAdServerParameter
    foreach ($dn in @($User.DirectReports | Select-Object -First 200)) {
        try {
            $report = Get-ADUser -Identity $dn -Properties DisplayName @parameters -ErrorAction Stop
            [pscustomobject]@{
                SamAccountName    = [string](Get-EobPropertyValue -InputObject $report -Name 'SamAccountName' -Default '')
                DisplayName       = [string](Get-EobPropertyValue -InputObject $report -Name 'DisplayName' -Default '')
                DistinguishedName = [string]$dn
            }
        }
        catch {
            [pscustomobject]@{ SamAccountName = ''; DisplayName = ''; DistinguishedName = [string]$dn }
        }
    }
}

function Get-EobAdPasswordPolicy {
    <#
    .SYNOPSIS
        Liest die Standard-Kennwortrichtlinie der Domäne.
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param()

    Assert-EobAdConnected
    $server = Get-EobAdServerParameter
    $policy = Get-ADDefaultDomainPasswordPolicy @server -ErrorAction Stop
    [pscustomobject]@{
        MinPasswordLength = [int](Get-EobPropertyValue -InputObject $policy -Name 'MinPasswordLength' -Default 0)
        ComplexityEnabled = [bool](Get-EobPropertyValue -InputObject $policy -Name 'ComplexityEnabled' -Default $true)
    }
}

function Test-EobAdRecycleBinEnabled {
    <#
    .SYNOPSIS
        Prüft, ob der AD-Papierkorb aktiviert ist.
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param()

    Assert-EobAdConnected
    try {
        $server = Get-EobAdServerParameter
        $feature = Get-ADOptionalFeature -Filter "Name -eq 'Recycle Bin Feature'" @server -ErrorAction Stop
        return (@(Get-EobPropertyValue -InputObject $feature -Name 'EnabledScopes' -Default @()).Count -gt 0)
    }
    catch {
        return $false
    }
}

function Get-EobAdUserSnapshot {
    <#
    .SYNOPSIS
        Erstellt einen serialisierbaren Zustandsbericht eines Benutzers (Attribute, Gruppen, Führungskraft, Unterstellte).
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param([Parameter(Mandatory)][string]$Identity)

    $user = Get-EobAdUser -Identity $Identity
    $groups = @(Get-EobAdUserGroupMembership -User $user)
    $manager = $null
    if ($user.Manager) {
        try {
            $managerUser = Resolve-EobAdUser -Identity $user.Manager
            $manager = [pscustomobject]@{ SamAccountName = $managerUser.SamAccountName; DisplayName = $managerUser.DisplayName; Mail = $managerUser.Mail; DistinguishedName = $managerUser.DistinguishedName }
        }
        catch {
            $manager = [pscustomobject]@{ SamAccountName = ''; DisplayName = ''; Mail = ''; DistinguishedName = $user.Manager }
        }
    }
    [pscustomobject]@{
        PSTypeName    = 'Eob.AdUserSnapshot'
        CapturedAt    = (Get-Date).ToString('o')
        Server        = $script:AdState.Server
        User          = $user
        Groups        = $groups
        Manager       = $manager
        DirectReports = @(Get-EobAdDirectReport -User $user)
    }
}

#endregion

#region Schreiben

function New-EobNotProcessedResult {
    <#
        Ergebnis, wenn ShouldProcess nicht bestätigt wurde (Simulation oder Ablehnung).
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param([Parameter(Mandatory)][string]$Message)

    if ($WhatIfPreference) {
        return New-EobResult -Status Succeeded -Message ('Simulation: ' + $Message)
    }
    return New-EobResult -Status Skipped -Message ('Nicht bestätigt: ' + $Message)
}

function Assert-EobAdTargetAllowed {
    <#
        Zweite Sicherheitslinie: prüft das Zielkonto unmittelbar vor einer Änderung.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Identity,
        [switch]$AllowPrivileged
    )

    $user = Resolve-EobAdUser -Identity $Identity
    $current = Get-EobCurrentIdentity
    $privileged = @(Get-EobAdPrivilegedMembership -DistinguishedName $user.DistinguishedName -PrimaryGroupId $user.PrimaryGroupId)
    $protection = Test-EobProtectedAccount -User $user -Config $script:AdState.Config -CurrentUserSid $current.Sid -PrivilegedGroups $privileged
    if ($protection.IsBlocked) {
        throw ("Änderung blockiert: " + ($protection.BlockReasons -join ' '))
    }
    if ($protection.RequiresAcknowledgement -and -not $AllowPrivileged) {
        throw ("Änderung blockiert (privilegiertes Konto, gesonderte Bestätigung erforderlich): " + ($protection.AcknowledgementReasons -join ' '))
    }
    return $user
}

function Assert-EobGroupAssignable {
    <#
        Blockiert die Zuweisung privilegierter bzw. geschützter Gruppen.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Group)

    $protection = Get-EobAdGroupProtection -Identity $Group
    if ($protection.IsProtected) {
        throw ("Zuweisung der Gruppe '$Group' blockiert: " + ($protection.Reasons -join ' '))
    }
    return $protection.GroupObject
}

function New-EobAdUserAccount {
    <#
    .SYNOPSIS
        Legt ein Benutzerkonto an.
    .PARAMETER Attributes
        Zusätzliche New-ADUser-Parameter (nur zulässige Parameternamen).
    .PARAMETER OtherAttributes
        Zusätzliche LDAP-Attribute (nur zulässige Attribute).
    #>
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Medium')]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory)][string]$SamAccountName,
        [Parameter(Mandatory)][string]$UserPrincipalName,
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][string]$DisplayName,
        [Parameter(Mandatory)][string]$GivenName,
        [Parameter(Mandatory)][string]$Surname,
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][System.Security.SecureString]$AccountPassword,
        [bool]$Enabled = $false,
        [bool]$ChangePasswordAtLogon = $true,
        [hashtable]$Attributes = @{},
        [hashtable]$OtherAttributes = @{}
    )

    foreach ($key in $Attributes.Keys) {
        if ($key -notin $script:AllowedNewUserParameters) { throw "Unzulässiger Parameter für New-ADUser: $key" }
    }
    foreach ($key in $OtherAttributes.Keys) {
        if ($key -notin $script:AllowedLdapAttributes) { throw "Unzulässiges LDAP-Attribut: $key" }
    }
    if (-not (Test-EobSamAccountName -Value $SamAccountName)) { throw "Ungültiger SamAccountName: $SamAccountName" }
    if (-not (Test-EobUserPrincipalName -Value $UserPrincipalName)) { throw "Ungültiger UPN: $UserPrincipalName" }

    if (-not $PSCmdlet.ShouldProcess("$SamAccountName ($Path)", 'Benutzerkonto anlegen')) {
        return New-EobNotProcessedResult -Message "Konto $SamAccountName würde in $Path angelegt."
    }
    Assert-EobAdConnected

    $parameters = Get-EobAdServerParameter
    $parameters['SamAccountName'] = $SamAccountName
    $parameters['UserPrincipalName'] = $UserPrincipalName
    $parameters['Name'] = $Name
    $parameters['DisplayName'] = $DisplayName
    $parameters['GivenName'] = $GivenName
    $parameters['Surname'] = $Surname
    $parameters['Path'] = $Path
    $parameters['AccountPassword'] = $AccountPassword
    $parameters['Enabled'] = $Enabled
    $parameters['ChangePasswordAtLogon'] = $ChangePasswordAtLogon
    foreach ($key in $Attributes.Keys) {
        if ($null -ne $Attributes[$key] -and "$($Attributes[$key])" -ne '') { $parameters[$key] = $Attributes[$key] }
    }
    $other = @{}
    foreach ($key in $OtherAttributes.Keys) {
        if ($null -ne $OtherAttributes[$key] -and "$($OtherAttributes[$key])" -ne '') { $other[$key] = $OtherAttributes[$key] }
    }
    if ($other.Count -gt 0) { $parameters['OtherAttributes'] = $other }

    try {
        $created = New-ADUser @parameters -PassThru -ErrorAction Stop
    }
    catch {
        throw (ConvertTo-EobAdErrorMessage -ErrorRecord $_)
    }
    $dn = [string](Get-EobPropertyValue -InputObject $created -Name 'DistinguishedName' -Default '')
    $guid = [string](Get-EobPropertyValue -InputObject $created -Name 'ObjectGUID' -Default '')
    return New-EobResult -Status Succeeded -Message "Konto $SamAccountName angelegt." -Data @{ UserDistinguishedName = $dn; UserObjectGuid = $guid }
}

function Set-EobAdUserAttribute {
    <#
    .SYNOPSIS
        Setzt bzw. löscht Attribute eines Benutzers (nur zulässige LDAP-Attribute).
    .PARAMETER Replace
        Attribut = Wert.
    .PARAMETER Clear
        Zu leerende Attribute.
    #>
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Medium')]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory)][string]$Identity,
        [hashtable]$Replace = @{},
        [string[]]$Clear = @(),
        [switch]$AllowPrivileged,
        [switch]$SkipTargetCheck
    )

    foreach ($key in @($Replace.Keys) + @($Clear)) {
        if ($key -notin $script:AllowedLdapAttributes) { throw "Unzulässiges LDAP-Attribut: $key" }
    }
    if ($Replace.Count -eq 0 -and $Clear.Count -eq 0) {
        return New-EobResult -Status Skipped -Message 'Keine Änderungen.'
    }
    $description = (@($Replace.Keys) + @($Clear | ForEach-Object { "$_ (leeren)" })) -join ', '
    if (-not $PSCmdlet.ShouldProcess($Identity, "Attribute setzen: $description")) {
        return New-EobNotProcessedResult -Message "$description"
    }
    Assert-EobAdConnected
    if (-not $SkipTargetCheck) { $null = Assert-EobAdTargetAllowed -Identity $Identity -AllowPrivileged:$AllowPrivileged }
    $parameters = Get-EobAdServerParameter
    $parameters['Identity'] = $Identity
    $values = @{}
    foreach ($key in $Replace.Keys) {
        if ($null -eq $Replace[$key] -or "$($Replace[$key])" -eq '') { $Clear += $key } else { $values[$key] = $Replace[$key] }
    }
    if ($values.Count -gt 0) { $parameters['Replace'] = $values }
    if ($Clear.Count -gt 0) { $parameters['Clear'] = @($Clear | Select-Object -Unique) }
    try {
        Set-ADUser @parameters -ErrorAction Stop
    }
    catch {
        throw (ConvertTo-EobAdErrorMessage -ErrorRecord $_)
    }
    return New-EobResult -Status Succeeded -Message "Attribute aktualisiert: $description"
}

function Set-EobAdUserManager {
    <#
    .SYNOPSIS
        Setzt oder entfernt die Führungskraft eines Benutzers.
    #>
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Medium')]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory)][string]$Identity,
        [string]$ManagerDistinguishedName,
        [switch]$SkipTargetCheck
    )

    $action = if ($ManagerDistinguishedName) { "Führungskraft setzen: $ManagerDistinguishedName" } else { 'Führungskraft entfernen' }
    if (-not $PSCmdlet.ShouldProcess($Identity, $action)) {
        return New-EobNotProcessedResult -Message "$action"
    }
    Assert-EobAdConnected
    if (-not $SkipTargetCheck) { $null = Assert-EobAdTargetAllowed -Identity $Identity }
    $parameters = Get-EobAdServerParameter
    try {
        if ($ManagerDistinguishedName) {
            Set-ADUser -Identity $Identity -Manager $ManagerDistinguishedName @parameters -ErrorAction Stop
        }
        else {
            Set-ADUser -Identity $Identity -Clear 'manager' @parameters -ErrorAction Stop
        }
    }
    catch {
        throw (ConvertTo-EobAdErrorMessage -ErrorRecord $_)
    }
    return New-EobResult -Status Succeeded -Message $action
}

function Add-EobAdGroupMembership {
    <#
    .SYNOPSIS
        Fügt einen Benutzer einer Gruppe hinzu (privilegierte Gruppen werden immer blockiert).
    #>
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Medium')]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory)][string]$Identity,
        [Parameter(Mandatory)][string]$Group
    )

    if (-not $PSCmdlet.ShouldProcess($Identity, "Zur Gruppe '$Group' hinzufügen")) {
        return New-EobNotProcessedResult -Message "Mitgliedschaft in '$Group'."
    }
    Assert-EobAdConnected
    $groupObject = Assert-EobGroupAssignable -Group $Group
    $parameters = Get-EobAdServerParameter
    try {
        Add-ADGroupMember -Identity $groupObject.DistinguishedName -Members $Identity @parameters -ErrorAction Stop
    }
    catch {
        if ($_.Exception.Message -match '(?i)already a member|bereits (ein )?Mitglied') {
            return New-EobResult -Status Succeeded -Message "Bereits Mitglied in '$Group'."
        }
        throw (ConvertTo-EobAdErrorMessage -ErrorRecord $_)
    }
    return New-EobResult -Status Succeeded -Message "Zur Gruppe '$Group' hinzugefügt."
}

function Remove-EobAdGroupMembership {
    <#
    .SYNOPSIS
        Entfernt einen Benutzer aus einer Gruppe.
    #>
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory)][string]$Identity,
        [Parameter(Mandatory)][string]$Group,
        [switch]$AllowPrivileged,
        [switch]$SkipTargetCheck
    )

    if (-not $PSCmdlet.ShouldProcess($Identity, "Aus Gruppe '$Group' entfernen")) {
        return New-EobNotProcessedResult -Message "Entfernen aus '$Group'."
    }
    Assert-EobAdConnected
    if (-not $SkipTargetCheck) { $null = Assert-EobAdTargetAllowed -Identity $Identity -AllowPrivileged:$AllowPrivileged }
    $parameters = Get-EobAdServerParameter
    try {
        Remove-ADGroupMember -Identity $Group -Members $Identity -Confirm:$false @parameters -ErrorAction Stop
    }
    catch {
        if ($_.Exception.Message -match '(?i)not a member|kein Mitglied|is not a member') {
            return New-EobResult -Status Skipped -Message "War kein Mitglied von '$Group'."
        }
        throw (ConvertTo-EobAdErrorMessage -ErrorRecord $_)
    }
    return New-EobResult -Status Succeeded -Message "Aus Gruppe '$Group' entfernt."
}

function Disable-EobAdUserAccount {
    <#
    .SYNOPSIS
        Deaktiviert ein Benutzerkonto.
    #>
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory)][string]$Identity,
        [switch]$AllowPrivileged
    )

    if (-not $PSCmdlet.ShouldProcess($Identity, 'Konto deaktivieren')) {
        return New-EobNotProcessedResult -Message 'Konto würde deaktiviert.'
    }
    Assert-EobAdConnected
    $null = Assert-EobAdTargetAllowed -Identity $Identity -AllowPrivileged:$AllowPrivileged
    try {
        $server = Get-EobAdServerParameter
        Disable-ADAccount -Identity $Identity @server -ErrorAction Stop
    }
    catch {
        throw (ConvertTo-EobAdErrorMessage -ErrorRecord $_)
    }
    return New-EobResult -Status Succeeded -Message 'Konto deaktiviert.'
}

function Enable-EobAdUserAccount {
    <#
    .SYNOPSIS
        Aktiviert ein Benutzerkonto.
    #>
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Medium')]
    [OutputType([pscustomobject])]
    param([Parameter(Mandatory)][string]$Identity)

    if (-not $PSCmdlet.ShouldProcess($Identity, 'Konto aktivieren')) {
        return New-EobNotProcessedResult -Message 'Konto würde aktiviert.'
    }
    Assert-EobAdConnected
    $null = Assert-EobAdTargetAllowed -Identity $Identity
    try {
        $server = Get-EobAdServerParameter
        Enable-ADAccount -Identity $Identity @server -ErrorAction Stop
    }
    catch {
        throw (ConvertTo-EobAdErrorMessage -ErrorRecord $_)
    }
    return New-EobResult -Status Succeeded -Message 'Konto aktiviert.'
}

function Reset-EobAdUserPassword {
    <#
    .SYNOPSIS
        Setzt das Kennwort eines Benutzers zurück (Wert wird nie protokolliert).
    #>
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory)][string]$Identity,
        [Parameter(Mandatory)][System.Security.SecureString]$NewPassword,
        [bool]$ChangePasswordAtLogon = $false,
        [switch]$AllowPrivileged
    )

    if (-not $PSCmdlet.ShouldProcess($Identity, 'Kennwort zurücksetzen')) {
        return New-EobNotProcessedResult -Message 'Kennwort würde zurückgesetzt.'
    }
    Assert-EobAdConnected
    $null = Assert-EobAdTargetAllowed -Identity $Identity -AllowPrivileged:$AllowPrivileged
    $parameters = Get-EobAdServerParameter
    try {
        Set-ADAccountPassword -Identity $Identity -Reset -NewPassword $NewPassword @parameters -ErrorAction Stop
        if ($ChangePasswordAtLogon) {
            Set-ADUser -Identity $Identity -ChangePasswordAtLogon $true @parameters -ErrorAction Stop
        }
    }
    catch {
        throw (ConvertTo-EobAdErrorMessage -ErrorRecord $_)
    }
    $suffix = if ($ChangePasswordAtLogon) { ' Änderung bei nächster Anmeldung erforderlich.' } else { '' }
    return New-EobResult -Status Succeeded -Message "Kennwort zurückgesetzt.$suffix"
}

function Set-EobAdUserPasswordChangeRequired {
    <#
    .SYNOPSIS
        Erzwingt die Kennwortänderung bei der nächsten Anmeldung.
    #>
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Medium')]
    [OutputType([pscustomobject])]
    param([Parameter(Mandatory)][string]$Identity)

    if (-not $PSCmdlet.ShouldProcess($Identity, 'Kennwortänderung erzwingen')) {
        return New-EobNotProcessedResult -Message 'Kennwortänderung würde erzwungen.'
    }
    Assert-EobAdConnected
    $null = Assert-EobAdTargetAllowed -Identity $Identity
    try {
        $server = Get-EobAdServerParameter
        Set-ADUser -Identity $Identity -ChangePasswordAtLogon $true @server -ErrorAction Stop
    }
    catch {
        throw (ConvertTo-EobAdErrorMessage -ErrorRecord $_)
    }
    return New-EobResult -Status Succeeded -Message 'Kennwortänderung bei nächster Anmeldung erzwungen.'
}

function Set-EobAdUserExpiration {
    <#
    .SYNOPSIS
        Setzt das Ablaufdatum (accountExpires) eines Kontos.
    .PARAMETER ExpiresAt
        Zeitpunkt, ab dem das Konto abgelaufen ist.
    #>
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Medium')]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory)][string]$Identity,
        [Parameter(Mandatory)][datetime]$ExpiresAt,
        [switch]$AllowPrivileged
    )

    $text = $ExpiresAt.ToString('dd.MM.yyyy HH:mm')
    if (-not $PSCmdlet.ShouldProcess($Identity, "Ablaufdatum setzen: $text")) {
        return New-EobNotProcessedResult -Message "Ablauf zum $text."
    }
    Assert-EobAdConnected
    $null = Assert-EobAdTargetAllowed -Identity $Identity -AllowPrivileged:$AllowPrivileged
    try {
        $server = Get-EobAdServerParameter
        Set-ADAccountExpiration -Identity $Identity -DateTime $ExpiresAt @server -ErrorAction Stop
    }
    catch {
        throw (ConvertTo-EobAdErrorMessage -ErrorRecord $_)
    }
    return New-EobResult -Status Succeeded -Message "Konto läuft ab: $text."
}

function Set-EobAdUserLogonHours {
    <#
    .SYNOPSIS
        Sperrt alle Anmeldezeiten (logonHours = 21 Null-Bytes).
    #>
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory)][string]$Identity,
        [switch]$AllowPrivileged
    )

    if (-not $PSCmdlet.ShouldProcess($Identity, 'Anmeldezeiten vollständig sperren')) {
        return New-EobNotProcessedResult -Message 'Anmeldezeiten würden gesperrt.'
    }
    Assert-EobAdConnected
    $null = Assert-EobAdTargetAllowed -Identity $Identity -AllowPrivileged:$AllowPrivileged
    try {
        $server = Get-EobAdServerParameter
        Set-ADUser -Identity $Identity -Replace @{ logonHours = [byte[]]::new(21) } @server -ErrorAction Stop
    }
    catch {
        throw (ConvertTo-EobAdErrorMessage -ErrorRecord $_)
    }
    return New-EobResult -Status Succeeded -Message 'Anmeldezeiten gesperrt.'
}

function Move-EobAdUserAccount {
    <#
    .SYNOPSIS
        Verschiebt ein Benutzerkonto in eine andere OU.
    #>
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Medium')]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory)][string]$Identity,
        [Parameter(Mandatory)][string]$TargetPath,
        [switch]$AllowPrivileged
    )

    if (-not $PSCmdlet.ShouldProcess($Identity, "Verschieben nach $TargetPath")) {
        return New-EobNotProcessedResult -Message "Verschieben nach $TargetPath."
    }
    Assert-EobAdConnected
    $user = Assert-EobAdTargetAllowed -Identity $Identity -AllowPrivileged:$AllowPrivileged
    if ($user.ParentContainer -ieq $TargetPath) {
        return New-EobResult -Status Skipped -Message 'Konto befindet sich bereits in der Ziel-OU.'
    }
    if (-not (Test-EobAdOrganizationalUnit -DistinguishedName $TargetPath)) {
        throw "Ziel-OU existiert nicht: $TargetPath"
    }
    if ($user.ProtectedFromAccidentalDeletion) {
        throw 'Das Konto ist vor versehentlichem Löschen geschützt und kann daher nicht verschoben werden. Schutz bewusst im AD entfernen oder Aktion deaktivieren.'
    }
    try {
        $server = Get-EobAdServerParameter
        $moved = Move-ADObject -Identity $user.DistinguishedName -TargetPath $TargetPath -PassThru @server -ErrorAction Stop
    }
    catch {
        throw (ConvertTo-EobAdErrorMessage -ErrorRecord $_)
    }
    $newDn = [string](Get-EobPropertyValue -InputObject $moved -Name 'DistinguishedName' -Default '')
    return New-EobResult -Status Succeeded -Message "Verschoben nach $TargetPath." -Data @{ UserDistinguishedName = $newDn }
}

function Remove-EobAdUserAccount {
    <#
    .SYNOPSIS
        Löscht ein Benutzerkonto endgültig (nur mit übereinstimmender ObjectGUID und Bestätigung).
    .DESCRIPTION
        Sicherungen: erneute Auflösung, Vergleich der erwarteten ObjectGUID, Schutzprüfung,
        Beachtung von ProtectedFromAccidentalDeletion, ConfirmImpact High.
    .PARAMETER ExpectedObjectGuid
        ObjectGUID aus dem Offboarding-Vorgang (Schutz vor Verwechslung bei Namenswiederverwendung).
    #>
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory)][string]$Identity,
        [Parameter(Mandatory)][string]$ExpectedObjectGuid
    )

    if (-not $PSCmdlet.ShouldProcess($Identity, 'Konto ENDGÜLTIG löschen')) {
        return New-EobNotProcessedResult -Message 'Konto würde endgültig gelöscht.'
    }
    Assert-EobAdConnected
    $user = Assert-EobAdTargetAllowed -Identity $Identity
    if ($user.ObjectGuid -ine $ExpectedObjectGuid) {
        throw "Löschung abgebrochen: Die ObjectGUID stimmt nicht mit dem Offboarding-Vorgang überein."
    }
    if ($user.Enabled) {
        throw 'Löschung abgebrochen: Das Konto ist noch aktiviert.'
    }
    if ($user.ProtectedFromAccidentalDeletion) {
        throw 'Löschung abgebrochen: Das Konto ist vor versehentlichem Löschen geschützt.'
    }
    try {
        $server = Get-EobAdServerParameter
        Remove-ADObject -Identity $user.DistinguishedName -Recursive -Confirm:$false @server -ErrorAction Stop
    }
    catch {
        throw (ConvertTo-EobAdErrorMessage -ErrorRecord $_)
    }
    return New-EobResult -Status Succeeded -Message 'Konto endgültig gelöscht.'
}

#endregion
