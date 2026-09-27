#Requires -Version 7.2
<#
    easyONB.Exchange
    Optionaler Exchange-Adapter (Exchange Online, Exchange Server, Hybrid): Status, Verbindung,
    Postfach anlegen, Abwesenheitsnotiz, Weiterleitung, Umwandlung in ein freigegebenes Postfach,
    Ausblenden aus Adresslisten und Dokumentation von Postfachberechtigungen.

    Implementierungsstand: vorbereitet. Die Funktionen sind mit Mocks getestet, aber nicht gegen
    eine reale Exchange-Umgebung. Der Status meldet das über Implementation = 'Prepared'.
    Es werden keine Module installiert; ExchangeOnlineManagement muss bei Bedarf vom Administrator
    bereitgestellt werden (Install-Module ExchangeOnlineManagement -Scope CurrentUser).
#>

Set-StrictMode -Version 3.0

$script:ExchangeState = @{
    Mode          = 'None'
    Session       = $null
    ImportedName  = ''
    ConnectedAt   = $null
    LastError     = ''
}

# Befehle, die aus einer Exchange-Server-Sitzung importiert werden (Allowlist).
$script:OnPremisesCommands = @(
    'Get-Mailbox', 'Set-Mailbox', 'Enable-Mailbox', 'Enable-RemoteMailbox', 'Set-RemoteMailbox', 'Get-RemoteMailbox',
    'Set-MailboxAutoReplyConfiguration', 'Get-MailboxAutoReplyConfiguration', 'Get-MailboxPermission'
)

#region Hilfsfunktionen

function New-EobNotProcessedResult {
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param([Parameter(Mandatory)][string]$Message)

    if ($WhatIfPreference) {
        return New-EobResult -Status Succeeded -Message ('Simulation: ' + $Message)
    }
    return New-EobResult -Status Skipped -Message ('Nicht bestätigt: ' + $Message)
}

function Test-EobExchangeOnlineConnected {
    [CmdletBinding()]
    [OutputType([bool])]
    param()

    if ($null -eq (Get-Command -Name 'Get-ConnectionInformation' -ErrorAction SilentlyContinue)) { return $false }
    try {
        $connections = @(Get-ConnectionInformation -ErrorAction Stop)
        return (@($connections | Where-Object { [string](Get-EobPropertyValue -InputObject $_ -Name 'State' -Default '') -eq 'Connected' }).Count -gt 0)
    }
    catch {
        return $false
    }
}

function Test-EobExchangeConnected {
    <#
    .SYNOPSIS
        Prüft, ob für den Modus eine nutzbare Exchange-Verbindung besteht.
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param([ValidateSet('None', 'Online', 'OnPremises', 'Hybrid')][string]$Mode = $script:ExchangeState.Mode)

    switch ($Mode) {
        'Online' { return (Test-EobExchangeOnlineConnected) }
        { $_ -in @('OnPremises', 'Hybrid') } {
            $session = $script:ExchangeState.Session
            return ($null -ne $session -and [string](Get-EobPropertyValue -InputObject $session -Name 'State' -Default '') -eq 'Opened')
        }
        default { return $false }
    }
}

function Assert-EobExchangeConnected {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Mode)

    if ($Mode -eq 'None') { throw 'Die Exchange-Integration ist deaktiviert ([Exchange] Mode=None).' }
    if (-not (Test-EobExchangeConnected -Mode $Mode)) {
        throw "Keine Verbindung zu Exchange ($Mode). Bitte zuerst unter Tools/Einstellungen verbinden."
    }
}

#endregion

#region Status und Verbindung

function Get-EobExchangeStatus {
    <#
    .SYNOPSIS
        Status der Exchange-Integration für Dashboard und Einstellungen.
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param([AllowNull()][object]$Config)

    $mode = [string](Get-EobConfigValue -Config $Config -Section 'Exchange' -Key 'Mode')
    if ($mode -eq 'None' -or [string]::IsNullOrEmpty($mode)) {
        return New-EobIntegrationStatus -Name 'Exchange' -State Disabled -Detail 'Exchange-Integration deaktiviert ([Exchange] Mode=None).' -Implementation Prepared
    }
    if ($mode -eq 'Online') {
        $module = Get-Module -ListAvailable -Name 'ExchangeOnlineManagement' -ErrorAction SilentlyContinue | Sort-Object -Property Version -Descending | Select-Object -First 1
        if ($null -eq $module -and $null -eq (Get-Command -Name 'Get-ConnectionInformation' -ErrorAction SilentlyContinue)) {
            return New-EobIntegrationStatus -Name 'Exchange' -State NotInstalled -Detail 'Modul ExchangeOnlineManagement nicht gefunden.' `
                -Hint 'Install-Module ExchangeOnlineManagement -Scope CurrentUser (durch den Administrator, keine automatische Installation).' -Implementation Prepared
        }
        $version = if ($null -ne $module) { [string]$module.Version } else { '' }
        if (Test-EobExchangeOnlineConnected) {
            return New-EobIntegrationStatus -Name 'Exchange' -State Connected -Detail 'Exchange Online verbunden.' -Version $version -Implementation Prepared
        }
        return New-EobIntegrationStatus -Name 'Exchange' -State NotConnected -Detail 'Exchange Online nicht verbunden.' -Version $version `
            -Hint 'Verbindung über Tools > Exchange verbinden herstellen.' -Implementation Prepared
    }
    $uri = [string](Get-EobConfigValue -Config $Config -Section 'Exchange' -Key 'OnPremisesUri')
    if (-not $uri) {
        return New-EobIntegrationStatus -Name 'Exchange' -State NotConfigured -Detail "Modus $mode ohne [Exchange] OnPremisesUri." `
            -Hint 'OnPremisesUri setzen, z. B. http://exchange.example.local/PowerShell/' -Implementation Prepared
    }
    if (Test-EobExchangeConnected -Mode $mode) {
        return New-EobIntegrationStatus -Name 'Exchange' -State Connected -Detail "Exchange Server ($mode) verbunden: $uri" -Implementation Prepared
    }
    return New-EobIntegrationStatus -Name 'Exchange' -State NotConnected -Detail "Exchange Server ($mode) nicht verbunden: $uri" `
        -Hint 'Verbindung über Tools > Exchange verbinden herstellen (Kerberos).' -Implementation Prepared
}

function Connect-EobExchange {
    <#
    .SYNOPSIS
        Stellt die Exchange-Verbindung gemäß [Exchange] Mode her.
    .DESCRIPTION
        Online: Connect-ExchangeOnline (moderne Authentifizierung, interaktiv).
        OnPremises/Hybrid: PSSession mit Kerberos zu OnPremisesUri; importiert werden nur die
        benötigten Befehle (Allowlist). Es werden keine Kennwörter gespeichert.
    #>
    [CmdletBinding(SupportsShouldProcess)]
    [OutputType([pscustomobject])]
    param(
        [AllowNull()][pscustomobject]$Config,
        [string]$UserPrincipalName,
        [pscredential]$Credential
    )

    $mode = [string](Get-EobConfigValue -Config $Config -Section 'Exchange' -Key 'Mode')
    if ($mode -eq 'None') { throw 'Die Exchange-Integration ist deaktiviert ([Exchange] Mode=None).' }
    if (-not $PSCmdlet.ShouldProcess($mode, 'Exchange verbinden')) {
        return New-EobNotProcessedResult -Message "Exchange ($mode) würde verbunden."
    }

    try {
        if ($mode -eq 'Online') {
            if ($null -eq (Get-Command -Name 'Connect-ExchangeOnline' -ErrorAction SilentlyContinue)) {
                throw 'Das Modul ExchangeOnlineManagement ist nicht installiert.'
            }
            $parameters = @{ ShowBanner = $false; ErrorAction = 'Stop' }
            if ($UserPrincipalName) { $parameters['UserPrincipalName'] = $UserPrincipalName }
            Connect-ExchangeOnline @parameters
        }
        else {
            $uri = [string](Get-EobConfigValue -Config $Config -Section 'Exchange' -Key 'OnPremisesUri')
            if (-not $uri) { throw '[Exchange] OnPremisesUri ist nicht gesetzt.' }
            $parsed = $null
            if (-not [uri]::TryCreate($uri, [System.UriKind]::Absolute, [ref]$parsed) -or $parsed.Scheme -notin @('http', 'https')) {
                throw "Ungültige OnPremisesUri '$uri'."
            }
            if ($null -ne $script:ExchangeState.Session) { Disconnect-EobExchange -Confirm:$false | Out-Null }
            $sessionParameters = @{
                ConfigurationName = 'Microsoft.Exchange'
                ConnectionUri     = $parsed.AbsoluteUri
                Authentication    = 'Kerberos'
                ErrorAction       = 'Stop'
            }
            if ($null -ne $Credential) { $sessionParameters['Credential'] = $Credential }
            $session = New-PSSession @sessionParameters
            $imported = Import-PSSession -Session $session -CommandName $script:OnPremisesCommands -AllowClobber -DisableNameChecking -ErrorAction Stop
            Import-Module -ModuleInfo $imported -Global -DisableNameChecking -ErrorAction Stop
            $script:ExchangeState.Session = $session
            $script:ExchangeState.ImportedName = $imported.Name
        }
        $script:ExchangeState.Mode = $mode
        $script:ExchangeState.ConnectedAt = Get-Date
        $script:ExchangeState.LastError = ''
        Write-EobLog -Level Information -Action 'ExchangeConnect' -Target $mode -Result 'Succeeded' -Message "Exchange ($mode) verbunden."
        return New-EobResult -Status Succeeded -Message "Exchange ($mode) verbunden."
    }
    catch {
        $script:ExchangeState.LastError = $_.Exception.Message
        Write-EobLog -Level Warning -Action 'ExchangeConnect' -Target $mode -Result 'Failed' -Message 'Exchange-Verbindung fehlgeschlagen.' -ErrorRecord $_
        throw
    }
}

function Disconnect-EobExchange {
    <#
    .SYNOPSIS
        Trennt die Exchange-Verbindung.
    #>
    [CmdletBinding(SupportsShouldProcess)]
    [OutputType([pscustomobject])]
    param()

    if (-not $PSCmdlet.ShouldProcess($script:ExchangeState.Mode, 'Exchange trennen')) {
        return New-EobNotProcessedResult -Message 'Exchange würde getrennt.'
    }
    if ($script:ExchangeState.Mode -eq 'Online' -and $null -ne (Get-Command -Name 'Disconnect-ExchangeOnline' -ErrorAction SilentlyContinue)) {
        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
    }
    if ($script:ExchangeState.ImportedName) {
        Remove-Module -Name $script:ExchangeState.ImportedName -Force -ErrorAction SilentlyContinue
    }
    if ($null -ne $script:ExchangeState.Session) {
        Remove-PSSession -Session $script:ExchangeState.Session -ErrorAction SilentlyContinue
    }
    $script:ExchangeState.Session = $null
    $script:ExchangeState.ImportedName = ''
    $script:ExchangeState.Mode = 'None'
    $script:ExchangeState.ConnectedAt = $null
    return New-EobResult -Status Succeeded -Message 'Exchange-Verbindung getrennt.'
}

#endregion

#region Lesen

function Get-EobMailboxInfo {
    <#
    .SYNOPSIS
        Liest Postfachdaten für Vorschau und Offboarding-Snapshot (liefert $null ohne Postfach).
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory)][string]$Identity,
        [Parameter(Mandatory)][ValidateSet('Online', 'OnPremises', 'Hybrid')][string]$Mode
    )

    Assert-EobExchangeConnected -Mode $Mode
    $mailbox = $null
    try {
        $mailbox = Get-Mailbox -Identity $Identity -ErrorAction Stop
    }
    catch {
        if ($_.Exception.Message -match "(?i)couldn't be found|not found|nicht gefunden") { return $null }
        throw
    }
    if ($null -eq $mailbox) { return $null }
    [pscustomobject]@{
        PSTypeName                    = 'Eob.MailboxInfo'
        Identity                      = $Identity
        PrimarySmtpAddress            = [string](Get-EobPropertyValue -InputObject $mailbox -Name 'PrimarySmtpAddress' -Default '')
        RecipientTypeDetails          = [string](Get-EobPropertyValue -InputObject $mailbox -Name 'RecipientTypeDetails' -Default '')
        HiddenFromAddressListsEnabled = [bool](Get-EobPropertyValue -InputObject $mailbox -Name 'HiddenFromAddressListsEnabled' -Default $false)
        ForwardingSmtpAddress         = [string](Get-EobPropertyValue -InputObject $mailbox -Name 'ForwardingSmtpAddress' -Default '')
        ForwardingAddress             = [string](Get-EobPropertyValue -InputObject $mailbox -Name 'ForwardingAddress' -Default '')
        DeliverToMailboxAndForward    = [bool](Get-EobPropertyValue -InputObject $mailbox -Name 'DeliverToMailboxAndForward' -Default $false)
        LitigationHoldEnabled         = [bool](Get-EobPropertyValue -InputObject $mailbox -Name 'LitigationHoldEnabled' -Default $false)
    }
}

function Get-EobMailboxPermissionReport {
    <#
    .SYNOPSIS
        Dokumentiert explizite Postfachberechtigungen (ohne geerbte und SELF-Einträge).
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory)][string]$Identity,
        [Parameter(Mandatory)][ValidateSet('Online', 'OnPremises', 'Hybrid')][string]$Mode
    )

    Assert-EobExchangeConnected -Mode $Mode
    foreach ($permission in @(Get-MailboxPermission -Identity $Identity -ErrorAction Stop)) {
        $user = [string](Get-EobPropertyValue -InputObject $permission -Name 'User' -Default '')
        $inherited = [bool](Get-EobPropertyValue -InputObject $permission -Name 'IsInherited' -Default $false)
        if ($inherited -or $user -match '(?i)^NT AUTHORITY\\SELF$|^S-1-5-10$') { continue }
        [pscustomobject]@{
            Mailbox      = $Identity
            User         = $user
            AccessRights = (@(Get-EobPropertyValue -InputObject $permission -Name 'AccessRights' -Default @()) | ForEach-Object { [string]$_ }) -join ', '
            Deny         = [bool](Get-EobPropertyValue -InputObject $permission -Name 'Deny' -Default $false)
        }
    }
}

#endregion

#region Schreiben (Plan-Handler)

function Enable-EobExchangeMailbox {
    <#
    .SYNOPSIS
        Legt das Postfach an: OnPremises per Enable-Mailbox, Hybrid per Enable-RemoteMailbox.
    .DESCRIPTION
        Bei Exchange Online entsteht das Postfach über die Lizenzzuweisung; dafür ist kein Aufruf nötig.
    #>
    [CmdletBinding(SupportsShouldProcess)]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory)][string]$Identity,
        [Parameter(Mandatory)][ValidateSet('OnPremises', 'Hybrid')][string]$Mode,
        [string]$Database = '',
        [string]$RemoteRoutingDomain = '',
        [string]$Alias = ''
    )

    if ($Mode -eq 'Hybrid') {
        if (-not (Test-EobDomainName -Value $RemoteRoutingDomain)) {
            throw 'Für Hybrid ist [Exchange] RemoteRoutingDomain (z. B. contoso.mail.onmicrosoft.com) erforderlich.'
        }
        $local = if ($Alias) { $Alias } else { $Identity }
        $routing = "$local@$RemoteRoutingDomain"
        if (-not $PSCmdlet.ShouldProcess($Identity, "Remote-Postfach aktivieren ($routing)")) {
            return New-EobNotProcessedResult -Message "Remote-Postfach würde für $Identity aktiviert ($routing)."
        }
        Assert-EobExchangeConnected -Mode $Mode
        $parameters = @{ Identity = $Identity; RemoteRoutingAddress = $routing; ErrorAction = 'Stop' }
        if ($Alias) { $parameters['Alias'] = $Alias }
        $null = Enable-RemoteMailbox @parameters
        return New-EobResult -Status Succeeded -Message "Remote-Postfach für $Identity aktiviert ($routing)."
    }

    if (-not $PSCmdlet.ShouldProcess($Identity, 'Postfach aktivieren')) {
        return New-EobNotProcessedResult -Message "Postfach würde für $Identity aktiviert$(if ($Database) { " (Datenbank $Database)" })."
    }
    Assert-EobExchangeConnected -Mode $Mode
    $parameters = @{ Identity = $Identity; ErrorAction = 'Stop' }
    if ($Database) { $parameters['Database'] = $Database }
    if ($Alias) { $parameters['Alias'] = $Alias }
    $null = Enable-Mailbox @parameters
    return New-EobResult -Status Succeeded -Message "Postfach für $Identity aktiviert."
}

function Set-EobMailboxAutoReply {
    <#
    .SYNOPSIS
        Aktiviert die Abwesenheitsnotiz (optional zeitgesteuert).
    #>
    [CmdletBinding(SupportsShouldProcess)]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory)][string]$Identity,
        [Parameter(Mandatory)][ValidateSet('Online', 'OnPremises', 'Hybrid')][string]$Mode,
        [Parameter(Mandatory)][string]$Message,
        [string]$ExternalMessage,
        [ValidateSet('None', 'Known', 'All')][string]$ExternalAudience = 'All',
        [Nullable[datetime]]$StartTime,
        [Nullable[datetime]]$EndTime
    )

    if (-not $ExternalMessage) { $ExternalMessage = $Message }
    $html = { param($Text) '<html><body><p>' + ((ConvertTo-EobHtmlEncoded -Value $Text) -replace '\r?\n', '<br>') + '</p></body></html>' }
    if (-not $PSCmdlet.ShouldProcess($Identity, 'Abwesenheitsnotiz aktivieren')) {
        return New-EobNotProcessedResult -Message "Abwesenheitsnotiz würde für $Identity aktiviert."
    }
    Assert-EobExchangeConnected -Mode $Mode
    $parameters = @{
        Identity         = $Identity
        InternalMessage  = (& $html $Message)
        ExternalMessage  = (& $html $ExternalMessage)
        ExternalAudience = $ExternalAudience
        ErrorAction      = 'Stop'
    }
    if ($null -ne $StartTime -and $null -ne $EndTime) {
        if ($EndTime -le $StartTime) { throw 'Das Ende der Abwesenheitsnotiz muss nach dem Beginn liegen.' }
        $parameters['AutoReplyState'] = 'Scheduled'
        $parameters['StartTime'] = $StartTime
        $parameters['EndTime'] = $EndTime
    }
    else {
        $parameters['AutoReplyState'] = 'Enabled'
    }
    Set-MailboxAutoReplyConfiguration @parameters
    return New-EobResult -Status Succeeded -Message "Abwesenheitsnotiz für $Identity aktiviert ($($parameters['AutoReplyState']))."
}

function Set-EobMailboxForwarding {
    <#
    .SYNOPSIS
        Richtet eine Weiterleitung an eine interne Adresse ein (z. B. Führungskraft).
    .DESCRIPTION
        Externe Weiterleitungen werden nur mit -AllowExternal zugelassen, da sie häufig durch
        Richtlinien blockiert sind und ein Datenschutzrisiko darstellen.
    #>
    [CmdletBinding(SupportsShouldProcess)]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory)][string]$Identity,
        [Parameter(Mandatory)][ValidateSet('Online', 'OnPremises', 'Hybrid')][string]$Mode,
        [Parameter(Mandatory)][string]$ForwardTo,
        [string[]]$InternalDomains = @(),
        [switch]$DeliverToMailboxAndForward,
        [switch]$AllowExternal
    )

    if (-not (Test-EobEmailAddress -Value $ForwardTo)) { throw "Ungültige Weiterleitungsadresse '$ForwardTo'." }
    $domain = $ForwardTo.Substring($ForwardTo.IndexOf('@') + 1)
    $isInternal = @($InternalDomains | Where-Object { $_ -and $domain -ieq $_ }).Count -gt 0
    if (-not $isInternal -and -not $AllowExternal) {
        throw "Weiterleitung an '$ForwardTo' ist extern (Domäne $domain) und nicht freigegeben."
    }
    if (-not $PSCmdlet.ShouldProcess($Identity, "Weiterleitung an $ForwardTo einrichten")) {
        return New-EobNotProcessedResult -Message "Weiterleitung von $Identity an $ForwardTo würde eingerichtet."
    }
    Assert-EobExchangeConnected -Mode $Mode
    Set-Mailbox -Identity $Identity -ForwardingSmtpAddress "smtp:$ForwardTo" -DeliverToMailboxAndForward ([bool]$DeliverToMailboxAndForward) -ErrorAction Stop
    return New-EobResult -Status Succeeded -Message "Weiterleitung von $Identity an $ForwardTo eingerichtet."
}

function Set-EobMailboxHidden {
    <#
    .SYNOPSIS
        Blendet ein Postfach über Exchange (Set-Mailbox) aus den Adresslisten aus.
    .DESCRIPTION
        Für synchronisierte Benutzer (Online/Hybrid) muss das AD-Attribut msExchHideFromAddressLists
        gesetzt werden; das erledigt der Offboarding-Plan über den AD-Adapter.
    #>
    [CmdletBinding(SupportsShouldProcess)]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory)][string]$Identity,
        [Parameter(Mandatory)][ValidateSet('Online', 'OnPremises', 'Hybrid')][string]$Mode
    )

    if (-not $PSCmdlet.ShouldProcess($Identity, 'Aus Adresslisten ausblenden')) {
        return New-EobNotProcessedResult -Message "$Identity würde aus den Adresslisten ausgeblendet."
    }
    Assert-EobExchangeConnected -Mode $Mode
    Set-Mailbox -Identity $Identity -HiddenFromAddressListsEnabled $true -ErrorAction Stop
    return New-EobResult -Status Succeeded -Message "$Identity aus den Adresslisten ausgeblendet."
}

function ConvertTo-EobSharedMailbox {
    <#
    .SYNOPSIS
        Wandelt ein Benutzerpostfach in ein freigegebenes Postfach um.
    .DESCRIPTION
        Online/OnPremises: Set-Mailbox -Type Shared. Hybrid (Postfach in Exchange Online, Objekt
        lokal verwaltet): Set-RemoteMailbox -Type Shared in der lokalen Exchange-Sitzung.
        Hinweis: Freigegebene Postfächer über 50 GB oder mit Archiv benötigen weiterhin eine Lizenz.
    #>
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory)][string]$Identity,
        [Parameter(Mandatory)][ValidateSet('Online', 'OnPremises', 'Hybrid')][string]$Mode
    )

    if (-not $PSCmdlet.ShouldProcess($Identity, 'In freigegebenes Postfach umwandeln')) {
        return New-EobNotProcessedResult -Message "Postfach $Identity würde in ein freigegebenes Postfach umgewandelt."
    }
    Assert-EobExchangeConnected -Mode $Mode
    if ($Mode -eq 'Hybrid') {
        Set-RemoteMailbox -Identity $Identity -Type Shared -ErrorAction Stop
    }
    else {
        Set-Mailbox -Identity $Identity -Type Shared -ErrorAction Stop
    }
    return New-EobResult -Status Succeeded -Message "Postfach $Identity in ein freigegebenes Postfach umgewandelt. Lizenzbedarf prüfen (Größe/Archiv)."
}

#endregion
