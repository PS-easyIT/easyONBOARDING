#Requires -Version 7.2
<#
    easyONB.Entra
    Optionale Anbindung an Microsoft Graph (Sitzungen widerrufen) und Entra Connect Sync
    (Synchronisationszyklus auslösen).

    Implementierungsstand: vorbereitet. Mit Mocks getestet, nicht gegen einen realen Tenant oder
    Entra-Connect-Server. Es werden keine Module installiert; Microsoft.Graph.Authentication und
    Microsoft.Graph.Users.Actions (Revoke-MgUserSignInSession) stellt bei Bedarf der Administrator bereit (-Scope CurrentUser).
    Graph wird delegiert mit minimalen Berechtigungen verbunden (User.Read.All, User.RevokeSessions.All)
    oder - für die geplante Aufgabe - per Zertifikat als Anwendung ([Graph] ClientId, CertificateThumbprint,
    TenantId; Anwendungsberechtigung User.RevokeSessions.All). In der Konfiguration steht nur der
    Fingerabdruck; das Zertifikat liegt im Zertifikatspeicher des ausführenden Kontos.
    Entra Connect: Statusprüfung über Get-ADSyncScheduler (Erreichbarkeit, Zyklus, Stagingmodus).
#>

Set-StrictMode -Version 3.0

$script:GraphScopes = @('User.Read.All', 'User.RevokeSessions.All')
# Als Anwendung (App-only) genügt die Berechtigung zum Widerrufen.
$script:GraphAppOnlyScopes = @('User.RevokeSessions.All')

#region Hilfsfunktionen

function New-EobNotProcessedResult {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Erzeugt nur ein Objekt im Speicher; keine Systemänderung.')]
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param([Parameter(Mandatory)][string]$Message)

    if ($WhatIfPreference) {
        return New-EobResult -Status Succeeded -Message ('Simulation: ' + $Message)
    }
    return New-EobResult -Status Skipped -Message ('Nicht bestätigt: ' + $Message)
}

function Get-EobGraphContext {
    [CmdletBinding()]
    [OutputType([object])]
    param()

    if ($null -eq (Get-Command -Name 'Get-MgContext' -ErrorAction SilentlyContinue)) { return $null }
    try { return (Get-MgContext -ErrorAction Stop) } catch { return $null }
}

function Test-EobLocalComputerName {
    [CmdletBinding()]
    [OutputType([bool])]
    param([Parameter(Mandatory)][string]$Name)

    $short = ($Name -split '\.')[0]
    return ($Name -ieq 'localhost' -or $Name -eq '.' -or $short -ieq [Environment]::MachineName)
}

function Invoke-EobAdSyncCommand {
    <#
        Führt einen festen Skriptblock lokal oder per PowerShell-Remoting auf dem Entra-Connect-Server aus.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Server,
        [Parameter(Mandatory)][scriptblock]$ScriptBlock,
        [object[]]$ArgumentList = @()
    )

    if (Test-EobLocalComputerName -Name $Server) {
        return (& $ScriptBlock @ArgumentList)
    }
    return (Invoke-Command -ComputerName $Server -ScriptBlock $ScriptBlock -ArgumentList $ArgumentList -ErrorAction Stop)
}

function Test-EobGraphAppOnlyConfigured {
    <#
    .SYNOPSIS
        Prüft, ob die Zertifikatsanmeldung an Microsoft Graph konfiguriert ist (ClientId, Fingerabdruck, TenantId).
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param([AllowNull()][object]$Config)

    $values = foreach ($key in @('ClientId', 'CertificateThumbprint', 'TenantId')) { [string](Get-EobConfigValue -Config $Config -Section 'Graph' -Key $key) }
    return (@($values | Where-Object { $_ }).Count -eq 3)
}

function Invoke-EobAdSyncCycle {
    <#
        Führt Start-ADSyncSyncCycle lokal oder per PowerShell-Remoting aus. Der Befehl ist fest
        vorgegeben; ein frei konfigurierbarer Befehl (Legacy: SyncCommand) wird nicht unterstützt.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Server,
        [Parameter(Mandatory)][ValidateSet('Delta', 'Initial')][string]$PolicyType
    )

    $scriptBlock = {
        param([string]$Type)
        Import-Module -Name ADSync -ErrorAction Stop
        Start-ADSyncSyncCycle -PolicyType $Type -ErrorAction Stop
    }
    return (Invoke-EobAdSyncCommand -Server $Server -ScriptBlock $scriptBlock -ArgumentList @($PolicyType))
}

#endregion

#region Microsoft Graph

function Get-EobGraphStatus {
    <#
    .SYNOPSIS
        Status der Microsoft-Graph-Anbindung.
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param([AllowNull()][object]$Config)

    if (-not (Get-EobConfigValue -Config $Config -Section 'Graph' -Key 'Enabled' -As Bool)) {
        return New-EobIntegrationStatus -Name 'Microsoft Graph' -State Disabled -Detail 'Graph-Anbindung deaktiviert ([Graph] Enabled=0).' -Implementation Prepared
    }
    $module = Get-Module -ListAvailable -Name 'Microsoft.Graph.Authentication' -ErrorAction SilentlyContinue | Sort-Object -Property Version -Descending | Select-Object -First 1
    if ($null -eq $module -and $null -eq (Get-Command -Name 'Get-MgContext' -ErrorAction SilentlyContinue)) {
        return New-EobIntegrationStatus -Name 'Microsoft Graph' -State NotInstalled -Detail 'Microsoft.Graph.Authentication nicht gefunden.' `
            -Hint 'Install-Module Microsoft.Graph.Authentication, Microsoft.Graph.Users.Actions -Scope CurrentUser (durch den Administrator).' -Implementation Prepared
    }
    $version = if ($null -ne $module) { [string]$module.Version } else { '' }
    # Der Widerruf benötigt Microsoft.Graph.Users.Actions; ohne das Modul scheitert der Schritt erst bei der Ausführung.
    if ($null -eq (Get-Command -Name 'Revoke-MgUserSignInSession' -ErrorAction SilentlyContinue)) {
        return New-EobIntegrationStatus -Name 'Microsoft Graph' -State NotInstalled -Detail 'Microsoft.Graph.Users.Actions (Revoke-MgUserSignInSession) nicht gefunden.' -Version $version `
            -Hint 'Install-Module Microsoft.Graph.Users.Actions -Scope CurrentUser (durch den Administrator).' -Implementation Prepared
    }
    $context = Get-EobGraphContext
    if ($null -ne $context) {
        $appOnly = [string](Get-EobPropertyValue -InputObject $context -Name 'AuthType' -Default '') -eq 'AppOnly'
        $account = if ($appOnly) { 'Anwendung ' + [string](Get-EobPropertyValue -InputObject $context -Name 'ClientId' -Default '') } else { [string](Get-EobPropertyValue -InputObject $context -Name 'Account' -Default '') }
        $scopes = @(Get-EobPropertyValue -InputObject $context -Name 'Scopes' -Default @())
        $required = if ($appOnly) { $script:GraphAppOnlyScopes } else { $script:GraphScopes }
        $missing = @($required | Where-Object { $_ -notin $scopes })
        if ($missing.Count -gt 0) {
            return New-EobIntegrationStatus -Name 'Microsoft Graph' -State Error -Detail "Verbunden als $account, es fehlen Berechtigungen: $($missing -join ', ')" `
                -Version $version -Hint 'Neu verbinden (Tools > Microsoft Graph verbinden).' -Implementation Prepared
        }
        return New-EobIntegrationStatus -Name 'Microsoft Graph' -State Connected -Detail "Verbunden als $account" -Version $version -Implementation Prepared
    }
    return New-EobIntegrationStatus -Name 'Microsoft Graph' -State NotConnected -Detail 'Nicht verbunden.' -Version $version `
        -Hint 'Verbindung über Tools > Microsoft Graph verbinden herstellen.' -Implementation Prepared
}

function Connect-EobGraph {
    <#
    .SYNOPSIS
        Verbindet Microsoft Graph delegiert mit minimalen Berechtigungen.
    #>
    [CmdletBinding(SupportsShouldProcess)]
    [OutputType([pscustomobject])]
    param([AllowNull()][pscustomobject]$Config)

    if (-not (Get-EobConfigValue -Config $Config -Section 'Graph' -Key 'Enabled' -As Bool)) {
        throw 'Die Graph-Anbindung ist deaktiviert ([Graph] Enabled=0).'
    }
    if ($null -eq (Get-Command -Name 'Connect-MgGraph' -ErrorAction SilentlyContinue)) {
        throw 'Das Modul Microsoft.Graph.Authentication ist nicht installiert.'
    }
    $appOnly = Test-EobGraphAppOnlyConfigured -Config $Config
    $description = if ($appOnly) { 'als Anwendung per Zertifikat' } else { $script:GraphScopes -join ', ' }
    if (-not $PSCmdlet.ShouldProcess('Microsoft Graph', "Verbinden ($description)")) {
        return New-EobNotProcessedResult -Message 'Microsoft Graph würde verbunden.'
    }
    $tenant = [string](Get-EobConfigValue -Config $Config -Section 'Graph' -Key 'TenantId')
    if ($appOnly) {
        # App-only: Berechtigungen sind der Anwendung zugewiesen; -Scopes ist hier nicht zulässig.
        $parameters = @{
            ClientId              = [string](Get-EobConfigValue -Config $Config -Section 'Graph' -Key 'ClientId')
            CertificateThumbprint = [string](Get-EobConfigValue -Config $Config -Section 'Graph' -Key 'CertificateThumbprint')
            TenantId              = $tenant
            NoWelcome             = $true
            ErrorAction           = 'Stop'
        }
    }
    else {
        $parameters = @{ Scopes = $script:GraphScopes; NoWelcome = $true; ErrorAction = 'Stop' }
        if ($tenant) { $parameters['TenantId'] = $tenant }
    }
    try {
        $null = Connect-MgGraph @parameters
    }
    catch {
        Write-EobLog -Level Warning -Action 'GraphConnect' -Result 'Failed' -Message 'Graph-Verbindung fehlgeschlagen.' -ErrorRecord $_
        throw
    }
    Write-EobLog -Level Information -Action 'GraphConnect' -Result 'Succeeded' -Message "Microsoft Graph verbunden ($description)."
    return New-EobResult -Status Succeeded -Message "Microsoft Graph verbunden ($description)."
}

function Disconnect-EobGraph {
    <#
    .SYNOPSIS
        Trennt die Graph-Verbindung.
    #>
    [CmdletBinding(SupportsShouldProcess)]
    [OutputType([pscustomobject])]
    param()

    if (-not $PSCmdlet.ShouldProcess('Microsoft Graph', 'Trennen')) {
        return New-EobNotProcessedResult -Message 'Microsoft Graph würde getrennt.'
    }
    if ($null -ne (Get-Command -Name 'Disconnect-MgGraph' -ErrorAction SilentlyContinue)) {
        $null = Disconnect-MgGraph -ErrorAction SilentlyContinue
    }
    return New-EobResult -Status Succeeded -Message 'Microsoft Graph getrennt.'
}

function Revoke-EobEntraUserSession {
    <#
    .SYNOPSIS
        Widerruft alle Anmeldesitzungen und Aktualisierungstoken eines Benutzers (Entra ID).
    #>
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
    [OutputType([pscustomobject])]
    param([Parameter(Mandatory)][string]$UserPrincipalName)

    if (-not (Test-EobUserPrincipalName -Value $UserPrincipalName)) {
        throw "Ungültiger UPN '$UserPrincipalName'."
    }
    if (-not $PSCmdlet.ShouldProcess($UserPrincipalName, 'Anmeldesitzungen widerrufen')) {
        return New-EobNotProcessedResult -Message "Anmeldesitzungen von $UserPrincipalName würden widerrufen."
    }
    if ($null -eq (Get-EobGraphContext)) {
        throw 'Keine Verbindung zu Microsoft Graph. Bitte zuerst unter Tools verbinden.'
    }
    $null = Revoke-MgUserSignInSession -UserId $UserPrincipalName -ErrorAction Stop
    return New-EobResult -Status Succeeded -Message "Anmeldesitzungen von $UserPrincipalName widerrufen."
}

#endregion

#region Entra Connect Sync

function Get-EobEntraConnectStatus {
    <#
    .SYNOPSIS
        Status der Entra-Connect-Synchronisation.
    .PARAMETER TestConnection
        Fragt den Server ab (Get-ADSyncScheduler): Erreichbarkeit, Rechte, aktiver Zyklus und Stagingmodus.
        Ohne den Schalter wird nur die Konfiguration bewertet (Dashboard, ohne Netzwerkzugriff).
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [AllowNull()][object]$Config,
        [switch]$TestConnection
    )

    if (-not (Get-EobConfigValue -Config $Config -Section 'ADSync' -Key 'EnableADSync' -As Bool)) {
        return New-EobIntegrationStatus -Name 'Entra Connect Sync' -State Disabled -Detail 'Synchronisation deaktiviert ([ADSync] EnableADSync=0).' -Implementation Prepared
    }
    $server = [string](Get-EobConfigValue -Config $Config -Section 'ADSync' -Key 'ADSyncServer')
    if (-not $server) {
        return New-EobIntegrationStatus -Name 'Entra Connect Sync' -State NotConfigured -Detail 'Kein Server konfiguriert.' -Hint '[ADSync] ADSyncServer setzen.' -Implementation Prepared
    }
    $policy = [string](Get-EobConfigValue -Config $Config -Section 'ADSync' -Key 'PolicyType')
    $hint = 'Voraussetzung: PowerShell-Remoting auf dem Server und Mitgliedschaft in ADSyncOperators.'
    if (-not $TestConnection) {
        return New-EobIntegrationStatus -Name 'Entra Connect Sync' -State Available -Detail "Server $server, Zyklus $policy (Erreichbarkeit: Tools > Entra Connect prüfen)" `
            -Hint $hint -Implementation Prepared
    }
    try {
        $scheduler = Invoke-EobAdSyncCommand -Server $server -ScriptBlock {
            Import-Module -Name ADSync -ErrorAction Stop
            $value = Get-ADSyncScheduler -ErrorAction Stop
            [pscustomobject]@{
                SyncCycleEnabled    = [bool]$value.SyncCycleEnabled
                SyncCycleInProgress = [bool]$value.SyncCycleInProgress
                StagingModeEnabled  = [bool]$value.StagingModeEnabled
                NextSyncCycle       = $value.NextSyncCycleStartTimeInUTC
            }
        } | Select-Object -First 1
    }
    catch {
        return New-EobIntegrationStatus -Name 'Entra Connect Sync' -State Error -Detail "Server $server nicht abfragbar: $(Protect-EobSensitiveText -Text $_.Exception.Message)" `
            -Hint $hint -Implementation Prepared
    }
    if ($null -eq $scheduler) {
        return New-EobIntegrationStatus -Name 'Entra Connect Sync' -State Error -Detail "Server $server lieferte keinen Status (Get-ADSyncScheduler)." -Hint $hint -Implementation Prepared
    }
    if ([bool](Get-EobPropertyValue -InputObject $scheduler -Name 'StagingModeEnabled' -Default $false)) {
        return New-EobIntegrationStatus -Name 'Entra Connect Sync' -State Error -Detail "Server $server ist im Stagingmodus: Ein Zyklus exportiert keine Änderungen." `
            -Hint '[ADSync] ADSyncServer auf den aktiven Entra-Connect-Server setzen.' -Implementation Prepared
    }
    $parts = [System.Collections.Generic.List[string]]::new()
    $parts.Add("Server $server erreichbar, Zyklus $policy")
    if (-not [bool](Get-EobPropertyValue -InputObject $scheduler -Name 'SyncCycleEnabled' -Default $true)) { $parts.Add('automatischer Zyklus deaktiviert') }
    if ([bool](Get-EobPropertyValue -InputObject $scheduler -Name 'SyncCycleInProgress' -Default $false)) { $parts.Add('Synchronisation läuft gerade') }
    $next = Get-EobPropertyValue -InputObject $scheduler -Name 'NextSyncCycle' -Default $null
    if ($next -is [datetime]) { $parts.Add('nächster Zyklus ' + $next.ToLocalTime().ToString('dd.MM.yyyy HH:mm')) }
    return New-EobIntegrationStatus -Name 'Entra Connect Sync' -State Connected -Detail ($parts -join ', ') -Implementation Prepared
}

function Start-EobEntraConnectSync {
    <#
    .SYNOPSIS
        Löst einen Entra-Connect-Synchronisationszyklus aus (Delta oder Initial).
    #>
    [CmdletBinding(SupportsShouldProcess)]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory)][string]$Server,
        [ValidateSet('Delta', 'Initial')][string]$PolicyType = 'Delta'
    )

    if (-not (Test-EobLocalComputerName -Name $Server) -and -not (Test-EobDomainName -Value $Server) -and $Server -notmatch '^[A-Za-z0-9][A-Za-z0-9-]{0,62}$') {
        throw "Ungültiger Servername '$Server'."
    }
    if (-not $PSCmdlet.ShouldProcess($Server, "Entra Connect Sync ($PolicyType) starten")) {
        return New-EobNotProcessedResult -Message "Entra Connect Sync ($PolicyType) würde auf $Server gestartet."
    }
    try {
        $null = Invoke-EobAdSyncCycle -Server $Server -PolicyType $PolicyType
    }
    catch {
        if ($_.Exception.Message -match '(?i)busy|already running|läuft bereits') {
            return New-EobResult -Status Warning -Message "Auf $Server läuft bereits eine Synchronisation; die Änderungen werden im laufenden oder nächsten Zyklus übertragen."
        }
        throw
    }
    return New-EobResult -Status Succeeded -Message "Entra Connect Sync ($PolicyType) auf $Server gestartet."
}

#endregion
