#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.3.0' }
<#
    Tests der optionalen Adapter für Exchange, Microsoft Graph und Entra Connect.
    Alle externen Cmdlets sind Stubs bzw. Mocks; es wird kein reales System angesprochen.
#>

BeforeAll {
    . (Join-Path -Path $PSScriptRoot -ChildPath '../TestHelpers.ps1')
    Import-EobTestModule

    $script:ExModule = 'easyONB.Exchange'
    $script:EntraModule = 'easyONB.Entra'
    $null = Initialize-EobLogging -Directory (Join-Path $TestDrive 'logs')

    function script:New-TestConfig {
        param([string]$Content = '')
        $path = Join-Path $TestDrive ([guid]::NewGuid().ToString('N') + '.ini')
        Set-Content -LiteralPath $path -Value $Content -Encoding utf8
        return Import-EobConfiguration -Path $path -NoIncludes
    }
}

Describe 'Exchange: Status und Verbindung' {
    It 'meldet die deaktivierte Integration' {
        $status = Get-EobExchangeStatus -Config (New-TestConfig)
        $status.State | Should -Be 'Disabled'
        $status.Implementation | Should -Be 'Prepared'
    }

    It 'meldet Exchange Online ohne Verbindung als NotConnected' {
        Mock -ModuleName $script:ExModule Get-ConnectionInformation { @() }
        (Get-EobExchangeStatus -Config (New-TestConfig -Content "[Exchange]`nMode=Online")).State | Should -Be 'NotConnected'
    }

    It 'meldet eine bestehende Exchange-Online-Verbindung' {
        Mock -ModuleName $script:ExModule Get-ConnectionInformation { [pscustomobject]@{ State = 'Connected' } }
        (Get-EobExchangeStatus -Config (New-TestConfig -Content "[Exchange]`nMode=Online")).State | Should -Be 'Connected'
    }

    It 'meldet Exchange Server ohne URI als NotConfigured' {
        (Get-EobExchangeStatus -Config (New-TestConfig -Content "[Exchange]`nMode=OnPremises")).State | Should -Be 'NotConfigured'
    }

    It 'meldet Exchange Server mit URI, aber ohne Sitzung als NotConnected' {
        $config = New-TestConfig -Content "[Exchange]`nMode=Hybrid`nOnPremisesUri=http://exchange.example.local/PowerShell/"
        (Get-EobExchangeStatus -Config $config).State | Should -Be 'NotConnected'
    }

    It 'verbindet Exchange Online ohne Banner und ohne gespeicherte Anmeldedaten' {
        Mock -ModuleName $script:ExModule Connect-ExchangeOnline { }
        $result = Connect-EobExchange -Config (New-TestConfig -Content "[Exchange]`nMode=Online") -UserPrincipalName 'admin@example.com' -Confirm:$false
        $result.Status | Should -Be 'Succeeded'
        Should -Invoke -ModuleName $script:ExModule Connect-ExchangeOnline -Times 1 -ParameterFilter { $ShowBanner -eq $false -and $UserPrincipalName -eq 'admin@example.com' }
    }

    It 'verbindet in der Simulation nicht' {
        Mock -ModuleName $script:ExModule Connect-ExchangeOnline { }
        $null = Connect-EobExchange -Config (New-TestConfig -Content "[Exchange]`nMode=Online") -WhatIf
        Should -Invoke -ModuleName $script:ExModule Connect-ExchangeOnline -Times 0
    }

    It 'verweigert die Verbindung bei deaktivierter Integration' {
        { Connect-EobExchange -Config (New-TestConfig) -Confirm:$false } | Should -Throw '*deaktiviert*'
    }

    It 'verweigert ungültige Exchange-URIs' {
        $config = New-TestConfig -Content "[Exchange]`nMode=OnPremises"
        $config.Sections['Exchange']['OnPremisesUri'] = 'ftp://exchange.example.local'
        { Connect-EobExchange -Config $config -Confirm:$false } | Should -Throw '*OnPremisesUri*'
    }
}

Describe 'Exchange: Postfachaktionen' {
    BeforeEach {
        Mock -ModuleName $script:ExModule Test-EobExchangeConnected { $true }
        Mock -ModuleName $script:ExModule Enable-Mailbox { [pscustomobject]@{ Name = 'x' } }
        Mock -ModuleName $script:ExModule Enable-RemoteMailbox { [pscustomobject]@{ Name = 'x' } }
        Mock -ModuleName $script:ExModule Set-Mailbox { }
        Mock -ModuleName $script:ExModule Set-RemoteMailbox { }
        Mock -ModuleName $script:ExModule Set-MailboxAutoReplyConfiguration { }
    }

    It 'legt ein lokales Postfach in der angegebenen Datenbank an' {
        $result = Enable-EobExchangeMailbox -Identity 'mmustermann' -Mode OnPremises -Database 'DB01' -Confirm:$false
        $result.Status | Should -Be 'Succeeded'
        Should -Invoke -ModuleName $script:ExModule Enable-Mailbox -Times 1 -ParameterFilter { $Identity -eq 'mmustermann' -and $Database -eq 'DB01' }
    }

    It 'legt im Hybridbetrieb ein Remote-Postfach mit Routingadresse an' {
        $null = Enable-EobExchangeMailbox -Identity 'mmustermann' -Mode Hybrid -RemoteRoutingDomain 'example.mail.onmicrosoft.com' -Confirm:$false
        Should -Invoke -ModuleName $script:ExModule Enable-RemoteMailbox -Times 1 -ParameterFilter { $RemoteRoutingAddress -eq 'mmustermann@example.mail.onmicrosoft.com' }
    }

    It 'verlangt im Hybridbetrieb eine Routingdomäne' {
        { Enable-EobExchangeMailbox -Identity 'mmustermann' -Mode Hybrid -Confirm:$false } | Should -Throw '*RemoteRoutingDomain*'
    }

    It 'simuliert ohne Exchange-Aufruf und ohne Verbindung' {
        Mock -ModuleName $script:ExModule Test-EobExchangeConnected { $false }
        $result = Enable-EobExchangeMailbox -Identity 'mmustermann' -Mode OnPremises -WhatIf
        $result.Message | Should -Match '^Simulation'
        Should -Invoke -ModuleName $script:ExModule Enable-Mailbox -Times 0
    }

    It 'bricht ohne Verbindung mit einer klaren Meldung ab' {
        Mock -ModuleName $script:ExModule Test-EobExchangeConnected { $false }
        { Enable-EobExchangeMailbox -Identity 'mmustermann' -Mode OnPremises -Confirm:$false } | Should -Throw '*Keine Verbindung*'
    }

    It 'aktiviert eine zeitgesteuerte Abwesenheitsnotiz mit kodiertem Text' {
        $start = [datetime]'2027-01-01'
        $null = Set-EobMailboxAutoReply -Identity 'mmustermann' -Mode Online -Message "Nicht mehr im Unternehmen <b>`nBitte an x wenden" -StartTime $start -EndTime $start.AddDays(90) -Confirm:$false
        Should -Invoke -ModuleName $script:ExModule Set-MailboxAutoReplyConfiguration -Times 1 -ParameterFilter {
            $AutoReplyState -eq 'Scheduled' -and $InternalMessage -match '&lt;b&gt;<br>Bitte' -and $ExternalAudience -eq 'All'
        }
    }

    It 'lehnt ein Ende vor dem Beginn ab' {
        $start = [datetime]'2027-01-01'
        { Set-EobMailboxAutoReply -Identity 'x' -Mode Online -Message 'm' -StartTime $start -EndTime $start.AddDays(-1) -Confirm:$false } | Should -Throw '*Ende*'
    }

    It 'richtet eine interne Weiterleitung ein' {
        $null = Set-EobMailboxForwarding -Identity 'mmustermann' -Mode Online -ForwardTo 'chefin@example.com' -InternalDomains 'example.com' -Confirm:$false
        Should -Invoke -ModuleName $script:ExModule Set-Mailbox -Times 1 -ParameterFilter { $ForwardingSmtpAddress -eq 'smtp:chefin@example.com' -and $DeliverToMailboxAndForward -eq $false }
    }

    It 'blockiert externe Weiterleitungen ohne Freigabe' {
        { Set-EobMailboxForwarding -Identity 'mmustermann' -Mode Online -ForwardTo 'privat@example.net' -InternalDomains 'example.com' -Confirm:$false } | Should -Throw '*extern*'
        Should -Invoke -ModuleName $script:ExModule Set-Mailbox -Times 0
    }

    It 'wandelt in ein freigegebenes Postfach um (<Mode>)' -ForEach @(
        @{ Mode = 'Online'; Command = 'Set-Mailbox' }, @{ Mode = 'OnPremises'; Command = 'Set-Mailbox' }, @{ Mode = 'Hybrid'; Command = 'Set-RemoteMailbox' }
    ) {
        $null = ConvertTo-EobSharedMailbox -Identity 'mmustermann' -Mode $Mode -Confirm:$false
        Should -Invoke -ModuleName $script:ExModule $Command -Times 1 -ParameterFilter { $Type -eq 'Shared' }
    }

    It 'blendet ein Postfach aus den Adresslisten aus' {
        $null = Set-EobMailboxHidden -Identity 'mmustermann' -Mode OnPremises -Confirm:$false
        Should -Invoke -ModuleName $script:ExModule Set-Mailbox -Times 1 -ParameterFilter { $HiddenFromAddressListsEnabled -eq $true }
    }
}

Describe 'Exchange: Lesen' {
    BeforeEach {
        Mock -ModuleName $script:ExModule Test-EobExchangeConnected { $true }
    }

    It 'liefert $null, wenn kein Postfach existiert' {
        Mock -ModuleName $script:ExModule Get-Mailbox { throw "The operation couldn't be performed because object 'x' couldn't be found." }
        Get-EobMailboxInfo -Identity 'x' -Mode Online | Should -BeNullOrEmpty
    }

    It 'normalisiert Postfachdaten' {
        Mock -ModuleName $script:ExModule Get-Mailbox { [pscustomobject]@{ PrimarySmtpAddress = 'max@example.com'; RecipientTypeDetails = 'UserMailbox'; HiddenFromAddressListsEnabled = $false } }
        $info = Get-EobMailboxInfo -Identity 'mmustermann' -Mode Online
        $info.PrimarySmtpAddress | Should -Be 'max@example.com'
        $info.RecipientTypeDetails | Should -Be 'UserMailbox'
        $info.ForwardingSmtpAddress | Should -Be ''
    }

    It 'dokumentiert nur explizite Berechtigungen' {
        Mock -ModuleName $script:ExModule Get-MailboxPermission {
            [pscustomobject]@{ User = 'NT AUTHORITY\SELF'; AccessRights = @('FullAccess'); IsInherited = $false; Deny = $false }
            [pscustomobject]@{ User = 'EXAMPLE\Administratoren'; AccessRights = @('FullAccess'); IsInherited = $true; Deny = $false }
            [pscustomobject]@{ User = 'EXAMPLE\chefin'; AccessRights = @('FullAccess', 'ReadPermission'); IsInherited = $false; Deny = $false }
        }
        $report = @(Get-EobMailboxPermissionReport -Identity 'mmustermann' -Mode OnPremises)
        $report.Count | Should -Be 1
        $report[0].User | Should -Be 'EXAMPLE\chefin'
        $report[0].AccessRights | Should -Be 'FullAccess, ReadPermission'
    }
}

Describe 'Microsoft Graph' {
    It 'meldet die deaktivierte Anbindung' {
        (Get-EobGraphStatus -Config (New-TestConfig)).State | Should -Be 'Disabled'
    }

    It 'meldet fehlende Verbindung und fehlende Berechtigungen' {
        $config = New-TestConfig -Content "[Graph]`nEnabled=1"
        Mock -ModuleName $script:EntraModule Get-MgContext { $null }
        (Get-EobGraphStatus -Config $config).State | Should -Be 'NotConnected'
        Mock -ModuleName $script:EntraModule Get-MgContext { [pscustomobject]@{ Account = 'admin@example.com'; Scopes = @('User.Read.All') } }
        $status = Get-EobGraphStatus -Config $config
        $status.State | Should -Be 'Error'
        $status.Detail | Should -Match 'User.RevokeSessions.All'
        Mock -ModuleName $script:EntraModule Get-MgContext { [pscustomobject]@{ Account = 'admin@example.com'; Scopes = @('User.Read.All', 'User.RevokeSessions.All') } }
        (Get-EobGraphStatus -Config $config).State | Should -Be 'Connected'
    }

    It 'verbindet nur mit minimalen Berechtigungen' {
        Mock -ModuleName $script:EntraModule Connect-MgGraph { }
        $null = Connect-EobGraph -Config (New-TestConfig -Content "[Graph]`nEnabled=1`nTenantId=example.onmicrosoft.com") -Confirm:$false
        Should -Invoke -ModuleName $script:EntraModule Connect-MgGraph -Times 1 -ParameterFilter {
            (@($Scopes) -join ',') -eq 'User.Read.All,User.RevokeSessions.All' -and $TenantId -eq 'example.onmicrosoft.com' -and $NoWelcome
        }
    }

    It 'verweigert die Verbindung bei deaktivierter Anbindung' {
        { Connect-EobGraph -Config (New-TestConfig) -Confirm:$false } | Should -Throw '*deaktiviert*'
    }

    It 'widerruft Sitzungen nur mit Verbindung' {
        Mock -ModuleName $script:EntraModule Revoke-MgUserSignInSession { $true }
        Mock -ModuleName $script:EntraModule Get-MgContext { $null }
        { Revoke-EobEntraUserSession -UserPrincipalName 'max@example.com' -Confirm:$false } | Should -Throw '*Keine Verbindung*'
        Mock -ModuleName $script:EntraModule Get-MgContext { [pscustomobject]@{ Account = 'admin@example.com'; Scopes = @() } }
        (Revoke-EobEntraUserSession -UserPrincipalName 'max@example.com' -Confirm:$false).Status | Should -Be 'Succeeded'
        Should -Invoke -ModuleName $script:EntraModule Revoke-MgUserSignInSession -Times 1 -ParameterFilter { $UserId -eq 'max@example.com' }
    }

    It 'simuliert den Widerruf ohne Graph-Aufruf' {
        Mock -ModuleName $script:EntraModule Revoke-MgUserSignInSession { $true }
        (Revoke-EobEntraUserSession -UserPrincipalName 'max@example.com' -WhatIf).Message | Should -Match '^Simulation'
        Should -Invoke -ModuleName $script:EntraModule Revoke-MgUserSignInSession -Times 0
    }
}

Describe 'Entra Connect Sync' {
    It 'meldet Konfigurationszustände' {
        (Get-EobEntraConnectStatus -Config (New-TestConfig)).State | Should -Be 'Disabled'
        (Get-EobEntraConnectStatus -Config (New-TestConfig -Content "[ADSync]`nEnableADSync=1")).State | Should -Be 'NotConfigured'
        (Get-EobEntraConnectStatus -Config (New-TestConfig -Content "[ADSync]`nEnableADSync=1`nADSyncServer=sync01.example.local")).State | Should -Be 'Available'
    }

    It 'startet den festen Synchronisationsbefehl per Remoting' {
        Mock -ModuleName $script:EntraModule Invoke-Command { 'Result: Success' }
        $result = Start-EobEntraConnectSync -Server 'sync01.example.local' -PolicyType Delta -Confirm:$false
        $result.Status | Should -Be 'Succeeded'
        Should -Invoke -ModuleName $script:EntraModule Invoke-Command -Times 1 -ParameterFilter {
            $ComputerName -eq 'sync01.example.local' -and $ArgumentList[0] -eq 'Delta' -and $ScriptBlock.ToString() -match 'Start-ADSyncSyncCycle -PolicyType \$Type'
        }
    }

    It 'meldet einen bereits laufenden Zyklus als Warnung' {
        Mock -ModuleName $script:EntraModule Invoke-Command { throw 'Sync is already running. Connector busy.' }
        (Start-EobEntraConnectSync -Server 'sync01.example.local' -Confirm:$false).Status | Should -Be 'Warning'
    }

    It 'simuliert ohne Remoting' {
        Mock -ModuleName $script:EntraModule Invoke-Command { }
        (Start-EobEntraConnectSync -Server 'sync01.example.local' -WhatIf).Message | Should -Match '^Simulation'
        Should -Invoke -ModuleName $script:EntraModule Invoke-Command -Times 0
    }

    It 'lehnt ungültige Servernamen ab' {
        { Start-EobEntraConnectSync -Server 'sync01; Remove-Item C:\' -Confirm:$false } | Should -Throw '*Ungültiger Servername*'
    }
}

Describe 'Integrationsübersicht' {
    It 'enthält alle Adapter' {
        $names = @(Get-EobIntegrationStatus -Config (New-TestConfig)).Name
        foreach ($expected in @('Active Directory', 'Exchange', 'Microsoft Graph', 'Entra Connect Sync', 'SMTP', 'Dateiserver', 'PDF-Erzeugung')) {
            $names | Should -Contain $expected
        }
    }
}
