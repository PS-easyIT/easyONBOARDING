#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.3.0' }
<#
    Tests der Offboarding-Engine. Das AD ist vollständig gemockt (Adapter-Interna und Cmdlets);
    es werden weder Konten noch Gruppen verändert. Dateien (Snapshots, Warteschlange, Berichte)
    entstehen ausschließlich im TestDrive.
#>

BeforeAll {
    . (Join-Path -Path $PSScriptRoot -ChildPath '../TestHelpers.ps1')
    Import-EobTestModule

    $script:Off = 'easyONB.Offboarding'
    $script:Ad = 'easyONB.ActiveDirectory'
    $script:Guid = '11111111-2222-3333-4444-555555555555'
    $null = Initialize-EobLogging -Directory (Join-Path $TestDrive 'logs')

    $script:ConfigDir = Join-Path $TestDrive 'config'
    $null = New-Item -ItemType Directory -Path (Join-Path $script:ConfigDir 'templates') -Force
    Copy-Item -LiteralPath (Join-Path (Get-EobTestRepoRoot) 'Config/templates/offboarding-templates.ini') -Destination (Join-Path $script:ConfigDir 'templates')

    function script:New-TestConfig {
        param([string]$Extra = '')
        $id = [guid]::NewGuid().ToString('N')
        $root = Join-Path $TestDrive $id
        $ini = @"
[Company]
CompanyNameFirma=Example GmbH
CompanyActiveDirectoryDomain=example.com
CompanyMailDomain=@example.com
[MailEndungen]
Domain1=@example.com
[Security]
ServiceAccountPatterns=svc_*
[Offboarding]
DisabledUsersOU=OU=Ausgeschieden,DC=example,DC=local
KeepGroups=GRP-Behalten
[LicensesGroups]
KEINE=
MS365_E3=GRP-Lizenz-E3
[ActivateUserMS365ADSync]
ADSync=1
ADSyncADGroup=GRP-M365-Sync
[Paths]
DataDirectory=$(Join-Path $root 'data')
SnapshotDirectory=$(Join-Path $root 'snapshots')
[Report]
ReportPath=$(Join-Path $root 'reports')
$Extra
"@
        $path = Join-Path $script:ConfigDir "$id.ini"
        Set-Content -LiteralPath $path -Value $ini -Encoding utf8
        $config = Import-EobConfiguration -Path $path
        Add-Member -InputObject $config -NotePropertyName 'TestRoot' -NotePropertyValue $root -Force
        return $config
    }

    function script:New-TestUser {
        param(
            [string]$Sam = 'mmuster', [bool]$Enabled = $true, [string]$Sid = 'S-1-5-21-1-2-3-1500', [int]$AdminCount = 0,
            [string]$HomeDirectory = '', [string]$ProfilePath = '\\fs01.example.local\profiles$\mmuster',
            [string]$Manager = 'CN=Chefin Beispiel,OU=Mitarbeiter,DC=example,DC=local'
        )
        ConvertTo-EobAdUserInfo -AdUser ([pscustomobject]@{
                SamAccountName = $Sam; SID = $Sid; Enabled = $Enabled; DistinguishedName = 'CN=Max Muster,OU=Mitarbeiter,DC=example,DC=local'
                adminCount = $AdminCount; ObjectGUID = $script:Guid; ObjectClass = 'user'; DisplayName = 'Max Muster'; GivenName = 'Max'; Surname = 'Muster'
                UserPrincipalName = "$Sam@example.com"; mail = "$Sam@example.com"; primaryGroupID = 513; manager = $Manager; Description = 'Vertrieb Innendienst'
                HomeDirectory = $HomeDirectory; ProfilePath = $ProfilePath
            })
    }

    function script:New-TestMembership {
        $group = { param($Name, $Rid, $Category = 'Security', $Primary = $false)
            [pscustomobject]@{ Name = $Name; SamAccountName = $Name; DistinguishedName = "CN=$Name,OU=Gruppen,DC=example,DC=local"; SID = "S-1-5-21-1-2-3-$Rid"; AdminCount = 0; GroupCategory = $Category; IsPrimary = $Primary }
        }
        @(
            (& $group 'GRP-Vertrieb' 2001), (& $group 'VL-Alle' 2002 'Distribution'), (& $group 'GRP-Lizenz-E3' 2003),
            (& $group 'GRP-M365-Sync' 2004), (& $group 'GRP-Behalten' 2005), (& $group 'Domain Users' 513 'Security' $true)
        )
    }

    function script:Set-DirectoryMock {
        param([pscustomobject]$User = (New-TestUser), [string[]]$Privileged = @())
        $script:TestUser = $User
        $script:TestPrivileged = $Privileged
        Mock -ModuleName $script:Off Get-EobAdConnectionInfo {
            [pscustomobject]@{ Connected = $true; Server = 'dc01.example.local'; Domain = [pscustomobject]@{ NetBiosName = 'EXAMPLE'; DomainSid = 'S-1-5-21-1-2-3' } }
        }
        Mock -ModuleName $script:Off Resolve-EobAdUser {
            if ($Identity -like 'CN=Chefin*') { [pscustomobject]@{ SamAccountName = 'chefin'; DisplayName = 'Chefin Beispiel'; Mail = 'chefin@example.com'; DistinguishedName = $Identity } }
            else { $script:TestUser }
        }
        Mock -ModuleName $script:Off Get-EobAdPrivilegedMembership { $script:TestPrivileged }
        Mock -ModuleName $script:Off Get-EobAdDirectReport { @() }
        Mock -ModuleName $script:Off Get-EobAdUserGroupMembership { New-TestMembership }
        Mock -ModuleName $script:Off Test-EobAdOrganizationalUnit { $true }
        Mock -ModuleName $script:Off Test-EobAdRecycleBinEnabled { $true }
        Mock -ModuleName $script:Off Get-EobCurrentIdentity { [pscustomobject]@{ Name = 'EXAMPLE\admin'; Sid = 'S-1-5-21-1-2-3-4242' } }
        Mock -ModuleName $script:Off Get-EobAdUserSnapshot {
            [pscustomobject]@{ CapturedAt = (Get-Date).ToString('o'); Server = 'dc01.example.local'; User = $script:TestUser; Groups = @(New-TestMembership); Manager = $null; DirectReports = @() }
        }
        Mock -ModuleName $script:Ad Assert-EobAdConnected { }
        Mock -ModuleName $script:Ad Assert-EobAdTargetAllowed { $script:TestUser }
        Mock -ModuleName $script:Ad Test-EobAdOrganizationalUnit { $true }
        Mock -ModuleName $script:Ad Disable-ADAccount { }
        Mock -ModuleName $script:Ad Set-ADUser { }
        Mock -ModuleName $script:Ad Set-ADAccountExpiration { }
        Mock -ModuleName $script:Ad Set-ADAccountPassword { }
        Mock -ModuleName $script:Ad Remove-ADGroupMember { }
        Mock -ModuleName $script:Ad Move-ADObject { [pscustomobject]@{ DistinguishedName = 'CN=Max Muster,OU=Ausgeschieden,DC=example,DC=local' } }
        Mock -ModuleName $script:Ad Remove-ADObject { }
    }

    function script:New-TestRequest {
        param([hashtable]$Data = @{}, [pscustomobject]$Config)
        $values = @{ Identity = 'mmuster'; Template = 'Standard'; ExitDate = '2027-03-31'; Ticket = 'CHG-1001'; Reason = 'Eigenkündigung' }
        foreach ($key in $Data.Keys) { $values[$key] = $Data[$key] }
        return ConvertTo-EobOffboardingRequest -InputObject $values -Config $Config
    }

    $script:BeforeExit = [datetime]'2027-03-15 09:00'
}

AfterAll {
    Remove-EobRedactionValue -All
}

Describe 'Vorlagen und Anfragen' {
    It 'liest alle mitgelieferten Vorlagen' {
        $names = @(Get-EobOffboardingTemplate -Config (New-TestConfig)).Name
        foreach ($expected in @('Standard', 'SofortigeSperrung', 'Befristet', 'Extern', 'Ruhestand', 'InternerWechsel', 'Testmodus')) {
            $names | Should -Contain $expected
        }
    }

    It 'ordnet jeder Aktion eine Phase zu (Standard: Löschung nur in FinalDeletion)' {
        $template = Get-EobOffboardingTemplate -Config (New-TestConfig) -Name 'Standard'
        $template.Actions['DeleteAccount'] | Should -Be 'FinalDeletion'
        $template.Actions['DisableAccount'] | Should -Be 'ExitDate'
        $template.Actions['SetExpiration'] | Should -Be 'Immediate'
        $template.RetentionDays | Should -Be 180
    }

    It 'meldet Pflichtfelder und ungültige Angaben' {
        $config = New-TestConfig
        $request = ConvertTo-EobOffboardingRequest -InputObject @{ Identity = ''; Template = 'Gibtsnicht'; ExitDate = '31.02.2027' } -Config $config
        $codes = @(Test-EobOffboardingRequest -Request $request -Config $config).Code
        $codes | Should -Contain 'OFF_IDENTITY_MISSING'
        $codes | Should -Contain 'OFF_TEMPLATE_NOT_FOUND'
        $codes | Should -Contain 'OFF_EXITDATE_INVALID'
        $codes | Should -Contain 'OFF_TICKET_REQUIRED'
    }

    It 'verlangt keine Ticketnummer im Testmodus' {
        $config = New-TestConfig
        $request = New-TestRequest -Config $config -Data @{ Template = 'Testmodus'; Ticket = '' }
        @(Test-EobOffboardingRequest -Request $request -Config $config).Code | Should -Not -Contain 'OFF_TICKET_REQUIRED'
    }

    It 'lässt die Löschung nur in der Phase FinalDeletion zu' {
        $config = New-TestConfig
        $request = New-TestRequest -Config $config -Data @{ ActionPhases = @{ DeleteAccount = 'Immediate'; DisableAccount = 'FinalDeletion'; Unbekannt = 'Immediate' } }
        $codes = @(Test-EobOffboardingRequest -Request $request -Config $config).Code
        $codes | Should -Contain 'OFF_DELETE_PHASE'
        $codes | Should -Contain 'OFF_PHASE_FINALDELETION'
        $codes | Should -Contain 'OFF_ACTION_UNKNOWN'
    }

    It 'berechnet die Phasentermine (Austritt = letzter Arbeitstag)' {
        $template = Get-EobOffboardingTemplate -Config (New-TestConfig) -Name 'Standard'
        $dates = Get-EobOffboardingPhaseDate -ExitDate ([datetime]'2027-03-31') -Template $template -Now $script:BeforeExit
        $dates['ExitDate'] | Should -Be ([datetime]'2027-04-01')
        $dates['Retention'] | Should -Be ([datetime]'2027-04-30')
        $dates['FinalDeletion'] | Should -Be ([datetime]'2027-09-27')
    }
}

Describe 'Schutzprüfung' {
    It 'blockiert Dienstkonten ohne Schritte' {
        Set-DirectoryMock -User (New-TestUser -Sam 'svc_backup')
        $config = New-TestConfig
        $plan = New-EobOffboardingPlan -Request (New-TestRequest -Config $config -Data @{ Identity = 'svc_backup' }) -Config $config -Now $script:BeforeExit
        @($plan.Findings | Where-Object Code -EQ 'OFF_PROTECTED').Count | Should -BeGreaterThan 0
        $plan.Steps.Count | Should -Be 0
        Test-EobPlanExecutable -Plan $plan | Should -BeFalse
    }

    It 'blockiert das eigene Konto' {
        Set-DirectoryMock -User (New-TestUser -Sid 'S-1-5-21-1-2-3-4242')
        $config = New-TestConfig
        $plan = New-EobOffboardingPlan -Request (New-TestRequest -Config $config) -Config $config -Now $script:BeforeExit
        ($plan.Findings | Where-Object Code -EQ 'OFF_PROTECTED').Message | Should -Match 'eigene'
    }

    It 'blockiert privilegierte Konten, solange die Konfiguration es nicht erlaubt' {
        Set-DirectoryMock -Privileged @('Domain Admins')
        $config = New-TestConfig
        $plan = New-EobOffboardingPlan -Request (New-TestRequest -Config $config -Data @{ AcknowledgePrivileged = $true }) -Config $config -Now $script:BeforeExit
        @($plan.Findings.Code) | Should -Contain 'OFF_PRIVILEGED'
        $plan.Steps.Count | Should -Be 0
    }

    It 'verlangt für privilegierte Konten die gesonderte Bestätigung' {
        Set-DirectoryMock -Privileged @('Domain Admins')
        $config = New-TestConfig -Extra "[Security]`nAllowPrivilegedOffboarding=1"
        $plan = New-EobOffboardingPlan -Request (New-TestRequest -Config $config) -Config $config -Now $script:BeforeExit
        @($plan.Findings.Code) | Should -Contain 'OFF_PRIVILEGED_ACK_REQUIRED'
        $plan.Steps.Count | Should -Be 0
    }

    It 'plant privilegierte Konten nach Bestätigung mit AllowPrivileged' {
        Set-DirectoryMock -Privileged @('Domain Admins')
        $config = New-TestConfig -Extra "[Security]`nAllowPrivilegedOffboarding=1"
        $plan = New-EobOffboardingPlan -Request (New-TestRequest -Config $config -Data @{ AcknowledgePrivileged = $true }) -Config $config -Now $script:BeforeExit
        @($plan.Findings.Code) | Should -Contain 'OFF_PRIVILEGED_ACKNOWLEDGED'
        $disable = $plan.Steps | Where-Object Action -EQ 'DisableAccount'
        $disable.Parameters['AllowPrivileged'] | Should -BeTrue
        $plan.Offboarding.Privileged | Should -BeTrue
    }

    It 'blockiert, wenn keine AD-Verbindung besteht' {
        Mock -ModuleName $script:Off Get-EobAdConnectionInfo { [pscustomobject]@{ Connected = $false; Server = ''; Domain = $null } }
        $config = New-TestConfig
        $plan = New-EobOffboardingPlan -Request (New-TestRequest -Config $config) -Config $config -Now $script:BeforeExit
        @($plan.Findings.Code) | Should -Contain 'OFF_AD_OFFLINE'
    }
}

Describe 'Planinhalt' {
    BeforeEach {
        Set-DirectoryMock
        $script:Config = New-TestConfig
        $script:Plan = New-EobOffboardingPlan -Request (New-TestRequest -Config $script:Config) -Config $script:Config -Now $script:BeforeExit
    }

    It 'ist ohne blockierende Befunde ausführbar' {
        Test-EobPlanExecutable -Plan $script:Plan | Should -BeTrue
        $script:Plan.Subject['ObjectGuid'] | Should -Be $script:Guid
    }

    It 'sichert den Zustand vorher (kritisch) und nachher' {
        $script:Plan.Steps[0].Action | Should -Be 'SnapshotBefore'
        $script:Plan.Steps[0].Critical | Should -BeTrue
        $script:Plan.Steps[-1].Action | Should -Be 'SnapshotAfter'
    }

    It 'adressiert alle AD-Schritte über die ObjectGUID' {
        foreach ($step in @($script:Plan.Steps | Where-Object { $_.Parameters.ContainsKey('Identity') -and $_.Handler -like '*EobAd*' })) {
            $step.Parameters['Identity'] | Should -Be $script:Guid
        }
    }

    It 'ordnet die Aktionen den Phasen der Vorlage zu' {
        ($script:Plan.Steps | Where-Object Action -EQ 'SetExpiration').Phase | Should -Be 'Immediate'
        ($script:Plan.Steps | Where-Object Action -EQ 'DisableAccount').Phase | Should -Be 'ExitDate'
        ($script:Plan.Steps | Where-Object Action -EQ 'ClearProfilePath').Phase | Should -Be 'Retention'
    }

    It 'setzt das Ablaufdatum auf das Ende des Austrittstags' {
        ($script:Plan.Steps | Where-Object Action -EQ 'SetExpiration').Parameters['ExpiresAt'] | Should -Be ([datetime]'2027-04-01')
    }

    It 'erzeugt die Beschreibung aus der Vorlage' {
        $text = ($script:Plan.Steps | Where-Object Action -EQ 'UpdateDescription').Parameters['Replace']['description']
        $text | Should -Be 'Ausgetreten 31.03.2027 | Ticket CHG-1001 | EXAMPLE\admin'
    }

    It 'entfernt nie primäre Gruppe, Synchronisations-, Behalten- oder Lizenzgruppen über RemoveGroups' {
        $removed = @($script:Plan.Steps | Where-Object { $_.Action -eq 'RemoveGroups' } | ForEach-Object { $_.Parameters['Group'] })
        $removed | Should -Contain 'CN=GRP-Vertrieb,OU=Gruppen,DC=example,DC=local'
        $removed | Should -Contain 'CN=VL-Alle,OU=Gruppen,DC=example,DC=local'
        foreach ($kept in @('Domain Users', 'GRP-M365-Sync', 'GRP-Behalten', 'GRP-Lizenz-E3')) {
            $removed | Should -Not -Contain "CN=$kept,OU=Gruppen,DC=example,DC=local"
        }
        ($script:Plan.Steps | Where-Object Action -EQ 'RemoveLicenseGroups').Parameters['Group'] | Should -Be 'CN=GRP-Lizenz-E3,OU=Gruppen,DC=example,DC=local'
        ($script:Plan.Steps | Where-Object Action -EQ 'RemoveLicenseGroups').Phase | Should -Be 'Retention'
    }

    It 'verschiebt als letzte Aktion der Austrittsphase in die Austritts-OU' {
        $exitSteps = @($script:Plan.Steps | Where-Object Phase -EQ 'ExitDate')
        $exitSteps[-1].Action | Should -Be 'MoveToOU'
        $exitSteps[-1].Parameters['TargetPath'] | Should -Be 'OU=Ausgeschieden,DC=example,DC=local'
    }

    It 'setzt das Kennwort nur über einen Secret-Verweis zurück' {
        $reset = $script:Plan.Steps | Where-Object Action -EQ 'ResetPassword'
        $reset.Parameters['NewPassword'] | Should -Be '@Secret:OffboardingPassword'
        $script:Plan.Secrets['OffboardingPassword'] | Should -BeOfType [securestring]
    }

    It 'plant die Löschung nur deaktiviert und mit Termin' {
        $delete = $script:Plan.Steps | Where-Object Action -EQ 'DeleteAccount'
        $delete.Enabled | Should -BeFalse
        $delete.Phase | Should -Be 'FinalDeletion'
        $delete.DueDate | Should -Be ([datetime]'2027-09-27')
    }

    It 'meldet deaktivierte Postfach- und Graph-Funktionen' {
        @($script:Plan.Findings.Code) | Should -Contain 'OFF_EXCHANGE_DISABLED'
        @($script:Plan.Findings.Code) | Should -Contain 'OFF_GRAPH_DISABLED'
        @($script:Plan.Steps | Where-Object Handler -Like '*Mailbox*').Count | Should -Be 0
    }

    It 'fasst den Vorgang zusammen' {
        $script:Plan.Summary['Vorlage'] | Should -Be 'Standardaustritt'
        $script:Plan.Summary['Phase Austritt ab'] | Should -Be '01.04.2027'
        $script:Plan.Summary['Endgültige Löschung'] | Should -Match '27\.09\.2027'
        $script:Plan.Summary['Führungskraft'] | Should -Be 'Chefin Beispiel'
    }

    It 'ermittelt die fälligen Phasen' {
        Get-EobOffboardingDuePhase -Plan $script:Plan -Now $script:BeforeExit | Should -Be @('Immediate')
        Get-EobOffboardingDuePhase -Plan $script:Plan -Now ([datetime]'2027-04-01') | Should -Be @('Immediate', 'ExitDate')
        Get-EobOffboardingDuePhase -Plan $script:Plan -Now ([datetime]'2027-05-02') | Should -Be @('Immediate', 'ExitDate', 'Retention')
    }

    It 'entfernt bei RemoveGroupsMode=Selected nur ausgewählte Gruppen' {
        $request = New-TestRequest -Config $script:Config -Data @{ Template = 'InternerWechsel'; SelectedGroups = 'GRP-Vertrieb' }
        $plan = New-EobOffboardingPlan -Request $request -Config $script:Config -Now $script:BeforeExit
        @($plan.Steps | Where-Object Action -EQ 'RemoveGroups').Parameters.Group | Should -Be @('CN=GRP-Vertrieb,OU=Gruppen,DC=example,DC=local')
        @($plan.Steps | Where-Object Action -EQ 'DisableAccount').Count | Should -Be 0
        @($plan.Steps | Where-Object Action -EQ 'DeleteAccount').Count | Should -Be 0
    }

    It 'erzwingt die Simulation bei der Testvorlage' {
        $plan = New-EobOffboardingPlan -Request (New-TestRequest -Config $script:Config -Data @{ Template = 'Testmodus' }) -Config $script:Config -Now $script:BeforeExit
        $plan.Simulation | Should -BeTrue
        @($plan.Findings.Code) | Should -Contain 'OFF_TEST_TEMPLATE'
    }

    It 'plant Postfachaktionen bei verbundenem Exchange Online' {
        Mock -ModuleName $script:Off Test-EobExchangeConnected { $true }
        Mock -ModuleName $script:Off Get-EobMailboxInfo { [pscustomobject]@{ PrimarySmtpAddress = 'mmuster@example.com'; RecipientTypeDetails = 'UserMailbox' } }
        $config = New-TestConfig -Extra "[Exchange]`nMode=Online"
        $plan = New-EobOffboardingPlan -Request (New-TestRequest -Config $config -Data @{ Template = 'Ruhestand' }) -Config $config -Now $script:BeforeExit
        $forward = $plan.Steps | Where-Object Action -EQ 'SetForwarding'
        $forward.Parameters['ForwardTo'] | Should -Be 'chefin@example.com'
        $forward.Parameters['InternalDomains'] | Should -Contain 'example.com'
        ($plan.Steps | Where-Object Action -EQ 'SetAutoReply').Parameters['Message'] | Should -Match 'Max Muster ist nicht mehr im Unternehmen'
        ($plan.Steps | Where-Object Action -EQ 'HideFromAddressLists').Parameters['Replace']['msExchHideFromAddressLists'] | Should -BeTrue
        $convert = $plan.Steps | Where-Object Action -EQ 'ConvertToSharedMailbox'
        $license = $plan.Steps | Where-Object Action -EQ 'RemoveLicenseGroups'
        [array]::IndexOf(@($plan.Steps), $convert) | Should -BeLessThan ([array]::IndexOf(@($plan.Steps), $license))
        $plan.Offboarding.RequiresIntegration | Should -Contain 'ExitDate'
    }
}

Describe 'Ausführung und Warteschlange' {
    BeforeEach {
        Set-DirectoryMock
        $script:Config = New-TestConfig
    }

    It 'führt nur die fällige Sofortphase aus und legt den Vorgang in die Warteschlange' {
        $plan = New-EobOffboardingPlan -Request (New-TestRequest -Config $script:Config) -Config $script:Config -Now $script:BeforeExit
        $result = Invoke-EobOffboardingPlan -Plan $plan -Now $script:BeforeExit -Confirm:$false
        $result.Status | Should -Be 'Succeeded'
        Should -Invoke -ModuleName $script:Ad Set-ADAccountExpiration -Times 1 -ParameterFilter { $DateTime -eq [datetime]'2027-04-01' }
        Should -Invoke -ModuleName $script:Ad Set-ADUser -Times 1 -ParameterFilter { $Replace['description'] -like 'Ausgetreten 31.03.2027*' }
        Should -Invoke -ModuleName $script:Ad Disable-ADAccount -Times 0
        ($result.Steps | Where-Object Action -EQ 'DisableAccount').Status | Should -Be 'Planned'

        $queue = @(Get-EobOffboardingQueue -Config $script:Config -Now $script:BeforeExit)
        $queue.Count | Should -Be 1
        $queue[0].Status | Should -Be 'Open'
        $queue[0].NextPhase | Should -Be 'ExitDate'
        $queue[0].NextDueDate | Should -Be ([datetime]'2027-04-01')
        $queue[0].IsDue | Should -BeFalse
        @($queue[0].Entry['CompletedPhases']) | Should -Be @('Immediate')
        @($queue[0].Entry['SnapshotPaths']).Count | Should -Be 2
        Test-Path -LiteralPath @($queue[0].Entry['SnapshotPaths'])[0] | Should -BeTrue
        foreach ($snapshot in @($queue[0].Entry['SnapshotPaths'])) {
            $queue[0].Entry['SnapshotHashes'][$snapshot] | Should -Be (Get-FileHash -LiteralPath $snapshot -Algorithm SHA256).Hash
        }
    }

    It 'speichert weder Kennwörter noch Secrets in Warteschlange, Snapshots, Logs oder Berichten' {
        $plan = New-EobOffboardingPlan -Request (New-TestRequest -Config $script:Config -Data @{ Template = 'SofortigeSperrung' }) -Config $script:Config -Now $script:BeforeExit
        $secret = ConvertTo-EobPlainText -SecureString $plan.Secrets['OffboardingPassword']
        $result = Invoke-EobOffboardingPlan -Plan $plan -Now $script:BeforeExit -Confirm:$false
        Should -Invoke -ModuleName $script:Ad Set-ADAccountPassword -Times 1 -ParameterFilter { $NewPassword -is [securestring] -and $Reset }
        $null = Export-EobPlanReport -Plan $result -Config $script:Config -Format Html, Json, Csv, Txt -Confirm:$false
        $files = Get-ChildItem -LiteralPath $script:Config.TestRoot -Recurse -File
        $files.Count | Should -BeGreaterThan 3
        foreach ($file in $files) {
            (Get-Content -LiteralPath $file.FullName -Raw) | Should -Not -Match ([regex]::Escape($secret))
        }
        foreach ($log in Get-ChildItem -LiteralPath (Join-Path $TestDrive 'logs') -Recurse -File) {
            (Get-Content -LiteralPath $log.FullName -Raw) | Should -Not -Match ([regex]::Escape($secret))
        }
        (Get-Content -LiteralPath @(Get-EobOffboardingQueue -Config $script:Config)[0].Path -Raw) | Should -Not -Match 'Secret|Password"'
    }

    It 'simuliert ohne Änderungen und ohne Warteschlangeneintrag' {
        $plan = New-EobOffboardingPlan -Request (New-TestRequest -Config $script:Config -Data @{ Template = 'SofortigeSperrung' }) -Config $script:Config -Now $script:BeforeExit
        $result = Invoke-EobOffboardingPlan -Plan $plan -Now $script:BeforeExit -WhatIf
        $result.Status | Should -Be 'Simulated'
        Should -Invoke -ModuleName $script:Ad Disable-ADAccount -Times 0
        Should -Invoke -ModuleName $script:Ad Set-ADUser -Times 0
        @(Get-EobOffboardingQueue -Config $script:Config).Count | Should -Be 0
    }

    It 'verarbeitet Folgephasen aus der Warteschlange bis zur Löschfreigabe' {
        $plan = New-EobOffboardingPlan -Request (New-TestRequest -Config $script:Config) -Config $script:Config -Now $script:BeforeExit
        $null = Invoke-EobOffboardingPlan -Plan $plan -Now $script:BeforeExit -Confirm:$false

        $exitRun = @(Invoke-EobDueOffboardingPhase -Config $script:Config -Now ([datetime]'2027-04-01 06:00') -Confirm:$false)
        $exitRun.Count | Should -Be 1
        $exitRun[0].Phases | Should -Be @('ExitDate')
        $exitRun[0].Outcome | Should -Be 'Succeeded'
        Should -Invoke -ModuleName $script:Ad Disable-ADAccount -Times 1
        Should -Invoke -ModuleName $script:Ad Move-ADObject -Times 1 -ParameterFilter { $TargetPath -eq 'OU=Ausgeschieden,DC=example,DC=local' }
        Should -Invoke -ModuleName $script:Ad Remove-ADGroupMember -Times 2
        Should -Invoke -ModuleName $script:Ad Remove-ADGroupMember -Times 0 -ParameterFilter { $Identity -match 'GRP-M365-Sync|GRP-Behalten|GRP-Lizenz-E3|Domain Users' }

        $retentionRun = @(Invoke-EobDueOffboardingPhase -Config $script:Config -Now ([datetime]'2027-05-02') -Confirm:$false)
        $retentionRun[0].Phases | Should -Be @('Retention')
        Should -Invoke -ModuleName $script:Ad Remove-ADGroupMember -Times 1 -ParameterFilter { $Identity -eq 'CN=GRP-Lizenz-E3,OU=Gruppen,DC=example,DC=local' }
        Should -Invoke -ModuleName $script:Ad Set-ADUser -Times 1 -ParameterFilter { $Clear -contains 'profilePath' }

        $entry = @(Get-EobOffboardingQueue -Config $script:Config)[0]
        $entry.Status | Should -Be 'AwaitingDeletion'
        @($entry.Entry['CompletedPhases']) | Should -Be @('Immediate', 'ExitDate', 'Retention')
        @(Get-ChildItem -LiteralPath (Join-Path $script:Config.TestRoot 'reports') -Recurse -Filter '*.html').Count | Should -Be 2

        $noop = @(Invoke-EobDueOffboardingPhase -Config $script:Config -Now ([datetime]'2027-06-01') -Confirm:$false)
        $noop.Count | Should -Be 0
        Should -Invoke -ModuleName $script:Ad Remove-ADObject -Times 0
    }

    It 'meldet fällige Löschungen nur, statt sie auszuführen' {
        $plan = New-EobOffboardingPlan -Request (New-TestRequest -Config $script:Config -Data @{ Template = 'Extern' }) -Config $script:Config -Now $script:BeforeExit
        $null = Invoke-EobOffboardingPlan -Plan $plan -Now ([datetime]'2027-04-10') -Confirm:$false
        $result = @(Invoke-EobDueOffboardingPhase -Config $script:Config -Now ([datetime]'2027-05-15') -Confirm:$false)
        $result[0].Action | Should -Be 'DeletionDue'
        $result[0].Outcome | Should -Be 'ManualApprovalRequired'
        Should -Invoke -ModuleName $script:Ad Remove-ADObject -Times 0
    }

    It 'führt privilegierte Konten nicht unbeaufsichtigt weiter' {
        Set-DirectoryMock -Privileged @('Domain Admins')
        $config = New-TestConfig -Extra "[Security]`nAllowPrivilegedOffboarding=1"
        $plan = New-EobOffboardingPlan -Request (New-TestRequest -Config $config -Data @{ AcknowledgePrivileged = $true }) -Config $config -Now $script:BeforeExit
        $null = Invoke-EobOffboardingPlan -Plan $plan -Now $script:BeforeExit -Confirm:$false
        $result = @(Invoke-EobDueOffboardingPhase -Config $config -Now ([datetime]'2027-04-02') -Confirm:$false)
        $result[0].Outcome | Should -Be 'ManualRequired'
        Should -Invoke -ModuleName $script:Ad Disable-ADAccount -Times 0
    }

    It 'führt Phasen mit Exchange-Aktionen ohne Verbindung nicht unbeaufsichtigt aus' {
        Mock -ModuleName $script:Off Test-EobExchangeConnected { $true }
        Mock -ModuleName $script:Off Get-EobMailboxInfo { [pscustomobject]@{ PrimarySmtpAddress = 'mmuster@example.com' } }
        Mock -ModuleName $script:Off Get-EobMailboxPermissionReport { @() }
        $config = New-TestConfig -Extra "[Exchange]`nMode=Online"
        $plan = New-EobOffboardingPlan -Request (New-TestRequest -Config $config) -Config $config -Now $script:BeforeExit
        $null = Invoke-EobOffboardingPlan -Plan $plan -Now $script:BeforeExit -Confirm:$false
        Mock -ModuleName $script:Off Test-EobExchangeConnected { $false }
        $result = @(Invoke-EobDueOffboardingPhase -Config $config -Now ([datetime]'2027-04-02') -Confirm:$false)
        $result[0].Outcome | Should -Be 'ManualRequired'
        Should -Invoke -ModuleName $script:Ad Disable-ADAccount -Times 0
    }

    It 'führt Phasen mit Sitzungswiderruf ohne Graph-Verbindung nicht unbeaufsichtigt aus' {
        Mock -ModuleName $script:Off Get-EobGraphStatus { [pscustomobject]@{ State = 'NotConnected' } }
        $config = New-TestConfig -Extra "[Graph]`nEnabled=1"
        $plan = New-EobOffboardingPlan -Request (New-TestRequest -Config $config) -Config $config -Now $script:BeforeExit
        $plan.Offboarding.RequiresGraph | Should -Be @('ExitDate')
        $null = Invoke-EobOffboardingPlan -Plan $plan -Now $script:BeforeExit -Confirm:$false
        $result = @(Invoke-EobDueOffboardingPhase -Config $config -Now ([datetime]'2027-04-02') -AllowIntegration -Confirm:$false)
        $result[0].Outcome | Should -Be 'ManualRequired'
        $result[0].Message | Should -Match 'Microsoft Graph'
        Should -Invoke -ModuleName $script:Ad Disable-ADAccount -Times 0
    }

    It 'erzwingt beim Vorziehen von Phasen die Reihenfolge' {
        $plan = New-EobOffboardingPlan -Request (New-TestRequest -Config $script:Config) -Config $script:Config -Now $script:BeforeExit
        { Invoke-EobOffboardingPlan -Plan $plan -Phase ExitDate -Now $script:BeforeExit -Confirm:$false } | Should -Throw '*setzt die Phase*Immediate*'
        { Invoke-EobOffboardingPlan -Plan $plan -Phase Immediate, Retention -Now $script:BeforeExit -Confirm:$false } | Should -Throw '*setzt die Phase*ExitDate*'
        $result = Invoke-EobOffboardingPlan -Plan $plan -Phase Immediate, ExitDate -Now $script:BeforeExit -Confirm:$false
        Should -Invoke -ModuleName $script:Ad Disable-ADAccount -Times 1
        @($result.Offboarding.CompletedPhases) | Should -Be @('Immediate', 'ExitDate')
    }

    It 'überspringt defekte Warteschlangeneinträge, statt das Lesen abzubrechen' {
        $plan = New-EobOffboardingPlan -Request (New-TestRequest -Config $script:Config) -Config $script:Config -Now $script:BeforeExit
        $null = Invoke-EobOffboardingPlan -Plan $plan -Now $script:BeforeExit -Confirm:$false
        $directory = Split-Path -Parent @(Get-EobOffboardingQueue -Config $script:Config)[0].Path
        Set-Content -LiteralPath (Join-Path $directory '0-leer.json') -Value '' -NoNewline
        Set-Content -LiteralPath (Join-Path $directory '1-liste.json') -Value '[1, 2]'
        Set-Content -LiteralPath (Join-Path $directory '2-ohne-phasen.json') -Value '{"OperationId":"x","Status":"Open"}'
        $queue = @(Get-EobOffboardingQueue -Config $script:Config -WarningAction SilentlyContinue)
        $queue.Count | Should -Be 1
        $queue[0].OperationId | Should -Be $plan.OperationId
    }

    It 'bricht einen Vorgang in der Warteschlange ab' {
        $plan = New-EobOffboardingPlan -Request (New-TestRequest -Config $script:Config) -Config $script:Config -Now $script:BeforeExit
        $null = Invoke-EobOffboardingPlan -Plan $plan -Now $script:BeforeExit -Confirm:$false
        $null = Stop-EobOffboardingQueueEntry -Config $script:Config -OperationId $plan.OperationId -Reason 'Kündigung zurückgenommen' -Confirm:$false
        @(Get-EobOffboardingQueue -Config $script:Config)[0].Status | Should -Be 'Cancelled'
        @(Invoke-EobDueOffboardingPhase -Config $script:Config -Now ([datetime]'2027-04-02') -Confirm:$false).Count | Should -Be 0
        Should -Invoke -ModuleName $script:Ad Disable-ADAccount -Times 0
    }

    It 'bricht Folgephasen ab, wenn das Konto eine andere ObjectGUID hat' {
        $plan = New-EobOffboardingPlan -Request (New-TestRequest -Config $script:Config) -Config $script:Config -Now $script:BeforeExit
        $null = Invoke-EobOffboardingPlan -Plan $plan -Now $script:BeforeExit -Confirm:$false
        $script:TestUser = ConvertTo-EobAdUserInfo -AdUser ([pscustomobject]@{ SamAccountName = 'mmuster'; SID = 'S-1-5-21-1-2-3-9999'; ObjectGUID = '99999999-2222-3333-4444-555555555555'; Enabled = $true; DistinguishedName = 'CN=Max Muster,OU=Mitarbeiter,DC=example,DC=local'; ObjectClass = 'user' })
        $result = @(Invoke-EobDueOffboardingPhase -Config $script:Config -Now ([datetime]'2027-04-02') -Confirm:$false)
        $result[0].Outcome | Should -Be 'Blocked'
        $result[0].Message | Should -Match 'ObjectGUID'
        Should -Invoke -ModuleName $script:Ad Disable-ADAccount -Times 0
    }
}

Describe 'Endgültige Löschung' {
    BeforeEach {
        Set-DirectoryMock
        $script:Config = New-TestConfig
        $plan = New-EobOffboardingPlan -Request (New-TestRequest -Config $script:Config) -Config $script:Config -Now $script:BeforeExit
        $null = Invoke-EobOffboardingPlan -Plan $plan -Phase Immediate, ExitDate, Retention -Now $script:BeforeExit -Confirm:$false
        $script:OperationId = $plan.OperationId
        $script:TestUser = New-TestUser -Enabled $false
    }

    It 'verweigert die Löschung vor Ablauf der Aufbewahrungsfrist' {
        $deletion = New-EobOffboardingDeletionPlan -OperationId $script:OperationId -Config $script:Config -Now ([datetime]'2027-06-01')
        @($deletion.Findings.Code) | Should -Contain 'OFF_DEL_RETENTION_ACTIVE'
        { Invoke-EobOffboardingFinalDeletion -Plan $deletion -ConfirmationText 'LÖSCHEN mmuster' -Confirm:$false } | Should -Throw
        Should -Invoke -ModuleName $script:Ad Remove-ADObject -Times 0
    }

    It 'verweigert die Löschung aktivierter Konten' {
        $script:TestUser = New-TestUser -Enabled $true
        $deletion = New-EobOffboardingDeletionPlan -OperationId $script:OperationId -Config $script:Config -Now ([datetime]'2027-10-01')
        @($deletion.Findings.Code) | Should -Contain 'OFF_DEL_ENABLED'
    }

    It 'verweigert die Löschung ohne Zustandsbericht' {
        Get-ChildItem -LiteralPath (Join-Path $script:Config.TestRoot 'snapshots') -Recurse -File | Remove-Item -Force
        $deletion = New-EobOffboardingDeletionPlan -OperationId $script:OperationId -Config $script:Config -Now ([datetime]'2027-10-01')
        @($deletion.Findings.Code) | Should -Contain 'OFF_DEL_NO_BACKUP'
    }

    It 'verweigert die Löschung, wenn der Zustandsbericht verändert wurde' {
        $snapshot = @(Get-ChildItem -LiteralPath (Join-Path $script:Config.TestRoot 'snapshots') -Recurse -File -Filter '*_vorher_*')[0]
        Add-Content -LiteralPath $snapshot.FullName -Value ' '
        $deletion = New-EobOffboardingDeletionPlan -OperationId $script:OperationId -Config $script:Config -Now ([datetime]'2027-10-01')
        @($deletion.Findings.Code) | Should -Contain 'OFF_DEL_BACKUP_MODIFIED'
        Test-EobPlanExecutable -Plan $deletion | Should -BeFalse
    }

    It 'verlangt den exakten Bestätigungstext' {
        $deletion = New-EobOffboardingDeletionPlan -OperationId $script:OperationId -Config $script:Config -Now ([datetime]'2027-10-01')
        Test-EobPlanExecutable -Plan $deletion | Should -BeTrue
        $deletion.Summary['Bestätigungstext'] | Should -Be 'LÖSCHEN mmuster'
        { Invoke-EobOffboardingFinalDeletion -Plan $deletion -Confirm:$false } | Should -Throw '*LÖSCHEN mmuster*'
        { Invoke-EobOffboardingFinalDeletion -Plan $deletion -ConfirmationText 'löschen mmuster' -Confirm:$false } | Should -Throw
        Should -Invoke -ModuleName $script:Ad Remove-ADObject -Times 0
    }

    It 'simuliert die Löschung mit -WhatIf' {
        $deletion = New-EobOffboardingDeletionPlan -OperationId $script:OperationId -Config $script:Config -Now ([datetime]'2027-10-01')
        $result = Invoke-EobOffboardingFinalDeletion -Plan $deletion -WhatIf
        $result.Status | Should -Be 'Simulated'
        Should -Invoke -ModuleName $script:Ad Remove-ADObject -Times 0
        @(Get-EobOffboardingQueue -Config $script:Config)[0].Status | Should -Be 'AwaitingDeletion'
    }

    It 'löscht nach Freigabe mit Sicherung und GUID-Abgleich' {
        $deletion = New-EobOffboardingDeletionPlan -OperationId $script:OperationId -Config $script:Config -Now ([datetime]'2027-10-01')
        $deletion.Steps[0].Action | Should -Be 'SnapshotBeforeDeletion'
        $result = Invoke-EobOffboardingFinalDeletion -Plan $deletion -ConfirmationText 'LÖSCHEN mmuster' -Confirm:$false
        $result.Status | Should -Be 'Succeeded'
        Should -Invoke -ModuleName $script:Ad Remove-ADObject -Times 1 -ParameterFilter { $Identity -eq 'CN=Max Muster,OU=Mitarbeiter,DC=example,DC=local' }
        $entry = @(Get-EobOffboardingQueue -Config $script:Config)[0]
        $entry.Status | Should -Be 'Deleted'
        @($entry.Entry['SnapshotPaths'] | Where-Object { $_ -match 'vor-loeschung' }).Count | Should -Be 1
        @(Get-EobAuditEntry -OperationId $script:OperationId).Action | Should -Contain 'DeleteAccount'
    }

    It 'warnt bei deaktiviertem AD-Papierkorb und synchronisiertem Postfach' {
        Mock -ModuleName $script:Off Test-EobAdRecycleBinEnabled { $false }
        $config = New-TestConfig -Extra "[Exchange]`nMode=Hybrid"
        $plan = New-EobOffboardingPlan -Request (New-TestRequest -Config $config) -Config $config -Now $script:BeforeExit
        $null = Invoke-EobOffboardingPlan -Plan $plan -Phase Immediate, ExitDate, Retention -Now $script:BeforeExit -Confirm:$false
        $deletion = New-EobOffboardingDeletionPlan -OperationId $plan.OperationId -Config $config -Now ([datetime]'2027-10-01')
        @($deletion.Findings.Code) | Should -Contain 'OFF_DEL_NO_RECYCLEBIN'
        @($deletion.Findings.Code) | Should -Contain 'OFF_DEL_CLOUD_MAILBOX'
    }
}

Describe 'Home-Verzeichnis-Handler' {
    BeforeEach {
        $script:Root = Join-Path $TestDrive ('fs-' + [guid]::NewGuid().ToString('N'))
        $script:HomeDir = Join-Path $script:Root 'home/mmuster'
        $script:Archive = Join-Path $script:Root 'archiv'
        $null = New-Item -ItemType Directory -Path $script:HomeDir -Force
        Set-Content -LiteralPath (Join-Path $script:HomeDir 'datei.txt') -Value 'Inhalt'
    }

    It 'verschiebt das Verzeichnis in das Archiv' {
        $result = Move-EobHomeDirectoryToArchive -Path $script:HomeDir -ArchiveRoot $script:Archive -AllowedRoots $script:Root -SamAccountName 'mmuster' -Confirm:$false
        $result.Status | Should -Be 'Succeeded'
        Test-Path -LiteralPath $script:HomeDir | Should -BeFalse
        Test-Path -LiteralPath (Join-Path $result.Data['ArchivedHomeDirectory'] 'datei.txt') | Should -BeTrue
    }

    It 'verweigert Pfade außerhalb der Stammpfade und Stammpfade selbst' {
        { Move-EobHomeDirectoryToArchive -Path $script:HomeDir -ArchiveRoot $script:Archive -AllowedRoots (Join-Path $TestDrive 'anderes') -SamAccountName 'x' -Confirm:$false } | Should -Throw '*erlaubten Stammpfad*'
        { Move-EobHomeDirectoryToArchive -Path $script:Root -ArchiveRoot $script:Archive -AllowedRoots $script:Root -SamAccountName 'x' -Confirm:$false } | Should -Throw '*Stammpfad selbst*'
        Test-Path -LiteralPath $script:HomeDir | Should -BeTrue
    }

    It 'simuliert ohne Verschieben' {
        $result = Move-EobHomeDirectoryToArchive -Path $script:HomeDir -ArchiveRoot $script:Archive -AllowedRoots $script:Root -SamAccountName 'mmuster' -WhatIf
        $result.Message | Should -Match '^Simulation'
        Test-Path -LiteralPath $script:HomeDir | Should -BeTrue
    }

    It 'plant Archivierung und Austragen des Home-Verzeichnisses' {
        Set-DirectoryMock -User (New-TestUser -HomeDirectory $script:HomeDir)
        $config = New-TestConfig -Extra "[FileServer]`nAllowedRoots=$($script:Root)`nArchiveRoot=$($script:Archive)"
        $plan = New-EobOffboardingPlan -Request (New-TestRequest -Config $config) -Config $config -Now $script:BeforeExit
        $archive = $plan.Steps | Where-Object Action -EQ 'ArchiveHomeDirectory' | Select-Object -First 1
        $archive.Handler | Should -Be 'Move-EobHomeDirectoryToArchive'
        $clear = $plan.Steps | Where-Object { $_.Action -eq 'ArchiveHomeDirectory' -and $_.Handler -eq 'Set-EobAdUserAttribute' }
        $clear.DependsOn | Should -Be @($archive.Id)
        $clear.Parameters['Clear'] | Should -Be @('homeDirectory', 'homeDrive')
    }
}

Describe 'Massen-Offboarding' {
    BeforeEach {
        Set-DirectoryMock
        $script:Config = New-TestConfig
    }

    It 'liest die CSV und erkennt Duplikate' {
        $csv = Join-Path $TestDrive 'off.csv'
        Set-Content -LiteralPath $csv -Encoding utf8 -Value "SamAccountName;Austrittsdatum;Ticket;Grund`nmmuster;31.03.2027;CHG-1;Test`nmmuster;31.03.2027;CHG-2;Test"
        $import = Import-EobOffboardingCsv -Path $csv -Config $script:Config
        @($import.Findings | Where-Object Severity -EQ 'Error').Count | Should -Be 0
        $batch = New-EobOffboardingBatch -Import $import -Config $script:Config -Now $script:BeforeExit
        $batch.Items[0].State | Should -BeIn @('Ready', 'Warning')
        $batch.Items[1].State | Should -Be 'Error'
        @($batch.Items[1].Plan.Findings.Code) | Should -Contain 'CSV_DUPLICATE_USER'
    }

    It 'meldet fehlende Pflichtspalten' {
        $csv = Join-Path $TestDrive 'off2.csv'
        Set-Content -LiteralPath $csv -Encoding utf8 -Value "Name;Ticket`nMax;1"
        $codes = @((Import-EobOffboardingCsv -Path $csv -Config $script:Config).Findings.Code)
        $codes | Should -Contain 'CSV_MISSING_COLUMN'
        $codes | Should -Contain 'CSV_UNKNOWN_COLUMN'
    }

    It 'führt den Stapel aus und exportiert das Ergebnis mit Formelschutz' {
        $csv = Join-Path $TestDrive 'off3.csv'
        Set-Content -LiteralPath $csv -Encoding utf8 -Value "SamAccountName;ExitDate;Ticket;Reason`nmmuster;2027-03-31;=CHG-1;Test"
        $batch = New-EobOffboardingBatch -Import (Import-EobOffboardingCsv -Path $csv -Config $script:Config) -Config $script:Config -Now $script:BeforeExit
        $batch.ExecutableCount | Should -Be 1
        $null = Invoke-EobOffboardingBatch -Batch $batch -Now $script:BeforeExit -Confirm:$false
        $batch.Items[0].Outcome | Should -Be 'Succeeded'
        Should -Invoke -ModuleName $script:Ad Set-ADAccountExpiration -Times 1
        $out = Export-EobOffboardingBatchResult -Batch $batch -Path (Join-Path $TestDrive 'ergebnis.csv') -Confirm:$false
        (Import-Csv -LiteralPath $out -Delimiter ';')[0].Ticket | Should -BeExactly '''=CHG-1'
    }
}
