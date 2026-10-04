#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.3.0' }
<#
    Tests der Oberfläche.
    - Plattformunabhängig: XAML-Regeln, Ressourcen, Farbschemata, Abgleich der im Code verwendeten
      Steuerelementnamen und Datenbindungen mit den XAML-Dateien, reine Hilfsfunktionen.
    - Nur Windows mit STA-Thread (z. B. GitHub-Runner windows-latest): echtes Laden aller XAML-Dateien
      über WPF und Initialisieren jeder Ansicht ohne Anzeige des Fensters.
#>

BeforeDiscovery {
    $script:CanLoadWpf = $IsWindows -and [System.Threading.Thread]::CurrentThread.GetApartmentState() -eq [System.Threading.ApartmentState]::STA
    $guiRoot = Join-Path (Split-Path -Parent (Split-Path -Parent $PSScriptRoot)) 'GUI'
    $script:XamlFiles = @(Get-ChildItem -Path $guiRoot -Filter '*.xaml' -Recurse | ForEach-Object {
            @{ Name = $_.FullName.Substring($guiRoot.Length + 1).Replace('\', '/'); Path = $_.FullName }
        })
}

BeforeAll {
    . (Join-Path -Path $PSScriptRoot -ChildPath '../TestHelpers.ps1')
    Import-EobTestModule -Exclude @()

    $script:GuiRoot = Join-Path (Get-EobTestRepoRoot) 'GUI'
    $script:UiSource = Get-Content -LiteralPath (Join-Path (Get-EobTestRepoRoot) 'Modules/UI/easyONB.UI.psm1') -Raw
    $script:ControlFiles = @(Get-ChildItem -Path $script:GuiRoot -Filter '*.xaml' -Recurse | Where-Object { $_.FullName -notmatch '[\\/]Styles[\\/]' })
    $script:AllNames = @(foreach ($file in $script:ControlFiles) { Get-EobXamlElementName -Xaml (Get-Content -LiteralPath $file.FullName -Raw) })
    $script:ResourceKeys = @(foreach ($file in @('Styles/Theme.Light.xaml', 'Styles/Controls.xaml')) {
            [regex]::Matches((Get-Content -LiteralPath (Join-Path $script:GuiRoot $file) -Raw), 'x:Key="([^"]+)"') | ForEach-Object { $_.Groups[1].Value }
        })
}

Describe 'XAML-Dateien' {
    It '<Name> ist wohlgeformt, ohne x:Class und ohne Ereignisattribute' -ForEach $script:XamlFiles {
        $problems = @(Test-EobXamlContent -Xaml (Get-Content -LiteralPath $Path -Raw))
        $problems | Should -BeNullOrEmpty
    }

    It '<Name> verwendet nur definierte Ressourcen' -ForEach $script:XamlFiles {
        $text = Get-Content -LiteralPath $Path -Raw
        $used = @([regex]::Matches($text, '\{(?:Static|Dynamic)Resource (?!\{)([A-Za-z0-9_.]+)\}') | ForEach-Object { $_.Groups[1].Value } | Select-Object -Unique)
        @($used | Where-Object { $_ -notin $script:ResourceKeys }) | Should -BeNullOrEmpty
    }

    It '<Name> enthält keine fest codierten Farben' -ForEach @($script:XamlFiles | Where-Object { $_.Name -notlike 'Styles/*' }) {
        $text = Get-Content -LiteralPath $Path -Raw
        # Nur echte Farbattribute (Leerzeichen davor), nicht z. B. LastChildFill.
        [regex]::Matches($text, '(?<=\s)(?:Background|Foreground|BorderBrush|Fill|Stroke)="(?!\{)(?!Transparent")([^"]+)"').Count | Should -Be 0
    }

    It 'hat eindeutige Steuerelementnamen über alle Fenster, Ansichten und Dialoge' {
        @($script:AllNames | Group-Object | Where-Object Count -GT 1 | ForEach-Object Name) | Should -BeNullOrEmpty
    }

    It 'definiert in beiden Farbschemata dieselben Schlüssel' {
        $light = [regex]::Matches((Get-Content -LiteralPath (Join-Path $script:GuiRoot 'Styles/Theme.Light.xaml') -Raw), 'x:Key="([^"]+)"') | ForEach-Object { $_.Groups[1].Value }
        $dark = [regex]::Matches((Get-Content -LiteralPath (Join-Path $script:GuiRoot 'Styles/Theme.Dark.xaml') -Raw), 'x:Key="([^"]+)"') | ForEach-Object { $_.Groups[1].Value }
        Compare-Object -ReferenceObject @($light) -DifferenceObject @($dark) | Should -BeNullOrEmpty
    }

    It 'erkennt Verstöße gegen die XAML-Regeln' {
        $bad = '<Grid xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" x:Class="A.B"><Button x:Name="A" Click="Do" /><TextBox x:Name="A" /></Grid>'
        $problems = @(Test-EobXamlContent -Xaml $bad)
        ($problems -join ' ') | Should -Match 'x:Class'
        ($problems -join ' ') | Should -Match 'Click'
        ($problems -join ' ') | Should -Match "'A'"
        @(Test-EobXamlContent -Xaml '<Grid') | Should -Match 'wohlgeformt'
    }
}

Describe 'Abgleich Code und XAML' {
    It 'verwendet im Code nur vorhandene Steuerelemente' {
        $referenced = [System.Collections.Generic.HashSet[string]]::new()
        foreach ($pattern in @('\$c\.([A-Z][A-Za-z0-9]+)', '\$script:Ui\.C\.([A-Z][A-Za-z0-9]+)')) {
            foreach ($match in [regex]::Matches($script:UiSource, $pattern)) { $null = $referenced.Add($match.Groups[1].Value) }
        }
        # Konfigurationsschlüssel (-Key 'SimulationByDefault' usw.) sind keine Steuerelemente.
        $configKeys = @([regex]::Matches($script:UiSource, "-Key\s+'([A-Za-z0-9]+)'") | ForEach-Object { $_.Groups[1].Value } | Select-Object -Unique)
        $prefixes = 'Onb|Off|Upd|Bulk|Reports|Audit|Tools|Set|Info|Dash|Nav|Dlg|Cred|Header|Status|Mode|Theme|Simulation|Main'
        foreach ($match in [regex]::Matches($script:UiSource, "'((?:$prefixes)[A-Z][A-Za-z0-9]*)'")) {
            $name = $match.Groups[1].Value
            if ($name -notmatch '(Brush|Style)$' -and $name -notin $configKeys) { $null = $referenced.Add($name) }
        }
        foreach ($name in @('Keys', 'Values', 'Count')) { $null = $referenced.Remove($name) }
        @($referenced | Where-Object { $_ -notin $script:AllNames } | Sort-Object) | Should -BeNullOrEmpty
        $referenced.Count | Should -BeGreaterThan 150
    }

    It 'findet die nummerierten Schritt-Steuerelemente der Assistenten' {
        foreach ($index in 1..8) { $script:AllNames | Should -Contain "OnbPanel$index"; $script:AllNames | Should -Contain "OnbStep$index" }
        foreach ($index in 1..4) { $script:AllNames | Should -Contain "OffPanel$index"; $script:AllNames | Should -Contain "OffStep$index" }
    }

    It 'verwendet in Datenbindungen nur Eigenschaften, die der Code erzeugt' {
        $paths = foreach ($file in $script:ControlFiles) {
            [regex]::Matches((Get-Content -LiteralPath $file.FullName -Raw), 'Binding="\{Binding ([A-Za-z]+)\}"') | ForEach-Object { $_.Groups[1].Value }
        }
        $missing = @($paths | Select-Object -Unique | Where-Object { $script:UiSource -notmatch "\b$_\s*=" })
        $missing | Should -BeNullOrEmpty
    }

    It 'verweist nur auf vorhandene Ansichten und Funktionen' {
        InModuleScope 'easyONB.UI' {
            foreach ($name in $script:ViewDefinitions.Keys) {
                $definition = $script:ViewDefinitions[$name]
                Test-Path -LiteralPath (Join-Path $script:GuiRoot $definition.File) | Should -BeTrue
                Get-Command -Name $definition.Initialize -ErrorAction SilentlyContinue | Should -Not -BeNullOrEmpty
                if ($definition.Refresh) { Get-Command -Name $definition.Refresh -ErrorAction SilentlyContinue | Should -Not -BeNullOrEmpty }
            }
        }
        foreach ($match in [regex]::Matches($script:UiSource, "-On(?:Progress|Complete) '([A-Za-z-]+)'")) {
            InModuleScope 'easyONB.UI' -Parameters @{ Name = $match.Groups[1].Value } { param($Name) Get-Command -Name $Name -ErrorAction SilentlyContinue | Should -Not -BeNullOrEmpty }
        }
    }

    It 'enthält keine Ereignis-Handler ohne Fehlerbehandlung für Aktionen mit Fachlogik' {
        $handlers = [regex]::Matches($script:UiSource, '\.Add_(Click|SelectionChanged|MouseDoubleClick|Tick)\(\{([^\n]*)')
        foreach ($handler in $handlers) {
            $body = $handler.Groups[2].Value
            if ($body -match 'Invoke-EobUiSafely' -or $body.Trim() -eq '' -or $body -match 'DialogResult = \$true' -or $body -match '^\s*\$script:Ui\.Bulk\.Cancel') { continue }
            throw "Handler ohne Invoke-EobUiSafely: $($handler.Value)"
        }
    }
}

Describe 'Hilfsfunktionen der Oberfläche' {
    It 'wählt die Vordergrundfarbe nach WCAG-Kontrast' {
        (Get-EobAccentPalette -Color '#0F6CBD' -Theme Light).Foreground | Should -Be '#FFFFFF'
        (Get-EobAccentPalette -Color '#FFD700' -Theme Light).Foreground | Should -Be '#000000'
        (Get-EobAccentPalette -Color '#0F6CBD' -Theme Light).Contrast | Should -BeGreaterOrEqual 4.5
    }

    It 'hellt dunkle Akzentfarben im dunklen Schema auf' {
        $dark = Get-EobAccentPalette -Color '#102040' -Theme Dark
        $dark.Accent | Should -Not -Be '#102040'
        { Get-EobAccentPalette -Color 'blau' } | Should -Throw
    }

    It 'zeigt im Testmodus keine modalen Dialoge an' {
        InModuleScope 'easyONB.UI' {
            $script:Ui = New-EobUiState
            $script:Ui.Headless = $true
            (Show-EobDialog -Title 'Titel' -Message 'Meldung' -Details @('Detail') -Kind Warning).Confirmed | Should -BeFalse
            $script:Ui.DialogLog[0] | Should -Be '[Warning] Titel: Meldung Detail'
            Show-EobCredentialDialog -SamAccountName 'mmuster' -Password (New-Object System.Security.SecureString)
            $script:Ui.DialogLog[1] | Should -Match 'mmuster'
        }
    }

    It 'verhindert Live-Modus und Neuladen bei Fehlern in der Konfiguration' {
        $missing = Join-Path $TestDrive 'nicht-vorhanden.ini'
        $null = Initialize-EobLogging -Directory (Join-Path $TestDrive 'logs-config')
        InModuleScope 'easyONB.UI' -Parameters @{ Missing = $missing } {
            param($Missing)
            $script:Ui = New-EobUiState
            $script:Ui.Headless = $true
            $script:Ui.Config = Import-EobConfiguration -Path $Missing -NoIncludes
            (Get-EobUiConfigError -Config $script:Ui.Config) -join ' ' | Should -Match 'CFG_FILE_NOT_FOUND'
            Switch-EobExecutionMode
            $script:Ui.Simulation | Should -BeTrue
            $script:Ui.DialogLog[0] | Should -Match 'Live-Modus nicht möglich'

            $previous = [pscustomobject]@{ Findings = @(); Path = 'bisher' }
            $script:Ui.Config = $previous
            $script:Ui.ConfigPath = $Missing
            Update-EobUiConfiguration
            $script:Ui.Config.Path | Should -Be 'bisher'
            $script:Ui.DialogLog[1] | Should -Match 'nicht übernommen'
        }
    }

    It 'gibt nach einem Fehler in der Massenverarbeitung die Ausführungssperre frei und verwirft Kennwörter' {
        $null = Initialize-EobLogging -Directory (Join-Path $TestDrive 'logs-bulk')
        InModuleScope 'easyONB.UI' {
            $new = { param([string]$Property) [pscustomobject]@{ $Property = $null } }
            $script:Ui = New-EobUiState
            $script:Ui.Headless = $true
            $c = @{ BulkProgressBar = & $new 'Value'; BulkProgressText = & $new 'Text'; ModeToggleButton = & $new 'IsEnabled' }
            foreach ($name in 'BulkCancelButton', 'BulkImportButton', 'BulkExportButton', 'BulkModeOnboardingRadio', 'BulkModeOffboardingRadio') { $c[$name] = & $new 'IsEnabled' }
            foreach ($view in $script:ViewDefinitions.Values) { $c[$view.Nav] = & $new 'IsEnabled' }
            $script:Ui.C = $c
            $timer = [pscustomobject]@{ Running = $false }
            $timer | Add-Member -MemberType ScriptMethod -Name Stop -Value { $this.Running = $false }
            $timer | Add-Member -MemberType ScriptMethod -Name Start -Value { $this.Running = $true }
            $plans = foreach ($index in 1..2) {
                $plan = New-EobPlan -Kind 'Onboarding' -Context (New-EobOperationContext -Kind 'Onboarding' -Simulation) -Simulation
                $plan.Secrets['InitialPassword'] = New-EobPassword
                Add-EobRedactionValue -Value $plan.Secrets['InitialPassword']
                $plan
            }
            $plain = ConvertTo-EobPlainText -SecureString $plans[1].Secrets['InitialPassword']
            $items = @($plans | ForEach-Object { [pscustomobject]@{ RowNumber = 2; State = 'Error'; Plan = $_; Messages = @() } })
            $script:Ui.Bulk = @{ Batch = [pscustomobject]@{ Items = $items; OperationId = 'op-bulk' }; Mode = 'Onboarding'; Index = 0; Timer = $timer; Cancel = $false; Credentials = [System.Collections.Generic.List[object]]::new() }
            $script:Ui.Execution = @{ Plan = $null; Timer = $timer; Bulk = $true }
            Mock Update-EobBulkGrid { throw 'Anzeige fehlgeschlagen' }

            { Invoke-EobBulkTick } | Should -Throw 'Anzeige fehlgeschlagen'
            $script:Ui.Execution | Should -BeNullOrEmpty
            $timer.Running | Should -BeFalse
            $c.BulkModeOnboardingRadio.IsEnabled | Should -BeTrue
            @($items | ForEach-Object Outcome) | Should -Be @('NotExecuted', 'NotExecuted')
            @($plans | Where-Object { $_.Secrets.Count -gt 0 }).Count | Should -Be 0
            Protect-EobSensitiveText -Text $plain | Should -Be $plain
        }
    }

    It 'filtert das Audit-Log ohne Platzhaltersyntax' {
        $null = Initialize-EobLogging -Directory (Join-Path $TestDrive 'logs-auditfilter')
        Write-EobLog -Level Audit -Action 'Test' -Target 'konto[1]' -Message 'Filtertest'
        Write-EobLog -Level Audit -Action 'Test' -Target 'anderes' -Message 'Filtertest'
        InModuleScope 'easyONB.UI' {
            $script:Ui = New-EobUiState
            $script:Ui.Headless = $true
            $script:Ui.C = @{
                AuditFilterText = [pscustomobject]@{ Text = 'KONTO[' }; AuditFromPicker = [pscustomobject]@{ SelectedDate = (Get-Date).AddDays(-1) }
                AuditToPicker = [pscustomobject]@{ SelectedDate = (Get-Date) }; AuditGrid = [pscustomobject]@{ ItemsSource = $null }; AuditSummaryText = [pscustomobject]@{ Text = '' }
            }
            { Update-EobAuditList } | Should -Not -Throw
            @($script:Ui.C.AuditGrid.ItemsSource).Target | Should -Be @('konto[1]')
        }
    }

    It 'meldet die Oberfläche außerhalb von Windows als nicht verfügbar' -Skip:$IsWindows {
        $result = Test-EobGuiEnvironment
        $result.IsSupported | Should -BeFalse
        $result.Reason | Should -Match 'Windows'
    }

    It 'übergibt automatisch ermittelte Kontonamen nicht als feste Vorgabe' {
        InModuleScope 'easyONB.UI' {
            $values = @{ GivenName = 'Max'; Surname = 'Muster'; SamAccountName = 'mmuster'; UserPrincipalNamePrefix = 'max.eigen'; MailLocalPart = ''; StartDate = [datetime]'2027-01-04' }
            $data = Get-EobOnboardingFormRequestData -Values $values -AutoIdentity @{ SamAccountName = 'mmuster'; UserPrincipalNamePrefix = 'max.muster' }
            $data.ContainsKey('SamAccountName') | Should -BeFalse
            $data['UserPrincipalNamePrefix'] | Should -Be 'max.eigen'
            $data.ContainsKey('MailLocalPart') | Should -BeFalse
            $data['StartDate'] | Should -Be '2027-01-04'
        }
    }

    It 'übersetzt das Offboarding-Formular' {
        InModuleScope 'easyONB.UI' {
            $data = Get-EobOffboardingFormRequestData -Values @{ Identity = 'guid'; Template = 'Standard'; ExitDate = [datetime]'2027-03-31'; AssetsText = "Notebook 123`r`n`r`nHandy 456 "; AcknowledgePrivileged = $true }
            $data['ExitDate'] | Should -Be '2027-03-31'
            $data['Assets'] | Should -Be @('Notebook 123', 'Handy 456')
            $data['AcknowledgePrivileged'] | Should -BeTrue
        }
    }

    It 'bereitet Schritte, Befunde und Warteschlangeneinträge für die Anzeige auf' {
        InModuleScope 'easyONB.UI' {
            $step = [pscustomobject]@{ Id = 'S01'; Phase = 'ExitDate'; DueDate = [datetime]'2027-04-01'; Title = 'Konto deaktivieren'; Target = 'mmuster'; Risk = 'High'; Details = @('a', 'b'); Status = 'Planned'; Message = ''; Enabled = $true }
            $row = ConvertTo-EobUiStepRow -Step $step
            $row.PhaseText | Should -Be 'Zum Austritt'
            $row.DueText | Should -Be '01.04.2027'
            $row.RiskText | Should -Be 'Hoch'
            $row.DetailsText | Should -Be 'a | b'
            $step.Enabled = $false
            (ConvertTo-EobUiStepRow -Step $step).StatusText | Should -Be 'Nur nach manueller Freigabe'
            (ConvertTo-EobUiFindingRow -Finding ([pscustomobject]@{ Severity = 'Warning'; Code = 'X'; Field = 'F'; Message = 'M' })).SeverityText | Should -Be 'Warnung'
            $queue = [pscustomobject]@{ OperationId = '1'; DisplayName = 'Max'; SamAccountName = 'mmuster'; Template = 'Standard'; ExitDate = [datetime]'2027-03-31'; Status = 'AwaitingDeletion'
                NextPhase = ''; NextDueDate = $null; DeletionDate = [datetime]'2027-09-27'; IsDue = $false; DeletionDue = $true; Privileged = $false; RequiresIntegration = @()
            }
            $queueRow = ConvertTo-EobUiQueueRow -Item $queue
            $queueRow.Account | Should -Be 'Max (mmuster)'
            $queueRow.NoteText | Should -Match 'Löschung fällig'
            $queueRow.StatusText | Should -Be 'Wartet auf Löschfreigabe'
        }
    }

    It 'formatiert ISO-Zeitstempel des Audit-Logs' {
        InModuleScope 'easyONB.UI' {
            ConvertTo-EobUiDateText -Value '2027-01-04' | Should -Be '04.01.2027'
            ConvertTo-EobUiDateText -Value ([datetime]'2027-01-04 13:05') -WithTime | Should -Be '04.01.2027 13:05'
            ConvertTo-EobUiDateText -Value '' | Should -Be ''
            ConvertTo-EobUiDateText -Value $null | Should -Be ''
            ConvertTo-EobUiDateText -Value '2027-01-04T13:05:00.0000000+00:00' -WithTime | Should -Match '^04\.01\.2027 \d\d:05$'
        }
    }
}

Describe 'WPF-Ladeprüfung (Windows)' -Tag 'Windows' {
    BeforeAll {
        $script:TestConfigPath = Join-Path $TestDrive 'ui.ini'
        $template = Get-Content -LiteralPath (Join-Path (Get-EobTestRepoRoot) 'Config/easyONB.ini.template') -Raw
        $template = $template -replace '(?m)^LogFile=.*$', ('LogFile=' + (Join-Path $TestDrive 'logs')) -replace '(?m)^ReportPath=.*$', ('ReportPath=' + (Join-Path $TestDrive 'reports'))
        Set-Content -LiteralPath $script:TestConfigPath -Value $template -Encoding utf8
        $script:UiConfig = Import-EobConfiguration -Path $script:TestConfigPath -NoIncludes
        $null = Initialize-EobLogging -Directory (Join-Path $TestDrive 'logs') -EnableQueue
    }

    It 'lädt alle XAML-Dateien über WPF' -Skip:(-not $script:CanLoadWpf) {
        InModuleScope 'easyONB.UI' -Parameters @{ Config = $script:UiConfig } {
            param($Config)
            Import-EobWpfAssembly
            $script:Ui = New-EobUiState
            $script:Ui.Headless = $true
            $script:Ui.Config = $Config
            Initialize-EobApplication -Theme Light -AccentColor '#0F6CBD'
            foreach ($file in @('MainWindow.xaml', 'Dialogs/ConfirmDialog.xaml', 'Dialogs/CredentialDialog.xaml') + @($script:ViewDefinitions.Values | ForEach-Object File)) {
                $loaded = Import-EobXamlWithControl -RelativePath $file
                $loaded.Root | Should -Not -BeNullOrEmpty
                $loaded.Controls.Count | Should -BeGreaterThan 0
            }
            Set-EobTheme -Theme Dark -AccentColor '#0F6CBD'
            $script:Ui.Theme | Should -Be 'Dark'
            @($script:Ui.DialogLog) | Should -BeNullOrEmpty
        }
    }

    It 'initialisiert Hauptfenster und alle Ansichten ohne Anzeige' -Skip:(-not $script:CanLoadWpf) {
        InModuleScope 'easyONB.UI' -Parameters @{ Config = $script:UiConfig; Path = $script:TestConfigPath } {
            param($Config, $Path)
            Import-EobWpfAssembly
            $script:Ui = New-EobUiState
            $script:Ui.Headless = $true
            $script:Ui.Config = $Config
            $script:Ui.ConfigPath = $Path
            Initialize-EobApplication -Theme Light
            $loaded = Import-EobXamlWithControl -RelativePath 'MainWindow.xaml'
            $script:Ui.Window = $loaded.Root
            $script:Ui.C = $loaded.Controls
            Initialize-EobMainWindow
            foreach ($name in $script:ViewDefinitions.Keys) {
                Show-EobView -Name $name
                $script:Ui.CurrentView | Should -Be $name
                $script:Ui.C.MainContent.Content | Should -Be $script:Ui.Views[$name]
            }
            Set-EobOnboardingStep -Step 7
            $script:Ui.C.OnbExecuteButton.Visibility | Should -Be 'Visible'
            # Fehler in Initialisierung oder Aktualisierung würden als Dialog gemeldet.
            @($script:Ui.DialogLog) | Should -BeNullOrEmpty
            Stop-EobUi
            $script:Ui.Window.Close()
        }
    }
}
