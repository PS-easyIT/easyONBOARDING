#Requires -Version 7.2
<#
.SYNOPSIS
    Verarbeitet fällige Offboarding-Phasen aus der Warteschlange (für eine geplante Aufgabe).

.DESCRIPTION
    - Führt die fälligen Phasen "Zum Austritt" und "Aufbewahrung" aus, die bei einem Offboarding
      in der Oberfläche geplant wurden.
    - Führt NIE eine endgültige Löschung aus. Fällige Löschungen werden gemeldet und müssen in
      der Oberfläche (Offboarding > Warteschlange) mit Vorschau und Tippbestätigung freigegeben werden.
    - Privilegierte Konten werden nicht unbeaufsichtigt weiterverarbeitet.
    - Phasen mit Exchange- oder Graph-Aktionen werden nur mit -AllowIntegration und bestehender
      Verbindung ausgeführt; sonst als "manuell erforderlich" gemeldet. Exchange Server/Hybrid
      wird mit -AllowIntegration per Kerberos verbunden; Exchange Online und Microsoft Graph
      erfordern eine interaktive Anmeldung und werden daher unbeaufsichtigt nicht verbunden.
    - Mit -WhatIf wird nur simuliert.

    Das Skript verändert keine Ausführungsrichtlinie und fordert keine Administratorrechte an.
    Das ausführende Konto benötigt die delegierten AD-Rechte für die Offboarding-Aktionen.

.PARAMETER ConfigPath
    Pfad zur easyONB.ini (Standard: Config\easyONB.ini bzw. Umgebungsvariable EASYONB_CONFIG).

.PARAMETER AllowIntegration
    Erlaubt fällige Exchange-/Graph-Aktionen, sofern eine Verbindung besteht bzw. (Exchange Server)
    per Kerberos hergestellt werden kann.

.EXAMPLE
    pwsh -NoProfile -File .\Scripts\Invoke-DueOffboardingPhases.ps1 -WhatIf

.EXAMPLE
    # Geplante Aufgabe (täglich 06:00), Konto mit delegierten AD-Rechten:
    pwsh.exe -NoProfile -NonInteractive -File "D:\easyONBOARDING\V3.0.0\Scripts\Invoke-DueOffboardingPhases.ps1"

.OUTPUTS
    Ein Ergebnisobjekt je betroffenem Vorgang. Exitcodes: 0 = ok, 1 = mindestens ein Vorgang
    blockiert oder fehlgeschlagen, 2 = Konfiguration/Verbindung fehlerhaft.
#>
[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
param(
    [string]$ConfigPath,
    [switch]$AllowIntegration
)

Set-StrictMode -Version 3.0
$ErrorActionPreference = 'Stop'

$appRoot = Split-Path -Parent $PSScriptRoot
. (Join-Path -Path $appRoot -ChildPath 'Modules/easyONB.Modules.ps1')
Import-EobModuleSet -ModulesRoot (Join-Path -Path $appRoot -ChildPath 'Modules') -Exclude 'UI'

if (-not $ConfigPath) { $ConfigPath = Get-EobDefaultConfigurationPath }
$config = Import-EobConfiguration -Path $ConfigPath
$logging = Get-EobLoggingParameter -Config $config
$null = Initialize-EobLogging @logging

$errors = @($config.Findings | Where-Object Severity -EQ 'Error')
if ($errors.Count -gt 0) {
    foreach ($finding in $errors) { Write-EobLog -Level Error -Action 'ConfigCheck' -Message "$($finding.Code): $($finding.Message)" }
    Write-Warning "Die Konfiguration enthält $($errors.Count) Fehler. Details: pwsh -File Start-easyONBOARDING.ps1 -CheckOnly"
    exit 2
}

$connection = Initialize-EobAdConnection -Config $config
if (-not $connection.Connected) {
    Write-EobLog -Level Error -Action 'AdConnect' -Message "Keine AD-Verbindung: $($connection.LastError)"
    Write-Warning "Keine Active-Directory-Verbindung: $($connection.LastError)"
    exit 2
}

if ($AllowIntegration) {
    $mode = [string](Get-EobConfigValue -Config $config -Section 'Exchange' -Key 'Mode')
    if ($mode -in @('OnPremises', 'Hybrid')) {
        try {
            $null = Connect-EobExchange -Config $config -Confirm:$false
        }
        catch {
            Write-EobLog -Level Warning -Action 'ExchangeConnect' -Message "Exchange-Verbindung nicht möglich: $($_.Exception.Message)"
        }
    }
}

Write-EobLog -Level Information -Action 'DueOffboardingRun' -Message ("Verarbeitung fälliger Offboarding-Phasen gestartet (Simulation: {0})." -f [bool]$WhatIfPreference)
$results = @(Invoke-EobDueOffboardingPhase -Config $config -AllowIntegration:$AllowIntegration -WhatIf:$WhatIfPreference -Confirm:$false)
$results

$problems = @($results | Where-Object { $_.Outcome -in @('Blocked', 'Failed') })
Write-EobLog -Level Information -Action 'DueOffboardingRun' -Result $(if ($problems.Count -gt 0) { 'Failed' } else { 'Succeeded' }) `
    -Message ("{0} Vorgang/Vorgänge verarbeitet, {1} mit Problemen, {2} zur manuellen Bearbeitung." -f $results.Count, $problems.Count, @($results | Where-Object { $_.Outcome -in @('ManualRequired', 'ManualApprovalRequired') }).Count)
if ($problems.Count -gt 0) { exit 1 }
exit 0
