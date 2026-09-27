#Requires -Version 7.2
<#
.SYNOPSIS
    Startet easyONBOARDING 3 (Oberfläche) oder prüft Konfiguration und Umgebung (-CheckOnly).

.DESCRIPTION
    - Oberfläche: Windows mit PowerShell 7.2 oder höher, STA-Thread (Standard bei pwsh auf Windows).
    - -CheckOnly: lädt alle Fachmodule, prüft die Konfiguration und die Integrationen und gibt das
      Ergebnis aus. Funktioniert ohne Oberfläche (auch in CI und unter Linux/macOS).

    Das Skript ändert keine Ausführungsrichtlinie, fordert keine Administratorrechte an und
    installiert keine Module. Neue Vorgänge starten gemäß [Security] SimulationByDefault im
    Simulationsmodus; der Live-Modus wird in der Oberfläche bewusst und mit Bestätigung aktiviert.

.PARAMETER ConfigPath
    Pfad zur INI-Datei (Standard: Config\easyONB.ini bzw. Umgebungsvariable EASYONB_CONFIG).

.PARAMETER CheckOnly
    Nur prüfen, keine Oberfläche starten. Exitcode 0 = keine Fehler, 2 = Fehler in der Konfiguration.

.PARAMETER Simulation
    Startet die Oberfläche unabhängig von der Konfiguration im Simulationsmodus.

.PARAMETER Theme
    Farbschema der Oberfläche (Light oder Dark); überschreibt [UI] Theme für diese Sitzung.

.EXAMPLE
    pwsh -File .\Start-easyONBOARDING.ps1

.EXAMPLE
    pwsh -File .\Start-easyONBOARDING.ps1 -CheckOnly

.NOTES
    Autor: Andreas Hepp (PHINIT.DE)
#>
[CmdletBinding()]
param(
    [string]$ConfigPath,
    [switch]$CheckOnly,
    [switch]$Simulation,
    [ValidateSet('Light', 'Dark')][string]$Theme
)

Set-StrictMode -Version 3.0
$ErrorActionPreference = 'Stop'

$appRoot = $PSScriptRoot
. (Join-Path -Path $appRoot -ChildPath 'Modules/easyONB.Modules.ps1')
$exclude = if ($CheckOnly) { @('UI') } else { @() }
Import-EobModuleSet -ModulesRoot (Join-Path -Path $appRoot -ChildPath 'Modules') -Exclude $exclude

if (-not $ConfigPath) { $ConfigPath = Get-EobDefaultConfigurationPath }
$config = Import-EobConfiguration -Path $ConfigPath
$logging = Get-EobLoggingParameter -Config $config
$null = Initialize-EobLogging @logging -EnableQueue:(-not $CheckOnly)
$adConnection = Initialize-EobAdConnection -Config $config

if ($CheckOnly) {
    $summary = Get-EobConfigFindingSummary -Config $config
    $log = Get-EobLogStatus
    Write-Output ('easyONBOARDING {0} - Prüfung' -f (Get-EobVersion))
    Write-Output ('PowerShell {0} ({1}), {2}' -f $PSVersionTable.PSVersion, $PSVersionTable.PSEdition, [System.Runtime.InteropServices.RuntimeInformation]::OSDescription)
    Write-Output ('Konfiguration: {0} ({1})' -f $ConfigPath, $(if ($config.Exists) { 'vorhanden' } else { 'FEHLT' }))
    Write-Output ('Befunde: {0} Fehler, {1} Warnungen, {2} Hinweise' -f $summary.Errors, $summary.Warnings, $summary.Informations)
    Write-Output ('Logs: {0} | Audit: {1}{2}' -f $log.Directory, $log.AuditDirectory, $(if ($log.FallbackActive) { ' (Ersatzverzeichnis aktiv)' } else { '' }))
    Write-Output ('Active Directory: {0}' -f $(if ($adConnection.Connected) { "verbunden mit $($adConnection.Server)" } else { "nicht verbunden ($($adConnection.LastError))" }))
    $gui = if ($IsWindows) { 'Windows - Oberfläche verfügbar' } else { 'kein Windows - nur Prüfung/Skripte' }
    Write-Output ('Oberfläche: {0}' -f $gui)
    Write-Output ''
    Write-Output 'Integrationen:'
    Get-EobIntegrationStatus -Config $config | Select-Object -Property Name, State, Implementation, Detail, Hint | Format-Table -AutoSize -Wrap | Out-String -Width 220 | Write-Output
    if (@($config.Findings).Count -gt 0) {
        Write-Output 'Befunde der Konfiguration:'
        $config.Findings | Select-Object -Property Severity, Code, Field, Message | Format-Table -AutoSize -Wrap | Out-String -Width 220 | Write-Output
    }
    if ($summary.Errors -gt 0) { exit 2 }
    exit 0
}

$guiParameters = @{ Config = $config; ConfigPath = $ConfigPath }
if ($Simulation) { $guiParameters['ForceSimulation'] = $true }
if ($Theme) { $guiParameters['Theme'] = $Theme }
Start-EobGui @guiParameters
