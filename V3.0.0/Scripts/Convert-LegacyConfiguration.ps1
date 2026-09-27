#Requires -Version 7.2
<#
.SYNOPSIS
    Überführt eine easyONB.ini der Versionen 1.3/1.4 in eine 3.0-kompatible Konfiguration.

.DESCRIPTION
    Kommandozeilenvariante von "Tools > Konfiguration migrieren" (Convert-EobLegacyConfiguration):
    - Klartext-Secrets (fixPassword, CompanyVPNPassword, JiraToken, SMTP-Kennwort, Webhooks) werden
      entfernt und durch einen Kommentar ohne Wert ersetzt.
    - SyncCommand wird auskommentiert (keine Codeausführung über die Konfiguration).
    - Fehlende neue Abschnitte werden aus der Vorlage ergänzt; alle übrigen Zeilen, Kommentare und
      unbekannten Schlüssel bleiben unverändert.
    Die Quelldatei wird nie verändert. Anschließend die neue Datei mit
    "Start-easyONBOARDING.ps1 -CheckOnly -ConfigPath <Ziel>" prüfen.

.PARAMETER SourcePath
    INI-Datei der Version 1.x.

.PARAMETER DestinationPath
    Zieldatei (Standard: Config\easyONB.ini). Eine vorhandene Datei wird nur mit -Force überschrieben.

.PARAMETER Force
    Vorhandene Zieldatei überschreiben.

.EXAMPLE
    pwsh -File .\Scripts\Convert-LegacyConfiguration.ps1 -SourcePath ..\V1.4.XX\assets\easyONB.ini -WhatIf

.EXAMPLE
    pwsh -File .\Scripts\Convert-LegacyConfiguration.ps1 -SourcePath D:\alt\easyONB.ini

.OUTPUTS
    Die durchgeführten Änderungen (Zeile, Abschnitt, Schlüssel, Aktion, Meldung).

.NOTES
    Autor: Andreas Hepp (PHINIT.DE)
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory)][string]$SourcePath,
    [string]$DestinationPath,
    [switch]$Force
)

Set-StrictMode -Version 3.0
$ErrorActionPreference = 'Stop'

$appRoot = Split-Path -Parent $PSScriptRoot
. (Join-Path -Path $appRoot -ChildPath 'Modules/easyONB.Modules.ps1')
Import-EobModuleSet -ModulesRoot (Join-Path -Path $appRoot -ChildPath 'Modules') -Exclude 'UI'

if (-not $DestinationPath) {
    $DestinationPath = Join-Path -Path (Join-Path -Path $appRoot -ChildPath 'Config') -ChildPath 'easyONB.ini'
}
$source = (Resolve-Path -LiteralPath $SourcePath).ProviderPath
$target = [System.IO.Path]::GetFullPath($DestinationPath)
if ($source -ieq $target) {
    throw 'Quelle und Ziel dürfen nicht identisch sein (die Quelldatei bleibt unverändert).'
}
if ($PSCmdlet.ShouldProcess($target, "Konfiguration aus '$source' migrieren")) {
    $changes = @(Convert-EobLegacyConfiguration -SourcePath $source -DestinationPath $target -Force:$Force -Confirm:$false)
    $changes
    # Hinweis standardmäßig anzeigen; ein ausdrückliches -InformationAction des Aufrufers hat Vorrang.
    $informationAction = if ($PSBoundParameters.ContainsKey('InformationAction')) { $InformationPreference } else { 'Continue' }
    Write-Information -MessageData ("{0} Änderung(en) geschrieben nach {1}. Nächster Schritt: Start-easyONBOARDING.ps1 -CheckOnly -ConfigPath '{1}'" -f $changes.Count, $target) -InformationAction $informationAction
}
