#Requires -Version 7.2
<#
.SYNOPSIS
    Installer von easyONBOARDING: erstellt die Konfiguration (Config\easyONB.ini) über eine Oberfläche.

.DESCRIPTION
    - Auf einem Domänencontroller werden Domäne, UPN-Suffixe, Domänencontroller und OUs sowie,
      falls vorhanden, Entra Connect und Exchange aus dem Active Directory vorbelegt.
    - Auf einem Client oder Mitgliedsserver werden alle Werte manuell eingetragen.
    - Zu jedem Feld und Bereich erscheint nach kurzer Zeit ein Hinweistext.
    - Die erzeugte Datei wird mit derselben Prüfung geladen wie in der Anwendung. Eine vorhandene
      Konfiguration wird nur nach Rückfrage ersetzt und vorher gesichert.

    Das Skript verändert keine Ausführungsrichtlinie, fordert keine Administratorrechte an und
    installiert keine Module.

.PARAMETER ConfigPath
    Zieldatei (Standard: Config\easyONB.ini bzw. Umgebungsvariable EASYONB_CONFIG).

.EXAMPLE
    pwsh -STA -File .\Install-easyONBOARDING.ps1

.NOTES
    Autor: Andreas Hepp (PHINIT.DE)
#>
[CmdletBinding()]
param(
    [string]$ConfigPath
)

Set-StrictMode -Version 3.0
$ErrorActionPreference = 'Stop'

$appRoot = $PSScriptRoot
. (Join-Path -Path $appRoot -ChildPath 'Modules/easyONB.Modules.ps1')
Import-EobModuleSet -ModulesRoot (Join-Path -Path $appRoot -ChildPath 'Modules') -Exclude @('UI')

if (-not $ConfigPath) { $ConfigPath = Get-EobDefaultConfigurationPath }
Start-EobSetupGui -ConfigPath $ConfigPath
