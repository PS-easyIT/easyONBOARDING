<#
.SYNOPSIS
    Lädt die easyONBOARDING-Module in der festgelegten Abhängigkeitsreihenfolge.

.DESCRIPTION
    Wird vom Einstiegsskript, von Hilfsskripten und von den Pester-Tests per Dot-Sourcing
    eingebunden. Die Reihenfolge entspricht den Abhängigkeiten (Core zuerst, UI zuletzt).
#>

Set-StrictMode -Version 3.0

$script:EobModuleOrder = @(
    'Core'
    'Configuration'
    'Security'
    'ActiveDirectory'
    'Exchange'
    'Entra'
    'Reporting'
    'Onboarding'
    'Offboarding'
    'UI'
)

function Get-EobModuleManifest {
    <#
    .SYNOPSIS
        Liefert die Manifestpfade aller Module in Ladereihenfolge.
    .PARAMETER ModulesRoot
        Wurzelverzeichnis der Module (Standard: Verzeichnis dieses Skripts).
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [string]$ModulesRoot = $PSScriptRoot
    )

    foreach ($name in $script:EobModuleOrder) {
        [pscustomobject]@{
            Name         = "easyONB.$name"
            ShortName    = $name
            ManifestPath = Join-Path -Path (Join-Path -Path $ModulesRoot -ChildPath $name) -ChildPath "easyONB.$name.psd1"
        }
    }
}

function Import-EobModuleSet {
    <#
    .SYNOPSIS
        Importiert alle easyONBOARDING-Module global.
    .PARAMETER ModulesRoot
        Wurzelverzeichnis der Module.
    .PARAMETER Exclude
        Kurznamen von Modulen, die nicht geladen werden sollen (z. B. 'UI' für Tests oder CLI).
    .PARAMETER Force
        Erzwingt das erneute Laden bereits geladener Module.
    #>
    [CmdletBinding()]
    param(
        [string]$ModulesRoot = $PSScriptRoot,
        [string[]]$Exclude = @(),
        [switch]$Force
    )

    foreach ($module in Get-EobModuleManifest -ModulesRoot $ModulesRoot) {
        if ($Exclude -contains $module.ShortName) {
            continue
        }
        if (-not (Test-Path -LiteralPath $module.ManifestPath -PathType Leaf)) {
            throw "Modulmanifest nicht gefunden: $($module.ManifestPath)"
        }
        Import-Module -Name $module.ManifestPath -Global -Force:$Force -ErrorAction Stop -DisableNameChecking
    }
}
