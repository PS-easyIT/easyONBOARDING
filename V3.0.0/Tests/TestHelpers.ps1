<#
.SYNOPSIS
    Gemeinsame Hilfsfunktionen für alle Pester-Tests.

.DESCRIPTION
    - Ermittelt das Repository-Wurzelverzeichnis.
    - Definiert Stub-Funktionen für AD-, Exchange- und Graph-Cmdlets. Globale Funktionen haben
      Vorrang vor Cmdlets; damit ist ausgeschlossen, dass Tests versehentlich echte Systeme
      ansprechen. Jeder Stub wirft, wenn er nicht gemockt wurde.
    - Lädt die Module in Abhängigkeitsreihenfolge.
#>

$script:EobRepoRoot = Split-Path -Parent $PSScriptRoot

function Get-EobTestRepoRoot {
    return $script:EobRepoRoot
}

function Import-EobTestModule {
    param(
        [string[]]$Exclude = @('UI')
    )

    . (Join-Path -Path $script:EobRepoRoot -ChildPath 'Tests/Fixtures/ExternalCommandStubs.ps1')
    . (Join-Path -Path $script:EobRepoRoot -ChildPath 'Modules/easyONB.Modules.ps1')

    # Entwicklungshilfe: EOB_TEST_ONLY="Core,Configuration" lädt nur die genannten Module.
    if (-not [string]::IsNullOrWhiteSpace($env:EOB_TEST_ONLY)) {
        $only = $env:EOB_TEST_ONLY -split ','
        $Exclude = @($Exclude) + @((Get-EobModuleManifest).ShortName | Where-Object { $_ -notin $only })
    }
    Import-EobModuleSet -ModulesRoot (Join-Path -Path $script:EobRepoRoot -ChildPath 'Modules') -Exclude $Exclude -Force
}

function New-EobTestSecureString {
    <#
        Erzeugt einen SecureString für Tests, ohne einen Klartextwert auszugeben.
    #>
    param([string]$Value = ('T3st!' + [guid]::NewGuid().ToString('N').Substring(0, 12)))

    return (ConvertTo-SecureString -String $Value -AsPlainText -Force)
}
