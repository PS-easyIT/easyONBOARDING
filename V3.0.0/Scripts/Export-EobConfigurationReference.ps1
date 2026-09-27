#Requires -Version 7.2
<#
.SYNOPSIS
    Erzeugt docs/CONFIGURATION-REFERENCE.md aus dem Konfigurationsschema.

.DESCRIPTION
    Die Referenz wird vollständig aus Config/schemas/easyONB.schema.psd1 erzeugt; Abschnitte und
    Schlüssel erscheinen in der Reihenfolge des Schemas. Die Qualitätsprüfung
    (Scripts/Invoke-EobQualityCheck.ps1 -Check Docs) meldet eine veraltete Referenz.

.PARAMETER Root
    Anwendungsverzeichnis (Standard: übergeordnetes Verzeichnis von Scripts).

.PARAMETER PassThru
    Gibt den Markdown-Text zurück, statt die Datei zu schreiben.

.EXAMPLE
    pwsh -File ./Scripts/Export-EobConfigurationReference.ps1

.NOTES
    Autor: Andreas Hepp (PHINIT.DE)
#>
[CmdletBinding(SupportsShouldProcess)]
[OutputType([string])]
param(
    [string]$Root = (Split-Path -Parent $PSScriptRoot),
    [switch]$PassThru
)

Set-StrictMode -Version 3.0
$ErrorActionPreference = 'Stop'

$schemaPath = Join-Path -Path $Root -ChildPath 'Config/schemas/easyONB.schema.psd1'
$targetPath = Join-Path -Path $Root -ChildPath 'docs/CONFIGURATION-REFERENCE.md'
$schema = Import-PowerShellDataFile -LiteralPath $schemaPath

# Hashtables verlieren beim Import ihre Reihenfolge; die Reihenfolge kommt daher aus dem Syntaxbaum.
$parseErrors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseFile($schemaPath, [ref]$null, [ref]$parseErrors)
if (@($parseErrors).Count -gt 0) { throw "Schema enthält Syntaxfehler: $($parseErrors[0].Message)" }
$isHashtable = { param($node) $node -is [System.Management.Automation.Language.HashtableAst] }

function Get-AstPairValue {
    param($HashtableAst, [string]$Name)
    foreach ($pair in $HashtableAst.KeyValuePairs) {
        if ([string]$pair.Item1.SafeGetValue() -eq $Name) { return $pair.Item2.Find($isHashtable, $true) }
    }
    return $null
}

function Get-AstKeyOrder {
    param($HashtableAst)
    if ($null -eq $HashtableAst) { return @() }
    return @($HashtableAst.KeyValuePairs | ForEach-Object { [string]$_.Item1.SafeGetValue() })
}

function ConvertTo-MarkdownCell {
    # Für Codeabschnitte in Tabellenzellen: nur Pipe und Zeilenumbrüche maskieren.
    param([AllowNull()][object]$Text)
    return ([string]$Text).Replace('|', '\|').Replace("`r", ' ').Replace("`n", ' ')
}

function ConvertTo-MarkdownText {
    # Für Fließtext: zusätzlich spitze Klammern maskieren (sonst als HTML-Tag interpretiert).
    param([AllowNull()][object]$Text)
    return (ConvertTo-MarkdownCell -Text $Text).Replace('<', '&lt;').Replace('>', '&gt;')
}

function Get-TypeText {
    param([hashtable]$Definition)
    $type = [string]$Definition['Type']
    if ($Definition.ContainsKey('Values')) { return "$type ($(@($Definition['Values']) -join ', '))" }
    if ($Definition.ContainsKey('Min') -and $Definition.ContainsKey('Max')) { return "$type ($($Definition['Min'])–$($Definition['Max']))" }
    return $type
}

function Get-DescriptionText {
    param([hashtable]$Definition)
    $parts = [System.Collections.Generic.List[string]]::new()
    if ($Definition['Required']) { $parts.Add('**Pflicht.**') }
    if ($Definition['Secret']) { $parts.Add('**Secret – wird ignoriert.**') }
    if ($Definition['Unused']) { $parts.Add('*Ohne Wirkung.*') }
    elseif ($Definition['Deprecated']) { $parts.Add('*Veraltet.*') }
    if ($Definition['Replacement']) { $parts.Add("Ersatz: ``$($Definition['Replacement'])``.") }
    $parts.Add((ConvertTo-MarkdownText -Text $Definition['Description']))
    return ($parts -join ' ')
}

function Get-KeyTable {
    param([hashtable]$Keys, [string[]]$Order, [object[]]$KeyPatterns = @())
    $lines = [System.Collections.Generic.List[string]]::new()
    $lines.Add('| Schlüssel | Typ | Standard | Beschreibung |')
    $lines.Add('|---|---|---|---|')
    foreach ($name in $Order) {
        $definition = $Keys[$name]
        $default = if ($definition.ContainsKey('Default') -and [string]$definition['Default'] -ne '') { '`' + (ConvertTo-MarkdownCell -Text $definition['Default']) + '`' } else { '–' }
        $lines.Add("| ``$name`` | $(ConvertTo-MarkdownText -Text (Get-TypeText -Definition $definition)) | $default | $(Get-DescriptionText -Definition $definition) |")
    }
    # Ein leeres Array aus einem if-Ausdruck kommt als $null an; @($null) hätte ein Element.
    foreach ($pattern in @($KeyPatterns | Where-Object { $null -ne $_ })) {
        $lines.Add("| Muster ``$(ConvertTo-MarkdownCell -Text $pattern['Pattern'])`` | $(ConvertTo-MarkdownCell -Text (Get-TypeText -Definition $pattern)) | – | $(Get-DescriptionText -Definition $pattern) |")
    }
    return $lines
}

$topAst = $ast.Find($isHashtable, $true)
$sectionsAst = Get-AstPairValue -HashtableAst $topAst -Name 'Sections'
$output = [System.Collections.Generic.List[string]]::new()
$output.Add('# Konfigurationsreferenz')
$output.Add('')
$output.Add('> Automatisch erzeugt aus `Config/schemas/easyONB.schema.psd1` mit')
$output.Add('> `Scripts/Export-EobConfigurationReference.ps1`. Bitte nicht von Hand bearbeiten.')
$output.Add('')
$output.Add('Erläuterungen zu Aufbau, Zusammenführung und Beispielen: [CONFIGURATION.md](CONFIGURATION.md).')
$output.Add('')
$output.Add('Kennzeichnungen: **Pflicht** = muss gesetzt sein · *Veraltet* = wird noch gelesen, bitte durch den')
$output.Add('Ersatz austauschen · *Ohne Wirkung* = wird in 3.x nicht ausgewertet · **Secret – wird ignoriert** =')
$output.Add('Klartext-Geheimnisse werden nicht unterstützt, gemeldet und nie verwendet.')

foreach ($sectionName in Get-AstKeyOrder -HashtableAst $sectionsAst) {
    $section = $schema.Sections[$sectionName]
    $output.Add('')
    $output.Add("## [$sectionName]")
    $output.Add('')
    $description = ConvertTo-MarkdownText -Text $section['Description']
    if ($section['Deprecated']) { $description = "*Veralteter Abschnitt.* $description" }
    $output.Add($description)
    if ($section['FreeForm']) {
        $output.Add('')
        $output.Add("Freie Schlüssel (Bezeichnung=Wert), Werttyp: $($section['ValueType']).")
    }
    if ($section.ContainsKey('Keys')) {
        $keyOrder = Get-AstKeyOrder -HashtableAst (Get-AstPairValue -HashtableAst (Get-AstPairValue -HashtableAst $sectionsAst -Name $sectionName) -Name 'Keys')
        $patterns = if ($section.ContainsKey('KeyPatterns')) { @($section['KeyPatterns']) } else { @() }
        $output.Add('')
        foreach ($line in Get-KeyTable -Keys $section['Keys'] -Order $keyOrder -KeyPatterns $patterns) { $output.Add($line) }
    }
}

$output.Add('')
$output.Add('## Dynamische Abschnitte')
$patternAsts = @($topAst.FindAll({ param($node) $node -is [System.Management.Automation.Language.HashtableAst] -and @($node.KeyValuePairs | Where-Object { [string]$_.Item1.SafeGetValue() -eq 'Pattern' }).Count -gt 0 -and @($node.KeyValuePairs | Where-Object { [string]$_.Item1.SafeGetValue() -eq 'Name' }).Count -gt 0 }, $true))
foreach ($sectionPattern in @($schema.SectionPatterns)) {
    $output.Add('')
    $output.Add("### $($sectionPattern['Name']) (Abschnittsmuster ``$($sectionPattern['Pattern'])``)")
    $output.Add('')
    $output.Add((ConvertTo-MarkdownText -Text $sectionPattern['Description']))
    $keys = if ($sectionPattern.ContainsKey('Keys')) { $sectionPattern['Keys'] } else { @{} }
    $patternAst = $patternAsts | Where-Object { @($_.KeyValuePairs | Where-Object { [string]$_.Item1.SafeGetValue() -eq 'Name' -and [string]$_.Item2.Find({ param($node) $node -is [System.Management.Automation.Language.StringConstantExpressionAst] }, $true).Value -eq $sectionPattern['Name'] }).Count -gt 0 } | Select-Object -First 1
    $keyOrder = if ($null -ne $patternAst) { Get-AstKeyOrder -HashtableAst (Get-AstPairValue -HashtableAst $patternAst -Name 'Keys') } else { @($keys.Keys | Sort-Object) }
    $patterns = if ($sectionPattern.ContainsKey('KeyPatterns')) { @($sectionPattern['KeyPatterns']) } else { @() }
    $output.Add('')
    foreach ($line in Get-KeyTable -Keys $keys -Order $keyOrder -KeyPatterns $patterns) { $output.Add($line) }
}

$output.Add('')
$output.Add('## Offboarding-Aktionen')
$output.Add('')
$output.Add('Gültige Aktionsnamen für `Action.<Aktion>=<Phase>` in Offboarding-Vorlagen:')
$output.Add('')
$output.Add((@($schema.OffboardingActions) | ForEach-Object { "``$_``" }) -join ', ')

$text = ($output -join "`n") + "`n"
if ($PassThru) {
    return $text
}
if ($PSCmdlet.ShouldProcess($targetPath, 'Konfigurationsreferenz schreiben')) {
    [System.IO.File]::WriteAllText($targetPath, $text, [System.Text.UTF8Encoding]::new($false))
    Write-Information -MessageData "Geschrieben: $targetPath" -InformationAction Continue
}
