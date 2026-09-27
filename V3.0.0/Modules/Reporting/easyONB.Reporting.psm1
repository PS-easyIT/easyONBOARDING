#Requires -Version 7.2
<#
    easyONB.Reporting
    Vorgangsberichte (HTML, JSON, CSV, TXT, optional PDF), Willkommensdokument und Welcome-Mail
    ohne Kennwort, Audit-Auswertung sowie Status von SMTP, Dateiserver und PDF-Erzeugung.

    Grundsätze:
    - Kennwörter, Handler-Parameter und Konfigurationswerte mit Secret-Kennzeichnung gelangen
      nie in Berichte, Dokumente oder E-Mails. Kennwort-Platzhalter werden durch einen Hinweis ersetzt.
    - Alle Werte werden HTML-kodiert eingesetzt, CSV-Werte gegen Formel-Injektion geschützt.
    - PDF ist optional: Edge (headless) -> wkhtmltopdf -> nur HTML.
#>

Set-StrictMode -Version 3.0

#region Konstanten

$script:PlaceholderPattern = '\{\{\s*([A-Za-z][A-Za-z0-9_]*)(?:\s*:\s*([^{}]*?))?\s*\}\}'
# Kennwort-Platzhalter werden nie befüllt.
$script:SecretPlaceholderPattern = '(?i)(passw(or)?[dt]|kennwort|custompw|secret|token)'
# Platzhalter, deren Wert intern als (bereits kodiertes) HTML erzeugt wird.
$script:RawHtmlPlaceholders = @('PasswordPolicyList', 'WebsitesHTML', 'LogoTag')
# Platzhalter mit "Password" im Namen, die kein Kennwort enthalten.
$script:NonSecretPlaceholders = @('PasswordPolicyList')
$script:AssetMimeTypes = @{
    '.png' = 'image/png'; '.jpg' = 'image/jpeg'; '.jpeg' = 'image/jpeg'; '.gif' = 'image/gif'; '.svg' = 'image/svg+xml'; '.webp' = 'image/webp'
}
$script:MaxAssetBytes = 2MB
$script:ReportFormats = @('Html', 'Json', 'Csv', 'Txt', 'Pdf')
$script:KindTitles = @{
    Onboarding      = 'Onboarding'
    Offboarding     = 'Offboarding'
    UserUpdate      = 'Benutzer aktualisieren'
    PasswordReset   = 'Kennwort zurücksetzen'
    BulkOnboarding  = 'Massen-Onboarding'
    BulkOffboarding = 'Massen-Offboarding'
}
$script:StatusTexts = @{
    Planned               = 'Geplant'
    Running               = 'Läuft'
    Succeeded             = 'Erfolgreich'
    Simulated             = 'Simuliert'
    Warning               = 'Warnung'
    Skipped               = 'Übersprungen'
    Failed                = 'Fehlgeschlagen'
    Cancelled             = 'Abgebrochen'
    CompletedWithWarnings = 'Abgeschlossen mit Warnungen'
}

#endregion

#region Hilfsfunktionen

function Get-EobStatusText {
    [CmdletBinding()]
    [OutputType([string])]
    param([AllowNull()][string]$Status)

    if ([string]::IsNullOrEmpty($Status)) { return '' }
    if ($script:StatusTexts.ContainsKey($Status)) { return $script:StatusTexts[$Status] }
    return $Status
}

function Format-EobReportDate {
    [CmdletBinding()]
    [OutputType([string])]
    param([AllowNull()][object]$Value)

    if ($null -eq $Value) { return '' }
    if ($Value -is [datetime]) { return $Value.ToString('dd.MM.yyyy HH:mm:ss', [System.Globalization.CultureInfo]::InvariantCulture) }
    return [string]$Value
}

function New-EobNotProcessedResult {
    <#
        Ergebnis, wenn ShouldProcess die Aktion nicht zulässt (Simulation oder keine Bestätigung).
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory)][string]$Message,
        [hashtable]$Data
    )

    if ($WhatIfPreference) {
        return New-EobResult -Status Succeeded -Message ('Simulation: ' + $Message) -Data $Data
    }
    return New-EobResult -Status Skipped -Message ('Nicht bestätigt: ' + $Message) -Data $Data
}

function Write-EobTextFile {
    <#
        Schreibt eine Textdatei atomar (temporäre Datei + Umbenennen) in UTF-8.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][AllowEmptyString()][string]$Content,
        [switch]$Bom
    )

    $directory = Split-Path -Parent $Path
    if ($directory -and -not (Test-Path -LiteralPath $directory -PathType Container)) {
        $null = New-Item -ItemType Directory -Path $directory -Force -ErrorAction Stop
    }
    $temporary = "$Path.$([guid]::NewGuid().ToString('N')).tmp"
    try {
        [System.IO.File]::WriteAllText($temporary, $Content, [System.Text.UTF8Encoding]::new([bool]$Bom))
        Move-Item -LiteralPath $temporary -Destination $Path -Force -ErrorAction Stop
    }
    finally {
        if (Test-Path -LiteralPath $temporary) { Remove-Item -LiteralPath $temporary -Force -ErrorAction SilentlyContinue }
    }
}

function Get-EobReportDirectory {
    <#
    .SYNOPSIS
        Liefert das Berichtsverzeichnis ([Report] ReportPath), optional mit Unterordner je Vorgangsart.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [AllowNull()][pscustomobject]$Config,
        [string]$Kind
    )

    $root = [string](Get-EobConfigValue -Config $Config -Section 'Report' -Key 'ReportPath' -As Path)
    if ([string]::IsNullOrWhiteSpace($root)) {
        $root = Join-Path -Path (Get-EobAppRoot) -ChildPath 'Reports'
    }
    if ($Kind) {
        return (Join-Path -Path $root -ChildPath (Get-EobSafeFileName -Name $Kind))
    }
    return $root
}

#endregion

#region Vorgangsberichte

function New-EobPlanReport {
    <#
    .SYNOPSIS
        Erzeugt das Berichtsmodell eines Plans.
    .DESCRIPTION
        Übernommen werden nur Anzeige- und Ergebnisdaten. Kennwörter (Plan.Secrets), Handler-Parameter
        und die Konfiguration sind nie Bestandteil des Berichts; alle Texte werden zusätzlich redigiert.
    .PARAMETER Plan
        Ausgeführter oder simulierter Plan.
    .PARAMETER Title
        Optionaler Titel (Standard: Vorgangsart und Konto).
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory)][pscustomobject]$Plan,
        [string]$Title
    )

    $kindTitle = if ($script:KindTitles.ContainsKey($Plan.Kind)) { $script:KindTitles[$Plan.Kind] } else { [string]$Plan.Kind }
    $subjectName = [string](Get-EobPropertyValue -InputObject $Plan.Subject -Name 'SamAccountName' -Default '')
    if ([string]::IsNullOrWhiteSpace($Title)) {
        $Title = if ($subjectName) { "$kindTitle - $subjectName" } else { $kindTitle }
    }

    $steps = [System.Collections.Generic.List[object]]::new()
    foreach ($step in $Plan.Steps) {
        $steps.Add([pscustomobject][ordered]@{
                Id         = $step.Id
                Phase      = [string]$step.Phase
                Action     = $step.Action
                Title      = Protect-EobSensitiveText -Text ([string]$step.Title)
                Target     = Protect-EobSensitiveText -Text ([string]$step.Target)
                Risk       = $step.Risk
                Enabled    = [bool]$step.Enabled
                DueDate    = $step.DueDate
                Status     = $step.Status
                Message    = Protect-EobSensitiveText -Text ([string]$step.Message)
                DurationMs = $step.DurationMs
                Details    = @($step.Details | ForEach-Object { Protect-EobSensitiveText -Text ([string]$_) })
            })
    }
    $findings = [System.Collections.Generic.List[object]]::new()
    foreach ($finding in $Plan.Findings) {
        $findings.Add([pscustomobject][ordered]@{
                Severity = $finding.Severity
                Code     = $finding.Code
                Field    = [string]$finding.Field
                Message  = Protect-EobSensitiveText -Text ([string]$finding.Message)
            })
    }

    $summary = [ordered]@{}
    foreach ($key in $Plan.Summary.Keys) {
        $summary[$key] = Protect-EobSensitiveText -Text ([string]$Plan.Summary[$key])
    }
    $subject = [ordered]@{}
    foreach ($key in @($Plan.Subject.Keys | Sort-Object)) {
        $subject[$key] = Protect-EobSensitiveText -Text ([string]$Plan.Subject[$key])
    }

    [pscustomobject][ordered]@{
        PSTypeName    = 'Eob.Report'
        SchemaVersion = 1
        Title         = $Title
        Kind          = $Plan.Kind
        KindTitle     = $kindTitle
        OperationId   = $Plan.OperationId
        Tool          = [string](Get-EobPropertyValue -InputObject $Plan.Context -Name 'Tool' -Default "easyONBOARDING $(Get-EobVersion)")
        Actor         = [string](Get-EobPropertyValue -InputObject $Plan.Context -Name 'Actor' -Default '')
        Computer      = [string](Get-EobPropertyValue -InputObject $Plan.Context -Name 'Computer' -Default '')
        CreatedAt     = Get-Date
        PlannedAt     = $Plan.CreatedAt
        StartedAt     = $Plan.StartedAt
        CompletedAt   = $Plan.CompletedAt
        Simulation    = [bool]$Plan.Simulation
        Status        = $Plan.Status
        Outcome       = Get-EobPlanOutcome -Plan $Plan
        SubjectName   = $subjectName
        Subject       = $subject
        Summary       = $summary
        Statistic     = Get-EobPlanStatistic -Plan $Plan
        Steps         = $steps.ToArray()
        Findings      = $findings.ToArray()
    }
}

function ConvertTo-EobReportHtml {
    <#
    .SYNOPSIS
        Rendert einen Vorgangsbericht als eigenständige HTML-Seite (ohne externe Ressourcen).
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory)][pscustomobject]$Report,
        [AllowNull()][pscustomobject]$Config
    )

    $e = { param($Value) ConvertTo-EobHtmlEncoded -Value $Value }
    $footer = [string](Get-EobConfigValue -Config $Config -Section 'Report' -Key 'ReportFooter')
    $mode = if ($Report.Simulation) { 'Simulation (keine Änderungen)' } else { 'Ausführung' }
    $outcomeClass = switch ($Report.Outcome) {
        'Succeeded' { 'ok' } 'Simulated' { 'info' } 'Failed' { 'error' } 'Cancelled' { 'warn' } default { 'warn' }
    }

    $builder = [System.Text.StringBuilder]::new()
    [void]$builder.AppendLine('<!DOCTYPE html>')
    [void]$builder.AppendLine('<html lang="de"><head><meta charset="utf-8">')
    [void]$builder.AppendLine('<meta name="viewport" content="width=device-width, initial-scale=1">')
    [void]$builder.AppendLine("<title>$(& $e $Report.Title)</title>")
    [void]$builder.AppendLine(@'
<style>
  :root { --fg:#1f2328; --muted:#59636e; --line:#d0d7de; --accent:#0f6cbd; --ok:#1a7f37; --warn:#9a6700; --error:#cf222e; --info:#0969da; }
  * { box-sizing: border-box; }
  body { font-family: "Segoe UI", Arial, sans-serif; color: var(--fg); margin: 24px; font-size: 13px; line-height: 1.45; }
  h1 { font-size: 20px; margin: 0 0 4px 0; color: var(--accent); }
  h2 { font-size: 15px; margin: 24px 0 8px 0; border-bottom: 1px solid var(--line); padding-bottom: 4px; }
  table { width: 100%; border-collapse: collapse; margin-bottom: 8px; }
  th, td { text-align: left; vertical-align: top; padding: 5px 8px; border-bottom: 1px solid var(--line); }
  th { background: #f6f8fa; font-weight: 600; }
  table.kv th { width: 28%; }
  .badge { display: inline-block; padding: 1px 8px; border-radius: 10px; font-weight: 600; font-size: 12px; border: 1px solid currentColor; }
  .ok { color: var(--ok); } .warn { color: var(--warn); } .error { color: var(--error); } .info { color: var(--info); }
  .muted { color: var(--muted); } ul.details { margin: 2px 0 0 16px; padding: 0; color: var(--muted); }
  footer { margin-top: 32px; color: var(--muted); font-size: 11px; border-top: 1px solid var(--line); padding-top: 8px; }
  @media print { body { margin: 12mm; } h2 { page-break-after: avoid; } tr { page-break-inside: avoid; } }
</style>
'@)
    [void]$builder.AppendLine('</head><body>')
    [void]$builder.AppendLine("<h1>$(& $e $Report.Title)</h1>")
    [void]$builder.AppendLine("<div class=""muted"">$(& $e $Report.Tool)</div>")

    [void]$builder.AppendLine('<h2>Vorgang</h2><table class="kv">')
    $meta = [ordered]@{
        'Vorgangsart'    = $Report.KindTitle
        'Operation-ID'   = $Report.OperationId
        'Modus'          = $mode
        'Ausgeführt von' = $Report.Actor
        'Computer'       = $Report.Computer
        'Geplant'        = Format-EobReportDate -Value $Report.PlannedAt
        'Gestartet'      = Format-EobReportDate -Value $Report.StartedAt
        'Abgeschlossen'  = Format-EobReportDate -Value $Report.CompletedAt
    }
    foreach ($key in $meta.Keys) {
        [void]$builder.AppendLine("<tr><th>$(& $e $key)</th><td>$(& $e $meta[$key])</td></tr>")
    }
    [void]$builder.AppendLine("<tr><th>Ergebnis</th><td><span class=""badge $outcomeClass"">$(& $e (Get-EobStatusText -Status $Report.Outcome))</span></td></tr>")
    $statistic = $Report.Statistic
    $statText = "{0} Schritte: {1} erfolgreich, {2} simuliert, {3} Warnung, {4} übersprungen, {5} fehlgeschlagen, {6} geplant" -f `
        $statistic.Total, $statistic.Succeeded, $statistic.Simulated, $statistic.Warning, $statistic.Skipped, $statistic.Failed, $statistic.Planned
    [void]$builder.AppendLine("<tr><th>Schritte</th><td>$(& $e $statText)</td></tr>")
    [void]$builder.AppendLine('</table>')

    if ($Report.Summary.Count -gt 0) {
        [void]$builder.AppendLine('<h2>Zusammenfassung</h2><table class="kv">')
        foreach ($key in $Report.Summary.Keys) {
            [void]$builder.AppendLine("<tr><th>$(& $e $key)</th><td>$(& $e $Report.Summary[$key])</td></tr>")
        }
        [void]$builder.AppendLine('</table>')
    }

    [void]$builder.AppendLine('<h2>Schritte</h2><table><thead><tr><th>#</th><th>Phase</th><th>Schritt</th><th>Ziel</th><th>Risiko</th><th>Status</th><th>Meldung</th></tr></thead><tbody>')
    foreach ($step in $Report.Steps) {
        $class = switch ($step.Status) { 'Succeeded' { 'ok' } 'Simulated' { 'info' } 'Failed' { 'error' } 'Warning' { 'warn' } 'Skipped' { 'muted' } default { '' } }
        $details = ''
        if (@($step.Details).Count -gt 0) {
            $details = '<ul class="details">' + ((@($step.Details) | ForEach-Object { "<li>$(& $e $_)</li>" }) -join '') + '</ul>'
        }
        [void]$builder.AppendLine(("<tr><td>{0}</td><td>{1}</td><td>{2}{3}</td><td>{4}</td><td>{5}</td><td class=""{6}"">{7}</td><td>{8}</td></tr>" -f `
                    (& $e $step.Id), (& $e $step.Phase), (& $e $step.Title), $details, (& $e $step.Target), (& $e $step.Risk), $class,
                (& $e (Get-EobStatusText -Status $step.Status)), (& $e $step.Message)))
    }
    [void]$builder.AppendLine('</tbody></table>')

    if (@($Report.Findings).Count -gt 0) {
        [void]$builder.AppendLine('<h2>Hinweise</h2><table><thead><tr><th>Schwere</th><th>Code</th><th>Feld</th><th>Meldung</th></tr></thead><tbody>')
        foreach ($finding in $Report.Findings) {
            $class = switch ($finding.Severity) { 'Error' { 'error' } 'Warning' { 'warn' } default { 'info' } }
            [void]$builder.AppendLine(("<tr><td class=""{0}"">{1}</td><td>{2}</td><td>{3}</td><td>{4}</td></tr>" -f `
                        $class, (& $e $finding.Severity), (& $e $finding.Code), (& $e $finding.Field), (& $e $finding.Message)))
        }
        [void]$builder.AppendLine('</tbody></table>')
    }

    $footerText = "Erstellt am $(Format-EobReportDate -Value $Report.CreatedAt) mit $($Report.Tool). Kennwörter sind nicht Bestandteil dieses Berichts."
    [void]$builder.AppendLine("<footer>$(& $e $footerText)$(if ($footer) { '<br>' + (& $e $footer) })</footer>")
    [void]$builder.AppendLine('</body></html>')
    return $builder.ToString()
}

function ConvertTo-EobReportText {
    <#
    .SYNOPSIS
        Rendert einen Vorgangsbericht als Klartext.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param([Parameter(Mandatory)][pscustomobject]$Report)

    $lines = [System.Collections.Generic.List[string]]::new()
    $lines.Add($Report.Title)
    $lines.Add(('=' * [Math]::Min(80, [Math]::Max(10, $Report.Title.Length))))
    $lines.Add("Vorgangsart:    $($Report.KindTitle)")
    $lines.Add("Operation-ID:   $($Report.OperationId)")
    $lines.Add("Modus:          $(if ($Report.Simulation) { 'Simulation (keine Änderungen)' } else { 'Ausführung' })")
    $lines.Add("Ausgeführt von: $($Report.Actor) auf $($Report.Computer)")
    $lines.Add("Gestartet:      $(Format-EobReportDate -Value $Report.StartedAt)")
    $lines.Add("Abgeschlossen:  $(Format-EobReportDate -Value $Report.CompletedAt)")
    $lines.Add("Ergebnis:       $(Get-EobStatusText -Status $Report.Outcome)")
    if ($Report.Summary.Count -gt 0) {
        $lines.Add('')
        $lines.Add('Zusammenfassung')
        $lines.Add('---------------')
        foreach ($key in $Report.Summary.Keys) { $lines.Add(("{0,-16} {1}" -f "$($key):", $Report.Summary[$key])) }
    }
    $lines.Add('')
    $lines.Add('Schritte')
    $lines.Add('--------')
    foreach ($step in $Report.Steps) {
        $phase = if ($step.Phase) { " [$($step.Phase)]" } else { '' }
        $lines.Add(("{0}{1} {2} -> {3}: {4}" -f $step.Id, $phase, $step.Title, (Get-EobStatusText -Status $step.Status), $step.Message))
    }
    if (@($Report.Findings).Count -gt 0) {
        $lines.Add('')
        $lines.Add('Hinweise')
        $lines.Add('--------')
        foreach ($finding in $Report.Findings) { $lines.Add(("[{0}] {1}: {2}" -f $finding.Severity, $finding.Code, $finding.Message)) }
    }
    $lines.Add('')
    $lines.Add("Erstellt mit $($Report.Tool). Kennwörter sind nicht Bestandteil dieses Berichts.")
    return ($lines -join [Environment]::NewLine)
}

function Export-EobReport {
    <#
    .SYNOPSIS
        Schreibt einen Vorgangsbericht in die gewünschten Formate.
    .DESCRIPTION
        Html, Json, Csv (Schritte, Semikolon, UTF-8 mit BOM, gegen Formel-Injektion geschützt), Txt, Pdf.
        Pdf setzt eine PDF-Engine voraus; ist keine verfügbar, bleibt die HTML-Datei erhalten und
        das Ergebnis enthält eine Warnung.
    .PARAMETER Report
        Berichtsmodell (New-EobPlanReport).
    .PARAMETER Format
        Formate; Standard aus [Report] Formats.
    .PARAMETER Directory
        Zielverzeichnis; Standard: <ReportPath>\<Vorgangsart>.
    .OUTPUTS
        Objekte mit Format, Path und Message.
    #>
    [CmdletBinding(SupportsShouldProcess)]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory)][pscustomobject]$Report,
        [ValidateSet('Html', 'Json', 'Csv', 'Txt', 'Pdf')][string[]]$Format,
        [string]$Directory,
        [string]$BaseName,
        [AllowNull()][pscustomobject]$Config
    )

    if (-not $PSBoundParameters.ContainsKey('Format')) {
        $configured = @(Get-EobConfigValue -Config $Config -Section 'Report' -Key 'Formats' -As List)
        $Format = @($configured | Where-Object { $_ -in $script:ReportFormats })
        if (Get-EobConfigValue -Config $Config -Section 'Report' -Key 'UserOnboardingCreateTXT' -As Bool) { $Format += 'Txt' }
        if ($Format.Count -eq 0) { $Format = @('Html', 'Json') }
    }
    $Format = @($Format | Select-Object -Unique)
    if (-not $Directory) { $Directory = Get-EobReportDirectory -Config $Config -Kind $Report.Kind }
    if (-not $BaseName) {
        $stamp = if ($Report.CreatedAt -is [datetime]) { $Report.CreatedAt.ToString('yyyyMMdd-HHmmss') } else { (Get-Date).ToString('yyyyMMdd-HHmmss') }
        $name = if ($Report.SubjectName) { $Report.SubjectName } else { $Report.Kind }
        $suffix = if ($Report.Simulation) { '_Simulation' } else { '' }
        $BaseName = "{0}_{1}_{2}{3}" -f $stamp, $name, ([string]$Report.OperationId).Split('-')[0], $suffix
    }
    $BaseName = Get-EobSafeFileName -Name $BaseName -MaxLength 120

    if (-not $PSCmdlet.ShouldProcess($Directory, "Bericht '$BaseName' ($($Format -join ', ')) schreiben")) {
        return
    }

    $htmlPath = Join-Path -Path $Directory -ChildPath "$BaseName.html"
    $needsHtml = ($Format -contains 'Html') -or ($Format -contains 'Pdf')
    if ($needsHtml) {
        Write-EobTextFile -Path $htmlPath -Content (ConvertTo-EobReportHtml -Report $Report -Config $Config)
        if ($Format -contains 'Html') {
            [pscustomobject]@{ Format = 'Html'; Path = $htmlPath; Message = '' }
        }
    }
    if ($Format -contains 'Json') {
        $jsonPath = Join-Path -Path $Directory -ChildPath "$BaseName.json"
        Write-EobTextFile -Path $jsonPath -Content (ConvertTo-Json -InputObject $Report -Depth 8)
        [pscustomobject]@{ Format = 'Json'; Path = $jsonPath; Message = '' }
    }
    if ($Format -contains 'Csv') {
        $csvPath = Join-Path -Path $Directory -ChildPath "$BaseName.csv"
        $rows = foreach ($step in $Report.Steps) {
            [pscustomobject][ordered]@{
                OperationId = $Report.OperationId
                Schritt     = $step.Id
                Phase       = $step.Phase
                Aktion      = $step.Action
                Titel       = $step.Title
                Ziel        = $step.Target
                Risiko      = $step.Risk
                Status      = $step.Status
                Meldung     = $step.Message
                DauerMs     = $step.DurationMs
            }
        }
        $directoryPath = Split-Path -Parent $csvPath
        if (-not (Test-Path -LiteralPath $directoryPath)) { $null = New-Item -ItemType Directory -Path $directoryPath -Force }
        @($rows) | ConvertTo-EobCsvSafeObject | Export-Csv -LiteralPath $csvPath -NoTypeInformation -Delimiter ';' -Encoding utf8BOM
        [pscustomobject]@{ Format = 'Csv'; Path = $csvPath; Message = '' }
    }
    if ($Format -contains 'Txt') {
        $txtPath = Join-Path -Path $Directory -ChildPath "$BaseName.txt"
        Write-EobTextFile -Path $txtPath -Content (ConvertTo-EobReportText -Report $Report) -Bom
        [pscustomobject]@{ Format = 'Txt'; Path = $txtPath; Message = '' }
    }
    if ($Format -contains 'Pdf') {
        $pdfPath = Join-Path -Path $Directory -ChildPath "$BaseName.pdf"
        $pdf = ConvertTo-EobPdf -HtmlPath $htmlPath -PdfPath $pdfPath -Config $Config -Confirm:$false
        if ($pdf.Status -eq 'Succeeded') {
            [pscustomobject]@{ Format = 'Pdf'; Path = $pdfPath; Message = '' }
        }
        else {
            [pscustomobject]@{ Format = 'Pdf'; Path = ''; Message = $pdf.Message }
            if ($Format -notcontains 'Html') {
                [pscustomobject]@{ Format = 'Html'; Path = $htmlPath; Message = 'HTML statt PDF (keine PDF-Engine verfügbar).' }
            }
        }
    }
}

function Export-EobPlanReport {
    <#
    .SYNOPSIS
        Erzeugt Berichtsmodell und Dateien eines Plans in einem Schritt.
    #>
    [CmdletBinding(SupportsShouldProcess)]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory)][pscustomobject]$Plan,
        [ValidateSet('Html', 'Json', 'Csv', 'Txt', 'Pdf')][string[]]$Format,
        [string]$Directory,
        [AllowNull()][pscustomobject]$Config
    )

    if ($null -eq $Config) { $Config = $Plan.Config }
    $report = New-EobPlanReport -Plan $Plan
    $parameters = @{ Report = $report; Config = $Config; Confirm = $false }
    if ($PSBoundParameters.ContainsKey('Format')) { $parameters['Format'] = $Format }
    if ($Directory) { $parameters['Directory'] = $Directory }
    if (-not $PSCmdlet.ShouldProcess($Plan.OperationId, 'Vorgangsbericht schreiben')) { return }
    $files = @(Export-EobReport @parameters)
    Write-EobLog -Level Information -OperationId $Plan.OperationId -Action 'ReportWritten' -Target $report.SubjectName `
        -Message ("Bericht geschrieben: {0}" -f ((@($files | Where-Object Path) | ForEach-Object Path) -join ', '))
    return $files
}

#endregion

#region PDF

function Find-EobEdgeExecutable {
    [CmdletBinding()]
    [OutputType([string])]
    param([AllowNull()][pscustomobject]$Config)

    $configured = [string](Get-EobConfigValue -Config $Config -Section 'Report' -Key 'EdgePath' -As Path)
    if ($configured) {
        if (Test-Path -LiteralPath $configured -PathType Leaf) { return $configured }
        return ''
    }
    $candidates = [System.Collections.Generic.List[string]]::new()
    foreach ($base in @(${env:ProgramFiles(x86)}, $env:ProgramFiles, $env:LOCALAPPDATA)) {
        if ($base) { $candidates.Add((Join-Path -Path $base -ChildPath 'Microsoft\Edge\Application\msedge.exe')) }
    }
    foreach ($candidate in $candidates) {
        if (Test-Path -LiteralPath $candidate -PathType Leaf) { return $candidate }
    }
    foreach ($name in @('msedge', 'microsoft-edge', 'microsoft-edge-stable')) {
        $command = Get-Command -Name $name -CommandType Application -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($null -ne $command) { return $command.Source }
    }
    return ''
}

function Find-EobWkhtmltopdfExecutable {
    [CmdletBinding()]
    [OutputType([string])]
    param([AllowNull()][pscustomobject]$Config)

    $configured = [string](Get-EobConfigValue -Config $Config -Section 'Report' -Key 'wkhtmltopdfPath' -As Path)
    if ($configured) {
        if (Test-Path -LiteralPath $configured -PathType Leaf) { return $configured }
        return ''
    }
    if ($env:ProgramFiles) {
        $candidate = Join-Path -Path $env:ProgramFiles -ChildPath 'wkhtmltopdf\bin\wkhtmltopdf.exe'
        if (Test-Path -LiteralPath $candidate -PathType Leaf) { return $candidate }
    }
    $command = Get-Command -Name 'wkhtmltopdf' -CommandType Application -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($null -ne $command) { return $command.Source }
    return ''
}

function Get-EobPdfEngine {
    <#
    .SYNOPSIS
        Ermittelt die PDF-Engine gemäß [Report] PdfEngine (Auto: Edge, dann wkhtmltopdf).
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param([AllowNull()][pscustomobject]$Config)

    $preference = [string](Get-EobConfigValue -Config $Config -Section 'Report' -Key 'PdfEngine')
    if ($preference -eq 'None') {
        return [pscustomobject]@{ Name = 'None'; Path = ''; Preference = $preference; Reason = 'PDF-Erzeugung ist deaktiviert ([Report] PdfEngine=None).' }
    }
    $edge = if ($preference -in @('Auto', 'Edge')) { Find-EobEdgeExecutable -Config $Config } else { '' }
    $wkhtml = if ($preference -in @('Auto', 'Wkhtmltopdf')) { Find-EobWkhtmltopdfExecutable -Config $Config } else { '' }
    if ($edge) { return [pscustomobject]@{ Name = 'Edge'; Path = $edge; Preference = $preference; Reason = '' } }
    if ($wkhtml) { return [pscustomobject]@{ Name = 'Wkhtmltopdf'; Path = $wkhtml; Preference = $preference; Reason = '' } }
    $reason = switch ($preference) {
        'Edge' { 'Microsoft Edge wurde nicht gefunden.' }
        'Wkhtmltopdf' { 'wkhtmltopdf wurde nicht gefunden ([Report] wkhtmltopdfPath).' }
        default { 'Weder Microsoft Edge noch wkhtmltopdf wurden gefunden.' }
    }
    return [pscustomobject]@{ Name = 'None'; Path = ''; Preference = $preference; Reason = $reason }
}

function Invoke-EobExternalProcess {
    <#
        Startet ein externes Programm mit Argumentliste (ohne Shell) und Zeitlimit.
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory)][string]$FilePath,
        [Parameter(Mandatory)][string[]]$ArgumentList,
        [ValidateRange(1, 600)][int]$TimeoutSeconds = 60
    )

    $startInfo = [System.Diagnostics.ProcessStartInfo]::new($FilePath)
    foreach ($argument in $ArgumentList) { $startInfo.ArgumentList.Add($argument) }
    $startInfo.UseShellExecute = $false
    $startInfo.CreateNoWindow = $true
    $startInfo.RedirectStandardOutput = $true
    $startInfo.RedirectStandardError = $true

    $process = [System.Diagnostics.Process]::new()
    $process.StartInfo = $startInfo
    try {
        $null = $process.Start()
        $stdout = $process.StandardOutput.ReadToEndAsync()
        $stderr = $process.StandardError.ReadToEndAsync()
        if (-not $process.WaitForExit($TimeoutSeconds * 1000)) {
            try { $process.Kill($true) } catch { Write-Verbose "Prozess konnte nicht beendet werden: $($_.Exception.Message)" }
            return [pscustomobject]@{ ExitCode = -1; TimedOut = $true; Output = ''; ErrorOutput = "Zeitlimit von $TimeoutSeconds s überschritten." }
        }
        $process.WaitForExit()
        return [pscustomobject]@{ ExitCode = $process.ExitCode; TimedOut = $false; Output = $stdout.Result; ErrorOutput = $stderr.Result }
    }
    finally {
        $process.Dispose()
    }
}

function ConvertTo-EobPdf {
    <#
    .SYNOPSIS
        Wandelt eine HTML-Datei in PDF um (Edge headless oder wkhtmltopdf).
    .DESCRIPTION
        Edge läuft mit einem temporären, isolierten Profil. Programme werden ohne Shell mit
        Argumentliste gestartet. Ist keine Engine verfügbar, wird eine Warnung zurückgegeben.
    #>
    [CmdletBinding(SupportsShouldProcess)]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory)][string]$HtmlPath,
        [Parameter(Mandatory)][string]$PdfPath,
        [AllowNull()][pscustomobject]$Config
    )

    if (-not (Test-Path -LiteralPath $HtmlPath -PathType Leaf)) {
        return New-EobResult -Status Warning -Message "PDF nicht erzeugt: HTML-Datei '$HtmlPath' fehlt."
    }
    $engine = Get-EobPdfEngine -Config $Config
    if ($engine.Name -eq 'None') {
        return New-EobResult -Status Warning -Message "PDF nicht erzeugt: $($engine.Reason)"
    }
    if (-not $PSCmdlet.ShouldProcess($PdfPath, "PDF mit $($engine.Name) erzeugen")) {
        return New-EobNotProcessedResult -Message "PDF würde mit $($engine.Name) erzeugt: $PdfPath"
    }

    $timeout = [int](Get-EobConfigValue -Config $Config -Section 'Report' -Key 'PdfTimeoutSeconds' -As Int)
    if ($timeout -lt 5) { $timeout = 60 }
    $htmlFull = (Resolve-Path -LiteralPath $HtmlPath).ProviderPath
    $pdfFull = [System.IO.Path]::GetFullPath($PdfPath)
    if (Test-Path -LiteralPath $pdfFull) { Remove-Item -LiteralPath $pdfFull -Force }

    $profileDirectory = $null
    try {
        if ($engine.Name -eq 'Edge') {
            $profileDirectory = Join-Path -Path ([System.IO.Path]::GetTempPath()) -ChildPath ("easyONB-pdf-" + [guid]::NewGuid().ToString('N'))
            $null = New-Item -ItemType Directory -Path $profileDirectory -Force
            $arguments = @(
                '--headless', '--disable-gpu', '--no-first-run', '--no-default-browser-check', '--disable-extensions',
                '--disable-sync', '--no-pdf-header-footer', '--print-to-pdf-no-header', "--user-data-dir=$profileDirectory", "--print-to-pdf=$pdfFull",
                [System.Uri]::new($htmlFull, [System.UriKind]::Absolute).AbsoluteUri
            )
        }
        else {
            $arguments = @('--quiet', '--encoding', 'utf-8', '--page-size', 'A4', $htmlFull, $pdfFull)
        }
        $result = Invoke-EobExternalProcess -FilePath $engine.Path -ArgumentList $arguments -TimeoutSeconds $timeout
    }
    finally {
        if ($profileDirectory -and (Test-Path -LiteralPath $profileDirectory)) {
            Remove-Item -LiteralPath $profileDirectory -Recurse -Force -ErrorAction SilentlyContinue
        }
    }

    if ((Test-Path -LiteralPath $pdfFull -PathType Leaf) -and (Get-Item -LiteralPath $pdfFull).Length -gt 0) {
        return New-EobResult -Status Succeeded -Message "PDF erzeugt ($($engine.Name)): $pdfFull" -Data @{ PdfPath = $pdfFull }
    }
    $detail = if ($result.TimedOut) { $result.ErrorOutput } else { "Exitcode $($result.ExitCode)" }
    return New-EobResult -Status Warning -Message "PDF konnte nicht erzeugt werden ($($engine.Name), $detail). Die HTML-Datei bleibt erhalten."
}

#endregion

#region Willkommensdokument

function Get-EobAssetDataUri {
    <#
    .SYNOPSIS
        Liefert eine Bilddatei unterhalb von Assets als data-URI (Dokumente bleiben eigenständig).
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param([Parameter(Mandatory)][string]$RelativePath)

    $assetsRoot = Join-Path -Path (Get-EobAppRoot) -ChildPath 'Assets'
    $relative = $RelativePath.Trim().Replace('\', [System.IO.Path]::DirectorySeparatorChar).Replace('/', [System.IO.Path]::DirectorySeparatorChar)
    $fullPath = [System.IO.Path]::GetFullPath((Join-Path -Path $assetsRoot -ChildPath $relative))
    if (-not (Test-EobPathWithin -Path $fullPath -Root $assetsRoot)) {
        throw "Asset '$RelativePath' liegt außerhalb des Assets-Verzeichnisses."
    }
    $extension = [System.IO.Path]::GetExtension($fullPath).ToLowerInvariant()
    if (-not $script:AssetMimeTypes.ContainsKey($extension)) {
        throw "Asset '$RelativePath' hat keinen zulässigen Bildtyp."
    }
    if (-not (Test-Path -LiteralPath $fullPath -PathType Leaf)) {
        throw "Asset '$RelativePath' wurde nicht gefunden."
    }
    $file = Get-Item -LiteralPath $fullPath
    if ($file.Length -gt $script:MaxAssetBytes) {
        throw "Asset '$RelativePath' ist größer als 2 MB."
    }
    return ('data:{0};base64,{1}' -f $script:AssetMimeTypes[$extension], [Convert]::ToBase64String([System.IO.File]::ReadAllBytes($fullPath)))
}

function Get-EobWebsiteLinkHtml {
    [CmdletBinding()]
    [OutputType([string])]
    param([AllowNull()][pscustomobject]$Config)

    $section = Get-EobConfigSection -Config $Config -Name 'Websites'
    $items = [System.Collections.Generic.List[string]]::new()
    foreach ($key in @($section.Keys | Sort-Object)) {
        $parts = @(([string]$section[$key]) -split '\|' | ForEach-Object { $_.Trim() })
        if ($parts.Count -lt 2 -or -not ($parts[1] -match '^https?://')) { continue }
        $uri = $null
        if (-not [uri]::TryCreate($parts[1], [System.UriKind]::Absolute, [ref]$uri)) { continue }
        $description = if ($parts.Count -ge 3 -and $parts[2]) { ' - ' + (ConvertTo-EobHtmlEncoded -Value $parts[2]) } else { '' }
        $items.Add(('<li><a href="{0}">{1}</a>{2}</li>' -f (ConvertTo-EobHtmlEncoded -Value $uri.AbsoluteUri), (ConvertTo-EobHtmlEncoded -Value $parts[0]), $description))
    }
    if ($items.Count -eq 0) { return '' }
    return ('<ul>' + ($items -join '') + '</ul>')
}

function Get-EobWelcomePlaceholderValue {
    <#
    .SYNOPSIS
        Stellt die Platzhalterwerte des Willkommensdokuments zusammen.
    .DESCRIPTION
        Reihenfolge (spätere Quellen haben Vorrang): Berichtseinstellungen, Unternehmensdaten,
        [CompanyHelpdesk], [CompanyWLAN], [CompanyVPN], [ReportPlaceholders], Benutzerwerte (-Values).
        Konfigurationswerte mit Secret-Kennzeichnung werden nie übernommen.
    .PARAMETER Language
        de oder en (Kennwortkriterien).
    #>
    [CmdletBinding()]
    [OutputType([System.Collections.Generic.Dictionary[string, string]])]
    param(
        [AllowNull()][pscustomobject]$Config,
        [hashtable]$Values = @{},
        [ValidateSet('de', 'en')][string]$Language = 'de'
    )

    $result = [System.Collections.Generic.Dictionary[string, string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    $reportTitle = [string](Get-EobConfigValue -Config $Config -Section 'Report' -Key 'ReportHeader')
    if (-not $reportTitle) { $reportTitle = [string](Get-EobConfigValue -Config $Config -Section 'Report' -Key 'ReportTitle') }
    $result['ReportTitle'] = $reportTitle
    $result['ReportFooter'] = [string](Get-EobConfigValue -Config $Config -Section 'Report' -Key 'ReportFooter')
    $result['ReportDate'] = (Get-Date).ToString('yyyy-MM-dd', [System.Globalization.CultureInfo]::InvariantCulture)
    $result['Admin'] = (Get-EobCurrentIdentity).Name

    $companyId = [string]$Values['CompanyId']
    $company = if ($companyId) { Get-EobCompany -Config $Config -Id $companyId } else { Get-EobCompany -Config $Config | Select-Object -First 1 }
    if ($null -ne $company) {
        # Legacy-kompatibel: {{CompanyDomain}} entspricht [Company] CompanyDomain (Webadresse), sonst der Mail-Domäne.
        $website = [string]$company.Website
        $websiteUrl = if ($website -and $website -notmatch '^(?i)https?://') { "https://$website" } else { $website }
        $result['CompanyName'] = $company.DisplayName
        $result['CompanyStreet'] = $company.Street
        $result['CompanyZIP'] = $company.PostalCode
        $result['CompanyCity'] = $company.City
        $result['CompanyCountry'] = $company.Country
        $result['CompanyPhone'] = $company.Phone
        $result['CompanyDomain'] = if ($website) { $website } else { $company.MailDomain }
        $result['CompanyMailDomain'] = $company.MailDomain
        $result['CompanyWebsite'] = $websiteUrl
    }

    foreach ($sectionName in @('CompanyHelpdesk', 'CompanyWLAN', 'CompanyVPN', 'ReportPlaceholders')) {
        $section = Get-EobConfigSection -Config $Config -Name $sectionName
        foreach ($key in $section.Keys) {
            if ($key -match $script:SecretPlaceholderPattern) { continue }
            $value = [string](Get-EobConfigValue -Config $Config -Section $sectionName -Key $key)
            # Leere Einträge überschreiben keine vorhandenen Werte (z. B. Unternehmensdaten).
            if ($value -or -not $result.ContainsKey($key)) { $result[$key] = $value }
        }
    }

    foreach ($key in $Values.Keys) {
        if ($key -match $script:SecretPlaceholderPattern) { continue }
        $value = $Values[$key]
        $result[$key] = if ($null -eq $value) { '' } elseif ($value -is [datetime]) { $value.ToString('dd.MM.yyyy') } else { [string]$value }
    }

    $policy = Get-EobPasswordPolicy -Config $Config
    $result['PasswordPolicyList'] = (@(Get-EobPasswordPolicyText -Policy $policy -Language $Language) | ForEach-Object { '<li>' + (ConvertTo-EobHtmlEncoded -Value $_) + '</li>' }) -join ''
    $result['WebsitesHTML'] = Get-EobWebsiteLinkHtml -Config $Config

    $result['LogoTag'] = ''
    $logo = [string](Get-EobConfigValue -Config $Config -Section 'Report' -Key 'TemplateLogo' -As Path)
    if ($logo -and (Test-Path -LiteralPath $logo -PathType Leaf)) {
        $extension = [System.IO.Path]::GetExtension($logo).ToLowerInvariant()
        if ($script:AssetMimeTypes.ContainsKey($extension) -and (Get-Item -LiteralPath $logo).Length -le $script:MaxAssetBytes) {
            $dataUri = 'data:{0};base64,{1}' -f $script:AssetMimeTypes[$extension], [Convert]::ToBase64String([System.IO.File]::ReadAllBytes($logo))
            $result['LogoTag'] = '<img src="' + $dataUri + '" alt="Logo">'
        }
    }
    return , $result
}

function Expand-EobTemplate {
    <#
    .SYNOPSIS
        Ersetzt {{Platzhalter}} in einer Vorlage.
    .DESCRIPTION
        - Werte werden HTML-kodiert (Ausnahme: intern erzeugte HTML-Blöcke).
        - Kennwort-Platzhalter werden nie befüllt, sondern durch -SecretHint ersetzt.
        - {{AssetDataUri:Pfad}} bettet Bilder aus dem Assets-Verzeichnis ein.
        - Unbekannte Platzhalter werden entfernt und in MissingPlaceholders gemeldet.
    .OUTPUTS
        Objekt mit Content, MissingPlaceholders, SecretPlaceholders und Warnings.
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory)][AllowEmptyString()][string]$Template,
        [Parameter(Mandatory)][System.Collections.IDictionary]$Values,
        [string]$SecretHint = 'Wird separat übermittelt.',
        [switch]$PlainText
    )

    $missing = [System.Collections.Generic.List[string]]::new()
    $secrets = [System.Collections.Generic.List[string]]::new()
    $warnings = [System.Collections.Generic.List[string]]::new()
    $encode = -not $PlainText
    $lookup = [System.Collections.Generic.Dictionary[string, string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($key in $Values.Keys) { $lookup[[string]$key] = [string]$Values[$key] }

    $evaluator = [System.Text.RegularExpressions.MatchEvaluator] {
        param($match)
        $name = $match.Groups[1].Value
        $argument = $match.Groups[2].Value
        if ($name -ieq 'AssetDataUri') {
            try {
                return (Get-EobAssetDataUri -RelativePath $argument)
            }
            catch {
                $warnings.Add($_.Exception.Message)
                return ''
            }
        }
        if ($name -match $script:SecretPlaceholderPattern -and $name -notin $script:NonSecretPlaceholders) {
            if (-not $secrets.Contains($name)) { $secrets.Add($name) }
            if ($name -match '(?i)label') { return '' }
            return $(if ($encode) { ConvertTo-EobHtmlEncoded -Value $SecretHint } else { $SecretHint })
        }
        if ($lookup.ContainsKey($name)) {
            $value = $lookup[$name]
            if (-not $encode -or ($name -in $script:RawHtmlPlaceholders)) { return $value }
            return (ConvertTo-EobHtmlEncoded -Value $value)
        }
        if (-not $missing.Contains($name)) { $missing.Add($name) }
        return ''
    }
    $content = [regex]::Replace($Template, $script:PlaceholderPattern, $evaluator)

    [pscustomobject]@{
        Content             = $content
        MissingPlaceholders = $missing.ToArray()
        SecretPlaceholders  = $secrets.ToArray()
        Warnings            = $warnings.ToArray()
    }
}

function New-EobWelcomeDocument {
    <#
    .SYNOPSIS
        Erzeugt das Willkommensdokument (HTML, optional PDF) aus der HTML-Vorlage - ohne Kennwort.
    .DESCRIPTION
        Plan-Handler des Onboardings. In der Simulation wird die Vorlage gerendert und geprüft,
        aber keine Datei geschrieben.
    .PARAMETER Values
        Benutzerbezogene Platzhalterwerte (Vorname, Nachname, LoginName, ...).
    .PARAMETER OutputDirectory
        Zielverzeichnis; Standard: <ReportPath>\Welcome.
    #>
    [CmdletBinding(SupportsShouldProcess)]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory)][string]$SamAccountName,
        [Parameter(Mandatory)][string]$TemplatePath,
        [hashtable]$Values = @{},
        [string]$OperationId = '',
        [AllowNull()][pscustomobject]$Config,
        [string]$OutputDirectory
    )

    if (-not (Test-Path -LiteralPath $TemplatePath -PathType Leaf)) {
        throw "Vorlage '$TemplatePath' wurde nicht gefunden."
    }
    $template = (Get-EobTextFileContent -Path $TemplatePath).Text
    $language = if ($template -match '<html[^>]*\blang\s*=\s*"en') { 'en' } else { 'de' }
    $hint = if ($language -eq 'en') { 'Provided to you separately and in person.' } else { 'Erhältst Du separat und persönlich.' }
    $placeholderValues = Get-EobWelcomePlaceholderValue -Config $Config -Values $Values -Language $language
    $rendered = Expand-EobTemplate -Template $template -Values $placeholderValues -SecretHint $hint

    if (-not $OutputDirectory) { $OutputDirectory = Get-EobReportDirectory -Config $Config -Kind 'Welcome' }
    $baseName = Get-EobSafeFileName -Name ("{0}_Willkommen_{1}" -f $SamAccountName, (Get-Date).ToString('yyyyMMdd-HHmmss'))
    $htmlPath = Join-Path -Path $OutputDirectory -ChildPath "$baseName.html"

    $notes = [System.Collections.Generic.List[string]]::new()
    if ($rendered.MissingPlaceholders.Count -gt 0) { $notes.Add("Platzhalter ohne Wert: $($rendered.MissingPlaceholders -join ', ')") }
    if ($rendered.SecretPlaceholders.Count -gt 0) { $notes.Add("Kennwort-Platzhalter neutralisiert: $($rendered.SecretPlaceholders -join ', ')") }
    foreach ($warning in $rendered.Warnings) { $notes.Add($warning) }
    $noteText = if ($notes.Count -gt 0) { ' (' + ($notes -join '; ') + ')' } else { '' }

    if (-not $PSCmdlet.ShouldProcess($htmlPath, 'Willkommensdokument erzeugen')) {
        return New-EobNotProcessedResult -Message "Willkommensdokument würde erzeugt: $htmlPath$noteText"
    }

    Write-EobTextFile -Path $htmlPath -Content $rendered.Content
    $data = @{ WelcomeDocumentPath = $htmlPath }
    $status = if ($rendered.MissingPlaceholders.Count -gt 0 -or $rendered.Warnings.Count -gt 0) { 'Warning' } else { 'Succeeded' }
    $message = "Willkommensdokument erstellt: $htmlPath$noteText"

    $formats = @(Get-EobConfigValue -Config $Config -Section 'Report' -Key 'Formats' -As List)
    if ($formats -contains 'Pdf') {
        $pdf = ConvertTo-EobPdf -HtmlPath $htmlPath -PdfPath ([System.IO.Path]::ChangeExtension($htmlPath, '.pdf')) -Config $Config -Confirm:$false
        if ($pdf.Status -eq 'Succeeded') {
            $data['WelcomeDocumentPdfPath'] = $pdf.Data['PdfPath']
            $message += " | PDF: $($pdf.Data['PdfPath'])"
        }
        else {
            $status = 'Warning'
            $message += " | $($pdf.Message)"
        }
    }
    Write-EobLog -Level Information -OperationId $OperationId -Action 'WelcomeDocument' -Target $SamAccountName -Result $status -Message $message
    return New-EobResult -Status $status -Message $message -Data $data
}

#endregion

#region Welcome-Mail

function Get-EobSmtpSetting {
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param([AllowNull()][pscustomobject]$Config)

    $from = [string](Get-EobConfigValue -Config $Config -Section 'EmailSettings' -Key 'FromAddress')
    if (-not $from) { $from = [string](Get-EobConfigValue -Config $Config -Section 'EmailSettings' -Key 'From') }
    [pscustomobject]@{
        Server                = ([string](Get-EobConfigValue -Config $Config -Section 'EmailSettings' -Key 'SMTPServer')).Trim()
        Port                  = [int](Get-EobConfigValue -Config $Config -Section 'EmailSettings' -Key 'SMTPPort' -As Int)
        UseSsl                = [bool](Get-EobConfigValue -Config $Config -Section 'EmailSettings' -Key 'UseSSL' -As Bool)
        UseDefaultCredentials = [bool](Get-EobConfigValue -Config $Config -Section 'EmailSettings' -Key 'UseDefaultCredentials' -As Bool)
        From                  = $from.Trim()
        Copy                  = ([string](Get-EobConfigValue -Config $Config -Section 'EmailSettings' -Key 'CopyAddress')).Trim()
        Subject               = [string](Get-EobConfigValue -Config $Config -Section 'EmailSettings' -Key 'WelcomeEmailSubject')
        TemplatePath          = [string](Get-EobConfigValue -Config $Config -Section 'EmailSettings' -Key 'WelcomeEmailTemplate' -As Path)
    }
}

function Get-EobDefaultWelcomeMailBody {
    [CmdletBinding()]
    [OutputType([string])]
    param()

    return @'
<!DOCTYPE html>
<html lang="de"><head><meta charset="utf-8"><title>{{ReportTitle}}</title></head>
<body style="font-family: Segoe UI, Arial, sans-serif; font-size: 14px; color: #1f2328;">
<p>Hallo {{DisplayName}},</p>
<p>herzlich willkommen bei {{CompanyName}}! Dein Benutzerkonto ist eingerichtet.</p>
<table style="border-collapse: collapse;">
<tr><td style="padding: 2px 12px 2px 0;">Anmeldename</td><td><b>{{LoginName}}</b></td></tr>
<tr><td style="padding: 2px 12px 2px 0;">Anmeldung (UPN)</td><td>{{UPN}}</td></tr>
<tr><td style="padding: 2px 12px 2px 0;">E-Mail-Adresse</td><td>{{MailAddress}}</td></tr>
<tr><td style="padding: 2px 12px 2px 0;">Startdatum</td><td>{{StartDate}}</td></tr>
</table>
<p>Dein Kennwort für die erste Anmeldung erhältst Du separat und persönlich. Bitte ändere es bei der ersten Anmeldung.</p>
<p>Bei Fragen hilft Dir der IT-Support: {{CompanyHelpdeskMail}} {{CompanyHelpdeskTel}}</p>
<p>Viele Grüße<br>{{CompanyITMitarbeiter}}</p>
</body></html>
'@
}

function Send-EobSmtpMessage {
    <#
        Versendet eine Nachricht über System.Net.Mail (kapselt den Transport, in Tests gemockt).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][pscustomobject]$Setting,
        [Parameter(Mandatory)][string]$To,
        [Parameter(Mandatory)][string]$Subject,
        [Parameter(Mandatory)][string]$HtmlBody
    )

    $message = [System.Net.Mail.MailMessage]::new()
    $client = [System.Net.Mail.SmtpClient]::new($Setting.Server, $Setting.Port)
    try {
        $message.From = [System.Net.Mail.MailAddress]::new($Setting.From)
        $message.To.Add([System.Net.Mail.MailAddress]::new($To))
        if ($Setting.Copy) { $message.CC.Add([System.Net.Mail.MailAddress]::new($Setting.Copy)) }
        $message.Subject = $Subject
        $message.SubjectEncoding = [System.Text.Encoding]::UTF8
        $message.Body = $HtmlBody
        $message.BodyEncoding = [System.Text.Encoding]::UTF8
        $message.IsBodyHtml = $true
        $client.EnableSsl = $Setting.UseSsl
        $client.UseDefaultCredentials = $Setting.UseDefaultCredentials
        $client.Timeout = 30000
        $client.Send($message)
    }
    finally {
        $message.Dispose()
        $client.Dispose()
    }
}

function Send-EobWelcomeMail {
    <#
    .SYNOPSIS
        Versendet die Welcome-Mail an den neuen Benutzer - ohne Kennwort.
    .DESCRIPTION
        SMTP über [EmailSettings]. Anmeldedaten für das Relay werden nicht aus der Konfiguration
        gelesen; unterstützt werden anonyme interne Relays und die Windows-Anmeldung
        (UseDefaultCredentials=1). Kennwort-Platzhalter eigener Vorlagen werden neutralisiert.
    #>
    [CmdletBinding(SupportsShouldProcess)]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory)][string]$To,
        [string]$DisplayName = '',
        [string]$SamAccountName = '',
        [string]$UserPrincipalName = '',
        [AllowNull()][object]$StartDate,
        [AllowNull()][pscustomobject]$Config,
        [string]$CompanyId = ''
    )

    if (-not (Test-EobEmailAddress -Value $To)) {
        throw "Ungültige Empfängeradresse '$To'."
    }
    $setting = Get-EobSmtpSetting -Config $Config
    if (-not $setting.Server -or -not $setting.From) {
        return New-EobResult -Status Warning -Message 'Welcome-Mail nicht gesendet: [EmailSettings] SMTPServer und FromAddress sind nicht konfiguriert.'
    }
    if (-not (Test-EobEmailAddress -Value $setting.From)) {
        return New-EobResult -Status Warning -Message "Welcome-Mail nicht gesendet: Absenderadresse '$($setting.From)' ist ungültig."
    }

    $template = Get-EobDefaultWelcomeMailBody
    if ($setting.TemplatePath) {
        if (-not (Test-Path -LiteralPath $setting.TemplatePath -PathType Leaf)) {
            return New-EobResult -Status Warning -Message "Welcome-Mail nicht gesendet: Vorlage '$($setting.TemplatePath)' fehlt."
        }
        $template = (Get-EobTextFileContent -Path $setting.TemplatePath).Text
    }
    $startText = ''
    if ($null -ne $StartDate) {
        $parsed = if ($StartDate -is [datetime]) { $StartDate } else { ConvertTo-EobDate -Value ([string]$StartDate) }
        if ($null -ne $parsed) { $startText = $parsed.ToString('dd.MM.yyyy') }
    }
    $values = @{
        DisplayName = $DisplayName; LoginName = $SamAccountName; UPN = $UserPrincipalName; MailAddress = $To
        StartDate = $startText; CompanyId = $CompanyId
    }
    $placeholderValues = Get-EobWelcomePlaceholderValue -Config $Config -Values $values
    $body = Expand-EobTemplate -Template $template -Values $placeholderValues -SecretHint 'Erhältst Du separat und persönlich.'
    $subjectTemplate = if ($setting.Subject) { $setting.Subject } else { 'Willkommen' }
    $subject = (Expand-EobTemplate -Template $subjectTemplate -Values $placeholderValues -PlainText -SecretHint '').Content
    $subject = [regex]::Replace($subject, '[\x00-\x1F\x7F]', ' ').Trim()

    if (-not $PSCmdlet.ShouldProcess($To, "Welcome-Mail über $($setting.Server) senden")) {
        return New-EobNotProcessedResult -Message "Welcome-Mail würde an $To gesendet (Betreff: $subject)."
    }
    Send-EobSmtpMessage -Setting $setting -To $To -Subject $subject -HtmlBody $body.Content
    $message = "Welcome-Mail an $To gesendet."
    if ($body.MissingPlaceholders.Count -gt 0) { $message += " Platzhalter ohne Wert: $($body.MissingPlaceholders -join ', ')" }
    return New-EobResult -Status Succeeded -Message $message
}

#endregion

#region Audit-Auswertung

function Get-EobAuditSummary {
    <#
    .SYNOPSIS
        Fasst das Audit-Log zusammen (Dashboard: letzte Vorgänge, Fehler, Zähler je Vorgangsart).
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [datetime]$From = (Get-Date).AddDays(-30),
        [datetime]$To = (Get-Date),
        [ValidateRange(1, 500)][int]$Last = 20,
        [string]$Directory
    )

    $parameters = @{ From = $From; To = $To }
    if ($Directory) { $parameters['Directory'] = $Directory }
    $entries = @(Get-EobAuditEntry @parameters)
    $completed = @($entries | Where-Object { [string]$_.Action -like '*Completed' })
    $operations = foreach ($entry in $completed) {
        [pscustomobject]@{
            Timestamp   = $entry.Timestamp
            Kind        = ([string]$entry.Action) -replace 'Completed$', ''
            Target      = [string]$entry.Target
            Result      = [string]$entry.Result
            Actor       = [string]$entry.Actor
            OperationId = [string]$entry.OperationId
        }
    }
    $byKind = [ordered]@{}
    foreach ($group in @($operations | Group-Object -Property Kind | Sort-Object -Property Name)) {
        $byKind[$group.Name] = $group.Count
    }
    [pscustomobject]@{
        From             = $From
        To               = $To
        EntryCount       = $entries.Count
        OperationCount   = @($operations).Count
        FailedOperations = @($operations | Where-Object Result -EQ 'Failed').Count
        FailedSteps      = @($entries | Where-Object { [string]$_.Result -eq 'Failed' -and [string]$_.Action -notlike '*Completed' }).Count
        ByKind           = $byKind
        LastOperations   = @($operations | Sort-Object -Property Timestamp -Descending | Select-Object -First $Last)
    }
}

function Export-EobAuditReport {
    <#
    .SYNOPSIS
        Exportiert Audit-Einträge eines Zeitraums (Csv, Json oder Html).
    #>
    [CmdletBinding(SupportsShouldProcess)]
    [OutputType([string])]
    param(
        [Parameter(Mandatory)][string]$Path,
        [datetime]$From = (Get-Date).AddDays(-30),
        [datetime]$To = (Get-Date),
        [string]$OperationId,
        [ValidateSet('Csv', 'Json', 'Html')][string]$Format = 'Csv',
        [string]$Directory
    )

    $parameters = @{ From = $From; To = $To }
    if ($OperationId) { $parameters['OperationId'] = $OperationId }
    if ($Directory) { $parameters['Directory'] = $Directory }
    $entries = @(Get-EobAuditEntry @parameters)
    $rows = foreach ($entry in $entries) {
        [pscustomobject][ordered]@{
            Zeitpunkt   = [string]$entry.Timestamp
            OperationId = [string]$entry.OperationId
            Akteur      = [string]$entry.Actor
            Computer    = [string]$entry.Computer
            Aktion      = [string]$entry.Action
            Ziel        = [string]$entry.Target
            Ergebnis    = [string]$entry.Result
            DauerMs     = $entry.DurationMs
            Meldung     = Protect-EobSensitiveText -Text ([string]$entry.Message)
        }
    }
    if (-not $PSCmdlet.ShouldProcess($Path, "Audit-Export ($($entries.Count) Einträge)")) { return $Path }
    $parent = Split-Path -Parent $Path
    if ($parent -and -not (Test-Path -LiteralPath $parent)) { $null = New-Item -ItemType Directory -Path $parent -Force }
    switch ($Format) {
        'Csv' { @($rows) | ConvertTo-EobCsvSafeObject | Export-Csv -LiteralPath $Path -NoTypeInformation -Delimiter ';' -Encoding utf8BOM }
        'Json' { Write-EobTextFile -Path $Path -Content (ConvertTo-Json -InputObject @($rows) -Depth 4) }
        'Html' {
            $builder = [System.Text.StringBuilder]::new()
            [void]$builder.AppendLine('<!DOCTYPE html><html lang="de"><head><meta charset="utf-8"><title>Audit-Export</title>')
            [void]$builder.AppendLine('<style>body{font-family:"Segoe UI",Arial,sans-serif;font-size:12px;margin:20px}table{border-collapse:collapse;width:100%}th,td{border-bottom:1px solid #d0d7de;padding:4px 6px;text-align:left;vertical-align:top}th{background:#f6f8fa}</style></head><body>')
            [void]$builder.AppendLine(("<h1>Audit-Export</h1><p>{0} bis {1}, {2} Einträge</p>" -f (ConvertTo-EobHtmlEncoded -Value (Format-EobReportDate -Value $From)), (ConvertTo-EobHtmlEncoded -Value (Format-EobReportDate -Value $To)), $entries.Count))
            [void]$builder.AppendLine('<table><thead><tr><th>Zeitpunkt</th><th>Operation</th><th>Akteur</th><th>Aktion</th><th>Ziel</th><th>Ergebnis</th><th>Meldung</th></tr></thead><tbody>')
            foreach ($row in $rows) {
                [void]$builder.AppendLine(("<tr><td>{0}</td><td>{1}</td><td>{2}</td><td>{3}</td><td>{4}</td><td>{5}</td><td>{6}</td></tr>" -f `
                            (ConvertTo-EobHtmlEncoded -Value $row.Zeitpunkt), (ConvertTo-EobHtmlEncoded -Value $row.OperationId), (ConvertTo-EobHtmlEncoded -Value $row.Akteur),
                        (ConvertTo-EobHtmlEncoded -Value $row.Aktion), (ConvertTo-EobHtmlEncoded -Value $row.Ziel), (ConvertTo-EobHtmlEncoded -Value $row.Ergebnis), (ConvertTo-EobHtmlEncoded -Value $row.Meldung)))
            }
            [void]$builder.AppendLine('</tbody></table></body></html>')
            Write-EobTextFile -Path $Path -Content $builder.ToString()
        }
    }
    return $Path
}

function Get-EobReportFile {
    <#
    .SYNOPSIS
        Listet erzeugte Berichte (neueste zuerst), z. B. für die Ansicht "Reports und Audit".
    #>
    [CmdletBinding()]
    [OutputType([System.IO.FileInfo])]
    param(
        [AllowNull()][pscustomobject]$Config,
        [string]$Kind,
        [ValidateRange(1, 3650)][int]$Days = 30,
        [ValidateRange(1, 5000)][int]$First = 200
    )

    $directory = Get-EobReportDirectory -Config $Config -Kind $Kind
    if (-not (Test-Path -LiteralPath $directory -PathType Container)) { return }
    $since = (Get-Date).AddDays(-$Days)
    Get-ChildItem -LiteralPath $directory -File -Recurse -Include '*.html', '*.pdf', '*.json', '*.csv', '*.txt' -ErrorAction SilentlyContinue |
        Where-Object { $_.LastWriteTime -ge $since } |
        Sort-Object -Property LastWriteTime -Descending |
        Select-Object -First $First
}

#endregion

#region Integrationsstatus

function Get-EobSmtpStatus {
    <#
    .SYNOPSIS
        Status des SMTP-Versands. Mit -TestConnection wird eine TCP-Verbindung zum Relay geprüft.
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [AllowNull()][object]$Config,
        [switch]$TestConnection,
        [ValidateRange(200, 30000)][int]$TimeoutMilliseconds = 3000
    )

    $setting = Get-EobSmtpSetting -Config $Config
    if (-not $setting.Server) {
        return New-EobIntegrationStatus -Name 'SMTP' -State NotConfigured -Detail 'Kein SMTP-Relay konfiguriert.' -Hint 'Für Welcome-Mails [EmailSettings] SMTPServer und FromAddress setzen.'
    }
    if (-not $setting.From) {
        return New-EobIntegrationStatus -Name 'SMTP' -State NotConfigured -Detail "Relay $($setting.Server), aber keine Absenderadresse." -Hint '[EmailSettings] FromAddress setzen.'
    }
    $detail = "{0}:{1}{2}" -f $setting.Server, $setting.Port, $(if ($setting.UseSsl) { ' (TLS)' } else { ' (ohne TLS)' })
    if (-not $TestConnection) {
        return New-EobIntegrationStatus -Name 'SMTP' -State Available -Detail $detail -Hint 'Verbindung nicht geprüft.'
    }
    $client = [System.Net.Sockets.TcpClient]::new()
    try {
        $task = $client.ConnectAsync($setting.Server, $setting.Port)
        if ($task.Wait($TimeoutMilliseconds) -and $client.Connected) {
            return New-EobIntegrationStatus -Name 'SMTP' -State Connected -Detail $detail
        }
        return New-EobIntegrationStatus -Name 'SMTP' -State Error -Detail "$detail nicht erreichbar (Zeitlimit)." -Hint 'Firewall, DNS und Port prüfen.'
    }
    catch {
        $reason = if ($null -ne $_.Exception.InnerException) { $_.Exception.InnerException.Message } else { $_.Exception.Message }
        return New-EobIntegrationStatus -Name 'SMTP' -State Error -Detail "$detail nicht erreichbar: $reason" -Hint 'Firewall, DNS und Port prüfen.'
    }
    finally {
        $client.Dispose()
    }
}

function Get-EobFileServerStatus {
    <#
    .SYNOPSIS
        Status der Dateiserver-Stammpfade ([FileServer] AllowedRoots).
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [AllowNull()][object]$Config,
        [switch]$TestConnection
    )

    $roots = @(Get-EobConfigValue -Config $Config -Section 'FileServer' -Key 'AllowedRoots' -As List)
    if ($roots.Count -eq 0) {
        return New-EobIntegrationStatus -Name 'Dateiserver' -State NotConfigured -Detail 'Keine freigegebenen Stammpfade.' -Hint 'Home-Verzeichnisse werden nur unter [FileServer] AllowedRoots angelegt oder archiviert.'
    }
    if (-not $TestConnection) {
        return New-EobIntegrationStatus -Name 'Dateiserver' -State Available -Detail ("{0} Stammpfad(e) konfiguriert" -f $roots.Count) -Hint 'Erreichbarkeit nicht geprüft.'
    }
    $unreachable = @($roots | Where-Object { -not (Test-Path -LiteralPath $_ -PathType Container) })
    if ($unreachable.Count -eq 0) {
        return New-EobIntegrationStatus -Name 'Dateiserver' -State Connected -Detail ("{0} Stammpfad(e) erreichbar" -f $roots.Count)
    }
    return New-EobIntegrationStatus -Name 'Dateiserver' -State Error -Detail ("Nicht erreichbar: {0}" -f ($unreachable -join ', ')) -Hint 'Freigabe, Berechtigungen und Netzwerk prüfen.'
}

function Get-EobPdfEngineStatus {
    <#
    .SYNOPSIS
        Status der PDF-Erzeugung.
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param([AllowNull()][object]$Config)

    $engine = Get-EobPdfEngine -Config $Config
    if ($engine.Preference -eq 'None') {
        return New-EobIntegrationStatus -Name 'PDF-Erzeugung' -State Disabled -Detail $engine.Reason
    }
    if ($engine.Name -eq 'None') {
        return New-EobIntegrationStatus -Name 'PDF-Erzeugung' -State NotInstalled -Detail "$($engine.Reason) Berichte werden als HTML erzeugt." `
            -Hint 'Microsoft Edge ist unter Windows üblicherweise vorhanden; alternativ [Report] wkhtmltopdfPath setzen.'
    }
    return New-EobIntegrationStatus -Name 'PDF-Erzeugung' -State Available -Detail "$($engine.Name): $($engine.Path)"
}

#endregion
