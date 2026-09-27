#Requires -Version 7.2
<#
    easyONB.Core
    Basisfunktionen: Version, Pfade, Konvertierungen, strukturiertes Logging mit Audit und
    Redaktion sensibler Werte sowie die Plan-Engine (Schritte, Abhängigkeiten, WhatIf).
#>

Set-StrictMode -Version 3.0

# Zeichen, mit denen Tabellenkalkulationen eine Formel beginnen (CSV-Injektion)
$script:CsvFormulaPrefixes = [char[]]@('=', '+', '-', '@', "`t", "`r")

#region Modulzustand

$script:AppRoot = [System.IO.Path]::GetFullPath((Join-Path -Path $PSScriptRoot -ChildPath (Join-Path -Path '..' -ChildPath '..')))
$script:VersionCache = $null

$script:LevelOrder = @{
    Debug       = 0
    Information = 1
    Warning     = 2
    Error       = 3
    Critical    = 4
    Audit       = 5
}

$script:LogState = @{
    Initialized     = $false
    Directory       = $null
    AuditDirectory  = $null
    MinimumLevel    = 'Information'
    RetentionDays   = 90
    AuditRetentionDays = 0
    MaxFileSizeBytes = 20MB
    ConsoleOutput   = $false
    Queue           = $null
    FailureCount    = 0
    LastError       = $null
    LastFailureAt   = $null
    LastWarningAt   = [datetime]::MinValue
    FallbackActive  = $false
    FallbackReason  = $null
    CurrentFiles    = @{}
}

$script:RedactionValues = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::Ordinal)
$script:RedactionMask = '***'
$script:SensitiveKeyPattern = '(?i)(password|passwort|kennwort|pwd|secret|token|apikey|api_key|credential|clientsecret)'
$script:TextRedactionRules = @(
    @{
        Pattern     = '(?i)\b([A-Za-z0-9_\-\.]*(?:password|passwort|kennwort|pwd|secret|token|apikey|api_key|credential))("?\s*[:=]\s*)("[^"]*"|''[^'']*''|[^\s;,}\]]+)'
        Replacement = '$1$2***'
    }
    @{
        Pattern     = '(?i)\bBearer\s+[A-Za-z0-9\-\._~\+/]+=*'
        Replacement = 'Bearer ***'
    }
    @{
        Pattern     = '\beyJ[A-Za-z0-9_\-]{5,}\.[A-Za-z0-9_\-]{5,}\.[A-Za-z0-9_\-]{5,}'
        Replacement = '***'
    }
)

$script:AllowedHandlerModulePattern = 'easyONB.*'
$script:CurrentIdentityCache = $null

#endregion

#region Version und Pfade

function Get-EobAppRoot {
    <#
    .SYNOPSIS
        Liefert das Wurzelverzeichnis der Anwendung.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param()

    return $script:AppRoot
}

function Get-EobVersion {
    <#
    .SYNOPSIS
        Liefert die Anwendungsversion aus der Datei VERSION (einzige Versionsquelle).
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param()

    if ($null -eq $script:VersionCache) {
        $versionFile = Join-Path -Path $script:AppRoot -ChildPath 'VERSION'
        if (-not (Test-Path -LiteralPath $versionFile -PathType Leaf)) {
            throw "Versionsdatei nicht gefunden: $versionFile"
        }
        $raw = (Get-Content -LiteralPath $versionFile -Raw -ErrorAction Stop).Trim()
        if ($raw -notmatch '^\d+\.\d+\.\d+(-[0-9A-Za-z\.\-]+)?$') {
            throw "Ungültiges Versionsformat in ${versionFile}: '$raw'"
        }
        $script:VersionCache = $raw
    }
    return $script:VersionCache
}

function Test-EobWindowsStylePath {
    [CmdletBinding()]
    [OutputType([bool])]
    param([string]$Path)

    return ($Path -match '^[A-Za-z]:[\\/]' -or $Path -match '^\\\\[^\\]+\\')
}

function Resolve-EobPath {
    <#
    .SYNOPSIS
        Löst einen Pfad auf: Umgebungsvariablen, relative Pfade (bezogen auf BasePath) und Normalisierung.
    .DESCRIPTION
        Unterstützt %VAR%, $env:VAR und ${env:VAR}. Relative Pfade werden gegen BasePath
        (Standard: Anwendungswurzel) aufgelöst. Leere Werte liefern einen Leerstring.
    .PARAMETER Path
        Aufzulösender Pfad.
    .PARAMETER BasePath
        Basis für relative Pfade.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory)]
        [AllowEmptyString()]
        [AllowNull()]
        [string]$Path,

        [string]$BasePath = $script:AppRoot
    )

    if ([string]::IsNullOrWhiteSpace($Path)) {
        return ''
    }

    $value = $Path.Trim().Trim('"')

    $value = [regex]::Replace($value, '\$\{env:([A-Za-z0-9_\(\)]+)\}|\$env:([A-Za-z0-9_]+)', {
            param($match)
            $name = if ($match.Groups[1].Success) { $match.Groups[1].Value } else { $match.Groups[2].Value }
            $envValue = [Environment]::GetEnvironmentVariable($name)
            if ($null -eq $envValue) { return $match.Value }
            return $envValue
        })
    $value = [Environment]::ExpandEnvironmentVariables($value)

    if (Test-EobWindowsStylePath -Path $value) {
        if ($IsWindows) {
            return [System.IO.Path]::GetFullPath($value)
        }
        return $value
    }

    if ([System.IO.Path]::IsPathRooted($value)) {
        return [System.IO.Path]::GetFullPath($value)
    }

    if ([string]::IsNullOrWhiteSpace($BasePath)) {
        $BasePath = $script:AppRoot
    }
    $normalized = $value -replace '[\\/]', [System.IO.Path]::DirectorySeparatorChar
    return [System.IO.Path]::GetFullPath((Join-Path -Path $BasePath -ChildPath $normalized))
}

function Test-EobPathWithin {
    <#
    .SYNOPSIS
        Prüft, ob ein Pfad innerhalb eines Wurzelverzeichnisses liegt (Schutz vor Path-Traversal).
    .PARAMETER Path
        Zu prüfender Pfad (absolut oder relativ zu Root).
    .PARAMETER Root
        Erlaubtes Wurzelverzeichnis.
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$Root
    )

    $normalize = {
        param([string]$p)
        $n = $p -replace '[\\/]+', '/'
        if ($n.Length -gt 1) { $n = $n.TrimEnd('/') }
        return $n
    }

    $canonicalWindows = {
        param([string]$p)
        $segments = [System.Collections.Generic.List[string]]::new()
        foreach ($part in ($p -split '[\\/]+')) {
            if ($part -eq '..') {
                if ($segments.Count -gt 0) { $segments.RemoveAt($segments.Count - 1) }
            }
            elseif ($part -ne '.' -and $part -ne '') {
                $segments.Add($part)
            }
        }
        $prefix = if ($p -match '^[\\/]{2}') { '//' } else { '' }
        return $prefix + ($segments -join '/')
    }

    if ((Test-EobWindowsStylePath -Path $Path) -or (Test-EobWindowsStylePath -Path $Root)) {
        $candidate = if (Test-EobWindowsStylePath -Path $Path) { $Path } else { $Root.TrimEnd('\', '/') + '\' + $Path }
        $fullPath = & $canonicalWindows $candidate
        $fullRoot = & $canonicalWindows $Root
    }
    else {
        $fullRoot = & $normalize ([System.IO.Path]::GetFullPath($Root))
        # [System.IO.Path]::Combine statt Join-Path: konfigurierte Laufwerke müssen nicht existieren.
        $candidate = if ([System.IO.Path]::IsPathRooted($Path)) { $Path } else { [System.IO.Path]::Combine($Root, $Path) }
        $fullPath = & $normalize ([System.IO.Path]::GetFullPath($candidate))
    }

    $comparison = [System.StringComparison]::OrdinalIgnoreCase
    if ($fullPath.Equals($fullRoot, $comparison)) {
        return $true
    }
    return $fullPath.StartsWith($fullRoot + '/', $comparison)
}

function Get-EobSafeFileName {
    <#
    .SYNOPSIS
        Erzeugt einen sicheren Dateinamen (ohne Pfadtrenner, reservierte Namen oder Steuerzeichen).
    .PARAMETER Name
        Ausgangswert.
    .PARAMETER MaxLength
        Maximale Länge.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory)][AllowEmptyString()][string]$Name,
        [ValidateRange(8, 200)][int]$MaxLength = 80
    )

    $safe = [regex]::Replace($Name, '[<>:"/\\|?*\x00-\x1F]', '_')
    $safe = $safe.Trim().Trim('.', ' ')
    if ([string]::IsNullOrWhiteSpace($safe)) {
        $safe = 'unbenannt'
    }
    if ($safe -match '^(?i)(CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])(\..*)?$') {
        $safe = '_' + $safe
    }
    if ($safe.Length -gt $MaxLength) {
        $safe = $safe.Substring(0, $MaxLength)
    }
    return $safe
}

function New-EobDirectory {
    <#
    .SYNOPSIS
        Legt ein Verzeichnis an (falls nötig) und liefert den vollständigen Pfad.
    #>
    [CmdletBinding(SupportsShouldProcess)]
    [OutputType([string])]
    param(
        [Parameter(Mandatory)][string]$Path
    )

    if (-not (Test-Path -LiteralPath $Path -PathType Container)) {
        if ($PSCmdlet.ShouldProcess($Path, 'Verzeichnis anlegen')) {
            $null = New-Item -ItemType Directory -Path $Path -Force -ErrorAction Stop
        }
    }
    return $Path
}

#endregion

#region Konvertierung und Hilfsfunktionen

function ConvertTo-EobHtmlEncoded {
    <#
    .SYNOPSIS
        HTML-kodiert einen Wert für Berichte, Vorlagen und E-Mails (null ergibt einen Leerstring).
    .DESCRIPTION
        Kodiert die in Text und Attributwerten relevanten Zeichen & < > " '. Umlaute bleiben
        lesbar erhalten (Dokumente werden als UTF-8 geschrieben).
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param([Parameter(ValueFromPipeline)][AllowNull()][object]$Value)

    process {
        if ($null -eq $Value) { return '' }
        return ([string]$Value).Replace('&', '&amp;').Replace('<', '&lt;').Replace('>', '&gt;').Replace('"', '&quot;').Replace("'", '&#39;')
    }
}

function ConvertTo-EobCsvSafeValue {
    <#
    .SYNOPSIS
        Schützt Texte in CSV-Exporten vor Formel-Injektion (Beginn mit =, +, -, @, Tab oder CR).
    .DESCRIPTION
        Betroffenen Texten wird ein Apostroph vorangestellt, damit Tabellenkalkulationen sie
        nicht als Formel auswerten. Zahlen und andere Werttypen bleiben unverändert.
    #>
    [CmdletBinding()]
    [OutputType([object])]
    param([AllowNull()][object]$Value)

    if ($Value -isnot [string]) { return $Value }
    if ($Value.Length -gt 0 -and $Value.IndexOfAny($script:CsvFormulaPrefixes) -eq 0) {
        return "'" + $Value
    }
    return $Value
}

function ConvertTo-EobCsvSafeObject {
    <#
    .SYNOPSIS
        Wendet ConvertTo-EobCsvSafeValue auf alle Eigenschaften eines Objekts an.
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param([Parameter(Mandatory, ValueFromPipeline)][AllowNull()][object]$InputObject)

    process {
        if ($null -eq $InputObject) { return }
        $copy = [ordered]@{}
        foreach ($property in $InputObject.PSObject.Properties) {
            $copy[$property.Name] = ConvertTo-EobCsvSafeValue -Value $property.Value
        }
        [pscustomobject]$copy
    }
}

function ConvertTo-EobBoolean {
    <#
    .SYNOPSIS
        Wandelt Konfigurations- und CSV-Werte (1/0, true/false, ja/nein, yes/no) in Boolean um.
    .PARAMETER Value
        Eingabewert.
    .PARAMETER Default
        Rückgabewert für leere oder unbekannte Werte.
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param(
        [AllowNull()][object]$Value,
        [bool]$Default = $false
    )

    if ($null -eq $Value) { return $Default }
    if ($Value -is [bool]) { return $Value }
    $text = ([string]$Value).Trim().ToLowerInvariant()
    if ($text -in @('1', 'true', 'yes', 'y', 'ja', 'j', 'on', 'wahr', 'x')) { return $true }
    if ($text -in @('0', 'false', 'no', 'n', 'nein', 'off', 'falsch')) { return $false }
    return $Default
}

function Test-EobBooleanText {
    <#
    .SYNOPSIS
        Prüft, ob ein Text als Boolean interpretierbar ist.
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param([AllowNull()][AllowEmptyString()][string]$Value)

    if ([string]::IsNullOrWhiteSpace($Value)) { return $true }
    return $Value.Trim().ToLowerInvariant() -in @('1', 'true', 'yes', 'y', 'ja', 'j', 'on', 'wahr', 'x', '0', 'false', 'no', 'n', 'nein', 'off', 'falsch')
}

function ConvertTo-EobDate {
    <#
    .SYNOPSIS
        Wandelt Datumswerte (ISO 8601, dd.MM.yyyy, d.M.yyyy) kulturunabhängig um.
    .OUTPUTS
        [datetime] oder $null bei ungültigen/leeren Werten.
    #>
    [CmdletBinding()]
    param(
        [AllowNull()][object]$Value
    )

    if ($null -eq $Value) { return $null }
    if ($Value -is [datetime]) { return $Value }
    $text = ([string]$Value).Trim()
    if ($text -eq '') { return $null }

    $formats = [string[]]@(
        'yyyy-MM-dd', 'yyyy-MM-ddTHH:mm:ss', 'yyyy-MM-ddTHH:mm:ssK', 'yyyy-MM-ddTHH:mm:ss.fffffffK',
        'yyyy-MM-dd HH:mm', 'yyyy-MM-dd HH:mm:ss',
        'dd.MM.yyyy', 'd.M.yyyy', 'dd.MM.yyyy HH:mm', 'd.M.yyyy HH:mm', 'dd.MM.yy'
    )
    $parsed = [datetime]::MinValue
    $styles = [System.Globalization.DateTimeStyles]::AllowWhiteSpaces
    if ([datetime]::TryParseExact($text, $formats, [System.Globalization.CultureInfo]::InvariantCulture, $styles, [ref]$parsed)) {
        return $parsed
    }
    return $null
}

function Test-EobDistinguishedName {
    <#
    .SYNOPSIS
        Prüft die Syntax eines Distinguished Name (RFC 4514, mindestens eine DC-Komponente).
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param([AllowNull()][AllowEmptyString()][string]$Value)

    if ([string]::IsNullOrWhiteSpace($Value)) { return $false }
    $rdn = '(?:[A-Za-z][A-Za-z0-9-]*|[0-9]+(?:\.[0-9]+)*)=(?:[^,\\+"<>;]|\\[,\\+"<>;=#\s]|\\[0-9A-Fa-f]{2})+'
    if ($Value -notmatch "^\s*$rdn(?:\s*,\s*$rdn)*\s*$") { return $false }
    return ($Value -match '(?i)(^|,)\s*DC=')
}

function Test-EobEmailAddress {
    <#
    .SYNOPSIS
        Prüft das Format einer E-Mail-Adresse (pragmatisch, ohne DNS-Abfrage).
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param([AllowNull()][AllowEmptyString()][string]$Value)

    if ([string]::IsNullOrWhiteSpace($Value) -or $Value.Length -gt 254) { return $false }
    if ($Value -notmatch '^(?<local>[A-Za-z0-9.!#$%&''*+/=?^_`{|}~-]+)@(?<domain>[^@]+)$') { return $false }
    $local = $Matches['local']
    $domain = $Matches['domain']
    if ($local.Length -gt 64 -or $local.StartsWith('.') -or $local.EndsWith('.') -or $local.Contains('..')) { return $false }
    return ((Test-EobDomainName -Value $domain) -and $domain.Contains('.'))
}

function Test-EobDomainName {
    <#
    .SYNOPSIS
        Prüft die Syntax eines DNS-Domänennamens.
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param([AllowNull()][AllowEmptyString()][string]$Value)

    if ([string]::IsNullOrWhiteSpace($Value)) { return $false }
    return ($Value -match '^(?=.{1,253}$)(?!-)[A-Za-z0-9-]{1,63}(?<!-)(\.(?!-)[A-Za-z0-9-]{1,63}(?<!-))*$')
}

function Get-EobPropertyValue {
    <#
    .SYNOPSIS
        Liest eine Eigenschaft bzw. einen Schlüssel sicher (auch unter StrictMode).
    .PARAMETER InputObject
        Objekt, Hashtable oder Dictionary.
    .PARAMETER Name
        Eigenschafts- bzw. Schlüsselname.
    .PARAMETER Default
        Rückgabewert, falls nicht vorhanden oder $null.
    #>
    [CmdletBinding()]
    param(
        [AllowNull()][object]$InputObject,
        [Parameter(Mandatory)][string]$Name,
        [AllowNull()][object]$Default = $null
    )

    if ($null -eq $InputObject) { return $Default }
    if ($InputObject -is [System.Collections.IDictionary]) {
        if ($InputObject.Contains($Name) -and $null -ne $InputObject[$Name]) {
            return $InputObject[$Name]
        }
        return $Default
    }
    $property = $InputObject.PSObject.Properties[$Name]
    if ($null -ne $property -and $null -ne $property.Value) {
        return $property.Value
    }
    return $Default
}

function Get-EobCurrentIdentity {
    <#
    .SYNOPSIS
        Liefert Name und SID des ausführenden Benutzers.
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param()

    if ($IsWindows) {
        $identity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
        try {
            return [pscustomobject]@{
                Name = $identity.Name
                Sid  = $identity.User.Value
            }
        }
        finally {
            $identity.Dispose()
        }
    }
    $userName = if ($env:USER) { $env:USER } elseif ($env:USERNAME) { $env:USERNAME } else { 'unbekannt' }
    return [pscustomobject]@{
        Name = $userName
        Sid  = $null
    }
}

function New-EobFinding {
    <#
    .SYNOPSIS
        Erzeugt ein einheitliches Befundobjekt (Validierung, Konfiguration, Planung).
    #>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Erzeugt nur ein Objekt im Speicher; keine Systemänderung.')]
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory)][ValidateSet('Error', 'Warning', 'Information')][string]$Severity,
        [Parameter(Mandatory)][string]$Code,
        [Parameter(Mandatory)][string]$Message,
        [string]$Field,
        [string]$Source
    )

    [pscustomobject]@{
        PSTypeName = 'Eob.Finding'
        Severity   = $Severity
        Code       = $Code
        Message    = $Message
        Field      = $Field
        Source     = $Source
    }
}

function New-EobResult {
    <#
    .SYNOPSIS
        Standard-Rückgabeobjekt für Schritt-Handler.
    .PARAMETER Status
        Succeeded, Skipped oder Warning.
    #>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Erzeugt nur ein Objekt im Speicher; keine Systemänderung.')]
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [ValidateSet('Succeeded', 'Skipped', 'Warning')][string]$Status = 'Succeeded',
        [string]$Message = '',
        [hashtable]$Data = @{}
    )

    [pscustomobject]@{
        PSTypeName = 'Eob.StepResult'
        Status     = $Status
        Message    = $Message
        Data       = $Data
    }
}

#endregion

#region Redaktion

function Add-EobRedactionValue {
    <#
    .SYNOPSIS
        Registriert einen konkreten sensiblen Wert (z. B. generiertes Kennwort) zur Maskierung in Logs.
    .PARAMETER Value
        Klartextwert (mindestens 4 Zeichen). SecureStrings werden intern konvertiert.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][object]$Value
    )

    $plain = $null
    if ($Value -is [System.Security.SecureString]) {
        $plain = ConvertFrom-EobSecureStringInternal -SecureString $Value
    }
    else {
        $plain = [string]$Value
    }
    if (-not [string]::IsNullOrEmpty($plain) -and $plain.Length -ge 4) {
        $null = $script:RedactionValues.Add($plain)
    }
}

function Remove-EobRedactionValue {
    <#
    .SYNOPSIS
        Entfernt registrierte Redaktionswerte (einzeln oder alle).
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [object]$Value,
        [switch]$All
    )

    if ($All) {
        if ($PSCmdlet.ShouldProcess('Redaktionsliste', 'Alle Werte entfernen')) {
            $script:RedactionValues.Clear()
        }
        return
    }
    if ($null -ne $Value) {
        $plain = if ($Value -is [System.Security.SecureString]) { ConvertFrom-EobSecureStringInternal -SecureString $Value } else { [string]$Value }
        if ($PSCmdlet.ShouldProcess('Redaktionsliste', 'Wert entfernen')) {
            $null = $script:RedactionValues.Remove($plain)
        }
    }
}

function ConvertFrom-EobSecureStringInternal {
    [CmdletBinding()]
    [OutputType([string])]
    param([Parameter(Mandatory)][System.Security.SecureString]$SecureString)

    $pointer = [System.Runtime.InteropServices.Marshal]::SecureStringToGlobalAllocUnicode($SecureString)
    try {
        return [System.Runtime.InteropServices.Marshal]::PtrToStringUni($pointer)
    }
    finally {
        [System.Runtime.InteropServices.Marshal]::ZeroFreeGlobalAllocUnicode($pointer)
    }
}

function Protect-EobSensitiveText {
    <#
    .SYNOPSIS
        Maskiert Kennwörter, Tokens und registrierte sensible Werte in einem Text.
    .PARAMETER Text
        Eingabetext.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [AllowNull()][AllowEmptyString()][string]$Text
    )

    if ([string]::IsNullOrEmpty($Text)) { return $Text }
    $result = $Text
    if ($script:RedactionValues.Count -gt 0) {
        foreach ($secret in ($script:RedactionValues | Sort-Object -Property Length -Descending)) {
            $result = $result.Replace($secret, $script:RedactionMask)
        }
    }
    foreach ($rule in $script:TextRedactionRules) {
        $result = [regex]::Replace($result, $rule.Pattern, $rule.Replacement)
    }
    return $result
}

function Protect-EobSensitiveData {
    <#
    .SYNOPSIS
        Erzeugt eine redigierte Kopie von Hashtables, Dictionaries, Objekten und Listen.
    .PARAMETER InputObject
        Zu redigierende Daten.
    .PARAMETER Depth
        Maximale Rekursionstiefe.
    #>
    [CmdletBinding()]
    param(
        [AllowNull()][object]$InputObject,
        [ValidateRange(1, 10)][int]$Depth = 5
    )

    if ($null -eq $InputObject) { return $null }
    if ($InputObject -is [System.Security.SecureString] -or $InputObject -is [pscredential]) {
        return $script:RedactionMask
    }
    if ($InputObject -is [string]) {
        return Protect-EobSensitiveText -Text $InputObject
    }
    if ($InputObject -is [ValueType]) {
        return $InputObject
    }
    if ($Depth -le 1) {
        return Protect-EobSensitiveText -Text ([string]$InputObject)
    }
    if ($InputObject -is [System.Collections.IDictionary]) {
        $copy = [ordered]@{}
        foreach ($key in $InputObject.Keys) {
            if ([string]$key -match $script:SensitiveKeyPattern) {
                $copy[[string]$key] = $script:RedactionMask
            }
            else {
                $copy[[string]$key] = Protect-EobSensitiveData -InputObject $InputObject[$key] -Depth ($Depth - 1)
            }
        }
        return $copy
    }
    if ($InputObject -is [System.Collections.IEnumerable]) {
        $list = [System.Collections.Generic.List[object]]::new()
        foreach ($item in $InputObject) {
            $list.Add((Protect-EobSensitiveData -InputObject $item -Depth ($Depth - 1)))
        }
        return , $list.ToArray()
    }
    if ($InputObject -is [psobject]) {
        $copy = [ordered]@{}
        foreach ($property in $InputObject.PSObject.Properties) {
            if ($property.Name -match $script:SensitiveKeyPattern) {
                $copy[$property.Name] = $script:RedactionMask
            }
            else {
                $copy[$property.Name] = Protect-EobSensitiveData -InputObject $property.Value -Depth ($Depth - 1)
            }
        }
        return [pscustomobject]$copy
    }
    return Protect-EobSensitiveText -Text ([string]$InputObject)
}

#endregion

#region Logging

function Get-EobLogMutexName {
    [CmdletBinding()]
    [OutputType([string])]
    param([Parameter(Mandatory)][string]$Path)

    $sha = [System.Security.Cryptography.SHA256]::Create()
    try {
        $bytes = $sha.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($Path.ToLowerInvariant()))
    }
    finally {
        $sha.Dispose()
    }
    return 'easyONB_log_' + ([System.BitConverter]::ToString($bytes, 0, 12) -replace '-', '')
}

function Write-EobFileLine {
    <#
        Schreibt eine Zeile unter prozessübergreifender Sperre. Wirft bei Fehlern.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$Line
    )

    $mutex = [System.Threading.Mutex]::new($false, (Get-EobLogMutexName -Path $Path))
    $acquired = $false
    try {
        try {
            $acquired = $mutex.WaitOne(5000)
        }
        catch [System.Threading.AbandonedMutexException] {
            $acquired = $true
        }
        if (-not $acquired) {
            throw "Zeitüberschreitung beim Warten auf die Log-Sperre ($Path)."
        }
        [System.IO.File]::AppendAllText($Path, $Line + [Environment]::NewLine, [System.Text.UTF8Encoding]::new($false))
    }
    finally {
        if ($acquired) {
            $mutex.ReleaseMutex()
        }
        $mutex.Dispose()
    }
}

function Test-EobDirectoryWritable {
    [CmdletBinding()]
    [OutputType([bool])]
    param([Parameter(Mandatory)][string]$Path)

    try {
        if (-not (Test-Path -LiteralPath $Path -PathType Container)) {
            $null = New-Item -ItemType Directory -Path $Path -Force -ErrorAction Stop
        }
        $probe = [System.IO.Path]::Combine($Path, '.write-test-' + [guid]::NewGuid().ToString('N'))
        [System.IO.File]::WriteAllText($probe, '')
        Remove-Item -LiteralPath $probe -Force -ErrorAction Stop
        return $true
    }
    catch {
        return $false
    }
}

function Initialize-EobLogging {
    <#
    .SYNOPSIS
        Initialisiert Text- und Audit-Logging.
    .DESCRIPTION
        Legt Verzeichnisse an, prüft die Schreibbarkeit, aktiviert bei Bedarf ein
        Ausweichverzeichnis (sichtbar über Get-EobLogStatus) und bereinigt alte Logdateien.
    .PARAMETER Directory
        Verzeichnis für Textlogs.
    .PARAMETER AuditDirectory
        Verzeichnis für Audit-Logs (Standard: <Directory>/audit).
    .PARAMETER MinimumLevel
        Minimales Level für Textlogs (Audit wird immer geschrieben).
    .PARAMETER RetentionDays
        Aufbewahrung der Textlogs in Tagen.
    .PARAMETER AuditRetentionDays
        Aufbewahrung der Audit-Logs in Tagen (0 = unbegrenzt).
    .PARAMETER MaxFileSizeMB
        Maximale Größe einer Textlogdatei vor dem Wechsel auf eine Folgedatei.
    .PARAMETER EnableQueue
        Stellt Einträge zusätzlich in eine Warteschlange (Anzeige in der GUI).
    .PARAMETER FallbackRoot
        Basis des Ausweichverzeichnisses (Standard: lokales Anwendungsdatenverzeichnis des Benutzers).
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory)][string]$Directory,
        [string]$AuditDirectory,
        [ValidateSet('Debug', 'Information', 'Warning', 'Error', 'Critical')][string]$MinimumLevel = 'Information',
        [ValidateRange(1, 3650)][int]$RetentionDays = 90,
        [ValidateRange(0, 36500)][int]$AuditRetentionDays = 0,
        [ValidateRange(1, 1024)][int]$MaxFileSizeMB = 20,
        [switch]$EnableQueue,
        [switch]$ConsoleOutput,
        [string]$FallbackRoot
    )

    $script:LogState.FallbackActive = $false
    $script:LogState.FallbackReason = $null
    $script:LogState.FailureCount = 0
    $script:LogState.LastError = $null
    $script:LogState.CurrentFiles = @{}

    $logDir = $Directory
    # [System.IO.Path]::Combine statt Join-Path: ein konfiguriertes Laufwerk, das es auf diesem Rechner
    # nicht gibt, soll zum Ausweichverzeichnis führen und nicht zu einem Fehler.
    $auditDir = if ([string]::IsNullOrWhiteSpace($AuditDirectory)) { [System.IO.Path]::Combine($Directory, 'audit') } else { $AuditDirectory }

    if (-not (Test-EobDirectoryWritable -Path $logDir) -or -not (Test-EobDirectoryWritable -Path $auditDir)) {
        $base = if ($FallbackRoot) { $FallbackRoot } else { [Environment]::GetFolderPath([Environment+SpecialFolder]::LocalApplicationData) }
        if ([string]::IsNullOrWhiteSpace($base)) { $base = [System.IO.Path]::GetTempPath() }
        $fallback = Join-Path -Path (Join-Path -Path $base -ChildPath 'easyONBOARDING') -ChildPath 'Logs'
        $script:LogState.FallbackActive = $true
        $script:LogState.FallbackReason = "Logverzeichnis '$logDir' bzw. '$auditDir' ist nicht beschreibbar. Ausweichverzeichnis: $fallback"
        $logDir = $fallback
        $auditDir = Join-Path -Path $fallback -ChildPath 'audit'
        if (-not (Test-EobDirectoryWritable -Path $logDir) -or -not (Test-EobDirectoryWritable -Path $auditDir)) {
            $script:LogState.Initialized = $false
            $script:LogState.LastError = 'Weder das konfigurierte Logverzeichnis noch das Ausweichverzeichnis sind beschreibbar.'
            Write-Warning $script:LogState.LastError
            return Get-EobLogStatus
        }
        Write-Warning $script:LogState.FallbackReason
    }

    $script:LogState.Directory = $logDir
    $script:LogState.AuditDirectory = $auditDir
    $script:LogState.MinimumLevel = $MinimumLevel
    $script:LogState.RetentionDays = $RetentionDays
    $script:LogState.AuditRetentionDays = $AuditRetentionDays
    $script:LogState.MaxFileSizeBytes = [long]$MaxFileSizeMB * 1MB
    $script:LogState.ConsoleOutput = [bool]$ConsoleOutput
    if ($EnableQueue) {
        if ($null -eq $script:LogState.Queue) {
            $script:LogState.Queue = [System.Collections.Concurrent.ConcurrentQueue[psobject]]::new()
        }
    }
    else {
        $script:LogState.Queue = $null
    }
    $script:LogState.Initialized = $true

    Invoke-EobLogMaintenance | Out-Null
    return Get-EobLogStatus
}

function Invoke-EobLogMaintenance {
    <#
    .SYNOPSIS
        Löscht Text- und Audit-Logs, die älter als die konfigurierte Aufbewahrung sind.
    .OUTPUTS
        Anzahl gelöschter Dateien.
    #>
    [CmdletBinding(SupportsShouldProcess)]
    [OutputType([int])]
    param()

    if (-not $script:LogState.Initialized) { return 0 }
    $removed = 0
    $now = Get-Date
    $targets = @(
        @{ Dir = $script:LogState.Directory; Filter = 'easyONB_*.log'; Days = $script:LogState.RetentionDays }
        @{ Dir = $script:LogState.AuditDirectory; Filter = 'easyONB_audit_*.jsonl'; Days = $script:LogState.AuditRetentionDays }
    )
    foreach ($target in $targets) {
        if ($target.Days -le 0 -or -not (Test-Path -LiteralPath $target.Dir)) { continue }
        $cutoff = $now.AddDays(-$target.Days)
        foreach ($file in Get-ChildItem -LiteralPath $target.Dir -Filter $target.Filter -File -ErrorAction SilentlyContinue) {
            if ($file.LastWriteTime -lt $cutoff -and $PSCmdlet.ShouldProcess($file.FullName, 'Alte Logdatei löschen')) {
                try {
                    Remove-Item -LiteralPath $file.FullName -Force -ErrorAction Stop
                    $removed++
                }
                catch {
                    Register-EobLogFailure -Message "Logbereinigung fehlgeschlagen für $($file.Name): $($_.Exception.Message)"
                }
            }
        }
    }
    return $removed
}

function Register-EobLogFailure {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Message)

    $script:LogState.FailureCount++
    $script:LogState.LastError = $Message
    $script:LogState.LastFailureAt = Get-Date
    if (((Get-Date) - $script:LogState.LastWarningAt).TotalSeconds -ge 60) {
        $script:LogState.LastWarningAt = Get-Date
        Write-Warning "Protokollierung gestört: $Message"
    }
}

function Get-EobTextLogPath {
    [CmdletBinding()]
    [OutputType([string])]
    param([datetime]$Timestamp)

    $dateKey = $Timestamp.ToString('yyyyMMdd')
    $current = $script:LogState.CurrentFiles[$dateKey]
    if ($null -eq $current) {
        $current = Join-Path -Path $script:LogState.Directory -ChildPath "easyONB_$dateKey.log"
    }
    $index = 1
    while ((Test-Path -LiteralPath $current -PathType Leaf) -and ((Get-Item -LiteralPath $current).Length -ge $script:LogState.MaxFileSizeBytes)) {
        $current = Join-Path -Path $script:LogState.Directory -ChildPath ('easyONB_{0}_{1}.log' -f $dateKey, $index)
        $index++
        if ($index -gt 999) { break }
    }
    $script:LogState.CurrentFiles[$dateKey] = $current
    return $current
}

function Write-EobLog {
    <#
    .SYNOPSIS
        Schreibt einen strukturierten Log- bzw. Audit-Eintrag.
    .DESCRIPTION
        Alle Texte und Daten werden vor dem Schreiben redigiert. Fehler beim Schreiben werden
        nicht geworfen, sondern gezählt und über Get-EobLogStatus bzw. eine Warnung gemeldet.
    .PARAMETER Message
        Meldungstext.
    .PARAMETER Level
        Debug, Information, Warning, Error, Critical oder Audit.
    .PARAMETER OperationId
        Vorgangs-ID (Correlation ID).
    .PARAMETER Action
        Aktionsname (z. B. CreateUser).
    .PARAMETER Target
        Zielobjekt (z. B. SamAccountName).
    .PARAMETER Result
        Ergebnis (z. B. Succeeded, Failed).
    .PARAMETER DurationMs
        Dauer in Millisekunden.
    .PARAMETER Data
        Zusätzliche Daten (werden redigiert).
    .PARAMETER ErrorRecord
        Fehlerobjekt.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, Position = 0)][AllowEmptyString()][string]$Message,
        [Parameter(Position = 1)][ValidateSet('Debug', 'Information', 'Warning', 'Error', 'Critical', 'Audit')][string]$Level = 'Information',
        [string]$OperationId,
        [string]$Action,
        [string]$Target,
        [string]$Result,
        [long]$DurationMs = -1,
        [System.Collections.IDictionary]$Data,
        [System.Management.Automation.ErrorRecord]$ErrorRecord
    )

    try {
        $timestamp = Get-Date
        $isAudit = $Level -eq 'Audit'
        $threshold = $script:LevelOrder[$script:LogState.MinimumLevel]
        if (-not $isAudit -and $script:LevelOrder[$Level] -lt $threshold) {
            return
        }

        if ($null -eq $script:CurrentIdentityCache) {
            $script:CurrentIdentityCache = Get-EobCurrentIdentity
        }
        $identity = $script:CurrentIdentityCache
        $entry = [ordered]@{
            Timestamp   = $timestamp.ToString('o')
            Level       = $Level
            OperationId = $OperationId
            Actor       = $identity.Name
            Computer    = [Environment]::MachineName
            Action      = $Action
            Target      = $Target
            Result      = $Result
            DurationMs  = if ($DurationMs -ge 0) { $DurationMs } else { $null }
            Message     = Protect-EobSensitiveText -Text $Message
            Error       = $null
            Data        = $null
        }
        if ($null -ne $ErrorRecord) {
            $entry.Error = Protect-EobSensitiveText -Text ('{0}: {1}' -f $ErrorRecord.Exception.GetType().Name, $ErrorRecord.Exception.Message)
        }
        if ($null -ne $Data -and $Data.Count -gt 0) {
            $entry.Data = Protect-EobSensitiveData -InputObject $Data
        }

        Write-Verbose ('[{0}] {1}' -f $Level, $entry.Message)

        if ($null -ne $script:LogState.Queue) {
            $script:LogState.Queue.Enqueue([pscustomobject]$entry)
        }

        if (-not $script:LogState.Initialized) {
            return
        }

        $parts = [System.Collections.Generic.List[string]]::new()
        $parts.Add($entry.Timestamp)
        $parts.Add("[$Level]")
        if ($OperationId) { $parts.Add("[op:$OperationId]") }
        $parts.Add("[actor:$($entry.Actor)]")
        if ($Action) { $parts.Add("Action=$Action") }
        if ($Target) { $parts.Add("Target=$Target") }
        if ($Result) { $parts.Add("Result=$Result") }
        if ($DurationMs -ge 0) { $parts.Add("DurationMs=$DurationMs") }
        $parts.Add($entry.Message)
        if ($entry.Error) { $parts.Add("| Error: $($entry.Error)") }
        if ($null -ne $entry.Data) { $parts.Add('| Data: ' + ($entry.Data | ConvertTo-Json -Compress -Depth 5)) }
        $line = ($parts -join ' ') -replace '[\r\n]+', ' '

        try {
            Write-EobFileLine -Path (Get-EobTextLogPath -Timestamp $timestamp) -Line $line
        }
        catch {
            Register-EobLogFailure -Message "Textlog: $($_.Exception.Message)"
        }

        if ($isAudit) {
            try {
                $auditPath = Join-Path -Path $script:LogState.AuditDirectory -ChildPath ('easyONB_audit_{0}.jsonl' -f $timestamp.ToString('yyyyMM'))
                Write-EobFileLine -Path $auditPath -Line ($entry | ConvertTo-Json -Compress -Depth 6)
            }
            catch {
                Register-EobLogFailure -Message "Audit-Log: $($_.Exception.Message)"
            }
        }

        if ($script:LogState.ConsoleOutput -and $script:LevelOrder[$Level] -ge $script:LevelOrder['Warning'] -and -not $isAudit) {
            Write-Warning $entry.Message
        }
    }
    catch {
        Register-EobLogFailure -Message "Interner Logfehler: $($_.Exception.Message)"
    }
}

function Get-EobLogStatus {
    <#
    .SYNOPSIS
        Liefert den Zustand des Loggings (Pfade, Fehlerzähler, Ausweichbetrieb).
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param()

    [pscustomobject]@{
        PSTypeName     = 'Eob.LogStatus'
        Initialized    = $script:LogState.Initialized
        Directory      = $script:LogState.Directory
        AuditDirectory = $script:LogState.AuditDirectory
        MinimumLevel   = $script:LogState.MinimumLevel
        Healthy        = ($script:LogState.Initialized -and $script:LogState.FailureCount -eq 0)
        FailureCount   = $script:LogState.FailureCount
        LastError      = $script:LogState.LastError
        FallbackActive = $script:LogState.FallbackActive
        FallbackReason = $script:LogState.FallbackReason
    }
}

function Test-EobAuditWritable {
    <#
    .SYNOPSIS
        Prüft, ob das Audit-Log aktuell beschreibbar ist (Voraussetzung für Live-Ausführungen).
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param()

    if (-not $script:LogState.Initialized) { return $false }
    try {
        $auditPath = Join-Path -Path $script:LogState.AuditDirectory -ChildPath ('easyONB_audit_{0}.jsonl' -f (Get-Date).ToString('yyyyMM'))
        $stream = [System.IO.File]::Open($auditPath, [System.IO.FileMode]::Append, [System.IO.FileAccess]::Write, [System.IO.FileShare]::ReadWrite)
        $stream.Dispose()
        return $true
    }
    catch {
        Register-EobLogFailure -Message "Audit-Log nicht beschreibbar: $($_.Exception.Message)"
        return $false
    }
}

function Receive-EobLogEntry {
    <#
    .SYNOPSIS
        Entnimmt Einträge aus der Log-Warteschlange (für die Anzeige in der GUI).
    .PARAMETER MaxCount
        Maximale Anzahl.
    #>
    [CmdletBinding()]
    param([ValidateRange(1, 10000)][int]$MaxCount = 500)

    if ($null -eq $script:LogState.Queue) { return }
    $items = [System.Collections.Generic.List[psobject]]::new()
    $item = $null
    while ($items.Count -lt $MaxCount -and $script:LogState.Queue.TryDequeue([ref]$item)) {
        $items.Add($item)
    }
    return $items.ToArray()
}

function Get-EobAuditEntry {
    <#
    .SYNOPSIS
        Liest Audit-Einträge (JSON Lines) mit optionalen Filtern.
    .PARAMETER From
        Beginn des Zeitraums.
    .PARAMETER To
        Ende des Zeitraums.
    .PARAMETER OperationId
        Filter auf Vorgangs-ID.
    .PARAMETER Directory
        Audit-Verzeichnis (Standard: aktuell konfiguriert).
    #>
    [CmdletBinding()]
    param(
        [datetime]$From = [datetime]::MinValue,
        [datetime]$To = [datetime]::MaxValue,
        [string]$OperationId,
        [string]$Directory = $script:LogState.AuditDirectory
    )

    if ([string]::IsNullOrWhiteSpace($Directory) -or -not (Test-Path -LiteralPath $Directory)) {
        return
    }
    foreach ($file in Get-ChildItem -LiteralPath $Directory -Filter 'easyONB_audit_*.jsonl' -File | Sort-Object -Property Name) {
        foreach ($line in [System.IO.File]::ReadLines($file.FullName)) {
            if ([string]::IsNullOrWhiteSpace($line)) { continue }
            try {
                $entry = $line | ConvertFrom-Json -ErrorAction Stop
            }
            catch {
                continue
            }
            $ts = ConvertTo-EobDate -Value $entry.Timestamp
            if ($null -eq $ts) {
                $parsed = [datetime]::MinValue
                if ([datetime]::TryParse([string]$entry.Timestamp, [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::RoundtripKind, [ref]$parsed)) {
                    $ts = $parsed
                }
            }
            if ($null -ne $ts -and ($ts -lt $From -or $ts -gt $To)) { continue }
            if ($OperationId -and $entry.OperationId -ne $OperationId) { continue }
            $entry
        }
    }
}

#endregion

#region Plan-Engine

function New-EobOperationContext {
    <#
    .SYNOPSIS
        Erzeugt einen Vorgangskontext mit eindeutiger Operation-ID.
    #>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Erzeugt nur ein Objekt im Speicher; keine Systemänderung.')]
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory)][string]$Kind,
        [switch]$Simulation
    )

    $identity = Get-EobCurrentIdentity
    [pscustomobject]@{
        PSTypeName  = 'Eob.OperationContext'
        OperationId = [guid]::NewGuid().ToString()
        Kind        = $Kind
        Actor       = $identity.Name
        ActorSid    = $identity.Sid
        Computer    = [Environment]::MachineName
        StartedAt   = Get-Date
        Simulation  = [bool]$Simulation
        Tool        = "easyONBOARDING $(Get-EobVersion)"
    }
}

function New-EobPlan {
    <#
    .SYNOPSIS
        Erzeugt einen leeren Ausführungsplan.
    .PARAMETER Kind
        Onboarding, Offboarding, UserUpdate, PasswordReset, BulkOnboarding, BulkOffboarding.
    .PARAMETER Context
        Vorgangskontext (New-EobOperationContext).
    .PARAMETER Subject
        Betroffenes Objekt (Anzeige- und Reportdaten).
    .PARAMETER Config
        Konfiguration; Schrittparameter mit dem Wert '@Config' erhalten sie bei der Ausführung.
    #>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Erzeugt nur ein Objekt im Speicher; keine Systemänderung.')]
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory)][string]$Kind,
        [Parameter(Mandatory)][pscustomobject]$Context,
        [hashtable]$Subject = @{},
        [AllowNull()][pscustomobject]$Config,
        [switch]$Simulation
    )

    [pscustomobject]@{
        PSTypeName      = 'Eob.Plan'
        Kind            = $Kind
        OperationId     = $Context.OperationId
        CreatedAt       = Get-Date
        CreatedBy       = $Context.Actor
        Context         = $Context
        Config          = $Config
        Simulation      = ([bool]$Simulation -or [bool]$Context.Simulation)
        Subject         = $Subject
        Summary         = [ordered]@{}
        Steps           = [System.Collections.Generic.List[object]]::new()
        Findings        = [System.Collections.Generic.List[object]]::new()
        Runtime         = @{}
        Secrets         = @{}
        Status          = 'Planned'
        StartedAt       = $null
        CompletedAt     = $null
        CancelRequested = $false
        Aborted         = $false
    }
}

function Add-EobPlanFinding {
    <#
    .SYNOPSIS
        Fügt einem Plan einen Befund hinzu (Error blockiert die Ausführung).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][pscustomobject]$Plan,
        [Parameter(Mandatory)][ValidateSet('Error', 'Warning', 'Information')][string]$Severity,
        [Parameter(Mandatory)][string]$Code,
        [Parameter(Mandatory)][string]$Message,
        [string]$Field
    )

    $finding = New-EobFinding -Severity $Severity -Code $Code -Message $Message -Field $Field -Source $Plan.Kind
    $Plan.Findings.Add($finding)
}

function Add-EobPlanStep {
    <#
    .SYNOPSIS
        Fügt einem Plan einen Schritt hinzu.
    .PARAMETER Action
        Technischer Aktionsname (z. B. CreateUser).
    .PARAMETER Title
        Anzeigename des Schritts.
    .PARAMETER Handler
        Name der ausführenden Funktion (muss aus einem easyONB-Modul stammen).
    .PARAMETER Parameters
        Parameter für den Handler; Secrets nur als '@Secret:<Name>'-Verweis.
    .PARAMETER Risk
        Low, Medium, High oder Destructive.
    .PARAMETER Critical
        Bei Fehler werden alle folgenden Schritte übersprungen.
    .PARAMETER DependsOn
        IDs von Schritten, die erfolgreich sein müssen.
    .PARAMETER Phase
        Phase (Offboarding: Immediate, ExitDate, Retention, FinalDeletion; sonst Main).
    .PARAMETER Details
        Vorschauzeilen (z. B. Änderungen alt → neu).
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory)][pscustomobject]$Plan,
        [Parameter(Mandatory)][string]$Action,
        [Parameter(Mandatory)][string]$Title,
        [Parameter(Mandatory)][string]$Handler,
        [hashtable]$Parameters = @{},
        [string]$Target = '',
        [ValidateSet('Low', 'Medium', 'High', 'Destructive')][string]$Risk = 'Low',
        [switch]$Critical,
        [string[]]$DependsOn = @(),
        [string]$Phase = 'Main',
        [string[]]$Details = @(),
        [bool]$Enabled = $true,
        [datetime]$DueDate = [datetime]::MinValue
    )

    $id = 'S{0:D2}' -f ($Plan.Steps.Count + 1)
    $step = [pscustomobject]@{
        PSTypeName = 'Eob.PlanStep'
        Id         = $id
        Order      = $Plan.Steps.Count + 1
        Phase      = $Phase
        DueDate    = if ($DueDate -eq [datetime]::MinValue) { $null } else { $DueDate }
        Action     = $Action
        Title      = $Title
        Target     = $Target
        Handler    = $Handler
        Parameters = $Parameters
        Risk       = $Risk
        Critical   = [bool]$Critical
        DependsOn  = @($DependsOn)
        Enabled    = $Enabled
        Details    = @($Details)
        Status     = 'Planned'
        Message    = ''
        StartedAt  = $null
        DurationMs = $null
        Error      = $null
    }
    $Plan.Steps.Add($step)
    return $step
}

function Test-EobPlanExecutable {
    <#
    .SYNOPSIS
        Prüft, ob ein Plan ausgeführt werden darf (keine Error-Befunde, mindestens ein Schritt).
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param([Parameter(Mandatory)][pscustomobject]$Plan)

    $errors = @($Plan.Findings | Where-Object { $_.Severity -eq 'Error' })
    return ($errors.Count -eq 0 -and $Plan.Steps.Count -gt 0)
}

function Resolve-EobStepParameter {
    [CmdletBinding()]
    [OutputType([hashtable])]
    param(
        [Parameter(Mandatory)][pscustomobject]$Plan,
        [Parameter(Mandatory)][pscustomobject]$Step
    )

    $resolved = @{}
    foreach ($key in $Step.Parameters.Keys) {
        $value = $Step.Parameters[$key]
        if ($value -is [string] -and $value.StartsWith('@Secret:')) {
            $secretName = $value.Substring(8)
            if (-not $Plan.Secrets.ContainsKey($secretName) -or $null -eq $Plan.Secrets[$secretName]) {
                throw "Geheimer Wert '$secretName' ist für Schritt $($Step.Id) nicht verfügbar."
            }
            $resolved[$key] = $Plan.Secrets[$secretName]
        }
        elseif ($value -is [string] -and $value -ceq '@Config') {
            $resolved[$key] = $Plan.Config
        }
        elseif ($value -is [string] -and $value.StartsWith('@Runtime:')) {
            $runtimeName = $value.Substring(9)
            if ($Plan.Runtime.ContainsKey($runtimeName)) {
                $resolved[$key] = $Plan.Runtime[$runtimeName]
            }
            else {
                throw "Laufzeitwert '$runtimeName' ist für Schritt $($Step.Id) nicht verfügbar."
            }
        }
        else {
            $resolved[$key] = $value
        }
    }
    return $resolved
}

function Get-EobAllowedHandler {
    <#
        Liefert den Handler-Befehl, sofern er aus einem easyONB-Modul stammt.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Name)

    $command = Get-Command -Name $Name -CommandType Function -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($null -eq $command) {
        throw "Handler '$Name' wurde nicht gefunden."
    }
    if ([string]::IsNullOrEmpty($command.ModuleName) -or $command.ModuleName -notlike $script:AllowedHandlerModulePattern) {
        throw "Handler '$Name' ist nicht zulässig (nur Funktionen aus easyONB-Modulen)."
    }
    return $command
}

function Invoke-EobPlanStep {
    <#
    .SYNOPSIS
        Führt genau einen Planschritt aus (Abhängigkeiten, Simulation, Fehlerbehandlung, Audit).
    .DESCRIPTION
        Wird von Invoke-EobPlan und von der GUI (kooperative Ausführung) verwendet.
        In der Simulation erhalten Handler mit ShouldProcess-Unterstützung -WhatIf.
    .PARAMETER Plan
        Plan.
    .PARAMETER Step
        Auszuführender Schritt.
    .PARAMETER Simulation
        Erzwingt Simulation.
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory)][pscustomobject]$Plan,
        [Parameter(Mandatory)][pscustomobject]$Step,
        [switch]$Simulation
    )

    $simulate = [bool]$Simulation -or [bool]$Plan.Simulation

    if (-not $Step.Enabled) {
        $Step.Status = 'Skipped'
        $Step.Message = 'Schritt deaktiviert.'
        return $Step
    }
    if ($Plan.CancelRequested) {
        $Step.Status = 'Skipped'
        $Step.Message = 'Ausführung abgebrochen.'
        return $Step
    }
    if ($Plan.Aborted) {
        $Step.Status = 'Skipped'
        $Step.Message = 'Übersprungen nach kritischem Fehler.'
        return $Step
    }
    foreach ($dependencyId in $Step.DependsOn) {
        $dependency = $Plan.Steps | Where-Object { $_.Id -eq $dependencyId } | Select-Object -First 1
        if ($null -eq $dependency -or $dependency.Status -notin @('Succeeded', 'Simulated', 'Warning')) {
            $Step.Status = 'Skipped'
            $Step.Message = "Voraussetzung $dependencyId nicht erfüllt."
            return $Step
        }
    }

    $Step.Status = 'Running'
    $Step.StartedAt = Get-Date
    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    try {
        $command = Get-EobAllowedHandler -Name $Step.Handler
        $parameters = Resolve-EobStepParameter -Plan $Plan -Step $Step
        $supportsShouldProcess = $command.Parameters.ContainsKey('WhatIf')
        if ($supportsShouldProcess) {
            $parameters['WhatIf'] = $simulate
            $parameters['Confirm'] = $false
        }
        $parameters['ErrorAction'] = 'Stop'

        $output = & $command @parameters
        $result = @($output | Where-Object { $null -ne $_ -and $_.PSObject.TypeNames -contains 'Eob.StepResult' }) | Select-Object -Last 1

        if ($simulate -and $supportsShouldProcess) {
            $Step.Status = 'Simulated'
            $Step.Message = if ($null -ne $result -and $result.Message) { $result.Message } else { 'Simulation: keine Änderung vorgenommen.' }
        }
        elseif ($null -ne $result) {
            $Step.Status = $result.Status
            $Step.Message = $result.Message
        }
        else {
            $Step.Status = 'Succeeded'
            $Step.Message = 'Erfolgreich.'
        }
        if ($null -ne $result -and $null -ne $result.Data) {
            foreach ($key in $result.Data.Keys) {
                $Plan.Runtime[$key] = $result.Data[$key]
            }
        }
    }
    catch {
        $Step.Status = 'Failed'
        $Step.Error = Protect-EobSensitiveText -Text $_.Exception.Message
        $Step.Message = $Step.Error
        if ($Step.Critical) {
            $Plan.Aborted = $true
        }
    }
    finally {
        $stopwatch.Stop()
        $Step.DurationMs = $stopwatch.ElapsedMilliseconds
    }

    $logLevel = if ($simulate) { 'Information' } else { 'Audit' }
    Write-EobLog -Level $logLevel -OperationId $Plan.OperationId -Action $Step.Action -Target $Step.Target `
        -Result $Step.Status -DurationMs $Step.DurationMs -Message ('{0}: {1}' -f $Step.Title, $Step.Message)
    if ($Step.Status -eq 'Failed') {
        Write-EobLog -Level 'Error' -OperationId $Plan.OperationId -Action $Step.Action -Target $Step.Target -Result 'Failed' -Message $Step.Message
    }
    return $Step
}

function Get-EobPlanOutcome {
    <#
    .SYNOPSIS
        Ermittelt den Gesamtstatus eines Plans aus den Schrittstatus.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory)][pscustomobject]$Plan,
        [string[]]$IncludePhase
    )

    $steps = @($Plan.Steps | Where-Object { -not $IncludePhase -or $_.Phase -in $IncludePhase })
    if ($Plan.CancelRequested) { return 'Cancelled' }
    if (@($steps | Where-Object Status -EQ 'Failed').Count -gt 0) { return 'Failed' }
    if ($Plan.Simulation) { return 'Simulated' }
    if (@($steps | Where-Object { $_.Status -in @('Warning', 'Skipped') -and $_.Enabled }).Count -gt 0) { return 'CompletedWithWarnings' }
    return 'Succeeded'
}

function Start-EobPlanRun {
    <#
    .SYNOPSIS
        Bereitet die Ausführung eines Plans vor (Prüfungen, Status, Protokoll).
    .DESCRIPTION
        Wird von Invoke-EobPlan und von der kooperativen Ausführung der Oberfläche verwendet.
        Live-Ausführungen setzen ein beschreibbares Audit-Log voraus.
    #>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Setzt nur den Planstatus und protokolliert den Start; bestätigt wird in Invoke-EobPlan bzw. in der Oberfläche.')]
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][pscustomobject]$Plan,
        [switch]$Simulation
    )

    if (-not (Test-EobPlanExecutable -Plan $Plan)) {
        $messages = @($Plan.Findings | Where-Object Severity -EQ 'Error' | ForEach-Object Message)
        throw ('Der Plan kann nicht ausgeführt werden: ' + ($messages -join ' | '))
    }
    if ($Simulation) { $Plan.Simulation = $true }
    if (-not $Plan.Simulation -and -not (Test-EobAuditWritable)) {
        throw 'Das Audit-Log ist nicht beschreibbar. Live-Ausführungen sind aus Gründen der Nachvollziehbarkeit gesperrt.'
    }
    $Plan.StartedAt = Get-Date
    $Plan.Status = 'Running'
    $level = if ($Plan.Simulation) { 'Information' } else { 'Audit' }
    Write-EobLog -Level $level -OperationId $Plan.OperationId -Action ('{0}Started' -f $Plan.Kind) -Target (Get-EobPlanSubjectText -Plan $Plan) `
        -Message ('Plan gestartet ({0} Schritte, Simulation: {1})' -f $Plan.Steps.Count, $Plan.Simulation)
}

function Complete-EobPlanRun {
    <#
    .SYNOPSIS
        Schließt die Ausführung eines Plans ab (Gesamtstatus, Protokoll).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][pscustomobject]$Plan,
        [string[]]$IncludePhase
    )

    $Plan.CompletedAt = Get-Date
    $Plan.Status = Get-EobPlanOutcome -Plan $Plan -IncludePhase $IncludePhase
    $level = if ($Plan.Simulation) { 'Information' } else { 'Audit' }
    $duration = if ($null -ne $Plan.StartedAt) { [long]($Plan.CompletedAt - $Plan.StartedAt).TotalMilliseconds } else { 0 }
    Write-EobLog -Level $level -OperationId $Plan.OperationId -Action ('{0}Completed' -f $Plan.Kind) -Target (Get-EobPlanSubjectText -Plan $Plan) `
        -Result $Plan.Status -DurationMs $duration -Message 'Plan abgeschlossen.'
}

function Get-EobPlanSubjectText {
    [CmdletBinding()]
    [OutputType([string])]
    param([Parameter(Mandatory)][pscustomobject]$Plan)

    return [string](Get-EobPropertyValue -InputObject $Plan.Subject -Name 'SamAccountName' -Default $Plan.Kind)
}

function Invoke-EobPlan {
    <#
    .SYNOPSIS
        Führt einen Plan vollständig aus (oder simuliert ihn mit -WhatIf).
    .DESCRIPTION
        Bestätigung erfolgt einmalig auf Planebene (ConfirmImpact High). Live-Ausführungen setzen
        ein beschreibbares Audit-Log voraus. Mit -WhatIf bzw. Plan.Simulation werden alle
        schreibenden Handler mit -WhatIf aufgerufen.
    .PARAMETER Plan
        Auszuführender Plan.
    .PARAMETER IncludePhase
        Nur Schritte dieser Phasen ausführen (übrige bleiben 'Planned').
    #>
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory)][pscustomobject]$Plan,
        [string[]]$IncludePhase
    )

    if (-not (Test-EobPlanExecutable -Plan $Plan)) {
        $messages = @($Plan.Findings | Where-Object Severity -EQ 'Error' | ForEach-Object Message)
        throw ('Der Plan kann nicht ausgeführt werden: ' + ($messages -join ' | '))
    }
    if ($WhatIfPreference) {
        $Plan.Simulation = $true
    }
    if (-not $Plan.Simulation) {
        if (-not $PSCmdlet.ShouldProcess((Get-EobPlanSubjectText -Plan $Plan), ('{0}: {1} Schritt(e) ausführen' -f $Plan.Kind, $Plan.Steps.Count))) {
            $Plan.Status = 'Cancelled'
            return $Plan
        }
    }

    Start-EobPlanRun -Plan $Plan
    foreach ($step in $Plan.Steps) {
        if ($IncludePhase -and $step.Phase -notin $IncludePhase) {
            continue
        }
        $null = Invoke-EobPlanStep -Plan $Plan -Step $step -Simulation:$Plan.Simulation
    }
    Complete-EobPlanRun -Plan $Plan -IncludePhase $IncludePhase
    return $Plan
}

function Stop-EobPlan {
    <#
    .SYNOPSIS
        Fordert den Abbruch eines laufenden Plans an (wirksam vor dem nächsten Schritt).
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param([Parameter(Mandatory)][pscustomobject]$Plan)

    if ($PSCmdlet.ShouldProcess($Plan.OperationId, 'Abbruch anfordern')) {
        $Plan.CancelRequested = $true
    }
}

function Get-EobPlanStatistic {
    <#
    .SYNOPSIS
        Zählt Planschritte nach Status.
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param([Parameter(Mandatory)][pscustomobject]$Plan)

    $steps = @($Plan.Steps)
    [pscustomobject]@{
        Total     = $steps.Count
        Succeeded = @($steps | Where-Object Status -EQ 'Succeeded').Count
        Simulated = @($steps | Where-Object Status -EQ 'Simulated').Count
        Warning   = @($steps | Where-Object Status -EQ 'Warning').Count
        Skipped   = @($steps | Where-Object Status -EQ 'Skipped').Count
        Failed    = @($steps | Where-Object Status -EQ 'Failed').Count
        Planned   = @($steps | Where-Object Status -EQ 'Planned').Count
    }
}

#endregion

#region Integrationsstatus

function New-EobIntegrationStatus {
    <#
    .SYNOPSIS
        Erzeugt ein einheitliches Statusobjekt für eine Integration (Dashboard).
    .PARAMETER State
        Available, Connected, NotConnected, NotConfigured, NotInstalled, Error, Disabled.
    #>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Erzeugt nur ein Objekt im Speicher; keine Systemänderung.')]
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][ValidateSet('Available', 'Connected', 'NotConnected', 'NotConfigured', 'NotInstalled', 'Error', 'Disabled')][string]$State,
        [string]$Detail = '',
        [string]$Version = '',
        [string]$Hint = '',
        [ValidateSet('Productive', 'Prepared', 'Simulated', 'Experimental')][string]$Implementation = 'Productive'
    )

    [pscustomobject]@{
        PSTypeName     = 'Eob.IntegrationStatus'
        Name           = $Name
        State          = $State
        IsUsable       = $State -in @('Available', 'Connected')
        Detail         = $Detail
        Version        = $Version
        Hint           = $Hint
        Implementation = $Implementation
    }
}

function Get-EobIntegrationStatus {
    <#
    .SYNOPSIS
        Sammelt den Status aller Integrationen über die Statusfunktionen der geladenen Module.
    .DESCRIPTION
        Fehlende Module führen nicht zum Abbruch, sondern zu einem Statuseintrag "NotInstalled".
    .PARAMETER Config
        Konfigurationsobjekt (Import-EobConfiguration).
    #>
    [CmdletBinding()]
    param([AllowNull()][object]$Config)

    $providers = [ordered]@{
        'Active Directory'      = 'Get-EobAdStatus'
        'Exchange'              = 'Get-EobExchangeStatus'
        'Microsoft Graph'       = 'Get-EobGraphStatus'
        'Entra Connect Sync'    = 'Get-EobEntraConnectStatus'
        'SMTP'                  = 'Get-EobSmtpStatus'
        'Dateiserver'           = 'Get-EobFileServerStatus'
        'PDF-Erzeugung'         = 'Get-EobPdfEngineStatus'
    }
    foreach ($name in $providers.Keys) {
        $functionName = $providers[$name]
        $command = Get-Command -Name $functionName -ErrorAction SilentlyContinue
        if ($null -eq $command) {
            New-EobIntegrationStatus -Name $name -State 'NotInstalled' -Detail "Statusfunktion $functionName nicht geladen."
            continue
        }
        try {
            & $command -Config $Config
        }
        catch {
            New-EobIntegrationStatus -Name $name -State 'Error' -Detail (Protect-EobSensitiveText -Text $_.Exception.Message)
        }
    }
}

#endregion
