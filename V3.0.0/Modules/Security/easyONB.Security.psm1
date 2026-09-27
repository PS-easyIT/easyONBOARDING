#Requires -Version 7.2
<#
    easyONB.Security
    Kennwortgenerierung (CSPRNG) und -prüfung, Erkennung privilegierter Gruppen und geschützter
    Konten, LDAP-Escaping und Eingabevalidierung. Enthält keine AD-Abfragen; AD-Daten werden von
    den aufrufenden Modulen übergeben.
#>

Set-StrictMode -Version 3.0

#region Konstanten

$script:UpperChars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
$script:LowerChars = 'abcdefghijklmnopqrstuvwxyz'
$script:DigitChars = '0123456789'
$script:AmbiguousChars = 'Il1O0o'
$script:DefaultSpecialChars = '!#%*+-=?@_'

# Domänenrelative RIDs privilegierter bzw. sicherheitskritischer Gruppen
$script:PrivilegedDomainRids = @{
    498 = 'Enterprise Read-only Domain Controllers'
    512 = 'Domain Admins'
    516 = 'Domain Controllers'
    517 = 'Cert Publishers'
    518 = 'Schema Admins'
    519 = 'Enterprise Admins'
    520 = 'Group Policy Creator Owners'
    521 = 'Read-only Domain Controllers'
    526 = 'Key Admins'
    527 = 'Enterprise Key Admins'
}

# Builtin-Gruppen (S-1-5-32-x)
$script:PrivilegedBuiltinSids = @{
    'S-1-5-32-544' = 'Administrators'
    'S-1-5-32-548' = 'Account Operators'
    'S-1-5-32-549' = 'Server Operators'
    'S-1-5-32-550' = 'Print Operators'
    'S-1-5-32-551' = 'Backup Operators'
    'S-1-5-32-552' = 'Replicator'
    'S-1-5-32-557' = 'Incoming Forest Trust Builders'
    'S-1-5-32-569' = 'Cryptographic Operators'
    'S-1-5-32-578' = 'Hyper-V Administrators'
    'S-1-5-32-582' = 'Storage Replica Administrators'
}

# Namens-Fallback (englische und deutsche Standardnamen sowie Exchange-Rollengruppen)
$script:PrivilegedGroupNames = @(
    'Domain Admins', 'Domänen-Admins', 'Enterprise Admins', 'Organisations-Admins', 'Schema Admins', 'Schema-Admins',
    'Administrators', 'Administratoren', 'Account Operators', 'Konten-Operatoren', 'Server Operators', 'Server-Operatoren',
    'Backup Operators', 'Sicherungs-Operatoren', 'Print Operators', 'Druck-Operatoren', 'Replicator', 'Replikations-Operator',
    'Group Policy Creator Owners', 'Richtlinien-Ersteller-Besitzer', 'Key Admins', 'Schlüsseladministratoren',
    'Enterprise Key Admins', 'Unternehmensschlüsseladministratoren', 'Domain Controllers', 'Domänencontroller',
    'Read-only Domain Controllers', 'Schreibgeschützte Domänencontroller', 'Enterprise Read-only Domain Controllers',
    'Schreibgeschützte Domänencontroller der Organisation', 'Cert Publishers', 'Zertifikatherausgeber', 'DnsAdmins',
    'Cryptographic Operators', 'Kryptografie-Operatoren', 'Hyper-V Administrators', 'Hyper-V-Administratoren',
    'Organization Management', 'Organisationsverwaltung', 'Exchange Trusted Subsystem', 'Exchange Windows Permissions',
    'Exchange Organization Administrators'
)

# RIDs eingebauter Konten, die nie bearbeitet werden
$script:BuiltinAccountRids = @{
    500 = 'Eingebautes Administratorkonto'
    501 = 'Gastkonto'
    502 = 'krbtgt (Kerberos-Dienst)'
    503 = 'DefaultAccount'
    504 = 'WDAGUtilityAccount'
}

$script:CommonPasswords = @(
    'password', 'passwort', 'kennwort', '123456', '12345678', '123456789', 'qwertz', 'qwerty', 'willkommen',
    'welcome', 'sommer', 'winter', 'fruehling', 'herbst', 'letmein', 'admin', 'changeme', 'geheim'
)

#endregion

#region Kennwörter

function Get-EobRandomIndex {
    <#
        Kryptografisch sichere Zufallszahl im Bereich [0, MaxExclusive).
    #>
    [CmdletBinding()]
    [OutputType([int])]
    param([Parameter(Mandatory)][ValidateRange(1, [int]::MaxValue)][int]$MaxExclusive)

    return [System.Security.Cryptography.RandomNumberGenerator]::GetInt32(0, $MaxExclusive)
}

function Get-EobPasswordPolicy {
    <#
    .SYNOPSIS
        Ermittelt die wirksame Kennwortrichtlinie aus Konfiguration und optional der Domänenrichtlinie.
    .PARAMETER Config
        Konfigurationsobjekt (Import-EobConfiguration). Ohne Konfiguration gelten Schema-Standards.
    .PARAMETER DomainPolicy
        Optionales Objekt mit MinPasswordLength und ComplexityEnabled (z. B. aus Get-EobAdPasswordPolicy).
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [AllowNull()][pscustomobject]$Config,
        [AllowNull()][object]$DomainPolicy
    )

    $read = {
        param([string]$Key, [string]$As = 'Int')
        Get-EobConfigValue -Config $Config -Section 'PasswordFixGenerate' -Key $Key -As $As
    }
    $includeSpecial = & $read 'IncludeSpecialChars' 'Bool'
    $special = [string](& $read 'SpecialCharacters' 'String')
    if ([string]::IsNullOrEmpty($special)) { $special = $script:DefaultSpecialChars }

    $policy = [ordered]@{
        PSTypeName           = 'Eob.PasswordPolicy'
        Length               = [int](& $read 'DefaultPasswordLength')
        MinUpperCase         = [int](& $read 'MinUpperCase')
        MinLowerCase         = [int](& $read 'MinLowerCase')
        MinDigits            = [int](& $read 'MinDigits')
        MinSpecialChars      = if ($includeSpecial) { [int](& $read 'MinSpecialChars') } else { 0 }
        MinNonAlpha          = [int](& $read 'MinNonAlpha')
        IncludeSpecialChars  = [bool]$includeSpecial
        SpecialCharacters    = $special
        ExcludeAmbiguous     = [bool](& $read 'AvoidAmbiguousChars' 'Bool')
        MinManualLength      = [int](& $read 'MinManualPasswordLength')
        ComplexityEnabled    = $true
        DomainMinLength      = $null
        Source               = 'Konfiguration'
    }

    $useDomain = Get-EobConfigValue -Config $Config -Section 'Security' -Key 'UseDomainPasswordPolicy' -As Bool
    if ($useDomain -and $null -ne $DomainPolicy) {
        $domainMin = [int](Get-EobPropertyValue -InputObject $DomainPolicy -Name 'MinPasswordLength' -Default 0)
        $complexity = Get-EobPropertyValue -InputObject $DomainPolicy -Name 'ComplexityEnabled' -Default $true
        $policy.DomainMinLength = $domainMin
        $policy.ComplexityEnabled = [bool]$complexity
        if ($domainMin -gt $policy.Length) { $policy.Length = $domainMin }
        if ($domainMin -gt $policy.MinManualLength) { $policy.MinManualLength = $domainMin }
        $policy.Source = 'Konfiguration + Domänenrichtlinie'
    }
    return [pscustomobject]$policy
}

function New-EobPassword {
    <#
    .SYNOPSIS
        Erzeugt ein Kennwort mit einem kryptografisch sicheren Zufallsgenerator.
    .DESCRIPTION
        Standardmäßig wird ein SecureString geliefert. Der Klartext entsteht nur mit -AsPlainText
        (z. B. für die einmalige Anzeige) und wird nie protokolliert.
    .PARAMETER Policy
        Richtlinie (Get-EobPasswordPolicy). Ohne Angabe gelten die Schema-Standards.
    .PARAMETER Length
        Überschreibt die Länge der Richtlinie.
    .PARAMETER AsPlainText
        Liefert einen String statt eines SecureString.
    #>
    [CmdletBinding()]
    param(
        [AllowNull()][pscustomobject]$Policy,
        [ValidateRange(8, 256)][int]$Length,
        [switch]$AsPlainText
    )

    if ($null -eq $Policy) {
        $Policy = Get-EobPasswordPolicy -Config $null
    }
    $targetLength = if ($PSBoundParameters.ContainsKey('Length')) { $Length } else { [int]$Policy.Length }

    $filter = {
        param([string]$Pool)
        if (-not $Policy.ExcludeAmbiguous) { return $Pool }
        return -join ($Pool.ToCharArray() | Where-Object { $script:AmbiguousChars.IndexOf($_) -lt 0 })
    }
    $upper = & $filter $script:UpperChars
    $lower = & $filter $script:LowerChars
    $digits = & $filter $script:DigitChars
    $special = if ($Policy.IncludeSpecialChars) { & $filter ([string]$Policy.SpecialCharacters) } else { '' }

    $required = [int]$Policy.MinUpperCase + [int]$Policy.MinLowerCase + [int]$Policy.MinDigits + [int]$Policy.MinSpecialChars
    if ($required -gt $targetLength) {
        throw "Die Kennwortlänge $targetLength ist kleiner als die Summe der Mindestanzahlen ($required)."
    }
    if ($Policy.MinSpecialChars -gt 0 -and [string]::IsNullOrEmpty($special)) {
        throw 'Sonderzeichen sind erforderlich, aber es ist kein zulässiges Sonderzeichen konfiguriert.'
    }

    $chars = [System.Collections.Generic.List[char]]::new($targetLength)
    $pick = { param([string]$Pool) $Pool[(Get-EobRandomIndex -MaxExclusive $Pool.Length)] }

    for ($i = 0; $i -lt $Policy.MinUpperCase; $i++) { $chars.Add((& $pick $upper)) }
    for ($i = 0; $i -lt $Policy.MinLowerCase; $i++) { $chars.Add((& $pick $lower)) }
    for ($i = 0; $i -lt $Policy.MinDigits; $i++) { $chars.Add((& $pick $digits)) }
    for ($i = 0; $i -lt $Policy.MinSpecialChars; $i++) { $chars.Add((& $pick $special)) }

    $nonAlphaPool = $digits + $special
    $nonAlphaCount = [int]$Policy.MinDigits + [int]$Policy.MinSpecialChars
    while ($nonAlphaCount -lt $Policy.MinNonAlpha -and $chars.Count -lt $targetLength) {
        $chars.Add((& $pick $nonAlphaPool))
        $nonAlphaCount++
    }

    $all = $upper + $lower + $digits + $special
    while ($chars.Count -lt $targetLength) {
        $chars.Add((& $pick $all))
    }

    $array = $chars.ToArray()
    for ($i = $array.Length - 1; $i -gt 0; $i--) {
        $j = Get-EobRandomIndex -MaxExclusive ($i + 1)
        $tmp = $array[$i]
        $array[$i] = $array[$j]
        $array[$j] = $tmp
    }

    try {
        if ($AsPlainText) {
            return [string]::new($array)
        }
        $secure = [System.Security.SecureString]::new()
        foreach ($c in $array) { $secure.AppendChar($c) }
        $secure.MakeReadOnly()
        return $secure
    }
    finally {
        [array]::Clear($array, 0, $array.Length)
        $chars.Clear()
    }
}

function ConvertTo-EobPlainText {
    <#
    .SYNOPSIS
        Wandelt einen SecureString für die einmalige Anzeige in Klartext um.
    .DESCRIPTION
        Nur für die Anzeige gegenüber dem Administrator bzw. den Druck aus dem Speicher verwenden.
        Das Ergebnis darf nicht protokolliert oder gespeichert werden.
    #>
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

function Test-EobPasswordPolicy {
    <#
    .SYNOPSIS
        Prüft ein Kennwort gegen die Richtlinie, ohne es auszugeben.
    .PARAMETER Password
        SecureString (bevorzugt) oder String.
    .PARAMETER Policy
        Richtlinie (Get-EobPasswordPolicy).
    .PARAMETER SamAccountName
        Kontoname (darf nicht enthalten sein).
    .PARAMETER DisplayName
        Anzeigename bzw. Namensbestandteile (Tokens ab drei Zeichen dürfen nicht enthalten sein).
    .PARAMETER MinimumLength
        Überschreibt die Mindestlänge (Standard: MinManualLength der Richtlinie).
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory)][object]$Password,
        [AllowNull()][pscustomobject]$Policy,
        [string]$SamAccountName,
        [string[]]$DisplayName = @(),
        [int]$MinimumLength
    )

    if ($null -eq $Policy) {
        $Policy = Get-EobPasswordPolicy -Config $null
    }
    $plain = if ($Password -is [System.Security.SecureString]) { ConvertTo-EobPlainText -SecureString $Password } else { [string]$Password }
    $violations = [System.Collections.Generic.List[string]]::new()
    try {
        $minLength = if ($PSBoundParameters.ContainsKey('MinimumLength')) { $MinimumLength } else { [int]$Policy.MinManualLength }
        if ($plain.Length -lt $minLength) {
            $violations.Add("Das Kennwort muss mindestens $minLength Zeichen lang sein.")
        }
        $categories = 0
        if ($plain -cmatch '[A-Z]') { $categories++ }
        if ($plain -cmatch '[a-z]') { $categories++ }
        if ($plain -match '[0-9]') { $categories++ }
        if ($plain -match '[^A-Za-z0-9]') { $categories++ }
        if ($Policy.ComplexityEnabled -and $categories -lt 3) {
            $violations.Add('Das Kennwort muss Zeichen aus mindestens drei Kategorien enthalten (Groß-, Kleinbuchstaben, Ziffern, Sonderzeichen).')
        }
        if (-not [string]::IsNullOrWhiteSpace($SamAccountName) -and $SamAccountName.Length -ge 3 -and
            $plain.IndexOf($SamAccountName, [System.StringComparison]::OrdinalIgnoreCase) -ge 0) {
            $violations.Add('Das Kennwort darf den Kontonamen nicht enthalten.')
        }
        foreach ($part in $DisplayName) {
            foreach ($token in ([string]$part -split '[,\.\-_#\s\t]+')) {
                if ($token.Length -ge 3 -and $plain.IndexOf($token, [System.StringComparison]::OrdinalIgnoreCase) -ge 0) {
                    $violations.Add('Das Kennwort darf keine Namensbestandteile enthalten.')
                    break
                }
            }
        }
        $lowerPlain = $plain.ToLowerInvariant()
        foreach ($common in $script:CommonPasswords) {
            if ($lowerPlain -eq $common -or ($lowerPlain.Contains($common) -and $common.Length -ge 6 -and $lowerPlain.Length -le $common.Length + 4)) {
                $violations.Add('Das Kennwort ist zu einfach oder zu verbreitet.')
                break
            }
        }
    }
    finally {
        $plain = $null
    }
    $unique = @($violations | Select-Object -Unique)
    [pscustomobject]@{
        PSTypeName = 'Eob.PasswordValidation'
        IsValid    = ($unique.Count -eq 0)
        Violations = $unique
    }
}

function Get-EobPasswordPolicyText {
    <#
    .SYNOPSIS
        Beschreibt die Richtlinie in lesbarer Form (z. B. für das Willkommensdokument).
    #>
    [CmdletBinding()]
    [OutputType([string[]])]
    param([Parameter(Mandatory)][pscustomobject]$Policy)

    $lines = [System.Collections.Generic.List[string]]::new()
    $lines.Add("Mindestlänge: $($Policy.MinManualLength) Zeichen")
    if ($Policy.ComplexityEnabled) {
        $lines.Add('Zeichen aus mindestens drei Kategorien: Großbuchstaben, Kleinbuchstaben, Ziffern, Sonderzeichen')
    }
    $lines.Add('Keine Namensbestandteile und nicht der Benutzername')
    return $lines.ToArray()
}

#endregion

#region Privilegierte Gruppen und geschützte Konten

function ConvertTo-EobSidString {
    [CmdletBinding()]
    [OutputType([string])]
    param([AllowNull()][object]$Value)

    if ($null -eq $Value) { return '' }
    $text = if ($Value -is [System.Security.Principal.SecurityIdentifier]) { $Value.Value } else { [string]$Value }
    if ($text -match '^S-1-[0-9\-]+$') { return $text.ToUpperInvariant() }
    return ''
}

function Get-EobSidRid {
    [CmdletBinding()]
    [OutputType([int])]
    param([string]$Sid)

    if ($Sid -match '^S-1-5-21-\d+-\d+-\d+-(\d+)$') {
        return [int]$Matches[1]
    }
    return -1
}

function Get-EobNameFromIdentity {
    <#
        Liefert Name, SamAccountName, DN und SID aus Strings oder AD-Objekten.
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param([Parameter(Mandatory)][object]$InputObject)

    if ($InputObject -is [string]) {
        $text = $InputObject.Trim()
        $sid = ConvertTo-EobSidString -Value $text
        if ($sid) {
            return [pscustomobject]@{ Name = ''; SamAccountName = ''; DistinguishedName = ''; Sid = $sid; AdminCount = 0 }
        }
        if ($text -match '^(?i)CN=((?:[^,\\]|\\.)+),') {
            return [pscustomobject]@{ Name = ($Matches[1] -replace '\\(.)', '$1'); SamAccountName = ''; DistinguishedName = $text; Sid = ''; AdminCount = 0 }
        }
        return [pscustomobject]@{ Name = $text; SamAccountName = $text; DistinguishedName = ''; Sid = ''; AdminCount = 0 }
    }
    [pscustomobject]@{
        Name              = [string](Get-EobPropertyValue -InputObject $InputObject -Name 'Name' -Default '')
        SamAccountName    = [string](Get-EobPropertyValue -InputObject $InputObject -Name 'SamAccountName' -Default '')
        DistinguishedName = [string](Get-EobPropertyValue -InputObject $InputObject -Name 'DistinguishedName' -Default '')
        Sid               = ConvertTo-EobSidString -Value (Get-EobPropertyValue -InputObject $InputObject -Name 'SID' -Default (Get-EobPropertyValue -InputObject $InputObject -Name 'ObjectSid' -Default $null))
        AdminCount        = [int](Get-EobPropertyValue -InputObject $InputObject -Name 'AdminCount' -Default 0)
    }
}

function Test-EobPrivilegedGroup {
    <#
    .SYNOPSIS
        Prüft, ob eine Gruppe privilegiert bzw. geschützt ist.
    .DESCRIPTION
        Erkennung über bekannte SIDs/RIDs (sprachunabhängig), Standardnamen (DE/EN), adminCount,
        konfigurierte Gruppen ([Security] ProtectedGroups) und Namensmuster (ProtectedGroupPatterns).
        Verschachtelte Mitgliedschaften prüft das AD-Modul und übergibt sie über -NestedIn.
    .PARAMETER Group
        Gruppenname, DN, SID oder AD-Gruppenobjekt.
    .PARAMETER Config
        Konfiguration.
    .PARAMETER NestedIn
        Namen privilegierter Gruppen, in denen die Gruppe (transitiv) Mitglied ist.
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory)][object]$Group,
        [AllowNull()][pscustomobject]$Config,
        [string[]]$NestedIn = @()
    )

    $info = Get-EobNameFromIdentity -InputObject $Group
    $reasons = [System.Collections.Generic.List[string]]::new()
    $names = @($info.Name, $info.SamAccountName) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique

    if ($info.Sid) {
        if ($script:PrivilegedBuiltinSids.ContainsKey($info.Sid)) {
            $reasons.Add("Eingebaute privilegierte Gruppe ($($script:PrivilegedBuiltinSids[$info.Sid])).")
        }
        $rid = Get-EobSidRid -Sid $info.Sid
        if ($script:PrivilegedDomainRids.ContainsKey($rid)) {
            $reasons.Add("Privilegierte Domänengruppe ($($script:PrivilegedDomainRids[$rid])).")
        }
    }
    foreach ($name in $names) {
        if (@($script:PrivilegedGroupNames | Where-Object { $_ -ieq $name }).Count -gt 0) {
            $reasons.Add("Privilegierte Standardgruppe '$name'.")
        }
    }
    if ($info.AdminCount -eq 1) {
        $reasons.Add('Gruppe ist durch AdminSDHolder geschützt (adminCount=1).')
    }

    foreach ($entry in @(Get-EobConfigValue -Config $Config -Section 'Security' -Key 'ProtectedGroups' -As List)) {
        $entrySid = ConvertTo-EobSidString -Value $entry
        if (($entrySid -and $entrySid -eq $info.Sid) -or
            ($info.DistinguishedName -and $entry -ieq $info.DistinguishedName) -or
            (@($names | Where-Object { $_ -ieq $entry }).Count -gt 0)) {
            $reasons.Add("Als geschützt konfiguriert ([Security] ProtectedGroups: $entry).")
        }
    }
    foreach ($pattern in @(Get-EobConfigValue -Config $Config -Section 'Security' -Key 'ProtectedGroupPatterns' -As List)) {
        foreach ($name in $names) {
            if ($name -like $pattern) {
                $reasons.Add("Entspricht dem Schutzmuster '$pattern'.")
            }
        }
    }
    foreach ($parent in $NestedIn) {
        if (-not [string]::IsNullOrWhiteSpace($parent)) {
            $reasons.Add("Verschachtelt in privilegierter Gruppe '$parent'.")
        }
    }

    $distinct = @($reasons | Select-Object -Unique)
    [pscustomobject]@{
        PSTypeName  = 'Eob.GroupProtection'
        Group       = if ($info.Name) { $info.Name } elseif ($info.Sid) { $info.Sid } else { [string]$Group }
        IsProtected = ($distinct.Count -gt 0)
        Reasons     = $distinct
    }
}

function Test-EobProtectedAccount {
    <#
    .SYNOPSIS
        Prüft, ob ein Benutzerkonto bearbeitet werden darf.
    .DESCRIPTION
        Sperrt: eigenes Konto, eingebaute Konten (RID 500-504), Break-Glass-, geschützte und
        Dienstkonten, Konten in geschützten OUs, Nicht-Benutzerobjekte und kritische Systemobjekte.
        Erfordert gesonderte Bestätigung: adminCount=1 oder Mitgliedschaft in privilegierten Gruppen.
    .PARAMETER User
        AD-Benutzerobjekt bzw. Objekt mit SamAccountName, SID, DistinguishedName, AdminCount,
        ObjectClass und IsCriticalSystemObject.
    .PARAMETER Config
        Konfiguration.
    .PARAMETER CurrentUserSid
        SID des ausführenden Administrators.
    .PARAMETER PrivilegedGroups
        Namen privilegierter Gruppen, in denen das Konto (transitiv) Mitglied ist.
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory)][object]$User,
        [AllowNull()][pscustomobject]$Config,
        [string]$CurrentUserSid,
        [string[]]$PrivilegedGroups = @()
    )

    $info = Get-EobNameFromIdentity -InputObject $User
    $block = [System.Collections.Generic.List[string]]::new()
    $acknowledge = [System.Collections.Generic.List[string]]::new()
    $sam = $info.SamAccountName

    if ($CurrentUserSid -and $info.Sid -and $info.Sid -eq (ConvertTo-EobSidString -Value $CurrentUserSid)) {
        $block.Add('Das eigene, aktuell angemeldete Konto darf nicht bearbeitet werden.')
    }
    $rid = Get-EobSidRid -Sid $info.Sid
    if ($script:BuiltinAccountRids.ContainsKey($rid)) {
        $block.Add("Eingebautes Systemkonto: $($script:BuiltinAccountRids[$rid]).")
    }
    if ($sam -ieq 'krbtgt') {
        $block.Add('krbtgt ist ein Systemkonto.')
    }

    $matchesEntry = {
        param([string]$Entry)
        $entrySid = ConvertTo-EobSidString -Value $Entry
        return (($entrySid -and $entrySid -eq $info.Sid) -or ($sam -and $Entry -ieq $sam))
    }
    foreach ($entry in @(Get-EobConfigValue -Config $Config -Section 'Security' -Key 'BreakGlassAccounts' -As List)) {
        if (& $matchesEntry $entry) { $block.Add('Notfallkonto (Break-Glass) laut Konfiguration.') }
    }
    foreach ($entry in @(Get-EobConfigValue -Config $Config -Section 'Security' -Key 'ProtectedAccounts' -As List)) {
        if (& $matchesEntry $entry) { $block.Add('Geschütztes Konto laut Konfiguration.') }
    }
    foreach ($pattern in @(Get-EobConfigValue -Config $Config -Section 'Security' -Key 'ServiceAccountPatterns' -As List)) {
        if ($sam -and $sam -like $pattern) { $block.Add("Dienstkonto (Muster '$pattern').") }
    }
    if ($info.DistinguishedName) {
        foreach ($ou in @(Get-EobConfigValue -Config $Config -Section 'Security' -Key 'ProtectedOUs' -As List)) {
            $normalizedOu = ($ou -replace '\s*,\s*', ',').Trim()
            $normalizedDn = ($info.DistinguishedName -replace '\s*,\s*', ',').Trim()
            if ($normalizedDn.EndsWith(",$normalizedOu", [System.StringComparison]::OrdinalIgnoreCase)) {
                $block.Add("Konto liegt in einer geschützten OU ($ou).")
            }
        }
    }

    $objectClass = Get-EobPropertyValue -InputObject $User -Name 'ObjectClass' -Default 'user'
    $classes = @($objectClass) | ForEach-Object { [string]$_ }
    if (@($classes | Where-Object { $_ -in @('computer', 'msDS-GroupManagedServiceAccount', 'msDS-ManagedServiceAccount', 'msDS-DelegatedManagedServiceAccount', 'group') }).Count -gt 0) {
        $block.Add("Objektklasse '$($classes[-1])' wird nicht bearbeitet (nur Benutzerkonten).")
    }
    if ([bool](Get-EobPropertyValue -InputObject $User -Name 'IsCriticalSystemObject' -Default $false)) {
        $block.Add('Kritisches Systemobjekt (isCriticalSystemObject).')
    }

    if ($info.AdminCount -eq 1) {
        $acknowledge.Add('Konto ist bzw. war privilegiert (adminCount=1).')
    }
    foreach ($group in $PrivilegedGroups) {
        if (-not [string]::IsNullOrWhiteSpace($group)) {
            $acknowledge.Add("Mitglied der privilegierten Gruppe '$group'.")
        }
    }

    $blockReasons = @($block | Select-Object -Unique)
    $ackReasons = @($acknowledge | Select-Object -Unique)
    [pscustomobject]@{
        PSTypeName              = 'Eob.AccountProtection'
        SamAccountName          = $sam
        IsBlocked               = ($blockReasons.Count -gt 0)
        RequiresAcknowledgement = ($ackReasons.Count -gt 0)
        BlockReasons            = $blockReasons
        AcknowledgementReasons  = $ackReasons
    }
}

#endregion

#region Escaping und Eingabeformate

function ConvertTo-EobLdapFilterValue {
    <#
    .SYNOPSIS
        Maskiert einen Wert für LDAP-Filter nach RFC 4515 (Schutz vor LDAP-Injection).
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param([Parameter(Mandatory)][AllowEmptyString()][string]$Value)

    $builder = [System.Text.StringBuilder]::new()
    foreach ($c in $Value.ToCharArray()) {
        switch ([int]$c) {
            0x5C { [void]$builder.Append('\5c') }
            0x2A { [void]$builder.Append('\2a') }
            0x28 { [void]$builder.Append('\28') }
            0x29 { [void]$builder.Append('\29') }
            0x00 { [void]$builder.Append('\00') }
            default { [void]$builder.Append($c) }
        }
    }
    return $builder.ToString()
}

function ConvertTo-EobAdFilterLiteral {
    <#
    .SYNOPSIS
        Erzeugt ein Literal für die PowerShell-AD-Filtersyntax ('O''Brien').
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param([Parameter(Mandatory)][AllowEmptyString()][string]$Value)

    return "'" + $Value.Replace("'", "''") + "'"
}

function Test-EobSamAccountName {
    <#
    .SYNOPSIS
        Prüft einen SamAccountName (max. 20 Zeichen, keine unzulässigen Zeichen, nicht auf '.' endend).
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param([AllowNull()][AllowEmptyString()][string]$Value)

    if ([string]::IsNullOrWhiteSpace($Value) -or $Value.Length -gt 20) { return $false }
    if ($Value -match '["/\\\[\]:;|=,+*?<>@\s]' -or $Value -match '[\x00-\x1F]') { return $false }
    if ($Value.EndsWith('.') -or $Value -match '^\.+$') { return $false }
    return $true
}

function Test-EobUserPrincipalName {
    <#
    .SYNOPSIS
        Prüft einen UPN (lokaler Teil max. 64 Zeichen, zulässige Zeichen, gültige Domäne).
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param([AllowNull()][AllowEmptyString()][string]$Value)

    if ([string]::IsNullOrWhiteSpace($Value) -or $Value -notmatch '^(?<local>[^@]+)@(?<domain>[^@]+)$') { return $false }
    $local = $Matches['local']
    $domain = $Matches['domain']
    if ($local.Length -gt 64 -or $local -notmatch "^[A-Za-z0-9'._!#^~-]+$") { return $false }
    if ($local.StartsWith('.') -or $local.EndsWith('.') -or $local.Contains('..')) { return $false }
    return (Test-EobDomainName -Value $domain)
}

function Test-EobPersonName {
    <#
    .SYNOPSIS
        Prüft Vor- bzw. Nachnamen (Buchstaben inkl. Umlaute/Akzente, Leerzeichen, Bindestrich, Apostroph, Punkt).
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param(
        [AllowNull()][AllowEmptyString()][string]$Value,
        [int]$MinLength = 1,
        [int]$MaxLength = 64
    )

    if ([string]::IsNullOrWhiteSpace($Value)) { return $false }
    $trimmed = $Value.Trim()
    if ($trimmed.Length -lt $MinLength -or $trimmed.Length -gt $MaxLength) { return $false }
    return ($trimmed -match "^[\p{L}\p{M}][\p{L}\p{M} '’\-\.]*$")
}

function Test-EobPhoneNumber {
    <#
    .SYNOPSIS
        Prüft eine Telefonnummer (Ziffern, Leerzeichen, + ( ) / . -), optional mit eigenem Muster.
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param(
        [AllowNull()][AllowEmptyString()][string]$Value,
        [string]$Pattern
    )

    if ([string]::IsNullOrWhiteSpace($Value)) { return $false }
    if ($Value -notmatch '^\+?[0-9 ()/.\-]{3,40}$' -or ($Value -replace '\D', '').Length -lt 3) { return $false }
    if (-not [string]::IsNullOrWhiteSpace($Pattern) -and $Value -notmatch $Pattern) { return $false }
    return $true
}

function Test-EobSafeText {
    <#
    .SYNOPSIS
        Prüft Freitext auf Steuerzeichen und Maximallänge (z. B. Beschreibung, Abteilung).
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param(
        [AllowNull()][AllowEmptyString()][string]$Value,
        [int]$MaxLength = 1024
    )

    if ($null -eq $Value) { return $true }
    if ($Value.Length -gt $MaxLength) { return $false }
    return ($Value -notmatch '[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]')
}

#endregion
