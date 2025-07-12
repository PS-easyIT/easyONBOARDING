#region [AD Manager Module]
# Manages Active Directory operations, group assignments, and user lifecycle
# Version: 1.4.XX
# Author: IT Team

#region [Module Variables]
$global:ADManager = @{
    IsConnected = $false
    DomainController = ""
    Domain = ""
    AvailableOUs = @()
    AvailableGroups = @()
    LastError = ""
}
#endregion

#region [Connection Functions]

function Test-ADConnection {
    <#
    .SYNOPSIS
    Testet die Verbindung zu Active Directory
    
    .PARAMETER Server
    Domain Controller Server (optional)
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [string]$Server = "autodiscover"
    )
    
    try {
        Write-LogMessage "Teste AD-Verbindung..." "INFO"
        
        # Teste AD-Modul
        if (-not (Get-Module -Name ActiveDirectory -ErrorAction SilentlyContinue)) {
            Import-Module ActiveDirectory -ErrorAction Stop
        }
        
        # Teste Verbindung
        if ($Server -eq "autodiscover" -or [string]::IsNullOrWhiteSpace($Server)) {
            $domain = Get-ADDomain -ErrorAction Stop
            $domainController = $domain.PDCEmulator
        }
        else {
            $domain = Get-ADDomain -Server $Server -ErrorAction Stop
            $domainController = $Server
        }
        
        $global:ADManager.IsConnected = $true
        $global:ADManager.DomainController = $domainController
        $global:ADManager.Domain = $domain.DNSRoot
        $global:ADManager.LastError = ""
        
        Write-LogMessage "AD-Verbindung erfolgreich: $($domain.DNSRoot) via $domainController" "INFO"
        
        # Lade verfügbare OUs und Gruppen
        Get-AvailableOUs
        Get-AvailableGroups
        
        return $true
    }
    catch {
        $global:ADManager.IsConnected = $false
        $global:ADManager.LastError = $_.Exception.Message
        Write-LogMessage "AD-Verbindung fehlgeschlagen: $($_.Exception.Message)" "ERROR"
        return $false
    }
}

function Get-ADConnectionStatus {
    <#
    .SYNOPSIS
    Gibt den aktuellen AD-Verbindungsstatus zurück
    #>
    
    return @{
        IsConnected = $global:ADManager.IsConnected
        DomainController = $global:ADManager.DomainController
        Domain = $global:ADManager.Domain
        LastError = $global:ADManager.LastError
    }
}

#endregion

#region [OU Management Functions]

function Get-AvailableOUs {
    <#
    .SYNOPSIS
    Ermittelt alle verfügbaren Organisationseinheiten
    #>
    
    try {
        if (-not $global:ADManager.IsConnected) {
            throw "Keine AD-Verbindung verfügbar"
        }
        
        Write-LogMessage "Lade verfügbare OUs..." "INFO"
        
        $ous = Get-ADOrganizationalUnit -Filter * -Server $global:ADManager.DomainController | 
               Sort-Object DistinguishedName |
               Select-Object Name, DistinguishedName, @{
                   Name = "DisplayName"
                   Expression = {
                       $pathParts = $_.DistinguishedName -split "," | Where-Object { $_ -like "OU=*" }
                       $pathParts = $pathParts | ForEach-Object { $_ -replace "OU=" }
                       [array]::Reverse($pathParts)
                       $pathParts -join " > "
                   }
               }
        
        $global:ADManager.AvailableOUs = $ous
        
        Write-LogMessage "$($ous.Count) OUs geladen" "INFO"
        
        return $ous
    }
    catch {
        Write-LogMessage "Fehler beim Laden der OUs: $($_.Exception.Message)" "ERROR"
        return @()
    }
}

function Get-OUByName {
    <#
    .SYNOPSIS
    Sucht eine OU anhand des Namens
    
    .PARAMETER Name
    Name oder Teil des OU-Namens
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name
    )
    
    if (-not $global:ADManager.IsConnected) {
        throw "Keine AD-Verbindung verfügbar"
    }
    
    return $global:ADManager.AvailableOUs | Where-Object { 
        $_.Name -like "*$Name*" -or $_.DisplayName -like "*$Name*" 
    }
}

#endregion

#region [Group Management Functions]

function Get-AvailableGroups {
    <#
    .SYNOPSIS
    Ermittelt alle verfügbaren AD-Gruppen
    #>
    
    try {
        if (-not $global:ADManager.IsConnected) {
            throw "Keine AD-Verbindung verfügbar"
        }
        
        Write-LogMessage "Lade verfügbare AD-Gruppen..." "INFO"
        
        $groups = Get-ADGroup -Filter * -Server $global:ADManager.DomainController -Properties Description, GroupScope, GroupCategory |
                  Sort-Object Name |
                  Select-Object Name, Description, GroupScope, GroupCategory, DistinguishedName, @{
                      Name = "TypeDisplay"
                      Expression = {
                          if ($_.GroupCategory -eq "Security") {
                              "Security ($($_.GroupScope))"
                          }
                          else {
                              "Distribution ($($_.GroupScope))"
                          }
                      }
                  }
        
        $global:ADManager.AvailableGroups = $groups
        
        Write-LogMessage "$($groups.Count) AD-Gruppen geladen" "INFO"
        
        return $groups
    }
    catch {
        Write-LogMessage "Fehler beim Laden der AD-Gruppen: $($_.Exception.Message)" "ERROR"
        return @()
    }
}

function Get-GroupsByCategory {
    <#
    .SYNOPSIS
    Filtert Gruppen nach Kategorie und Typ
    
    .PARAMETER Category
    Security oder Distribution
    
    .PARAMETER Scope
    Global, DomainLocal oder Universal
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [ValidateSet("Security", "Distribution")]
        [string]$Category,
        
        [Parameter(Mandatory = $false)]
        [ValidateSet("Global", "DomainLocal", "Universal")]
        [string]$Scope
    )
    
    $filteredGroups = $global:ADManager.AvailableGroups
    
    if ($Category) {
        $filteredGroups = $filteredGroups | Where-Object { $_.GroupCategory -eq $Category }
    }
    
    if ($Scope) {
        $filteredGroups = $filteredGroups | Where-Object { $_.GroupScope -eq $Scope }
    }
    
    return $filteredGroups
}

function Test-GroupExists {
    <#
    .SYNOPSIS
    Prüft ob eine AD-Gruppe existiert
    
    .PARAMETER GroupName
    Name der zu prüfenden Gruppe
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$GroupName
    )
    
    try {
        if (-not $global:ADManager.IsConnected) {
            throw "Keine AD-Verbindung verfügbar"
        }
        
        $group = Get-ADGroup -Identity $GroupName -Server $global:ADManager.DomainController -ErrorAction SilentlyContinue
        return $null -ne $group
    }
    catch {
        return $false
    }
}

function Add-UserToGroups {
    <#
    .SYNOPSIS
    Fügt einen Benutzer zu mehreren AD-Gruppen hinzu
    
    .PARAMETER UserName
    SamAccountName des Benutzers
    
    .PARAMETER GroupNames
    Array von Gruppennamen
    
    .PARAMETER SkipValidation
    Überspringt die Validierung der Gruppennamen
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserName,
        
        [Parameter(Mandatory = $true)]
        [string[]]$GroupNames,
        
        [Parameter(Mandatory = $false)]
        [switch]$SkipValidation
    )
    
    try {
        if (-not $global:ADManager.IsConnected) {
            throw "Keine AD-Verbindung verfügbar"
        }
        
        Write-LogMessage "Füge Benutzer '$UserName' zu Gruppen hinzu: $($GroupNames -join ', ')" "INFO"
        
        $successCount = 0
        $errorCount = 0
        $results = @()
        
        foreach ($groupName in $GroupNames) {
            try {
                # Validiere Gruppe falls nicht übersprungen
                if (-not $SkipValidation) {
                    if (-not (Test-GroupExists -GroupName $groupName)) {
                        Write-LogMessage "Warnung: Gruppe '$groupName' existiert nicht" "WARN"
                        $results += @{
                            Group = $groupName
                            Success = $false
                            Error = "Gruppe existiert nicht"
                        }
                        $errorCount++
                        continue
                    }
                }
                
                # Füge Benutzer zur Gruppe hinzu
                Add-ADGroupMember -Identity $groupName -Members $UserName -Server $global:ADManager.DomainController -ErrorAction Stop
                
                Write-LogMessage "Benutzer '$UserName' erfolgreich zu Gruppe '$groupName' hinzugefügt" "INFO"
                $results += @{
                    Group = $groupName
                    Success = $true
                    Error = ""
                }
                $successCount++
            }
            catch {
                Write-LogMessage "Fehler beim Hinzufügen zu Gruppe '$groupName': $($_.Exception.Message)" "ERROR"
                $results += @{
                    Group = $groupName
                    Success = $false
                    Error = $_.Exception.Message
                }
                $errorCount++
            }
        }
        
        Write-LogMessage "Gruppenzuweisung abgeschlossen: $successCount erfolgreich, $errorCount Fehler" "INFO"
        
        return @{
            Success = $errorCount -eq 0
            SuccessCount = $successCount
            ErrorCount = $errorCount
            Results = $results
        }
    }
    catch {
        Write-LogMessage "Fehler bei der Gruppenzuweisung: $($_.Exception.Message)" "ERROR"
        return @{
            Success = $false
            SuccessCount = 0
            ErrorCount = $GroupNames.Count
            Results = @()
            Error = $_.Exception.Message
        }
    }
}

#endregion

#region [User Management Functions]

function Test-UserExists {
    <#
    .SYNOPSIS
    Prüft ob ein Benutzer bereits existiert
    
    .PARAMETER SamAccountName
    SamAccountName des zu prüfenden Benutzers
    
    .PARAMETER UserPrincipalName
    UPN des zu prüfenden Benutzers
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [string]$SamAccountName,
        
        [Parameter(Mandatory = $false)]
        [string]$UserPrincipalName
    )
    
    try {
        if (-not $global:ADManager.IsConnected) {
            throw "Keine AD-Verbindung verfügbar"
        }
        
        $filter = ""
        if ($SamAccountName) {
            $filter = "SamAccountName -eq '$SamAccountName'"
        }
        elseif ($UserPrincipalName) {
            $filter = "UserPrincipalName -eq '$UserPrincipalName'"
        }
        else {
            throw "SamAccountName oder UserPrincipalName muss angegeben werden"
        }
        
        $user = Get-ADUser -Filter $filter -Server $global:ADManager.DomainController -ErrorAction SilentlyContinue
        return $null -ne $user
    }
    catch {
        Write-LogMessage "Fehler bei der Benutzerprüfung: $($_.Exception.Message)" "ERROR"
        return $false
    }
}

function New-ADUserExtended {
    <#
    .SYNOPSIS
    Erstellt einen neuen AD-Benutzer mit erweiterten Optionen
    
    .PARAMETER UserData
    Hashtable mit Benutzerdaten
    
    .PARAMETER TargetOU
    Ziel-OU für den neuen Benutzer
    
    .PARAMETER GroupAssignments
    Array von Gruppennamen für automatische Zuweisung
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$UserData,
        
        [Parameter(Mandatory = $true)]
        [string]$TargetOU,
        
        [Parameter(Mandatory = $false)]
        [string[]]$GroupAssignments = @()
    )
    
    try {
        if (-not $global:ADManager.IsConnected) {
            throw "Keine AD-Verbindung verfügbar"
        }
        
        Write-LogMessage "Erstelle neuen AD-Benutzer: $($UserData.SamAccountName)" "INFO"
        
        # Validiere erforderliche Felder
        $requiredFields = @('GivenName', 'Surname', 'SamAccountName', 'UserPrincipalName')
        foreach ($field in $requiredFields) {
            if (-not $UserData.ContainsKey($field) -or [string]::IsNullOrWhiteSpace($UserData[$field])) {
                throw "Erforderliches Feld fehlt: $field"
            }
        }
        
        # Prüfe ob Benutzer bereits existiert
        if (Test-UserExists -SamAccountName $UserData.SamAccountName) {
            throw "Benutzer mit SamAccountName '$($UserData.SamAccountName)' existiert bereits"
        }
        
        if (Test-UserExists -UserPrincipalName $UserData.UserPrincipalName) {
            throw "Benutzer mit UPN '$($UserData.UserPrincipalName)' existiert bereits"
        }
        
        # Bereite Parameter für New-ADUser vor
        $adUserParams = @{
            Server = $global:ADManager.DomainController
            Path = $TargetOU
            GivenName = $UserData.GivenName
            Surname = $UserData.Surname
            SamAccountName = $UserData.SamAccountName
            UserPrincipalName = $UserData.UserPrincipalName
            Name = $UserData.DisplayName ?? "$($UserData.GivenName) $($UserData.Surname)"
            DisplayName = $UserData.DisplayName ?? "$($UserData.GivenName) $($UserData.Surname)"
            Enabled = $UserData.Enabled ?? $true
        }
        
        # Optionale Felder hinzufügen
        if ($UserData.ContainsKey('EmailAddress') -and -not [string]::IsNullOrWhiteSpace($UserData.EmailAddress)) {
            $adUserParams.EmailAddress = $UserData.EmailAddress
        }
        
        if ($UserData.ContainsKey('Department') -and -not [string]::IsNullOrWhiteSpace($UserData.Department)) {
            $adUserParams.Department = $UserData.Department
        }
        
        if ($UserData.ContainsKey('Title') -and -not [string]::IsNullOrWhiteSpace($UserData.Title)) {
            $adUserParams.Title = $UserData.Title
        }
        
        if ($UserData.ContainsKey('Description') -and -not [string]::IsNullOrWhiteSpace($UserData.Description)) {
            $adUserParams.Description = $UserData.Description
        }
        
        if ($UserData.ContainsKey('Password') -and -not [string]::IsNullOrWhiteSpace($UserData.Password)) {
            $securePassword = ConvertTo-SecureString -String $UserData.Password -AsPlainText -Force
            $adUserParams.AccountPassword = $securePassword
        }
        
        # Erstelle Benutzer
        $newUser = New-ADUser @adUserParams -PassThru -ErrorAction Stop
        
        Write-LogMessage "Benutzer '$($UserData.SamAccountName)' erfolgreich erstellt" "INFO"
        
        # Gruppenzuweisungen
        if ($GroupAssignments.Count -gt 0) {
            Write-LogMessage "Weise Benutzer zu Gruppen zu..." "INFO"
            $groupResult = Add-UserToGroups -UserName $UserData.SamAccountName -GroupNames $GroupAssignments
            
            if (-not $groupResult.Success) {
                Write-LogMessage "Warnung: Nicht alle Gruppenzuweisungen waren erfolgreich ($($groupResult.ErrorCount) Fehler)" "WARN"
            }
        }
        
        return @{
            Success = $true
            User = $newUser
            GroupAssignmentResult = $groupResult ?? @{ Success = $true; SuccessCount = 0; ErrorCount = 0 }
            Message = "Benutzer erfolgreich erstellt"
        }
    }
    catch {
        Write-LogMessage "Fehler beim Erstellen des Benutzers: $($_.Exception.Message)" "ERROR"
        return @{
            Success = $false
            User = $null
            GroupAssignmentResult = @{ Success = $false; SuccessCount = 0; ErrorCount = 0 }
            Message = $_.Exception.Message
            Error = $_.Exception.Message
        }
    }
}

function Search-ADUsers {
    <#
    .SYNOPSIS
    Sucht AD-Benutzer anhand verschiedener Kriterien
    
    .PARAMETER SearchTerm
    Suchbegriff für Name, SamAccountName oder E-Mail
    
    .PARAMETER MaxResults
    Maximale Anzahl der Ergebnisse
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$SearchTerm,
        
        [Parameter(Mandatory = $false)]
        [int]$MaxResults = 50
    )
    
    try {
        if (-not $global:ADManager.IsConnected) {
            throw "Keine AD-Verbindung verfügbar"
        }
        
        Write-LogMessage "Suche AD-Benutzer: '$SearchTerm'" "INFO"
        
        $filter = "(Name -like '*$SearchTerm*') -or (SamAccountName -like '*$SearchTerm*') -or (EmailAddress -like '*$SearchTerm*') -or (UserPrincipalName -like '*$SearchTerm*')"
        
        $users = Get-ADUser -Filter $filter -Server $global:ADManager.DomainController -Properties DisplayName, EmailAddress, Department, Title, Enabled -ResultSetSize $MaxResults |
                 Sort-Object DisplayName |
                 Select-Object SamAccountName, DisplayName, EmailAddress, Department, Title, Enabled, DistinguishedName
        
        Write-LogMessage "$($users.Count) Benutzer gefunden" "INFO"
        
        return $users
    }
    catch {
        Write-LogMessage "Fehler bei der Benutzersuche: $($_.Exception.Message)" "ERROR"
        return @()
    }
}

function Get-UserGroupMembership {
    <#
    .SYNOPSIS
    Ermittelt die Gruppenmitgliedschaften eines Benutzers
    
    .PARAMETER SamAccountName
    SamAccountName des Benutzers
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$SamAccountName
    )
    
    try {
        if (-not $global:ADManager.IsConnected) {
            throw "Keine AD-Verbindung verfügbar"
        }
        
        $user = Get-ADUser -Identity $SamAccountName -Server $global:ADManager.DomainController -Properties MemberOf -ErrorAction Stop
        
        $groups = $user.MemberOf | ForEach-Object {
            Get-ADGroup -Identity $_ -Server $global:ADManager.DomainController -Properties Description, GroupScope, GroupCategory
        } | Select-Object Name, Description, GroupScope, GroupCategory, DistinguishedName
        
        return $groups | Sort-Object Name
    }
    catch {
        Write-LogMessage "Fehler beim Ermitteln der Gruppenmitgliedschaften: $($_.Exception.Message)" "ERROR"
        return @()
    }
}

#endregion

#region [Validation Functions]

function Test-ADPermissions {
    <#
    .SYNOPSIS
    Testet die erforderlichen AD-Berechtigungen für Benutzerverwaltung
    #>
    
    try {
        if (-not $global:ADManager.IsConnected) {
            throw "Keine AD-Verbindung verfügbar"
        }
        
        Write-LogMessage "Teste AD-Berechtigungen..." "INFO"
        
        $permissions = @{
            CanCreateUsers = $false
            CanModifyUsers = $false
            CanReadUsers = $false
            CanManageGroups = $false
            CanReadOUs = $false
        }
        
        # Teste Leseberechtigung für Benutzer
        try {
            Get-ADUser -Filter "Name -like '*'" -ResultSetSize 1 -Server $global:ADManager.DomainController | Out-Null
            $permissions.CanReadUsers = $true
        }
        catch {
            Write-LogMessage "Keine Leseberechtigung für Benutzer" "WARN"
        }
        
        # Teste Leseberechtigung für OUs
        try {
            Get-ADOrganizationalUnit -Filter * -ResultSetSize 1 -Server $global:ADManager.DomainController | Out-Null
            $permissions.CanReadOUs = $true
        }
        catch {
            Write-LogMessage "Keine Leseberechtigung für OUs" "WARN"
        }
        
        # Teste Gruppenverwaltung
        try {
            Get-ADGroup -Filter "Name -like '*'" -ResultSetSize 1 -Server $global:ADManager.DomainController | Out-Null
            $permissions.CanManageGroups = $true
        }
        catch {
            Write-LogMessage "Keine Berechtigung für Gruppenverwaltung" "WARN"
        }
        
        # Weitere Tests könnten hier implementiert werden
        # z.B. Testen von Schreibberechtigungen durch Erstellen eines Test-Users
        
        Write-LogMessage "AD-Berechtigungstest abgeschlossen" "INFO"
        
        return $permissions
    }
    catch {
        Write-LogMessage "Fehler beim Testen der AD-Berechtigungen: $($_.Exception.Message)" "ERROR"
        return @{
            CanCreateUsers = $false
            CanModifyUsers = $false
            CanReadUsers = $false
            CanManageGroups = $false
            CanReadOUs = $false
            Error = $_.Exception.Message
        }
    }
}

#endregion

#region [Export Functions]
Export-ModuleMember -Function @(
    'Test-ADConnection',
    'Get-ADConnectionStatus',
    'Get-AvailableOUs',
    'Get-OUByName',
    'Get-AvailableGroups',
    'Get-GroupsByCategory',
    'Test-GroupExists',
    'Add-UserToGroups',
    'Test-UserExists',
    'New-ADUserExtended',
    'Search-ADUsers',
    'Get-UserGroupMembership',
    'Test-ADPermissions'
)
#endregion 
# SIG # Begin signature block
# MIIcCAYJKoZIhvcNAQcCoIIb+TCCG/UCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCD9F8kgpEGmCEb
# 1NJ2ugz254AU0Dxx8l4Fy6VitPiGYqCCFk4wggMQMIIB+KADAgECAhB3jzsyX9Cg
# jEi+sBC2rBMTMA0GCSqGSIb3DQEBCwUAMCAxHjAcBgNVBAMMFVBoaW5JVC1QU3Nj
# cmlwdHNfU2lnbjAeFw0yNTA3MDUwODI4MTZaFw0yNzA3MDUwODM4MTZaMCAxHjAc
# BgNVBAMMFVBoaW5JVC1QU3NjcmlwdHNfU2lnbjCCASIwDQYJKoZIhvcNAQEBBQAD
# ggEPADCCAQoCggEBALmz3o//iDA5MvAndTjGX7/AvzTSACClfuUR9WYK0f6Ut2dI
# mPxn+Y9pZlLjXIpZT0H2Lvxq5aSI+aYeFtuJ8/0lULYNCVT31Bf+HxervRBKsUyi
# W9+4PH6STxo3Pl4l56UNQMcWLPNjDORWRPWHn0f99iNtjI+L4tUC/LoWSs3obzxN
# 3uTypzlaPBxis2qFSTR5SWqFdZdRkcuI5LNsJjyc/QWdTYRrfmVqp0QrvcxzCv8u
# EiVuni6jkXfiE6wz+oeI3L2iR+ywmU6CUX4tPWoS9VTtmm7AhEpasRTmrrnSg20Q
# jiBa1eH5TyLAH3TcYMxhfMbN9a2xDX5pzM65EJUCAwEAAaNGMEQwDgYDVR0PAQH/
# BAQDAgeAMBMGA1UdJQQMMAoGCCsGAQUFBwMDMB0GA1UdDgQWBBQO7XOqiE/EYi+n
# IaR6YO5M2MUuVTANBgkqhkiG9w0BAQsFAAOCAQEAjYOKIwBu1pfbdvEFFaR/uY88
# peKPk0NnvNEc3dpGdOv+Fsgbz27JPvItITFd6AKMoN1W48YjQLaU22M2jdhjGN5i
# FSobznP5KgQCDkRsuoDKiIOTiKAAknjhoBaCCEZGw8SZgKJtWzbST36Thsdd/won
# ihLsuoLxfcFnmBfrXh3rTIvTwvfujob68s0Sf5derHP/F+nphTymlg+y4VTEAijk
# g2dhy8RAsbS2JYZT7K5aEJpPXMiOLBqd7oTGfM7y5sLk2LIM4cT8hzgz3v5yPMkF
# H2MdR//K403e1EKH9MsGuGAJZddVN8ppaiESoPLoXrgnw2SY5KCmhYw1xRFdjTCC
# BY0wggR1oAMCAQICEA6bGI750C3n79tQ4ghAGFowDQYJKoZIhvcNAQEMBQAwZTEL
# MAkGA1UEBhMCVVMxFTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3
# LmRpZ2ljZXJ0LmNvbTEkMCIGA1UEAxMbRGlnaUNlcnQgQXNzdXJlZCBJRCBSb290
# IENBMB4XDTIyMDgwMTAwMDAwMFoXDTMxMTEwOTIzNTk1OVowYjELMAkGA1UEBhMC
# VVMxFTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0
# LmNvbTEhMB8GA1UEAxMYRGlnaUNlcnQgVHJ1c3RlZCBSb290IEc0MIICIjANBgkq
# hkiG9w0BAQEFAAOCAg8AMIICCgKCAgEAv+aQc2jeu+RdSjwwIjBpM+zCpyUuySE9
# 8orYWcLhKac9WKt2ms2uexuEDcQwH/MbpDgW61bGl20dq7J58soR0uRf1gU8Ug9S
# H8aeFaV+vp+pVxZZVXKvaJNwwrK6dZlqczKU0RBEEC7fgvMHhOZ0O21x4i0MG+4g
# 1ckgHWMpLc7sXk7Ik/ghYZs06wXGXuxbGrzryc/NrDRAX7F6Zu53yEioZldXn1RY
# jgwrt0+nMNlW7sp7XeOtyU9e5TXnMcvak17cjo+A2raRmECQecN4x7axxLVqGDgD
# EI3Y1DekLgV9iPWCPhCRcKtVgkEy19sEcypukQF8IUzUvK4bA3VdeGbZOjFEmjNA
# vwjXWkmkwuapoGfdpCe8oU85tRFYF/ckXEaPZPfBaYh2mHY9WV1CdoeJl2l6SPDg
# ohIbZpp0yt5LHucOY67m1O+SkjqePdwA5EUlibaaRBkrfsCUtNJhbesz2cXfSwQA
# zH0clcOP9yGyshG3u3/y1YxwLEFgqrFjGESVGnZifvaAsPvoZKYz0YkH4b235kOk
# GLimdwHhD5QMIR2yVCkliWzlDlJRR3S+Jqy2QXXeeqxfjT/JvNNBERJb5RBQ6zHF
# ynIWIgnffEx1P2PsIV/EIFFrb7GrhotPwtZFX50g/KEexcCPorF+CiaZ9eRpL5gd
# LfXZqbId5RsCAwEAAaOCATowggE2MA8GA1UdEwEB/wQFMAMBAf8wHQYDVR0OBBYE
# FOzX44LScV1kTN8uZz/nupiuHA9PMB8GA1UdIwQYMBaAFEXroq/0ksuCMS1Ri6en
# IZ3zbcgPMA4GA1UdDwEB/wQEAwIBhjB5BggrBgEFBQcBAQRtMGswJAYIKwYBBQUH
# MAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBDBggrBgEFBQcwAoY3aHR0cDov
# L2NhY2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEUm9vdENBLmNy
# dDBFBgNVHR8EPjA8MDqgOKA2hjRodHRwOi8vY3JsMy5kaWdpY2VydC5jb20vRGln
# aUNlcnRBc3N1cmVkSURSb290Q0EuY3JsMBEGA1UdIAQKMAgwBgYEVR0gADANBgkq
# hkiG9w0BAQwFAAOCAQEAcKC/Q1xV5zhfoKN0Gz22Ftf3v1cHvZqsoYcs7IVeqRq7
# IviHGmlUIu2kiHdtvRoU9BNKei8ttzjv9P+Aufih9/Jy3iS8UgPITtAq3votVs/5
# 9PesMHqai7Je1M/RQ0SbQyHrlnKhSLSZy51PpwYDE3cnRNTnf+hZqPC/Lwum6fI0
# POz3A8eHqNJMQBk1RmppVLC4oVaO7KTVPeix3P0c2PR3WlxUjG/voVA9/HYJaISf
# b8rbII01YBwCA8sgsKxYoA5AY8WYIsGyWfVVa88nq2x2zm8jLfR+cWojayL/ErhU
# LSd+2DrZ8LaHlv1b0VysGMNNn3O3AamfV6peKOK5lDCCBrQwggScoAMCAQICEA3H
# rFcF/yGZLkBDIgw6SYYwDQYJKoZIhvcNAQELBQAwYjELMAkGA1UEBhMCVVMxFTAT
# BgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNvbTEh
# MB8GA1UEAxMYRGlnaUNlcnQgVHJ1c3RlZCBSb290IEc0MB4XDTI1MDUwNzAwMDAw
# MFoXDTM4MDExNDIzNTk1OVowaTELMAkGA1UEBhMCVVMxFzAVBgNVBAoTDkRpZ2lD
# ZXJ0LCBJbmMuMUEwPwYDVQQDEzhEaWdpQ2VydCBUcnVzdGVkIEc0IFRpbWVTdGFt
# cGluZyBSU0E0MDk2IFNIQTI1NiAyMDI1IENBMTCCAiIwDQYJKoZIhvcNAQEBBQAD
# ggIPADCCAgoCggIBALR4MdMKmEFyvjxGwBysddujRmh0tFEXnU2tjQ2UtZmWgyxU
# 7UNqEY81FzJsQqr5G7A6c+Gh/qm8Xi4aPCOo2N8S9SLrC6Kbltqn7SWCWgzbNfiR
# +2fkHUiljNOqnIVD/gG3SYDEAd4dg2dDGpeZGKe+42DFUF0mR/vtLa4+gKPsYfwE
# u7EEbkC9+0F2w4QJLVSTEG8yAR2CQWIM1iI5PHg62IVwxKSpO0XaF9DPfNBKS7Za
# zch8NF5vp7eaZ2CVNxpqumzTCNSOxm+SAWSuIr21Qomb+zzQWKhxKTVVgtmUPAW3
# 5xUUFREmDrMxSNlr/NsJyUXzdtFUUt4aS4CEeIY8y9IaaGBpPNXKFifinT7zL2gd
# FpBP9qh8SdLnEut/GcalNeJQ55IuwnKCgs+nrpuQNfVmUB5KlCX3ZA4x5HHKS+rq
# BvKWxdCyQEEGcbLe1b8Aw4wJkhU1JrPsFfxW1gaou30yZ46t4Y9F20HHfIY4/6vH
# espYMQmUiote8ladjS/nJ0+k6MvqzfpzPDOy5y6gqztiT96Fv/9bH7mQyogxG9QE
# PHrPV6/7umw052AkyiLA6tQbZl1KhBtTasySkuJDpsZGKdlsjg4u70EwgWbVRSX1
# Wd4+zoFpp4Ra+MlKM2baoD6x0VR4RjSpWM8o5a6D8bpfm4CLKczsG7ZrIGNTAgMB
# AAGjggFdMIIBWTASBgNVHRMBAf8ECDAGAQH/AgEAMB0GA1UdDgQWBBTvb1NK6eQG
# fHrK4pBW9i/USezLTjAfBgNVHSMEGDAWgBTs1+OC0nFdZEzfLmc/57qYrhwPTzAO
# BgNVHQ8BAf8EBAMCAYYwEwYDVR0lBAwwCgYIKwYBBQUHAwgwdwYIKwYBBQUHAQEE
# azBpMCQGCCsGAQUFBzABhhhodHRwOi8vb2NzcC5kaWdpY2VydC5jb20wQQYIKwYB
# BQUHMAKGNWh0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0
# ZWRSb290RzQuY3J0MEMGA1UdHwQ8MDowOKA2oDSGMmh0dHA6Ly9jcmwzLmRpZ2lj
# ZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRSb290RzQuY3JsMCAGA1UdIAQZMBcwCAYG
# Z4EMAQQCMAsGCWCGSAGG/WwHATANBgkqhkiG9w0BAQsFAAOCAgEAF877FoAc/gc9
# EXZxML2+C8i1NKZ/zdCHxYgaMH9Pw5tcBnPw6O6FTGNpoV2V4wzSUGvI9NAzaoQk
# 97frPBtIj+ZLzdp+yXdhOP4hCFATuNT+ReOPK0mCefSG+tXqGpYZ3essBS3q8nL2
# UwM+NMvEuBd/2vmdYxDCvwzJv2sRUoKEfJ+nN57mQfQXwcAEGCvRR2qKtntujB71
# WPYAgwPyWLKu6RnaID/B0ba2H3LUiwDRAXx1Neq9ydOal95CHfmTnM4I+ZI2rVQf
# jXQA1WSjjf4J2a7jLzWGNqNX+DF0SQzHU0pTi4dBwp9nEC8EAqoxW6q17r0z0noD
# js6+BFo+z7bKSBwZXTRNivYuve3L2oiKNqetRHdqfMTCW/NmKLJ9M+MtucVGyOxi
# Df06VXxyKkOirv6o02OoXN4bFzK0vlNMsvhlqgF2puE6FndlENSmE+9JGYxOGLS/
# D284NHNboDGcmWXfwXRy4kbu4QFhOm0xJuF2EZAOk5eCkhSxZON3rGlHqhpB/8Ml
# uDezooIs8CVnrpHMiD2wL40mm53+/j7tFaxYKIqL0Q4ssd8xHZnIn/7GELH3IdvG
# 2XlM9q7WP/UwgOkw/HQtyRN62JK4S1C8uw3PdBunvAZapsiI5YKdvlarEvf8EA+8
# hcpSM9LHJmyrxaFtoza2zNaQ9k+5t1wwggbtMIIE1aADAgECAhAKgO8YS43xBYLR
# xHanlXRoMA0GCSqGSIb3DQEBCwUAMGkxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5E
# aWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1l
# U3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYgMjAyNSBDQTEwHhcNMjUwNjA0MDAwMDAw
# WhcNMzYwOTAzMjM1OTU5WjBjMQswCQYDVQQGEwJVUzEXMBUGA1UEChMORGlnaUNl
# cnQsIEluYy4xOzA5BgNVBAMTMkRpZ2lDZXJ0IFNIQTI1NiBSU0E0MDk2IFRpbWVz
# dGFtcCBSZXNwb25kZXIgMjAyNSAxMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIIC
# CgKCAgEA0EasLRLGntDqrmBWsytXum9R/4ZwCgHfyjfMGUIwYzKomd8U1nH7C8Dr
# 0cVMF3BsfAFI54um8+dnxk36+jx0Tb+k+87H9WPxNyFPJIDZHhAqlUPt281mHrBb
# ZHqRK71Em3/hCGC5KyyneqiZ7syvFXJ9A72wzHpkBaMUNg7MOLxI6E9RaUueHTQK
# WXymOtRwJXcrcTTPPT2V1D/+cFllESviH8YjoPFvZSjKs3SKO1QNUdFd2adw44wD
# cKgH+JRJE5Qg0NP3yiSyi5MxgU6cehGHr7zou1znOM8odbkqoK+lJ25LCHBSai25
# CFyD23DZgPfDrJJJK77epTwMP6eKA0kWa3osAe8fcpK40uhktzUd/Yk0xUvhDU6l
# vJukx7jphx40DQt82yepyekl4i0r8OEps/FNO4ahfvAk12hE5FVs9HVVWcO5J4dV
# mVzix4A77p3awLbr89A90/nWGjXMGn7FQhmSlIUDy9Z2hSgctaepZTd0ILIUbWuh
# KuAeNIeWrzHKYueMJtItnj2Q+aTyLLKLM0MheP/9w6CtjuuVHJOVoIJ/DtpJRE7C
# e7vMRHoRon4CWIvuiNN1Lk9Y+xZ66lazs2kKFSTnnkrT3pXWETTJkhd76CIDBbTR
# ofOsNyEhzZtCGmnQigpFHti58CSmvEyJcAlDVcKacJ+A9/z7eacCAwEAAaOCAZUw
# ggGRMAwGA1UdEwEB/wQCMAAwHQYDVR0OBBYEFOQ7/PIx7f391/ORcWMZUEPPYYzo
# MB8GA1UdIwQYMBaAFO9vU0rp5AZ8esrikFb2L9RJ7MtOMA4GA1UdDwEB/wQEAwIH
# gDAWBgNVHSUBAf8EDDAKBggrBgEFBQcDCDCBlQYIKwYBBQUHAQEEgYgwgYUwJAYI
# KwYBBQUHMAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBdBggrBgEFBQcwAoZR
# aHR0cDovL2NhY2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0VHJ1c3RlZEc0VGlt
# ZVN0YW1waW5nUlNBNDA5NlNIQTI1NjIwMjVDQTEuY3J0MF8GA1UdHwRYMFYwVKBS
# oFCGTmh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRHNFRp
# bWVTdGFtcGluZ1JTQTQwOTZTSEEyNTYyMDI1Q0ExLmNybDAgBgNVHSAEGTAXMAgG
# BmeBDAEEAjALBglghkgBhv1sBwEwDQYJKoZIhvcNAQELBQADggIBAGUqrfEcJwS5
# rmBB7NEIRJ5jQHIh+OT2Ik/bNYulCrVvhREafBYF0RkP2AGr181o2YWPoSHz9iZE
# N/FPsLSTwVQWo2H62yGBvg7ouCODwrx6ULj6hYKqdT8wv2UV+Kbz/3ImZlJ7YXwB
# D9R0oU62PtgxOao872bOySCILdBghQ/ZLcdC8cbUUO75ZSpbh1oipOhcUT8lD8QA
# GB9lctZTTOJM3pHfKBAEcxQFoHlt2s9sXoxFizTeHihsQyfFg5fxUFEp7W42fNBV
# N4ueLaceRf9Cq9ec1v5iQMWTFQa0xNqItH3CPFTG7aEQJmmrJTV3Qhtfparz+BW6
# 0OiMEgV5GWoBy4RVPRwqxv7Mk0Sy4QHs7v9y69NBqycz0BZwhB9WOfOu/CIJnzkQ
# TwtSSpGGhLdjnQ4eBpjtP+XB3pQCtv4E5UCSDag6+iX8MmB10nfldPF9SVD7weCC
# 3yXZi/uuhqdwkgVxuiMFzGVFwYbQsiGnoa9F5AaAyBjFBtXVLcKtapnMG3VH3EmA
# p/jsJ3FVF3+d1SVDTmjFjLbNFZUWMXuZyvgLfgyPehwJVxwC+UpX2MSey2ueIu9T
# HFVkT+um1vshETaWyQo8gmBto/m3acaP9QsuLj3FNwFlTxq25+T4QwX9xa6ILs84
# ZPvmpovq90K8eWyG2N01c4IhSOxqt81nMYIFEDCCBQwCAQEwNDAgMR4wHAYDVQQD
# DBVQaGluSVQtUFNzY3JpcHRzX1NpZ24CEHePOzJf0KCMSL6wELasExMwDQYJYIZI
# AWUDBAIBBQCggYQwGAYKKwYBBAGCNwIBDDEKMAigAoAAoQKAADAZBgkqhkiG9w0B
# CQMxDAYKKwYBBAGCNwIBBDAcBgorBgEEAYI3AgELMQ4wDAYKKwYBBAGCNwIBFTAv
# BgkqhkiG9w0BCQQxIgQgPjEXRyrCRv05yfBRsPpIzv9/xD8P9bgQEL5NWIrfyN8w
# DQYJKoZIhvcNAQEBBQAEggEAFdLVPiQ9+NHSFmDsIbadItTb8In3aXki4OB6Lfhw
# hktOD/G/vCWv5yciW7wh+h/uZBl0VXbbgbszocXLvyr3A1WUX0TTmuhJ2MlL4mtk
# piNMzaQc5NsjkoEq66pNIuGYC+JxfRB724jmMkiP9qCuSd2qQTVRbPhCimxeHz5W
# zOchHwmVa5vuOzCHhhSRQQMPx1e+wVMcHyaCcDITq1fyv3EUXuYnlTMma2+i+tFx
# 3p/MX5kQtZoLdIEUnK52QyCyQjRxx3TKWJXHKeo522S681DLa5naJgpjSHAXMF1T
# WATDaq0sOJ+ui2NuFpEwBAhIpRiW/kiJLiaeI1V0GG83JKGCAyYwggMiBgkqhkiG
# 9w0BCQYxggMTMIIDDwIBATB9MGkxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdp
# Q2VydCwgSW5jLjFBMD8GA1UEAxM4RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3Rh
# bXBpbmcgUlNBNDA5NiBTSEEyNTYgMjAyNSBDQTECEAqA7xhLjfEFgtHEdqeVdGgw
# DQYJYIZIAWUDBAIBBQCgaTAYBgkqhkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqG
# SIb3DQEJBTEPFw0yNTA3MTIwODAwMjlaMC8GCSqGSIb3DQEJBDEiBCBj5ePL2NQm
# bAqljAzrd4P1444cvFAhXnWExDLXComDbDANBgkqhkiG9w0BAQEFAASCAgBnc2Dv
# vOLwBQWaxcIWwhm8con58flGlV9eV+i+gVcvTdViSYMLVXG0BxypbwwXnjTzMIY3
# TpG6v922zbFZNE+uIP1gRp5vUAzqsNkhSgdShk1zkWhdtxYLZW7YDaHcTqB2IYwp
# 6fNnnwMj1Poy6Xgk/vS8/smIwDn85sjh6e+hL12TrJAmlwIEpiruKoP5Dgl/73f/
# 6h+SzZGmECKu1XO8PiqAExUf2Iy6P3TUTlLEEgRcXdpQLOosL4/t2sR3ohj4S9Hj
# +K05EXHKPjzuzvKSGMcEP2yBajiF2Xdx0G5F6V9eKLovbqrle2jaRvW627xwuWK8
# qKZU5WRWsxtQrDBEtUFooWxwIIYAGdyCoFxmr3KQgXoUx0okq3H1XpOTBJQAa3qm
# KaJCW06SL/yE1v3fPzWn3jypBLwYNiidGr0ATYjk548CGQydzjSzSdc2GBY6bhvo
# djC7l1LUNLRntBLAb4bphR1+EHFkux0GynqWfDk/EaZf110RPe4GsvBU/+HQ3C/a
# w1t7R/SH3j7xgCqtf9timV9WvEU23fIjABzcq99m9nHtf/PbpMTb03EODF43NoBP
# Roo4/aF7OD4Y94lwHbqJcDfTAlYdaQeOg+lX2rQ8TV8xfTk7Qp8tMBx5yUQ7lbiV
# zax7EZUiyuNAuP2I7KgJirwC3l18ZmvlTbewJA==
# SIG # End signature block
