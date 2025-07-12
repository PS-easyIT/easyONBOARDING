#region [Settings Manager Module]
# Manages application settings, AD group mappings, and configuration persistence
# Version: 1.4.XX
# Author: IT Team

#region [Module Variables]
$global:SettingsManager = @{
    ConfigPath = ""
    SettingsData = @{}
    GroupMappings = @{}
    ADDomains = @()
    IsInitialized = $false
}
#endregion

#region [Core Settings Functions]

function Initialize-SettingsManager {
    <#
    .SYNOPSIS
    Initialisiert den Settings Manager und lädt alle Konfigurationsdaten
    
    .PARAMETER ConfigPath
    Pfad zur INI-Konfigurationsdatei
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ConfigPath
    )
    
    try {
        Write-LogMessage "Initialisiere Settings Manager..." "INFO"
        
        $global:SettingsManager.ConfigPath = $ConfigPath
        
        # Lade Basis-Konfiguration
        $global:SettingsManager.SettingsData = Get-IniContent -Path $ConfigPath
        
        # Initialisiere Standard-Gruppenmappings
        Initialize-DefaultGroupMappings
        
        # Lade gespeicherte Gruppenmappings
        Import-GroupMappings
        
        # Lade verfügbare AD-Domains
        Get-AvailableADDomains
        
        $global:SettingsManager.IsInitialized = $true
        Write-LogMessage "Settings Manager erfolgreich initialisiert" "INFO"
        
        return $true
    }
    catch {
        Write-LogMessage "Fehler beim Initialisieren des Settings Manager: $($_.Exception.Message)" "ERROR"
        return $false
    }
}

function Initialize-DefaultGroupMappings {
    <#
    .SYNOPSIS
    Initialisiert die Standard-Gruppenmappings für verschiedene Funktionsbereiche
    #>
    
    $global:SettingsManager.GroupMappings = @{
        # Terminal Services / RDP Access
        TerminalUsers = @{
            DisplayName = "Terminal Services Benutzer"
            Description = "Benutzer mit Remote Desktop Zugriff"
            Type = "Security"
            Scope = "DomainLocal"
            DefaultGroups = @("Remote Desktop Users", "Terminal Server Users")
            MappedGroups = @()
            AutoAssign = $false
        }
        
        # Microsoft 365 Groups
        MS365Groups = @{
            DisplayName = "Microsoft 365 Gruppen"
            Description = "MS365 Lizenzen und Collaboration"
            Type = "Both"
            Scope = "Universal"
            DefaultGroups = @("MS365-BasicUsers", "MS365-E3Users", "MS365-E5Users")
            MappedGroups = @()
            AutoAssign = $true
        }
        
        # VPN Access Groups
        VPNGroups = @{
            DisplayName = "VPN Zugangsgruppen"
            Description = "VPN-Verbindungen und Remote Access"
            Type = "Security"
            Scope = "Global"
            DefaultGroups = @("VPN-Users", "VPN-RemoteWorkers", "VPN-Admins")
            MappedGroups = @()
            AutoAssign = $false
        }
        
        # Security Groups
        SecurityGroups = @{
            DisplayName = "Sicherheitsgruppen"
            Description = "Basis-Sicherheitsgruppen für Ressourcenzugriff"
            Type = "Security"
            Scope = "DomainLocal"
            DefaultGroups = @("Domain Users", "Authenticated Users")
            MappedGroups = @()
            AutoAssign = $true
        }
        
        # Application Groups (Jira/Confluence etc.)
        ApplicationGroups = @{
            DisplayName = "Anwendungsgruppen"
            Description = "Jira, Confluence und andere Anwendungen"
            Type = "Security"
            Scope = "Global"
            DefaultGroups = @("Jira-Users", "Confluence-Users", "Wiki-Contributors")
            MappedGroups = @()
            AutoAssign = $false
        }
        
        # Departmental Groups
        DepartmentalGroups = @{
            DisplayName = "Abteilungsgruppen"
            Description = "Organisatorische Abteilungen"
            Type = "Security"
            Scope = "Global"
            DefaultGroups = @("IT-Department", "HR-Department", "Finance-Department", "Marketing-Department")
            MappedGroups = @()
            AutoAssign = $false
        }
        
        # Management Groups
        ManagementGroups = @{
            DisplayName = "Management Gruppen"
            Description = "Führungskräfte und erweiterte Berechtigungen"
            Type = "Security"
            Scope = "Universal"
            DefaultGroups = @("Team-Leads", "Department-Managers", "Senior-Management")
            MappedGroups = @()
            AutoAssign = $false
        }
        
        # Distribution Lists
        DistributionLists = @{
            DisplayName = "Verteilerlisten"
            Description = "E-Mail-Verteiler für Kommunikation"
            Type = "Distribution"
            Scope = "Universal"
            DefaultGroups = @("All-Employees", "IT-Announcements", "Company-News")
            MappedGroups = @()
            AutoAssign = $true
        }
    }
}

function Get-GroupMappingCategories {
    <#
    .SYNOPSIS
    Gibt alle verfügbaren Gruppenmapping-Kategorien zurück
    #>
    
    if (-not $global:SettingsManager.IsInitialized) {
        throw "Settings Manager ist nicht initialisiert"
    }
    
    return $global:SettingsManager.GroupMappings.Keys
}

function Get-GroupMappingByCategory {
    <#
    .SYNOPSIS
    Gibt die Gruppenmappings für eine spezifische Kategorie zurück
    
    .PARAMETER Category
    Die Kategorie der Gruppenmappings
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Category
    )
    
    if (-not $global:SettingsManager.IsInitialized) {
        throw "Settings Manager ist nicht initialisiert"
    }
    
    if ($global:SettingsManager.GroupMappings.ContainsKey($Category)) {
        return $global:SettingsManager.GroupMappings[$Category]
    }
    
    return $null
}

function Set-GroupMappingForCategory {
    <#
    .SYNOPSIS
    Setzt die Gruppenmappings für eine spezifische Kategorie
    
    .PARAMETER Category
    Die Kategorie der Gruppenmappings
    
    .PARAMETER Groups
    Array von AD-Gruppennamen
    
    .PARAMETER AutoAssign
    Ob die Gruppen automatisch zugewiesen werden sollen
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Category,
        
        [Parameter(Mandatory = $true)]
        [string[]]$Groups,
        
        [Parameter(Mandatory = $false)]
        [bool]$AutoAssign = $false
    )
    
    try {
        if (-not $global:SettingsManager.IsInitialized) {
            throw "Settings Manager ist nicht initialisiert"
        }
        
        if ($global:SettingsManager.GroupMappings.ContainsKey($Category)) {
            $global:SettingsManager.GroupMappings[$Category].MappedGroups = $Groups
            $global:SettingsManager.GroupMappings[$Category].AutoAssign = $AutoAssign
            
            Write-LogMessage "Gruppenmapping für Kategorie '$Category' aktualisiert: $($Groups -join ', ')" "INFO"
            
            # Speichere Änderungen
            Export-GroupMappings
            
            return $true
        }
        else {
            throw "Unbekannte Kategorie: $Category"
        }
    }
    catch {
        Write-LogMessage "Fehler beim Setzen der Gruppenmappings: $($_.Exception.Message)" "ERROR"
        return $false
    }
}

function Get-AvailableADDomains {
    <#
    .SYNOPSIS
    Ermittelt verfügbare Active Directory Domains
    #>
    
    try {
        Write-LogMessage "Ermittle verfügbare AD-Domains..." "INFO"
        
        # Versuche lokale Domain zu ermitteln
        $localDomain = $env:USERDNSDOMAIN
        $domains = @()
        
        if ($localDomain) {
            $domains += $localDomain
        }
        
        # Versuche weitere Domains über AD zu ermitteln
        try {
            Import-Module ActiveDirectory -ErrorAction SilentlyContinue
            $forest = Get-ADForest -ErrorAction SilentlyContinue
            if ($forest) {
                $domains += $forest.Domains
            }
        }
        catch {
            Write-LogMessage "Warnung: Konnte AD-Forest nicht abfragen: $($_.Exception.Message)" "WARN"
        }
        
        # Entferne Duplikate
        $global:SettingsManager.ADDomains = $domains | Sort-Object | Get-Unique
        
        Write-LogMessage "Gefundene AD-Domains: $($global:SettingsManager.ADDomains -join ', ')" "INFO"
        
        return $global:SettingsManager.ADDomains
    }
    catch {
        Write-LogMessage "Fehler beim Ermitteln der AD-Domains: $($_.Exception.Message)" "ERROR"
        return @()
    }
}

function Import-GroupMappings {
    <#
    .SYNOPSIS
    Importiert gespeicherte Gruppenmappings aus der Konfiguration
    #>
    
    try {
        $mappingsFile = Join-Path (Split-Path $global:SettingsManager.ConfigPath) "GroupMappings.json"
        
        if (Test-Path $mappingsFile) {
            $savedMappings = Get-Content $mappingsFile | ConvertFrom-Json
            
            foreach ($category in $savedMappings.PSObject.Properties.Name) {
                if ($global:SettingsManager.GroupMappings.ContainsKey($category)) {
                    $global:SettingsManager.GroupMappings[$category].MappedGroups = $savedMappings.$category.MappedGroups
                    $global:SettingsManager.GroupMappings[$category].AutoAssign = $savedMappings.$category.AutoAssign
                }
            }
            
            Write-LogMessage "Gruppenmappings erfolgreich importiert aus: $mappingsFile" "INFO"
        }
        else {
            Write-LogMessage "Keine gespeicherten Gruppenmappings gefunden, verwende Defaults" "INFO"
        }
    }
    catch {
        Write-LogMessage "Fehler beim Importieren der Gruppenmappings: $($_.Exception.Message)" "ERROR"
    }
}

function Export-GroupMappings {
    <#
    .SYNOPSIS
    Exportiert aktuelle Gruppenmappings in eine JSON-Datei
    #>
    
    try {
        $mappingsFile = Join-Path (Split-Path $global:SettingsManager.ConfigPath) "GroupMappings.json"
        
        $exportData = @{}
        foreach ($category in $global:SettingsManager.GroupMappings.Keys) {
            $exportData[$category] = @{
                MappedGroups = $global:SettingsManager.GroupMappings[$category].MappedGroups
                AutoAssign = $global:SettingsManager.GroupMappings[$category].AutoAssign
            }
        }
        
        $exportData | ConvertTo-Json -Depth 3 | Set-Content $mappingsFile
        
        Write-LogMessage "Gruppenmappings erfolgreich exportiert nach: $mappingsFile" "INFO"
    }
    catch {
        Write-LogMessage "Fehler beim Exportieren der Gruppenmappings: $($_.Exception.Message)" "ERROR"
    }
}

function Get-SettingValue {
    <#
    .SYNOPSIS
    Gibt einen spezifischen Einstellungswert zurück
    
    .PARAMETER Section
    Die INI-Sektion
    
    .PARAMETER Key
    Der Schlüssel
    
    .PARAMETER DefaultValue
    Standardwert falls nicht gefunden
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Section,
        
        [Parameter(Mandatory = $true)]
        [string]$Key,
        
        [Parameter(Mandatory = $false)]
        [string]$DefaultValue = ""
    )
    
    if (-not $global:SettingsManager.IsInitialized) {
        return $DefaultValue
    }
    
    if ($global:SettingsManager.SettingsData.ContainsKey($Section) -and 
        $global:SettingsManager.SettingsData[$Section].ContainsKey($Key)) {
        return $global:SettingsManager.SettingsData[$Section][$Key]
    }
    
    return $DefaultValue
}

function Set-SettingValue {
    <#
    .SYNOPSIS
    Setzt einen spezifischen Einstellungswert
    
    .PARAMETER Section
    Die INI-Sektion
    
    .PARAMETER Key
    Der Schlüssel
    
    .PARAMETER Value
    Der neue Wert
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Section,
        
        [Parameter(Mandatory = $true)]
        [string]$Key,
        
        [Parameter(Mandatory = $true)]
        [string]$Value
    )
    
    try {
        if (-not $global:SettingsManager.IsInitialized) {
            throw "Settings Manager ist nicht initialisiert"
        }
        
        if (-not $global:SettingsManager.SettingsData.ContainsKey($Section)) {
            $global:SettingsManager.SettingsData[$Section] = @{}
        }
        
        $global:SettingsManager.SettingsData[$Section][$Key] = $Value
        
        Write-LogMessage "Einstellung aktualisiert: [$Section] $Key = $Value" "INFO"
        
        return $true
    }
    catch {
        Write-LogMessage "Fehler beim Setzen der Einstellung: $($_.Exception.Message)" "ERROR"
        return $false
    }
}

#endregion

#region [Export Functions]
Export-ModuleMember -Function @(
    'Initialize-SettingsManager',
    'Get-GroupMappingCategories',
    'Get-GroupMappingByCategory',
    'Set-GroupMappingForCategory',
    'Get-AvailableADDomains',
    'Import-GroupMappings',
    'Export-GroupMappings',
    'Get-SettingValue',
    'Set-SettingValue'
)
#endregion 
# SIG # Begin signature block
# MIIcCAYJKoZIhvcNAQcCoIIb+TCCG/UCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDxdYP4ueJ0202a
# 0pMNGgVI9YuAJNlVdzFo/CVG2m3bg6CCFk4wggMQMIIB+KADAgECAhB3jzsyX9Cg
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
# BgkqhkiG9w0BCQQxIgQgA6p/Y3nzciVkFH0BXkCngWEr1ZyfMCAgBFPV728k2Pkw
# DQYJKoZIhvcNAQEBBQAEggEAt5vQc3y7Xxu18/HIdG1woP/R6BQYh5uYhCH+n8WN
# Om7pENKqd1C7spURvTolMlEExLQH5NDuqM41nQJsHsAHXpcQP+N6eqMD1N2k8vMI
# zrgDM8QC3d/ql3SQouJ4GwMOGCt47NpAYTQQiPrr710sxInRcDYYQlS4J9sOXikz
# d/hTCzBelFMJF0nhBUP0ASNxCcosGt2c+W0nfNWNvqpKv24pWcalryiyJdapO73J
# xutVnVPAoWmLvKgekJFQNWo8JFn0ox+8uh9j4sxzS4IfWE26nTLs+9f/AEvMrs9t
# i6s+HLEY455iBrfRtuKVvzXKudSJ/oN9KqS3Z6bqznBLpKGCAyYwggMiBgkqhkiG
# 9w0BCQYxggMTMIIDDwIBATB9MGkxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdp
# Q2VydCwgSW5jLjFBMD8GA1UEAxM4RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3Rh
# bXBpbmcgUlNBNDA5NiBTSEEyNTYgMjAyNSBDQTECEAqA7xhLjfEFgtHEdqeVdGgw
# DQYJYIZIAWUDBAIBBQCgaTAYBgkqhkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqG
# SIb3DQEJBTEPFw0yNTA3MTExOTMyMjdaMC8GCSqGSIb3DQEJBDEiBCCh9e1cmfry
# TCZTeeeYMPameivuJHZTpCt3eMthrJ0wwzANBgkqhkiG9w0BAQEFAASCAgBB0T/f
# vq+4jDotz4njhXHy6WVFNKvNrMEwWIeV/ZhlP6f7CwgXFr4S82wvOiw4yzOpZV3C
# /DQROzyiCT+9Nd6e+7Q3ScSBwLfO3Vz9xxQuxPpXt/VvsbcLuDYnnsCeOsf12JZF
# tRXrGm9gqzLjcteVTXx+iTEQFmI8MaAUVDnX/U3nFAoZmM/P6RA+4PltXJwKIWwo
# cg/dAqoy/mgaws5jkAjUPLzGulQXwQ8IzzpwBTOkR76SRwOeCX/yzxb9Zd/A5A8q
# sk85wM1XsQwIfyrjb4A7GS8lBOk10e+tG9bbzeTyaGHKGPuF1uOVfXKGuVmlKIBJ
# DD47O2zL1/iJH4gwbV+dD0+5w7BCq9qcb/VazD3u+aorl/mwbH094ahuZdlZG24L
# sBLEKh+Urd7C0bFEbnlsM3wU/8CbQRzmbpPR2CvkBZXXkVioQJ6hgFC3gAdQ7iNr
# dLsviBGhQHUe01SCUoG0TCH9XQjjEj7neqQlm5CKGJdxdqJNfhORuAMhsClsgoUk
# k/JaRiqnPBAHfgk18tlTIwcxBRGSyBZsRawWEkT0fBhbzTZf8MFG0G2L6mU9v26i
# zg5t7yJGl+bdFwTM1on4Aru80SSjrmA09j6CzKwTxH1SrFiq1hGFObLi+BNsDmeE
# 49TQCmhOwXD2yBRWx7tBbNPJYkvDQkACTw+Emw==
# SIG # End signature block
