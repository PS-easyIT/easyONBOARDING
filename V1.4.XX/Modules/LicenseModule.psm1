# LicenseModule.psm1
# Module for Microsoft 365 license management

# Module-level variables
$script:Configuration = @{}
$script:LogFunction = $null
$script:MSGraphConnected = $false

function Initialize-LicenseModule {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $false)]
        [hashtable]$Configuration = @{},
        
        [Parameter(Mandatory = $false)]
        [scriptblock]$LogFunction
    )
    
    $script:Configuration = $Configuration
    $script:LogFunction = $LogFunction
    
    # Check if Microsoft Graph module is available
    if (Get-Module -ListAvailable -Name Microsoft.Graph) {
        Write-ModuleLog "Microsoft Graph module is available" "INFO"
    } else {
        Write-ModuleLog "Microsoft Graph module not found. License management will be limited." "WARN"
    }
    
    Write-ModuleLog "LicenseModule initialized" "INFO"
}

function Write-ModuleLog {
    param(
        [string]$Message,
        [string]$Level = "INFO"
    )
    
    if ($script:LogFunction) {
        & $script:LogFunction $Message $Level
    } else {
        Write-Host "[$Level] $Message"
    }
}

function Connect-MSGraphForLicensing {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $false)]
        [string[]]$Scopes = @("User.ReadWrite.All", "Directory.ReadWrite.All")
    )
    
    try {
        if (-not (Get-Module -Name Microsoft.Graph -ErrorAction SilentlyContinue)) {
            Import-Module Microsoft.Graph -ErrorAction Stop
        }
        
        # Connect to Microsoft Graph
        Connect-MgGraph -Scopes $Scopes -ErrorAction Stop
        
        $script:MSGraphConnected = $true
        Write-ModuleLog "Connected to Microsoft Graph successfully" "INFO"
        
        return @{
            Success = $true
            Message = "Connected to Microsoft Graph"
        }
    }
    catch {
        Write-ModuleLog "Error connecting to Microsoft Graph: $($_.Exception.Message)" "ERROR"
        return @{
            Success = $false
            Message = "Error connecting to Microsoft Graph: $($_.Exception.Message)"
        }
    }
}

function Get-AvailableLicenses {
    [CmdletBinding()]
    param()
    
    try {
        if (-not $script:MSGraphConnected) {
            $connectResult = Connect-MSGraphForLicensing
            if (-not $connectResult.Success) {
                throw "Failed to connect to Microsoft Graph"
            }
        }
        
        Write-ModuleLog "Retrieving available licenses" "INFO"
        
        # Get subscribed SKUs
        $licenses = Get-MgSubscribedSku -All | Select-Object -Property SkuId, SkuPartNumber, 
            @{Name="Available"; Expression={$_.PrepaidUnits.Enabled - $_.ConsumedUnits}},
            @{Name="Total"; Expression={$_.PrepaidUnits.Enabled}},
            @{Name="Consumed"; Expression={$_.ConsumedUnits}}
        
        # Map to friendly names if configured
        if ($script:Configuration.ContainsKey("LicenseMapping")) {
            foreach ($license in $licenses) {
                if ($script:Configuration.LicenseMapping.ContainsKey($license.SkuPartNumber)) {
                    $license | Add-Member -NotePropertyName "FriendlyName" -NotePropertyValue $script:Configuration.LicenseMapping[$license.SkuPartNumber] -Force
                } else {
                    $license | Add-Member -NotePropertyName "FriendlyName" -NotePropertyValue $license.SkuPartNumber -Force
                }
            }
        }
        
        Write-ModuleLog "Retrieved $($licenses.Count) license SKUs" "INFO"
        
        return $licenses
    }
    catch {
        Write-ModuleLog "Error retrieving licenses: $($_.Exception.Message)" "ERROR"
        throw
    }
}

function Add-UserLicense {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$UserPrincipalName,
        
        [Parameter(Mandatory = $true)]
        [string]$LicenseSku,
        
        [Parameter(Mandatory = $false)]
        [string[]]$DisabledPlans = @()
    )
    
    try {
        if (-not $script:MSGraphConnected) {
            $connectResult = Connect-MSGraphForLicensing
            if (-not $connectResult.Success) {
                throw "Failed to connect to Microsoft Graph"
            }
        }
        
        Write-ModuleLog "Assigning license $LicenseSku to user $UserPrincipalName" "INFO"
        
        # Get the user
        $user = Get-MgUser -UserId $UserPrincipalName -ErrorAction Stop
        
        # Get the SKU ID
        $sku = Get-MgSubscribedSku -All | Where-Object { $_.SkuPartNumber -eq $LicenseSku }
        if (-not $sku) {
            throw "License SKU '$LicenseSku' not found"
        }
        
        # Check if license is available
        $available = $sku.PrepaidUnits.Enabled - $sku.ConsumedUnits
        if ($available -le 0) {
            throw "No available licenses for SKU '$LicenseSku'"
        }
        
        # Build license assignment
        $addLicenses = @{
            SkuId = $sku.SkuId
        }
        
        if ($DisabledPlans.Count -gt 0) {
            $addLicenses.DisabledPlans = $DisabledPlans
        }
        
        # Assign the license
        Set-MgUserLicense -UserId $user.Id -AddLicenses @($addLicenses) -RemoveLicenses @() -ErrorAction Stop
        
        Write-ModuleLog "License assigned successfully" "INFO"
        
        return @{
            Success = $true
            Message = "License '$LicenseSku' assigned to user '$UserPrincipalName'"
            User = $user.UserPrincipalName
            License = $LicenseSku
        }
    }
    catch {
        Write-ModuleLog "Error assigning license: $($_.Exception.Message)" "ERROR"
        return @{
            Success = $false
            Message = "Error assigning license: $($_.Exception.Message)"
            User = $UserPrincipalName
            License = $LicenseSku
        }
    }
}

function Remove-UserLicense {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$UserPrincipalName,
        
        [Parameter(Mandatory = $true)]
        [string]$LicenseSku
    )
    
    try {
        if (-not $script:MSGraphConnected) {
            $connectResult = Connect-MSGraphForLicensing
            if (-not $connectResult.Success) {
                throw "Failed to connect to Microsoft Graph"
            }
        }
        
        Write-ModuleLog "Removing license $LicenseSku from user $UserPrincipalName" "INFO"
        
        # Get the user
        $user = Get-MgUser -UserId $UserPrincipalName -ErrorAction Stop
        
        # Get the SKU ID
        $sku = Get-MgSubscribedSku -All | Where-Object { $_.SkuPartNumber -eq $LicenseSku }
        if (-not $sku) {
            throw "License SKU '$LicenseSku' not found"
        }
        
        # Remove the license
        Set-MgUserLicense -UserId $user.Id -AddLicenses @() -RemoveLicenses @($sku.SkuId) -ErrorAction Stop
        
        Write-ModuleLog "License removed successfully" "INFO"
        
        return @{
            Success = $true
            Message = "License '$LicenseSku' removed from user '$UserPrincipalName'"
        }
    }
    catch {
        Write-ModuleLog "Error removing license: $($_.Exception.Message)" "ERROR"
        return @{
            Success = $false
            Message = "Error removing license: $($_.Exception.Message)"
        }
    }
}

function Get-UserLicenses {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$UserPrincipalName
    )
    
    try {
        if (-not $script:MSGraphConnected) {
            $connectResult = Connect-MSGraphForLicensing
            if (-not $connectResult.Success) {
                throw "Failed to connect to Microsoft Graph"
            }
        }
        
        Write-ModuleLog "Getting licenses for user $UserPrincipalName" "INFO"
        
        # Get user with license details
        $user = Get-MgUser -UserId $UserPrincipalName -Property "id,displayName,assignedLicenses,assignedPlans" -ErrorAction Stop
        
        if (-not $user.AssignedLicenses) {
            return @()
        }
        
        # Get SKU details for assigned licenses
        $userLicenses = @()
        $allSkus = Get-MgSubscribedSku -All
        
        foreach ($license in $user.AssignedLicenses) {
            $sku = $allSkus | Where-Object { $_.SkuId -eq $license.SkuId }
            if ($sku) {
                $userLicenses += [PSCustomObject]@{
                    SkuId = $license.SkuId
                    SkuPartNumber = $sku.SkuPartNumber
                    FriendlyName = if ($script:Configuration.LicenseMapping.ContainsKey($sku.SkuPartNumber)) {
                        $script:Configuration.LicenseMapping[$sku.SkuPartNumber]
                    } else {
                        $sku.SkuPartNumber
                    }
                    DisabledPlans = $license.DisabledPlans
                    ServicePlans = $sku.ServicePlans | Where-Object { $_.ServicePlanId -notin $license.DisabledPlans }
                }
            }
        }
        
        return $userLicenses
    }
    catch {
        Write-ModuleLog "Error getting user licenses: $($_.Exception.Message)" "ERROR"
        throw
    }
}

function Add-BulkUserLicenses {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string[]]$UserPrincipalNames,
        
        [Parameter(Mandatory = $true)]
        [string]$LicenseSku,
        
        [Parameter(Mandatory = $false)]
        [string[]]$DisabledPlans = @(),
        
        [Parameter(Mandatory = $false)]
        [int]$BatchSize = 20
    )
    
    try {
        Write-ModuleLog "Starting bulk license assignment for $($UserPrincipalNames.Count) users" "INFO"
        
        $results = @{
            Total = $UserPrincipalNames.Count
            Success = 0
            Failed = 0
            Details = @()
        }
        
        # Process in batches
        for ($i = 0; $i -lt $UserPrincipalNames.Count; $i += $BatchSize) {
            $batch = $UserPrincipalNames[$i..([Math]::Min($i + $BatchSize - 1, $UserPrincipalNames.Count - 1))]
            
            foreach ($upn in $batch) {
                $result = Add-UserLicense -UserPrincipalName $upn -LicenseSku $LicenseSku -DisabledPlans $DisabledPlans
                
                if ($result.Success) {
                    $results.Success++
                } else {
                    $results.Failed++
                }
                
                $results.Details += $result
            }
            
            # Small delay between batches
            if ($i + $BatchSize -lt $UserPrincipalNames.Count) {
                Start-Sleep -Seconds 2
            }
        }
        
        Write-ModuleLog "Bulk license assignment complete. Success: $($results.Success), Failed: $($results.Failed)" "INFO"
        
        return $results
    }
    catch {
        Write-ModuleLog "Error in bulk license assignment: $($_.Exception.Message)" "ERROR"
        throw
    }
}

function Get-LicenseUsageReport {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $false)]
        [switch]$IncludeServicePlans
    )
    
    try {
        if (-not $script:MSGraphConnected) {
            $connectResult = Connect-MSGraphForLicensing
            if (-not $connectResult.Success) {
                throw "Failed to connect to Microsoft Graph"
            }
        }
        
        Write-ModuleLog "Generating license usage report" "INFO"
        
        $report = @()
        $allSkus = Get-MgSubscribedSku -All
        
        foreach ($sku in $allSkus) {
            $reportEntry = [PSCustomObject]@{
                SkuPartNumber = $sku.SkuPartNumber
                FriendlyName = if ($script:Configuration.LicenseMapping.ContainsKey($sku.SkuPartNumber)) {
                    $script:Configuration.LicenseMapping[$sku.SkuPartNumber]
                } else {
                    $sku.SkuPartNumber
                }
                Total = $sku.PrepaidUnits.Enabled
                Consumed = $sku.ConsumedUnits
                Available = $sku.PrepaidUnits.Enabled - $sku.ConsumedUnits
                PercentageUsed = if ($sku.PrepaidUnits.Enabled -gt 0) {
                    [Math]::Round(($sku.ConsumedUnits / $sku.PrepaidUnits.Enabled) * 100, 2)
                } else { 0 }
            }
            
            if ($IncludeServicePlans) {
                $reportEntry | Add-Member -NotePropertyName "ServicePlans" -NotePropertyValue $sku.ServicePlans
            }
            
            $report += $reportEntry
        }
        
        return $report
    }
    catch {
        Write-ModuleLog "Error generating license report: $($_.Exception.Message)" "ERROR"
        throw
    }
}

# Export module functions
Export-ModuleMember -Function @(
    'Initialize-LicenseModule',
    'Connect-MSGraphForLicensing',
    'Get-AvailableLicenses',
    'Add-UserLicense',
    'Remove-UserLicense',
    'Get-UserLicenses',
    'Add-BulkUserLicenses',
    'Get-LicenseUsageReport'
) 
# SIG # Begin signature block
# MIIcCAYJKoZIhvcNAQcCoIIb+TCCG/UCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCl2XmILVz0fpVu
# zB1Q1lvQ39sBEuDKclq0q3PnBQOh4KCCFk4wggMQMIIB+KADAgECAhB3jzsyX9Cg
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
# BgkqhkiG9w0BCQQxIgQg1JpYrGgeiYWMK888BpN5dtUtLmGFFBtYksYP6VyADm8w
# DQYJKoZIhvcNAQEBBQAEggEAm8rJDc9IOsRcL/L3cBL0AvIEj1n1GjNQL6ZvYFgv
# G1pzzCZMbtnUdYOTwABuiwf2O/AIyIZbPPMR4eTYnami2mty/5LFoJZClQKa93hg
# PSLBIWN/mW7bh9eMs/R+rNYbwrHzRVY1oTcDAKiBbcYX+uRjZJtuR97ImSk/17c4
# Miju/JPidtr61FUodOnzQl5jWBzKdCKRqLf9MM1lAMAOOhVVpxaREOj0VRCiRK5K
# bo9NSof/nF+NSlXsj0S+jTuV+MUHxw7kPpd5cFyNqYk2ODGgojC+eN4JsQr99rQV
# ZypJr5eYKKT1pqtWwoRaAqRA5ane10A/DOdg4XA/oYzwfaGCAyYwggMiBgkqhkiG
# 9w0BCQYxggMTMIIDDwIBATB9MGkxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdp
# Q2VydCwgSW5jLjFBMD8GA1UEAxM4RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3Rh
# bXBpbmcgUlNBNDA5NiBTSEEyNTYgMjAyNSBDQTECEAqA7xhLjfEFgtHEdqeVdGgw
# DQYJYIZIAWUDBAIBBQCgaTAYBgkqhkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqG
# SIb3DQEJBTEPFw0yNTA3MTExOTMxNDdaMC8GCSqGSIb3DQEJBDEiBCBALAsrCS3d
# duROuNhy5RbzjAySXwkXPshMbOeqQlK0cTANBgkqhkiG9w0BAQEFAASCAgCKbuat
# 0IBs1R5BYhO5AppQa7t6FqRscRq+DCoMpEYDz9RYpi20+G1Tzgh2jYmYKkHdRzYw
# JD7GO2+TKA1uigW7rQc0T3tDUF6eZDC1rpM2vL2Kk7HzcbZtpyAhdKKtsi89f990
# Hv+K+WHUXGpQ7eVdsX5AH4WckaticyJIARPU46Sxsy5cPtTJa/Ms8ZQfbB/GTS3U
# XlOrlrMk0kHZTyHQ35O+NiDmS7pQSrFSbfG1LCIr8V85OU0Kwj8aojSxlsrTKi6S
# XL7JRylv3LPe6lxDm0j/BzkjSrfzX9mQ3UrGGx9dcvwk2JWP9o8uvUhgn/jj6sur
# Na4xOYtEZwz/4EF7zNXegUdSTJmwkGvMU4jjT/0nu59VY8CqcmIYi2jMH5DZxv6/
# 7am7TluuFi/y3u7OWxGXsZfEi8unKgwaekmAfj8JirETt/Yyiw8dgJ1UbbLtJRHO
# hx+rrW5NZ8CmqS7JsVCC6laR5cm9cDxxLnhZF6cSf5dPQvtmpb2NeUgNM3X3MOR7
# vEOs4gNftW4dXUKf+dBZ7oNUGi/CxuL/TkuZeIweZTR5F9jCMkr90r6QrX4/TAUR
# xIkrk+WbQ7wSr0oMSN+mAUuS/6KI2hpyfy6VhhJX46kvurn6aux6BerLGAnwAOHO
# NuohnSMHC8FpieTyyKL5P/klNMQTap6nvYG1AQ==
# SIG # End signature block
