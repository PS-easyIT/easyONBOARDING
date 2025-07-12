# ValidationModule.psm1
# Module for input validation and data verification

# Module-level variables
$script:CustomRules = @{}
$script:LogFunction = $null

function Initialize-ValidationModule {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $false)]
        [hashtable]$CustomRules = @{},
        
        [Parameter(Mandatory = $false)]
        [scriptblock]$LogFunction
    )
    
    $script:CustomRules = $CustomRules
    $script:LogFunction = $LogFunction
    
    Write-ModuleLog "ValidationModule initialized with $($CustomRules.Count) custom rules" "INFO"
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

function Test-UserInput {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [PSCustomObject]$UserData,
        
        [Parameter(Mandatory = $false)]
        [switch]$StrictMode
    )
    
    $validationErrors = @()
    
    try {
        # Required field validation
        if ([string]::IsNullOrWhiteSpace($UserData.FirstName)) {
            $validationErrors += "First name is required"
        } elseif ($UserData.FirstName.Length -lt 2) {
            $validationErrors += "First name must be at least 2 characters"
        }
        
        if ([string]::IsNullOrWhiteSpace($UserData.LastName)) {
            $validationErrors += "Last name is required"
        } elseif ($UserData.LastName.Length -lt 2) {
            $validationErrors += "Last name must be at least 2 characters"
        }
        
        # Name format validation
        if ($UserData.FirstName -match "[^a-zA-ZäöüÄÖÜß\s\-']") {
            $validationErrors += "First name contains invalid characters"
        }
        
        if ($UserData.LastName -match "[^a-zA-ZäöüÄÖÜß\s\-']") {
            $validationErrors += "Last name contains invalid characters"
        }
        
        # Email validation
        if (-not [string]::IsNullOrWhiteSpace($UserData.EmailAddress)) {
            if ($UserData.EmailAddress -notmatch '^[a-zA-Z0-9._%+-]+(@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})?$') {
                $validationErrors += "Invalid email format"
            }
        }
        
        # Phone number validation
        if (-not [string]::IsNullOrWhiteSpace($UserData.PhoneNumber)) {
            if ($UserData.PhoneNumber -notmatch '^[\d\s\-\+\(\)]+$') {
                $validationErrors += "Phone number contains invalid characters"
            }
        }
        
        if (-not [string]::IsNullOrWhiteSpace($UserData.MobileNumber)) {
            if ($UserData.MobileNumber -notmatch '^[\d\s\-\+\(\)]+$') {
                $validationErrors += "Mobile number contains invalid characters"
            }
        }
        
        # Department validation
        if ($StrictMode -and [string]::IsNullOrWhiteSpace($UserData.DepartmentField)) {
            $validationErrors += "Department is required in strict mode"
        }
        
        # Position validation
        if ($StrictMode -and [string]::IsNullOrWhiteSpace($UserData.Position)) {
            $validationErrors += "Position is required in strict mode"
        }
        
        # Termination date validation
        if (-not [string]::IsNullOrWhiteSpace($UserData.Ablaufdatum)) {
            try {
                $termDate = [DateTime]::Parse($UserData.Ablaufdatum)
                if ($termDate -lt (Get-Date)) {
                    $validationErrors += "Termination date cannot be in the past"
                }
            } catch {
                $validationErrors += "Invalid termination date format"
            }
        }
        
        # Custom rule validation
        foreach ($rule in $script:CustomRules.GetEnumerator()) {
            $ruleName = $rule.Key
            $ruleValue = $rule.Value
            
            if ($ruleValue -match '^Required:(.+)$') {
                $fieldName = $matches[1]
                if ($UserData.PSObject.Properties.Name -contains $fieldName) {
                    if ([string]::IsNullOrWhiteSpace($UserData.$fieldName)) {
                        $validationErrors += "$fieldName is required by custom rule"
                    }
                }
            }
        }
        
        # Return validation result
        if ($validationErrors.Count -eq 0) {
            Write-ModuleLog "User input validation passed" "INFO"
            return @{
                IsValid = $true
                Errors = @()
                Message = "All validations passed"
            }
        } else {
            Write-ModuleLog "User input validation failed with $($validationErrors.Count) errors" "WARN"
            return @{
                IsValid = $false
                Errors = $validationErrors
                Message = "Validation failed: $($validationErrors -join '; ')"
            }
        }
    }
    catch {
        Write-ModuleLog "Error during validation: $($_.Exception.Message)" "ERROR"
        return @{
            IsValid = $false
            Errors = @("Validation error: $($_.Exception.Message)")
            Message = "Validation error occurred"
        }
    }
}

function Test-EmailAddress {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$EmailAddress,
        
        [Parameter(Mandatory = $false)]
        [switch]$CheckMX
    )
    
    try {
        # Basic format validation
        if ($EmailAddress -notmatch '^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$') {
            return @{
                IsValid = $false
                Message = "Invalid email format"
            }
        }
        
        # Check for common typos
        $commonTypos = @{
            'gmial.com' = 'gmail.com'
            'gmai.com' = 'gmail.com'
            'outlok.com' = 'outlook.com'
            'yahooo.com' = 'yahoo.com'
        }
        
        $domain = $EmailAddress.Split('@')[1]
        if ($commonTypos.ContainsKey($domain)) {
            return @{
                IsValid = $false
                Message = "Possible typo in domain. Did you mean $($commonTypos[$domain])?"
                SuggestedDomain = $commonTypos[$domain]
            }
        }
        
        # MX record check if requested
        if ($CheckMX) {
            try {
                $mx = Resolve-DnsName -Name $domain -Type MX -ErrorAction Stop
                if (-not $mx) {
                    return @{
                        IsValid = $false
                        Message = "No MX records found for domain $domain"
                    }
                }
            } catch {
                return @{
                    IsValid = $false
                    Message = "Could not verify domain $domain"
                }
            }
        }
        
        return @{
            IsValid = $true
            Message = "Email address is valid"
        }
    }
    catch {
        Write-ModuleLog "Error validating email: $($_.Exception.Message)" "ERROR"
        return @{
            IsValid = $false
            Message = "Error validating email address"
        }
    }
}

function Test-PasswordComplexity {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$Password,
        
        [Parameter(Mandatory = $false)]
        [int]$MinLength = 8,
        
        [Parameter(Mandatory = $false)]
        [switch]$RequireUppercase,
        
        [Parameter(Mandatory = $false)]
        [switch]$RequireLowercase,
        
        [Parameter(Mandatory = $false)]
        [switch]$RequireNumbers,
        
        [Parameter(Mandatory = $false)]
        [switch]$RequireSpecialChars
    )
    
    $issues = @()
    
    # Length check
    if ($Password.Length -lt $MinLength) {
        $issues += "Password must be at least $MinLength characters long"
    }
    
    # Uppercase check
    if ($RequireUppercase -and $Password -notmatch '[A-Z]') {
        $issues += "Password must contain at least one uppercase letter"
    }
    
    # Lowercase check
    if ($RequireLowercase -and $Password -notmatch '[a-z]') {
        $issues += "Password must contain at least one lowercase letter"
    }
    
    # Number check
    if ($RequireNumbers -and $Password -notmatch '\d') {
        $issues += "Password must contain at least one number"
    }
    
    # Special character check
    if ($RequireSpecialChars -and $Password -notmatch '[^a-zA-Z0-9]') {
        $issues += "Password must contain at least one special character"
    }
    
    # Common password check
    $commonPasswords = @('password', '123456', 'qwerty', 'abc123', 'password123', 'admin', 'letmein')
    if ($commonPasswords -contains $Password.ToLower()) {
        $issues += "Password is too common and easily guessable"
    }
    
    if ($issues.Count -eq 0) {
        return @{
            IsValid = $true
            Message = "Password meets complexity requirements"
            Score = Calculate-PasswordStrength -Password $Password
        }
    } else {
        return @{
            IsValid = $false
            Message = "Password does not meet complexity requirements"
            Issues = $issues
        }
    }
}

function Calculate-PasswordStrength {
    param([string]$Password)
    
    $score = 0
    
    # Length bonus
    $score += [Math]::Min($Password.Length * 4, 40)
    
    # Character variety bonus
    if ($Password -match '[a-z]') { $score += 10 }
    if ($Password -match '[A-Z]') { $score += 10 }
    if ($Password -match '\d') { $score += 10 }
    if ($Password -match '[^a-zA-Z0-9]') { $score += 20 }
    
    # Pattern penalty
    if ($Password -match '(.)\1{2,}') { $score -= 20 }  # Repeated characters
    if ($Password -match '(012|123|234|345|456|567|678|789|890|abc|bcd|cde|def)') { $score -= 20 }  # Sequential
    
    return [Math]::Max(0, [Math]::Min(100, $score))
}

function Test-ADGroupExists {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$GroupName
    )
    
    try {
        $group = Get-ADGroup -Identity $GroupName -ErrorAction Stop
        return @{
            Exists = $true
            Group = $group
            Message = "Group exists"
        }
    }
    catch {
        return @{
            Exists = $false
            Group = $null
            Message = "Group does not exist or cannot be accessed"
        }
    }
}

function Test-ADUserExists {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$Identity,
        
        [Parameter(Mandatory = $false)]
        [ValidateSet("SamAccountName", "UserPrincipalName", "EmailAddress")]
        [string]$IdentityType = "SamAccountName"
    )
    
    try {
        $filter = switch ($IdentityType) {
            "SamAccountName" { "SamAccountName -eq '$Identity'" }
            "UserPrincipalName" { "UserPrincipalName -eq '$Identity'" }
            "EmailAddress" { "EmailAddress -eq '$Identity'" }
        }
        
        $user = Get-ADUser -Filter $filter -ErrorAction Stop
        
        if ($user) {
            return @{
                Exists = $true
                User = $user
                Message = "User exists"
            }
        } else {
            return @{
                Exists = $false
                User = $null
                Message = "User does not exist"
            }
        }
    }
    catch {
        Write-ModuleLog "Error checking user existence: $($_.Exception.Message)" "ERROR"
        return @{
            Exists = $false
            User = $null
            Message = "Error checking user existence"
        }
    }
}

# Export module functions
Export-ModuleMember -Function @(
    'Initialize-ValidationModule',
    'Test-UserInput',
    'Test-EmailAddress',
    'Test-PasswordComplexity',
    'Test-ADGroupExists',
    'Test-ADUserExists'
) 
# SIG # Begin signature block
# MIIcCAYJKoZIhvcNAQcCoIIb+TCCG/UCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCwoLoAP0QMWKBb
# RtUtReQQF1K7dWDWKrjNgGLj12E+XqCCFk4wggMQMIIB+KADAgECAhB3jzsyX9Cg
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
# BgkqhkiG9w0BCQQxIgQgwGHB0EYEYeFAcAr1DJftYPSQHiVdhMeKMuw6dipHU1ow
# DQYJKoZIhvcNAQEBBQAEggEAYDz/0MflHl0R60oWMg9UAljFFA1+qLkL52yB0aLt
# kIyLXvhYkwYTY5QSB3u/QIE09nER6l83/4NUSZUeOXC62dBj/vEiH6jRscLZbDJB
# Z6vjjx5/6z5Pt9t2cQIwQ3YKMDXoJaosf4Zp1MAhzEDtj/0+lr9y4DUAfr20ktWx
# ElGBnyCHvb+BIgNER1GyXWeprL73go3PqSDkodJuWmq2ykVzHu+/aFE8QujEhVYh
# RgKoKmFMnajLaJV0sWHnnWRmLLL2sCMVJD9LXd/NtCa9htMBYxCLOXTLB/aXJYMg
# sUgJv47i1AoKnX7BU/ttB+yRsNxE3IpOR6qulCaBt5y22KGCAyYwggMiBgkqhkiG
# 9w0BCQYxggMTMIIDDwIBATB9MGkxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdp
# Q2VydCwgSW5jLjFBMD8GA1UEAxM4RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3Rh
# bXBpbmcgUlNBNDA5NiBTSEEyNTYgMjAyNSBDQTECEAqA7xhLjfEFgtHEdqeVdGgw
# DQYJYIZIAWUDBAIBBQCgaTAYBgkqhkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqG
# SIb3DQEJBTEPFw0yNTA3MTExOTMyNDBaMC8GCSqGSIb3DQEJBDEiBCDINuc9/6Ld
# rdIVfGQh+1AKOpWUOenF5G9dmwcWA9tm4DANBgkqhkiG9w0BAQEFAASCAgA8Og/2
# ta7hNBulE0yywYatSMzq+F+lWMsEff8jCoEDJ4xACiFIfUhtUlFQlEmfsAn3satg
# jaVGOEyldErCIG0/l52/JF9OZrsNRvC+cklEvZqQ0liXd7M6zlTosOOH7pv5AwZo
# lJryo2FL5czMx0XjAIXdU53aVnd3YkbY+ZTUP/D5cr1W39ZYEKiahi/01qngFgX7
# xKsgglKUk0zpm1kPOfriXoz5GypetY/k+VzF01EiChq5/ntGOFXAuagbMRUFX7Fi
# 2ln8xziS74O5HwNc9QN/2mJSQah+yQ/nwTaDzw55rdSOA8ORSiCzm34RwWwNonnh
# CiFY5MlI7e8ehfWNYz8cbV4iPEpPD4CJX1rRlpeb+08pQxXFr9FXHJ0qTHhQ5C+H
# H1cU8aKTo852dUQLWIC/Nkaz9DxqiyaaW1jCmerjt+9h02aMTT5J7zfMiOie83Qj
# RqhP5X859IFIHqvv5UCW0ldJoqVL+GQHihUfRdG3kGbDrMwcAGDBtyT+t75DOHx8
# wvoEnWOpGkvdM1tVB0/6LB0wx5qOwQgwiFCLTUC6vLjsSDoJwHp8UZdHDhV+aQYB
# artwW1ZdmyOuDcueb8ffZUUtkQw7oRWqiiI2nIpiXXbIAthoWIZUKOI0Fu/9P7z0
# PIKjijCpH5vZY3OVTqk/DJm3Zz3OmoQxTYQsgg==
# SIG # End signature block
