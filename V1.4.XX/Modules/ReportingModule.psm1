# ReportingModule.psm1
# Module for comprehensive reporting functionality

# Module-level variables
$script:ReportPath = ""
$script:LogFunction = $null
$script:Config = @{}

function Initialize-ReportingModule {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$ReportPath,
        
        [Parameter(Mandatory = $false)]
        [scriptblock]$LogFunction,
        
        [Parameter(Mandatory = $false)]
        [hashtable]$Configuration = @{}
    )
    
    $script:ReportPath = $ReportPath
    $script:LogFunction = $LogFunction
    $script:Config = $Configuration
    
    # Ensure report directory exists
    if (-not (Test-Path $ReportPath)) {
        New-Item -ItemType Directory -Path $ReportPath -Force | Out-Null
    }
    
    Write-ModuleLog "ReportingModule initialized with path: $ReportPath" "INFO"
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

function New-OnboardingSummaryReport {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [DateTime]$StartDate,
        
        [Parameter(Mandatory = $true)]
        [DateTime]$EndDate,
        
        [Parameter(Mandatory = $false)]
        [ValidateSet("HTML", "PDF", "Excel", "CSV")]
        [string]$OutputFormat = "HTML"
    )
    
    try {
        Write-ModuleLog "Generating onboarding summary report from $StartDate to $EndDate" "INFO"
        
        # Get onboarding logs
        $logPath = Split-Path $script:ReportPath -Parent
        $logFile = Join-Path $logPath "Logs\easyOnboarding.log"
        
        $onboardingData = @()
        
        if (Test-Path $logFile) {
            $logs = Get-Content $logFile | Where-Object { 
                $_ -match "ONBOARDING PERFORMED BY:" -and 
                $_ -match "(\d{4}-\d{2}-\d{2})"
            }
            
            foreach ($log in $logs) {
                if ($log -match "(\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}).*ONBOARDING PERFORMED BY: ([^,]+), SamAccountName: ([^,]+), Display Name: '([^']+)', UPN: '([^']+)', Location: '([^']+)', Company: '([^']+)', License: '([^']+)'") {
                    $logDate = [DateTime]::Parse($matches[1])
                    if ($logDate -ge $StartDate -and $logDate -le $EndDate) {
                        $onboardingData += [PSCustomObject]@{
                            Date = $logDate
                            PerformedBy = $matches[2]
                            SamAccountName = $matches[3]
                            DisplayName = $matches[4]
                            UPN = $matches[5]
                            Location = $matches[6]
                            Company = $matches[7]
                            License = $matches[8]
                        }
                    }
                }
            }
        }
        
        # Generate report based on format
        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $reportFileName = "OnboardingSummary_${timestamp}.$($OutputFormat.ToLower())"
        $reportPath = Join-Path $script:ReportPath $reportFileName
        
        switch ($OutputFormat) {
            "HTML" {
                $html = @"
<!DOCTYPE html>
<html>
<head>
    <title>Onboarding Summary Report</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; }
        h1 { color: #333; }
        table { border-collapse: collapse; width: 100%; margin-top: 20px; }
        th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }
        th { background-color: #4CAF50; color: white; }
        tr:nth-child(even) { background-color: #f2f2f2; }
        .summary { background-color: #e7f3fe; border-left: 6px solid #2196F3; padding: 10px; margin: 20px 0; }
    </style>
</head>
<body>
    <h1>Onboarding Summary Report</h1>
    <p>Period: $($StartDate.ToString('yyyy-MM-dd')) to $($EndDate.ToString('yyyy-MM-dd'))</p>
    <div class="summary">
        <h3>Summary</h3>
        <p>Total Users Onboarded: $($onboardingData.Count)</p>
        <p>Report Generated: $(Get-Date)</p>
    </div>
    <table>
        <thead>
            <tr>
                <th>Date</th>
                <th>Performed By</th>
                <th>Username</th>
                <th>Display Name</th>
                <th>UPN</th>
                <th>Location</th>
                <th>Company</th>
                <th>License</th>
            </tr>
        </thead>
        <tbody>
"@
                foreach ($user in $onboardingData | Sort-Object Date -Descending) {
                    $html += @"
            <tr>
                <td>$($user.Date.ToString('yyyy-MM-dd HH:mm'))</td>
                <td>$($user.PerformedBy)</td>
                <td>$($user.SamAccountName)</td>
                <td>$($user.DisplayName)</td>
                <td>$($user.UPN)</td>
                <td>$($user.Location)</td>
                <td>$($user.Company)</td>
                <td>$($user.License)</td>
            </tr>
"@
                }
                
                $html += @"
        </tbody>
    </table>
</body>
</html>
"@
                Set-Content -Path $reportPath -Value $html -Encoding UTF8
            }
            
            "CSV" {
                $onboardingData | Export-Csv -Path $reportPath -NoTypeInformation -Encoding UTF8
            }
            
            "Excel" {
                if (Get-Module -ListAvailable -Name ImportExcel) {
                    $onboardingData | Export-Excel -Path $reportPath -AutoSize -AutoFilter -TableName "OnboardingData"
                } else {
                    # Fallback to CSV if Excel module not available
                    $reportPath = $reportPath -replace '\.xlsx$', '.csv'
                    $onboardingData | Export-Csv -Path $reportPath -NoTypeInformation -Encoding UTF8
                    Write-ModuleLog "Excel module not available, exported as CSV instead" "WARN"
                }
            }
            
            "PDF" {
                # For PDF, we need to create HTML first then convert
                # This requires wkhtmltopdf or similar tool
                Write-ModuleLog "PDF generation requires external tool configuration" "WARN"
                return @{
                    Success = $false
                    Message = "PDF generation not yet implemented"
                    Path = $null
                }
            }
        }
        
        Write-ModuleLog "Report generated successfully: $reportPath" "INFO"
        
        return @{
            Success = $true
            Path = $reportPath
            RecordCount = $onboardingData.Count
            Data = $onboardingData
        }
    }
    catch {
        Write-ModuleLog "Error generating report: $($_.Exception.Message)" "ERROR"
        return @{
            Success = $false
            Message = $_.Exception.Message
            Path = $null
        }
    }
}

function New-UserReport {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$Identity,
        
        [Parameter(Mandatory = $false)]
        [ValidateSet("Summary", "Detailed", "Audit")]
        [string]$ReportType = "Summary",
        
        [Parameter(Mandatory = $false)]
        [string]$OutputPath
    )
    
    try {
        Write-ModuleLog "Generating $ReportType report for user: $Identity" "INFO"
        
        # Get user details from AD
        $user = Get-ADUser -Identity $Identity -Properties * -ErrorAction Stop
        
        if (-not $OutputPath) {
            $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
            $OutputPath = Join-Path $script:ReportPath "UserReport_${Identity}_${timestamp}.html"
        }
        
        # Generate HTML report
        $html = @"
<!DOCTYPE html>
<html>
<head>
    <title>User Report - $($user.DisplayName)</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; }
        h1, h2 { color: #333; }
        .section { margin: 20px 0; padding: 15px; border: 1px solid #ddd; border-radius: 5px; }
        .property { margin: 5px 0; }
        .label { font-weight: bold; display: inline-block; width: 200px; }
        .value { color: #555; }
        table { border-collapse: collapse; width: 100%; margin-top: 10px; }
        th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }
        th { background-color: #4CAF50; color: white; }
    </style>
</head>
<body>
    <h1>User Report: $($user.DisplayName)</h1>
    <p>Generated: $(Get-Date)</p>
    
    <div class="section">
        <h2>Basic Information</h2>
        <div class="property"><span class="label">Username:</span> <span class="value">$($user.SamAccountName)</span></div>
        <div class="property"><span class="label">Display Name:</span> <span class="value">$($user.DisplayName)</span></div>
        <div class="property"><span class="label">Email:</span> <span class="value">$($user.EmailAddress)</span></div>
        <div class="property"><span class="label">UPN:</span> <span class="value">$($user.UserPrincipalName)</span></div>
        <div class="property"><span class="label">Department:</span> <span class="value">$($user.Department)</span></div>
        <div class="property"><span class="label">Title:</span> <span class="value">$($user.Title)</span></div>
        <div class="property"><span class="label">Office:</span> <span class="value">$($user.Office)</span></div>
        <div class="property"><span class="label">Manager:</span> <span class="value">$($user.Manager)</span></div>
    </div>
"@
        
        if ($ReportType -in @("Detailed", "Audit")) {
            $html += @"
    <div class="section">
        <h2>Account Status</h2>
        <div class="property"><span class="label">Enabled:</span> <span class="value">$($user.Enabled)</span></div>
        <div class="property"><span class="label">Created:</span> <span class="value">$($user.Created)</span></div>
        <div class="property"><span class="label">Last Logon:</span> <span class="value">$($user.LastLogonDate)</span></div>
        <div class="property"><span class="label">Password Last Set:</span> <span class="value">$($user.PasswordLastSet)</span></div>
        <div class="property"><span class="label">Account Expires:</span> <span class="value">$($user.AccountExpirationDate)</span></div>
    </div>
    
    <div class="section">
        <h2>Group Memberships</h2>
        <table>
            <thead>
                <tr><th>Group Name</th><th>Type</th></tr>
            </thead>
            <tbody>
"@
            $groups = Get-ADPrincipalGroupMembership -Identity $Identity | Sort-Object Name
            foreach ($group in $groups) {
                $html += "<tr><td>$($group.Name)</td><td>$($group.GroupCategory)</td></tr>"
            }
            
            $html += @"
            </tbody>
        </table>
    </div>
"@
        }
        
        $html += @"
</body>
</html>
"@
        
        Set-Content -Path $OutputPath -Value $html -Encoding UTF8
        Write-ModuleLog "User report generated: $OutputPath" "INFO"
        
        return $OutputPath
    }
    catch {
        Write-ModuleLog "Error generating user report: $($_.Exception.Message)" "ERROR"
        throw
    }
}

function Get-InactiveUsersReport {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $false)]
        [int]$DaysInactive = 90,
        
        [Parameter(Mandatory = $false)]
        [ValidateSet("HTML", "CSV", "Excel")]
        [string]$OutputFormat = "HTML"
    )
    
    try {
        Write-ModuleLog "Generating inactive users report (>$DaysInactive days)" "INFO"
        
        $cutoffDate = (Get-Date).AddDays(-$DaysInactive)
        
        # Get inactive users
        $inactiveUsers = Get-ADUser -Filter "(LastLogonDate -lt '$cutoffDate' -or -not(LastLogonDate -like '*')) -and Enabled -eq 'True'" `
            -Properties LastLogonDate, Created, PasswordLastSet, Department, Title, Manager |
        Select-Object SamAccountName, DisplayName, EmailAddress, LastLogonDate, 
                      Created, PasswordLastSet, Department, Title, @{
                          Name = "DaysInactive"
                          Expression = {
                              if ($_.LastLogonDate) {
                                  (New-TimeSpan -Start $_.LastLogonDate -End (Get-Date)).Days
                              } else {
                                  "Never Logged In"
                              }
                          }
                      }
        
        # Generate report
        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $reportPath = Join-Path $script:ReportPath "InactiveUsers_${timestamp}.$($OutputFormat.ToLower())"
        
        switch ($OutputFormat) {
            "CSV" {
                $inactiveUsers | Export-Csv -Path $reportPath -NoTypeInformation
            }
            "Excel" {
                if (Get-Module -ListAvailable -Name ImportExcel) {
                    $inactiveUsers | Export-Excel -Path $reportPath -AutoSize -AutoFilter
                } else {
                    $reportPath = $reportPath -replace '\.xlsx$', '.csv'
                    $inactiveUsers | Export-Csv -Path $reportPath -NoTypeInformation
                }
            }
            "HTML" {
                # Generate HTML report
                $html = ConvertTo-Html -InputObject $inactiveUsers -Title "Inactive Users Report" `
                    -PreContent "<h1>Inactive Users Report</h1><p>Users inactive for more than $DaysInactive days</p>"
                Set-Content -Path $reportPath -Value $html
            }
        }
        
        Write-ModuleLog "Inactive users report generated: $reportPath" "INFO"
        return @{
            Path = $reportPath
            Count = $inactiveUsers.Count
            Data = $inactiveUsers
        }
    }
    catch {
        Write-ModuleLog "Error generating inactive users report: $($_.Exception.Message)" "ERROR"
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
        Write-ModuleLog "Generating license usage report" "INFO"
        
        # This is a placeholder - actual implementation would query Office 365/Azure AD
        # For now, we'll generate a sample report based on AD group memberships
        
        $licenseGroups = @()
        
        if ($script:Config.Contains("LicensesGroups")) {
            foreach ($key in $script:Config["LicensesGroups"].Keys) {
                $groupName = $script:Config["LicensesGroups"][$key]
                try {
                    $members = Get-ADGroupMember -Identity $groupName -ErrorAction Stop
                    $licenseGroups += [PSCustomObject]@{
                        LicenseType = $key
                        GroupName = $groupName
                        AssignedUsers = $members.Count
                        Users = $members | Select-Object Name, SamAccountName
                    }
                } catch {
                    Write-ModuleLog "Could not get members for group: $groupName" "WARN"
                }
            }
        }
        
        return $licenseGroups
    }
    catch {
        Write-ModuleLog "Error generating license report: $($_.Exception.Message)" "ERROR"
        throw
    }
}

# Export module functions
Export-ModuleMember -Function @(
    'Initialize-ReportingModule',
    'New-OnboardingSummaryReport',
    'New-UserReport',
    'Get-InactiveUsersReport',
    'Get-LicenseUsageReport'
) 
# SIG # Begin signature block
# MIIcCAYJKoZIhvcNAQcCoIIb+TCCG/UCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBNSmAA9B8CUmJk
# L7QVXB9utq6Mql67+Y9Q2eoEuGPRlaCCFk4wggMQMIIB+KADAgECAhB3jzsyX9Cg
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
# BgkqhkiG9w0BCQQxIgQgJ95Ho7VDmHMSqWc3YzuTt2eBWPMmtVBPbA3eNTIfk3ow
# DQYJKoZIhvcNAQEBBQAEggEAMyEW2tvYl2dCuP4I5DNI/3mI32XJWDMaLhRSVNPY
# v3U+oX2on3IK4qTkJ8z4Xzmvtb4xxfK3lEgWWnIH2+Fzik5PNRusQPBU2O8EPd6n
# WvB/8fuK9b7vusY61PGhVVbvHYtcPksOSNEF2oC4Xf9LJSEtraAGMuYCxu5lnnzw
# 07v+9365VUPXjbBZDe2VJYzqHRxv0PxqW/w9vRipDSMI0Yp8w4KIyDn8q5NUqmB+
# 4/3AcXmF2ce2DLWs6/7G1g96c4ZwhshCXvsiQ33e3JkMIqZJkvYmUe5QPEaoDZe6
# QbejT9NHplwaT1Seek3hOMmaQPaZZ69ZKwfLRP1h0aiDzKGCAyYwggMiBgkqhkiG
# 9w0BCQYxggMTMIIDDwIBATB9MGkxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdp
# Q2VydCwgSW5jLjFBMD8GA1UEAxM4RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3Rh
# bXBpbmcgUlNBNDA5NiBTSEEyNTYgMjAyNSBDQTECEAqA7xhLjfEFgtHEdqeVdGgw
# DQYJYIZIAWUDBAIBBQCgaTAYBgkqhkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqG
# SIb3DQEJBTEPFw0yNTA3MTExOTMyMDhaMC8GCSqGSIb3DQEJBDEiBCCmA9nSRIrP
# AQ55oED8uXaLsuHlewnS0pvF/PK7rw5DjDANBgkqhkiG9w0BAQEFAASCAgAVASwa
# rNe86sYfRXLXapGPZuftuZeA2tiqhJ2uoiO7y/zTJsVAFA4O/5UzoGqFnNfmmRoq
# 0zIBChoIzAjHujLcfJaroOwC3UTLiy48IDcQjGL4qvqqAkXj1K33KN0evUkINevz
# zRBNoS4YEcZ+yxdo4JeaVyA48+/H0v7ao5IL733mKWbI2fcuMvRMHWUFTkqQFVpf
# Yqz5dnv+hQXX9Z3V1acRy8zn+TBXieuGiMK0tnj+NgjGuGevufvwcaNl6QLV4B8v
# 5xHELLLV2iijJF3pmfxpoWoMv8WEcyN3iqK9fOcE/TXin8npMgKD0S599GabLipO
# 6J5KQn2bPSQB7BOrHlMSKZEB8OrGEMsb+yI2p4//GuG3eCnL8w+CGrvqw8iXUs4e
# 2Znb+2RBabGJxDT0/dcMvLpNrJ6dU1mc/G7GbKkl5P5yCkBIJsdohWbqXoRp3SQM
# KNhSsIwdB8x6UHpK1O1Ijyf2iP49cMBRfEnrr83O8gAlTTfU+zuzqCIzJSav+nry
# jKfmVEgfzOL4Wjyw1c5sGO+HwKRZ5a+ACPZoqOFk+8rS6EC+hsgs9UTeraCYmumv
# zFUtdN5td9bt8GMrajIJ8bvC9UpnTVZlozw3ssuTYSS7M8UxG8W/elPJk0A3Uv8g
# zzDv/QYboJdR1JrZj6Lv6OcBFzFBFLtKU/EpOg==
# SIG # End signature block
