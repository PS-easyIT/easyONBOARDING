# ExtendedFunctions.psm1 - Extended functionality for easyONBOARDING
# Version: 1.0.0
# Description: Additional functions for complete onboarding process

#region Module Variables
$script:ModuleVersion = "1.0.0"
$script:LogFunction = $null
#endregion

#region Initialization
function Initialize-ExtendedFunctions {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [ScriptBlock]$LogFunction
    )
    
    $script:LogFunction = $LogFunction
    Write-ModuleLog "ExtendedFunctions module initialized (v$script:ModuleVersion)" "INFO"
}

function Write-ModuleLog {
    param(
        [string]$Message,
        [string]$Level = "INFO"
    )
    
    if ($script:LogFunction) {
        & $script:LogFunction $Message $Level
    }
    else {
        Write-Host "[$Level] $Message"
    }
}
#endregion

#region Export Functions

# Export user data to various formats
function Export-UserData {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$SamAccountName,
        
        [Parameter(Mandatory = $true)]
        [ValidateSet('CSV', 'JSON', 'XML', 'HTML')]
        [string]$Format,
        
        [Parameter(Mandatory = $true)]
        [string]$OutputPath,
        
        [Parameter(Mandatory = $false)]
        [switch]$IncludeGroups,
        
        [Parameter(Mandatory = $false)]
        [switch]$IncludeManager,
        
        [Parameter(Mandatory = $false)]
        [switch]$IncludeExtendedProperties
    )
    
    try {
        Write-ModuleLog "Exporting user data for $SamAccountName to $Format format" "INFO"
        
        # Get user data
        $user = Get-ADUser -Identity $SamAccountName -Properties *
        
        if (-not $user) {
            throw "User $SamAccountName not found"
        }
        
        # Build export object
        $exportData = [PSCustomObject]@{
            SamAccountName = $user.SamAccountName
            UserPrincipalName = $user.UserPrincipalName
            DisplayName = $user.DisplayName
            FirstName = $user.GivenName
            LastName = $user.Surname
            Email = $user.EmailAddress
            Department = $user.Department
            Title = $user.Title
            Office = $user.Office
            Phone = $user.telephoneNumber
            Mobile = $user.mobile
            Created = $user.Created
            Modified = $user.Modified
            Enabled = $user.Enabled
        }
        
        # Add groups if requested
        if ($IncludeGroups) {
            $groups = Get-ADPrincipalGroupMembership -Identity $SamAccountName | Select-Object -ExpandProperty Name
            $exportData | Add-Member -MemberType NoteProperty -Name "Groups" -Value ($groups -join ";")
        }
        
        # Add manager if requested
        if ($IncludeManager -and $user.Manager) {
            $manager = Get-ADUser -Identity $user.Manager -Properties DisplayName
            $exportData | Add-Member -MemberType NoteProperty -Name "Manager" -Value $manager.DisplayName
        }
        
        # Add extended properties if requested
        if ($IncludeExtendedProperties) {
            $exportData | Add-Member -MemberType NoteProperty -Name "StreetAddress" -Value $user.StreetAddress
            $exportData | Add-Member -MemberType NoteProperty -Name "City" -Value $user.City
            $exportData | Add-Member -MemberType NoteProperty -Name "PostalCode" -Value $user.PostalCode
            $exportData | Add-Member -MemberType NoteProperty -Name "Country" -Value $user.Country
            $exportData | Add-Member -MemberType NoteProperty -Name "Company" -Value $user.Company
            $exportData | Add-Member -MemberType NoteProperty -Name "EmployeeID" -Value $user.EmployeeID
        }
        
        # Export based on format
        switch ($Format) {
            'CSV' {
                $exportData | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
            }
            'JSON' {
                $exportData | ConvertTo-Json -Depth 10 | Out-File -FilePath $OutputPath -Encoding UTF8
            }
            'XML' {
                $exportData | Export-Clixml -Path $OutputPath
            }
            'HTML' {
                $html = @"
<!DOCTYPE html>
<html>
<head>
    <title>User Export - $($user.DisplayName)</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; }
        table { border-collapse: collapse; width: 100%; }
        th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }
        th { background-color: #4CAF50; color: white; }
        tr:nth-child(even) { background-color: #f2f2f2; }
    </style>
</head>
<body>
    <h1>User Export: $($user.DisplayName)</h1>
    <table>
"@
                foreach ($property in $exportData.PSObject.Properties) {
                    $html += "<tr><th>$($property.Name)</th><td>$($property.Value)</td></tr>`n"
                }
                $html += @"
    </table>
    <p>Exported on: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</p>
</body>
</html>
"@
                $html | Out-File -FilePath $OutputPath -Encoding UTF8
            }
        }
        
        Write-ModuleLog "User data exported successfully to $OutputPath" "INFO"
        return $true
    }
    catch {
        Write-ModuleLog "Error exporting user data: $($_.Exception.Message)" "ERROR"
        throw
    }
}

# Send welcome email to new user
function Send-WelcomeEmail {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [PSCustomObject]$UserData,
        
        [Parameter(Mandatory = $true)]
        [hashtable]$Config,
        
        [Parameter(Mandatory = $false)]
        [string]$TemplatePath,
        
        [Parameter(Mandatory = $false)]
        [string[]]$Attachments,
        
        [Parameter(Mandatory = $false)]
        [switch]$SendCopy
    )
    
    try {
        Write-ModuleLog "Preparing welcome email for $($UserData.DisplayName)" "INFO"
        
        # Get email configuration
        if (-not $Config.Contains("EmailSettings")) {
            throw "EmailSettings section missing in configuration"
        }
        
        $emailConfig = $Config.EmailSettings
        $smtpServer = $emailConfig.SMTPServer
        $smtpPort = if ($emailConfig.SMTPPort) { [int]$emailConfig.SMTPPort } else { 25 }
        $fromAddress = $emailConfig.FromAddress
        $useSSL = $emailConfig.UseSSL -eq "1"
        
        # Build recipient address
        $toAddress = if ($UserData.EmailAddress -contains '@') {
            $UserData.EmailAddress
        } else {
            "$($UserData.EmailAddress)$($UserData.MailSuffix)"
        }
        
        # Load email template or use default
        $emailBody = if ($TemplatePath -and (Test-Path $TemplatePath)) {
            Get-Content -Path $TemplatePath -Raw
        } else {
            @"
<!DOCTYPE html>
<html>
<head>
    <style>
        body { font-family: Arial, sans-serif; line-height: 1.6; color: #333; }
        .container { max-width: 600px; margin: 0 auto; padding: 20px; }
        .header { background-color: #0078D4; color: white; padding: 20px; text-align: center; }
        .content { padding: 20px; background-color: #f4f4f4; }
        .info-box { background-color: white; padding: 15px; margin: 10px 0; border-radius: 5px; }
        .footer { text-align: center; padding: 20px; font-size: 12px; color: #666; }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>Welcome to {{CompanyName}}!</h1>
        </div>
        <div class="content">
            <h2>Hello {{FirstName}} {{LastName}},</h2>
            <p>We are pleased to welcome you to our team! Your account has been successfully created.</p>
            
            <div class="info-box">
                <h3>Your Account Information:</h3>
                <p><strong>Username:</strong> {{SamAccountName}}</p>
                <p><strong>Email:</strong> {{Email}}</p>
                <p><strong>Temporary Password:</strong> {{Password}}</p>
                <p><em>You will be required to change this password on your first login.</em></p>
            </div>
            
            <div class="info-box">
                <h3>Getting Started:</h3>
                <ul>
                    <li>Access your email at: <a href="https://outlook.office365.com">Outlook Web Access</a></li>
                    <li>Company portal: <a href="{{CompanyPortal}}">{{CompanyPortal}}</a></li>
                    <li>IT Support: {{ITSupport}}</li>
                </ul>
            </div>
            
            <p>If you have any questions, please don't hesitate to contact our IT department.</p>
            
            <p>Best regards,<br>IT Team</p>
        </div>
        <div class="footer">
            <p>This is an automated message. Please do not reply to this email.</p>
        </div>
    </div>
</body>
</html>
"@
        }
        
        # Replace placeholders
        $emailBody = $emailBody -replace '{{FirstName}}', $UserData.FirstName
        $emailBody = $emailBody -replace '{{LastName}}', $UserData.LastName
        $emailBody = $emailBody -replace '{{DisplayName}}', $UserData.DisplayName
        $emailBody = $emailBody -replace '{{SamAccountName}}', $UserData.SamAccountName
        $emailBody = $emailBody -replace '{{Email}}', $toAddress
        $emailBody = $emailBody -replace '{{Password}}', $UserData.Password
        $emailBody = $emailBody -replace '{{CompanyName}}', $Config.Company.CompanyNameFirma
        $emailBody = $emailBody -replace '{{CompanyPortal}}', $Config.Company.CompanyDomain
        $emailBody = $emailBody -replace '{{ITSupport}}', $Config.CompanyHelpdesk.CompanyHelpdeskMail
        
        # Create email message
        $mailParams = @{
            To = $toAddress
            From = $fromAddress
            Subject = "Welcome to $($Config.Company.CompanyNameFirma) - Account Information"
            Body = $emailBody
            BodyAsHtml = $true
            SmtpServer = $smtpServer
            Port = $smtpPort
            UseSsl = $useSSL
        }
        
        # Add CC if SendCopy is specified
        if ($SendCopy -and $Config.EmailSettings.CopyAddress) {
            $mailParams.Cc = $Config.EmailSettings.CopyAddress
        }
        
        # Add attachments if specified
        if ($Attachments) {
            $validAttachments = $Attachments | Where-Object { Test-Path $_ }
            if ($validAttachments) {
                $mailParams.Attachments = $validAttachments
            }
        }
        
        # Send email
        Send-MailMessage @mailParams
        
        Write-ModuleLog "Welcome email sent successfully to $toAddress" "INFO"
        return $true
    }
    catch {
        Write-ModuleLog "Error sending welcome email: $($_.Exception.Message)" "ERROR"
        throw
    }
}

# Create Exchange mailbox for user
function New-UserMailbox {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserPrincipalName,
        
        [Parameter(Mandatory = $false)]
        [string]$Database,
        
        [Parameter(Mandatory = $false)]
        [ValidateSet('Regular', 'Room', 'Equipment', 'Shared')]
        [string]$Type = 'Regular',
        
        [Parameter(Mandatory = $false)]
        [string]$Alias,
        
        [Parameter(Mandatory = $false)]
        [switch]$Archive
    )
    
    try {
        Write-ModuleLog "Creating mailbox for $UserPrincipalName" "INFO"
        
        # Check if Exchange management tools are available
        if (-not (Get-Command Enable-Mailbox -ErrorAction SilentlyContinue)) {
            throw "Exchange management tools not available"
        }
        
        # Prepare mailbox parameters
        $mailboxParams = @{
            Identity = $UserPrincipalName
        }
        
        if ($Database) {
            $mailboxParams.Database = $Database
        }
        
        if ($Alias) {
            $mailboxParams.Alias = $Alias
        }
        
        # Create mailbox based on type
        switch ($Type) {
            'Regular' {
                Enable-Mailbox @mailboxParams
            }
            'Room' {
                Enable-Mailbox @mailboxParams -Room
            }
            'Equipment' {
                Enable-Mailbox @mailboxParams -Equipment
            }
            'Shared' {
                Enable-Mailbox @mailboxParams -Shared
            }
        }
        
        # Enable archive if requested
        if ($Archive) {
            Enable-Mailbox -Identity $UserPrincipalName -Archive
        }
        
        Write-ModuleLog "Mailbox created successfully for $UserPrincipalName" "INFO"
        return $true
    }
    catch {
        Write-ModuleLog "Error creating mailbox: $($_.Exception.Message)" "ERROR"
        throw
    }
}

# Export module members
Export-ModuleMember -Function @(
    'Initialize-ExtendedFunctions',
    'Export-UserData',
    'Send-WelcomeEmail',
    'New-UserMailbox'  # Changed from Create-UserMailbox to use approved verb
) 
# SIG # Begin signature block
# MIIcCAYJKoZIhvcNAQcCoIIb+TCCG/UCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAP5MA3mf+Cx9ae
# bnIgtFfjZqTqpj9NsGAf2VDlYSeLb6CCFk4wggMQMIIB+KADAgECAhB3jzsyX9Cg
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
# BgkqhkiG9w0BCQQxIgQgYLW6hY17mWOT/9m+lyxaUcxc7qL+mLeuQmq6P87AnM0w
# DQYJKoZIhvcNAQEBBQAEggEAC94lWYGzmiv7YpijMt2uAHIrVhqP2o67hU/tLe6H
# s97yypChj2arenRf/UPq3+08PvqEBsEktG9YsN/4YwTWvUSc7ez4OwmFhrnL3YWI
# ESbIDrOPDtaFBNPvvaPkj9qGkxdEkW4aYqAw3/bJvd5AnpB2Xl0BpDqAYstOB8Is
# gVufQZXnPBwe/+dB3Hs9WAYjf/j963W2AcbkQq0mr+yskEPTs0LJDgSyUN6qUaPM
# ezTSOgw/qMweRw1T8F+fV4Cz6I2cfmOnY8QbyafNBkjoUILo9vGWglnkg0zx9CSK
# Xh9UiA+iM7J/43CuGPTH+UR9k7NwH5j7FLV68LyRSegcqaGCAyYwggMiBgkqhkiG
# 9w0BCQYxggMTMIIDDwIBATB9MGkxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdp
# Q2VydCwgSW5jLjFBMD8GA1UEAxM4RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3Rh
# bXBpbmcgUlNBNDA5NiBTSEEyNTYgMjAyNSBDQTECEAqA7xhLjfEFgtHEdqeVdGgw
# DQYJYIZIAWUDBAIBBQCgaTAYBgkqhkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqG
# SIb3DQEJBTEPFw0yNTA3MTExOTMxMzdaMC8GCSqGSIb3DQEJBDEiBCCCg7hmtjlr
# FzByfJakZJ4UsUTnQK3P90kdBZfI2BsIVjANBgkqhkiG9w0BAQEFAASCAgCsy9Dv
# XBlLkBzrZ8vI11L5Qg4/xNUWwOfZFB+sOX4A3jokwDVLT3RmnihTPBXR/A2ntQql
# +rvChmCPZkof4ILlacBZ6Cd17aJ9ZaebLL166JsNJvsdeR2gwbeapC6cbZ+T/y8P
# zVdUWSS6M+Nv9TXwdAL7b9h/wQbLqGktO6yqTfhhj4vNQ8kpL5/Y97V70/4hEYPz
# 1BLK39UL8ec6d0XUNGfUPnSivsJwjOq8/f8SncyE7J0sSK8H54ynRiTbCWgv8i3p
# toO6hCQhrrqO9rYKKS7hn4aUBX48j5yM6/F8QLLHXB50hxoQvYTVwCV62dQJWJSP
# OJIczkAhFvjIT9MDe6TPWaWo/U0vZMtUwdoz4mwUHrzHHkCm2FxlmAX4WY81XQmx
# 47FIbATVjLIyENXuUJ38itIbRybSomkeVTHGZYpR4NolrWaEvYIwMPeMJrpGgpBl
# E16YAJWxX+E4LCfHFo5NC3nmfB1aVhCr/HMyumCZklV1dsbdPze2gF8p4SBSY2Ze
# QRn0+rhqRPRq+BJP9a7KytLDzk7o3EDEC4fTzEPA6qtTpmyTz6EoDLh6vSCXVFEx
# dKSqahpfg//jW0VwemBN6x0mLjDefZ3dUUKmfIzdw7xEEDmabzMN/SVcHVQ/NiB6
# /now7/5uHzFmUNyhH3cU+lL2RPxeYIxRwxxKuA==
# SIG # End signature block
