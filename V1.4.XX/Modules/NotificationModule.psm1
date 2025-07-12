# NotificationModule.psm1
# Module for email and Teams notifications

# Module-level variables
$script:EmailConfig = @{}
$script:TeamsConfig = @{}
$script:LogFunction = $null

function Initialize-NotificationModule {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $false)]
        [hashtable]$EmailConfiguration = @{},
        
        [Parameter(Mandatory = $false)]
        [hashtable]$TeamsConfiguration = @{},
        
        [Parameter(Mandatory = $false)]
        [scriptblock]$LogFunction
    )
    
    $script:EmailConfig = $EmailConfiguration
    $script:TeamsConfig = $TeamsConfiguration
    $script:LogFunction = $LogFunction
    
    Write-ModuleLog "NotificationModule initialized" "INFO"
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

function Send-EmailNotification {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string[]]$To,
        
        [Parameter(Mandatory = $true)]
        [string]$Subject,
        
        [Parameter(Mandatory = $true)]
        [string]$Body,
        
        [Parameter(Mandatory = $false)]
        [string[]]$Cc,
        
        [Parameter(Mandatory = $false)]
        [string[]]$Bcc,
        
        [Parameter(Mandatory = $false)]
        [string[]]$Attachments,
        
        [Parameter(Mandatory = $false)]
        [switch]$BodyAsHtml,
        
        [Parameter(Mandatory = $false)]
        [ValidateSet("High", "Normal", "Low")]
        [string]$Priority = "Normal"
    )
    
    try {
        Write-ModuleLog "Sending email to: $($To -join ', ')" "INFO"
        
        # Validate email configuration
        $requiredKeys = @('SMTPServer', 'SMTPPort', 'From')
        foreach ($key in $requiredKeys) {
            if (-not $script:EmailConfig.ContainsKey($key) -or [string]::IsNullOrWhiteSpace($script:EmailConfig[$key])) {
                throw "Missing required email configuration: $key"
            }
        }
        
        # Build email parameters
        $mailParams = @{
            To = $To
            From = $script:EmailConfig.From
            Subject = $Subject
            Body = $Body
            SmtpServer = $script:EmailConfig.SMTPServer
            Port = [int]$script:EmailConfig.SMTPPort
            Priority = $Priority
        }
        
        # Add optional parameters
        if ($Cc) { $mailParams.Cc = $Cc }
        if ($Bcc) { $mailParams.Bcc = $Bcc }
        if ($Attachments) { $mailParams.Attachments = $Attachments }
        if ($BodyAsHtml) { $mailParams.BodyAsHtml = $true }
        
        # SSL configuration
        if ($script:EmailConfig.ContainsKey('UseSSL') -and $script:EmailConfig.UseSSL -eq "1") {
            $mailParams.UseSsl = $true
        }
        
        # Authentication
        if ($script:EmailConfig.ContainsKey('Username') -and $script:EmailConfig.ContainsKey('Password')) {
            $securePassword = ConvertTo-SecureString $script:EmailConfig.Password -AsPlainText -Force
            $credential = New-Object System.Management.Automation.PSCredential($script:EmailConfig.Username, $securePassword)
            $mailParams.Credential = $credential
        }
        
        # Send email
        Send-MailMessage @mailParams -ErrorAction Stop
        
        Write-ModuleLog "Email sent successfully" "INFO"
        
        return @{
            Success = $true
            Message = "Email sent successfully to $($To -join ', ')"
        }
    }
    catch {
        Write-ModuleLog "Error sending email: $($_.Exception.Message)" "ERROR"
        return @{
            Success = $false
            Message = "Error sending email: $($_.Exception.Message)"
        }
    }
}

function Send-WelcomeEmail {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [PSCustomObject]$UserData,
        
        [Parameter(Mandatory = $true)]
        [hashtable]$Config
    )
    
    try {
        Write-ModuleLog "Sending welcome email to: $($UserData.Email)" "INFO"
        
        # Check if welcome email is enabled
        if ($script:EmailConfig.SendWelcomeEmail -ne "1") {
            Write-ModuleLog "Welcome email is disabled in configuration" "INFO"
            return @{
                Success = $true
                Message = "Welcome email is disabled"
            }
        }
        
        # Get email template
        $templatePath = $script:EmailConfig.WelcomeEmailTemplate
        if ([string]::IsNullOrWhiteSpace($templatePath) -or -not (Test-Path $templatePath)) {
            # Use default template
            $emailBody = Get-DefaultWelcomeEmailTemplate
        } else {
            $emailBody = Get-Content -Path $templatePath -Raw
        }
        
        # Replace placeholders
        $emailBody = $emailBody -replace "{{FirstName}}", $UserData.FirstName
        $emailBody = $emailBody -replace "{{LastName}}", $UserData.LastName
        $emailBody = $emailBody -replace "{{DisplayName}}", $UserData.DisplayName
        $emailBody = $emailBody -replace "{{Username}}", $UserData.SamAccountName
        $emailBody = $emailBody -replace "{{Email}}", $UserData.Email
        $emailBody = $emailBody -replace "{{Password}}", $UserData.Password
        $emailBody = $emailBody -replace "{{Department}}", $UserData.DepartmentField
        $emailBody = $emailBody -replace "{{Position}}", $UserData.Position
        $emailBody = $emailBody -replace "{{Company}}", ($Config.Company.CompanyNameFirma ?? "")
        $emailBody = $emailBody -replace "{{Date}}", (Get-Date -Format "dd.MM.yyyy")
        
        # Get subject from config or use default
        $subject = $script:EmailConfig.WelcomeEmailSubject ?? "Welcome to $($Config.Company.CompanyNameFirma)"
        
        # Send email
        $result = Send-EmailNotification -To $UserData.Email -Subject $subject -Body $emailBody -BodyAsHtml
        
        return $result
    }
    catch {
        Write-ModuleLog "Error sending welcome email: $($_.Exception.Message)" "ERROR"
        return @{
            Success = $false
            Message = "Error sending welcome email: $($_.Exception.Message)"
        }
    }
}

function Get-DefaultWelcomeEmailTemplate {
    return @'
<!DOCTYPE html>
<html>
<head>
    <style>
        body { font-family: Arial, sans-serif; line-height: 1.6; color: #333; }
        .container { max-width: 600px; margin: 0 auto; padding: 20px; }
        .header { background-color: #0078D4; color: white; padding: 20px; text-align: center; }
        .content { padding: 20px; background-color: #f9f9f9; }
        .credentials { background-color: #fff; padding: 15px; border: 1px solid #ddd; margin: 20px 0; }
        .footer { text-align: center; padding: 20px; color: #666; font-size: 0.9em; }
        h1 { margin: 0; }
        .label { font-weight: bold; color: #555; }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>Welcome to {{Company}}!</h1>
        </div>
        <div class="content">
            <p>Dear {{FirstName}} {{LastName}},</p>
            
            <p>We are pleased to welcome you to our team! Your account has been successfully created.</p>
            
            <div class="credentials">
                <h3>Your Login Credentials:</h3>
                <p><span class="label">Username:</span> {{Username}}</p>
                <p><span class="label">Email:</span> {{Email}}</p>
                <p><span class="label">Temporary Password:</span> {{Password}}</p>
            </div>
            
            <p><strong>Important:</strong> You will be required to change your password on first login.</p>
            
            <h3>Your Information:</h3>
            <p><span class="label">Department:</span> {{Department}}</p>
            <p><span class="label">Position:</span> {{Position}}</p>
            
            <p>If you have any questions or need assistance, please contact our IT Help Desk.</p>
            
            <p>Best regards,<br>IT Department</p>
        </div>
        <div class="footer">
            <p>This is an automated message. Please do not reply to this email.</p>
            <p>&copy; {{Date}} {{Company}}</p>
        </div>
    </div>
</body>
</html>
'@
}

function Send-TeamsNotification {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$WebhookUrl,
        
        [Parameter(Mandatory = $true)]
        [string]$Title,
        
        [Parameter(Mandatory = $true)]
        [string]$Message,
        
        [Parameter(Mandatory = $false)]
        [ValidateSet("Success", "Warning", "Error", "Information")]
        [string]$Type = "Information",
        
        [Parameter(Mandatory = $false)]
        [hashtable[]]$Facts,
        
        [Parameter(Mandatory = $false)]
        [hashtable[]]$Actions
    )
    
    try {
        Write-ModuleLog "Sending Teams notification: $Title" "INFO"
        
        # Use configured webhook if not provided
        if ([string]::IsNullOrWhiteSpace($WebhookUrl) -and $script:TeamsConfig.ContainsKey('WebhookUrl')) {
            $WebhookUrl = $script:TeamsConfig.WebhookUrl
        }
        
        if ([string]::IsNullOrWhiteSpace($WebhookUrl)) {
            throw "Teams webhook URL is not configured"
        }
        
        # Set theme color based on type
        $themeColor = switch ($Type) {
            "Success" { "00FF00" }
            "Warning" { "FFA500" }
            "Error" { "FF0000" }
            default { "0078D4" }
        }
        
        # Build message card
        $messageCard = @{
            "@type" = "MessageCard"
            "@context" = "http://schema.org/extensions"
            "themeColor" = $themeColor
            "summary" = $Title
            "sections" = @(
                @{
                    "activityTitle" = $Title
                    "activitySubtitle" = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
                    "text" = $Message
                    "markdown" = $true
                }
            )
        }
        
        # Add facts if provided
        if ($Facts) {
            $messageCard.sections[0]["facts"] = $Facts
        }
        
        # Add actions if provided
        if ($Actions) {
            $messageCard["potentialAction"] = $Actions
        }
        
        # Convert to JSON
        $jsonBody = $messageCard | ConvertTo-Json -Depth 10
        
        # Send to Teams
        $response = Invoke-RestMethod -Uri $WebhookUrl -Method Post -Body $jsonBody -ContentType 'application/json'
        
        Write-ModuleLog "Teams notification sent successfully" "INFO"
        
        return @{
            Success = $true
            Message = "Teams notification sent successfully"
        }
    }
    catch {
        Write-ModuleLog "Error sending Teams notification: $($_.Exception.Message)" "ERROR"
        return @{
            Success = $false
            Message = "Error sending Teams notification: $($_.Exception.Message)"
        }
    }
}

function Send-OnboardingNotification {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [PSCustomObject]$UserData,
        
        [Parameter(Mandatory = $false)]
        [ValidateSet("Email", "Teams", "Both")]
        [string]$Method = "Email"
    )
    
    try {
        $results = @()
        
        # Send email notification
        if ($Method -in @("Email", "Both")) {
            if ($script:EmailConfig.ContainsKey('OnboardingNotificationRecipients')) {
                $recipients = $script:EmailConfig.OnboardingNotificationRecipients -split ';'
                
                $subject = "New User Onboarded: $($UserData.DisplayName)"
                $body = @"
A new user has been onboarded:

Name: $($UserData.DisplayName)
Username: $($UserData.SamAccountName)
Email: $($UserData.Email)
Department: $($UserData.DepartmentField)
Position: $($UserData.Position)
Manager: $($UserData.Manager)
Onboarded by: $env:USERNAME
Date: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")
"@
                
                $emailResult = Send-EmailNotification -To $recipients -Subject $subject -Body $body
                $results += $emailResult
            }
        }
        
        # Send Teams notification
        if ($Method -in @("Teams", "Both")) {
            if ($script:TeamsConfig.ContainsKey('OnboardingWebhook')) {
                $facts = @(
                    @{ name = "Name"; value = $UserData.DisplayName }
                    @{ name = "Username"; value = $UserData.SamAccountName }
                    @{ name = "Email"; value = $UserData.Email }
                    @{ name = "Department"; value = $UserData.DepartmentField }
                    @{ name = "Position"; value = $UserData.Position }
                    @{ name = "Onboarded by"; value = $env:USERNAME }
                )
                
                $teamsResult = Send-TeamsNotification `
                    -WebhookUrl $script:TeamsConfig.OnboardingWebhook `
                    -Title "New User Onboarded" `
                    -Message "User **$($UserData.DisplayName)** has been successfully onboarded." `
                    -Type "Success" `
                    -Facts $facts
                    
                $results += $teamsResult
            }
        }
        
        # Check if all notifications were successful
        $allSuccess = $results.Count -gt 0 -and ($results | Where-Object { -not $_.Success }).Count -eq 0
        
        return @{
            Success = $allSuccess
            Results = $results
            Message = if ($allSuccess) { "All notifications sent successfully" } else { "Some notifications failed" }
        }
    }
    catch {
        Write-ModuleLog "Error sending onboarding notification: $($_.Exception.Message)" "ERROR"
        return @{
            Success = $false
            Message = "Error sending onboarding notification: $($_.Exception.Message)"
        }
    }
}

# Export module functions
Export-ModuleMember -Function @(
    'Initialize-NotificationModule',
    'Send-EmailNotification',
    'Send-WelcomeEmail',
    'Send-TeamsNotification',
    'Send-OnboardingNotification'
) 
# SIG # Begin signature block
# MIIcCAYJKoZIhvcNAQcCoIIb+TCCG/UCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCUBxBxAkrrQJSm
# Thd9pcDg350dam8If5/BmJZq72mWPKCCFk4wggMQMIIB+KADAgECAhB3jzsyX9Cg
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
# BgkqhkiG9w0BCQQxIgQgoVeH5JPc3MzmUeLgd+/ChyjVn46hHJf8Sz3ottWHrLAw
# DQYJKoZIhvcNAQEBBQAEggEAMOpXHrojCUup1hFBWNx73H/Xtb+yN/dGgF8GWOZr
# EFF5l3X2Ry5Pd5LLom/unL98LhT97Q8NSiu0oGNUnnAt800yFpQ1RgMA5UMWffXl
# UBaip9oLb2QJqnSW2DxgkDoZBqmnhGdBI7vm1LFyIc/uitSYAb2FX+LXv3VmLF1K
# P1Qkofdge+v+V+SAnDGMkKvdhUlgG/z452WPEBCNl1MxGIRx6lmUPYoZNsvnk7yw
# jJFz6g2tJWVUXxbpkFo7PMoAxKO9TMmc+wgTyah37uv2SBISj8VRzGpPPRrkD8B7
# V53ahwLFd78xObWHolScOc2GgDeBb3BhW2Xp62iAvF2CUaGCAyYwggMiBgkqhkiG
# 9w0BCQYxggMTMIIDDwIBATB9MGkxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdp
# Q2VydCwgSW5jLjFBMD8GA1UEAxM4RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3Rh
# bXBpbmcgUlNBNDA5NiBTSEEyNTYgMjAyNSBDQTECEAqA7xhLjfEFgtHEdqeVdGgw
# DQYJYIZIAWUDBAIBBQCgaTAYBgkqhkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqG
# SIb3DQEJBTEPFw0yNTA3MTExOTMxNTJaMC8GCSqGSIb3DQEJBDEiBCAccTpjLfIF
# KRZ2JHP2Ord0yLIshw2bCXOVeYAwVT6ZeTANBgkqhkiG9w0BAQEFAASCAgBZn9Cv
# iUih4ticO0w3l+0Q3nCHNllYSqPeGFE5BjG9Gj5REpUxjI6Ivl/YF2urMJ8OJeIo
# fM0zbqxPEUWIjW3Ahd9zNqVL3xO1YzAG1NnrWyQ5htVaD0zLF8h+DJk9ZLVLrdp8
# 8jIDXqm0Sd29mFh6UiEMf/b7+hB6UY66BJNAcubCQivd3CyLg26hOqKxNUW7/2tn
# lTK/nizrTBqpXhoPcVZ1IZYpt3gfedp+Dt3kyPyAv4WpLZFUJ9vcxCb3J4D1+E2X
# m9p2ePlXm2iVTi8Gucw6ErrJFv+orIFuEfIoSTLcLFCLjjvmuvfenR9Ol8OoZxK3
# dh6rjs1JGu0n5oX0VCtA1G3UqYl6L0jEi3pblSemb31HZ0ZvPnjuZAzbFovthM/y
# 7Bt0jlxqU57B+1harehaWRjjW+0BkHUk+npxevtAG/GQUpBJTFuSCpwsinrjr0OO
# /RTZcVvXZ96DZppFwVc2Pj2+qMCZk00Tz735Oo3FRRcAzY1sFTLgdLUSCgDKF5Qd
# bbQ5zOqF3RwOZKODuka51KVHETc7ZOZ2dMYu0/N6koRTd8ebryQ0Rxj8f8/Xp45d
# SobzODtdOFfoS1Ij7Dd/5S3qw1K2lrDJCBNywDi+7BBzzQg9DtwJuj4yAv0flS/x
# ayto9EuI9Mz5cyhraG8LTo3yLGP04hHnF12iWA==
# SIG # End signature block
