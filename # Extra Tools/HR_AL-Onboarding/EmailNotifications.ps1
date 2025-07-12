function Send-OnboardingEmail {
    param (
        [Parameter(Mandatory=$true)]
        [string]$RecipientEmail,
        
        [Parameter(Mandatory=$true)]
        [string]$Subject,
        
        [Parameter(Mandatory=$true)]
        [string]$Body,
        
        [string]$SMTPServer = "smtp.yourcompany.com",
        [int]$SMTPPort = 25,
        [string]$FromAddress = "onboarding@yourcompany.com",
        [string]$FromName = "easyONBOARDING System",
        [switch]$HTML,
        [System.Net.NetworkCredential]$Credential,
        [switch]$UseSSL,
        [string[]]$CC,
        [string[]]$BCC,
        [string]$Priority = "Normal",
        [string[]]$AttachmentPaths
    )
    
    try {
        # E-Mail-Nachricht erstellen
        $message = New-Object System.Net.Mail.MailMessage
        
        # Absender einrichten
        if ([string]::IsNullOrWhiteSpace($FromName)) {
            $message.From = New-Object System.Net.Mail.MailAddress($FromAddress)
        } else {
            $message.From = New-Object System.Net.Mail.MailAddress($FromAddress, $FromName)
        }
        
        # Empfänger einrichten
        $message.To.Add($RecipientEmail)
        
        # CC-Empfänger hinzufügen, wenn vorhanden
        if ($CC -and $CC.Count -gt 0) {
            foreach ($ccAddress in $CC) {
                if (-not [string]::IsNullOrWhiteSpace($ccAddress)) {
                    $message.CC.Add($ccAddress)
                }
            }
        }
        
        # BCC-Empfänger hinzufügen, wenn vorhanden
        if ($BCC -and $BCC.Count -gt 0) {
            foreach ($bccAddress in $BCC) {
                if (-not [string]::IsNullOrWhiteSpace($bccAddress)) {
                    $message.Bcc.Add($bccAddress)
                }
            }
        }
        
        # Betreff und Inhalt einrichten
        $message.Subject = $Subject
        $message.Body = $Body
        $message.IsBodyHtml = $HTML.IsPresent
        
        # Priorität setzen
        switch ($Priority) {
            "High" { $message.Priority = [System.Net.Mail.MailPriority]::High }
            "Low" { $message.Priority = [System.Net.Mail.MailPriority]::Low }
            default { $message.Priority = [System.Net.Mail.MailPriority]::Normal }
        }
        
        # Anhänge hinzufügen, wenn vorhanden
        if ($AttachmentPaths -and $AttachmentPaths.Count -gt 0) {
            foreach ($attachment in $AttachmentPaths) {
                if (Test-Path $attachment) {
                    $message.Attachments.Add((New-Object System.Net.Mail.Attachment($attachment)))
                } else {
                    Write-Warning "Die Anhangsdatei '$attachment' wurde nicht gefunden und wird übersprungen."
                }
            }
        }
        
        # SMTP-Client erstellen
        $smtpClient = New-Object System.Net.Mail.SmtpClient($SMTPServer, $SMTPPort)
        
        # SSL einrichten, wenn aktiviert
        $smtpClient.EnableSsl = $UseSSL.IsPresent
        
        # Anmeldeinformationen einrichten, wenn vorhanden
        if ($Credential) {
            $smtpClient.Credentials = $Credential
        } else {
            $smtpClient.UseDefaultCredentials = $true
        }
        
        # E-Mail senden
        $smtpClient.Send($message)
        
        # Nachricht freigeben
        $message.Dispose()
        
        return @{
            Success = $true
            ErrorMessage = $null
        }
    }
    catch {
        $errorMessage = "Fehler beim Senden der E-Mail: $($_.Exception.Message)"
        Write-Log -Message $errorMessage -Level "ERROR"
        
        return @{
            Success = $false
            ErrorMessage = $errorMessage
        }
    }
}

function Send-OnboardingNotification {
    param (
        [Parameter(Mandatory=$true)]
        [string]$RecipientUsername,
        
        [Parameter(Mandatory=$true)]
        [string]$NotificationType,
        
        [Parameter(Mandatory=$true)]
        [hashtable]$RecordData,
        
        [string]$EmailDomain = "yourcompany.com",
        [string]$CustomMessage = "",
        [switch]$HighPriority
    )
    
    # E-Mail-Adresse aus Benutzername und Domain erstellen
    $recipientEmail = "$RecipientUsername@$EmailDomain"
    
    # Basisinformationen aus dem Datensatz extrahieren
    $employeeName = "$($RecordData.FirstName) $($RecordData.LastName)"
    $startDate = $RecordData.StartWorkDate
    $position = $RecordData.Position
    $office = $RecordData.OfficeRoom
    
    # Template basierend auf Benachrichtigungstyp auswählen
    switch ($NotificationType) {
        "NewRequest" {
            $subject = "Neue Onboarding-Anfrage für $employeeName"
            $body = @"
<html>
<body style="font-family: Arial, sans-serif; line-height: 1.6;">
<h2>Neue Onboarding-Anfrage</h2>
<p>Sehr geehrte(r) $RecipientUsername,</p>
<p>eine neue Onboarding-Anfrage wurde erstellt und wartet auf Ihre Bearbeitung:</p>
<ul>
  <li><strong>Name:</strong> $employeeName</li>
  <li><strong>Startdatum:</strong> $startDate</li>
  <li><strong>Position:</strong> $position</li>
  <li><strong>Büro:</strong> $office</li>
</ul>
<p>Bitte melden Sie sich im easyONBOARDING HR-AL Tool an, um die Anfrage zu bearbeiten.</p>
<p>$CustomMessage</p>
<p>Mit freundlichen Grüßen<br>
Ihr easyONBOARDING Team</p>
</body>
</html>
"@
        }
        
        "PendingVerification" {
            $subject = "Anfrage bereit zur Verifizierung: $employeeName"
            $body = @"
<html>
<body style="font-family: Arial, sans-serif; line-height: 1.6;">
<h2>Anfrage bereit zur Verifizierung</h2>
<p>Sehr geehrte(r) $RecipientUsername,</p>
<p>die Onboarding-Anfrage für $employeeName wurde vom Manager ergänzt und wartet auf Ihre Verifizierung:</p>
<ul>
  <li><strong>Name:</strong> $employeeName</li>
  <li><strong>Startdatum:</strong> $startDate</li>
  <li><strong>Position:</strong> $position</li>
  <li><strong>Büro:</strong> $office</li>
  <li><strong>Abteilung:</strong> $($RecordData.DepartmentField)</li>
</ul>
<p>Bitte melden Sie sich im easyONBOARDING HR-AL Tool an, um die Anfrage zu verifizieren.</p>
<p>$CustomMessage</p>
<p>Mit freundlichen Grüßen<br>
Ihr easyONBOARDING Team</p>
</body>
</html>
"@
        }
        
        "ReadyForIT" {
            $subject = "Neue IT-Anfrage: Onboarding für $employeeName"
            $body = @"
<html>
<body style="font-family: Arial, sans-serif; line-height: 1.6;">
<h2>Neue IT-Anfrage für Onboarding</h2>
<p>Sehr geehrte(r) $RecipientUsername,</p>
<p>eine neue Onboarding-Anfrage wurde verifiziert und ist bereit für die IT-Bearbeitung:</p>
<ul>
  <li><strong>Name:</strong> $employeeName</li>
  <li><strong>Startdatum:</strong> $startDate</li>
  <li><strong>Position:</strong> $position</li>
  <li><strong>Büro:</strong> $office</li>
  <li><strong>Abteilung:</strong> $($RecordData.DepartmentField)</li>
</ul>
<p>Bitte melden Sie sich im easyONBOARDING HR-AL Tool an, um die notwendigen IT-Maßnahmen durchzuführen.</p>
<p>$CustomMessage</p>
<p>Mit freundlichen Grüßen<br>
Ihr easyONBOARDING Team</p>
</body>
</html>
"@
        }
        
        "Completed" {
            $subject = "Onboarding abgeschlossen: $employeeName"
            $body = @"
<html>
<body style="font-family: Arial, sans-serif; line-height: 1.6;">
<h2>Onboarding-Prozess abgeschlossen</h2>
<p>Sehr geehrte(r) $RecipientUsername,</p>
<p>der Onboarding-Prozess für $employeeName wurde erfolgreich abgeschlossen:</p>
<ul>
  <li><strong>Name:</strong> $employeeName</li>
  <li><strong>Startdatum:</strong> $startDate</li>
  <li><strong>Position:</strong> $position</li>
  <li><strong>Büro:</strong> $office</li>
  <li><strong>Abteilung:</strong> $($RecordData.DepartmentField)</li>
  <li><strong>IT-Status:</strong> Einrichtung abgeschlossen</li>
</ul>
<p>Alle notwendigen Systeme wurden eingerichtet und die Ausstattung ist bereit.</p>
<p>$CustomMessage</p>
<p>Mit freundlichen Grüßen<br>
Ihr easyONBOARDING Team</p>
</body>
</html>
"@
        }
        
        default {
            $subject = "Onboarding-Benachrichtigung: $employeeName"
            $body = @"
<html>
<body style="font-family: Arial, sans-serif; line-height: 1.6;">
<h2>Onboarding-Benachrichtigung</h2>
<p>Sehr geehrte(r) $RecipientUsername,</p>
<p>es gibt eine Aktualisierung im Onboarding-Prozess für $employeeName:</p>
<p>$CustomMessage</p>
<p>Mit freundlichen Grüßen<br>
Ihr easyONBOARDING Team</p>
</body>
</html>
"@
        }
    }
    
    # E-Mail senden
    $emailParams = @{
        RecipientEmail = $recipientEmail
        Subject = $subject
        Body = $body
        HTML = $true
    }
    
    # Priority setzen, wenn angefordert
    if ($HighPriority) {
        $emailParams.Priority = "High"
    }
    
    # E-Mail senden und Ergebnis zurückgeben
    return Send-OnboardingEmail @emailParams
}

# SIG # Begin signature block
# MIIbywYJKoZIhvcNAQcCoIIbvDCCG7gCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAYMqVeTjorqjZ3
# Emev0gHjvgqk5ex/mZLtI8LpOGaBPaCCFhcwggMQMIIB+KADAgECAhB3jzsyX9Cg
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
# LSd+2DrZ8LaHlv1b0VysGMNNn3O3AamfV6peKOK5lDCCBq4wggSWoAMCAQICEAc2
# N7ckVHzYR6z9KGYqXlswDQYJKoZIhvcNAQELBQAwYjELMAkGA1UEBhMCVVMxFTAT
# BgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNvbTEh
# MB8GA1UEAxMYRGlnaUNlcnQgVHJ1c3RlZCBSb290IEc0MB4XDTIyMDMyMzAwMDAw
# MFoXDTM3MDMyMjIzNTk1OVowYzELMAkGA1UEBhMCVVMxFzAVBgNVBAoTDkRpZ2lD
# ZXJ0LCBJbmMuMTswOQYDVQQDEzJEaWdpQ2VydCBUcnVzdGVkIEc0IFJTQTQwOTYg
# U0hBMjU2IFRpbWVTdGFtcGluZyBDQTCCAiIwDQYJKoZIhvcNAQEBBQADggIPADCC
# AgoCggIBAMaGNQZJs8E9cklRVcclA8TykTepl1Gh1tKD0Z5Mom2gsMyD+Vr2EaFE
# FUJfpIjzaPp985yJC3+dH54PMx9QEwsmc5Zt+FeoAn39Q7SE2hHxc7Gz7iuAhIoi
# GN/r2j3EF3+rGSs+QtxnjupRPfDWVtTnKC3r07G1decfBmWNlCnT2exp39mQh0YA
# e9tEQYncfGpXevA3eZ9drMvohGS0UvJ2R/dhgxndX7RUCyFobjchu0CsX7LeSn3O
# 9TkSZ+8OpWNs5KbFHc02DVzV5huowWR0QKfAcsW6Th+xtVhNef7Xj3OTrCw54qVI
# 1vCwMROpVymWJy71h6aPTnYVVSZwmCZ/oBpHIEPjQ2OAe3VuJyWQmDo4EbP29p7m
# O1vsgd4iFNmCKseSv6De4z6ic/rnH1pslPJSlRErWHRAKKtzQ87fSqEcazjFKfPK
# qpZzQmiftkaznTqj1QPgv/CiPMpC3BhIfxQ0z9JMq++bPf4OuGQq+nUoJEHtQr8F
# nGZJUlD0UfM2SU2LINIsVzV5K6jzRWC8I41Y99xh3pP+OcD5sjClTNfpmEpYPtMD
# iP6zj9NeS3YSUZPJjAw7W4oiqMEmCPkUEBIDfV8ju2TjY+Cm4T72wnSyPx4Jduyr
# XUZ14mCjWAkBKAAOhFTuzuldyF4wEr1GnrXTdrnSDmuZDNIztM2xAgMBAAGjggFd
# MIIBWTASBgNVHRMBAf8ECDAGAQH/AgEAMB0GA1UdDgQWBBS6FtltTYUvcyl2mi91
# jGogj57IbzAfBgNVHSMEGDAWgBTs1+OC0nFdZEzfLmc/57qYrhwPTzAOBgNVHQ8B
# Af8EBAMCAYYwEwYDVR0lBAwwCgYIKwYBBQUHAwgwdwYIKwYBBQUHAQEEazBpMCQG
# CCsGAQUFBzABhhhodHRwOi8vb2NzcC5kaWdpY2VydC5jb20wQQYIKwYBBQUHMAKG
# NWh0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRSb290
# RzQuY3J0MEMGA1UdHwQ8MDowOKA2oDSGMmh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNv
# bS9EaWdpQ2VydFRydXN0ZWRSb290RzQuY3JsMCAGA1UdIAQZMBcwCAYGZ4EMAQQC
# MAsGCWCGSAGG/WwHATANBgkqhkiG9w0BAQsFAAOCAgEAfVmOwJO2b5ipRCIBfmbW
# 2CFC4bAYLhBNE88wU86/GPvHUF3iSyn7cIoNqilp/GnBzx0H6T5gyNgL5Vxb122H
# +oQgJTQxZ822EpZvxFBMYh0MCIKoFr2pVs8Vc40BIiXOlWk/R3f7cnQU1/+rT4os
# equFzUNf7WC2qk+RZp4snuCKrOX9jLxkJodskr2dfNBwCnzvqLx1T7pa96kQsl3p
# /yhUifDVinF2ZdrM8HKjI/rAJ4JErpknG6skHibBt94q6/aesXmZgaNWhqsKRcnf
# xI2g55j7+6adcq/Ex8HBanHZxhOACcS2n82HhyS7T6NJuXdmkfFynOlLAlKnN36T
# U6w7HQhJD5TNOXrd/yVjmScsPT9rp/Fmw0HNT7ZAmyEhQNC3EyTN3B14OuSereU0
# cZLXJmvkOHOrpgFPvT87eK1MrfvElXvtCl8zOYdBeHo46Zzh3SP9HSjTx/no8Zhf
# +yvYfvJGnXUsHicsJttvFXseGYs2uJPU5vIXmVnKcPA3v5gA3yAWTyf7YGcWoWa6
# 3VXAOimGsJigK+2VQbc61RWYMbRiCQ8KvYHZE/6/pNHzV9m8BPqC3jLfBInwAM1d
# wvnQI38AC+R2AibZ8GV2QqYphwlHK+Z/GqSFD/yYlvZVVCsfgPrA8g4r5db7qS9E
# FUrnEw4d2zc4GqEr9u3WfPwwgga8MIIEpKADAgECAhALrma8Wrp/lYfG+ekE4zME
# MA0GCSqGSIb3DQEBCwUAMGMxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2Vy
# dCwgSW5jLjE7MDkGA1UEAxMyRGlnaUNlcnQgVHJ1c3RlZCBHNCBSU0E0MDk2IFNI
# QTI1NiBUaW1lU3RhbXBpbmcgQ0EwHhcNMjQwOTI2MDAwMDAwWhcNMzUxMTI1MjM1
# OTU5WjBCMQswCQYDVQQGEwJVUzERMA8GA1UEChMIRGlnaUNlcnQxIDAeBgNVBAMT
# F0RpZ2lDZXJ0IFRpbWVzdGFtcCAyMDI0MIICIjANBgkqhkiG9w0BAQEFAAOCAg8A
# MIICCgKCAgEAvmpzn/aVIauWMLpbbeZZo7Xo/ZEfGMSIO2qZ46XB/QowIEMSvgjE
# dEZ3v4vrrTHleW1JWGErrjOL0J4L0HqVR1czSzvUQ5xF7z4IQmn7dHY7yijvoQ7u
# jm0u6yXF2v1CrzZopykD07/9fpAT4BxpT9vJoJqAsP8YuhRvflJ9YeHjes4fduks
# THulntq9WelRWY++TFPxzZrbILRYynyEy7rS1lHQKFpXvo2GePfsMRhNf1F41nyE
# g5h7iOXv+vjX0K8RhUisfqw3TTLHj1uhS66YX2LZPxS4oaf33rp9HlfqSBePejlY
# eEdU740GKQM7SaVSH3TbBL8R6HwX9QVpGnXPlKdE4fBIn5BBFnV+KwPxRNUNK6lY
# k2y1WSKour4hJN0SMkoaNV8hyyADiX1xuTxKaXN12HgR+8WulU2d6zhzXomJ2Ple
# I9V2yfmfXSPGYanGgxzqI+ShoOGLomMd3mJt92nm7Mheng/TBeSA2z4I78JpwGpT
# RHiT7yHqBiV2ngUIyCtd0pZ8zg3S7bk4QC4RrcnKJ3FbjyPAGogmoiZ33c1HG93V
# p6lJ415ERcC7bFQMRbxqrMVANiav1k425zYyFMyLNyE1QulQSgDpW9rtvVcIH7Wv
# G9sqYup9j8z9J1XqbBZPJ5XLln8mS8wWmdDLnBHXgYly/p1DhoQo5fkCAwEAAaOC
# AYswggGHMA4GA1UdDwEB/wQEAwIHgDAMBgNVHRMBAf8EAjAAMBYGA1UdJQEB/wQM
# MAoGCCsGAQUFBwMIMCAGA1UdIAQZMBcwCAYGZ4EMAQQCMAsGCWCGSAGG/WwHATAf
# BgNVHSMEGDAWgBS6FtltTYUvcyl2mi91jGogj57IbzAdBgNVHQ4EFgQUn1csA3cO
# KBWQZqVjXu5Pkh92oFswWgYDVR0fBFMwUTBPoE2gS4ZJaHR0cDovL2NybDMuZGln
# aWNlcnQuY29tL0RpZ2lDZXJ0VHJ1c3RlZEc0UlNBNDA5NlNIQTI1NlRpbWVTdGFt
# cGluZ0NBLmNybDCBkAYIKwYBBQUHAQEEgYMwgYAwJAYIKwYBBQUHMAGGGGh0dHA6
# Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBYBggrBgEFBQcwAoZMaHR0cDovL2NhY2VydHMu
# ZGlnaWNlcnQuY29tL0RpZ2lDZXJ0VHJ1c3RlZEc0UlNBNDA5NlNIQTI1NlRpbWVT
# dGFtcGluZ0NBLmNydDANBgkqhkiG9w0BAQsFAAOCAgEAPa0eH3aZW+M4hBJH2UOR
# 9hHbm04IHdEoT8/T3HuBSyZeq3jSi5GXeWP7xCKhVireKCnCs+8GZl2uVYFvQe+p
# PTScVJeCZSsMo1JCoZN2mMew/L4tpqVNbSpWO9QGFwfMEy60HofN6V51sMLMXNTL
# fhVqs+e8haupWiArSozyAmGH/6oMQAh078qRh6wvJNU6gnh5OruCP1QUAvVSu4kq
# VOcJVozZR5RRb/zPd++PGE3qF1P3xWvYViUJLsxtvge/mzA75oBfFZSbdakHJe2B
# VDGIGVNVjOp8sNt70+kEoMF+T6tptMUNlehSR7vM+C13v9+9ZOUKzfRUAYSyyEmY
# tsnpltD/GWX8eM70ls1V6QG/ZOB6b6Yum1HvIiulqJ1Elesj5TMHq8CWT/xrW7tw
# ipXTJ5/i5pkU5E16RSBAdOp12aw8IQhhA/vEbFkEiF2abhuFixUDobZaA0VhqAsM
# HOmaT3XThZDNi5U2zHKhUs5uHHdG6BoQau75KiNbh0c+hatSF+02kULkftARjsyE
# pHKsF7u5zKRbt5oK5YGwFvgc4pEVUNytmB3BpIiowOIIuDgP5M9WArHYSAR16gc0
# dP2XdkMEP5eBsX7bf/MGN4K3HP50v/01ZHo/Z5lGLvNwQ7XHBx1yomzLP8lx4Q1z
# ZKDyHcp4VQJLu2kWTsKsOqQxggUKMIIFBgIBATA0MCAxHjAcBgNVBAMMFVBoaW5J
# VC1QU3NjcmlwdHNfU2lnbgIQd487Ml/QoIxIvrAQtqwTEzANBglghkgBZQMEAgEF
# AKCBhDAYBgorBgEEAYI3AgEMMQowCKACgAChAoAAMBkGCSqGSIb3DQEJAzEMBgor
# BgEEAYI3AgEEMBwGCisGAQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3
# DQEJBDEiBCDxlS5RirFgCaaW/NsirD/BnH7axIkRCITYhKRT7XXMvDANBgkqhkiG
# 9w0BAQEFAASCAQANLVJIrmcUzWyu9bTy/54YnuEQRQ4Y9bBMHbWd7zagyTfObh+B
# +ShXizxzofMDe3SLEiuIrG7sSeP9iqhrKDD/RxpK/hRjtC+b9BQyLco9NmE6uoul
# 1ce5yJsTCSBYRYmasWzP7wMcB+XuHi5gaRhhc3xOAPZ3sG6Y4RKgFQPZEIvpZzR2
# xwRqgIz4B+ldz+3oRd4vxwy4DeJo4enMLj5Dorm1aC/WYhgFXdZX1OkOMeJLGWys
# MV4KfeECSRGvwaCkkatyT+9MPo8aTbudH+XmLaU+AxT0CBzrodrl0dRxReGI4ZFx
# YYdesd5u9BDSCLy12NxvwRznW15vybCbuiaOoYIDIDCCAxwGCSqGSIb3DQEJBjGC
# Aw0wggMJAgEBMHcwYzELMAkGA1UEBhMCVVMxFzAVBgNVBAoTDkRpZ2lDZXJ0LCBJ
# bmMuMTswOQYDVQQDEzJEaWdpQ2VydCBUcnVzdGVkIEc0IFJTQTQwOTYgU0hBMjU2
# IFRpbWVTdGFtcGluZyBDQQIQC65mvFq6f5WHxvnpBOMzBDANBglghkgBZQMEAgEF
# AKBpMBgGCSqGSIb3DQEJAzELBgkqhkiG9w0BBwEwHAYJKoZIhvcNAQkFMQ8XDTI1
# MDcwNTEwMTIwM1owLwYJKoZIhvcNAQkEMSIEII2swuVx90uHNFCJ7husyacf2sJF
# X3HIQLlZKNVBQ+LYMA0GCSqGSIb3DQEBAQUABIICAEpDLnPnCYIRU97utMGCMtT9
# 6SF/S6C/elC/tLcgwmzr5dif6wlUUykzDXgA39s/UhbcA+Ir1Nvo4cRD4GPI9RfD
# y2wxusY5AW04WPY3IYE7VqK8qAacDqiUPir3m1Uz1kJm4wCIJ/EURMOX2iAGzB24
# jz3s8fTHYx5hUHVF9nIRMor9N+N+5iKsFk09H5OkmZrGLQKooedO2+ZDMq+HPcYz
# CeoDmH6kK50q7ZbPIvQaLhMIgA9N1DEtRHZM974pkHWCjcErE++nH6wBj3oFfz++
# WPfRCB7JiZ6nqIvELlL9Xa6V1jBAUHD9hN/2/KXXgtzOhLkMIjZquKom3NO0mDJZ
# o87v0/SqfD6Nsw7q2YPredNm+gqzZA1Z63lqYgW8hmkSjJSosMG6KvQui2PwIlXM
# m1LH6thl7g4FAoV8yut3ktk5QBtHNaFxg8OKAL9xzoo0MOrQqfXOlpshyPx/ZDuK
# X6mf1zjs4mpQYCfx8gF02xt4UwrjU0oh1KLjDOo6mzblK0w2f9z3QR33dN773uFY
# 39DlOvrR/i+f44LbFVFWLDBZtTRvJiYjaBaH7XunZyWzCXGpCEA20unoeKgDxjiG
# SP8A1moQXaVZa1YYJ5PxxzOnJLoBs0MJm7fH9aNLM2cjYeue3d04xW/8LB/jfe1q
# OBX6s5/thcwYcoDqaPqy
# SIG # End signature block
