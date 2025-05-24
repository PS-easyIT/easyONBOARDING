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
