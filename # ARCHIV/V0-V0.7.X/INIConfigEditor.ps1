# Erweiterter INI Editor mit festen GUI-Einstellungen und Änderungsprotokollierung
# (ohne dynamische Buttons zum Hinzufügen/Löschen von Abschnitten und Schlüsseln)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

#############################################
# 0) Zusätzliche Funktion: Loggen
#############################################
function Log-Change {
    param(
        [string]$Message,
        [string]$LogPath = "ini_changes.log"
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "$timestamp - $Message"
    Add-Content -Path $LogPath -Value $logEntry
}

#############################################
# 1) INI einlesen (mit Kommentarunterstützung)
#############################################
function Read-IniFile {
    param([string]$Path)
    $ini = [ordered]@{}
    $currentSection = ""
    $pendingComment = ""

    foreach ($line in Get-Content $Path) {
        $trimmed = $line.Trim()

        # Kommentarzeile?
        if ($trimmed -match "^\s*;") {
            # Kommentartext ohne führendes ";"
            $commentText = $trimmed.Substring(1).Trim()
            if ($pendingComment -eq "") {
                $pendingComment = $commentText
            }
            else {
                $pendingComment += "`n" + $commentText
            }
            continue
        }

        # Abschnittszeile [Abschnitt]?
        if ($trimmed -match "^\[(.+?)\]") {
            $currentSection = $matches[1]

            # Falls Abschnitt noch nicht existiert => anlegen
            if (-not ($ini.Keys -contains $currentSection)) {
                $ini[$currentSection] = [ordered]@{}
            }
            # Kommentar gilt nicht mehr für nächsten Abschnitt
            $pendingComment = ""
        }

        # key = value ?
        elseif ($trimmed -match "^(.*?)\s*=\s*(.*)$") {
            $key = $matches[1].Trim()
            $value = $matches[2].Trim()

            if ($currentSection -ne "") {
                # Falls in diesem Abschnitt noch nichts existiert, anlegen
                if (-not ($ini.Keys -contains $currentSection)) {
                    $ini[$currentSection] = [ordered]@{}
                }
                # Speichere Schlüssel/Value + Kommentar
                $ini[$currentSection][$key] = [PSCustomObject]@{
                    Value   = $value
                    Comment = $pendingComment
                }
            }
            $pendingComment = ""
        }
        else {
            # Zeile nicht erkannt, Kommentar zurücksetzen
            $pendingComment = ""
        }
    }
    return $ini
}

#############################################
# 2) INI schreiben (Kommentare wieder einfügen)
#############################################
function Write-IniFile {
    param(
        [hashtable]$IniData,
        [string]$Path,
        [System.Collections.ArrayList]$DesiredOrder
    )
    # Abschnitte sortieren: zuerst DesiredOrder, dann Rest
    $remaining = [ordered]@{}
    foreach ($section in $IniData.Keys) {
        $remaining[$section] = $IniData[$section]
    }
    $orderedIni = [ordered]@{}

    foreach ($section in $DesiredOrder) {
        if ($remaining.Keys -contains $section) {
            $orderedIni[$section] = $remaining[$section]
            $remaining.Remove($section) | Out-Null
        }
    }
    foreach ($left in $remaining.Keys) {
        $orderedIni[$left] = $remaining[$left]
    }
    
    $lines = @()
    foreach ($sec in $orderedIni.Keys) {
        $lines += "[$sec]"
        if ($orderedIni[$sec].Count -gt 0) {
            foreach ($key in $orderedIni[$sec].Keys) {
                $entry = $orderedIni[$sec][$key]
                if (-not [string]::IsNullOrEmpty($entry.Comment)) {
                    foreach ($cl in $entry.Comment -split "`n") {
                        $lines += "; " + $cl
                    }
                }
                $lines += "$key=$($entry.Value)"
            }
        }
        $lines += ""  # Leere Zeile als Trenner
    }
    $lines | Set-Content -Path $Path -Encoding UTF8
}

#############################################
# 3) Gewünschte Reihenfolge festlegen
#############################################
$desiredOrder = New-Object System.Collections.ArrayList
[void]$desiredOrder.AddRange(@(
    "ScriptInfo",
    "General",
    "Branding-GUI",
    "Branding-Report",
    "Websites",
    "ADUserDefaults",
    "DisplayNameTemplates",
    "UserCreationDefaults",
    "PasswordFixGenerate",
    "ADGroups",
    "LicensesGroups",
    "ActivateUserMS365ADSync",
    "SignaturGruppe_Optional",
    "STANDORTE",
    "MailEndungen",
    "Company1",
    "Company2",
    "Company3",
    "Company4",
    "Company5"
))

#############################################
# 4) Pfad zur INI-Datei
#############################################
$iniPath = "easyOnboarding*Config.ini"

#############################################
# 5) INI einlesen
#############################################
$iniData = Read-IniFile -Path $iniPath

# Editor-Einstellungen aus [CONFIGEDITOR] lesen
# Fallback-Werte für den Fall, dass [CONFIGEDITOR] nicht existiert
$editorSettings = @{
    FormBackColor     = "#F0F0F0"
    ListViewBackColor = "#FFFFFF"
    DataGridBackColor = "#FAFAFA"
    FontName          = "Segoe UI"
    FontSize          = 10
    LogFilePath       = "ini_changes.log"
}

if ($iniData.Keys -contains "CONFIGEDITOR") {
    # Abschnitt existiert, auslesen
    foreach ($key in $iniData["CONFIGEDITOR"].Keys) {
        $valObj = $iniData["CONFIGEDITOR"][$key]
        if ($valObj -and $editorSettings.ContainsKey($key)) {
            $editorSettings[$key] = $valObj.Value
        }
    }
}

# Log-Dateipfad
$logFilePath = $editorSettings["LogFilePath"]

#############################################
# 6) Erstelle finalOrder (DesiredOrder zuerst, dann Rest)
#############################################
$allSections = $iniData.Keys
$finalOrder = New-Object System.Collections.ArrayList
foreach ($sec in $desiredOrder) {
    if ($allSections -contains $sec) {
        [void]$finalOrder.Add($sec)
    }
}
foreach ($s in $allSections) {
    if (-not $finalOrder.Contains($s)) {
        [void]$finalOrder.Add($s)
    }
}

#############################################
# 7) GUI erstellen: Formular, ListView und DataGridView
#############################################
$form = New-Object System.Windows.Forms.Form
$form.Text = "INI Editor - Feste Reihenfolge"
$form.Size = New-Object System.Drawing.Size(1250,900)
$form.StartPosition = "CenterScreen"
$form.BackColor = [System.Drawing.ColorTranslator]::FromHtml($editorSettings.FormBackColor)

# ListView (Abschnitte)
$listView = New-Object System.Windows.Forms.ListView
$listView.Location = New-Object System.Drawing.Point(20,20)
$listView.Size = New-Object System.Drawing.Size(300,750)
$listView.View = 'Details'
$listView.FullRowSelect = $true
$listView.Sorting = 'None'
$listView.ShowGroups = $false
$listView.Columns.Add("Abschnitt", 280) | Out-Null
$listView.BackColor = [System.Drawing.ColorTranslator]::FromHtml($editorSettings.ListViewBackColor)

foreach ($sec in $finalOrder) {
    $item = New-Object System.Windows.Forms.ListViewItem($sec)
    [void]$listView.Items.Add($item)
}
$form.Controls.Add($listView)

# DataGridView (Schlüssel, Werte, Kommentare)
$dataGrid = New-Object System.Windows.Forms.DataGridView
$dataGrid.Location = New-Object System.Drawing.Point(340,20)
$dataGrid.Size = New-Object System.Drawing.Size(880,750)
$dataGrid.ColumnCount = 3
$dataGrid.Columns[0].Name = "Key"
$dataGrid.Columns[1].Name = "Value"
$dataGrid.Columns[2].Name = "Comment"
$dataGrid.Columns[2].ReadOnly = $true
$dataGrid.AutoSizeColumnsMode = 'None'
$dataGrid.Columns[0].Width = 150
$dataGrid.Columns[1].Width = 400
$dataGrid.Columns[2].Width = 285
$dataGrid.AllowUserToAddRows = $false
$dataGrid.EditMode = 'EditOnKeystrokeOrF2'
$dataGrid.BackColor = [System.Drawing.ColorTranslator]::FromHtml($editorSettings.DataGridBackColor)
$form.Controls.Add($dataGrid)

#############################################
# 8) ListView-Auswahl-Event: Abschnitt laden
#############################################
$currentSection = $null
$listView.Add_SelectedIndexChanged({
    if ($listView.SelectedItems.Count -gt 0) {
        $currentSection = $listView.SelectedItems[0].Text
        $dataGrid.Rows.Clear()

        if ($iniData.Keys -contains $currentSection) {
            foreach ($key in $iniData[$currentSection].Keys) {
                $entry = $iniData[$currentSection][$key]
                $row = @($key, $entry.Value, $entry.Comment)
                [void]$dataGrid.Rows.Add($row)
            }
        }
    }
})

#############################################
# 9) Buttons: Save, Info, Close
#############################################

# Save Changes (hellgrün)
$saveButton = New-Object System.Windows.Forms.Button
$saveButton.Text = "Save Changes"
$saveButton.Size = New-Object System.Drawing.Size(120,40)
$saveButton.Location = New-Object System.Drawing.Point(20,780)
$saveButton.BackColor = [System.Drawing.Color]::LightGreen
$saveButton.Add_Click({
    if ($listView.SelectedItems.Count -gt 0) {
        $currentSection = $listView.SelectedItems[0].Text
    }
    if ([string]::IsNullOrEmpty($currentSection)) {
        [System.Windows.Forms.MessageBox]::Show("Bitte wählen Sie einen Abschnitt aus.", "Hinweis")
        return
    }

    # Neue Section-Daten aufbauen
    $sectionData = [ordered]@{}
    foreach ($row in $dataGrid.Rows) {
        $key = $row.Cells[0].Value
        $value = $row.Cells[1].Value
        $comment = $row.Cells[2].Value

        if (-not [string]::IsNullOrEmpty($key)) {
            $sectionData[$key] = [PSCustomObject]@{
                Value   = $value
                Comment = $comment
            }
        }
    }

    # Abschnitt ersetzen
    $iniData[$currentSection] = $sectionData

    # INI-Datei schreiben
    Write-IniFile -IniData $iniData -Path $iniPath -DesiredOrder $desiredOrder

    [System.Windows.Forms.MessageBox]::Show("INI-Datei erfolgreich gespeichert.", "Gespeichert")

    Log-Change -Message "Abschnitt '$currentSection' wurde gespeichert." -LogPath $logFilePath
})
$form.Controls.Add($saveButton)

# Info Button (hellblau) – öffnet info_editor.txt
$infoButton = New-Object System.Windows.Forms.Button
$infoButton.Text = "Info"
$infoButton.Size = New-Object System.Drawing.Size(80,40)
$infoButton.Location = New-Object System.Drawing.Point(160,780)
$infoButton.BackColor = [System.Drawing.Color]::LightBlue
$infoButton.Add_Click({
    if (Test-Path "info_editor.txt") {
        Invoke-Item "info_editor.txt"
    }
    else {
        [System.Windows.Forms.MessageBox]::Show("Die Datei 'info_editor.txt' wurde nicht gefunden.", "Info")
    }
})
$form.Controls.Add($infoButton)

# Close Button (hellrot)
$closeButton = New-Object System.Windows.Forms.Button
$closeButton.Text = "Close"
$closeButton.Size = New-Object System.Drawing.Size(80,40)
$closeButton.Location = New-Object System.Drawing.Point(260,780)
$closeButton.BackColor = [System.Drawing.Color]::LightCoral
$closeButton.Add_Click({
    $form.Close()
})
$form.Controls.Add($closeButton)

#############################################
# 10) Formular anzeigen
#############################################
$form.Add_Shown({ $form.Activate() })
[void]$form.ShowDialog()

# SIG # Begin signature block
# MIIbywYJKoZIhvcNAQcCoIIbvDCCG7gCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBspZB1ESkvTAUS
# fCQW26NEYRen/lYnSIQJQZ2F1ENbH6CCFhcwggMQMIIB+KADAgECAhB3jzsyX9Cg
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
# DQEJBDEiBCBvqACJZ4ecu5CehuvZzqRkzG7+xYuKaCLlmIC/gI4bnjANBgkqhkiG
# 9w0BAQEFAASCAQBmcgUYerOWEfqvFmGnw5blIrwx2DfcbgH6DXA+s4vopy5DWw5g
# pC1qH0vsESG5GZBZVBzwarg5kQAcqcxyvykHQq260G8+xz8Yr4jpXNk5J0PXhnUk
# bnCcdMKl/4zEZG4v+bMMf6DLbwNJqI+eSeXNQ0IYFAxln4BJob7mtiQBS2pi8ehL
# gDwiT8ds/IbBWsZGYLH1g7zzce2bRJNqfoVXxX68rWTxcxTsjXi18yyH3wPc/41r
# s2rbf6FvgNTZ6OwaQq8cVfUocWC3aSnXYs/7ZPGWXffBJWAL5kBWomQx50YWGFGZ
# 0Puh6Y4iRoyqn/of6oVXtrW9J7F2de3x5a9ZoYIDIDCCAxwGCSqGSIb3DQEJBjGC
# Aw0wggMJAgEBMHcwYzELMAkGA1UEBhMCVVMxFzAVBgNVBAoTDkRpZ2lDZXJ0LCBJ
# bmMuMTswOQYDVQQDEzJEaWdpQ2VydCBUcnVzdGVkIEc0IFJTQTQwOTYgU0hBMjU2
# IFRpbWVTdGFtcGluZyBDQQIQC65mvFq6f5WHxvnpBOMzBDANBglghkgBZQMEAgEF
# AKBpMBgGCSqGSIb3DQEJAzELBgkqhkiG9w0BBwEwHAYJKoZIhvcNAQkFMQ8XDTI1
# MDcwNTEwMTE1OFowLwYJKoZIhvcNAQkEMSIEINj/zkM4OXnPUlvKC87Y8lGe9Q8e
# uCJ/nHJs69XL6r7rMA0GCSqGSIb3DQEBAQUABIICAGOmDSTQmFuYbxWjfEFvW9pQ
# j9LBM/zZeExzM77+Gr2ZeGC/6TzHCrFEQ4A4Lm3yQm9Wx9n6Wj+oNjk0HKGmAZzM
# KrTmhlm6/xhz5Q9yXUXPJIrdMM88wYMax6xZ/nuE3dMtherzLyYJaaWHNLnyYECY
# quP7i8GwwH5MHHd0RBcYwTYUfDdLaFZKCASjUQXa1nzyweKhh6ptHPOI4E6Nlbuv
# 3ZCrcYw58S4Vao5OLFzbn9yNVieNmXOznOOWMpMw41uZi1QQz20tSxJVI1ktqwaP
# XViZsCwRK8aM/toPFXslhtGUcA/jqnuuY2HXLx3GI5SDRtPTWkaBSjWa1AKdqiYK
# gfTMw5CHOCUd3vkvfWAruIrZy2NL0qXDHl8yfTnGNWJ9o8sn1fCH31fMYgeUmyKq
# n/jPnVSWU2zWJ02KDv1mWmGA4hETy/7SJw314Zbn5dNVVKpOX4kciohqCny7m9ZI
# Ol6H2JwS05u0QbqOzZuZvVLAV4UWQ5aYjzUbOgcB3INjt0gLcV+cEze6yqwBmv42
# jBJm4F6PgOGAQmRvzQ5SrNFbevvXjbDD++wch4G22zujJpAOcTIJkoK2/mGNPL12
# wdzyc/yuSxC4iYahK+d/D+c3VheskKXlAm5Ko3UD5Dt0S9gOdxvr5PxBFzMdCkWY
# suXQG0KHsEu+LwmrMGA9
# SIG # End signature block
