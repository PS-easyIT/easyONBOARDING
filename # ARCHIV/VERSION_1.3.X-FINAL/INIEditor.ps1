# Erweiterter INI Editor mit festen GUI-Einstellungen und Änderungsprotokollierung
# (ohne dynamische Buttons zum Hinzufügen/Löschen von Abschnitten und Schlüsseln)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

#############################################
# Zusätzliche Funktion: Loggen
#############################################
function Log-Change {
    param(
        [string]$Message,
        [string]$LogPath = "LOG_easyINIEditor-Changes.log"
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "$timestamp - $Message"
    Add-Content -Path $LogPath -Value $logEntry
}

#############################################
# INI einlesen (mit Kommentarunterstützung)
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

        # key = value ? (abgesehen von reinen Leer-/Kommentarzeilen)
        elseif ($trimmed -match "^(.*?)\s*=\s*(.*)$") {
            $key = $matches[1].Trim()
            $value = $matches[2].Trim()

            if ($currentSection -ne "") {
                if (-not ($ini.Keys -contains $currentSection)) {
                    $ini[$currentSection] = [ordered]@{}
                }
                $ini[$currentSection][$key] = [PSCustomObject]@{
                    Value   = $value
                    Comment = $pendingComment
                }
            }
            $pendingComment = ""
        }
        else {
            # Unbekannte Zeile -> Kommentar zurücksetzen
            $pendingComment = ""
        }
    }
    return $ini
}

#############################################
# INI schreiben (Kommentare wieder einfügen)
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
# Gewünschte Reihenfolge festlegen
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
    "Company5",
    "CONFIGEDITOR"
))

#############################################
# Pfad zur INI-Datei (alle Dateien mit .ini laden)
#############################################
$iniPath = "*.ini"

#############################################
# Alle INI-Dateien einlesen und in $AllIniData speichern
#############################################
$AllIniData = @{}
$files = Get-ChildItem -Path $iniPath -File -ErrorAction SilentlyContinue
if (-not $files) {
    Write-Host "Keine INI-Dateien gefunden im aktuellen Verzeichnis."
    return
}

foreach ($f in $files) {
    $AllIniData[$f.FullName] = Read-IniFile -Path $f.FullName
}

#############################################
# Editor-Einstellungen aus [CONFIGEDITOR] aus der ersten INI-Datei laden (falls vorhanden)
#############################################
$editorSettings = @{
    FormBackColor     = "#F0F0F0"
    ListViewBackColor = "#FFFFFF"
    DataGridBackColor = "#FAFAFA"
    FontName          = "Segoe UI"
    FontSize          = 10
    LogFilePath       = "ini_changes.log"
}
$firstFile = $files[0].FullName
if ($AllIniData[$firstFile].Keys -contains "CONFIGEDITOR") {
    foreach ($key in $AllIniData[$firstFile]["CONFIGEDITOR"].Keys) {
        $valObj = $AllIniData[$firstFile]["CONFIGEDITOR"][$key]
        if ($valObj -and $editorSettings.ContainsKey($key)) {
            $editorSettings[$key] = $valObj.Value
        }
    }
}
$logFilePath = $editorSettings["LogFilePath"]

#############################################
# GUI erstellen: Formular, ListView und DataGridView
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
# Linke Ansicht befüllen
# Für jede INI-Datei: ein fetter Header ("INI File: ..."), dann direkt die Sections in finaler Reihenfolge
#############################################
function Populate-ListView {
    $listView.Items.Clear()

    foreach ($filePath in $AllIniData.Keys) {
        # Fetter Balken als Header mit Dateiname
        $fileItem = New-Object System.Windows.Forms.ListViewItem("INI File: " + (Split-Path $filePath -Leaf))
        $fileItem.Font = New-Object System.Drawing.Font($listView.Font.FontFamily, $listView.Font.Size, [System.Drawing.FontStyle]::Bold)
        $fileItem.Tag = [PSCustomObject]@{
            Type    = "FileHeader"
            File    = $filePath
            Section = $null
        }
        [void]$listView.Items.Add($fileItem)

        # Erstelle für diese Datei die gewünschte Reihenfolge der Sections
        $currentIni = $AllIniData[$filePath]
        $allSections = $currentIni.Keys
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

        # Sections einfügen (OHNE Separatoren)
        foreach ($sec in $finalOrder) {
            $item = New-Object System.Windows.Forms.ListViewItem($sec)
            $item.Tag = [PSCustomObject]@{
                Type    = "Section"
                File    = $filePath
                Section = $sec
            }
            [void]$listView.Items.Add($item)
        }
    }
}

Populate-ListView

#############################################
# ListView-Auswahl-Event: Beim Klick werden Datei und Section aus dem Tag-Feld ermittelt
#############################################
$currentFile = $null
$currentSection = $null

$listView.Add_SelectedIndexChanged({
    if ($listView.SelectedItems.Count -gt 0) {
        $selTag = $listView.SelectedItems[0].Tag
        if ($selTag -and $selTag.Type -eq "Section") {
            $currentFile = $selTag.File
            $currentSection = $selTag.Section

            $dataGrid.Rows.Clear()
            $thisIniData = $AllIniData[$currentFile]
            if ($thisIniData.Keys -contains $currentSection) {
                foreach ($key in $thisIniData[$currentSection].Keys) {
                    $entry = $thisIniData[$currentSection][$key]
                    $row = @($key, $entry.Value, $entry.Comment)
                    [void]$dataGrid.Rows.Add($row)
                }
            }
        }
        else {
            $currentFile = $null
            $currentSection = $null
            $dataGrid.Rows.Clear()
        }
    }
})

#############################################
# Buttons: Save, Info, Close
#############################################

# Save Changes
$saveButton = New-Object System.Windows.Forms.Button
$saveButton.Text = "Save Changes"
$saveButton.Size = New-Object System.Drawing.Size(120,40)
$saveButton.Location = New-Object System.Drawing.Point(20,780)
$saveButton.BackColor = [System.Drawing.Color]::LightGreen
$saveButton.Add_Click({
    if (-not $currentFile -or -not $currentSection) {
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
    $AllIniData[$currentFile][$currentSection] = $sectionData

    # INI-Datei schreiben (mit der originalen Reihenfolge)
    Write-IniFile -IniData $AllIniData[$currentFile] -Path $currentFile -DesiredOrder $desiredOrder

    [System.Windows.Forms.MessageBox]::Show("INI-Datei erfolgreich gespeichert.", "Gespeichert")

    Log-Change -Message "Abschnitt '$currentSection' in Datei '$currentFile' wurde gespeichert." -LogPath $logFilePath
})
$form.Controls.Add($saveButton)

# Info Button – öffnet info_editor.txt
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

# Close Button
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
# Formular anzeigen
#############################################
$form.Add_Shown({ $form.Activate() })
[void]$form.ShowDialog()

# SIG # Begin signature block
# MIIcCAYJKoZIhvcNAQcCoIIb+TCCG/UCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCATs1RDgd20Pn9j
# 5QqWPEhrKFYixtv6GFkV2ul8DlsUjqCCFk4wggMQMIIB+KADAgECAhB3jzsyX9Cg
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
# BgkqhkiG9w0BCQQxIgQg4tTSx0JHQNXFWhJwvBdepUny3lRjbSFQItSSPRM3rJcw
# DQYJKoZIhvcNAQEBBQAEggEAqGbaGNNpRO1E1w5CWffuNiKJ8DOCjjIaojeFPHfN
# wW/vI2LXjz+sFRXOWSpujWzEIJGKqMt3X9aV79XyBVyPkVf08rPaxA7oY1iL29O7
# Snbaq5SsAhbbI11l3X1Fifo69R3dOnpZ9DP5ICz7Em8tLcBfVVc7+kpkGPKPZ1AH
# M/qH5JXTygPkyLKrJcN57nFhrT6Vq9AnxGiBuYXDbIGPM1OinmfM2SN3m39iUa7F
# 24IU1Gw4Gfllgz+0e73HI++KdOKbBDJ06tioDJWeo7z0J/rwi8HY2r8UXdGW4GId
# vHKnTBuWJR/na0JxrBUka0ZPEp+lBxcQICMgJmFxU9epwaGCAyYwggMiBgkqhkiG
# 9w0BCQYxggMTMIIDDwIBATB9MGkxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdp
# Q2VydCwgSW5jLjFBMD8GA1UEAxM4RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3Rh
# bXBpbmcgUlNBNDA5NiBTSEEyNTYgMjAyNSBDQTECEAqA7xhLjfEFgtHEdqeVdGgw
# DQYJYIZIAWUDBAIBBQCgaTAYBgkqhkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqG
# SIb3DQEJBTEPFw0yNTA3MTIwODAyMzNaMC8GCSqGSIb3DQEJBDEiBCDVCwtjBz+i
# QkV7foyeldLYys/ZUhNBflWSt8kd4k/PBjANBgkqhkiG9w0BAQEFAASCAgC0pbT6
# RvYHAxEKxzFt6LOCitFGHyjnMXy94syXLwJM/mdBEa/HLg6Nvt2XddCvuBavzySC
# 3xjB3+/1vBDJl4X4mImawj+NBYfPXz+TED3Cz8AJsfFBDswXyy72lgwUhKBv0pND
# XLXK9kLsi1oBhW2sjLaRRhXpVOMDWc/pQY1jvZBrQzd5Gm3b6IZksGdp0EWDA90U
# LBooCqs4DPruyC6VXz+kM94dXJwkhFWUWMwstf1YlPJ9rGCWOQ8yIoKoOoncBrNH
# hlk9DeCCQTl1kAiR7UibWsFAcfNNFTWaFyZR+74YfwT/f8igEPg5kcUaIvwUdtLf
# 4XGRIKy7As4mTUNPkgOUAIQRp1xpTzioQPO/khGcrPYNZj9KNIfJVoeGT7jfi3W/
# aBPaPkKYFjaQkzBJPT32iVqNCNHfn94rW88aQz9OOaiptiN5b/3wXEkLxf6nMHwU
# oCUDjLnXKcyKourdhAWI8iUfscihRgQxdNgnEwH4gF5RbhHXnp+OogzYiKJC2IGu
# NC50NLBXOBqk1gB8cydF8wI74MRvrKMLINlOsZdPuHTgxt1644YK4FNRNHKSSdTy
# WlsgTGiMNxdx3dvovyA+EwLwCMqz5v1jDQnb1FXzuRMtcf7zYZmznHcYfLOY8mVj
# ZihP1kfAZrQp2Gf8Cms8+r1EHE4HDyL2lE/WRw==
# SIG # End signature block
