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
