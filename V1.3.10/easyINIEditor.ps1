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
