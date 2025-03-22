#
# INIEditor.psm1
# Funktionen zum Bearbeiten von INI-Dateien für das easyONBOARDING-Tool
#

# Funktion zum Parsen der INI-Datei
function Get-IniContent {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$Path
    )
    
    Write-DebugOutput "Lade INI-Datei: $Path" -Level DEBUG
    
    if (-not (Test-Path -Path $Path)) {
        Write-DebugOutput "INI-Datei nicht gefunden: $Path" -Level ERROR
        return $null
    }
    
    $ini = @{}
    $section = "DEFAULT"
    $ini[$section] = @{}
    
    # Lade INI-Datei
    $content = Get-Content -Path $Path
    
    foreach ($line in $content) {
        $line = $line.Trim()
        
        # Leere Zeilen und Kommentare überspringen
        if ($line -match "^\s*$" -or $line -match "^\s*[;#]") {
            continue
        }
        
        # Sektionsname
        if ($line -match "^\[(.+)\]$") {
            $section = $matches[1]
            if (-not $ini.ContainsKey($section)) {
                $ini[$section] = @{}
            }
            continue
        }
        
        # Key-Value-Paar
        if ($line -match "^([^=]+)=(.*)$") {
            $key = $matches[1].Trim()
            $value = $matches[2].Trim()
            $ini[$section][$key] = $value
        }
    }
    
    return $ini
}

# Funktion zum Speichern der INI-Datei
function Save-IniContent {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [hashtable]$Content,
        
        [Parameter(Mandatory=$true)]
        [string]$Path
    )
    
    Write-DebugOutput "Speichere INI-Datei: $Path" -Level DEBUG
    
    $output = ""
    
    foreach ($section in $Content.Keys) {
        $output += "[$section]`r`n"
        
        foreach ($key in $Content[$section].Keys) {
            $value = $Content[$section][$key]
            $output += "$key=$value`r`n"
        }
        
        $output += "`r`n"
    }
    
    try {
        $output | Out-File -FilePath $Path -Encoding UTF8
        return $true
    }
    catch {
        Write-DebugOutput "Fehler beim Speichern der INI-Datei: $($_.Exception.Message)" -Level ERROR
        return $false
    }
}

# Funktion zum Laden der INI-Sektionen in die ListView
function Import-INIEditorData {
    [CmdletBinding()]
    param()
    
    Write-DebugOutput "Initialisiere INI-Editor-Daten" -Level INFO
    
    try {
        $listViewINIEditor = $global:window.FindName("listViewINIEditor")
        if ($null -eq $listViewINIEditor) {
            Write-DebugOutput "ListView 'listViewINIEditor' nicht gefunden" -Level ERROR
            return $false
        }
        
        # ListView leeren
        $listViewINIEditor.Items.Clear()
        
        # Sektionen laden
        if ($null -eq $global:Config) {
            $global:Config = Get-IniContent -Path $global:INIPath
        }
        
        # Sektionen hinzufügen
        foreach ($section in $global:Config.Keys) {
            $item = New-Object PSObject -Property @{
                SectionName = $section
                KeyCount = $global:Config[$section].Count
            }
            
            $listViewINIEditor.Items.Add($item)
        }
        
        Write-DebugOutput "INI-Editordaten geladen: $($global:Config.Count) Sektionen gefunden" -Level INFO
        return $true
    }
    catch {
        Write-DebugOutput "Fehler beim Laden der INI-Editordaten: $($_.Exception.Message)" -Level ERROR
        return $false
    }
}

# Funktion zum Laden einer INI-Sektion in das DataGrid
function Import-SectionSettings {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$SectionName,
        
        [Parameter(Mandatory=$true)]
        [System.Windows.Controls.DataGrid]$DataGrid,
        
        [Parameter(Mandatory=$true)]
        [string]$INIPath
    )
    
    Write-DebugOutput "Lade Sektion: $SectionName" -Level DEBUG
    
    try {
        if ($null -eq $global:Config) {
            $global:Config = Get-IniContent -Path $INIPath
        }
        
        if (-not $global:Config.ContainsKey($SectionName)) {
            Write-DebugOutput "Sektion nicht gefunden: $SectionName" -Level WARNING
            return $false
        }
        
        # DataGrid leeren
        $DataGrid.Items.Clear()
        
        # Einträge hinzufügen
        foreach ($key in $global:Config[$SectionName].Keys) {
            $value = $global:Config[$SectionName][$key]
            
            $item = New-Object PSObject -Property @{
                Key = $key
                Value = $value
                OriginalKey = $key  # Für Tracking von Änderungen
            }
            
            $DataGrid.Items.Add($item)
        }
        
        Write-DebugOutput "Sektionseinstellungen geladen: $($global:Config[$SectionName].Count) Einträge" -Level INFO
        return $true
    }
    catch {
        Write-DebugOutput "Fehler beim Laden der Sektionseinstellungen: $($_.Exception.Message)" -Level ERROR
        return $false
    }
}

# Funktion zum Speichern der INI-Änderungen
function Save-INIChanges {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$INIPath
    )
    
    Write-DebugOutput "Speichere INI-Änderungen in: $INIPath" -Level INFO
    
    try {
        if ($null -eq $global:Config) {
            Write-DebugOutput "Keine Konfiguration zum Speichern vorhanden" -Level ERROR
            return $false
        }
        
        $result = Save-IniContent -Content $global:Config -Path $INIPath
        
        if ($result) {
            Write-DebugOutput "INI-Datei erfolgreich gespeichert" -Level INFO
        }
        else {
            Write-DebugOutput "Fehler beim Speichern der INI-Datei" -Level ERROR
        }
        
        return $result
    }
    catch {
        Write-DebugOutput "Fehler beim Speichern der INI-Änderungen: $($_.Exception.Message)" -Level ERROR
        return $false
    }
}

# Exportiere alle Funktionen
Export-ModuleMember -Function Get-IniContent, Save-IniContent, Import-INIEditorData, Import-SectionSettings, Save-INIChanges
