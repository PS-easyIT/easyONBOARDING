#
# DropdownHelpers.psm1
# Hilfsfunktionen zur Verwaltung von Dropdown-Menüs im easyONBOARDING-Tool
#

function Initialize-AllDropdowns {
    [CmdletBinding()]
    param()
    
    Write-DebugOutput "Initialisiere alle Dropdown-Menüs" -Level INFO
    
    $initialized = 0
    $failed = 0
    
    # Funktionen für verschiedene Dropdown-Initialisierungen
    $dropdownFunctions = @(
        "Initialize-DisplayNameTemplateDropdown",
        "Initialize-LicenseDropdown",
        "Initialize-TeamLeaderGroupDropdown",
        "Initialize-EmailDomainSuffixDropdown",
        "Initialize-LocationDropdown",
        "Initialize-OUDropdown"
    )
    
    foreach ($functionName in $dropdownFunctions) {
        if (Get-Command -Name $functionName -ErrorAction SilentlyContinue) {
            try {
                & $functionName
                $initialized++
                Write-DebugOutput "Dropdown initialisiert: $functionName" -Level DEBUG
            }
            catch {
                $failed++
                Write-DebugOutput "Fehler bei Dropdown-Initialisierung ($functionName): $($_.Exception.Message)" -Level WARNING
            }
        }
        else {
            Write-DebugOutput "Dropdown-Funktion nicht gefunden: $functionName" -Level DEBUG
        }
    }
    
    Write-DebugOutput "Dropdown-Initialisierung abgeschlossen. Erfolgreich: $initialized, Fehlgeschlagen: $failed" -Level INFO
    return ($failed -eq 0)
}

function Initialize-DisplayNameTemplateDropdown {
    [CmdletBinding()]
    param()
    
    Write-DebugOutput "Initialisiere DisplayName-Template Dropdown" -Level DEBUG
    
    $cmbDisplayTemplate = $global:window.FindName("cmbDisplayTemplate")
    if ($null -eq $cmbDisplayTemplate) {
        Write-DebugOutput "ComboBox 'cmbDisplayTemplate' nicht gefunden" -Level WARNING
        return $false
    }
    
    # Standardtemplates laden
    $templates = @(
        "Vorname.Nachname",
        "NachnameV",
        "Nachname.Vorname",
        "VornameN"
    )
    
    # Angepasste Templates aus der Konfiguration laden, wenn vorhanden
    if ($global:Config -and $global:Config.ContainsKey("UPN") -and $global:Config["UPN"].ContainsKey("Templates")) {
        $configTemplates = $global:Config["UPN"]["Templates"] -split ","
        if ($configTemplates.Count -gt 0) {
            $templates = $configTemplates
        }
    }
    
    # ComboBox leeren und befüllen
    $cmbDisplayTemplate.Items.Clear()
    foreach ($template in $templates) {
        $cmbDisplayTemplate.Items.Add($template.Trim())
    }
    
    # Standardauswahl setzen
    if ($cmbDisplayTemplate.Items.Count -gt 0) {
        $cmbDisplayTemplate.SelectedIndex = 0
    }
    
    Write-DebugOutput "DisplayName-Template Dropdown initialisiert mit $($templates.Count) Templates" -Level DEBUG
    return $true
}

function Initialize-OUDropdownRefreshEvent {
    [CmdletBinding()]
    param()
    
    Write-DebugOutput "Registriere OU-Dropdown Refresh-Event" -Level DEBUG
    
    $btnRefreshOU = $global:window.FindName("btnRefreshOU")
    if ($null -eq $btnRefreshOU) {
        Write-DebugOutput "Button 'btnRefreshOU' nicht gefunden" -Level WARNING
        return $false
    }
    
    $btnRefreshOU.Add_Click({
        Write-DebugOutput "OU-Refresh angefordert" -Level DEBUG
        if (Get-Command -Name "Load-OUDropdown" -ErrorAction SilentlyContinue) {
            Load-OUDropdown
        }
        elseif (Get-Command -Name "Initialize-OUDropdown" -ErrorAction SilentlyContinue) {
            Initialize-OUDropdown
        }
        else {
            Write-DebugOutput "Keine Funktion zum Laden des OU-Dropdowns gefunden" -Level WARNING
        }
    })
    
    Write-DebugOutput "OU-Dropdown Refresh-Event registriert" -Level DEBUG
    return $true
}

# Platzhalter für die nicht implementierten Funktionen, die in einem vollständigen Modul 
# implementiert werden sollten
function Initialize-LicenseDropdown { 
    Write-DebugOutput "LicenseDropdown-Initialisierung - Platzhalter" -Level DEBUG
    # Implementierung hier
}

function Initialize-TeamLeaderGroupDropdown {
    Write-DebugOutput "TeamLeaderGroupDropdown-Initialisierung - Platzhalter" -Level DEBUG
    # Implementierung hier
}

function Initialize-EmailDomainSuffixDropdown {
    Write-DebugOutput "EmailDomainSuffixDropdown-Initialisierung - Platzhalter" -Level DEBUG
    # Implementierung hier
}

function Initialize-LocationDropdown {
    Write-DebugOutput "LocationDropdown-Initialisierung - Platzhalter" -Level DEBUG
    # Implementierung hier
}

function Initialize-OUDropdown {
    Write-DebugOutput "OUDropdown-Initialisierung - Platzhalter" -Level DEBUG
    # Implementierung hier
}

# Exportiere alle Funktionen
Export-ModuleMember -Function Initialize-AllDropdowns, Initialize-DisplayNameTemplateDropdown, 
                              Initialize-LicenseDropdown, Initialize-TeamLeaderGroupDropdown,
                              Initialize-EmailDomainSuffixDropdown, Initialize-LocationDropdown,
                              Initialize-OUDropdown, Initialize-OUDropdownRefreshEvent
