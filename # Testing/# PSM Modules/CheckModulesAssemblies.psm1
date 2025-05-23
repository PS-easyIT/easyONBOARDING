# Funktion zum Laden aller notwendigen Assemblies
function Initialize-RequiredAssemblies {
    [CmdletBinding()]
    param()
    
    $assemblies = @(
        "PresentationFramework",
        "PresentationCore",
        "WindowsBase",
        "System.Windows.Forms",
        "System.Drawing",
        "Microsoft.VisualBasic"
    )
    
    foreach ($assembly in $assemblies) {
        try {
            Add-Type -AssemblyName $assembly -ErrorAction Stop
            Write-Verbose "Assembly $assembly erfolgreich geladen"
        }
        catch {
            Write-Error "Fehler beim Laden von Assembly $assembly: $($_.Exception.Message)"
            return $false
        }
    }
    
    # Spezifische Typen für die GUI registrieren
    try {
        # WPF-Namespace
        [System.Windows.Window] | Out-Null
        [System.Windows.Controls.Grid] | Out-Null
        [System.Windows.Controls.Button] | Out-Null
        [System.Windows.Controls.TextBox] | Out-Null
        
        Write-Verbose "Alle kritischen WPF-Typen erfolgreich registriert"
    }
    catch {
        Write-Error "Fehler beim Registrieren kritischer WPF-Typen: $($_.Exception.Message)"
        return $false
    }
    
    return $true
}

# Beim Laden des Moduls automatisch ausführen
$result = Initialize-RequiredAssemblies
if (-not $result) {
    Write-Error "KRITISCH: Notwendige Assemblies konnten nicht geladen werden. Die GUI wird nicht funktionieren."
}

# Funktion exportieren
Export-ModuleMember -Function Initialize-RequiredAssemblies