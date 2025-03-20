# Definiert eine Funktion zum Laden des Active Directory Moduls
function Initialize-ADModule {
    [CmdletBinding()]
    param()

    # Überprüft, ob die erforderlichen Logging-Funktionen existieren
    $requiredFunctions = @('Write-DebugMessage', 'Write-Log')
    foreach ($function in $requiredFunctions) {
        if (-not (Get-Command -Name $function -ErrorAction SilentlyContinue)) {
            Write-Warning "Die Funktion '$function' ist nicht verfügbar. Stellen Sie sicher, dass das entsprechende Modul geladen wurde."
        }
    }

    # Loads the Active Directory PowerShell module required for user management
    Write-DebugMessage "Loading Active Directory module."

    try {
        Write-DebugMessage "Checking if Active Directory module exists."
        if (Get-Module -Name ActiveDirectory -ListAvailable) {
            Write-DebugMessage "Loading AD module"
            Import-Module ActiveDirectory -SkipEditionCheck -ErrorAction Stop
            Write-DebugMessage "Active Directory module loaded successfully."
            return $true
        } else {
            Write-Log -Message "ActiveDirectory module is not installed." -LogLevel "ERROR"
            Throw "Critical error: ActiveDirectory module is not installed!"
        }
    } catch {
        Write-Log -Message "ActiveDirectory module could not be loaded: $($_.Exception.Message)" -LogLevel "ERROR"
        Throw "Critical error: $($_.Exception.Message)"
    }
}

# Automatisch das AD-Modul beim Import dieses Moduls initialisieren
Initialize-ADModule

# Export module functions
Export-ModuleMember -Function Initialize-ADModule