# Module zur Überprüfung und Verwaltung von Administratorrechten

# Überprüft, ob das Skript mit Administratorrechten ausgeführt wird
function Test-Admin {
    [CmdletBinding()]
    param()
    
    # Prüfen, ob Write-DebugMessage verfügbar ist
    if (Get-Command -Name "Write-DebugMessage" -ErrorAction SilentlyContinue) {
        Write-DebugMessage "Überprüfe Administratorrechte" -LogLevel "INFO"
    }
    
    $currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    $isAdmin = $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    
    if (Get-Command -Name "Write-DebugMessage" -ErrorAction SilentlyContinue) {
        if ($isAdmin) {
            Write-DebugMessage "Skript wird mit Administratorrechten ausgeführt" -LogLevel "INFO"
        } else {
            Write-DebugMessage "Skript wird OHNE Administratorrechte ausgeführt" -LogLevel "WARNING"
        }
    }
    
    return $isAdmin
}

# Startet das aktuelle Skript mit Administratorrechten neu
function Start-AsAdmin {
    [CmdletBinding()]
    param()
    
    if (Get-Command -Name "Write-DebugMessage" -ErrorAction SilentlyContinue) {
        Write-DebugMessage "Starte Skript mit Administratorrechten neu" -LogLevel "INFO"
    }
    
    try {
        $scriptPath = $MyInvocation.PSCommandPath
        if ([string]::IsNullOrEmpty($scriptPath)) {
            $scriptPath = $script:MyInvocation.MyCommand.Path
        }
        
        if ([string]::IsNullOrEmpty($scriptPath)) {
            throw "Konnte Skriptpfad nicht ermitteln"
        }
        
        $arguments = "-NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`""
        
        if (Get-Command -Name "Write-DebugMessage" -ErrorAction SilentlyContinue) {
            Write-DebugMessage "Starte PowerShell mit folgenden Argumenten: $arguments" -LogLevel "DEBUG"
        }
        
        Start-Process -FilePath PowerShell.exe -ArgumentList $arguments -Verb RunAs
        
        if (Get-Command -Name "Write-DebugMessage" -ErrorAction SilentlyContinue) {
            Write-DebugMessage "Prozess mit Rechteerweiterung gestartet" -LogLevel "INFO"
        }
        
        return $true
    }
    catch {
        if (Get-Command -Name "Write-DebugMessage" -ErrorAction SilentlyContinue) {
            Write-DebugMessage "Fehler beim Neustart mit Administratorrechten: $($_.Exception.Message)" -LogLevel "ERROR"
        }
        
        return $false
    }
}

# Check for admin privileges
if (-not (Test-Admin)) {
    # If not running as admin, show a dialog to restart with elevated permissions
    Add-Type -AssemblyName PresentationFramework
    $result = [System.Windows.MessageBox]::Show(
        "This script requires administrator privileges to run properly.`n`nDo you want to restart with elevated permissions?",
        "Administrator Rights Required",
        [System.Windows.MessageBoxButton]::YesNo,
        [System.Windows.MessageBoxImage]::Warning
    )
    
    if ($result -eq [System.Windows.MessageBoxResult]::Yes) {
        if (Start-AsAdmin) {
            # Exit this instance as we've started a new elevated instance
            exit
        }
    }
    else {
        Write-Warning "The script will continue without administrator privileges. Some functions may not work properly."
    }
}

# Check PowerShell version
if ($PSVersionTable.PSVersion.Major -lt 7) {
    Add-Type -AssemblyName PresentationFramework
    [System.Windows.MessageBox]::Show(
        "This script requires PowerShell 7 or higher.`n`nCurrent version: $($PSVersionTable.PSVersion)",
        "Version Error",
        [System.Windows.MessageBoxButton]::OK,
        [System.Windows.MessageBoxImage]::Error
    )
    exit
}

# Define and make the script directory available to importing scripts
$script:ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
#endregion

# Export module members
Export-ModuleMember -Function Test-Admin, Start-AsAdmin