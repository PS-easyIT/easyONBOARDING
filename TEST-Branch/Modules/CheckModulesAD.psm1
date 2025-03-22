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

# Functions for checking and validating Active Directory related functionality

function Test-ADModuleAvailable {
    [CmdletBinding()]
    param()
    
    try {
        if (Get-Command -Name "Write-DebugOutput" -ErrorAction SilentlyContinue) {
            Write-DebugOutput "Checking if ActiveDirectory module is available" -Level INFO
        } else {
            Write-Host "Checking if ActiveDirectory module is available"
        }
        
        if (Get-Module -Name ActiveDirectory -ListAvailable) {
            if (Get-Command -Name "Write-DebugOutput" -ErrorAction SilentlyContinue) {
                Write-DebugOutput "ActiveDirectory module is available" -Level INFO
            } else {
                Write-Host "ActiveDirectory module is available" -ForegroundColor Green
            }
            return $true
        } else {
            if (Get-Command -Name "Write-DebugOutput" -ErrorAction SilentlyContinue) {
                Write-DebugOutput "ActiveDirectory module is NOT available!" -Level ERROR
            } else {
                Write-Host "ActiveDirectory module is NOT available!" -ForegroundColor Red
            }
            return $false
        }
    }
    catch {
        if (Get-Command -Name "Write-DebugOutput" -ErrorAction SilentlyContinue) {
            Write-DebugOutput "Error checking ActiveDirectory module: $($_.Exception.Message)" -Level ERROR
        } else {
            Write-Host "Error checking ActiveDirectory module: $($_.Exception.Message)" -ForegroundColor Red
        }
        return $false
    }
}

function Import-ADModule {
    [CmdletBinding()]
    param()
    
    try {
        if (Get-Command -Name "Write-DebugOutput" -ErrorAction SilentlyContinue) {
            Write-DebugOutput "Importing ActiveDirectory module" -Level INFO
        } else {
            Write-Host "Importing ActiveDirectory module"
        }
        
        Import-Module -Name ActiveDirectory -ErrorAction Stop
        
        if (Get-Command -Name "Write-DebugOutput" -ErrorAction SilentlyContinue) {
            Write-DebugOutput "ActiveDirectory module imported successfully" -Level INFO
        } else {
            Write-Host "ActiveDirectory module imported successfully" -ForegroundColor Green
        }
        return $true
    }
    catch {
        if (Get-Command -Name "Write-DebugOutput" -ErrorAction SilentlyContinue) {
            Write-DebugOutput "Error importing ActiveDirectory module: $($_.Exception.Message)" -Level ERROR
        } else {
            Write-Host "Error importing ActiveDirectory module: $($_.Exception.Message)" -ForegroundColor Red
        }
        return $false
    }
}

function Test-ADConnection {
    [CmdletBinding()]
    param()
    
    try {
        if (Get-Command -Name "Write-DebugOutput" -ErrorAction SilentlyContinue) {
            Write-DebugOutput "Testing connection to Active Directory" -Level INFO
        } else {
            Write-Host "Testing connection to Active Directory"
        }
        
        # Try to get domain information
        $domain = Get-ADDomain -ErrorAction Stop
        
        if (Get-Command -Name "Write-DebugOutput" -ErrorAction SilentlyContinue) {
            Write-DebugOutput "Successfully connected to domain: $($domain.DNSRoot)" -Level INFO
        } else {
            Write-Host "Successfully connected to domain: $($domain.DNSRoot)" -ForegroundColor Green
        }
        return $true
    }
    catch {
        if (Get-Command -Name "Write-DebugOutput" -ErrorAction SilentlyContinue) {
            Write-DebugOutput "Error connecting to Active Directory: $($_.Exception.Message)" -Level ERROR
        } else {
            Write-Host "Error connecting to Active Directory: $($_.Exception.Message)" -ForegroundColor Red
        }
        return $false
    }
}

# Automatisch das AD-Modul beim Import dieses Moduls initialisieren
Initialize-ADModule

# Export module functions
Export-ModuleMember -Function Initialize-ADModule, Test-ADModuleAvailable, Import-ADModule, Test-ADConnection