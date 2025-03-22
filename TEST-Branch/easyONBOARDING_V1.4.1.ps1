<#
  This script integrates functions for onboarding, AD group creation and the INI editor in a WPF interface.
  
  Requirements: Administrative rights and PowerShell 7 or higher
    - INI file     ("easyONBOARDINGConfig.ini")
    - XAML file    ("MainGUI.xaml")
#>

# [Region 00 | SCRIPT INITIALIZATION]
# Sets up error handling and basic script environment
#requires -Version 7.0

$ErrorActionPreference = "Stop"
trap {
    $errorMessage = "ERROR: Unhandled error occurred! " +
                    "Error message: $($_.Exception.Message); " +
                    "Position: $($_.InvocationInfo.PositionMessage); " +
                    "StackTrace: $($_.ScriptStackTrace)"
    if ($null -ne $global:Config -and $null -ne $global:Config.Logging -and $global:Config.Logging.DebugMode -eq "1") {
        Write-Error $errorMessage
    } else {
        Write-Error "A critical error has occurred. Please check the log file."
    }
    exit 1
}

#region [Region 01 | ADMIN AND VERSION CHECK]
# Verifies administrator rights and PowerShell version before proceeding
function Test-Admin {
    $currentUser = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    return $currentUser.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Start-AsAdmin {
    $scriptPath = $MyInvocation.MyCommand.Definition
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = "pwsh.exe"
    $psi.Arguments = "-ExecutionPolicy Bypass -File `"$scriptPath`""
    $psi.Verb = "runas" # This triggers the UAC prompt
    try {
        [System.Diagnostics.Process]::Start($psi) | Out-Null
    }
    catch {
        Write-Warning "Failed to restart as administrator: $($_.Exception.Message)"
        return $false
    }
    return $true
}

# Check for admin privileges
if (-not (Test-Admin)) {
    # If not running as admin, show a dialog to restart with elevated permissions
    Add-Type -AssemblyName PresentationFramework
    $result = [System.Windows.MessageBox]::Show(
        "This script requires administrator privileges to run properly.`n`nDo you want to restart with elevated permissions?",
        "Administrator Rights Required",
        "YesNo",
        "Warning"
    )
    
    if ($result -eq "Yes") {
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
        "OK",
        "Error"
    )
    exit
}

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
#endregion

#region [Region 02 | INI FILE LOADING]
# Loads configuration data from the external INI file
function Get-IniContent {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Path
    )
    
    # [2.1 | VALIDATE INI FILE]
    # Checks if the INI file exists and is accessible
    if (-not (Test-Path $Path)) {
        Throw "INI file not found: $Path"
    }
    
    $ini = [ordered]@{}
    $currentSection = "Global"
    $ini[$currentSection] = [ordered]@{}
    
    try {
        $lines = Get-Content -Path $Path -ErrorAction Stop
    } catch {
        Throw "Error reading INI file: $($_.Exception.Message)"
    }
    
    # [2.2 | PARSE INI CONTENT]
    # Processes the INI file line by line to extract sections and key-value pairs
    foreach ($line in $lines) {
        $line = $line.Trim()
        # Skip empty lines and comment lines (starting with ; or #)
        if ([string]::IsNullOrWhiteSpace($line) -or $line -match '^\s*[;#]') { continue }
    
        if ($line -match '^\[(.+)\]$') {
            $currentSection = $matches[1].Trim()
            if (-not $ini.Contains($currentSection)) {
                $ini[$currentSection] = [ordered]@{}
            }
        } elseif ($line -match '^(.*?)=(.*)$') {
            $key = $matches[1].Trim()
            $value = $matches[2].Trim()
            $ini[$currentSection][$key] = $value
        }
    }
    return $ini
}

$INIPath = Join-Path $ScriptDir "easyONB.ini"
try {
    $global:Config = Get-IniContent -Path $INIPath
}
catch {
    Write-Host "Error loading INI file: $_"
    exit
}
#endregion

#region [Region 03 | FUNCTION DEFINITIONS]
# Contains all helper and core functionality functions

#region [Region 03.1 | LOGGING FUNCTIONS]
# Defines logging capabilities for different message levels
function Write-LogMessage {
    # [03.1.1 - Primary logging wrapper for consistent message formatting]
    param(
        [string]$message,
        [string]$logLevel = "INFO"
    )
    Write-Log -message $message -logLevel $logLevel
}

function Write-DebugMessage {
    # [03.1.2 - Debug-specific logging with conditional execution]
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true, ValueFromPipeline=$true)]
        [AllowEmptyString()]
        [string]$Message
    )
    process {
        # Only log debug messages if DebugMode is enabled in config
        if ($null -ne $global:Config -and 
            $null -ne $global:Config.Logging -and 
            $global:Config.Logging.DebugMode -eq "1") {
            
            # Call Write-Log without capturing output to avoid pipeline return
            Write-Log -Message $Message -LogLevel "DEBUG"
            
            # Also output to console for immediate feedback during debugging
            Write-Host "[DEBUG] $Message" -ForegroundColor Cyan
        }
        
        # No return value to avoid unwanted pipeline output
    }
}
#endregion

# Log level (default is "INFO"): "WARN", "ERROR", "DEBUG".

#region [Region 03.2 | LOG FILE WRITER] 
# Core logging function that writes messages to log files
function Write-Log {
    # [03.2.1 - Low-level file logging implementation]
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true, ValueFromPipeline=$true)]
        [AllowEmptyString()]
        [string]$Message,
        
        [Parameter(Mandatory=$false)]
        [ValidateSet("INFO", "WARN", "ERROR", "DEBUG")]
        [string]$LogLevel = "INFO"
    )
    process {
        # log path determination using null-conditional and coalescing operators
        $logFilePath = $global:Config?.Logging?.ExtraFile ?? 
                      $global:Config?.Logging?.LogFile ?? 
                      (Join-Path $ScriptDir "Logs")
        
        # Ensure the log directory exists
        if (-not (Test-Path $logFilePath)) {
            try {
                # Use -Force to create parent directories if needed
                [void](New-Item -ItemType Directory -Path $logFilePath -Force -ErrorAction Stop)
                Write-Host "Created log directory: $logFilePath"
            } catch {
                # Handle directory creation errors gracefully
                Write-Warning "Error creating log directory: $($_.Exception.Message)"
                # Try writing to a fallback location if possible
                $logFilePath = $env:TEMP
            }
        }

        # Define log file paths
        $logFile = Join-Path -Path $logFilePath -ChildPath "easyOnboarding.log"
        $errorLogFile = Join-Path -Path $logFilePath -ChildPath "easyOnboarding_error.log"
        
        # Generate timestamp and log entry
        $timeStamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        $logEntry = "$timeStamp [$LogLevel] $Message"

        try {
            Add-Content -Path $logFile -Value $logEntry -ErrorAction Stop
        } catch {
            # Log additional error if writing to log file fails
            try {
                Add-Content -Path $errorLogFile -Value "$timeStamp [ERROR] Error writing to log file: $($_.Exception.Message)" -ErrorAction SilentlyContinue
            } catch {
                # Last resort: output to console
                Write-Warning "Critical error: Cannot write to any log file: $($_.Exception.Message)"
            }
        }

        # Output to console based on DebugMode setting
        if ($null -ne $global:Config -and $null -ne $global:Config.Logging) {
            # When DebugMode=1, output all messages regardless of log level
            if ($global:Config.Logging.DebugMode -eq "1") {
                # Use color coding based on log level
                switch ($LogLevel) {
                    "DEBUG" { Write-Host "$logEntry" -ForegroundColor Cyan }
                    "INFO"  { Write-Host "$logEntry" -ForegroundColor White }
                    "WARN"  { Write-Host "$logEntry" -ForegroundColor Yellow }
                    "ERROR" { Write-Host "$logEntry" -ForegroundColor Red }
                    default { Write-Host "$logEntry" }
                }
            }
            # When DebugMode=0, no console output
        }
        
        # No return value to avoid unwanted pipeline output
    }
}
#endregion

Write-DebugMessage "INI, LOG, Module, etc. regions initialized."

#region [Region 04 | ACTIVE DIRECTORY MODULE]
# Loads the Active Directory PowerShell module required for user management
Write-DebugMessage "Loading Active Directory module."
try {
    Write-DebugMessage "Loading AD module"
    Import-Module ActiveDirectory -SkipEditionCheck -ErrorAction Stop
    Write-DebugMessage "Active Directory module loaded successfully."
} catch {
    Write-Log -Message "ActiveDirectory module could not be loaded: $($_.Exception.Message)" -LogLevel "ERROR"
    Throw "Critical error: ActiveDirectory module missing!"
}
#endregion

#region [Region 05 | WPF ASSEMBLIES]
# Loads required WPF assemblies for the GUI interface
Write-DebugMessage "Loading WPF assemblies."
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Windows.Forms
Write-DebugMessage "WPF assemblies loaded successfully."
#endregion

Write-DebugMessage "Determining XAML file path."

#region [Region 06 | XAML LOADING]
# Determines XAML file path and loads the GUI definition
Write-DebugMessage "Loading XAML file: $xamlPath"
$xamlPath = Join-Path $ScriptDir "MainGUI.xaml"
if (-not (Test-Path $xamlPath)) {
    Write-Error "XAML file not found: $xamlPath"
    exit
}
try {
    [xml]$xaml = Get-Content -Path $xamlPath
    $reader = New-Object System.Xml.XmlNodeReader $xaml
    $window = [Windows.Markup.XamlReader]::Load($reader)
    if (-not $window) {
        Throw "XAML could not be loaded."
    }
    Write-DebugMessage "XAML file loaded successfully."
    if ($global:Config.Logging.DebugMode -eq "1") {
        Write-Log "XAML file successfully loaded." "DEBUG"
    }
}
catch {
    Write-Error "Error loading XAML file. Please check the file content. $_"
    exit
}
#endregion

Write-DebugMessage "XAML file loaded"

#region [Region 07 | PASSWORD MANAGEMENT]
# Functions for setting and removing password change restrictions

#region [Region 07.1 | PREVENT PASSWORD CHANGE]
# Sets ACL restrictions to prevent users from changing their passwords
Write-DebugMessage "Defining Set-CannotChangePassword function."
function Set-CannotChangePassword {
    # [07.1.1 - Modifies ACL settings to deny password change permissions]
    param(
        [Parameter(Mandatory=$true)]
        [string]$SamAccountName
    )

    try {
        $adUser = Get-ADUser -Identity $SamAccountName -Properties DistinguishedName -ErrorAction Stop
        if (-not $adUser) {
            Write-Warning "User $SamAccountName not found."
            return
        }

        $user = [ADSI]"LDAP://$($adUser.DistinguishedName)"
        $acl = $user.psbase.ObjectSecurity

        Write-DebugMessage "Set-CannotChangePassword: Defining AccessRule"
        # Define AccessRule: SELF is not allowed to 'Change Password'
        $denyRule = New-Object System.DirectoryServices.ActiveDirectoryAccessRule(
            [System.Security.Principal.NTAccount]"NT AUTHORITY\\SELF",
            [System.DirectoryServices.ActiveDirectoryRights]::WriteProperty,
            [System.Security.AccessControl.AccessControlType]::Deny,
            [GUID]"ab721a53-1e2f-11d0-9819-00aa0040529b"  # GUID for 'User-Change-Password'
        )
        $acl.AddAccessRule($denyRule)
        $user.psbase.ObjectSecurity = $acl
        $user.psbase.CommitChanges()
        Write-Log "Prevent Password Change has been set for $SamAccountName." "DEBUG"
    }
    catch {
        Write-Warning "Error setting password change restriction"
    }
}
#endregion

Write-Host "easyONBOARDING v1.4.1 wird initialisiert..." -ForegroundColor Cyan
Write-Host "Skriptverzeichnis: $ScriptDir"
Write-Host "Modulverzeichnis: $ModulesDir"
Write-Host "INI-Dateipfad: $global:INIPath"
Write-Host "XAML-Dateipfad: $global:XAMLPath"
Write-Host "Protokollverzeichnis: $global:LogsPath"
Write-Host "Berichtsverzeichnis: $global:ReportsPath"

#region Critical Modules Verification and Loading
# Verbesserte Modul-Überprüfung mit Try-Catch
try {
    # Definition der kritischen Module
    $criticalModules = @(
        @{Name = "CoreEASYONBOARDING"; Critical = $true; Dependencies = @(
            "FunctionLogDebug", "CheckModulesAD", "FunctionADCommand", "FunctionUPNCreate", 
            "FunctionADUserCreate", "FunctionPassword")}
    )

    # Verbesserte Funktionen verwenden, wenn verfügbar
    if (Get-Command -Name "Import-RequiredModule" -ErrorAction SilentlyContinue) {
        Write-DebugOutput "Überprüfe und lade kritische Module mit verbesserter Methode..." -Level INFO
        
        $missingCritical = $false
        foreach ($module in $criticalModules) {
            $result = Import-RequiredModule -ModuleName $module.Name -Dependencies $module.Dependencies
            
            if (-not $result -and $module.Critical) {
                Write-DebugOutput "KRITISCHES MODUL FEHLT: $($module.Name)" -Level ERROR
                $missingCritical = $true
            }
        }
        
        if ($missingCritical) {
            $errorMsg = "Ein oder mehrere kritische Module fehlen. Die Anwendung kann ohne sie nicht ausgeführt werden."
            Write-DebugOutput $errorMsg -Level ERROR
            
            $continue = Read-Host "Möchtest du trotzdem fortfahren? (j/n)"
            if ($continue -ne "j") {
                exit
            }
            Write-DebugOutput "Fahre trotz fehlender Module fort. Fehler werden wahrscheinlich auftreten." -Level WARNING
        }
    }
    else {
        # Verwende Original-Funktionen aus dem Skript
        Write-Host "Überprüfe kritische Module mit Original-Methode..." -ForegroundColor Yellow
        
        # Original-Funktion importieren
        # ...existing code...
        
        # Module prüfen
        $missingCritical = $false
        Write-Host "Überprüfe kritische Module..." -ForegroundColor Yellow

        foreach ($module in $criticalModules) {
            $modulePath = Join-Path -Path $ModulesDir -ChildPath "$($module.Name).psm1"
            $exists = Test-Path $modulePath
            
            if (-not $exists -and $module.Critical) {
                Write-Host "KRITISCHES MODUL FEHLT: $($module.Name)" -ForegroundColor Red -BackgroundColor Black
                $missingCritical = $true
            }
        }

        if ($missingCritical) {
            Write-Host "Ein oder mehrere kritische Module fehlen. Die Anwendung kann ohne sie nicht ausgeführt werden." -ForegroundColor Red
            Write-Host "Stelle sicher, dass alle Module in diesem Verzeichnis vorhanden sind: $ModulesDir" -ForegroundColor Yellow
            
            $continue = Read-Host "Möchtest du trotzdem fortfahren? (j/n)"
            if ($continue -ne "j") {
                exit
            }
            Write-Host "Fahre trotz fehlender Module fort. Fehler werden wahrscheinlich auftreten." -ForegroundColor Red
        }
        
        # Module laden
        foreach ($module in $criticalModules) {
            Write-Host "Lade Modul: $($module.Name)" -ForegroundColor Cyan
            
            # Erst Abhängigkeiten laden
            foreach ($dependency in $module.Dependencies) {
                $dependencyPath = Join-Path -Path $ModulesDir -ChildPath "$dependency.psm1"
                if (Test-Path $dependencyPath) {
                    try {
                        Import-Module $dependencyPath -Force -Global -ErrorAction Stop
                        Write-Host "Abhängigkeitsmodul geladen: $dependency" -ForegroundColor Green
                    }
                    catch {
                        $errorMsg = "Fehler beim Laden des Abhängigkeitsmoduls '$dependency': $($_.Exception.Message)"
                        Write-Host $errorMsg -ForegroundColor Red
                        
                        if ($module.Critical) {
                            Write-Host "Kritische Abhängigkeit konnte nicht geladen werden. Das Skript kann nicht fortgesetzt werden." -ForegroundColor Red -BackgroundColor Black
                            exit
                        }
                    }
                }
                else {
                    Write-Host "Abhängigkeitsmodul nicht gefunden: $dependency" -ForegroundColor Red
                    
                    if ($module.Critical) {
                        Write-Host "Kritische Abhängigkeit nicht gefunden. Das Skript kann nicht fortgesetzt werden." -ForegroundColor Red -BackgroundColor Black
                        exit
                    }
                }
            }
            
            # Dann das Modul selbst laden
            $modulePath = Join-Path -Path $ModulesDir -ChildPath "$($module.Name).psm1"
            if (Test-Path $modulePath) {
                try {
                    Import-Module $modulePath -Force -Global -ErrorAction Stop
                    Write-Host "Modul geladen: $($module.Name)" -ForegroundColor Green
                }
                catch {
                    $errorMsg = "Fehler beim Laden des Moduls '$($module.Name)': $($_.Exception.Message)"
                    Write-Host $errorMsg -ForegroundColor Red
                    
                    if ($module.Critical) {
                        Write-Host "Kritisches Modul konnte nicht geladen werden. Das Skript kann nicht fortgesetzt werden." -ForegroundColor Red -BackgroundColor Black
                        exit
                    }
                }
            }
            else {
                Write-Host "Modul nicht gefunden: $($module.Name)" -ForegroundColor Red
                
                if ($module.Critical) {
                    Write-Host "Kritisches Modul nicht gefunden. Das Skript kann nicht fortgesetzt werden." -ForegroundColor Red -BackgroundColor Black
                    exit
                }
            }
        }
    }
}
catch {
    Write-Host "Kritischer Fehler bei der Modulverarbeitung: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}
#endregion

#region Load Additional Modules
# Lade zusätzliche nicht-kritische Module
try {
    # Liste zusätzlicher Module
    $additionalModules = @(
        @{Name = "GUIEventHandler"; Critical = $false; Dependencies = @("FunctionLogDebug", "LoadConfigXAML")}
        @{Name = "FunctionUPNCreate"; Critical = $false; Dependencies = @("FunctionLogDebug", "CheckModulesAD")}
        @{Name = "FunctionUPNTemplate"; Critical = $false; Dependencies = @("FunctionLogDebug")}
        @{Name = "FunctionADUserCreate"; Critical = $false; Dependencies = @("FunctionLogDebug", "CheckModulesAD", "FunctionADCommand")}
        @{Name = "FunctionDropDowns"; Critical = $false; Dependencies = @("FunctionLogDebug")}
        @{Name = "GUIDropDowns"; Critical = $false; Dependencies = @("FunctionLogDebug", "FunctionDropDowns")}
        @{Name = "GUIADGroup"; Critical = $false; Dependencies = @("FunctionLogDebug", "CheckModulesAD")}
        @{Name = "GUIUPNTemplateInfo"; Critical = $false; Dependencies = @("FunctionLogDebug", "FunctionUPNTemplate")}
        @{Name = "FindVisualChild"; Critical = $false; Dependencies = @("FunctionLogDebug")}
        @{Name = "FunctionADDataLoad"; Critical = $false; Dependencies = @("FunctionLogDebug", "CheckModulesAD", "FunctionADCommand")}
        @{Name = "FunctionSettingsImportINI"; Critical = $false; Dependencies = @("FunctionLogDebug", "LoadConfigINI")}
        @{Name = "FunctionSettingsLoadINI"; Critical = $false; Dependencies = @("FunctionLogDebug", "LoadConfigINI")}
        @{Name = "FunctionSettingsSaveINI"; Critical = $false; Dependencies = @("FunctionLogDebug", "LoadConfigINI")}
        @{Name = "FunctionPassword"; Critical = $false; Dependencies = @("FunctionLogDebug", "CheckModulesAD")}
        @{Name = "FunctionSetLogoSetPassword"; Critical = $false; Dependencies = @("FunctionLogDebug")}
        @{Name = "INIEditor"; Critical = $false; Dependencies = @("FunctionLogDebug")}
        @{Name = "DropdownHelpers"; Critical = $false; Dependencies = @("FunctionLogDebug")}
    )

    # Verbesserte Funktionen verwenden
    if (Get-Command -Name "Import-RequiredModule" -ErrorAction SilentlyContinue) {
        Write-DebugOutput "Lade zusätzliche Module mit verbesserter Methode..." -Level INFO
        
        # Parallelisierung für nicht-kritische Module verwenden
        $jobs = @()
        
        foreach ($module in $additionalModules) {
            # Kritische Module direkt im Hauptthread laden
            if ($module.Critical) {
                $result = Import-RequiredModule -ModuleName $module.Name -Critical -Dependencies $module.Dependencies
                
                if (-not $result) {
                    Write-DebugOutput "Fehler beim Laden des kritischen Moduls: $($module.Name)" -Level ERROR
                }
            }
            else {
                # Nicht-kritische Module parallel laden
                $jobs += Start-Job -ScriptBlock {
                    param($moduleName, $moduleDir, $dependencies)
                    
                    $modulePath = Join-Path -Path $moduleDir -ChildPath "$moduleName.psm1"
                    
                    if (Test-Path $modulePath) {
                        try {
                            Import-Module $modulePath -Force -Global
                            return @{Success = $true; Module = $moduleName}
                        }
                        catch {
                            return @{Success = $false; Module = $moduleName; Error = $_.Exception.Message}
                        }
                    }
                    else {
                        return @{Success = $false; Module = $moduleName; Error = "Modul nicht gefunden"}
                    }
                } -ArgumentList $module.Name, $ModulesDir, $module.Dependencies
            }
        }
        
        # Warte auf die Hintergrundaufgaben und verarbeite die Ergebnisse
        if ($jobs.Count -gt 0) {
            Write-DebugOutput "Warte auf das Laden von $($jobs.Count) Modulen im Hintergrund..." -Level INFO
            
            $timeout = 30  # Timeout in Sekunden
            $completed = Wait-Job -Job $jobs -Timeout $timeout
            
            foreach ($job in $completed) {
                $result = Receive-Job -Job $job
                
                if ($result.Success) {
                    Write-DebugOutput "Modul erfolgreich geladen: $($result.Module)" -Level INFO
                }
                else {
                    Write-DebugOutput "Fehler beim Laden des Moduls $($result.Module): $($result.Error)" -Level WARNING
                    
                    # Überprüfen, ob es sich um das INIEditor-Modul handelt
                    if ($result.Module -eq "INIEditor" -and $result.Error -match "Modul nicht gefunden") {
                        Write-DebugOutput "INIEditor-Modul nicht gefunden - erstelle es automatisch" -Level WARNING
                        
                        try {
                            # Erstelle einen minimalen INIEditor mit den notwendigen Funktionen
                            New-INIEditorModule
                        }
                        catch {
                            Write-DebugOutput "Fehler beim Erstellen des INIEditor-Moduls: $($_.Exception.Message)" -Level ERROR
                        }
                    }
                    
                    # Überprüfen, ob es sich um das DropdownHelpers-Modul handelt
                    if ($result.Module -eq "DropdownHelpers" -and $result.Error -match "Modul nicht gefunden") {
                        Write-DebugOutput "DropdownHelpers-Modul nicht gefunden - erstelle es automatisch" -Level WARNING
                        
                        try {
                            # Erstelle einen minimalen DropdownHelper mit den notwendigen Funktionen
                            New-DropdownHelpersModule
                        }
                        catch {
                            Write-DebugOutput "Fehler beim Erstellen des DropdownHelpers-Moduls: $($_.Exception.Message)" -Level ERROR
                        }
                    }
                }
                
                Remove-Job -Job $job
            }
            
            # Zeitüberschreitungen behandeln
            $remaining = $jobs | Where-Object { $_.State -ne "Completed" }
            if ($remaining) {
                Write-DebugOutput "Zeitüberschreitung beim Laden von $($remaining.Count) Modulen" -Level WARNING
                $remaining | Stop-Job
                $remaining | Remove-Job
            }
        }
    }
    else {
        # Original-Methode verwenden
        Write-Host "Lade zusätzliche Module mit Original-Methode..." -ForegroundColor Yellow
        
        foreach ($module in $additionalModules) {
            $modulePath = Join-Path -Path $ModulesDir -ChildPath "$($module.Name).psm1"
            if (Test-Path $modulePath) {
                try {
                    # Erst Abhängigkeiten prüfen
                    $dependenciesOk = $true
                    foreach ($dependency in $module.Dependencies) {
                        if (-not (Get-Module -Name $dependency -ErrorAction SilentlyContinue)) {
                            $dependencyPath = Join-Path -Path $ModulesDir -ChildPath "$dependency.psm1"
                            if (Test-Path $dependencyPath) {
                                try {
                                    Import-Module $dependencyPath -Force -Global
                                }
                                catch {
                                    Write-Host "Fehler beim Laden der Abhängigkeit $dependency für Modul $($module.Name): $($_.Exception.Message)" -ForegroundColor Yellow
                                    $dependenciesOk = $false
                                }
                            }
                            else {
                                Write-Host "Abhängigkeit $dependency für Modul $($module.Name) nicht gefunden." -ForegroundColor Yellow
                                $dependenciesOk = $false
                            }
                        }
                    }
                    
                    if ($dependenciesOk) {
                        Import-Module $modulePath -Force -Global
                        Write-Host "Modul geladen: $($module.Name)" -ForegroundColor Green
                    }
                    else {
                        Write-Host "Modul $($module.Name) wurde wegen fehlender Abhängigkeiten nicht geladen." -ForegroundColor Yellow
                    }
                }
                catch {
                    Write-Host "Fehler beim Laden des Moduls $($module.Name): $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else {
                Write-Host "Modul nicht gefunden: $($module.Name)" -ForegroundColor Yellow
            }
        }
    }
    
    # Prüfung, ob notwendige INI-Editor-Funktionen verfügbar sind 
    if (-not (Get-Command -Name "Import-INIEditorData" -ErrorAction SilentlyContinue)) {
        Write-DebugOutput "INI-Editor-Funktionen nicht gefunden. Erstelle Hilfsfunktionen." -Level WARNING
        
        # Definiere die notwendigen Funktionen
        function Import-INIEditorData {
            [CmdletBinding()]
            param()
            
            Write-DebugOutput "Notfall-INI-Daten-Import" -Level WARNING
            
            $listViewINIEditor = $global:window.FindName("listViewINIEditor")
            if ($null -eq $listViewINIEditor) {
                Write-DebugOutput "ListView 'listViewINIEditor' nicht gefunden" -Level ERROR
                return $false
            }
            
            $listViewINIEditor.Items.Clear()
            
            # Sektionen aus Config hinzufügen
            foreach ($section in $global:Config.Keys) {
                $item = New-Object PSObject -Property @{
                    SectionName = $section
                    KeyCount = $global:Config[$section].Count
                }
                
                $listViewINIEditor.Items.Add($item)
            }
            
            return $true
        }
        
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
            
            Write-DebugOutput "Notfall-Sektion-Import: $SectionName" -Level WARNING
            
            $DataGrid.Items.Clear()
            
            # Einträge hinzufügen
            foreach ($key in $global:Config[$SectionName].Keys) {
                $value = $global:Config[$SectionName][$key]
                
                $item = New-Object PSObject -Property @{
                    Key = $key
                    Value = $value
                    OriginalKey = $key
                }
                
                $DataGrid.Items.Add($item)
            }
            
            return $true
        }
        
        function Save-INIChanges {
            [CmdletBinding()]
            param(
                [Parameter(Mandatory=$true)]
                [string]$INIPath
            )
            
            Write-DebugOutput "Notfall-INI-Speicherung: $INIPath" -Level WARNING
            
            try {
                $output = ""
                
                foreach ($section in $global:Config.Keys) {
                    $output += "[$section]`r`n"
                    
                    foreach ($key in $global:Config[$section].Keys) {
                        $value = $global:Config[$section][$key]
                        $output += "$key=$value`r`n"
                    }
                    
                    $output += "`r`n"
                }
                
                $output | Out-File -FilePath $INIPath -Encoding UTF8
                return $true
            }
            catch {
                Write-DebugOutput "Fehler beim Speichern der INI-Datei: $($_.Exception.Message)" -Level ERROR
                return $false
            }
        }
    }
    
    # Stellt sicher, dass DropdownHelpers-Funktionen verfügbar sind
    if (-not (Get-Command -Name "Initialize-AllDropdowns" -ErrorAction SilentlyContinue)) {
        Write-DebugOutput "Dropdown-Hilfsfunktionen nicht gefunden. Erstelle Notfall-Funktionen." -Level WARNING
        
        function Initialize-AllDropdowns {
            [CmdletBinding()]
            param()
            
            Write-DebugOutput "Notfall-Dropdown-Initialisierung" -Level WARNING
            
            # Minimale Funktionalität hier
            $cmbDisplayTemplate = $global:window.FindName("cmbDisplayTemplate")
            if ($cmbDisplayTemplate) {
                $templates = @("Vorname.Nachname", "NachnameV")
                $cmbDisplayTemplate.Items.Clear()
                foreach ($template in $templates) {
                    $cmbDisplayTemplate.Items.Add($template)
                }
                if ($cmbDisplayTemplate.Items.Count -gt 0) {
                    $cmbDisplayTemplate.SelectedIndex = 0
                }
            }
            
            return $true
        }
    }
    
    # Hilfsfunktionen zum Erstellen fehlender Module
    function New-INIEditorModule {
        $moduleContent = @'
# INIEditor.psm1 - Automatisch erstellt
# Funktionen zum Bearbeiten von INI-Dateien für das easyONBOARDING-Tool

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
        
        # INI-Datei speichern
        $output = ""
        
        foreach ($section in $global:Config.Keys) {
            $output += "[$section]`r`n"
            
            foreach ($key in $global:Config[$section].Keys) {
                $value = $global:Config[$section][$key]
                $output += "$key=$value`r`n"
            }
            
            $output += "`r`n"
        }
        
        $output | Out-File -FilePath $INIPath -Encoding UTF8
        Write-DebugOutput "INI-Datei erfolgreich gespeichert" -Level INFO
        return $true
    }
    catch {
        Write-DebugOutput "Fehler beim Speichern der INI-Änderungen: $($_.Exception.Message)" -Level ERROR
        return $false
    }
}

# Exportiere alle Funktionen
Export-ModuleMember -Function Import-INIEditorData, Import-SectionSettings, Save-INIChanges
'@
        
        $iniEditorPath = Join-Path -Path $ModulesDir -ChildPath "INIEditor.psm1"
        Set-Content -Path $iniEditorPath -Value $moduleContent -Force
        
        # Lade das neu erstellte Modul
        Import-Module $iniEditorPath -Force -Global
        Write-DebugOutput "INIEditor-Modul erstellt und geladen." -Level INFO
    }
    
    function New-DropdownHelpersModule {
        $moduleContent = @'
# DropdownHelpers.psm1 - Automatisch erstellt
# Hilfsfunktionen zur Verwaltung von Dropdown-Menüs im easyONBOARDING-Tool

function Initialize-AllDropdowns {
    [CmdletBinding()]
    param()
    
    Write-DebugOutput "Initialisiere alle Dropdown-Menüs" -Level INFO
    
    $initialized = 0
    $failed = 0
    
    # Dropdown-Template initialisieren
    try {
        $cmbDisplayTemplate = $global:window.FindName("cmbDisplayTemplate")
        if ($cmbDisplayTemplate) {
            $templates = @("Vorname.Nachname", "NachnameV", "Nachname.Vorname", "VornameN")
            $cmbDisplayTemplate.Items.Clear()
            foreach ($template in $templates) {
                $cmbDisplayTemplate.Items.Add($template)
            }
            if ($cmbDisplayTemplate.Items.Count -gt 0) {
                $cmbDisplayTemplate.SelectedIndex = 0
            }
            $initialized++
        }
    }
    catch {
        $failed++
        Write-DebugOutput "Fehler bei Template-Dropdown-Initialisierung: $($_.Exception.Message)" -Level WARNING
    }
    
    # Lizenz-Dropdown initialisieren
    try {
        $cmbLicense = $global:window.FindName("cmbLicense")
        if ($cmbLicense) {
            $licenses = @("Standard", "Premium", "Enterprise", "Keine")
            $cmbLicense.Items.Clear()
            foreach ($license in $licenses) {
                $cmbLicense.Items.Add($license)
            }
            if ($cmbLicense.Items.Count -gt 0) {
                $cmbLicense.SelectedIndex = 0
            }
            $initialized++
        }
    }
    catch {
        $failed++
        Write-DebugOutput "Fehler bei Lizenz-Dropdown-Initialisierung: $($_.Exception.Message)" -Level WARNING
    }
    
    Write-DebugOutput "Dropdown-Initialisierung abgeschlossen. Erfolgreich: $initialized, Fehlgeschlagen: $failed" -Level INFO
    return ($failed -eq 0)
}

# Exportiere alle Funktionen
Export-ModuleMember -Function Initialize-AllDropdowns
'@
        
        $dropdownHelpersPath = Join-Path -Path $ModulesDir -ChildPath "DropdownHelpers.psm1"
        Set-Content -Path $dropdownHelpersPath -Value $moduleContent -Force
        
        # Lade das neu erstellte Modul
        Import-Module $dropdownHelpersPath -Force -Global
        Write-DebugOutput "DropdownHelpers-Modul erstellt und geladen." -Level INFO
    }
    
    # Debug-Module bei Bedarf laden
    if ($script:DebugMode -gt 1) {
        Write-DebugOutput "Lade erweiterte Debug-Helfer..." -Level DEBUG
        
        $debugModules = @(
            @{Name = "WPFDebugHelper"; Dependencies = @("FunctionLogDebug", "LoadConfigXAML")}
            @{Name = "ModuleStatusChecker"; Dependencies = @()}
        )
        
        foreach ($module in $debugModules) {
            if (Get-Command -Name "Import-RequiredModule" -ErrorAction SilentlyContinue) {
                Import-RequiredModule -ModuleName $module.Name -Dependencies $module.Dependencies
            }
            else {
                $modulePath = Join-Path -Path $ModulesDir -ChildPath "$($module.Name).psm1"
                if (Test-Path $modulePath) {
                    try {
                        Import-Module $modulePath -Force -Global
                        Write-Host "Debug-Modul geladen: $($module.Name)" -ForegroundColor Green
                    }
                    catch {
                        Write-Host "Fehler beim Laden des Debug-Moduls $($module.Name): $($_.Exception.Message)" -ForegroundColor Yellow
                    }
                }
            }
        }
        
        # Prüfe, ob Modulstatusüberprüfung verfügbar ist
        if (Get-Command -Name "Get-ModuleStatus" -ErrorAction SilentlyContinue) {
            Write-DebugOutput "Überprüfe Modulstatus..." -Level DEBUG
            $moduleStatus = Get-ModuleStatus -ModuleDirectory $ModulesDir
            Write-DebugOutput "Modulstatus: $($moduleStatus | Out-String)" -Level DEBUG
            
            if (Get-Command -Name "Test-ModuleDependencies" -ErrorAction SilentlyContinue) {
                $dependencyCheck = Test-ModuleDependencies -ModuleDirectory $ModulesDir
                Write-DebugOutput "Modulabhängigkeiten: $($dependencyCheck | Out-String)" -Level DEBUG
            }
        }
    }
}
catch {
    Write-Host "Fehler beim Laden der zusätzlichen Module: $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host "Alle Module erfolgreich geladen!" -ForegroundColor Green
if (Get-Command -Name "Write-DebugOutput" -ErrorAction SilentlyContinue) {
    Write-DebugOutput "Modulladeprozess erfolgreich abgeschlossen" -Level INFO
}
#endregion

#region XAML GUI Import
# Verbesserte XAML-Import Funktion mit optimiertem Error-Handling
try {
    Write-DebugOutput "Lade Haupt-GUI..." -Level INFO
    
    # Sicherstellen, dass Assemblies geladen sind
    if (-not ([System.Management.Automation.PSTypeName]'System.Windows.Markup.XamlReader').Type) {
        Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase
    }
    
    # Prüfe, ob XAML-Datei existiert
    if (-not (Test-Path $global:XAMLPath)) {
        throw "XAML-Datei nicht gefunden: $($global:XAMLPath)"
    }
    
    # Importiere XAML mit vorhandener oder fallback Funktion
    if (Get-Command -Name "Import-XamlWindow" -ErrorAction SilentlyContinue) {
        $global:window = Import-XamlWindow -XamlPath $global:XAMLPath
        
        if ($null -eq $global:window) {
            throw "Import-XamlWindow gab ein null-Window-Objekt zurück"
        }
    }
    else {
        # Fallback-Import
        [xml]$xaml = Get-Content -Path $global:XAMLPath -ErrorAction Stop
        $reader = New-Object System.Xml.XmlNodeReader $xaml
        $global:window = [Windows.Markup.XamlReader]::Load($reader)
        
        if ($null -eq $global:window) {
            throw "XamlReader.Load gab ein null-Window-Objekt zurück"
        }
    }
    
    Write-DebugOutput "GUI erfolgreich geladen: $($global:window.Title)" -Level INFO
}
catch {
    Write-DebugOutput "Fehler beim Laden der Haupt-GUI: $($_.Exception.Message)" -Level ERROR
    
    # Zeige Fehlermeldung an
    $errorMessage = "Die Anwendung konnte nicht gestartet werden: $($_.Exception.Message)"
    
    if ([System.Windows.MessageBox]::Show(
        $errorMessage,
        "Kritischer Fehler",
        [System.Windows.MessageBoxButton]::OKCancel,
        [System.Windows.MessageBoxImage]::Error
    ) -eq [System.Windows.MessageBoxResult]::Cancel) {
        exit 1
    }
    
    # Versuche, ein einfaches Ersatzfenster zu erstellen
    try {
        Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase
        
        $fallbackWindow = New-Object System.Windows.Window
        $fallbackWindow.Title = "easyONBOARDING - Notfallmodus"
        $fallbackWindow.Width = 600
        $fallbackWindow.Height = 400
        
        $stackPanel = New-Object System.Windows.Controls.StackPanel
        $fallbackWindow.Content = $stackPanel
        
        $warningText = New-Object System.Windows.Controls.TextBlock
        $warningText.Text = "Die Anwendung konnte nicht im Normalmodus gestartet werden. Fehlermeldung:`n$($_.Exception.Message)"
        $warningText.Margin = New-Object System.Windows.Thickness(10)
        $warningText.TextWrapping = [System.Windows.TextWrapping]::Wrap
        $warningText.Foreground = [System.Windows.Media.Brushes]::Red
        $stackPanel.Children.Add($warningText)
        
        $exitButton = New-Object System.Windows.Controls.Button
        $exitButton.Content = "Beenden"
        $exitButton.Width = 100
        $exitButton.Margin = New-Object System.Windows.Thickness(10)
        $exitButton.Add_Click({ $fallbackWindow.Close() })
        $stackPanel.Children.Add($exitButton)
        
        $global:window = $fallbackWindow
        $global:window.ShowDialog() | Out-Null
        
        exit 1
    }
    catch {
        Write-DebugOutput "Auch Ersatzfenster konnte nicht erstellt werden: $($_.Exception.Message)" -Level ERROR
        exit 1
    }
}
#endregion

#region GUI Element Setup and Event Registration
# Zentrale Funktion zur Registrierung von GUI-Events mit Fehlerbehandlung
function Register-GUIEventWithLogging {
    param(
        [Parameter(Mandatory=$true)]
        [System.Windows.UIElement]$Element,
        
        [Parameter(Mandatory=$true)]
        [string]$EventName,
        
        [Parameter(Mandatory=$true)]
        [scriptblock]$Handler,
        
        [Parameter(Mandatory=$false)]
        [string]$ElementDescription = ""
    )
    
    if ($null -eq $Element) {
        Write-DebugMessage "Element ist null - Event kann nicht registriert werden ($ElementDescription)" -LogLevel "WARNING"
        return $false
    }
    
    try {
        # GUIEventHandler Modul verwenden, wenn verfügbar
        if (Get-Command -Name "Register-GUIEvent" -ErrorAction SilentlyContinue) {
            $result = Register-GUIEvent -Control $Element -EventAction $Handler -EventName $EventName -ErrorMessagePrefix "Fehler in $ElementDescription"
            Write-DebugMessage "Event $EventName für $ElementDescription mit Register-GUIEvent registriert" -LogLevel "DEBUG"
            return $result
        }
        
        # Alternativ direkt registrieren mit Fehlerbehandlung
        $safeHandler = {
            param($sender, $e)
            
            try {
                # Original-Handler ausführen
                & $Handler $sender $e
            }
            catch {
                $errorMsg = "Fehler in $ElementDescription ($EventName): $($_.Exception.Message)"
                Write-DebugMessage $errorMsg -LogLevel "ERROR"
                
                [System.Windows.MessageBox]::Show(
                    $errorMsg,
                    "Fehler",
                    [System.Windows.MessageBoxButton]::OK,
                    [System.Windows.MessageBoxImage]::Error
                )
            }
        }
        
        $eventMethod = "Add_$EventName"
        $Element.$eventMethod($safeHandler)
        
        Write-DebugMessage "Event $EventName für $ElementDescription direkt registriert" -LogLevel "DEBUG"
        return $true
    }
    catch {
        Write-DebugMessage "Fehler beim Registrieren des Events $($_.Exception.Message)" -LogLevel "ERROR"
        return $false
    }
}

# Sichere GUI Control Zugriffsfunktion mit Null-Check
function Get-GUIElement {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Name,
        
        [Parameter(Mandatory=$false)]
        [System.Windows.DependencyObject]$ParentElement = $global:window
    )
    
    # Überprüfe, ob ParentElement nicht null ist
    if ($null -eq $ParentElement) {
        Write-DebugMessage "ParentElement ist null beim Suchen von '$Name'" -LogLevel "WARNING"
        return $null
    }
    
    # Zuerst über FindName versuchen (schnell)
    try {
        $element = $ParentElement.FindName($Name)
        if ($null -ne $element) {
            return $element
        }
    }
    catch {
        Write-DebugMessage "FindName-Methode für '$Name' fehlgeschlagen: $($_.Exception.Message)" -LogLevel "DEBUG"
    }
    
    # Wenn das nicht funktioniert und FindVisualChild verfügbar ist
    if (Get-Command -Name "FindVisualChild" -ErrorAction SilentlyContinue) {
        try {
            # FindVisualChild nur aufrufen, wenn Parent nicht null ist
            if ($null -ne $ParentElement) {
                $element = FindVisualChild -Parent $ParentElement -Name $Name
                if ($null -ne $element) {
                    return $element
                }
            }
        }
        catch {
            Write-DebugMessage "FindVisualChild für '$Name' fehlgeschlagen: $($_.Exception.Message)" -LogLevel "DEBUG"
        }
    }
    
    # Wenn Find-UIElement aus CoreHelpers verfügbar ist
    if (Get-Command -Name "Find-UIElement" -ErrorAction SilentlyContinue) {
        try {
            # Find-UIElement nur aufrufen, wenn Parent nicht null ist
            if ($null -ne $ParentElement) {
                $element = Find-UIElement -Name $Name -Parent $ParentElement
                if ($null -ne $element) {
                    return $element
                }
            }
        }
        catch {
            Write-DebugMessage "Find-UIElement für '$Name' fehlgeschlagen: $($_.Exception.Message)" -LogLevel "DEBUG"
        }
    }
    
    # Wenn nichts funktioniert hat
    Write-DebugMessage "Element '$Name' nicht gefunden" -LogLevel "WARNING"
    return $null
}

# Dropdown-Initialisierungsfunktion
function Initialize-AllDropdowns {
    try {
        Write-DebugMessage "Initializing all dropdown menus" -LogLevel "INFO"
        
        # Nutze die spezialisierte Funktion wenn verfügbar
        if (Get-Command -Name "Initialize-AllDropdowns" -ErrorAction SilentlyContinue) {
            $initResult = Initialize-AllDropdowns
            Write-DebugMessage "Used specialized Initialize-AllDropdowns function" -LogLevel "INFO"
            return $initResult
        }
        
        # Alternativ die einzelnen Dropdown-Funktionen aufrufen
        if (Get-Command -Name "Initialize-DisplayNameTemplateDropdown" -ErrorAction SilentlyContinue) {
            Initialize-DisplayNameTemplateDropdown
            Write-DebugMessage "Initialized DisplayName template dropdown" -LogLevel "INFO"
        }
        
        if (Get-Command -Name "Initialize-LicenseDropdown" -ErrorAction SilentlyContinue) {
            Initialize-LicenseDropdown
            Write-DebugMessage "Initialized license dropdown" -LogLevel "INFO"
        }
        
        if (Get-Command -Name "Initialize-TeamLeaderGroupDropdown" -ErrorAction SilentlyContinue) {
            Initialize-TeamLeaderGroupDropdown
            Write-DebugMessage "Initialized team leader group dropdown" -LogLevel "INFO"
        }
        
        if (Get-Command -Name "Initialize-EmailDomainSuffixDropdown" -ErrorAction SilentlyContinue) {
            Initialize-EmailDomainSuffixDropdown
            Write-DebugMessage "Initialized email domain suffix dropdown" -LogLevel "INFO"
        }
        
        if (Get-Command -Name "Initialize-LocationDropdown" -ErrorAction SilentlyContinue) {
            Initialize-LocationDropdown
            Write-DebugMessage "Initialized location dropdown" -LogLevel "INFO"
        }
        
        Write-DebugMessage "All dropdowns initialized" -LogLevel "INFO"
    }
    catch {
        Write-DebugMessage "Error initializing dropdowns: $($_.Exception.Message)" -LogLevel "ERROR"
    }
}

# AD-Gruppen laden und darstellen
function Load-ADGroupsToInterface {
    try {
        Write-DebugMessage "Loading AD groups to interface" -LogLevel "INFO"
        
        $onboardingTab = Get-GUIElement -Name "Tab_Onboarding"
        if (-not $onboardingTab) {
            Write-DebugMessage "Tab_Onboarding not found" -LogLevel "WARNING"
            return $false
        }
        
        $icADGroups = Get-GUIElement -Name "icADGroups" -ParentElement $onboardingTab
        if (-not $icADGroups) {
            Write-DebugMessage "icADGroups control not found" -LogLevel "WARNING"
            return $false
        }
        
        # Nutze die spezialisierte Funktion wenn verfügbar
        if (Get-Command -Name "Load-ADGroups" -ErrorAction SilentlyContinue) {
            Load-ADGroups
            Write-DebugMessage "AD groups loaded using specialized function" -LogLevel "INFO"
            return $true
        }
        else {
            Write-DebugMessage "Load-ADGroups function not found" -LogLevel "WARNING"
            return $false
        }
    }
    catch {
        Write-DebugMessage "Error loading AD groups: $($_.Exception.Message)" -LogLevel "ERROR"
        return $false
    }
}

# UPN-Template Display aktualisieren
function Update-UPNDisplayTemplate {
    try {
        Write-DebugMessage "Updating UPN template display" -LogLevel "INFO"
        
        # Nutze die spezialisierte Funktion wenn verfügbar
        if (Get-Command -Name "Update-UPNTemplateDisplay" -ErrorAction SilentlyContinue) {
            Update-UPNTemplateDisplay
            Write-DebugMessage "UPN template display updated using specialized function" -LogLevel "INFO"
            return $true
        }
        else {
            Write-DebugMessage "Update-UPNTemplateDisplay function not found" -LogLevel "WARNING"
            return $false
        }
    }
    catch {
        Write-DebugMessage "Error updating UPN template display: $($_.Exception.Message)" -LogLevel "ERROR"
        return $false
    }
}

# GUI-Initialisierungsfunktionen aufrufen
$global:window.Add_Loaded({
    try {
        Write-DebugMessage "Window Loaded event triggered - initializing interface" -LogLevel "INFO"
        
        # Dropdown-Menüs initialisieren
        Initialize-AllDropdowns
        
        # AD-Gruppen laden
        Load-ADGroupsToInterface
        
        # UPN-Templates aktualisieren
        Update-UPNDisplayTemplate
        
        # Prüfe, ob OU Dropdown einen Refresh-Event braucht
        if (Get-Command -Name "Register-OUDropdownRefreshEvent" -ErrorAction SilentlyContinue) {
            Register-OUDropdownRefreshEvent
            Write-DebugMessage "OU dropdown refresh event registered" -LogLevel "INFO"
        }
        
        Write-DebugMessage "Interface initialization completed" -LogLevel "INFO"
    }
    catch {
        Write-DebugMessage "Error in Window.Loaded event: $($_.Exception.Message)" -LogLevel "ERROR"
    }
})
#endregion

#region Logo Upload Button Setup
Write-DebugMessage "Setting up logo upload button."

# Sets up the logo upload button and its functionality
$btnUploadLogo = Get-GUIElement -Name "btnUploadLogo"
if ($btnUploadLogo) {
    Register-GUIEventWithLogging -Element $btnUploadLogo -EventName "Click" -ElementDescription "Upload Logo Button" -Handler {
        if (-not $global:Config -or -not $global:Config.Contains("Report")) {
            [System.Windows.MessageBox]::Show("Report section is missing in the configuration", "Error", "OK", "Error")
            return
        }
        
        # Rufe die Set-Logo Funktion auf
        if (Get-Command -Name "Set-Logo" -ErrorAction SilentlyContinue) {
            Set-Logo
        }
        else {
            Write-DebugMessage "Set-Logo function not found" -LogLevel "WARNING"
            [System.Windows.MessageBox]::Show("Logo upload functionality is not available", "Function Not Found", "OK", "Warning")
        }
    }
}
#endregion

#region Onboarding Tab Setup
Write-DebugMessage "Setting up onboarding tab functionality"

# Implements the main onboarding tab functionality
$onboardingTab = Get-GUIElement -Name "Tab_Onboarding"
if ($onboardingTab) {
    Write-DebugMessage "Onboarding tab found, setting up event handlers" -LogLevel "INFO"
    
    # Implements the start onboarding button click handler
    $btnStartOnboarding = Get-GUIElement -Name "btnOnboard" -ParentElement $onboardingTab
    if ($btnStartOnboarding) {
        Register-GUIEventWithLogging -Element $btnStartOnboarding -EventName "Click" -ElementDescription "Start Onboarding Button" -Handler {
            try {
                Write-DebugMessage "Onboarding button clicked, starting onboarding process" -LogLevel "INFO"
                
                # Sammle alle notwendigen Daten aus der GUI
                $userData = [PSCustomObject]@{
                    FirstName = (Get-GUIElement -Name "txtFirstName").Text
                    LastName = (Get-GUIElement -Name "txtLastName").Text
                    DisplayName = (Get-GUIElement -Name "txtDisplayName").Text
                    Description = (Get-GUIElement -Name "txtDescription").Text
                    OfficeRoom = (Get-GUIElement -Name "txtOffice").Text
                    PhoneNumber = (Get-GUIElement -Name "txtPhone").Text
                    MobileNumber = (Get-GUIElement -Name "txtMobile").Text
                    Position = (Get-GUIElement -Name "txtPosition").Text
                    DepartmentField = (Get-GUIElement -Name "txtDepartment").Text
                    Ablaufdatum = (Get-GUIElement -Name "txtTermination").Text
                    External = (Get-GUIElement -Name "chkExternal").IsChecked
                    AccountDisabled = -not (Get-GUIElement -Name "chkAccountDisabled").IsChecked
                    PasswordNeverExpires = (Get-GUIElement -Name "chkPWNeverExpires").IsChecked
                    SmartcardLogonRequired = (Get-GUIElement -Name "chkSmartcardLogon").IsChecked
                    CannotChangePassword = (Get-GUIElement -Name "chkPreventPWChange").IsChecked
                    MustChangePassword = (Get-GUIElement -Name "chkPWChangeLogon").IsChecked
                    PasswordMode = if ((Get-GUIElement -Name "rbRand").IsChecked) { 1 } else { 0 }
                    FixPassword = (Get-GUIElement -Name "txtFixPW").Text
                    EmailAddress = (Get-GUIElement -Name "txtEmail").Text
                    OU = (Get-GUIElement -Name "cmbOU").SelectedItem
                    UPNFormat = (Get-GUIElement -Name "cmbDisplayTemplate").SelectedValue
                    TL = (Get-GUIElement -Name "chkTL").IsChecked
                    AL = (Get-GUIElement -Name "chkAL").IsChecked
                    TLGroup = (Get-GUIElement -Name "cmbTLGroup").SelectedValue
                    OutputHTML = (Get-GUIElement -Name "chkHTML").IsChecked
                    OutputTXT = (Get-GUIElement -Name "chkTXT").IsChecked
                }
                
                # Prüfe auf Lizenz
                $licenseComboBox = Get-GUIElement -Name "cmbLicense"
                if ($licenseComboBox -and $licenseComboBox.SelectedValue) {
                    $userData | Add-Member -NotePropertyName License -NotePropertyValue $licenseComboBox.SelectedValue
                }
                
                # Prüfe auf Mail Suffix
                $mailSuffixComboBox = Get-GUIElement -Name "cmbSuffix"
                if ($mailSuffixComboBox -and $mailSuffixComboBox.SelectedValue) {
                    $userData | Add-Member -NotePropertyName MailSuffix -NotePropertyValue $mailSuffixComboBox.SelectedValue
                }
                
                # Hole die ausgewählten AD-Gruppen
                $icADGroups = Get-GUIElement -Name "icADGroups"
                if ($icADGroups) {
                    if (Get-Command -Name "Get-SelectedADGroups" -ErrorAction SilentlyContinue) {
                        $selectedGroups = Get-SelectedADGroups -Panel $icADGroups
                        $userData | Add-Member -NotePropertyName ADGroupsSelected -NotePropertyValue $selectedGroups
                    }
                }
                
                # Führe die Onboarding-Funktion aus
                if (-not (Get-Command -Name "Invoke-Onboarding" -ErrorAction SilentlyContinue)) {
                    throw "Onboarding function not found. Please check that the CoreEASYONBOARDING module is loaded."
                }
                
                $progressBar = Get-GUIElement -Name "progressBar"
                $lblStatus = Get-GUIElement -Name "lblStatus"
                
                if ($progressBar) { $progressBar.Value = 10 }
                if ($lblStatus) { $lblStatus.Text = "Starting onboarding process..." }
                
                # Invoke the onboarding function
                Write-DebugMessage "Calling Invoke-Onboarding with user data" -LogLevel "INFO"
                $result = Invoke-Onboarding -userData $userData -Config $global:Config
                
                if ($progressBar) { $progressBar.Value = 90 }
                if ($lblStatus) { $lblStatus.Text = "Onboarding completed." }
                
                # Zeige Erfolgsmeldung
                [System.Windows.MessageBox]::Show(
                    "The user $($result.SamAccountName) has been onboarded successfully.`n`n" +
                    "Username: $($result.SamAccountName)`n" +
                    "UPN: $($result.UPN)`n" +
                    "Password: $($result.Password)",
                    "Onboarding Successful",
                    "OK",
                    "Information"
                )
                
                if ($progressBar) { $progressBar.Value = 100 }
                if ($lblStatus) { $lblStatus.Text = "Ready" }
            }
            catch {
                # Zeige Fehlermeldung
                Write-DebugMessage "Error during onboarding: $($_.Exception.Message)" -LogLevel "ERROR"
                [System.Windows.MessageBox]::Show(
                    "An error occurred during onboarding:`n$($_.Exception.Message)",
                    "Onboarding Failed",
                    "OK",
                    "Error"
                )
                
                $progressBar = Get-GUIElement -Name "progressBar"
                $lblStatus = Get-GUIElement -Name "lblStatus"
                
                if ($progressBar) { $progressBar.Value = 0 }
                if ($lblStatus) { $lblStatus.Text = "Error: Onboarding failed" }
            }
        }
    }
    else {
        Write-DebugMessage "Onboarding button (btnOnboard) not found - check your XAML" -LogLevel "WARNING"
    }
}
else {
    Write-DebugMessage "Onboarding tab (Tab_Onboarding) not found - check your XAML" -LogLevel "WARNING"
}
#endregion

#region PDF Creation Button Setup
Write-DebugMessage "Setting up PDF creation button"

# Implements the PDF creation button functionality
$btnPDF = Get-GUIElement -Name "btnPDF"
if ($btnPDF) {
    Register-GUIEventWithLogging -Element $btnPDF -EventName "Click" -ElementDescription "PDF Creation Button" -Handler {
        try {
            # PDF-Erstellung noch zu implementieren
            [System.Windows.MessageBox]::Show(
                "PDF creation is not yet implemented in this version.",
                "Feature Not Implemented",
                "OK",
                "Information"
            )
        } catch {
            Write-DebugMessage "Error handling PDF button click: $($_.Exception.Message)" -LogLevel "ERROR"
            [System.Windows.MessageBox]::Show(
                "An error occurred: $($_.Exception.Message)",
                "Error",
                "OK",
                "Error"
            )
        }
    }
}
#endregion

#region Info Button Setup
Write-DebugMessage "Setting up info button."

# Implements the info button functionality to show application information
$btnInfo = Get-GUIElement -Name "btnInfo"
if ($btnInfo) {
    Register-GUIEventWithLogging -Element $btnInfo -EventName "Click" -ElementDescription "Info Button" -Handler {
        $infoFilePath = Join-Path $PSScriptRoot "easyIT.txt"
        
        if (Test-Path $infoFilePath) {
            Start-Process -FilePath $infoFilePath
        } else {
            [System.Windows.MessageBox]::Show(
                "Information file not found: $infoFilePath",
                "File Not Found",
                "OK",
                "Warning"
            )
        }
    }
}
#endregion

#region Close Button Setup
Write-DebugMessage "Setting up close button."

# Implements the close button functionality to exit the application
$btnClose = Get-GUIElement -Name "btnClose"
if ($btnClose) {
    Register-GUIEventWithLogging -Element $btnClose -EventName "Click" -ElementDescription "Close Button" -Handler {
        $global:window.Close()
    }
}
#endregion

#region AD Update Tab Setup
Write-DebugMessage "Setting up AD Update tab functionality"

# Implements the AD Update tab functionality for user management
$adUpdateTab = Get-GUIElement -Name "Tab_ADUpdate"
if ($adUpdateTab) {
    Write-DebugMessage "AD Update tab found, setting up event handlers" -LogLevel "INFO"
    
    try {
        # Retrieve search and basic controls – names must match the XAML definitions
        $txtSearchADUpdate   = Get-GUIElement -Name "txtSearchADUpdate" -ParentElement $adUpdateTab
        $btnSearchADUpdate   = Get-GUIElement -Name "btnSearchADUpdate" -ParentElement $adUpdateTab
        $lstUsersADUpdate    = Get-GUIElement -Name "lstUsersADUpdate" -ParentElement $adUpdateTab
        
        # The refresh button has a specific name in the XAML
        $btnRefreshADUserList = Get-GUIElement -Name "btnRefreshOU_Kopieren" -ParentElement $adUpdateTab
        
        # Action buttons for user management
        $btnADUserUpdate     = Get-GUIElement -Name "btnADUserUpdate" -ParentElement $adUpdateTab
        $btnADUserCancel     = Get-GUIElement -Name "btnADUserCancel" -ParentElement $adUpdateTab
        $btnAddGroupUpdate   = Get-GUIElement -Name "btnAddGroupUpdate" -ParentElement $adUpdateTab
        $btnRemoveGroupUpdate= Get-GUIElement -Name "btnRemoveGroupUpdate" -ParentElement $adUpdateTab
        
        # Basic information controls for user details
        $txtFirstNameUpdate  = Get-GUIElement -Name "txtFirstNameUpdate" -ParentElement $adUpdateTab
        $txtLastNameUpdate   = Get-GUIElement -Name "txtLastNameUpdate" -ParentElement $adUpdateTab
        $txtDisplayNameUpdate= Get-GUIElement -Name "txtDisplayNameUpdate" -ParentElement $adUpdateTab
        $txtEmailUpdate      = Get-GUIElement -Name "txtEmailUpdate" -ParentElement $adUpdateTab
        $txtDepartmentUpdate = Get-GUIElement -Name "txtDepartmentUpdate" -ParentElement $adUpdateTab
        
        # Contact information controls
        $txtPhoneUpdate      = Get-GUIElement -Name "txtPhoneUpdate" -ParentElement $adUpdateTab
        $txtMobileUpdate     = Get-GUIElement -Name "txtMobileUpdate" -ParentElement $adUpdateTab
        $txtOfficeUpdate     = Get-GUIElement -Name "txtOfficeUpdate" -ParentElement $adUpdateTab
        
        # Account options and settings
        $chkAccountEnabledUpdate      = Get-GUIElement -Name "chkAccountEnabledUpdate" -ParentElement $adUpdateTab
        $chkPasswordNeverExpiresUpdate= Get-GUIElement -Name "chkPasswordNeverExpiresUpdate" -ParentElement $adUpdateTab
        $chkMustChangePasswordUpdate  = Get-GUIElement -Name "chkMustChangePasswordUpdate" -ParentElement $adUpdateTab
        
        # Extended properties controls for additional user attributes
        $txtManagerUpdate    = Get-GUIElement -Name "txtManagerUpdate" -ParentElement $adUpdateTab
        $txtJobTitleUpdate   = Get-GUIElement -Name "txtJobTitleUpdate" -ParentElement $adUpdateTab
        $txtLocationUpdate   = Get-GUIElement -Name "txtLocationUpdate" -ParentElement $adUpdateTab
        $txtEmployeeIDUpdate = Get-GUIElement -Name "txtEmployeeIDUpdate" -ParentElement $adUpdateTab
        
        # Group management control
        $lstGroupsUpdate     = Get-GUIElement -Name "lstGroupsUpdate" -ParentElement $adUpdateTab
        if ($lstGroupsUpdate) {
            # Im Original-Skript gibt es hier weitere Initialisierung
        }
    }
    catch {
        Write-DebugMessage "Error loading AD Update tab controls: $($_.Exception.Message)" -LogLevel "ERROR"
        Write-LogMessage -Message "Failed to initialize AD Update tab controls: $($_.Exception.Message)" -LogLevel "ERROR"
    }
    
    # Register search event for finding AD users
    if ($btnSearchADUpdate -and $txtSearchADUpdate -and $lstUsersADUpdate) {
        Write-DebugMessage "Registering event for btnSearchADUpdate in Tab_ADUpdate" -LogLevel "INFO"
        Register-GUIEventWithLogging -Element $btnSearchADUpdate -EventName "Click" -ElementDescription "Search AD Update Button" -Handler {
            try {
                $searchTerm = $txtSearchADUpdate.Text
                
                if ([string]::IsNullOrWhiteSpace($searchTerm)) {
                    [System.Windows.MessageBox]::Show("Please enter a search term", "Input Required", "OK", "Warning")
                    return
                }
                
                # Define filter pattern for AD search
                $filter = "DisplayName -like '*$searchTerm*' -or SamAccountName -like '*$searchTerm*' -or UserPrincipalName -like '*$searchTerm*'"
                
                # Execute the search
                $users = Get-ADUser -Filter $filter -Properties DisplayName, SamAccountName, UserPrincipalName, Mail, Enabled -ResultSetSize 200
                
                if ($users.Count -eq 0) {
                    [System.Windows.MessageBox]::Show("No users found matching '$searchTerm'", "No Results", "OK", "Information")
                    return
                }
                
                # Clear and populate ListView
                $lstUsersADUpdate.Items.Clear()
                foreach ($user in $users) {
                    $item = New-Object PSObject -Property @{
                        DisplayName = $user.DisplayName
                        SamAccountName = $user.SamAccountName
                        Email = $user.Mail
                        Enabled = $user.Enabled
                        DistinguishedName = $user.DistinguishedName
                    }
                    $lstUsersADUpdate.Items.Add($item)
                }
                
                Write-DebugMessage "AD search completed with $($users.Count) results" -LogLevel "INFO"
            }
            catch {
                Write-DebugMessage "Error searching AD: $($_.Exception.Message)" -LogLevel "ERROR"
                [System.Windows.MessageBox]::Show("Error searching Active Directory: $($_.Exception.Message)", "Search Error", "OK", "Error")
            }
        }
    }
    else {
        Write-DebugMessage "Required search controls missing in Tab_ADUpdate" -LogLevel "WARNING"
    }

    # Register Refresh User List event to load active directory users
    if ($btnRefreshADUserList -and $lstUsersADUpdate) {
        Write-DebugMessage "Registering event for Refresh User List button in Tab_ADUpdate" -LogLevel "INFO"
        Register-GUIEventWithLogging -Element $btnRefreshADUserList -EventName "Click" -ElementDescription "Refresh AD User List Button" -Handler {
            try {
                # Attempt to load users (limited to recent or active)
                $users = Get-ADUser -Filter "Enabled -eq 'True'" -Properties DisplayName, SamAccountName, Mail, Enabled, Modified -ResultSetSize 200 |
                         Sort-Object -Property Modified -Descending
                
                # Clear and populate ListView
                $lstUsersADUpdate.Items.Clear()
                foreach ($user in $users) {
                    $item = New-Object PSObject -Property @{
                        DisplayName = $user.DisplayName
                        SamAccountName = $user.SamAccountName
                        Email = $user.Mail
                        Enabled = $user.Enabled
                        DistinguishedName = $user.DistinguishedName
                    }
                    $lstUsersADUpdate.Items.Add($item)
                }
                
                Write-DebugMessage "Refreshed AD user list with $($users.Count) users" -LogLevel "INFO"
            }
            catch {
                Write-DebugMessage "Error refreshing AD user list: $($_.Exception.Message)" -LogLevel "ERROR"
                [System.Windows.MessageBox]::Show("Error loading Active Directory users: $($_.Exception.Message)", "Refresh Error", "OK", "Error")
            }
        }
    }

    # Populate update fields when a user is selected from the list
    if ($lstUsersADUpdate) {
        Register-GUIEventWithLogging -Element $lstUsersADUpdate -EventName "SelectionChanged" -ElementDescription "AD User List" -Handler {
            try {
                $selectedUser = $lstUsersADUpdate.SelectedItem
                
                if ($null -eq $selectedUser) {
                    return
                }
                
                # Get full user details from AD
                $userDetail = Get-ADUser -Identity $selectedUser.DistinguishedName -Properties *
                
                # Populate form fields with user details
                $txtFirstNameUpdate.Text = $userDetail.GivenName
                $txtLastNameUpdate.Text = $userDetail.Surname
                $txtDisplayNameUpdate.Text = $userDetail.DisplayName
                $txtEmailUpdate.Text = $userDetail.mail
                $txtDepartmentUpdate.Text = $userDetail.Department
                
                # Contact information
                $txtPhoneUpdate.Text = $userDetail.OfficePhone
                $txtMobileUpdate.Text = $userDetail.Mobile
                $txtOfficeUpdate.Text = $userDetail.Office
                
                # Account options
                $chkAccountEnabledUpdate.IsChecked = $userDetail.Enabled
                $chkPasswordNeverExpiresUpdate.IsChecked = $userDetail.PasswordNeverExpires
                
                # Extended properties
                $txtManagerUpdate.Text = if ($userDetail.Manager) { (Get-ADUser -Identity $userDetail.Manager).Name } else { "" }
                $txtJobTitleUpdate.Text = $userDetail.Title
                $txtLocationUpdate.Text = $userDetail.l
                $txtEmployeeIDUpdate.Text = $userDetail.EmployeeID
                
                # Load group memberships
                $lstGroupsUpdate.Items.Clear()
                $groups = Get-ADPrincipalGroupMembership -Identity $userDetail.DistinguishedName
                foreach ($group in $groups) {
                    $lstGroupsUpdate.Items.Add($group.Name)
                }
                
                Write-DebugMessage "Loaded details for user: $($selectedUser.SamAccountName)" -LogLevel "INFO"
            }
            catch {
                Write-DebugMessage "Error loading user details: $($_.Exception.Message)" -LogLevel "ERROR"
                [System.Windows.MessageBox]::Show("Error loading user details: $($_.Exception.Message)", "User Details Error", "OK", "Error")
            }
        }
    }

    # Update User button event handler - save changes to Active Directory
    if ($btnADUserUpdate) {
        Write-DebugMessage "Registering event handler for btnADUserUpdate" -LogLevel "INFO"
        Register-GUIEventWithLogging -Element $btnADUserUpdate -EventName "Click" -ElementDescription "Update AD User Button" -Handler {
            try {
                $selectedUser = $lstUsersADUpdate.SelectedItem
                
                if ($null -eq $selectedUser) {
                    [System.Windows.MessageBox]::Show("Please select a user first", "No User Selected", "OK", "Warning")
                    return
                }
                
                # Collect all updated properties
                $updateParams = @{
                    Identity = $selectedUser.DistinguishedName
                    GivenName = $txtFirstNameUpdate.Text
                    Surname = $txtLastNameUpdate.Text
                    DisplayName = $txtDisplayNameUpdate.Text
                    EmailAddress = $txtEmailUpdate.Text
                    Department = $txtDepartmentUpdate.Text
                    OfficePhone = $txtPhoneUpdate.Text
                    MobilePhone = $txtMobileUpdate.Text
                    Office = $txtOfficeUpdate.Text
                    Title = $txtJobTitleUpdate.Text
                    Enabled = $chkAccountEnabledUpdate.IsChecked
                    PasswordNeverExpires = $chkPasswordNeverExpiresUpdate.IsChecked
                }
                
                # Update the user
                Set-ADUser @updateParams
                
                # Handle ChangePasswordAtNextLogon separately as it requires different parameters
                if ($chkMustChangePasswordUpdate.IsChecked) {
                    Set-ADUser -Identity $selectedUser.DistinguishedName -ChangePasswordAtLogon $true
                }
                
                [System.Windows.MessageBox]::Show("User updated successfully", "Update Successful", "OK", "Information")
                Write-DebugMessage "Updated user: $($selectedUser.SamAccountName)" -LogLevel "INFO"
            }
            catch {
                Write-DebugMessage "Error updating user: $($_.Exception.Message)" -LogLevel "ERROR"
                [System.Windows.MessageBox]::Show("Error updating user: $($_.Exception.Message)", "Update Error", "OK", "Error")
            }
        }
    }

    # Cancel button event handler - reset form fields
    if ($btnADUserCancel) {
        Write-DebugMessage "Registering event handler for btnADUserCancel" -LogLevel "INFO"
        Register-GUIEventWithLogging -Element $btnADUserCancel -EventName "Click" -ElementDescription "Cancel AD User Update Button" -Handler {
            # Clear all form fields
            $txtFirstNameUpdate.Text = ""
            $txtLastNameUpdate.Text = ""
            $txtDisplayNameUpdate.Text = ""
            $txtEmailUpdate.Text = ""
            $txtDepartmentUpdate.Text = ""
            $txtPhoneUpdate.Text = ""
            $txtMobileUpdate.Text = ""
            $txtOfficeUpdate.Text = ""
            $txtManagerUpdate.Text = ""
            $txtJobTitleUpdate.Text = ""
            $txtLocationUpdate.Text = ""
            $txtEmployeeIDUpdate.Text = ""
            
            $chkAccountEnabledUpdate.IsChecked = $true
            $chkPasswordNeverExpiresUpdate.IsChecked = $false
            $chkMustChangePasswordUpdate.IsChecked = $false
            
            $lstGroupsUpdate.Items.Clear()
            $lstUsersADUpdate.SelectedItem = $null
            
            Write-DebugMessage "Form fields reset" -LogLevel "INFO"
        }
    }

    # Add Group button functionality
    if ($btnAddGroupUpdate) {
        Register-GUIEventWithLogging -Element $btnAddGroupUpdate -EventName "Click" -ElementDescription "Add Group Button" -Handler {
            try {
                $selectedUser = $lstUsersADUpdate.SelectedItem
                
                if ($null -eq $selectedUser) {
                    [System.Windows.MessageBox]::Show("Please select a user first", "No User Selected", "OK", "Warning")
                    return
                }
                
                # Prompt for group name using a simple input dialog
                $groupName = [Microsoft.VisualBasic.Interaction]::InputBox("Enter group name:", "Add Group", "")
                
                if ([string]::IsNullOrWhiteSpace($groupName)) {
                    return
                }
                
                # Try to find the group
                try {
                    $group = Get-ADGroup -Identity $groupName
                }
                catch {
                    [System.Windows.MessageBox]::Show("Group '$groupName' not found", "Group Not Found", "OK", "Warning")
                    return
                }
                
                # Add user to group
                Add-ADGroupMember -Identity $groupName -Members $selectedUser.DistinguishedName
                
                # Refresh group list
                $lstGroupsUpdate.Items.Add($groupName)
                
                [System.Windows.MessageBox]::Show("User added to group '$groupName'", "Group Added", "OK", "Information")
                Write-DebugMessage "Added user $($selectedUser.SamAccountName) to group $groupName" -LogLevel "INFO"
            }
            catch {
                Write-DebugMessage "Error adding group: $($_.Exception.Message)" -LogLevel "ERROR"
                [System.Windows.MessageBox]::Show("Error adding group: $($_.Exception.Message)", "Group Error", "OK", "Error")
            }
        }
    }

    # Remove Group button functionality
    if ($btnRemoveGroupUpdate) {
        Register-GUIEventWithLogging -Element $btnRemoveGroupUpdate -EventName "Click" -ElementDescription "Remove Group Button" -Handler {
            try {
                $selectedUser = $lstUsersADUpdate.SelectedItem
                $selectedGroup = $lstGroupsUpdate.SelectedItem
                
                if ($null -eq $selectedUser) {
                    [System.Windows.MessageBox]::Show("Please select a user first", "No User Selected", "OK", "Warning")
                    return
                }
                
                if ($null -eq $selectedGroup) {
                    [System.Windows.MessageBox]::Show("Please select a group first", "No Group Selected", "OK", "Warning")
                    return
                }
                
                # Remove user from group
                Remove-ADGroupMember -Identity $selectedGroup -Members $selectedUser.DistinguishedName
                
                # Refresh group list
                $lstGroupsUpdate.Items.Remove($selectedGroup)
                
                [System.Windows.MessageBox]::Show("User removed from group '$selectedGroup'", "Group Removed", "OK", "Information")
                Write-DebugMessage "Removed user $($selectedUser.SamAccountName) from group $selectedGroup" -LogLevel "INFO"
            }
            catch {
                Write-DebugMessage "Error removing group: $($_.Exception.Message)" -LogLevel "ERROR"
                [System.Windows.MessageBox]::Show("Error removing group: $($_.Exception.Message)", "Group Error", "OK", "Error")
            }
        }
    }
}
else {
    Write-DebugMessage "AD Update tab (Tab_ADUpdate) not found in XAML." -LogLevel "WARNING"
}
#endregion

#region Settings Tab Setup
# Implements the Settings tab functionality for INI editing
Write-DebugMessage "Settings tab loaded."

$tabINIEditor = Get-GUIElement -Name "Tab_INIEditor"
if ($tabINIEditor) {
    # Register event handlers for the INI editor UI
    $listViewINIEditor = Get-GUIElement -Name "listViewINIEditor" -ParentElement $tabINIEditor
    $dataGridINIEditor = Get-GUIElement -Name "dataGridINIEditor" -ParentElement $tabINIEditor
    
    if ($listViewINIEditor) {
        Register-GUIEventWithLogging -Element $listViewINIEditor -EventName "SelectionChanged" -ElementDescription "INI Editor ListView" -Handler {
            if ($null -ne $listViewINIEditor.SelectedItem) {
                $selectedSection = $listViewINIEditor.SelectedItem.SectionName
                Write-DebugMessage "Section selected: $selectedSection" -LogLevel "INFO"
                Import-SectionSettings -SectionName $selectedSection -DataGrid $dataGridINIEditor -INIPath $global:INIPath
            }
        }
    }

    # Setup DataGrid cell edit handling
    if ($dataGridINIEditor) {
        # Make sure DataGrid is editable
        $dataGridINIEditor.IsReadOnly = $false
        
        Register-GUIEventWithLogging -Element $dataGridINIEditor -EventName "CellEditEnding" -ElementDescription "INI Editor DataGrid" -Handler {
            param($eventSender, $e)
            
            try {
                if ($e.EditAction -eq [System.Windows.Controls.DataGridEditAction]::Commit) {
                    $item = $e.Row.Item
                    $column = $e.Column
                    
                    if ($null -ne $listViewINIEditor.SelectedItem) {
                        $sectionName = $listViewINIEditor.SelectedItem.SectionName
                        
                        if ($column.Header -eq "Key") {
                            # Key was edited - update the dictionary
                            $oldKey = $item.OriginalKey
                            $newKey = $item.Key
                            
                            if ($oldKey -ne $newKey) {
                                Write-DebugMessage "Changing key from '$oldKey' to '$newKey' in section $sectionName" -LogLevel "INFO"
                                
                                # Get the current value
                                $value = $global:Config[$sectionName][$oldKey]
                                
                                # Remove old key and add new key
                                $global:Config[$sectionName].Remove($oldKey)
                                $global:Config[$sectionName][$newKey] = $value
                                
                                # Update OriginalKey for future edits
                                $item.OriginalKey = $newKey
                            }
                        }
                        elseif ($column.Header -eq "Value") {
                            # Value was edited
                            $key = $item.Key
                            $newValue = $item.Value
                            
                            Write-DebugMessage "Updating value for [$sectionName] $key to: $newValue" -LogLevel "INFO"
                            $global:Config[$sectionName][$key] = $newValue
                        }
                    }
                }
            }
            catch {
                Write-DebugMessage "Error handling cell edit: $($_.Exception.Message)" -LogLevel "ERROR"
            }
        }
    }

    # Add key button handler
    $btnAddKey = Get-GUIElement -Name "btnAddKey" -ParentElement $tabINIEditor
    if ($btnAddKey) {
        Register-GUIEventWithLogging -Element $btnAddKey -EventName "Click" -ElementDescription "Add Key Button" -Handler {
            try {
                if ($null -eq $listViewINIEditor.SelectedItem) {
                    [System.Windows.MessageBox]::Show("Please select a section first.", "Note", "OK", "Information")
                    return
                }
                
                $sectionName = $listViewINIEditor.SelectedItem.SectionName
                
                # Prompt for new key and value
                $keyName = [Microsoft.VisualBasic.Interaction]::InputBox("Enter the name of the new key:", "New Key", "")
                if ([string]::IsNullOrWhiteSpace($keyName)) { return }
                
                $keyValue = [Microsoft.VisualBasic.Interaction]::InputBox("Enter the value for '$keyName':", "New Value", "")
                
                # Add to config
                $global:Config[$sectionName][$keyName] = $keyValue
                
                # Refresh view
                Import-SectionSettings -SectionName $sectionName -DataGrid $dataGridINIEditor -INIPath $global:INIPath
                
                Write-DebugMessage "Added new key '$keyName' with value '$keyValue' to section '$sectionName'" -LogLevel "INFO"
            }
            catch {
                Write-DebugMessage "Error adding new key: $($_.Exception.Message)" -LogLevel "ERROR"
                [System.Windows.MessageBox]::Show("Error adding: $($_.Exception.Message)", "Error", "OK", "Error")
            }
        }
    }

    # Remove key button handler
    $btnRemoveKey = Get-GUIElement -Name "btnRemoveKey" -ParentElement $tabINIEditor
    if ($btnRemoveKey) {
        Register-GUIEventWithLogging -Element $btnRemoveKey -EventName "Click" -ElementDescription "Remove Key Button" -Handler {
            try {
                if ($null -eq $dataGridINIEditor.SelectedItem) {
                    [System.Windows.MessageBox]::Show("Please select a key first.", "Note", "OK", "Information")
                    return
                }
                
                $selectedItem = $dataGridINIEditor.SelectedItem
                $keyToRemove = $selectedItem.Key
                $sectionName = $listViewINIEditor.SelectedItem.SectionName
                
                $confirmation = [System.Windows.MessageBox]::Show(
                    "Do you really want to delete the key '$keyToRemove'?",
                    "Confirmation",
                    "YesNo",
                    "Question"
                )
                
                if ($confirmation -eq [System.Windows.MessageBoxResult]::Yes) {
                    # Remove from config
                    $global:Config[$sectionName].Remove($keyToRemove)
                    
                    # Refresh view
                    Import-SectionSettings -SectionName $sectionName -DataGrid $dataGridINIEditor -INIPath $global:INIPath
                    
                    Write-DebugMessage "Removed key '$keyToRemove' from section '$sectionName'" -LogLevel "INFO"
                }
            }
            catch {
                Write-DebugMessage "Error removing key: $($_.Exception.Message)" -LogLevel "ERROR"
                [System.Windows.MessageBox]::Show("Error deleting: $($_.Exception.Message)", "Error", "OK", "Error")
            }
        }
    }

    # Register the save button handler
    $btnSaveINIChanges = Get-GUIElement -Name "btnSaveINIChanges" -ParentElement $tabINIEditor
    if ($btnSaveINIChanges) {
        Register-GUIEventWithLogging -Element $btnSaveINIChanges -EventName "Click" -ElementDescription "Save INI Changes Button" -Handler {
            try {
                $result = Save-INIChanges -INIPath $global:INIPath
                if ($result) {
                    [System.Windows.MessageBox]::Show("INI file successfully saved.", "Success", "OK", "Information")
                    # Reload the INI to reflect any changes in the UI
                    $global:Config = Get-IniContent -Path $global:INIPath
                    Import-INIEditorData
                    
                    # Update current selection if available
                    if ($listViewINIEditor.SelectedItem -ne $null) {
                        $selectedSection = $listViewINIEditor.SelectedItem.SectionName
                        Import-SectionSettings -SectionName $selectedSection -DataGrid $dataGridINIEditor -INIPath $global:INIPath
                    }
                }
                else {
                    [System.Windows.MessageBox]::Show("There was a problem saving the INI file.", "Error", "OK", "Error")
                }
            }
            catch {
                Write-DebugMessage "Error in Save button handler: $($_.Exception.Message)" -LogLevel "ERROR"
                [System.Windows.MessageBox]::Show("Error: $($_.Exception.Message)", "Error", "OK", "Error")
            }
        }
    }
    
    # Info button event handler
    $btnINIEditorInfo = Get-GUIElement -Name "btnINIEditorInfo" -ParentElement $tabINIEditor
    if ($btnINIEditorInfo) {
        Register-GUIEventWithLogging -Element $btnINIEditorInfo -EventName "Click" -ElementDescription "INI Editor Info Button" -Handler {
            [System.Windows.MessageBox]::Show(
                "INI Editor Help:
                
1. Select a section from the left list
2. Edit values directly in the table
3. Use the buttons to add or remove entries
4. Save your changes with the Save button

INI file: $global:INIPath", 
                "INI Editor Help", "OK", "Information")
        }
    }
    
    # Initialize the editor with data when the tab is first loaded
    Register-GUIEventWithLogging -Element $tabINIEditor -EventName "Loaded" -ElementDescription "INI Editor Tab" -Handler {
        Write-DebugMessage "INI Editor loaded with file path: $global:INIPath" -LogLevel "INFO"
        Import-INIEditorData
    }
    
    $btnINIEditorClose = Get-GUIElement -Name "btnINIEditorClose" -ParentElement $tabINIEditor
    if ($btnINIEditorClose) {
        Register-GUIEventWithLogging -Element $btnINIEditorClose -EventName "Click" -ElementDescription "INI Editor Close Button" -Handler {
            $global:window.Close()
        }
    }
}
#endregion

#region Main GUI Execution
# Main code block that starts and handles the GUI dialog
Write-DebugMessage "GUI: Main GUI execution"

# Apply branding settings from INI file before showing the window
try {
    Write-DebugMessage "Applying branding settings from INI file" -LogLevel "INFO"
    
    # Set window properties
    if ($global:Config.Contains("WPFGUI")) {
        # Set window title with dynamic replacements
        $window.Title = $global:Config["WPFGUI"]["HeaderText"] -replace "{ScriptVersion}", 
            $global:Config["ScriptInfo"]["ScriptVersion"] -replace "{LastUpdate}", 
            $global:Config["ScriptInfo"]["LastUpdate"] -replace "{Author}", 
            $global:Config["ScriptInfo"]["Author"]
            
        # Set application name if specified
        if ($global:Config["WPFGUI"].Contains("APPName")) {
            $window.Title = $global:Config["WPFGUI"]["APPName"]
        }
        
        # Set window theme color if specified and valid
        if ($global:Config["WPFGUI"].Contains("ThemeColor")) {
            try {
                $color = $global:Config["WPFGUI"]["ThemeColor"]
                # Try to create a brush from the color string
                $themeColorBrush = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.ColorConverter]::ConvertFromString($color))
                $window.Background = $themeColorBrush
            } catch {
                Write-DebugMessage "Invalid ThemeColor: $color. Using named color." -LogLevel "WARNING"
                try {
                    $window.Background = [System.Windows.Media.Brushes]::$color
                } catch {
                    Write-DebugMessage "Named color also invalid. Using default." -LogLevel "ERROR"
                }
            }
        }
        
        # Set BoxColor for appropriate controls (like GroupBox, TextBox, etc.)
        if ($global:Config["WPFGUI"].Contains("BoxColor")) {
            try {
                $boxColor = $global:Config["WPFGUI"]["BoxColor"]
                $boxBrush = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.ColorConverter]::ConvertFromString($boxColor))
                
                # Function to recursively find and set background for appropriate controls
                function Set-BoxBackground {
                    param($parent)
                    
                    # Process all child elements recursively
                    for ($i = 0; $i -lt [System.Windows.Media.VisualTreeHelper]::GetChildrenCount($parent); $i++) {
                        $child = [System.Windows.Media.VisualTreeHelper]::GetChild($parent, $i)
                        
                        # Apply to specific control types
                        if ($child -is [System.Windows.Controls.GroupBox] -or 
                            $child -is [System.Windows.Controls.TextBox] -or 
                            $child -is [System.Windows.Controls.ComboBox] -or
                            $child -is [System.Windows.Controls.ListBox]) {
                            $child.Background = $boxBrush
                        }
                        
                        # Recursively process children
                        if ([System.Windows.Media.VisualTreeHelper]::GetChildrenCount($child) -gt 0) {
                            Set-BoxBackground -parent $child
                        }
                    }
                }
                
                # Apply when window is loaded to ensure visual tree is constructed
                $window.Add_Loaded({
                    Set-BoxBackground -parent $window
                })
                
            } catch {
                Write-DebugMessage "Error applying BoxColor: $($_.Exception.Message)" -LogLevel "ERROR"
            }
        }
        
        # Set RahmenColor (border color) for appropriate controls
        if ($global:Config["WPFGUI"].Contains("RahmenColor")) {
            try {
                $rahmenColor = $global:Config["WPFGUI"]["RahmenColor"]
                $rahmenBrush = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.ColorConverter]::ConvertFromString($rahmenColor))
                
                # Function to recursively find and set border color for appropriate controls
                function Set-BorderColor {
                    param($parent)
                    
                    # Process all child elements recursively
                    for ($i = 0; $i -lt [System.Windows.Media.VisualTreeHelper]::GetChildrenCount($parent); $i++) {
                        $child = [System.Windows.Media.VisualTreeHelper]::GetChild($parent, $i)
                        
                        # Apply to specific control types with BorderBrush property
                        if ($child -is [System.Windows.Controls.Border] -or 
                            $child -is [System.Windows.Controls.GroupBox] -or 
                            $child -is [System.Windows.Controls.TextBox] -or 
                            $child -is [System.Windows.Controls.ComboBox]) {
                            $child.BorderBrush = $rahmenBrush
                        }
                        
                        # Recursively process children
                        if ([System.Windows.Media.VisualTreeHelper]::GetChildrenCount($child) -gt 0) {
                            Set-BorderColor -parent $child
                        }
                    }
                }
                
                # Apply when window is loaded to ensure visual tree is constructed
                $window.Add_Loaded({
                    Set-BorderColor -parent $window
                })
                
            } catch {
                Write-DebugMessage "Error applying RahmenColor: $($_.Exception.Message)" -LogLevel "ERROR"
            }
        }
        
        # Set font properties
        if ($global:Config["WPFGUI"].Contains("FontFamily")) {
            $window.FontFamily = New-Object System.Windows.Media.FontFamily($global:Config["WPFGUI"]["FontFamily"])
        }
        
        if ($global:Config["WPFGUI"].Contains("FontSize")) {
            try {
                $window.FontSize = [double]$global:Config["WPFGUI"]["FontSize"]
            } catch {
                Write-DebugMessage "Invalid FontSize value" -LogLevel "WARNING"
            }
        }
        
        # Set background image if specified
        if ($global:Config["WPFGUI"].Contains("BackgroundImage")) {
            try {
                $bgImagePath = $global:Config["WPFGUI"]["BackgroundImage"]
                if (Test-Path $bgImagePath) {
                    $bgImage = New-Object System.Windows.Media.Imaging.BitmapImage
                    $bgImage.BeginInit()
                    $bgImage.UriSource = New-Object System.Uri($bgImagePath, [System.UriKind]::Absolute)
                    $bgImage.EndInit()
                    $window.Background = New-Object System.Windows.Media.ImageBrush($bgImage)
                }
            } catch {
                Write-DebugMessage "Error setting background image: $($_.Exception.Message)" -LogLevel "ERROR"
            }
        }
        
        # Set header logo
        $picLogo = Get-GUIElement -Name "picLogo"
        if ($picLogo -and $global:Config["WPFGUI"].Contains("HeaderLogo")) {
            try {
                $logoPath = $global:Config["WPFGUI"]["HeaderLogo"]
                if (Test-Path $logoPath) {
                    $logo = New-Object System.Windows.Media.Imaging.BitmapImage
                    $logo.BeginInit()
                    $logo.UriSource = New-Object System.Uri($logoPath, [System.UriKind]::Absolute) 
                    $logo.EndInit()
                    $picLogo.Source = $logo
                    
                    # Add click event if URL is specified
                    if ($global:Config["WPFGUI"].Contains("HeaderLogoURL")) {
                        $picLogo.Cursor = [System.Windows.Input.Cursors]::Hand
                        Register-GUIEventWithLogging -Element $picLogo -EventName "MouseLeftButtonUp" -ElementDescription "Header Logo" -Handler {
                            $url = $global:Config["WPFGUI"]["HeaderLogoURL"]
                            if (-not [string]::IsNullOrEmpty($url)) {
                                Start-Process $url
                            }
                        }
                    }
                }
            } catch {
                Write-DebugMessage "Error setting header logo: $($_.Exception.Message)" -LogLevel "ERROR"
            }
        }
        
        # Set footer website hyperlink
        $linkLabel = Get-GUIElement -Name "linkLabel"
        if ($linkLabel -and $global:Config["WPFGUI"].Contains("FooterWebseite")) {
            try {
                $websiteUrl = $global:Config["WPFGUI"]["FooterWebseite"]
                $linkLabel.Inlines.Clear()
                
                # Create a proper hyperlink with text
                $hyperlink = New-Object System.Windows.Documents.Hyperlink
                $hyperlink.NavigateUri = New-Object System.Uri($websiteUrl)
                $hyperlink.Inlines.Add($websiteUrl)
                Register-GUIEventWithLogging -Element $hyperlink -EventName "RequestNavigate" -ElementDescription "Footer Hyperlink" -Handler {
                    param($s, $e)
                    Start-Process $e.Uri.AbsoluteUri
                    $e.Handled = $true
                }
                
                $linkLabel.Inlines.Add($hyperlink)
            } catch {
                Write-DebugMessage "Error setting footer hyperlink: $($_.Exception.Message)" -LogLevel "ERROR"
            }
        }
        
        # Set footer info text
        $footerInfo = Get-GUIElement -Name "footerInfo"
        if ($footerInfo -and $global:Config["WPFGUI"].Contains("GUI_ExtraText")) {
            $footerInfo.Text = $global:Config["WPFGUI"]["GUI_ExtraText"]
        }
    }
    
    Write-DebugMessage "Branding settings applied successfully" -LogLevel "INFO"
} catch {
    Write-DebugMessage "Error applying branding settings: $($_.Exception.Message)" -LogLevel "ERROR"
    Write-DebugMessage "Stack trace: $($_.ScriptStackTrace)" -LogLevel "ERROR"
}

# Shows main window and handles any GUI initialization errors
try {
    $result = $global:window.ShowDialog()
    Write-DebugMessage "GUI started successfully, result: $result" -LogLevel "INFO"
} catch {
    Write-DebugMessage "ERROR: GUI could not be started!" -LogLevel "ERROR"
    Write-DebugMessage "Error message: $($_.Exception.Message)" -LogLevel "ERROR"
    Write-DebugMessage "Error details: $($_.InvocationInfo.PositionMessage)" -LogLevel "ERROR"
    exit 1
}

Write-DebugMessage "Main GUI execution completed." -LogLevel "INFO"