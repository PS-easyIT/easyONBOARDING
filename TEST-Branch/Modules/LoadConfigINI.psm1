# Lädt Konfiguration aus einer INI-Datei
# Version: 1.4.1 - Optimized

# Ensure Write-DebugMessage function is available
if (-not (Get-Command -Name Write-DebugMessage -ErrorAction SilentlyContinue)) {
    function Write-DebugMessage {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory=$true)]
            [string]$Message,
            [string]$LogLevel = "DEBUG"
        )
        Write-Verbose $Message
    }
}

# Ensure Write-LogMessage function is available
if (-not (Get-Command -Name Write-LogMessage -ErrorAction SilentlyContinue)) {
    function Write-LogMessage {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory=$true)]
            [string]$Message,
            [string]$LogLevel = "INFO"
        )
        Write-Output "[$LogLevel] $Message"
    }
}

Write-DebugMessage "Loading INI configuration."

function Get-IniContent {
    [CmdletBinding()]
    [OutputType([System.Collections.Specialized.OrderedDictionary])]
    param(
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$Path,
        
        [switch]$Ordered
    )
    
    try {
        Write-DebugMessage "Reading INI file: $Path"
        if (-not (Test-Path -Path $Path -PathType Leaf)) {
            throw "INI file not found: $Path"
        }
        
        # Use ordered dictionary to preserve section order
        $ini = if ($Ordered) { 
            [ordered]@{} 
        } else { 
            @{} 
        }
        
        $section = "Default"
        $ini[$section] = @{}
        
        # Read the file line by line
        $content = Get-Content -Path $Path -ErrorAction Stop
        
        foreach ($line in $content) {
            $line = $line.Trim()
            
            # Skip empty lines and comments
            if ([string]::IsNullOrWhiteSpace($line) -or $line.StartsWith(";") -or $line.StartsWith("#")) {
                continue
            }
            
            # Section header
            if ($line -match '^\[(.+)\]') {
                $section = $matches[1].Trim()
                if (-not $ini.ContainsKey($section)) {
                    $ini[$section] = @{}
                }
                continue
            }
            
            # Key-value pair
            if ($line -match '^([^=]+)=(.*)') {
                $key = $matches[1].Trim()
                $value = $matches[2].Trim()
                
                if ($key -and $value) {
                    if (-not $ini.ContainsKey($section)) {
                        $ini[$section] = @{}
                    }
                    $ini[$section][$key] = $value
                }
            }
        }
        
        Write-DebugMessage "INI file loaded successfully with $($ini.Keys.Count) sections."
        return $ini
    }
    catch {
        Write-DebugMessage "Error loading INI file: $($_.Exception.Message)"
        throw "Error loading INI file: $($_.Exception.Message)"
    }
}

function Set-IniContent {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$Path,
        
        [Parameter(Mandatory=$true)]
        [ValidateNotNull()]
        [System.Collections.IDictionary]$Content
    )
    
    try {
        Write-DebugMessage "Writing INI file: $Path"
        
        $output = New-Object System.Collections.Generic.List[string]
        
        foreach ($section in $Content.Keys) {
            # Add section header
            $output.Add("[$section]")
            
            # Add key-value pairs
            if ($Content[$section] -is [System.Collections.IDictionary]) {
                foreach ($key in $Content[$section].Keys) {
                    $value = $Content[$section][$key]
                    $output.Add("$key=$value")
                }
            }
            
            # Add a blank line after each section
            $output.Add("")
        }
        
        # Write the file
        Set-Content -Path $Path -Value $output -ErrorAction Stop
        Write-DebugMessage "INI file written successfully."
        return $true
    }
    catch {
        Write-DebugMessage "Error writing INI file: $($_.Exception.Message)"
        throw "Error writing INI file: $($_.Exception.Message)"
    }
}

# If INI path is specified in parameters, use that instead
if ($PSBoundParameters.ContainsKey('ConfigPath') -and -not [string]::IsNullOrEmpty($ConfigPath)) {
    $global:INIPath = $ConfigPath
    Write-DebugMessage "INI path set from parameter: $global:INIPath"
}

# Validate INI path
if (-not $global:INIPath) {
    $scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
    $global:INIPath = Join-Path -Path (Split-Path -Parent $scriptDir) -ChildPath 'easyONB.ini'
    Write-DebugMessage "INI path set to default: $global:INIPath"
}

# Load the INI file
try {
    $global:Config = Get-IniContent -Path $global:INIPath
    Write-DebugMessage "Global Config loaded from: $global:INIPath"
    
    # Set debug mode from config if available
    if ($global:Config.Contains("Logging") -and $global:Config.Logging.Contains("DebugMode")) {
        $debugMode = [int]::Parse($global:Config.Logging.DebugMode)
        $global:DebugEnabled = $debugMode -gt 0
        Write-DebugMessage "Debug mode set from INI: $($global:DebugEnabled)"
    }
} catch {
    Write-LogMessage -Message "Failed to load INI file: $($_.Exception.Message)" -LogLevel "ERROR"
    throw "Critical Error: Failed to load configuration. $($_.Exception.Message)"
}

Export-ModuleMember -Function Get-IniContent, Set-IniContent