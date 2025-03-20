# Loads configuration from an INI file

Write-DebugMessage "Loading INI configuration."
function Get-IniContent {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Path
    )
    
    try {
        Write-DebugMessage "Reading INI file: $Path"
        $ini = @{}
        if (Test-Path -Path $Path) {
            $section = "Default"
            $ini[$section] = @{}
            
            switch -regex -file $Path {
                "^\[(.+)\]" {
                    $section = $matches[1].Trim()
                    $ini[$section] = @{}
                }
                "^(.*?)=(.*)" {
                    $name, $value = $matches[1..2]
                    $name = $name.Trim()
                    $value = $value.Trim()
                    
                    # Handle special case where value contains =
                    if ($name -and $value) {
                        if (-not $ini[$section]) {
                            $ini[$section] = @{}
                        }
                        $ini[$section][$name] = $value
                    }
                }
            }
        }
        else {
            Write-DebugMessage "INI file not found: $Path"
            throw "INI file not found: $Path"
        }
        
        Write-DebugMessage "INI file loaded successfully."
        return $ini
    }
    catch {
        Write-DebugMessage "Error loading INI file: $($_.Exception.Message)"
        throw "Error loading INI file: $($_.Exception.Message)"
    }
}

# If INI path is specified in parameters, use that instead
if ($PSBoundParameters.ContainsKey('ConfigPath') -and -not [string]::IsNullOrEmpty($ConfigPath)) {
    $global:INIPath = $ConfigPath
}

# Load the INI file
try {
    $global:Config = Get-IniContent -Path $global:INIPath
    Write-DebugMessage "Global Config loaded from: $global:INIPath"
} catch {
    Write-LogMessage -Message "Failed to load INI file: $($_.Exception.Message)" -LogLevel "ERROR"
    throw "Critical Error: Failed to load configuration. $($_.Exception.Message)"
}