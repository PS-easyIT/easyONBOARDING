# Helper functions for configuration data
Write-DebugMessage "Loading configuration helpers."
function Get-ConfigValue {
    [CmdletBinding()]
    [OutputType([string])]
    param (
        [Parameter(Mandatory=$true)]
        [hashtable]$Section,
        
        [Parameter(Mandatory=$true)]
        [string]$Key,
        
        [Parameter(Mandatory=$false)]
        [string]$DefaultValue = ""
    )
    
    try {
        if ($Section.ContainsKey($Key) -and -not [string]::IsNullOrEmpty($Section[$Key])) {
            return $Section[$Key]
        }
        else {
            Write-DebugMessage "Key '$Key' not found in section or has empty value. Using default: '$DefaultValue'"
            return $DefaultValue
        }
    }
    catch {
        Write-ErrorMessage "Error in Get-ConfigValue: $($_.Exception.Message)"
        return $DefaultValue
    }
}

function Test-ConfigSection {
    [CmdletBinding()]
    [OutputType([bool])]
    param (
        [Parameter(Mandatory=$true)]
        [hashtable]$Config,
        
        [Parameter(Mandatory=$true)]
        [string]$SectionName,
        
        [Parameter(Mandatory=$false)]
        [string[]]$RequiredKeys = @()
    )
    
    try {
        # Check if section exists
        if (-not $Config.ContainsKey($SectionName)) {
            Write-DebugMessage "Section '$SectionName' not found in configuration"
            return $false
        }
        
        # If specific keys are required, check them too
        if ($RequiredKeys.Count -gt 0) {
            $section = $Config[$SectionName]
            foreach ($key in $RequiredKeys) {
                if (-not $section.ContainsKey($key) -or [string]::IsNullOrEmpty($section[$key])) {
                    Write-DebugMessage "Required key '$key' not found in section '$SectionName' or has empty value"
                    return $false
                }
            }
        }
        
        return $true
    }
    catch {
        Write-ErrorMessage "Error in Test-ConfigSection: $($_.Exception.Message)"
        return $false
    }
}

# Export the functions to make them available outside the module
Export-ModuleMember -Function Get-ConfigValue, Test-ConfigSection