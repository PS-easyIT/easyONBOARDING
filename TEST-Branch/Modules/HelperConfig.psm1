# Helper functions for configuration

function Get-IniContent {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$Path
    )
    
    if (Get-Command -Name "Write-DebugOutput" -ErrorAction SilentlyContinue) {
        Write-DebugOutput "Reading INI content from: $Path" -Level INFO
    } else {
        Write-Host "Reading INI content from: $Path"
    }
    
    try {
        $ini = @{}
        $section = "Default"
        $ini[$section] = @{}
        
        if (Test-Path -Path $Path) {
            switch -regex -file $Path {
                "^\[(.+)\]$" {
                    $section = $matches[1]
                    $ini[$section] = @{}
                }
                "^\s*([^#].+?)\s*=\s*(.*)" {
                    $name,$value = $matches[1..2]
                    $ini[$section][$name] = $value
                }
            }
        } else {
            if (Get-Command -Name "Write-DebugOutput" -ErrorAction SilentlyContinue) {
                Write-DebugOutput "INI file not found: $Path" -Level WARNING
            } else {
                Write-Host "INI file not found: $Path" -ForegroundColor Yellow
            }
            
            # Create default settings
            $ini["WPFGUI"] = @{
                "HeaderText" = "easyONBOARDING {ScriptVersion}"
                "APPName" = "easyONBOARDING Tool"
                "ThemeColor" = "LightGray"
                "BoxColor" = "White"
                "RahmenColor" = "Gray"
                "FontFamily" = "Segoe UI"
                "FontSize" = "12"
            }
            
            $ini["ScriptInfo"] = @{
                "ScriptVersion" = "1.4.1"
                "LastUpdate" = (Get-Date -Format "yyyy-MM-dd")
                "Author" = "PowerShell Admin"
            }
        }
        
        if (Get-Command -Name "Write-DebugOutput" -ErrorAction SilentlyContinue) {
            Write-DebugOutput "Loaded INI with $($ini.Count) sections" -Level INFO
        } else {
            Write-Host "Loaded INI with $($ini.Count) sections" -ForegroundColor Green
        }
        
        return $ini
    }
    catch {
        if (Get-Command -Name "Write-DebugOutput" -ErrorAction SilentlyContinue) {
            Write-DebugOutput "Error reading INI file: $($_.Exception.Message)" -Level ERROR
        } else {
            Write-Host "Error reading INI file: $($_.Exception.Message)" -ForegroundColor Red
        }
        return @{}
    }
}

function Set-IniContent {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [hashtable]$InputObject,
        
        [Parameter(Mandatory=$true)]
        [string]$Path
    )
    
    if (Get-Command -Name "Write-DebugOutput" -ErrorAction SilentlyContinue) {
        Write-DebugOutput "Writing INI content to: $Path" -Level INFO
    } else {
        Write-Host "Writing INI content to: $Path"
    }
    
    try {
        $output = ""
        
        foreach ($section in $InputObject.Keys) {
            $output += "[$section]`r`n"
            
            foreach ($key in $InputObject[$section].Keys) {
                $value = $InputObject[$section][$key]
                $output += "$key=$value`r`n"
            }
            
            $output += "`r`n"
        }
        
        $output | Out-File -FilePath $Path -Encoding UTF8
        
        if (Get-Command -Name "Write-DebugOutput" -ErrorAction SilentlyContinue) {
            Write-DebugOutput "INI file successfully saved" -Level INFO
        } else {
            Write-Host "INI file successfully saved" -ForegroundColor Green
        }
        
        return $true
    }
    catch {
        if (Get-Command -Name "Write-DebugOutput" -ErrorAction SilentlyContinue) {
            Write-DebugOutput "Error writing INI file: $($_.Exception.Message)" -Level ERROR
        } else {
            Write-Host "Error writing INI file: $($_.Exception.Message)" -ForegroundColor Red
        }
        return $false
    }
}

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
Export-ModuleMember -Function Get-ConfigValue, Test-ConfigSection, Get-IniContent, Set-IniContent