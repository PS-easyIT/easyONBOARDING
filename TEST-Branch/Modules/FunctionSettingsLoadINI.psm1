# Function to open the INI configuration editor
Write-DebugMessage "Opening INI editor."
function Open-INIEditor {
    try {
        # Check if $global:Config is defined
        if ($global:Config) {
            if ($global:Config.Logging.DebugMode -eq "1") {
                # Check if Write-Log function exists before calling it
                if (Get-Command Write-Log -ErrorAction SilentlyContinue) {
                    Write-Log " INI Editor is starting." "DEBUG"
                } else {
                    Write-Warning "Write-Log function not found. Debug logging will be skipped."
                }
            }
        } else {
            Write-Warning "\$global:Config is not defined. Please ensure the configuration is loaded."
        }
        Write-DebugMessage "INI Editor has started."
        Write-LogMessage -Message "INI Editor has started." -Level "Info"
        [System.Windows.MessageBox]::Show("INI Editor (Settings) – Implement functionality here.", "Settings")
    }
    catch {
        Throw "Error opening INI Editor: $_"
    }
}

# Functions for loading INI settings

function Load-INISettings {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$false)]
        [string]$INIPath = $global:INIPath
    )
    
    try {
        if (Get-Command -Name "Write-DebugOutput" -ErrorAction SilentlyContinue) {
            Write-DebugOutput "Loading INI settings from $INIPath" -Level INFO
        }
        
        if (-not (Test-Path -Path $INIPath)) {
            if (Get-Command -Name "Write-DebugOutput" -ErrorAction SilentlyContinue) {
                Write-DebugOutput "INI file not found: $INIPath" -Level WARNING
            }
            
            # Create a new INI file with default settings
            $defaultSettings = @{
                "WPFGUI" = @{
                    "HeaderText" = "easyONBOARDING {ScriptVersion}"
                    "APPName" = "easyONBOARDING Tool"
                    "ThemeColor" = "LightGray"
                    "BoxColor" = "White"
                    "RahmenColor" = "Gray"
                    "FontFamily" = "Segoe UI"
                    "FontSize" = "12"
                    "GUI_ExtraText" = "easyONBOARDING Tool by PowerShell Admin"
                }
                "ScriptInfo" = @{
                    "ScriptVersion" = "1.4.1"
                    "LastUpdate" = (Get-Date -Format "yyyy-MM-dd")
                    "Author" = "PowerShell Admin"
                }
                "DefaultValues" = @{
                    "DefaultDomain" = "example.com"
                    "DefaultOU" = "OU=Users,DC=example,DC=com"
                    "DefaultUPNTemplate" = "Vorname.Nachname"
                }
                "Report" = @{
                    "LogoPath" = ""
                    "CompanyName" = "Example Company"
                    "ReportHeader" = "User Onboarding Report"
                }
            }
            
            # Save the default settings
            if (Get-Command -Name "Set-IniContent" -ErrorAction SilentlyContinue) {
                Set-IniContent -InputObject $defaultSettings -Path $INIPath
            }
            else {
                # Manual save if Set-IniContent not available
                $output = ""
                foreach ($section in $defaultSettings.Keys) {
                    $output += "[$section]`r`n"
                    foreach ($key in $defaultSettings[$section].Keys) {
                        $value = $defaultSettings[$section][$key]
                        $output += "$key=$value`r`n"
                    }
                    $output += "`r`n"
                }
                $output | Out-File -FilePath $INIPath -Encoding UTF8
            }
            
            if (Get-Command -Name "Write-DebugOutput" -ErrorAction SilentlyContinue) {
                Write-DebugOutput "Created new INI file with default settings" -Level INFO
            }
        }
        
        # Load the INI settings
        if (Get-Command -Name "Get-IniContent" -ErrorAction SilentlyContinue) {
            $global:Config = Get-IniContent -Path $INIPath
        }
        else {
            # Manual load if Get-IniContent not available
            $global:Config = @{}
            $section = "Default"
            $global:Config[$section] = @{}
            
            switch -regex -file $INIPath {
                "^\[(.+)\]$" {
                    $section = $matches[1]
                    $global:Config[$section] = @{}
                }
                "^\s*([^#].+?)\s*=\s*(.*)" {
                    $name,$value = $matches[1..2]
                    $global:Config[$section][$name] = $value
                }
            }
        }
        
        if (Get-Command -Name "Write-DebugOutput" -ErrorAction SilentlyContinue) {
            Write-DebugOutput "INI settings loaded successfully with $($global:Config.Count) sections" -Level INFO
        }
        
        return $true
    }
    catch {
        if (Get-Command -Name "Write-DebugOutput" -ErrorAction SilentlyContinue) {
            Write-DebugOutput "Error loading INI settings: $($_.Exception.Message)" -Level ERROR
        }
        return $false
    }
}

function Apply-INIDefaults {
    [CmdletBinding()]
    param()
    
    try {
        if (Get-Command -Name "Write-DebugOutput" -ErrorAction SilentlyContinue) {
            Write-DebugOutput "Applying INI defaults to the interface" -Level INFO
        }
        
        if (-not $global:Config -or -not $global:Config.Contains("DefaultValues")) {
            if (Get-Command -Name "Write-DebugOutput" -ErrorAction SilentlyContinue) {
                Write-DebugOutput "No DefaultValues section found in configuration" -Level WARNING
            }
            return $false
        }
        
        # Apply domain default
        if ($global:Config["DefaultValues"].Contains("DefaultDomain")) {
            $cmbSuffix = $global:window.FindName("cmbSuffix")
            if ($cmbSuffix) {
                $defaultDomain = $global:Config["DefaultValues"]["DefaultDomain"]
                
                # Find and select the default domain
                for ($i = 0; $i -lt $cmbSuffix.Items.Count; $i++) {
                    if ($cmbSuffix.Items[$i] -eq $defaultDomain) {
                        $cmbSuffix.SelectedIndex = $i
                        break
                    }
                }
            }
        }
        
        # Apply OU default
        if ($global:Config["DefaultValues"].Contains("DefaultOU")) {
            $cmbOU = $global:window.FindName("cmbOU")
            if ($cmbOU) {
                $defaultOU = $global:Config["DefaultValues"]["DefaultOU"]
                
                # Find and select the default OU
                for ($i = 0; $i -lt $cmbOU.Items.Count; $i++) {
                    if ($cmbOU.Items[$i] -eq $defaultOU) {
                        $cmbOU.SelectedIndex = $i
                        break
                    }
                }
            }
        }
        
        # Apply UPN template default
        if ($global:Config["DefaultValues"].Contains("DefaultUPNTemplate")) {
            $cmbDisplayTemplate = $global:window.FindName("cmbDisplayTemplate")
            if ($cmbDisplayTemplate) {
                $defaultUPNTemplate = $global:Config["DefaultValues"]["DefaultUPNTemplate"]
                
                # Find and select the default UPN template
                for ($i = 0; $i -lt $cmbDisplayTemplate.Items.Count; $i++) {
                    if ($cmbDisplayTemplate.Items[$i] -eq $defaultUPNTemplate) {
                        $cmbDisplayTemplate.SelectedIndex = $i
                        break
                    }
                }
            }
        }
        
        if (Get-Command -Name "Write-DebugOutput" -ErrorAction SilentlyContinue) {
            Write-DebugOutput "INI defaults applied successfully" -Level INFO
        }
        
        return $true
    }
    catch {
        if (Get-Command -Name "Write-DebugOutput" -ErrorAction SilentlyContinue) {
            Write-DebugOutput "Error applying INI defaults: $($_.Exception.Message)" -Level ERROR
        }
        return $false
    }
}

# Export the function to make it available outside the module
Export-ModuleMember -Function Open-INIEditor, Load-INISettings, Apply-INIDefaults

