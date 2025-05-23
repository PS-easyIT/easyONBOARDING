# Check if required functions exist, if not create stub functions
if (-not (Get-Command -Name "Write-DebugMessage" -ErrorAction SilentlyContinue)) {
    function Write-DebugMessage { param([string]$Message) Write-Verbose $Message }
}
if (-not (Get-Command -Name "Write-LogMessage" -ErrorAction SilentlyContinue)) {
    function Write-LogMessage { param([string]$Message, [string]$LogLevel) Write-Verbose $Message }
}

function Save-INIChanges {
    param(
        [Parameter(Mandatory=$false)]
        [string]$INIPath
    )
    
    try {
        Write-DebugMessage "Starting Save-INIChanges..."
        # Make sure INIPath is defined
        if ([string]::IsNullOrEmpty($INIPath)) {
            $INIPath = $global:INIPath
            Write-DebugMessage "Set INI path from global: $INIPath"
        }
        
        if (-not (Test-Path -Path $INIPath)) {
            Write-DebugMessage "INI file not found: $INIPath"
            throw "INI file not found at: $INIPath"
        }
        
        # Create a backup before making changes
        $backupPath = "$INIPath.backup"
        try {
            Copy-Item -Path $INIPath -Destination $backupPath -Force -ErrorAction Stop
            Write-DebugMessage "Created backup at: $backupPath"
        } catch {
            Write-DebugMessage "Error creating backup: $($_.Exception.Message)"
            Write-LogMessage -Message "Error creating INI backup file: $($_.Exception.Message)" -LogLevel "ERROR"
            # Consider whether to continue if backup fails.  For now, continue.
        }
        
        # Create a temporary string builder to build the INI content
        $iniContent = New-Object System.Text.StringBuilder
        
        # Read original INI to preserve comments and structure
        Write-DebugMessage "Reading original INI content from: $INIPath"
        try {
            $originalContent = Get-Content -Path $INIPath -ErrorAction Stop
        } catch {
            Write-DebugMessage "Error reading INI file: $($_.Exception.Message)"
            Write-LogMessage -Message "Error reading INI file: $($_.Exception.Message)" -LogLevel "ERROR"
            return $false # Or handle the error as appropriate
        }
        $currentSection = ""
        $inSection = $false
        $processedSections = @{ }
        $processedKeys = @{ }
        
        Write-DebugMessage "Processing original content with $($originalContent.Count) lines"
        
        # First pass: Preserve structure and update existing values
        foreach ($line in $originalContent) {
            $trimmedLine = $line.Trim()
            
            # Check if it's a section header
            if ($trimmedLine -match '^\[(.+)\]$') {
                $currentSection = $matches[1].Trim()
                Write-DebugMessage "Found section: [$currentSection]"
                
                # Initialize tracking for this section if needed
                if (-not $processedKeys.ContainsKey($currentSection)) {
                    $processedKeys[$currentSection] = @{ }
                }
                
                $inSection = $global:Config.ContainsKey($currentSection)
                $processedSections[$currentSection] = $true
                
                # Add the section header
                [void]$iniContent.AppendLine($line)
                continue
            }
            
            # Check if it's a comment or empty line
            if ([string]::IsNullOrWhiteSpace($trimmedLine) -or $trimmedLine -match '^[;#]') {
                [void]$iniContent.AppendLine($line)
                continue
            }
            
            # Check if it's a key-value pair
            if ($trimmedLine -match '^(.*?)=(.*)$') {
                $key = $matches[1].Trim()
                $value = $matches[2].Trim()
                
                # If we're in a tracked section and the key exists in our config, use the updated value
                if ($inSection -and $global:Config[$currentSection].ContainsKey($key)) {
                    $newValue = $global:Config[$currentSection][$key]
                    [void]$iniContent.AppendLine("$key=$newValue")
                    
                    # Mark this key as processed
                    $processedKeys[$currentSection][$key] = $true
                    Write-DebugMessage "Updated key: [$currentSection] $key=$newValue"
                }
                else {
                    if ($inSection) {
                        Write-DebugMessage "Key no longer exists in config: [$currentSection] $key"
                    }
                    # Keep the original line for keys not in our config or not in tracked sections
                    [void]$iniContent.AppendLine($line)
                }
            }
            else {
                # Keep other lines as-is
                [void]$iniContent.AppendLine($line)
            }
        }
       # Second pass: Add any new sections or keys that weren't in the original file
       Write-DebugMessage "Adding new sections and keys..."
       foreach ($sectionName in $global:Config.Keys) {
           # Add new section if it wasn't in the original file
           if (-not $processedSections.ContainsKey($sectionName)) {
               Write-DebugMessage "Adding new section: [$sectionName]"
               [void]$iniContent.AppendLine("")
               [void]$iniContent.AppendLine("[$sectionName]")
               
               # Initialize tracking for this section
               if (-not $processedKeys.ContainsKey($sectionName)) {
                   $processedKeys[$sectionName] = @{ }
               }
           }
           
           # Add any keys that weren't in the original file
           if ($global:Config.ContainsKey($sectionName)) {
                if (-not $processedKeys.ContainsKey($sectionName)) {
                    $processedKeys[$sectionName] = @{} 
                }
                foreach ($key in $global:Config[$sectionName].Keys) {
                    if (-not $processedKeys[$sectionName].ContainsKey($key)) {
                        $value = $global:Config[$sectionName][$key]
                        [void]$iniContent.AppendLine("$key=$value")
                        Write-DebugMessage "Added new key: [$sectionName] $key=$value"
                    }
                }
           }
       }
       
       # Save the content back to the file
       Write-DebugMessage "Saving new content to INI file..."
       try {
            $finalContent = $iniContent.ToString()
            [System.IO.File]::WriteAllText($INIPath, $finalContent, [System.Text.Encoding]::UTF8)
       } catch {
            Write-DebugMessage "Error writing to INI file: $($_.Exception.Message)"
            Write-LogMessage -Message "Error writing to INI file: $($_.Exception.Message)" -LogLevel "ERROR"
            return $false
       }
       
       Write-DebugMessage "INI file saved successfully: $INIPath"
       Write-LogMessage -Message "INI file edited by $($env:USERNAME)" -LogLevel "INFO"
       
       return $true
   }
   catch {
       Write-DebugMessage "Error saving INI changes: $($_.Exception.Message)"
       Write-DebugMessage "Stack trace: $($_.ScriptStackTrace)"
       Write-LogMessage -Message "Error saving INI file: $($_.Exception.Message)" -LogLevel "ERROR"
       return $false
   }
}

# Export the function so it's available outside this module.
Export-ModuleMember -Function Save-INIChanges