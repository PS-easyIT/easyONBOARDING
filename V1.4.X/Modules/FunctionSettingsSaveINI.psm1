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
        Copy-Item -Path $INIPath -Destination $backupPath -Force -ErrorAction Stop
        Write-DebugMessage "Created backup at: $backupPath"
        
        # Create a temporary string builder to build the INI content
        $iniContent = New-Object System.Text.StringBuilder
        
        # Read original INI to preserve comments and structure
        Write-DebugMessage "Reading original INI content from: $INIPath"
        $originalContent = Get-Content -Path $INIPath -ErrorAction Stop
        $currentSection = ""
        $inSection = $false
        $processedSections = @{ }
        $processedKeys = @{ }
        
        Write-DebugMessage "Processing original content with ${$originalContent.Count} lines"
        
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
                
                $inSection = $global:Config.Contains($currentSection)
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
                
                # If we're in a tracked section and the key exists in our config, use the updated value
                if ($inSection -and $global:Config[$currentSection].Contains($key)) {
                    $value = $global:Config[$currentSection][$key]
                    [void]$iniContent.AppendLine("$key=$value")
                    
                    # Mark this key as processed
                    $processedKeys[$currentSection][$key] = $true
                    Write-DebugMessage "Updated key: [$currentSection] $key=$value"
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
           foreach ($key in $global:Config[$sectionName].Keys) {
               if (-not ($processedKeys.ContainsKey($sectionName) -and $processedKeys[$sectionName].ContainsKey($key))) {
                   $value = $global:Config[$sectionName][$key]
                   [void]$iniContent.AppendLine("$key=$value")
                   Write-DebugMessage "Added new key: [$sectionName] $key=$value"
               }
           }
       }
       
       # Save the content back to the file
       Write-DebugMessage "Saving new content to INI file..."
       $finalContent = $iniContent.ToString()
       [System.IO.File]::WriteAllText($INIPath, $finalContent, [System.Text.Encoding]::UTF8)
       
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