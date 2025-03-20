function Import-INIEditorData {
    try {
        Write-DebugMessage "Loading INI sections and settings for the editor"
        
        # Get the ListView and DataGrid controls
        $listViewINIEditor = $tabINIEditor.FindName("listViewINIEditor")
        $dataGridINIEditor = $tabINIEditor.FindName("dataGridINIEditor")
        
        if ($null -eq $listViewINIEditor -or $null -eq $dataGridINIEditor) {
            Write-DebugMessage "ListView or DataGrid not found in XAML"
            return
        }
        
        # Create a collection to store the sections
        $sectionItems = New-Object System.Collections.ObjectModel.ObservableCollection[PSObject]
        
        # Add all sections from the global Config
        foreach ($sectionName in $global:Config.Keys) {
            $sectionItems.Add([PSCustomObject]@{
                SectionName = $sectionName
            })
        }
        
        # Set the ListView ItemsSource
        $listViewINIEditor.ItemsSource = $sectionItems
        
        # Ensure ListView displays section names
        if ($listViewINIEditor.View -is [System.Windows.Controls.GridView]) {
            # GridView already defined in XAML
        } else {
            # Create GridView programmatically
            $gridView = New-Object System.Windows.Controls.GridView
            $column = New-Object System.Windows.Controls.GridViewColumn
            $column.Header = "Section"
            $column.DisplayMemberBinding = New-Object System.Windows.Data.Binding("SectionName")
            $gridView.Columns.Add($column)
            $listViewINIEditor.View = $gridView
        }
        
        # Ensure DataGrid has columns defined
        if ($dataGridINIEditor.Columns.Count -eq 0) {
            # Add columns programmatically
            $keyColumn = New-Object System.Windows.Controls.DataGridTextColumn
            $keyColumn.Header = "Key"
            $keyColumn.Binding = New-Object System.Windows.Data.Binding("Key")
            $keyColumn.Width = New-Object System.Windows.Controls.DataGridLength(1, [System.Windows.Controls.DataGridLengthUnitType]::Star)
            $dataGridINIEditor.Columns.Add($keyColumn)
            
            $valueColumn = New-Object System.Windows.Controls.DataGridTextColumn
            $valueColumn.Header = "Value"
            $valueColumn.Binding = New-Object System.Windows.Data.Binding("Value") 
            $valueColumn.Width = New-Object System.Windows.Controls.DataGridLength(2, [System.Windows.Controls.DataGridLengthUnitType]::Star)
            $dataGridINIEditor.Columns.Add($valueColumn)
            
            $commentColumn = New-Object System.Windows.Controls.DataGridTextColumn
            $commentColumn.Header = "Comment"
            $commentColumn.Binding = New-Object System.Windows.Data.Binding("Comment")
            $commentColumn.Width = New-Object System.Windows.Controls.DataGridLength(2, [System.Windows.Controls.DataGridLengthUnitType]::Star)
            $dataGridINIEditor.Columns.Add($commentColumn)
        }
        
        Write-DebugMessage "Loaded $($sectionItems.Count) INI sections"
    }
    catch {
        Write-DebugMessage "Error loading INI editor data: $($_.Exception.Message)"
        [System.Windows.MessageBox]::Show("Error loading INI data: $($_.Exception.Message)", "Error", "OK", "Error")
    }
}

function Import-SectionSettings {
    param(
        [Parameter(Mandatory=$true)]
        [string]$SectionName,
        
        [Parameter(Mandatory=$true)]
        [System.Windows.Controls.DataGrid]$DataGrid,
        
        [Parameter(Mandatory=$false)]
        [string]$INIPath
    )
    
    try {
        Write-DebugMessage "Loading settings for section: $SectionName"
        
        if ($null -eq $DataGrid) {
            Write-DebugMessage "DataGrid parameter is null"
            return
        }
        
        # Get the section data from the global Config
        $sectionData = $global:Config[$SectionName]
        if ($null -eq $sectionData) {
            Write-DebugMessage "Section $SectionName not found in Config"
            return
        }
        
        Write-DebugMessage "Found section with $($sectionData.Count) keys"
        
        # Create a collection to store the key-value pairs
        $settingsItems = New-Object System.Collections.ObjectModel.ObservableCollection[PSObject]
        
        # Extract comments from the original INI file if available
        $commentsByKey = @{ }
        try {
            if ([string]::IsNullOrEmpty($INIPath)) {
                $INIPath = $global:INIPath
            }
            
            $iniContent = Get-Content -Path $INIPath -ErrorAction SilentlyContinue
            $currentSection = ""
            $keyComment = ""
            
            foreach ($line in $iniContent) {
                $line = $line.Trim()
                
                # Check if it's a section header
                if ($line -match '^\[(.+)\]$') {
                    $currentSection = $matches[1].Trim()
                    continue
                }
                
                # Check if it's a comment
                if ($line -match '^[;#](.*)$') {
                    $keyComment = $matches[1].Trim()
                    continue
                }
                
                # Check if it's a key-value pair
                if ($line -match '^(.*?)=(.*)$' -and $currentSection -eq $SectionName) {
                    $key = $matches[1].Trim()
                    $commentsByKey[$key] = $keyComment
                    $keyComment = ""  # Reset for next key
                }
            }
        }
        catch {
            Write-DebugMessage "Error extracting comments from INI: $($_.Exception.Message)"
        }
        
        # Add all key-value pairs from the section
        foreach ($key in $sectionData.Keys) {
            $value = $sectionData[$key]
            $comment = if ($commentsByKey.ContainsKey($key)) { $commentsByKey[$key] } else { "" }
            
            Write-DebugMessage "Adding key: '$key' with value: '$value'"
            
            $settingsItems.Add([PSCustomObject]@{
                Key = $key
                Value = $value
                Comment = $comment
                OriginalKey = $key  # Store original key in case it's edited
            })
        }
        
        # Clear current data and set the new DataGrid ItemsSource
        $DataGrid.ItemsSource = $null
        $DataGrid.Items.Clear()
        $DataGrid.ItemsSource = $settingsItems
        
        Write-DebugMessage "Loaded $($settingsItems.Count) settings for section $SectionName"
        
        # Force UI update
        $DataGrid.UpdateLayout()
    }
    catch {
        Write-DebugMessage "Error loading section settings: $($_.Exception.Message)"
        [System.Windows.MessageBox]::Show("Error loading settings for section '$SectionName': $($_.Exception.Message)", "Error", "OK", "Error")
    }
}

