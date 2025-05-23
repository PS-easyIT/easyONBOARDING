function Show-SearchDialog {
    param(
        [string]$InitialFilter = "",
        [scriptblock]$OnSearch = {}
    )
    
    # Hauptsuchfenster erstellen
    $searchWindow = New-Object System.Windows.Window
    $searchWindow.Title = "Erweiterte Suche"
    $searchWindow.Width = 600
    $searchWindow.Height = 400
    $searchWindow.WindowStartupLocation = "CenterScreen"
    $searchWindow.Background = "#F0F0F0"
    $searchWindow.ResizeMode = "CanResize"
    $searchWindow.MinWidth = 500
    $searchWindow.MinHeight = 350
    
    # Hauptcontainer
    $mainGrid = New-Object System.Windows.Controls.Grid
    $mainGrid.Margin = New-Object System.Windows.Thickness(15)
    
    # Zeilen definieren
    $mainGrid.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition -Property @{Height = "Auto"}))
    $mainGrid.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition -Property @{Height = "Auto"}))
    $mainGrid.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition -Property @{Height = "Auto"}))
    $mainGrid.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition -Property @{Height = "Auto"}))
    $mainGrid.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition -Property @{Height = "*"}))
    $mainGrid.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition -Property @{Height = "Auto"}))
    
    # Titel
    $title = New-Object System.Windows.Controls.Label
    $title.Content = "Erweiterte Such- und Filteroptionen"
    $title.FontSize = 16
    $title.FontWeight = "Bold"
    $title.Foreground = "#5B87A9"
    $title.Margin = New-Object System.Windows.Thickness(0, 0, 0, 10)
    [System.Windows.Controls.Grid]::SetRow($title, 0)
    $mainGrid.Children.Add($title)
    
    # Schnellsuchbereich
    $quickSearchPanel = New-Object System.Windows.Controls.StackPanel
    $quickSearchPanel.Orientation = "Horizontal"
    $quickSearchPanel.Margin = New-Object System.Windows.Thickness(0, 5, 0, 10)
    [System.Windows.Controls.Grid]::SetRow($quickSearchPanel, 1)
    
    $quickSearchLabel = New-Object System.Windows.Controls.Label
    $quickSearchLabel.Content = "Schnellsuche:"
    $quickSearchLabel.Width = 100
    $quickSearchLabel.VerticalContentAlignment = "Center"
    $quickSearchPanel.Children.Add($quickSearchLabel)
    
    $quickSearchBox = New-Object System.Windows.Controls.TextBox
    $quickSearchBox.Margin = New-Object System.Windows.Thickness(0, 0, 10, 0)
    $quickSearchBox.Height = 30
    $quickSearchBox.MinWidth = 200
    $quickSearchBox.VerticalContentAlignment = "Center"
    $quickSearchBox.Text = $InitialFilter
    $quickSearchBox.HorizontalAlignment = "Stretch"
    $quickSearchPanel.Children.Add($quickSearchBox)
    
    $searchButton = New-Object System.Windows.Controls.Button
    $searchButton.Content = "Suchen"
    $searchButton.Height = 30
    $searchButton.Width = 100
    $searchButton.Background = "#5B87A9"
    $searchButton.Foreground = "White"
    $searchButton.FontWeight = "SemiBold"
    $searchPanel.Children.Add($searchButton)
    
    $mainGrid.Children.Add($quickSearchPanel)
    
    # Datumsfilterbereich
    $dateFilterPanel = New-Object System.Windows.Controls.StackPanel
    $dateFilterPanel.Orientation = "Horizontal"
    $dateFilterPanel.Margin = New-Object System.Windows.Thickness(0, 5, 0, 10)
    [System.Windows.Controls.Grid]::SetRow($dateFilterPanel, 2)
    
    $dateFilterLabel = New-Object System.Windows.Controls.Label
    $dateFilterLabel.Content = "Zeitraum:"
    $dateFilterLabel.Width = 100
    $dateFilterLabel.VerticalContentAlignment = "Center"
    $dateFilterPanel.Children.Add($dateFilterLabel)
    
    $startDateLabel = New-Object System.Windows.Controls.Label
    $startDateLabel.Content = "Von:"
    $startDateLabel.Margin = New-Object System.Windows.Thickness(0, 0, 5, 0)
    $startDateLabel.VerticalContentAlignment = "Center"
    $dateFilterPanel.Children.Add($startDateLabel)
    
    $startDatePicker = New-Object System.Windows.Controls.DatePicker
    $startDatePicker.Margin = New-Object System.Windows.Thickness(0, 0, 15, 0)
    $startDatePicker.Width = 120
    $dateFilterPanel.Children.Add($startDatePicker)
    
    $endDateLabel = New-Object System.Windows.Controls.Label
    $endDateLabel.Content = "Bis:"
    $endDateLabel.Margin = New-Object System.Windows.Thickness(0, 0, 5, 0)
    $endDateLabel.VerticalContentAlignment = "Center"
    $dateFilterPanel.Children.Add($endDateLabel)
    
    $endDatePicker = New-Object System.Windows.Controls.DatePicker
    $endDatePicker.Width = 120
    $dateFilterPanel.Children.Add($endDatePicker)
    
    $mainGrid.Children.Add($dateFilterPanel)
    
    # Abteilungsfilter
    $deptFilterPanel = New-Object System.Windows.Controls.StackPanel
    $deptFilterPanel.Orientation = "Horizontal"
    $deptFilterPanel.Margin = New-Object System.Windows.Thickness(0, 5, 0, 10)
    [System.Windows.Controls.Grid]::SetRow($deptFilterPanel, 3)
    
    $deptFilterLabel = New-Object System.Windows.Controls.Label
    $deptFilterLabel.Content = "Abteilung:"
    $deptFilterLabel.Width = 100
    $deptFilterLabel.VerticalContentAlignment = "Center"
    $deptFilterPanel.Children.Add($deptFilterLabel)
    
    $deptFilterCombo = New-Object System.Windows.Controls.ComboBox
    $deptFilterCombo.Margin = New-Object System.Windows.Thickness(0, 0, 10, 0)
    $deptFilterCombo.Height = 30
    $deptFilterCombo.MinWidth = 200
    $deptFilterCombo.HorizontalAlignment = "Stretch"
    
    # Abteilungsliste laden (hier könnte eine dynamische Ladung implementiert werden)
    $departments = @("Alle", "IT", "Personal", "Marketing", "Finanzen", "Vertrieb", "Entwicklung")
    foreach ($dept in $departments) {
        $deptFilterCombo.Items.Add($dept) | Out-Null
    }
    $deptFilterCombo.SelectedIndex = 0
    
    $deptFilterPanel.Children.Add($deptFilterCombo)
    
    $mainGrid.Children.Add($deptFilterPanel)
    
    # Status-Filter Checkboxen
    $statusPanel = New-Object System.Windows.Controls.GroupBox
    $statusPanel.Header = "Status"
    $statusPanel.Margin = New-Object System.Windows.Thickness(0, 10, 0, 10)
    [System.Windows.Controls.Grid]::SetRow($statusPanel, 4)
    
    $statusGrid = New-Object System.Windows.Controls.Grid
    $statusGrid.Margin = New-Object System.Windows.Thickness(10)
    
    # Status-Grid mit 3 Spalten aufteilen
    $statusGrid.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{Width = "*"}))
    $statusGrid.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{Width = "*"}))
    $statusGrid.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{Width = "*"}))
    
    # Zeilen definieren
    $statusGrid.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition -Property @{Height = "Auto"}))
    $statusGrid.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition -Property @{Height = "Auto"}))
    
    # Status-Checkboxen
    $statusCheckboxes = @{}
    
    # Alle Status-Checkboxen erstellen
    $statuses = @(
        @{ Name = "Neu"; Key = "New" },
        @{ Name = "Warte auf Manager"; Key = "PendingManagerInput" },
        @{ Name = "Warte auf HR Prüfung"; Key = "PendingHRVerification" },
        @{ Name = "Bereit für IT"; Key = "ReadyForIT" },
        @{ Name = "Abgeschlossen"; Key = "Completed" },
        @{ Name = "Alle"; Key = "All" }
    )
    
    for ($i = 0; $i -lt $statuses.Count; $i++) {
        $status = $statuses[$i]
        $row = [math]::Floor($i / 3)
        $col = $i % 3
        
        $checkBox = New-Object System.Windows.Controls.CheckBox
        $checkBox.Content = $status.Name
        $checkBox.Margin = New-Object System.Windows.Thickness(5, 5, 15, 5)
        $checkBox.Tag = $status.Key
        $checkBox.IsChecked = ($status.Key -eq "All")
        
        [System.Windows.Controls.Grid]::SetRow($checkBox, $row)
        [System.Windows.Controls.Grid]::SetColumn($checkBox, $col)
        $statusGrid.Children.Add($checkBox)
        
        $statusCheckboxes[$status.Key] = $checkBox
    }
    
    # Handler für "Alle" Checkbox
    $statusCheckboxes["All"].Add_Click({
        $isChecked = $statusCheckboxes["All"].IsChecked
        foreach ($key in $statusCheckboxes.Keys) {
            if ($key -ne "All") {
                $statusCheckboxes[$key].IsChecked = $isChecked
            }
        }
    })
    
    # Handler für einzelne Checkboxen
    foreach ($key in $statusCheckboxes.Keys) {
        if ($key -ne "All") {
            $statusCheckboxes[$key].Add_Click({
                # Prüfen, ob alle einzelnen Checkboxen markiert sind
                $allChecked = $true
                foreach ($k in $statusCheckboxes.Keys) {
                    if ($k -ne "All" -and -not $statusCheckboxes[$k].IsChecked) {
                        $allChecked = $false
                        break
                    }
                }
                $statusCheckboxes["All"].IsChecked = $allChecked
            })
        }
    }
    
    $statusPanel.Content = $statusGrid
    $mainGrid.Children.Add($statusPanel)
    
    # Button-Bereich
    $buttonPanel = New-Object System.Windows.Controls.StackPanel
    $buttonPanel.Orientation = "Horizontal"
    $buttonPanel.HorizontalAlignment = "Right"
    $buttonPanel.Margin = New-Object System.Windows.Thickness(0, 10, 0, 0)
    [System.Windows.Controls.Grid]::SetRow($buttonPanel, 5)
    
    $resetButton = New-Object System.Windows.Controls.Button
    $resetButton.Content = "Zurücksetzen"
    $resetButton.Width = 120
    $resetButton.Height = 34
    $resetButton.Margin = New-Object System.Windows.Thickness(0, 0, 10, 0)
    $resetButton.Background = "#CCCCCC"
    $resetButton.Foreground = "#333333"
    $buttonPanel.Children.Add($resetButton)
    
    $applyButton = New-Object System.Windows.Controls.Button
    $applyButton.Content = "Filter anwenden"
    $applyButton.Width = 150
    $applyButton.Height = 34
    $applyButton.Background = "#5B87A9"
    $applyButton.Foreground = "White"
    $applyButton.FontWeight = "SemiBold"
    $buttonPanel.Children.Add($applyButton)
    
    $mainGrid.Children.Add($buttonPanel)
    
    # Event-Handler für Reset-Button
    $resetButton.Add_Click({
        $quickSearchBox.Text = ""
        $startDatePicker.SelectedDate = $null
        $endDatePicker.SelectedDate = $null
        $deptFilterCombo.SelectedIndex = 0
        
        foreach ($key in $statusCheckboxes.Keys) {
            $statusCheckboxes[$key].IsChecked = ($key -eq "All")
        }
    })
    
    # Event-Handler für Apply-Button
    $applyButton.Add_Click({
        # Filter-Parameter sammeln
        $filterParams = @{
            SearchText = $quickSearchBox.Text
            Department = if ($deptFilterCombo.SelectedIndex -gt 0) { $deptFilterCombo.SelectedItem.ToString() } else { "" }
            FromDate = $startDatePicker.SelectedDate
            ToDate = $endDatePicker.SelectedDate
            StatusFilters = @()
        }
        
        # Status-Filter sammeln
        foreach ($key in $statusCheckboxes.Keys) {
            if ($key -ne "All" -and $statusCheckboxes[$key].IsChecked) {
                $filterParams.StatusFilters += $key
            }
        }
        
        # Wenn "All" oder keine Status ausgewählt sind, alle einschließen
        if ($statusCheckboxes["All"].IsChecked -or $filterParams.StatusFilters.Count -eq 0) {
            $filterParams.StatusFilters = $statuses | Where-Object { $_.Key -ne "All" } | ForEach-Object { $_.Key }
        }
        
        # OnSearch-Scriptblock mit Filter-Parametern aufrufen
        if ($OnSearch -ne $null) {
            & $OnSearch $filterParams
        }
        
        $searchWindow.DialogResult = $true
        $searchWindow.Close()
    })
    
    $searchWindow.Content = $mainGrid
    
    # Fenster anzeigen und Ergebnis zurückgeben
    return $searchWindow.ShowDialog()
}

function Update-FilteredRecords {
    param (
        [hashtable]$FilterParams,
        [System.Windows.Controls.ListBox]$TargetListBox
    )
    
    # Alle Datensätze laden
    $allRecords = Get-OnboardingRecords
    
    # Textfilter anwenden
    if (-not [string]::IsNullOrWhiteSpace($FilterParams.SearchText)) {
        $searchPattern = $FilterParams.SearchText.ToLower()
        $allRecords = $allRecords | Where-Object {
            $_.FirstName -like "*$searchPattern*" -or 
            $_.LastName -like "*$searchPattern*" -or
            $_.Description -like "*$searchPattern*" -or
            $_.Position -like "*$searchPattern*" -or
            $_.AssignedManager -like "*$searchPattern*" -or
            $_.OfficeRoom -like "*$searchPattern*"
        }
    }
    
    # Abteilungsfilter anwenden
    if (-not [string]::IsNullOrWhiteSpace($FilterParams.Department)) {
        $allRecords = $allRecords | Where-Object {
            $_.DepartmentField -eq $FilterParams.Department
        }
    }
    
    # Datumsfilter anwenden
    if ($FilterParams.FromDate -ne $null) {
        $allRecords = $allRecords | Where-Object {
            try {
                $recordDate = [DateTime]::Parse($_.CreatedDate)
                return $recordDate -ge $FilterParams.FromDate
            } catch {
                return $true # Bei Datum-Parsing-Fehler trotzdem einschließen
            }
        }
    }
    
    if ($FilterParams.ToDate -ne $null) {
        $allRecords = $allRecords | Where-Object {
            try {
                $recordDate = [DateTime]::Parse($_.CreatedDate)
                return $recordDate -le $FilterParams.ToDate
            } catch {
                return $true # Bei Datum-Parsing-Fehler trotzdem einschließen
            }
        }
    }
    
    # Status-Filter anwenden
    if ($FilterParams.StatusFilters -and $FilterParams.StatusFilters.Count -gt 0) {
        # Status-Werte in numerische Werte umwandeln
        $statusValues = @()
        foreach ($statusKey in $FilterParams.StatusFilters) {
            $statusValues += $workflowStates[$statusKey]
        }
        
        $allRecords = $allRecords | Where-Object {
            $statusValues -contains $_.WorkflowState
        }
    }
    
    # Ziel-Listbox aktualisieren
    if ($TargetListBox -ne $null) {
        $TargetListBox.Items.Clear()
        
        # Update-RecordsList mit gefilterten Daten aufrufen
        # oder alternativ direkt hier im Skript aktualisieren
        
        if ($allRecords.Count -eq 0) {
            $noRecordsItem = New-Object System.Windows.Controls.ListBoxItem
            $noRecordsItem.Content = "Keine Einträge gefunden"
            $noRecordsItem.IsEnabled = $false
            $TargetListBox.Items.Add($noRecordsItem)
        } else {
            foreach ($record in $allRecords) {
                # Item erstellen (hier könnte der Code aus Update-RecordsList wiederverwendet werden)
                $item = New-Object System.Windows.Controls.ListBoxItem
                $item.Content = "$($record.FirstName) $($record.LastName)"
                $item.Tag = $record
                $TargetListBox.Items.Add($item)
            }
        }
    }
    
    return $allRecords
}
