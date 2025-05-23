function Show-ReportingDashboard {
    param (
        [string]$ReportType = "General"
    )
    
    # Hauptfenster für Reporting-Dashboard erstellen
    $dashboardWindow = New-Object System.Windows.Window
    $dashboardWindow.Title = "easyONBOARDING - Berichtswesen"
    $dashboardWindow.Width = 900
    $dashboardWindow.Height = 700
    $dashboardWindow.WindowStartupLocation = "CenterScreen"
    $dashboardWindow.Background = "#F0F0F0"
    
    # Haupt-Grid erstellen
    $mainGrid = New-Object System.Windows.Controls.Grid
    $mainGrid.Margin = New-Object System.Windows.Thickness(15)
    
    # Zeilen definieren
    $mainGrid.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition -Property @{Height = "Auto"}))
    $mainGrid.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition -Property @{Height = "Auto"}))
    $mainGrid.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition -Property @{Height = "*"}))
    $mainGrid.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition -Property @{Height = "Auto"}))
    
    # Titel
    $title = New-Object System.Windows.Controls.Label
    $title.Content = "easyONBOARDING - Berichtswesen"
    $title.FontSize = 20
    $title.FontWeight = "Bold"
    $title.Foreground = "#5B87A9"
    $title.Margin = New-Object System.Windows.Thickness(0, 0, 0, 15)
    [System.Windows.Controls.Grid]::SetRow($title, 0)
    $mainGrid.Children.Add($title)
    
    # Filter-Bereich
    $filterPanel = New-Object System.Windows.Controls.GroupBox
    $filterPanel.Header = "Filter"
    $filterPanel.Margin = New-Object System.Windows.Thickness(0, 0, 0, 15)
    [System.Windows.Controls.Grid]::SetRow($filterPanel, 1)
    
    $filterGrid = New-Object System.Windows.Controls.Grid
    $filterGrid.Margin = New-Object System.Windows.Thickness(10)
    
    # Filter-Grid aufteilen
    $filterGrid.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{Width = "Auto"}))
    $filterGrid.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{Width = "*"}))
    $filterGrid.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{Width = "Auto"}))
    $filterGrid.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{Width = "*"}))
    
    $filterGrid.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition -Property @{Height = "Auto"}))
    $filterGrid.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition -Property @{Height = "Auto"}))
    
    # Berichtstyp-Label und Dropdown
    $reportTypeLabel = New-Object System.Windows.Controls.Label
    $reportTypeLabel.Content = "Berichtstyp:"
    $reportTypeLabel.VerticalAlignment = "Center"
    [System.Windows.Controls.Grid]::SetRow($reportTypeLabel, 0)
    [System.Windows.Controls.Grid]::SetColumn($reportTypeLabel, 0)
    $filterGrid.Children.Add($reportTypeLabel)
    
    $reportTypeCombo = New-Object System.Windows.Controls.ComboBox
    $reportTypeCombo.Margin = New-Object System.Windows.Thickness(5, 5, 15, 5)
    $reportTypeCombo.Height = 30
    $reportTypeCombo.MinWidth = 200
    [System.Windows.Controls.Grid]::SetRow($reportTypeCombo, 0)
    [System.Windows.Controls.Grid]::SetColumn($reportTypeCombo, 1)
    
    # Berichtstypen hinzufügen
    $reportTypes = @(
        @{ Name = "Allgemeine Übersicht"; Key = "General" },
        @{ Name = "Onboarding nach Abteilung"; Key = "ByDepartment" },
        @{ Name = "Onboarding nach Zeitraum"; Key = "ByTimeframe" },
        @{ Name = "Status-Übersicht"; Key = "StatusOverview" },
        @{ Name = "Durchlaufzeiten"; Key = "ProcessingTime" }
    )
    
    foreach ($type in $reportTypes) {
        $item = New-Object System.Windows.Controls.ComboBoxItem
        $item.Content = $type.Name
        $item.Tag = $type.Key
        $reportTypeCombo.Items.Add($item)
        
        if ($type.Key -eq $ReportType) {
            $reportTypeCombo.SelectedItem = $item
        }
    }
    
    if ($reportTypeCombo.SelectedItem -eq $null -and $reportTypeCombo.Items.Count -gt 0) {
        $reportTypeCombo.SelectedIndex = 0
    }
    
    $filterGrid.Children.Add($reportTypeCombo)
    
    # Zeitraum-Label und DatePicker
    $timeframeLabel = New-Object System.Windows.Controls.Label
    $timeframeLabel.Content = "Zeitraum:"
    $timeframeLabel.VerticalAlignment = "Center"
    [System.Windows.Controls.Grid]::SetRow($timeframeLabel, 0)
    [System.Windows.Controls.Grid]::SetColumn($timeframeLabel, 2)
    $filterGrid.Children.Add($timeframeLabel)
    
    $timeframePanel = New-Object System.Windows.Controls.StackPanel
    $timeframePanel.Orientation = "Horizontal"
    $timeframePanel.Margin = New-Object System.Windows.Thickness(5)
    [System.Windows.Controls.Grid]::SetRow($timeframePanel, 0)
    [System.Windows.Controls.Grid]::SetColumn($timeframePanel, 3)
    
    $fromDatePicker = New-Object System.Windows.Controls.DatePicker
    $fromDatePicker.Width = 120
    $fromDatePicker.Margin = New-Object System.Windows.Thickness(0, 0, 5, 0)
    $timeframePanel.Children.Add($fromDatePicker)
    
    $toLabel = New-Object System.Windows.Controls.Label
    $toLabel.Content = "bis"
    $toLabel.VerticalAlignment = "Center"
    $toLabel.Margin = New-Object System.Windows.Thickness(0, 0, 5, 0)
    $timeframePanel.Children.Add($toLabel)
    
    $toDatePicker = New-Object System.Windows.Controls.DatePicker
    $toDatePicker.Width = 120
    $timeframePanel.Children.Add($toDatePicker)
    
    $filterGrid.Children.Add($timeframePanel)
    
    # Zusätzliche Filter-Optionen (zweite Zeile)
    $additionalFilterLabel = New-Object System.Windows.Controls.Label
    $additionalFilterLabel.Content = "Abteilung:"
    $additionalFilterLabel.VerticalAlignment = "Center"
    [System.Windows.Controls.Grid]::SetRow($additionalFilterLabel, 1)
    [System.Windows.Controls.Grid]::SetColumn($additionalFilterLabel, 0)
    $filterGrid.Children.Add($additionalFilterLabel)
    
    $departmentCombo = New-Object System.Windows.Controls.ComboBox
    $departmentCombo.Margin = New-Object System.Windows.Thickness(5)
    $departmentCombo.Height = 30
    $departmentCombo.MinWidth = 200
    [System.Windows.Controls.Grid]::SetRow($departmentCombo, 1)
    [System.Windows.Controls.Grid]::SetColumn($department