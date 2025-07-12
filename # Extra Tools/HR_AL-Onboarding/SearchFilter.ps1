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

# SIG # Begin signature block
# MIIbywYJKoZIhvcNAQcCoIIbvDCCG7gCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBDWAqzyBmw8sCH
# KN5ZQYBM+D3NBS3kbSyaNOvYnGAxDKCCFhcwggMQMIIB+KADAgECAhB3jzsyX9Cg
# jEi+sBC2rBMTMA0GCSqGSIb3DQEBCwUAMCAxHjAcBgNVBAMMFVBoaW5JVC1QU3Nj
# cmlwdHNfU2lnbjAeFw0yNTA3MDUwODI4MTZaFw0yNzA3MDUwODM4MTZaMCAxHjAc
# BgNVBAMMFVBoaW5JVC1QU3NjcmlwdHNfU2lnbjCCASIwDQYJKoZIhvcNAQEBBQAD
# ggEPADCCAQoCggEBALmz3o//iDA5MvAndTjGX7/AvzTSACClfuUR9WYK0f6Ut2dI
# mPxn+Y9pZlLjXIpZT0H2Lvxq5aSI+aYeFtuJ8/0lULYNCVT31Bf+HxervRBKsUyi
# W9+4PH6STxo3Pl4l56UNQMcWLPNjDORWRPWHn0f99iNtjI+L4tUC/LoWSs3obzxN
# 3uTypzlaPBxis2qFSTR5SWqFdZdRkcuI5LNsJjyc/QWdTYRrfmVqp0QrvcxzCv8u
# EiVuni6jkXfiE6wz+oeI3L2iR+ywmU6CUX4tPWoS9VTtmm7AhEpasRTmrrnSg20Q
# jiBa1eH5TyLAH3TcYMxhfMbN9a2xDX5pzM65EJUCAwEAAaNGMEQwDgYDVR0PAQH/
# BAQDAgeAMBMGA1UdJQQMMAoGCCsGAQUFBwMDMB0GA1UdDgQWBBQO7XOqiE/EYi+n
# IaR6YO5M2MUuVTANBgkqhkiG9w0BAQsFAAOCAQEAjYOKIwBu1pfbdvEFFaR/uY88
# peKPk0NnvNEc3dpGdOv+Fsgbz27JPvItITFd6AKMoN1W48YjQLaU22M2jdhjGN5i
# FSobznP5KgQCDkRsuoDKiIOTiKAAknjhoBaCCEZGw8SZgKJtWzbST36Thsdd/won
# ihLsuoLxfcFnmBfrXh3rTIvTwvfujob68s0Sf5derHP/F+nphTymlg+y4VTEAijk
# g2dhy8RAsbS2JYZT7K5aEJpPXMiOLBqd7oTGfM7y5sLk2LIM4cT8hzgz3v5yPMkF
# H2MdR//K403e1EKH9MsGuGAJZddVN8ppaiESoPLoXrgnw2SY5KCmhYw1xRFdjTCC
# BY0wggR1oAMCAQICEA6bGI750C3n79tQ4ghAGFowDQYJKoZIhvcNAQEMBQAwZTEL
# MAkGA1UEBhMCVVMxFTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3
# LmRpZ2ljZXJ0LmNvbTEkMCIGA1UEAxMbRGlnaUNlcnQgQXNzdXJlZCBJRCBSb290
# IENBMB4XDTIyMDgwMTAwMDAwMFoXDTMxMTEwOTIzNTk1OVowYjELMAkGA1UEBhMC
# VVMxFTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0
# LmNvbTEhMB8GA1UEAxMYRGlnaUNlcnQgVHJ1c3RlZCBSb290IEc0MIICIjANBgkq
# hkiG9w0BAQEFAAOCAg8AMIICCgKCAgEAv+aQc2jeu+RdSjwwIjBpM+zCpyUuySE9
# 8orYWcLhKac9WKt2ms2uexuEDcQwH/MbpDgW61bGl20dq7J58soR0uRf1gU8Ug9S
# H8aeFaV+vp+pVxZZVXKvaJNwwrK6dZlqczKU0RBEEC7fgvMHhOZ0O21x4i0MG+4g
# 1ckgHWMpLc7sXk7Ik/ghYZs06wXGXuxbGrzryc/NrDRAX7F6Zu53yEioZldXn1RY
# jgwrt0+nMNlW7sp7XeOtyU9e5TXnMcvak17cjo+A2raRmECQecN4x7axxLVqGDgD
# EI3Y1DekLgV9iPWCPhCRcKtVgkEy19sEcypukQF8IUzUvK4bA3VdeGbZOjFEmjNA
# vwjXWkmkwuapoGfdpCe8oU85tRFYF/ckXEaPZPfBaYh2mHY9WV1CdoeJl2l6SPDg
# ohIbZpp0yt5LHucOY67m1O+SkjqePdwA5EUlibaaRBkrfsCUtNJhbesz2cXfSwQA
# zH0clcOP9yGyshG3u3/y1YxwLEFgqrFjGESVGnZifvaAsPvoZKYz0YkH4b235kOk
# GLimdwHhD5QMIR2yVCkliWzlDlJRR3S+Jqy2QXXeeqxfjT/JvNNBERJb5RBQ6zHF
# ynIWIgnffEx1P2PsIV/EIFFrb7GrhotPwtZFX50g/KEexcCPorF+CiaZ9eRpL5gd
# LfXZqbId5RsCAwEAAaOCATowggE2MA8GA1UdEwEB/wQFMAMBAf8wHQYDVR0OBBYE
# FOzX44LScV1kTN8uZz/nupiuHA9PMB8GA1UdIwQYMBaAFEXroq/0ksuCMS1Ri6en
# IZ3zbcgPMA4GA1UdDwEB/wQEAwIBhjB5BggrBgEFBQcBAQRtMGswJAYIKwYBBQUH
# MAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBDBggrBgEFBQcwAoY3aHR0cDov
# L2NhY2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEUm9vdENBLmNy
# dDBFBgNVHR8EPjA8MDqgOKA2hjRodHRwOi8vY3JsMy5kaWdpY2VydC5jb20vRGln
# aUNlcnRBc3N1cmVkSURSb290Q0EuY3JsMBEGA1UdIAQKMAgwBgYEVR0gADANBgkq
# hkiG9w0BAQwFAAOCAQEAcKC/Q1xV5zhfoKN0Gz22Ftf3v1cHvZqsoYcs7IVeqRq7
# IviHGmlUIu2kiHdtvRoU9BNKei8ttzjv9P+Aufih9/Jy3iS8UgPITtAq3votVs/5
# 9PesMHqai7Je1M/RQ0SbQyHrlnKhSLSZy51PpwYDE3cnRNTnf+hZqPC/Lwum6fI0
# POz3A8eHqNJMQBk1RmppVLC4oVaO7KTVPeix3P0c2PR3WlxUjG/voVA9/HYJaISf
# b8rbII01YBwCA8sgsKxYoA5AY8WYIsGyWfVVa88nq2x2zm8jLfR+cWojayL/ErhU
# LSd+2DrZ8LaHlv1b0VysGMNNn3O3AamfV6peKOK5lDCCBq4wggSWoAMCAQICEAc2
# N7ckVHzYR6z9KGYqXlswDQYJKoZIhvcNAQELBQAwYjELMAkGA1UEBhMCVVMxFTAT
# BgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNvbTEh
# MB8GA1UEAxMYRGlnaUNlcnQgVHJ1c3RlZCBSb290IEc0MB4XDTIyMDMyMzAwMDAw
# MFoXDTM3MDMyMjIzNTk1OVowYzELMAkGA1UEBhMCVVMxFzAVBgNVBAoTDkRpZ2lD
# ZXJ0LCBJbmMuMTswOQYDVQQDEzJEaWdpQ2VydCBUcnVzdGVkIEc0IFJTQTQwOTYg
# U0hBMjU2IFRpbWVTdGFtcGluZyBDQTCCAiIwDQYJKoZIhvcNAQEBBQADggIPADCC
# AgoCggIBAMaGNQZJs8E9cklRVcclA8TykTepl1Gh1tKD0Z5Mom2gsMyD+Vr2EaFE
# FUJfpIjzaPp985yJC3+dH54PMx9QEwsmc5Zt+FeoAn39Q7SE2hHxc7Gz7iuAhIoi
# GN/r2j3EF3+rGSs+QtxnjupRPfDWVtTnKC3r07G1decfBmWNlCnT2exp39mQh0YA
# e9tEQYncfGpXevA3eZ9drMvohGS0UvJ2R/dhgxndX7RUCyFobjchu0CsX7LeSn3O
# 9TkSZ+8OpWNs5KbFHc02DVzV5huowWR0QKfAcsW6Th+xtVhNef7Xj3OTrCw54qVI
# 1vCwMROpVymWJy71h6aPTnYVVSZwmCZ/oBpHIEPjQ2OAe3VuJyWQmDo4EbP29p7m
# O1vsgd4iFNmCKseSv6De4z6ic/rnH1pslPJSlRErWHRAKKtzQ87fSqEcazjFKfPK
# qpZzQmiftkaznTqj1QPgv/CiPMpC3BhIfxQ0z9JMq++bPf4OuGQq+nUoJEHtQr8F
# nGZJUlD0UfM2SU2LINIsVzV5K6jzRWC8I41Y99xh3pP+OcD5sjClTNfpmEpYPtMD
# iP6zj9NeS3YSUZPJjAw7W4oiqMEmCPkUEBIDfV8ju2TjY+Cm4T72wnSyPx4Jduyr
# XUZ14mCjWAkBKAAOhFTuzuldyF4wEr1GnrXTdrnSDmuZDNIztM2xAgMBAAGjggFd
# MIIBWTASBgNVHRMBAf8ECDAGAQH/AgEAMB0GA1UdDgQWBBS6FtltTYUvcyl2mi91
# jGogj57IbzAfBgNVHSMEGDAWgBTs1+OC0nFdZEzfLmc/57qYrhwPTzAOBgNVHQ8B
# Af8EBAMCAYYwEwYDVR0lBAwwCgYIKwYBBQUHAwgwdwYIKwYBBQUHAQEEazBpMCQG
# CCsGAQUFBzABhhhodHRwOi8vb2NzcC5kaWdpY2VydC5jb20wQQYIKwYBBQUHMAKG
# NWh0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRSb290
# RzQuY3J0MEMGA1UdHwQ8MDowOKA2oDSGMmh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNv
# bS9EaWdpQ2VydFRydXN0ZWRSb290RzQuY3JsMCAGA1UdIAQZMBcwCAYGZ4EMAQQC
# MAsGCWCGSAGG/WwHATANBgkqhkiG9w0BAQsFAAOCAgEAfVmOwJO2b5ipRCIBfmbW
# 2CFC4bAYLhBNE88wU86/GPvHUF3iSyn7cIoNqilp/GnBzx0H6T5gyNgL5Vxb122H
# +oQgJTQxZ822EpZvxFBMYh0MCIKoFr2pVs8Vc40BIiXOlWk/R3f7cnQU1/+rT4os
# equFzUNf7WC2qk+RZp4snuCKrOX9jLxkJodskr2dfNBwCnzvqLx1T7pa96kQsl3p
# /yhUifDVinF2ZdrM8HKjI/rAJ4JErpknG6skHibBt94q6/aesXmZgaNWhqsKRcnf
# xI2g55j7+6adcq/Ex8HBanHZxhOACcS2n82HhyS7T6NJuXdmkfFynOlLAlKnN36T
# U6w7HQhJD5TNOXrd/yVjmScsPT9rp/Fmw0HNT7ZAmyEhQNC3EyTN3B14OuSereU0
# cZLXJmvkOHOrpgFPvT87eK1MrfvElXvtCl8zOYdBeHo46Zzh3SP9HSjTx/no8Zhf
# +yvYfvJGnXUsHicsJttvFXseGYs2uJPU5vIXmVnKcPA3v5gA3yAWTyf7YGcWoWa6
# 3VXAOimGsJigK+2VQbc61RWYMbRiCQ8KvYHZE/6/pNHzV9m8BPqC3jLfBInwAM1d
# wvnQI38AC+R2AibZ8GV2QqYphwlHK+Z/GqSFD/yYlvZVVCsfgPrA8g4r5db7qS9E
# FUrnEw4d2zc4GqEr9u3WfPwwgga8MIIEpKADAgECAhALrma8Wrp/lYfG+ekE4zME
# MA0GCSqGSIb3DQEBCwUAMGMxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2Vy
# dCwgSW5jLjE7MDkGA1UEAxMyRGlnaUNlcnQgVHJ1c3RlZCBHNCBSU0E0MDk2IFNI
# QTI1NiBUaW1lU3RhbXBpbmcgQ0EwHhcNMjQwOTI2MDAwMDAwWhcNMzUxMTI1MjM1
# OTU5WjBCMQswCQYDVQQGEwJVUzERMA8GA1UEChMIRGlnaUNlcnQxIDAeBgNVBAMT
# F0RpZ2lDZXJ0IFRpbWVzdGFtcCAyMDI0MIICIjANBgkqhkiG9w0BAQEFAAOCAg8A
# MIICCgKCAgEAvmpzn/aVIauWMLpbbeZZo7Xo/ZEfGMSIO2qZ46XB/QowIEMSvgjE
# dEZ3v4vrrTHleW1JWGErrjOL0J4L0HqVR1czSzvUQ5xF7z4IQmn7dHY7yijvoQ7u
# jm0u6yXF2v1CrzZopykD07/9fpAT4BxpT9vJoJqAsP8YuhRvflJ9YeHjes4fduks
# THulntq9WelRWY++TFPxzZrbILRYynyEy7rS1lHQKFpXvo2GePfsMRhNf1F41nyE
# g5h7iOXv+vjX0K8RhUisfqw3TTLHj1uhS66YX2LZPxS4oaf33rp9HlfqSBePejlY
# eEdU740GKQM7SaVSH3TbBL8R6HwX9QVpGnXPlKdE4fBIn5BBFnV+KwPxRNUNK6lY
# k2y1WSKour4hJN0SMkoaNV8hyyADiX1xuTxKaXN12HgR+8WulU2d6zhzXomJ2Ple
# I9V2yfmfXSPGYanGgxzqI+ShoOGLomMd3mJt92nm7Mheng/TBeSA2z4I78JpwGpT
# RHiT7yHqBiV2ngUIyCtd0pZ8zg3S7bk4QC4RrcnKJ3FbjyPAGogmoiZ33c1HG93V
# p6lJ415ERcC7bFQMRbxqrMVANiav1k425zYyFMyLNyE1QulQSgDpW9rtvVcIH7Wv
# G9sqYup9j8z9J1XqbBZPJ5XLln8mS8wWmdDLnBHXgYly/p1DhoQo5fkCAwEAAaOC
# AYswggGHMA4GA1UdDwEB/wQEAwIHgDAMBgNVHRMBAf8EAjAAMBYGA1UdJQEB/wQM
# MAoGCCsGAQUFBwMIMCAGA1UdIAQZMBcwCAYGZ4EMAQQCMAsGCWCGSAGG/WwHATAf
# BgNVHSMEGDAWgBS6FtltTYUvcyl2mi91jGogj57IbzAdBgNVHQ4EFgQUn1csA3cO
# KBWQZqVjXu5Pkh92oFswWgYDVR0fBFMwUTBPoE2gS4ZJaHR0cDovL2NybDMuZGln
# aWNlcnQuY29tL0RpZ2lDZXJ0VHJ1c3RlZEc0UlNBNDA5NlNIQTI1NlRpbWVTdGFt
# cGluZ0NBLmNybDCBkAYIKwYBBQUHAQEEgYMwgYAwJAYIKwYBBQUHMAGGGGh0dHA6
# Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBYBggrBgEFBQcwAoZMaHR0cDovL2NhY2VydHMu
# ZGlnaWNlcnQuY29tL0RpZ2lDZXJ0VHJ1c3RlZEc0UlNBNDA5NlNIQTI1NlRpbWVT
# dGFtcGluZ0NBLmNydDANBgkqhkiG9w0BAQsFAAOCAgEAPa0eH3aZW+M4hBJH2UOR
# 9hHbm04IHdEoT8/T3HuBSyZeq3jSi5GXeWP7xCKhVireKCnCs+8GZl2uVYFvQe+p
# PTScVJeCZSsMo1JCoZN2mMew/L4tpqVNbSpWO9QGFwfMEy60HofN6V51sMLMXNTL
# fhVqs+e8haupWiArSozyAmGH/6oMQAh078qRh6wvJNU6gnh5OruCP1QUAvVSu4kq
# VOcJVozZR5RRb/zPd++PGE3qF1P3xWvYViUJLsxtvge/mzA75oBfFZSbdakHJe2B
# VDGIGVNVjOp8sNt70+kEoMF+T6tptMUNlehSR7vM+C13v9+9ZOUKzfRUAYSyyEmY
# tsnpltD/GWX8eM70ls1V6QG/ZOB6b6Yum1HvIiulqJ1Elesj5TMHq8CWT/xrW7tw
# ipXTJ5/i5pkU5E16RSBAdOp12aw8IQhhA/vEbFkEiF2abhuFixUDobZaA0VhqAsM
# HOmaT3XThZDNi5U2zHKhUs5uHHdG6BoQau75KiNbh0c+hatSF+02kULkftARjsyE
# pHKsF7u5zKRbt5oK5YGwFvgc4pEVUNytmB3BpIiowOIIuDgP5M9WArHYSAR16gc0
# dP2XdkMEP5eBsX7bf/MGN4K3HP50v/01ZHo/Z5lGLvNwQ7XHBx1yomzLP8lx4Q1z
# ZKDyHcp4VQJLu2kWTsKsOqQxggUKMIIFBgIBATA0MCAxHjAcBgNVBAMMFVBoaW5J
# VC1QU3NjcmlwdHNfU2lnbgIQd487Ml/QoIxIvrAQtqwTEzANBglghkgBZQMEAgEF
# AKCBhDAYBgorBgEEAYI3AgEMMQowCKACgAChAoAAMBkGCSqGSIb3DQEJAzEMBgor
# BgEEAYI3AgEEMBwGCisGAQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3
# DQEJBDEiBCAwrJua3EUxGyTUgRRfqiYG8ubjeWu4N722jz+m2ypIJDANBgkqhkiG
# 9w0BAQEFAASCAQA7UDE+DnUy8DUzOR6Mr+BEXyj0XEkoPgxLEM5mKSLM34dn/FyI
# U3X05zXn4Ij0rFHFkouD7QB3muoaucyLxqHGXM2ffR3TwtTicxRLlAapjUrHkuzs
# LHW6GAy6yiBVByxIMiq/t0yiJqtpVLA+obPZbp84mopCDmlBeeJ7FEo+6F71llYN
# 1koobyXjWDTQ7V0REuHmgy9I+UEy4ri6bW4f0Lfw8OXc6ZSk0MYYT9C3Bjo06bZL
# z/ZHvQqPmUeiSpQwDIcnrIcShjur9ao2q1JZ3apvFO/xt1OMWw8mrp6EJdiD1wAn
# r9gf3LT4+LdhCx92RzVKjoJzZBYHmwMVKA8poYIDIDCCAxwGCSqGSIb3DQEJBjGC
# Aw0wggMJAgEBMHcwYzELMAkGA1UEBhMCVVMxFzAVBgNVBAoTDkRpZ2lDZXJ0LCBJ
# bmMuMTswOQYDVQQDEzJEaWdpQ2VydCBUcnVzdGVkIEc0IFJTQTQwOTYgU0hBMjU2
# IFRpbWVTdGFtcGluZyBDQQIQC65mvFq6f5WHxvnpBOMzBDANBglghkgBZQMEAgEF
# AKBpMBgGCSqGSIb3DQEJAzELBgkqhkiG9w0BBwEwHAYJKoZIhvcNAQkFMQ8XDTI1
# MDcwNTEwMTIwNFowLwYJKoZIhvcNAQkEMSIEIGYp5QH1Y348QgnsMHDBtNOdmKrJ
# QGZH05Jc+peXK20AMA0GCSqGSIb3DQEBAQUABIICAEZCiUcL4uOBUSqAOm2/8kzZ
# afmw/Th51VPr0qul+OJVFJIYFrl5+PCLSnFEecBisSEwvvRpIDiBf9DawaZ6TBUo
# yaW35CvpjXw7HaO2i0YyGjgxTW6pkoZbO+4RGL0aL23Dal8qMSpOt+aKDGYmm0qi
# B+7XaJ9AoWcvWUt0/cByruJ5s0XsE8rr/5pdXYHcxzN9zF66FAtUbM3+Aba47sne
# jbTzLq9c1hin8VW8tPT/Y/hTsE2R8drpqnc+VssH9Z0OwWevqqYVepXgB5rNIY4k
# C+PvQSZ0xvMIPGuFgZeJv4B0i4WeF27xXur/LAeBiukRWl6nVwUbf5Iv0NWBra++
# oBTGUdgeaYWtvrg9nYkCTFTq6BoETJenyw2WmXJbDqIYGJwWk00u4L1+XN5NHnTl
# vy+mNHYzPZCa1gQXf95o9gDeU/5lPMwiK+DMzgxS0A5bUc5941+KODUpIL0FtgHy
# 2ci65wzFtyTJasWNbptKrBstGYQ1D9l7sRKkri76zav2UIulxk0DVwVSDPyXmQ17
# JEsnYRbOhWCe8ntRENNOAXcVFTZvFt7TYhwp6/YmHQLT7v95tg263AatR/s8jPuX
# oLIhYcdpq38KRdr1BKctX5jnX9spsLqMByePTgnqbKjlXi5CdWY97xC/JEgVoyz0
# wHXjHqEWgFg9alnPf+fL
# SIG # End signature block
