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
# SIG # Begin signature block
# MIIbywYJKoZIhvcNAQcCoIIbvDCCG7gCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCPK0M1m2Ge0GJs
# GPsIIXmv6xcYayGkyyCovFzCtgi4VqCCFhcwggMQMIIB+KADAgECAhB3jzsyX9Cg
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
# DQEJBDEiBCDWfpRQG4u8LrxeFajxrGK/Qwo2dH4Ytf2ft8CN6TaGQTANBgkqhkiG
# 9w0BAQEFAASCAQA1ATYrs38rUOZjwCCedQYtyBwWvdW3FJ6PhrL59S2lX5H5h0Ra
# z/fW5L7g3SWcke3YyBtqEvZjwvohDbwUBzMEtBljJSMXjgVmOpYDnK5G0Q348To4
# MEiFPRrtIEr+E/BRVwRxyXtOGhCCZc8A5DG3YIL/Sr1rdEJBmD/wxeWl5ZKZJkcw
# 51l7z+z8/i9Qn8Tz41QoDSdZMUYb5rxispGUDiaO5FqighjVaeSYu8xla1kv0d3x
# 7YAgB0AjKe2vkB6huvrYGzTEl+ELkqtrT3Qb4GXnlOTA/po/1BQNJv1UUgpS3mh2
# n15jmVD4LNZzCIXIIWJJd+B2WHUM5EcM5lrroYIDIDCCAxwGCSqGSIb3DQEJBjGC
# Aw0wggMJAgEBMHcwYzELMAkGA1UEBhMCVVMxFzAVBgNVBAoTDkRpZ2lDZXJ0LCBJ
# bmMuMTswOQYDVQQDEzJEaWdpQ2VydCBUcnVzdGVkIEc0IFJTQTQwOTYgU0hBMjU2
# IFRpbWVTdGFtcGluZyBDQQIQC65mvFq6f5WHxvnpBOMzBDANBglghkgBZQMEAgEF
# AKBpMBgGCSqGSIb3DQEJAzELBgkqhkiG9w0BBwEwHAYJKoZIhvcNAQkFMQ8XDTI1
# MDcwNTEwMTIwNFowLwYJKoZIhvcNAQkEMSIEINnDsOiI/W3HcffLIX52Fwm1RH5s
# 24OrlIeXPpZBc+cXMA0GCSqGSIb3DQEBAQUABIICAA1UXxrWQuMKb2f0l/zTC6TS
# gOO3GYDhY0Mel5Rv389lLTbmeHIf76RqY3UGtY6zj04E7jdUFDVgyiG/qaAvk2dp
# RFgR48ZgD+8tNXZIWAP2h7tzqoZDu6vaCgowz8eIBsChE10YfWCcoH+jTsxdqj5A
# 8256DAy+MBZ4iYWVkD8nFcd4PafX/bUIg6g4aKwlEELqGm3/qjDA8XY9V60CUcCr
# jF/BVP7YASsKA92YOxTO+inoNrd8NhrVv+w4YG2BZ2QpGJzzKB/PMdzUwlM2hG1p
# lsVDEdM8iwYzEpBEhHtecmNIoE2SOgXvgwM+2yq6Tx37O5dLOFYVf5jKsbV9oSvK
# a00Ego25TJS+OfVgtDQRexYNzcQHnwB180+kyRXhz6Vw2Sxv6GTGv8kRBzGQXr0r
# pG8IPCeo2fPsssH4qBtEyTKa4Yb8bqLHUdXIi5EKwAjOm8V4ir9IWC/j+8Vu4GsL
# 7OrHI6fADWxlG7H3P6FAQcWEMkja1lgAz+8XrOqGTI0X2esS/BLRqOTLyepXVtrK
# 1JAJlibZdqG5FYe5LWhzmo0TTDGAalOdySiGItiyAYFArELAt7nm20cTjNJz/q/u
# dYJJNb6uJHop5EN98A4Iq6Nk4/86+6vnrp4DmsvCrgyEBEsre4aaQGi4mtAF6/ub
# 8TXaX4KO0bGbPZm1GOH3
# SIG # End signature block
