# Parameterblock – alle erforderlichen Parameter werden hier deklariert.
param(
    [string]$htmlFile = "",
    [string]$pdfFile = "",
    [string]$wkhtmltopdfPath = "",
    # URL zum PowerShell 7 Setup-Paket (anpassen, falls benötigt)
    [string]$PS7SetupURL = "https://github.com/PowerShell/PowerShell/releases/download/v7.5.0/PowerShell-7.5.0-win-x64.exe"
)

# Falls alle Parameter für die PDF-Konvertierung vorhanden sind, wird diese direkt ausgeführt.
if ($htmlFile -and $pdfFile -and $wkhtmltopdfPath) {
    $htmlFileURL = "file:///" + ($htmlFile -replace '\\', '/')
    $arguments = @("--enable-local-file-access", $htmlFileURL, $pdfFile)
    [System.Windows.Forms.MessageBox]::Show("Starte wkhtmltopdf mit folgenden Argumenten:`n$($arguments -join ' ')", "wkhtmltopdf")
    Start-Process -FilePath $wkhtmltopdfPath -ArgumentList $arguments -Wait -NoNewWindow
    if (Test-Path $pdfFile) {
        [System.Windows.Forms.MessageBox]::Show("PDF erfolgreich erstellt: $pdfFile", "Erfolg")
    }
    else {
        [System.Windows.Forms.MessageBox]::Show("PDF konnte nicht erstellt werden.", "Fehler", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
    return
}

# -----------------------------------
# Funktion: Anzeige des Installations-Assistenten (GUI)
# -----------------------------------
function Show-InstallerForm {
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Installations-Assistent"
    $form.Size = New-Object System.Drawing.Size(500,400)
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $form.MaximizeBox = $false

    # Status-Label (alle Meldungen werden hier angezeigt)
    $lblStatus = New-Object System.Windows.Forms.Label
    $lblStatus.Location = New-Object System.Drawing.Point(20,20)
    $lblStatus.Size = New-Object System.Drawing.Size(440,60)
    $lblStatus.Text = "Bereit..."
    $lblStatus.Font = New-Object System.Drawing.Font("Segoe UI",10)
    $form.Controls.Add($lblStatus)

    # ProgressBar
    $progressBar = New-Object System.Windows.Forms.ProgressBar
    $progressBar.Location = New-Object System.Drawing.Point(20,90)
    $progressBar.Size = New-Object System.Drawing.Size(440,25)
    $progressBar.Minimum = 0
    $progressBar.Maximum = 100
    $progressBar.Style = "Continuous"
    $form.Controls.Add($progressBar)

    # Checkbox: Ausführungsrichtlinie automatisch auf RemoteSigned setzen
    $chkSetPolicy = New-Object System.Windows.Forms.CheckBox
    $chkSetPolicy.Location = New-Object System.Drawing.Point(20,130)
    $chkSetPolicy.Size = New-Object System.Drawing.Size(300,20)
    $chkSetPolicy.Text = "Automatisch auf RemoteSigned setzen"
    $chkSetPolicy.Checked = $true
    $form.Controls.Add($chkSetPolicy)

    # Checkbox: PowerShell 7 installieren (via lokales Setup-Paket)
    $chkInstallPS7 = New-Object System.Windows.Forms.CheckBox
    $chkInstallPS7.Location = New-Object System.Drawing.Point(20,160)
    $chkInstallPS7.Size = New-Object System.Drawing.Size(300,20)
    $chkInstallPS7.Text = "PowerShell 7 installieren (Setup.exe)"
    $chkInstallPS7.Checked = $true
    $form.Controls.Add($chkInstallPS7)

    # Checkbox: ActiveDirectory-Modul installieren
    $chkInstallADModule = New-Object System.Windows.Forms.CheckBox
    $chkInstallADModule.Location = New-Object System.Drawing.Point(20,190)
    $chkInstallADModule.Size = New-Object System.Drawing.Size(300,20)
    $chkInstallADModule.Text = "ActiveDirectory-Modul installieren"
    $chkInstallADModule.Checked = $true
    $form.Controls.Add($chkInstallADModule)

    # Checkbox: PSWritePDF-Modul installieren
    $chkInstallPSWritePDF = New-Object System.Windows.Forms.CheckBox
    $chkInstallPSWritePDF.Location = New-Object System.Drawing.Point(20,220)
    $chkInstallPSWritePDF.Size = New-Object System.Drawing.Size(300,20)
    $chkInstallPSWritePDF.Text = "PSWritePDF-Modul installieren"
    $chkInstallPSWritePDF.Checked = $true
    $form.Controls.Add($chkInstallPSWritePDF)

    # Button "Install" (hellgrün)
    $btnInstall = New-Object System.Windows.Forms.Button
    $btnInstall.Text = "Install"
    $btnInstall.Size = New-Object System.Drawing.Size(100,30)
    $btnInstall.Location = New-Object System.Drawing.Point(100,300)
    $btnInstall.BackColor = [System.Drawing.Color]::LightGreen
    $form.Controls.Add($btnInstall)

    # Button "Close" (hellrot)
    $btnClose = New-Object System.Windows.Forms.Button
    $btnClose.Text = "Close"
    $btnClose.Size = New-Object System.Drawing.Size(100,30)
    $btnClose.Location = New-Object System.Drawing.Point(300,300)
    $btnClose.BackColor = [System.Drawing.Color]::LightCoral
    $form.Controls.Add($btnClose)

    return @{
        Form = $form;
        Label = $lblStatus;
        ProgressBar = $progressBar;
        SetPolicyCheckBox = $chkSetPolicy;
        InstallPS7CheckBox = $chkInstallPS7;
        InstallADModuleCheckBox = $chkInstallADModule;
        InstallPSWritePDFCheckBox = $chkInstallPSWritePDF;
        InstallButton = $btnInstall;
        CloseButton = $btnClose
    }
}

# -----------------------------------
# Funktion: Installation ausführen (alle Schritte)
# -----------------------------------
function Run-Installation {
    param(
        $guiElements
    )
    $lbl = $guiElements.Label
    $pb = $guiElements.ProgressBar
    $chkSetPolicy = $guiElements.SetPolicyCheckBox
    $chkInstallPS7 = $guiElements.InstallPS7CheckBox
    $chkInstallAD = $guiElements.InstallADModuleCheckBox
    $chkInstallPSWritePDF = $guiElements.InstallPSWritePDFCheckBox

    try {
        # Schritt 1: Ausführungsrichtlinie prüfen
        $lbl.Text = "Prüfe Ausführungsrichtlinie..."
        [System.Windows.Forms.Application]::DoEvents()
        $currentPolicy = Get-ExecutionPolicy
        if ($currentPolicy -eq "Restricted" -or $currentPolicy -eq "AllSigned") {
            if ($chkSetPolicy.Checked) {
                Set-ExecutionPolicy RemoteSigned -Scope Process -Force
                $lbl.Text = "Ausführungsrichtlinie auf RemoteSigned gesetzt."
            }
            else {
                $lbl.Text = "Ausführungsrichtlinie ($currentPolicy) ist zu restriktiv."
            }
        }
        else {
            $lbl.Text = "Ausführungsrichtlinie ($currentPolicy) ist akzeptabel."
        }
        Start-Sleep -Seconds 1
        $pb.Value = 20
        [System.Windows.Forms.Application]::DoEvents()

        # Schritt 2: msstore-Quelle entfernen, um manuelle Bestätigungen zu vermeiden
        try {
            winget source remove msstore -ErrorAction Stop | Out-Null
            $lbl.Text = "msstore-Quelle entfernt."
        }
        catch {
            $lbl.Text = "msstore-Quelle konnte nicht entfernt werden oder ist bereits entfernt."
        }
        Start-Sleep -Seconds 1
        $pb.Value = 30
        [System.Windows.Forms.Application]::DoEvents()

        # Schritt 3: PowerShell 7 prüfen und ggf. installieren via lokales Setup-Paket
        if ($chkInstallPS7.Checked) {
            $lbl.Text = "Prüfe PowerShell 7..."
            [System.Windows.Forms.Application]::DoEvents()
            $ps7 = Get-Command pwsh.exe -ErrorAction SilentlyContinue
            if ($ps7) {
                $lbl.Text = "PowerShell 7 ist bereits installiert."
            }
            else {
                $lbl.Text = "PowerShell 7 nicht gefunden."
                [System.Windows.Forms.Application]::DoEvents()
                $setupPath = Join-Path $PSScriptRoot "PowerShell7Setup.exe"
                # Wenn das Setup-Paket nicht vorhanden ist, wird kein Download gestartet.
                if (-not (Test-Path $setupPath)) {
                    $lbl.Text = "Setup-Paket für PowerShell 7 nicht vorhanden."
                }
                else {
                    $lbl.Text = "Starte PowerShell 7 Setup (silent)..."
                    [System.Windows.Forms.Application]::DoEvents()
                    Start-Process -FilePath $setupPath -ArgumentList "/quiet", "/norestart" -Wait -NoNewWindow
                    # Nach der Installation erneut prüfen
                    $ps7 = Get-Command pwsh.exe -ErrorAction SilentlyContinue
                    if ($ps7) {
                        $lbl.Text = "PowerShell 7 erfolgreich installiert."
                    }
                    else {
                        $lbl.Text = "Installation von PowerShell 7 via Setup-Paket fehlgeschlagen."
                    }
                }
            }
            Start-Sleep -Seconds 1
            $pb.Value = 50
            [System.Windows.Forms.Application]::DoEvents()
        }

        # Schritt 4: ActiveDirectory-Modul prüfen und installieren
        if ($chkInstallAD.Checked) {
            $lbl.Text = "Prüfe ActiveDirectory-Modul..."
            [System.Windows.Forms.Application]::DoEvents()
            if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
                $lbl.Text = "ActiveDirectory-Modul nicht gefunden. Versuche Installation..."
                [System.Windows.Forms.Application]::DoEvents()
                try {
                    Add-WindowsCapability -Online -Name "Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0" -ErrorAction Stop
                    $lbl.Text = "ActiveDirectory-Modul erfolgreich installiert."
                }
                catch {
                    $lbl.Text = "RSAT-Installation fehlgeschlagen. Bitte manuell installieren."
                }
            }
            else {
                $lbl.Text = "ActiveDirectory-Modul ist vorhanden."
            }
            Start-Sleep -Seconds 1
            $pb.Value = 70
            [System.Windows.Forms.Application]::DoEvents()
        }

        # Schritt 5: PSWritePDF-Modul prüfen, installieren und importieren
        if ($chkInstallPSWritePDF.Checked) {
            $lbl.Text = "Prüfe PSWritePDF-Modul..."
            [System.Windows.Forms.Application]::DoEvents()
            if (-not (Get-Module -ListAvailable -Name PSWritePDF)) {
                $lbl.Text = "PSWritePDF-Modul nicht gefunden. Starte Installation..."
                [System.Windows.Forms.Application]::DoEvents()
                try {
                    Install-Module PSWritePDF -Scope CurrentUser -Force -AllowClobber
                    $lbl.Text = "PSWritePDF-Modul erfolgreich installiert."
                }
                catch {
                    $lbl.Text = "Fehler bei der Installation von PSWritePDF: $($_.Exception.Message)"
                }
            }
            else {
                $lbl.Text = "PSWritePDF-Modul ist bereits installiert."
            }
            Start-Sleep -Seconds 1
            $pb.Value = 80
            [System.Windows.Forms.Application]::DoEvents()

            try {
                Import-Module PSWritePDF -ErrorAction Stop
                $lbl.Text = "PSWritePDF-Modul erfolgreich importiert."
            }
            catch {
                $lbl.Text = "Fehler beim Import von PSWritePDF."
            }
            Start-Sleep -Seconds 1
            $pb.Value = 90
            [System.Windows.Forms.Application]::DoEvents()
        }

        # Abschluss
        $lbl.Text = "Installation abgeschlossen. Alle ausgewählten Module und Tools wurden geprüft."
        $pb.Value = 100
        [System.Windows.Forms.Application]::DoEvents()
        Start-Sleep -Seconds 2
    }
    catch {
        $lbl.Text = "Fehler: $($_.Exception.Message)"
    }
}

# -----------------------------------
# Hauptteil: GUI anzeigen und Button-Ereignisse verarbeiten
# -----------------------------------
$gui = Show-InstallerForm
$gui.InstallButton.Add_Click({ Run-Installation -guiElements $gui })
$gui.CloseButton.Add_Click({ $gui.Form.Close() })

[System.Windows.Forms.Application]::Run($gui.Form)

# SIG # Begin signature block
# MIIcCAYJKoZIhvcNAQcCoIIb+TCCG/UCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCD0jQP/qZKzLraR
# g0qgZTiEoi8FRFWpknFyjjpO4LwWE6CCFk4wggMQMIIB+KADAgECAhB3jzsyX9Cg
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
# LSd+2DrZ8LaHlv1b0VysGMNNn3O3AamfV6peKOK5lDCCBrQwggScoAMCAQICEA3H
# rFcF/yGZLkBDIgw6SYYwDQYJKoZIhvcNAQELBQAwYjELMAkGA1UEBhMCVVMxFTAT
# BgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNvbTEh
# MB8GA1UEAxMYRGlnaUNlcnQgVHJ1c3RlZCBSb290IEc0MB4XDTI1MDUwNzAwMDAw
# MFoXDTM4MDExNDIzNTk1OVowaTELMAkGA1UEBhMCVVMxFzAVBgNVBAoTDkRpZ2lD
# ZXJ0LCBJbmMuMUEwPwYDVQQDEzhEaWdpQ2VydCBUcnVzdGVkIEc0IFRpbWVTdGFt
# cGluZyBSU0E0MDk2IFNIQTI1NiAyMDI1IENBMTCCAiIwDQYJKoZIhvcNAQEBBQAD
# ggIPADCCAgoCggIBALR4MdMKmEFyvjxGwBysddujRmh0tFEXnU2tjQ2UtZmWgyxU
# 7UNqEY81FzJsQqr5G7A6c+Gh/qm8Xi4aPCOo2N8S9SLrC6Kbltqn7SWCWgzbNfiR
# +2fkHUiljNOqnIVD/gG3SYDEAd4dg2dDGpeZGKe+42DFUF0mR/vtLa4+gKPsYfwE
# u7EEbkC9+0F2w4QJLVSTEG8yAR2CQWIM1iI5PHg62IVwxKSpO0XaF9DPfNBKS7Za
# zch8NF5vp7eaZ2CVNxpqumzTCNSOxm+SAWSuIr21Qomb+zzQWKhxKTVVgtmUPAW3
# 5xUUFREmDrMxSNlr/NsJyUXzdtFUUt4aS4CEeIY8y9IaaGBpPNXKFifinT7zL2gd
# FpBP9qh8SdLnEut/GcalNeJQ55IuwnKCgs+nrpuQNfVmUB5KlCX3ZA4x5HHKS+rq
# BvKWxdCyQEEGcbLe1b8Aw4wJkhU1JrPsFfxW1gaou30yZ46t4Y9F20HHfIY4/6vH
# espYMQmUiote8ladjS/nJ0+k6MvqzfpzPDOy5y6gqztiT96Fv/9bH7mQyogxG9QE
# PHrPV6/7umw052AkyiLA6tQbZl1KhBtTasySkuJDpsZGKdlsjg4u70EwgWbVRSX1
# Wd4+zoFpp4Ra+MlKM2baoD6x0VR4RjSpWM8o5a6D8bpfm4CLKczsG7ZrIGNTAgMB
# AAGjggFdMIIBWTASBgNVHRMBAf8ECDAGAQH/AgEAMB0GA1UdDgQWBBTvb1NK6eQG
# fHrK4pBW9i/USezLTjAfBgNVHSMEGDAWgBTs1+OC0nFdZEzfLmc/57qYrhwPTzAO
# BgNVHQ8BAf8EBAMCAYYwEwYDVR0lBAwwCgYIKwYBBQUHAwgwdwYIKwYBBQUHAQEE
# azBpMCQGCCsGAQUFBzABhhhodHRwOi8vb2NzcC5kaWdpY2VydC5jb20wQQYIKwYB
# BQUHMAKGNWh0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0
# ZWRSb290RzQuY3J0MEMGA1UdHwQ8MDowOKA2oDSGMmh0dHA6Ly9jcmwzLmRpZ2lj
# ZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRSb290RzQuY3JsMCAGA1UdIAQZMBcwCAYG
# Z4EMAQQCMAsGCWCGSAGG/WwHATANBgkqhkiG9w0BAQsFAAOCAgEAF877FoAc/gc9
# EXZxML2+C8i1NKZ/zdCHxYgaMH9Pw5tcBnPw6O6FTGNpoV2V4wzSUGvI9NAzaoQk
# 97frPBtIj+ZLzdp+yXdhOP4hCFATuNT+ReOPK0mCefSG+tXqGpYZ3essBS3q8nL2
# UwM+NMvEuBd/2vmdYxDCvwzJv2sRUoKEfJ+nN57mQfQXwcAEGCvRR2qKtntujB71
# WPYAgwPyWLKu6RnaID/B0ba2H3LUiwDRAXx1Neq9ydOal95CHfmTnM4I+ZI2rVQf
# jXQA1WSjjf4J2a7jLzWGNqNX+DF0SQzHU0pTi4dBwp9nEC8EAqoxW6q17r0z0noD
# js6+BFo+z7bKSBwZXTRNivYuve3L2oiKNqetRHdqfMTCW/NmKLJ9M+MtucVGyOxi
# Df06VXxyKkOirv6o02OoXN4bFzK0vlNMsvhlqgF2puE6FndlENSmE+9JGYxOGLS/
# D284NHNboDGcmWXfwXRy4kbu4QFhOm0xJuF2EZAOk5eCkhSxZON3rGlHqhpB/8Ml
# uDezooIs8CVnrpHMiD2wL40mm53+/j7tFaxYKIqL0Q4ssd8xHZnIn/7GELH3IdvG
# 2XlM9q7WP/UwgOkw/HQtyRN62JK4S1C8uw3PdBunvAZapsiI5YKdvlarEvf8EA+8
# hcpSM9LHJmyrxaFtoza2zNaQ9k+5t1wwggbtMIIE1aADAgECAhAKgO8YS43xBYLR
# xHanlXRoMA0GCSqGSIb3DQEBCwUAMGkxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5E
# aWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1l
# U3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYgMjAyNSBDQTEwHhcNMjUwNjA0MDAwMDAw
# WhcNMzYwOTAzMjM1OTU5WjBjMQswCQYDVQQGEwJVUzEXMBUGA1UEChMORGlnaUNl
# cnQsIEluYy4xOzA5BgNVBAMTMkRpZ2lDZXJ0IFNIQTI1NiBSU0E0MDk2IFRpbWVz
# dGFtcCBSZXNwb25kZXIgMjAyNSAxMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIIC
# CgKCAgEA0EasLRLGntDqrmBWsytXum9R/4ZwCgHfyjfMGUIwYzKomd8U1nH7C8Dr
# 0cVMF3BsfAFI54um8+dnxk36+jx0Tb+k+87H9WPxNyFPJIDZHhAqlUPt281mHrBb
# ZHqRK71Em3/hCGC5KyyneqiZ7syvFXJ9A72wzHpkBaMUNg7MOLxI6E9RaUueHTQK
# WXymOtRwJXcrcTTPPT2V1D/+cFllESviH8YjoPFvZSjKs3SKO1QNUdFd2adw44wD
# cKgH+JRJE5Qg0NP3yiSyi5MxgU6cehGHr7zou1znOM8odbkqoK+lJ25LCHBSai25
# CFyD23DZgPfDrJJJK77epTwMP6eKA0kWa3osAe8fcpK40uhktzUd/Yk0xUvhDU6l
# vJukx7jphx40DQt82yepyekl4i0r8OEps/FNO4ahfvAk12hE5FVs9HVVWcO5J4dV
# mVzix4A77p3awLbr89A90/nWGjXMGn7FQhmSlIUDy9Z2hSgctaepZTd0ILIUbWuh
# KuAeNIeWrzHKYueMJtItnj2Q+aTyLLKLM0MheP/9w6CtjuuVHJOVoIJ/DtpJRE7C
# e7vMRHoRon4CWIvuiNN1Lk9Y+xZ66lazs2kKFSTnnkrT3pXWETTJkhd76CIDBbTR
# ofOsNyEhzZtCGmnQigpFHti58CSmvEyJcAlDVcKacJ+A9/z7eacCAwEAAaOCAZUw
# ggGRMAwGA1UdEwEB/wQCMAAwHQYDVR0OBBYEFOQ7/PIx7f391/ORcWMZUEPPYYzo
# MB8GA1UdIwQYMBaAFO9vU0rp5AZ8esrikFb2L9RJ7MtOMA4GA1UdDwEB/wQEAwIH
# gDAWBgNVHSUBAf8EDDAKBggrBgEFBQcDCDCBlQYIKwYBBQUHAQEEgYgwgYUwJAYI
# KwYBBQUHMAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBdBggrBgEFBQcwAoZR
# aHR0cDovL2NhY2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0VHJ1c3RlZEc0VGlt
# ZVN0YW1waW5nUlNBNDA5NlNIQTI1NjIwMjVDQTEuY3J0MF8GA1UdHwRYMFYwVKBS
# oFCGTmh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRHNFRp
# bWVTdGFtcGluZ1JTQTQwOTZTSEEyNTYyMDI1Q0ExLmNybDAgBgNVHSAEGTAXMAgG
# BmeBDAEEAjALBglghkgBhv1sBwEwDQYJKoZIhvcNAQELBQADggIBAGUqrfEcJwS5
# rmBB7NEIRJ5jQHIh+OT2Ik/bNYulCrVvhREafBYF0RkP2AGr181o2YWPoSHz9iZE
# N/FPsLSTwVQWo2H62yGBvg7ouCODwrx6ULj6hYKqdT8wv2UV+Kbz/3ImZlJ7YXwB
# D9R0oU62PtgxOao872bOySCILdBghQ/ZLcdC8cbUUO75ZSpbh1oipOhcUT8lD8QA
# GB9lctZTTOJM3pHfKBAEcxQFoHlt2s9sXoxFizTeHihsQyfFg5fxUFEp7W42fNBV
# N4ueLaceRf9Cq9ec1v5iQMWTFQa0xNqItH3CPFTG7aEQJmmrJTV3Qhtfparz+BW6
# 0OiMEgV5GWoBy4RVPRwqxv7Mk0Sy4QHs7v9y69NBqycz0BZwhB9WOfOu/CIJnzkQ
# TwtSSpGGhLdjnQ4eBpjtP+XB3pQCtv4E5UCSDag6+iX8MmB10nfldPF9SVD7weCC
# 3yXZi/uuhqdwkgVxuiMFzGVFwYbQsiGnoa9F5AaAyBjFBtXVLcKtapnMG3VH3EmA
# p/jsJ3FVF3+d1SVDTmjFjLbNFZUWMXuZyvgLfgyPehwJVxwC+UpX2MSey2ueIu9T
# HFVkT+um1vshETaWyQo8gmBto/m3acaP9QsuLj3FNwFlTxq25+T4QwX9xa6ILs84
# ZPvmpovq90K8eWyG2N01c4IhSOxqt81nMYIFEDCCBQwCAQEwNDAgMR4wHAYDVQQD
# DBVQaGluSVQtUFNzY3JpcHRzX1NpZ24CEHePOzJf0KCMSL6wELasExMwDQYJYIZI
# AWUDBAIBBQCggYQwGAYKKwYBBAGCNwIBDDEKMAigAoAAoQKAADAZBgkqhkiG9w0B
# CQMxDAYKKwYBBAGCNwIBBDAcBgorBgEEAYI3AgELMQ4wDAYKKwYBBAGCNwIBFTAv
# BgkqhkiG9w0BCQQxIgQgoqq6P4PAGWGZnjgvY8T3J9oUJFcIT7ZFZZov8OrvUXIw
# DQYJKoZIhvcNAQEBBQAEggEAS4ZoqtS4qAdCE64gF4Tz/umfuWyi9C7UX9Tz0jYC
# Cu86a+aUryQzSYd4AMFKWfphsHtxtabDAw4jMWt6TJPyNHcqbW3p06XQPgUUC7JJ
# DP1z7Lyu/WziDN06rm+tzyP6l9vdJMAmEWhjZ5+Rulxlyl3goJHaLfOXJx+yQtOE
# GSG+GKdYo/fVJuGFAER19HAqpkcyyFhsklw3s/351O4DEvDGa85LXp3D4ajkUlb3
# 38mi8ZUgUG2KEExMr/f4ej6wjBthrJTaxBY6UZWPvdOTFsFq8J8xBr5+TGqoV54+
# QJXx6/WEk+bSCpbK5PP1uQyIZMSxoDaViC6O7Pz1BEF0saGCAyYwggMiBgkqhkiG
# 9w0BCQYxggMTMIIDDwIBATB9MGkxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdp
# Q2VydCwgSW5jLjFBMD8GA1UEAxM4RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3Rh
# bXBpbmcgUlNBNDA5NiBTSEEyNTYgMjAyNSBDQTECEAqA7xhLjfEFgtHEdqeVdGgw
# DQYJYIZIAWUDBAIBBQCgaTAYBgkqhkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqG
# SIb3DQEJBTEPFw0yNTA3MTIwODAyMzNaMC8GCSqGSIb3DQEJBDEiBCAyJPbFyTZW
# gpiRyWMy3afwCriX4QpVQNqoD+A5CS/x1TANBgkqhkiG9w0BAQEFAASCAgCVTmfD
# +jKwCw5SsuD3qJxvIveUVLW8ZRKG46MDiy7/gkuhztNjd4+D5sPLQAxSS04O1LwP
# 15dpOfN/WABUYMvv48pCcqK+PgkkzSfL+NEyTjcrPxESz3Lh3DkR6ChNQYvPQQ7H
# QRVs4hYYLcKFQIB3qxrZMbHJT7PcDNreFVQ7lP2svrVm1/EyP8L0xND/QmMrPZrL
# c6MvHJW01SDmso7CBmjJH6qI7crRkxvCh7JXF1I54jkqysG05wX8CYaDUIynQplj
# vNdQ+LFOKaDYH7uYeAgHqDgaPmO3Its4o/xAwGmWPHplxe6V9higkHMPEXJs+9WE
# HGgRzVmeKKkb+vxqdgdPeIqvk49DCUt2y0bdu1rI0bQ8CHkbfZUCjvacDv0m9z+8
# /prXWBKfNIqsfUhROtPpP+lWH1e0gX+azfInD6bxE+IX2jBaoHo/xtH2tMOMLnFS
# XZZIj8ZXhz/wxeB51YA6xJWFP+CYjZP2JyE3Tbb3SiZn7s1KQfqkdJH3gVgtnfm6
# /MnmMYVXFmC6Hv6dVmSBoSVibhGCK1t19WSoMAMX18PyYsMeAObvMijHkROtbwJe
# sLYWZwt5Wxk+7aSfpq+LZ1Xvuu8j+hkNljGwhS79sN3PTpE3mD5estmr4EeDmc0t
# VX761ZlDRMH7Ytb+Mt78E91emktw3J0C+2RwOQ==
# SIG # End signature block
