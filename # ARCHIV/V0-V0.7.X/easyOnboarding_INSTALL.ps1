param(
    [string]$htmlFile,
    [string]$pdfFile,
    [string]$wkhtmltopdfPath
)

# Falls alle Parameter vorhanden sind, fuehre die PDF-Konvertierung aus
if ($htmlFile -and $pdfFile -and $wkhtmltopdfPath) {
    # Lokalen Pfad in URL umwandeln
    $htmlFileURL = "file:///" + ($htmlFile -replace '\\', '/')
    $arguments = @("--enable-local-file-access", $htmlFileURL, $pdfFile)
    Write-Host "Starte wkhtmltopdf mit folgenden Argumenten:" $arguments
    Start-Process -FilePath $wkhtmltopdfPath -ArgumentList $arguments -Wait -NoNewWindow
    if (Test-Path $pdfFile) {
        Write-Host "PDF erfolgreich erstellt: $pdfFile"
    }
    else {
        Write-Error "PDF konnte nicht erstellt werden."
    }
    return
}

# Andernfalls wird die Installations-GUI angezeigt

function Show-InstallerForm {
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Installations-Assistent"
    $form.Size = New-Object System.Drawing.Size(500,300)
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $form.MaximizeBox = $false

    # Status-Label
    $lblStatus = New-Object System.Windows.Forms.Label
    $lblStatus.Location = New-Object System.Drawing.Point(20,20)
    $lblStatus.Size = New-Object System.Drawing.Size(440,40)
    $lblStatus.Text = "Bereit..."
    $lblStatus.Font = New-Object System.Drawing.Font("Segoe UI",10)
    $form.Controls.Add($lblStatus)

    # ProgressBar
    $progressBar = New-Object System.Windows.Forms.ProgressBar
    $progressBar.Location = New-Object System.Drawing.Point(20,70)
    $progressBar.Size = New-Object System.Drawing.Size(440,25)
    $progressBar.Minimum = 0
    $progressBar.Maximum = 100
    $progressBar.Style = "Continuous"
    $form.Controls.Add($progressBar)

    # Checkbox: Automatisch Ausfuehrungsrichtlinie auf RemoteSigned setzen
    $chkSetPolicy = New-Object System.Windows.Forms.CheckBox
    $chkSetPolicy.Location = New-Object System.Drawing.Point(20,105)
    $chkSetPolicy.Size = New-Object System.Drawing.Size(300,20)
    $chkSetPolicy.Text = "Automatisch auf RemoteSigned setzen"
    $chkSetPolicy.Checked = $true
    $form.Controls.Add($chkSetPolicy)

    # Button "Install" (hellgruen)
    $btnInstall = New-Object System.Windows.Forms.Button
    $btnInstall.Text = "Install"
    $btnInstall.Size = New-Object System.Drawing.Size(100,30)
    $btnInstall.Location = New-Object System.Drawing.Point(100,200)
    $btnInstall.BackColor = [System.Drawing.Color]::LightGreen
    $form.Controls.Add($btnInstall)

    # Button "Close" (hellrot)
    $btnClose = New-Object System.Windows.Forms.Button
    $btnClose.Text = "Close"
    $btnClose.Size = New-Object System.Drawing.Size(100,30)
    $btnClose.Location = New-Object System.Drawing.Point(300,200)
    $btnClose.BackColor = [System.Drawing.Color]::LightCoral
    $form.Controls.Add($btnClose)

    # Rueckgabe eines Hashtables mit den GUI-Elementen
    return @{
        Form = $form;
        Label = $lblStatus;
        ProgressBar = $progressBar;
        SetPolicyCheckBox = $chkSetPolicy;
        InstallButton = $btnInstall;
        CloseButton = $btnClose
    }
}

function Run-Installation {
    param(
        $guiElements
    )
    $lbl = $guiElements.Label
    $pb = $guiElements.ProgressBar
    $setPolicy = $guiElements.SetPolicyCheckBox

    try {
        # Schritt 1: Ausfuehrungsrichtlinie pruefen
        $lbl.Text = "Pruefe Ausfuehrungsrichtlinie..."
        [System.Windows.Forms.Application]::DoEvents()
        $currentPolicy = Get-ExecutionPolicy
        if ($currentPolicy -eq "Restricted" -or $currentPolicy -eq "AllSigned") {
            if ($setPolicy.Checked) {
                Set-ExecutionPolicy RemoteSigned -Scope Process -Force
                $lbl.Text = "Ausfuehrungsrichtlinie auf RemoteSigned gesetzt."
            }
            else {
                $lbl.Text = "Ausfuehrungsrichtlinie ($currentPolicy) ist zu restriktiv."
            }
        }
        else {
            $lbl.Text = "Ausfuehrungsrichtlinie ($currentPolicy) ist akzeptabel."
        }
        Start-Sleep -Seconds 1
        $pb.Value = 20
        [System.Windows.Forms.Application]::DoEvents()

        # Schritt 2: ActiveDirectory-Modul pruefen
        $lbl.Text = "Pruefe ActiveDirectory-Modul..."
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
        $pb.Value = 40
        [System.Windows.Forms.Application]::DoEvents()

        # Schritt 3: PSWritePDF-Modul pruefen, installieren und importieren
        $lbl.Text = "Pruefe PSWritePDF-Modul..."
        [System.Windows.Forms.Application]::DoEvents()
        if (-not (Get-Module -ListAvailable -Name PSWritePDF)) {
            $lbl.Text = "PSWritePDF-Modul nicht gefunden. Installation wird gestartet..."
            [System.Windows.Forms.Application]::DoEvents()
            try {
                Install-Module PSWritePDF -Scope CurrentUser -Force -AllowClobber
                $lbl.Text = "PSWritePDF-Modul erfolgreich installiert."
            }
            catch {
                Throw "Fehler bei der Installation von PSWritePDF: $($_.Exception.Message)"
            }
        }
        else {
            $lbl.Text = "PSWritePDF-Modul ist bereits installiert."
        }
        Start-Sleep -Seconds 1
        $pb.Value = 60
        [System.Windows.Forms.Application]::DoEvents()

        try {
            Import-Module PSWritePDF -ErrorAction Stop
            $lbl.Text = "PSWritePDF-Modul erfolgreich importiert."
        }
        catch {
            $lbl.Text = "Fehler beim Import von PSWritePDF."
        }
        Start-Sleep -Seconds 1
        $pb.Value = 80
        [System.Windows.Forms.Application]::DoEvents()

        # Weitere Schritte koennen hier ergänzt werden...
        $lbl.Text = "Installation abgeschlossen."
        $pb.Value = 90
        [System.Windows.Forms.Application]::DoEvents()
        Start-Sleep -Seconds 2
        $lbl.Text = "Alle erforderlichen Module und Tools wurden geprueft."
        $pb.Value = 100
        [System.Windows.Forms.Application]::DoEvents()
        Start-Sleep -Seconds 2
    }
    catch {
        $lbl.Text = "Fehler: $($_.Exception.Message)"
    }
}

# Hauptteil: GUI anzeigen und Button-Ereignisse verarbeiten
$gui = Show-InstallerForm
$gui.InstallButton.Add_Click({ Run-Installation -guiElements $gui })
$gui.CloseButton.Add_Click({ $gui.Form.Close() })

[System.Windows.Forms.Application]::Run($gui.Form)

# SIG # Begin signature block
# MIIbywYJKoZIhvcNAQcCoIIbvDCCG7gCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBgY2Tg0tHehFpi
# 4Q9YlmVmGB318MCEqGH8X1vwY+jTY6CCFhcwggMQMIIB+KADAgECAhB3jzsyX9Cg
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
# DQEJBDEiBCBWIWPHh6NCWZ7rxu6tJIWVulvDs+YXVHJKkhqhamRRfTANBgkqhkiG
# 9w0BAQEFAASCAQBBnboDM8H6cuSciHXnS/Ly2MzrBVymxv91j2mhBM9C5zEPzQyG
# pkqojgYwya/B5YLHQaDA0iApG9XDdTC8Dio26GRMSBQLGNYybaRzU05uZvoicLYx
# KRHuYwF+0OT8Ixb1WTxgRzm2xC/3iVcO7GBtyGVvIIoGQbQKAaBCymsdwoDuWlUJ
# i2tKCzMBFNhOdJ7VbIlyf26TON/nRkVmuYfZhab8DwYs4d0wGbWZ/7rEu6IXx9nv
# kjiytG8v+gcIo55k3/QZtSpkQbeOQV06/iQeSKG0RyVBaX6jnKWNyVCyrhFwh1J/
# x4MZJEkvzPjHDMOsijAytsnUwsJSCYV8c5yCoYIDIDCCAxwGCSqGSIb3DQEJBjGC
# Aw0wggMJAgEBMHcwYzELMAkGA1UEBhMCVVMxFzAVBgNVBAoTDkRpZ2lDZXJ0LCBJ
# bmMuMTswOQYDVQQDEzJEaWdpQ2VydCBUcnVzdGVkIEc0IFJTQTQwOTYgU0hBMjU2
# IFRpbWVTdGFtcGluZyBDQQIQC65mvFq6f5WHxvnpBOMzBDANBglghkgBZQMEAgEF
# AKBpMBgGCSqGSIb3DQEJAzELBgkqhkiG9w0BBwEwHAYJKoZIhvcNAQkFMQ8XDTI1
# MDcwNTEwMTE1NFowLwYJKoZIhvcNAQkEMSIEIKLbVU5sYTFuINpJhstNUxYVhs+Z
# KywX+SxV1a64q3p7MA0GCSqGSIb3DQEBAQUABIICAFU4O21x6uuwKp1jUhQ1Q71+
# wTwh3+p3+rDyMQ5NLeJNFEY7nMVAzM2iUpGvRNTSnDkKldjVg/b/nQYuKuuHXPN5
# EXVjczT3WwDi2taayJcVAJ8WbEcYTxx+wjHtjGSnp04aC5ZO44pvMl/ZQ8Eu91QY
# 0F9O+qMuVDl60vLXpmpiVPYYlWSlnlN+aX3bCn7iADbKUD/uV7dTanea22KQRbdA
# Yee8SOumkQXNNMMd5SiVeEaLkZtwajw6bLu2nsrFi8dtqZp2tRg+4WujF42EBQF5
# 31MFYFFIJrHXsJh5SnQDQX7ViE9VRS9wLbSYdIiigZ5+/8ffJPeKADxTXaqEz/eO
# ppztPJybXH6tuBBv6iJaY58snt+gqx8nvPTxmYz8Wb/k/Nm1BTLgKf3EOxw5yRgX
# C/EbgCsDLQdSvgCq2eW1wED5FOv/VABBFudbv6HmZJPkhcru/aY9gr3jKSDHEZQx
# /CPkIgVtGniqvOPSh0FCfhFF+6IIr4ypGOMPakx9SiSEeStODAMJr+jsa9AxhAiX
# BB7mLGJ+NhS461KpI4vLqpra+guY493EnLA+EBHZVSeawngEi/q+mxnlIlqwlZzN
# vQfH44lF/z3OdJ8f6Lud+qFtzVcy/nkXASqKLs0qYNIIwEF01vPZNyJSSRm2lxG+
# K1hHg2EttG+TwmfP4e1T
# SIG # End signature block
