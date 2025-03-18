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
