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
