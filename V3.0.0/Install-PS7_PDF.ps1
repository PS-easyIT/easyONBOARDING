# Parameter block - all required parameters are declared here
param(
    [string]$wkhtmltopdfPath = "",
    # URL to PowerShell 7 setup package (adjust if needed)
    [string]$PS7SetupURL = "https://github.com/PowerShell/PowerShell/releases/download/v7.5.0/PowerShell-7.5.0-win-x64.msi",
    # URL to wkhtmltopdf setup package
    [string]$wkhtmltopdfURL = "https://github.com/wkhtmltopdf/packaging/releases/download/0.12.6-1/wkhtmltox-0.12.6-1.msvc2015-win64.exe"
)

# Add required .NET assemblies
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# -----------------------------------
# Function: Display Installation Wizard (GUI)
# -----------------------------------
function Show-InstallerForm {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Installation Wizard"
    $form.Size = New-Object System.Drawing.Size(500,400)
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $form.MaximizeBox = $false

    # Status Label (all messages are displayed here)
    $lblStatus = New-Object System.Windows.Forms.Label
    $lblStatus.Location = New-Object System.Drawing.Point(20,20)
    $lblStatus.Size = New-Object System.Drawing.Size(440,60)
    $lblStatus.Text = "Ready..."
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

    # Checkbox: Automatically set execution policy to RemoteSigned
    $chkSetPolicy = New-Object System.Windows.Forms.CheckBox
    $chkSetPolicy.Location = New-Object System.Drawing.Point(20,130)
    $chkSetPolicy.Size = New-Object System.Drawing.Size(300,20)
    $chkSetPolicy.Text = "Automatically set to RemoteSigned"
    $chkSetPolicy.Checked = $true
    $form.Controls.Add($chkSetPolicy)
    
    # Status label for Execution Policy
    $lblPolicyStatus = New-Object System.Windows.Forms.Label
    $lblPolicyStatus.Location = New-Object System.Drawing.Point(330,130)
    $lblPolicyStatus.Size = New-Object System.Drawing.Size(150,20)
    $lblPolicyStatus.Text = "CHECKING..."
    $lblPolicyStatus.Font = New-Object System.Drawing.Font("Segoe UI",9,[System.Drawing.FontStyle]::Bold)
    $form.Controls.Add($lblPolicyStatus)

    # Checkbox: Install PowerShell 7
    $chkInstallPS7 = New-Object System.Windows.Forms.CheckBox
    $chkInstallPS7.Location = New-Object System.Drawing.Point(20,160)
    $chkInstallPS7.Size = New-Object System.Drawing.Size(300,20)
    $chkInstallPS7.Text = "Install PowerShell 7"
    $chkInstallPS7.Checked = $true
    $form.Controls.Add($chkInstallPS7)
    
    # Status label for PowerShell 7
    $lblPS7Status = New-Object System.Windows.Forms.Label
    $lblPS7Status.Location = New-Object System.Drawing.Point(330,160)
    $lblPS7Status.Size = New-Object System.Drawing.Size(150,20)
    $lblPS7Status.Text = "CHECKING..."
    $lblPS7Status.Font = New-Object System.Drawing.Font("Segoe UI",9,[System.Drawing.FontStyle]::Bold)
    $form.Controls.Add($lblPS7Status)

    # Checkbox: Install ActiveDirectory module
    $chkInstallADModule = New-Object System.Windows.Forms.CheckBox
    $chkInstallADModule.Location = New-Object System.Drawing.Point(20,190)
    $chkInstallADModule.Size = New-Object System.Drawing.Size(300,20)
    $chkInstallADModule.Text = "Install ActiveDirectory module"
    $chkInstallADModule.Checked = $true
    $form.Controls.Add($chkInstallADModule)
    
    # Status label for ActiveDirectory module
    $lblADStatus = New-Object System.Windows.Forms.Label
    $lblADStatus.Location = New-Object System.Drawing.Point(330,190)
    $lblADStatus.Size = New-Object System.Drawing.Size(150,20)
    $lblADStatus.Text = "CHECKING..."
    $lblADStatus.Font = New-Object System.Drawing.Font("Segoe UI",9,[System.Drawing.FontStyle]::Bold)
    $form.Controls.Add($lblADStatus)

    # Checkbox: Install wkhtmltopdf
    $chkInstallwkhtmltopdf = New-Object System.Windows.Forms.CheckBox
    $chkInstallwkhtmltopdf.Location = New-Object System.Drawing.Point(20,220)
    $chkInstallwkhtmltopdf.Size = New-Object System.Drawing.Size(300,20)
    $chkInstallwkhtmltopdf.Text = "Install wkhtmltopdf"
    $chkInstallwkhtmltopdf.Checked = $true
    $form.Controls.Add($chkInstallwkhtmltopdf)
    
    # Status label for wkhtmltopdf
    $lblWkStatus = New-Object System.Windows.Forms.Label
    $lblWkStatus.Location = New-Object System.Drawing.Point(330,220)
    $lblWkStatus.Size = New-Object System.Drawing.Size(150,20)
    $lblWkStatus.Text = "CHECKING..."
    $lblWkStatus.Font = New-Object System.Drawing.Font("Segoe UI",9,[System.Drawing.FontStyle]::Bold)
    $form.Controls.Add($lblWkStatus)

    # New Checkbox: Copy current folder contents
    $chkCopyContents = New-Object System.Windows.Forms.CheckBox
    $chkCopyContents.Location = New-Object System.Drawing.Point(20,250)
    $chkCopyContents.Size = New-Object System.Drawing.Size(440,20)
    $chkCopyContents.Text = "Kopiere Inhalte des Ordners nach C:\easyIT\easyONBOARDING"
    $chkCopyContents.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $chkCopyContents.Checked = $true
    $form.Controls.Add($chkCopyContents)

    # Button "Install" (light green)
    $btnInstall = New-Object System.Windows.Forms.Button
    $btnInstall.Text = "Install"
    $btnInstall.Size = New-Object System.Drawing.Size(100,30)
    $btnInstall.Location = New-Object System.Drawing.Point(100,300)
    $btnInstall.BackColor = [System.Drawing.Color]::LightGreen
    $form.Controls.Add($btnInstall)

    # Button "Close" (light red)
    $btnClose = New-Object System.Windows.Forms.Button
    $btnClose.Text = "Close"
    $btnClose.Size = New-Object System.Drawing.Size(100,30)
    $btnClose.Location = New-Object System.Drawing.Point(300,300)
    $btnClose.BackColor = [System.Drawing.Color]::LightCoral
    $form.Controls.Add($btnClose)
    
    # Load Form event to immediately check installations
    # Use script: prefix to access these variables in the scriptblock
    $script:lblStatus = $lblStatus
    $script:lblPolicyStatus = $lblPolicyStatus
    $script:lblPS7Status = $lblPS7Status
    $script:lblADStatus = $lblADStatus
    $script:lblWkStatus = $lblWkStatus
    
    $form.Add_Load({
        $script:lblStatus.Text = "Checking installation status..."
        [System.Windows.Forms.Application]::DoEvents()
        
        # Check execution policy status
        $currentPolicy = Get-ExecutionPolicy
        if ($currentPolicy -eq "RemoteSigned" -or $currentPolicy -eq "Unrestricted") {
            $script:lblPolicyStatus.Text = "CONFIGURED"
            $script:lblPolicyStatus.ForeColor = [System.Drawing.Color]::Green
        } else {
            $script:lblPolicyStatus.Text = "NOT CONFIGURED"
            $script:lblPolicyStatus.ForeColor = [System.Drawing.Color]::Red
        }
        
        # Check PowerShell 7 status
        $ps7 = Get-Command pwsh.exe -ErrorAction SilentlyContinue
        if ($ps7) {
            $script:lblPS7Status.Text = "INSTALLED"
            $script:lblPS7Status.ForeColor = [System.Drawing.Color]::Green
        } else {
            $script:lblPS7Status.Text = "NOT INSTALLED"
            $script:lblPS7Status.ForeColor = [System.Drawing.Color]::Red
        }
        
        # Check ActiveDirectory module status
        if (Get-Module -ListAvailable -Name ActiveDirectory) {
            $script:lblADStatus.Text = "INSTALLED"
            $script:lblADStatus.ForeColor = [System.Drawing.Color]::Green
        } else {
            $script:lblADStatus.Text = "NOT INSTALLED"
            $script:lblADStatus.ForeColor = [System.Drawing.Color]::Red
        }
        
        # Check wkhtmltopdf status
        $wkhtmltopdfInstalled = $false
        $possiblePaths = @(
            "${env:ProgramFiles}\wkhtmltopdf\bin\wkhtmltopdf.exe",
            "${env:ProgramFiles(x86)}\wkhtmltopdf\bin\wkhtmltopdf.exe",
            "C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe",
            "C:\Program Files (x86)\wkhtmltopdf\bin\wkhtmltopdf.exe"
        )
        
        foreach ($path in $possiblePaths) {
            if (Test-Path $path) {
                $wkhtmltopdfInstalled = $true
                $script:wkhtmltopdfPath = $path
                break
            }
        }
        
        if ($wkhtmltopdfInstalled) {
            $script:lblWkStatus.Text = "INSTALLED"
            $script:lblWkStatus.ForeColor = [System.Drawing.Color]::Green
        } else {
            $script:lblWkStatus.Text = "NOT INSTALLED"
            $script:lblWkStatus.ForeColor = [System.Drawing.Color]::Red
        }
        
        $script:lblStatus.Text = "Ready..."
    })

    return @{
        Form = $form;
        Label = $lblStatus;
        ProgressBar = $progressBar;
        SetPolicyCheckBox = $chkSetPolicy;
        InstallPS7CheckBox = $chkInstallPS7;
        InstallADModuleCheckBox = $chkInstallADModule;
        InstallwkhtmltopdfCheckBox = $chkInstallwkhtmltopdf;
        InstallCopyContentsCheckBox = $chkCopyContents;
        InstallButton = $btnInstall;
        CloseButton = $btnClose;
        PolicyStatusLabel = $lblPolicyStatus;
        PS7StatusLabel = $lblPS7Status;
        ADStatusLabel = $lblADStatus;
        WkStatusLabel = $lblWkStatus
    }
}

# -----------------------------------
# Function: Check installation status of components
# -----------------------------------
function Update-InstallationStatus {
    param(
        $guiElements
    )
    
    # Check execution policy status
    $currentPolicy = Get-ExecutionPolicy
    if ($currentPolicy -eq "RemoteSigned" -or $currentPolicy -eq "Unrestricted") {
        $guiElements.PolicyStatusLabel.Text = "CONFIGURED"
        $guiElements.PolicyStatusLabel.ForeColor = [System.Drawing.Color]::Green
    } else {
        $guiElements.PolicyStatusLabel.Text = "NOT CONFIGURED"
        $guiElements.PolicyStatusLabel.ForeColor = [System.Drawing.Color]::Red
    }
    
    # Check PowerShell 7 status
    $ps7 = Get-Command pwsh.exe -ErrorAction SilentlyContinue
    if ($ps7) {
        $guiElements.PS7StatusLabel.Text = "INSTALLED"
        $guiElements.PS7StatusLabel.ForeColor = [System.Drawing.Color]::Green
    } else {
        $guiElements.PS7StatusLabel.Text = "NOT INSTALLED"
        $guiElements.PS7StatusLabel.ForeColor = [System.Drawing.Color]::Red
    }
    
    # Check ActiveDirectory module status
    if (Get-Module -ListAvailable -Name ActiveDirectory) {
        $guiElements.ADStatusLabel.Text = "INSTALLED"
        $guiElements.ADStatusLabel.ForeColor = [System.Drawing.Color]::Green
    } else {
        $guiElements.ADStatusLabel.Text = "NOT INSTALLED"
        $guiElements.ADStatusLabel.ForeColor = [System.Drawing.Color]::Red
    }
    
    # Check wkhtmltopdf status
    $wkhtmltopdfInstalled = $false
    $possiblePaths = @(
        "${env:ProgramFiles}\wkhtmltopdf\bin\wkhtmltopdf.exe",
        "${env:ProgramFiles(x86)}\wkhtmltopdf\bin\wkhtmltopdf.exe",
        "C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe",
        "C:\Program Files (x86)\wkhtmltopdf\bin\wkhtmltopdf.exe"
    )
    
    foreach ($path in $possiblePaths) {
        if (Test-Path $path) {
            $wkhtmltopdfInstalled = $true
            $global:wkhtmltopdfPath = $path
            break
        }
    }
    
    if ($wkhtmltopdfInstalled) {
        $guiElements.WkStatusLabel.Text = "INSTALLED"
        $guiElements.WkStatusLabel.ForeColor = [System.Drawing.Color]::Green
    } else {
        $guiElements.WkStatusLabel.Text = "NOT INSTALLED"
        $guiElements.WkStatusLabel.ForeColor = [System.Drawing.Color]::Red
    }
}

# -----------------------------------
# Function: Execute Installation (all steps)
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
    $chkInstallwkhtmltopdf = $guiElements.InstallwkhtmltopdfCheckBox
    $chkCopyContents = $guiElements.InstallCopyContentsCheckBox

    try {
        # Step 1: Check execution policy
        $lbl.Text = "Checking execution policy..."
        [System.Windows.Forms.Application]::DoEvents()
        $currentPolicy = Get-ExecutionPolicy
        if ($currentPolicy -eq "Restricted" -or $currentPolicy -eq "AllSigned") {
            if ($chkSetPolicy.Checked) {
                Set-ExecutionPolicy RemoteSigned -Scope Process -Force
                $lbl.Text = "Execution policy set to RemoteSigned."
                $guiElements.PolicyStatusLabel.Text = "CONFIGURED"
                $guiElements.PolicyStatusLabel.ForeColor = [System.Drawing.Color]::Green
            }
            else {
                $lbl.Text = "Execution policy ($currentPolicy) is too restrictive."
                $guiElements.PolicyStatusLabel.Text = "NOT CONFIGURED"
                $guiElements.PolicyStatusLabel.ForeColor = [System.Drawing.Color]::Red
            }
        }
        else {
            $lbl.Text = "Execution policy ($currentPolicy) is acceptable."
            $guiElements.PolicyStatusLabel.Text = "CONFIGURED"
            $guiElements.PolicyStatusLabel.ForeColor = [System.Drawing.Color]::Green
        }
        Start-Sleep -Seconds 1
        $pb.Value = 20
        [System.Windows.Forms.Application]::DoEvents()

        # Step 2: Remove msstore source to avoid manual confirmations
        try {
            winget source remove msstore -ErrorAction Stop | Out-Null
            $lbl.Text = "msstore source removed."
        }
        catch {
            $lbl.Text = "msstore source could not be removed or is already removed."
        }
        Start-Sleep -Seconds 1
        $pb.Value = 30
        [System.Windows.Forms.Application]::DoEvents()

        # Step 3: Check PowerShell 7 and install if needed
        if ($chkInstallPS7.Checked) {
            $lbl.Text = "Checking PowerShell 7..."
            [System.Windows.Forms.Application]::DoEvents()
            $ps7 = Get-Command pwsh.exe -ErrorAction SilentlyContinue
            if ($ps7) {
                $lbl.Text = "PowerShell 7 is already installed."
                $guiElements.PS7StatusLabel.Text = "INSTALLED"
                $guiElements.PS7StatusLabel.ForeColor = [System.Drawing.Color]::Green
            }
            else {
                $lbl.Text = "PowerShell 7 not found. Checking installation options..."
                [System.Windows.Forms.Application]::DoEvents()
                
                # Check if WinGet is available
                $winget = Get-Command winget -ErrorAction SilentlyContinue
                $installSuccess = $false
                
                if ($winget) {
                    $lbl.Text = "Installing PowerShell 7 using WinGet..."
                    [System.Windows.Forms.Application]::DoEvents()
                    
                    try {
                        # Install PowerShell 7 using WinGet
                        Start-Process -FilePath "winget" -ArgumentList "install --id Microsoft.PowerShell -e" -Wait -NoNewWindow
                        
                        # Check if WinGet installation was successful
                        $ps7 = Get-Command pwsh.exe -ErrorAction SilentlyContinue
                        if ($ps7) {
                            $lbl.Text = "PowerShell 7 successfully installed using WinGet."
                            $guiElements.PS7StatusLabel.Text = "INSTALLED"
                            $guiElements.PS7StatusLabel.ForeColor = [System.Drawing.Color]::Green
                            $installSuccess = $true
                        }
                        else {
                            $lbl.Text = "WinGet installation completed but PowerShell 7 not detected. Falling back to MSI..."
                            [System.Windows.Forms.Application]::DoEvents()
                        }
                    }
                    catch {
                        $lbl.Text = "Error using WinGet: $($_.Exception.Message). Falling back to MSI installation..."
                        [System.Windows.Forms.Application]::DoEvents()
                    }
                }
                
                # If WinGet is not available or failed, use MSI installation
                if (-not $installSuccess) {
                    $lbl.Text = "Using MSI installation method..."
                    [System.Windows.Forms.Application]::DoEvents()
                    
                    # Download PS7 MSI
                    $setupPath = Join-Path $env:TEMP "PowerShell-7.5.0-win-x64.msi"
                    try {
                        $lbl.Text = "Downloading PowerShell 7 MSI..."
                        [System.Windows.Forms.Application]::DoEvents()
                        Invoke-WebRequest -Uri $PS7SetupURL -OutFile $setupPath -ErrorAction Stop
                        $lbl.Text = "PowerShell 7 downloaded. Starting installation..."
                        [System.Windows.Forms.Application]::DoEvents()
                        
                        # Install PS7
                        Start-Process -FilePath "msiexec.exe" -ArgumentList "/i", $setupPath, "/quiet", "/norestart" -Wait -NoNewWindow
                        
                        # Check if installed
                        $ps7 = Get-Command pwsh.exe -ErrorAction SilentlyContinue
                        if ($ps7) {
                            $lbl.Text = "PowerShell 7 successfully installed using MSI."
                            $guiElements.PS7StatusLabel.Text = "INSTALLED"
                            $guiElements.PS7StatusLabel.ForeColor = [System.Drawing.Color]::Green
                        }
                        else {
                            $lbl.Text = "PowerShell 7 installation failed. Please try manually."
                            $guiElements.PS7StatusLabel.Text = "NOT INSTALLED"
                            $guiElements.PS7StatusLabel.ForeColor = [System.Drawing.Color]::Red
                        }
                        
                        # Clean up
                        if (Test-Path $setupPath) {
                            Remove-Item $setupPath -Force
                        }
                    }
                    catch {
                        $lbl.Text = "Failed to install PowerShell 7: $($_.Exception.Message)"
                        $guiElements.PS7StatusLabel.Text = "NOT INSTALLED"
                        $guiElements.PS7StatusLabel.ForeColor = [System.Drawing.Color]::Red
                    }
                }
            }
            Start-Sleep -Seconds 1
            $pb.Value = 50
            [System.Windows.Forms.Application]::DoEvents()
        }

        # Step 4: Check and install ActiveDirectory module
        if ($chkInstallAD.Checked) {
            $lbl.Text = "Checking ActiveDirectory module..."
            [System.Windows.Forms.Application]::DoEvents()
            if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
                $lbl.Text = "ActiveDirectory module not found. Attempting installation..."
                [System.Windows.Forms.Application]::DoEvents()
                try {
                    Add-WindowsCapability -Online -Name "Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0" -ErrorAction Stop
                    $lbl.Text = "ActiveDirectory module successfully installed."
                    $guiElements.ADStatusLabel.Text = "INSTALLED"
                    $guiElements.ADStatusLabel.ForeColor = [System.Drawing.Color]::Green
                }
                catch {
                    $lbl.Text = "RSAT installation failed. Please install manually."
                    $guiElements.ADStatusLabel.Text = "NOT INSTALLED"
                    $guiElements.ADStatusLabel.ForeColor = [System.Drawing.Color]::Red
                }
            }
            else {
                $lbl.Text = "ActiveDirectory module is available."
                $guiElements.ADStatusLabel.Text = "INSTALLED"
                $guiElements.ADStatusLabel.ForeColor = [System.Drawing.Color]::Green
            }
            Start-Sleep -Seconds 1
            $pb.Value = 70
            [System.Windows.Forms.Application]::DoEvents()
        }

        # Step 5: Check and install wkhtmltopdf
        if ($chkInstallwkhtmltopdf.Checked) {
            $lbl.Text = "Checking wkhtmltopdf installation..."
            [System.Windows.Forms.Application]::DoEvents()
            
            # Check if wkhtmltopdf is already installed
            $wkhtmltopdfInstalled = $false
            $possiblePaths = @(
                "${env:ProgramFiles}\wkhtmltopdf\bin\wkhtmltopdf.exe",
                "${env:ProgramFiles(x86)}\wkhtmltopdf\bin\wkhtmltopdf.exe",
                "C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe",
                "C:\Program Files (x86)\wkhtmltopdf\bin\wkhtmltopdf.exe"
            )
            
            foreach ($path in $possiblePaths) {
                if (Test-Path $path) {
                    $wkhtmltopdfInstalled = $true
                    $global:wkhtmltopdfPath = $path
                    break
                }
            }
            
            if ($wkhtmltopdfInstalled) {
                $lbl.Text = "wkhtmltopdf is already installed at: $global:wkhtmltopdfPath"
                $guiElements.WkStatusLabel.Text = "INSTALLED"
                $guiElements.WkStatusLabel.ForeColor = [System.Drawing.Color]::Green
            } else {
                $lbl.Text = "wkhtmltopdf not found. Starting installation..."
                [System.Windows.Forms.Application]::DoEvents()
                
                try {
                    # Download wkhtmltopdf installer
                    $setupPath = Join-Path $env:TEMP "wkhtmltox-setup.exe"
                    $lbl.Text = "Downloading wkhtmltopdf installer..."
                    [System.Windows.Forms.Application]::DoEvents()
                    
                    Invoke-WebRequest -Uri $wkhtmltopdfURL -OutFile $setupPath -ErrorAction Stop
                    
                    $lbl.Text = "Installing wkhtmltopdf..."
                    [System.Windows.Forms.Application]::DoEvents()
                    
                    # Install wkhtmltopdf (silent installation)
                    Start-Process -FilePath $setupPath -ArgumentList "/S" -Wait -NoNewWindow
                    
                    # Check if installation was successful
                    Start-Sleep -Seconds 2
                    foreach ($path in $possiblePaths) {
                        if (Test-Path $path) {
                            $wkhtmltopdfInstalled = $true
                            $global:wkhtmltopdfPath = $path
                            break
                        }
                    }
                    
                    if ($wkhtmltopdfInstalled) {
                        $lbl.Text = "wkhtmltopdf successfully installed at: $global:wkhtmltopdfPath"
                        $guiElements.WkStatusLabel.Text = "INSTALLED"
                        $guiElements.WkStatusLabel.ForeColor = [System.Drawing.Color]::Green
                    } else {
                        $lbl.Text = "wkhtmltopdf installation completed but executable not found."
                        $guiElements.WkStatusLabel.Text = "NOT INSTALLED"
                        $guiElements.WkStatusLabel.ForeColor = [System.Drawing.Color]::Red
                    }
                    
                    # Clean up
                    if (Test-Path $setupPath) {
                        Remove-Item $setupPath -Force
                    }
                }
                catch {
                    $lbl.Text = "Error installing wkhtmltopdf: $($_.Exception.Message)"
                    $guiElements.WkStatusLabel.Text = "NOT INSTALLED"
                    $guiElements.WkStatusLabel.ForeColor = [System.Drawing.Color]::Red
                }
            }
            
            Start-Sleep -Seconds 1
            $pb.Value = 90
            [System.Windows.Forms.Application]::DoEvents()
        }

        # Completion
        $lbl.Text = "Installation completed. All selected tools have been checked."
        $pb.Value = 100
        [System.Windows.Forms.Application]::DoEvents()
        Start-Sleep -Seconds 2

        # Conditionally copy contents if checkbox is checked
        if ($chkCopyContents.Checked) {
            # Neue Funktion: Inhalte des aktuellen Ordners kopieren
            $sourceDir = Get-Location
            $destDir = "C:\easyIT\easyONBOARDING"
            if (-not (Test-Path $destDir)) {
                New-Item -Path $destDir -ItemType Directory -Force | Out-Null
            }
            Copy-Item -Path "$($sourceDir.Path)\*" -Destination $destDir -Recurse -Force

            # Create shortcut on the desktop
            $WScriptShell = New-Object -ComObject WScript.Shell
            $Shortcut = $WScriptShell.CreateShortcut("$([Environment]::GetFolderPath("Desktop"))\easyONBOARDING.lnk")
            $Shortcut.TargetPath = $destDir
            $Shortcut.Save()
        }

    }
    catch {
        $lbl.Text = "Error: $($_.Exception.Message)"
    }
}

# -----------------------------------
# Main: Display GUI and process button events
# -----------------------------------
$gui = Show-InstallerForm
$gui.InstallButton.Add_Click({ Run-Installation -guiElements $gui })
$gui.CloseButton.Add_Click({ $gui.Form.Close() })

# No need to call Update-InstallationStatus here as it's now handled in the Form's Load event
[System.Windows.Forms.Application]::Run($gui.Form)

# SIG # Begin signature block
# MIIcCAYJKoZIhvcNAQcCoIIb+TCCG/UCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDwQL/yTIXTuNit
# vrbScyzilMmANpqStGJfaGnOl/ED9KCCFk4wggMQMIIB+KADAgECAhB3jzsyX9Cg
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
# BgkqhkiG9w0BCQQxIgQgmB28SoYTvKPrOVfzEbQsCR8baXB8xrMn3XpWPMlTndkw
# DQYJKoZIhvcNAQEBBQAEggEAPm5382R2Js6gxK8v6Sl7JiT5w7sEJWJkmroi2iHu
# rs9dE+3qb4xkOTOWyRDJbypRr23GfgwjKoHlDF7ctJtvPGvOs+hgxm49+8NHVInj
# 4I0Wj1jfSIMkbu+SwMaq1ruOWU3cnxSqnj5Z6StKV8gpxejQrQmvDcDdcio13rIi
# KaeP2nv4GL1UG7SubjXYhMZheguU918W1pAcrPaSM+iBEPyo9TfDu/JxgcOFlItC
# /BOshE7GSRmveDmZggDskjhccwovZcw1sTo3FNmJkZAGnFPjVO6KUsbeOdQLAHkV
# nS3/ofgrtTZQEwpYFGneErjn8bcQ8XOHmAxy/GPQmcywb6GCAyYwggMiBgkqhkiG
# 9w0BCQYxggMTMIIDDwIBATB9MGkxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdp
# Q2VydCwgSW5jLjFBMD8GA1UEAxM4RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3Rh
# bXBpbmcgUlNBNDA5NiBTSEEyNTYgMjAyNSBDQTECEAqA7xhLjfEFgtHEdqeVdGgw
# DQYJYIZIAWUDBAIBBQCgaTAYBgkqhkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqG
# SIb3DQEJBTEPFw0yNTA3MTIwODAwMTZaMC8GCSqGSIb3DQEJBDEiBCCmS2cWPWTK
# xhSie9Huz9mXa2V90X8rcOKgSWbMY0oaAzANBgkqhkiG9w0BAQEFAASCAgAefRLh
# FZP66YxueB1dQUjfiZa8jBYixWDKJLnKW/rVlL5EtBm2OqGh/SummZ1jbqJsVdRt
# eyMX3IOGXQI3DPOGizPNjGiYbFI+PrdU17cN463GsNDruSFu+GF0jbzgR/+a1OxN
# w0u7UzebyYpciwIulHCmThxYgou/LIVE4ZSteRZRBZN3IelJ/77uQ9yFKECS+C0S
# FA1dzEZcjnij43oAzLBOBmXhu7klRAx8A5Ga63NeAPJu+qH65cRLLTzZuW/mekjD
# qaf9cViBSftyybzRf+U3HfPCNmlA1zmtMw6lMJvfs+fRrGGy4wLCe6S8w2L/GYi7
# rvu5bBqKYdo0sf6c5nBzFvobfr4GQDbNyLi61CU0x4tzwDYh5iX7MCo2TvI3rEPb
# JsKRWCvF7hX/IUBkwfN1wH4j9tYUHmo0DVbqIfOzPZRaOn4UR7CBOtkUtfJb8484
# JpIXdzKEQ+q4FddeBb0zv5XEmqfiwyYp1pcUBQ0lko2En84qZYvk6XXJXSjUCMIK
# eOM3S209WudEhrrGgGrUhxZeZ94OsRIfX399BR0YO9++FQ5E2m3T/ZitAMA356GO
# wNcmWaUou3lCJ4osl1Bqw/pqTgj9sJYMtB9piscMxqBEPQ9pfLjXid5gmv7ysPES
# lyvFcvouuD5DrQ6bxEG1Egqd0+Sf6NjnQnSNHg==
# SIG # End signature block
