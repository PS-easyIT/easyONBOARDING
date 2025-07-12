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
# MIIbywYJKoZIhvcNAQcCoIIbvDCCG7gCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAPawz1mnItD1eC
# oqRLvdrUziQSCFdG+z0lRfLS1fwWc6CCFhcwggMQMIIB+KADAgECAhB3jzsyX9Cg
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
# DQEJBDEiBCBY7LsSAQuGitazznw/vQjQwwYN6L9mnwLf4rqoGbq+ojANBgkqhkiG
# 9w0BAQEFAASCAQA3RT951UuM86fOF7ih9ZDz8rYmO1Bg0CYHnf5YrlgPtpQqz0/j
# dnm3Ay599nyZaXZOkQq39Q+qRgFQZmb0nWCqgLsUVItvArbSRKuOsVyyKs36Q3Fh
# o4+cgefYOUWkgVA92QfdfT3l/sxLl4GtUFg3uMKLUR9atgElXO5qkYCcbT5Hv9v7
# YRiuxFawL4jnrVxlR3121C1lWfMZQnxMm9em3vN6Z8EIRTw258SrxnhJPCbujKk7
# dM2286oci4CJpQRE7g9falYUm1XITYqGLzodFJoAnbwNTZNNsQdFYKvPoaY6VMoB
# gfZcvrDFN0d0jPJlcqGzfqkJKp845MIry4FnoYIDIDCCAxwGCSqGSIb3DQEJBjGC
# Aw0wggMJAgEBMHcwYzELMAkGA1UEBhMCVVMxFzAVBgNVBAoTDkRpZ2lDZXJ0LCBJ
# bmMuMTswOQYDVQQDEzJEaWdpQ2VydCBUcnVzdGVkIEc0IFJTQTQwOTYgU0hBMjU2
# IFRpbWVTdGFtcGluZyBDQQIQC65mvFq6f5WHxvnpBOMzBDANBglghkgBZQMEAgEF
# AKBpMBgGCSqGSIb3DQEJAzELBgkqhkiG9w0BBwEwHAYJKoZIhvcNAQkFMQ8XDTI1
# MDcwODE1MDk1NlowLwYJKoZIhvcNAQkEMSIEIOF6ZoIrwhwlWKM6gLNd+RtMoE/3
# dHLU3PV3SM1qHHw2MA0GCSqGSIb3DQEBAQUABIICAA0c0WmDzEYdQ3y8VU+hR34O
# UYmaC5w4/gPiiG+U4c+hcxq4X1Zo0Fhv6AcT7jA6qfezrtxZgkl1ojhVicTTbZcf
# sslwmVW6EP2yqPsgEjt0sNx+uaJTtkDFfcOaSj44Dz130w/GMLFIqqxcXPR58Wll
# SAiOyfm1l6uQ3zPBzTB9Nu1MUdgTjFa3M8pjDtf/FVPpBi5FDX0PI65MnkTy59bc
# 8+vDAftm1gYe2LjnFWatVyzV08+fEFxMwjlxi95czAfVLyt3E7eYqq8Wu8daXtMH
# moH6xJCDBkuq3LUYiYee9nutUfxW3mlseHWYVKGbcIEC3O4xzauNpX7IQc4QlerF
# f/y9G1llNXU3Pyx0NKFiq/Xee13kpacUOYLC2IGSWebJQ9iwiaowtj5RvczJFIKA
# JnPaX3vdlDzIe6i9OZCrnbKIFatv1an9VS5Uol57SgB44dLQxRefEuB98ki/hx+3
# 51n/FHU5Pa3jcDo1mbvkgXn9cwpTQQrjqeHXp96jFSoOqHHzE1rlhrbFpPBI2rL4
# d3YkWIE0tk2nch/h+mnCz99FxC6uvQa4UuwErQqE0hXou6ENIJgf3SKSBUIqqvRf
# ZOksqDtGOmslEjpoB9v2jJEXipGPIpD6oPVPCd7bPyhvesN2bTMPv08hMmdAFrYH
# 1fJPln+Cfp1/XU6m/R23
# SIG # End signature block
