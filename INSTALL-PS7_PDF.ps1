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
