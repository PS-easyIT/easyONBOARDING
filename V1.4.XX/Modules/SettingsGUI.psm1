# SettingsGUI.psm1 - Modern Settings Interface for easyONBOARDING
# Version: 1.0.0
# Description: Provides a modern, user-friendly settings interface

#region Module Variables
$script:SettingsWindow = $null
$script:ConfigPath = $null
$script:Config = $null
$script:LogFunction = $null
$script:HasChanges = $false
#endregion

#region Initialization
function Initialize-SettingsGUI {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$IniPath,
        
        [Parameter(Mandatory = $false)]
        [ScriptBlock]$LogFunction
    )
    
    $script:ConfigPath = $IniPath
    $script:LogFunction = $LogFunction
    Write-SettingsLog "SettingsGUI module initialized" "INFO"
}

function Write-SettingsLog {
    param(
        [string]$Message,
        [string]$Level = "INFO"
    )
    
    if ($script:LogFunction) {
        & $script:LogFunction $Message $Level
    }
    else {
        Write-Host "[$Level] $Message"
    }
}
#endregion

#region Settings Window
function Show-SettingsWindow {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [System.Windows.Window]$ParentWindow
    )
    
    try {
        Write-SettingsLog "Opening Settings window" "INFO"
        
        # Load configuration
        $script:Config = Get-IniContent -Path $script:ConfigPath
        
        # Create XAML for settings window
        $xaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="easyONBOARDING Settings"
        Height="700" Width="1000"
        WindowStartupLocation="CenterScreen"
        Background="#F5F5F5">
    
    <Window.Resources>
        <!-- Modern Styles -->
        <Style x:Key="CategoryButton" TargetType="Button">
            <Setter Property="Background" Value="Transparent"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Padding" Value="20,15"/>
            <Setter Property="HorizontalAlignment" Value="Stretch"/>
            <Setter Property="HorizontalContentAlignment" Value="Left"/>
            <Setter Property="FontSize" Value="14"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="border" Background="{TemplateBinding Background}" 
                                BorderThickness="0" Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="{TemplateBinding HorizontalContentAlignment}"
                                            VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="border" Property="Background" Value="#E3E3E3"/>
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter TargetName="border" Property="Background" Value="#D0D0D0"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        
        <Style x:Key="SelectedCategoryButton" TargetType="Button" BasedOn="{StaticResource CategoryButton}">
            <Setter Property="Background" Value="#0078D4"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
        </Style>
        
        <Style x:Key="ModernTextBox" TargetType="TextBox">
            <Setter Property="BorderBrush" Value="#D0D0D0"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding" Value="8,6"/>
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="Background" Value="White"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="TextBox">
                        <Border Background="{TemplateBinding Background}"
                                BorderBrush="{TemplateBinding BorderBrush}"
                                BorderThickness="{TemplateBinding BorderThickness}"
                                CornerRadius="4">
                            <ScrollViewer x:Name="PART_ContentHost" Margin="{TemplateBinding Padding}"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsFocused" Value="True">
                                <Setter Property="BorderBrush" Value="#0078D4"/>
                                <Setter Property="BorderThickness" Value="2"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        
        <Style x:Key="SaveButton" TargetType="Button">
            <Setter Property="Background" Value="#0078D4"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Padding" Value="20,10"/>
            <Setter Property="FontSize" Value="14"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}" 
                                CornerRadius="4"
                                Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter Property="Background" Value="#106EBE"/>
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter Property="Background" Value="#005A9E"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
    </Window.Resources>
    
    <Grid>
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="250"/>
            <ColumnDefinition Width="*"/>
        </Grid.ColumnDefinitions>
        
        <!-- Left Navigation -->
        <Border Grid.Column="0" Background="White" BorderBrush="#E0E0E0" BorderThickness="0,0,1,0">
            <ScrollViewer VerticalScrollBarVisibility="Auto">
                <StackPanel Margin="0,20">
                    <TextBlock Text="Settings" FontSize="20" FontWeight="Bold" 
                               Margin="20,0,20,20" Foreground="#1A1A1A"/>
                    
                    <!-- Category Buttons -->
                    <Button Name="btnGeneral" Content="🏢 General" Style="{StaticResource SelectedCategoryButton}"/>
                    <Button Name="btnCompany" Content="🏢 Company Information" Style="{StaticResource CategoryButton}"/>
                    <Button Name="btnActiveDirectory" Content="🌐 Active Directory" Style="{StaticResource CategoryButton}"/>
                    <Button Name="btnEmail" Content="📧 Email Settings" Style="{StaticResource CategoryButton}"/>
                    <Button Name="btnPasswords" Content="🔐 Password Policy" Style="{StaticResource CategoryButton}"/>
                    <Button Name="btnGroups" Content="👥 Groups & Licenses" Style="{StaticResource CategoryButton}"/>
                    <Button Name="btnReporting" Content="📊 Reporting" Style="{StaticResource CategoryButton}"/>
                    <Button Name="btnAdvanced" Content="⚙️ Advanced" Style="{StaticResource CategoryButton}"/>
                </StackPanel>
            </ScrollViewer>
        </Border>
        
        <!-- Right Content Area -->
        <Grid Grid.Column="1">
            <Grid.RowDefinitions>
                <RowDefinition Height="*"/>
                <RowDefinition Height="Auto"/>
            </Grid.RowDefinitions>
            
            <!-- Content ScrollViewer -->
            <ScrollViewer Grid.Row="0" VerticalScrollBarVisibility="Auto" Padding="30">
                <StackPanel Name="ContentPanel">
                    <!-- General Settings Panel (Default) -->
                    <StackPanel Name="GeneralPanel" Visibility="Visible">
                        <TextBlock Text="General Settings" FontSize="24" FontWeight="SemiBold" 
                                   Margin="0,0,0,20" Foreground="#1A1A1A"/>
                        
                        <Border Background="White" CornerRadius="8" Padding="20" Margin="0,0,0,20">
                            <Border.Effect>
                                <DropShadowEffect BlurRadius="10" Opacity="0.1" ShadowDepth="2"/>
                            </Border.Effect>
                            
                            <StackPanel>
                                <TextBlock Text="Application Settings" FontSize="16" FontWeight="SemiBold" 
                                           Margin="0,0,0,15" Foreground="#1A1A1A"/>
                                
                                <Grid>
                                    <Grid.ColumnDefinitions>
                                        <ColumnDefinition Width="200"/>
                                        <ColumnDefinition Width="*"/>
                                    </Grid.ColumnDefinitions>
                                    <Grid.RowDefinitions>
                                        <RowDefinition Height="Auto"/>
                                        <RowDefinition Height="Auto"/>
                                        <RowDefinition Height="Auto"/>
                                    </Grid.RowDefinitions>
                                    
                                    <TextBlock Grid.Row="0" Grid.Column="0" Text="Application Name:" 
                                               VerticalAlignment="Center" Margin="0,0,0,10"/>
                                    <TextBox Grid.Row="0" Grid.Column="1" Name="txtAppName" 
                                             Style="{StaticResource ModernTextBox}" Margin="0,0,0,10"/>
                                    
                                    <TextBlock Grid.Row="1" Grid.Column="0" Text="Theme Color:" 
                                               VerticalAlignment="Center" Margin="0,0,0,10"/>
                                    <ComboBox Grid.Row="1" Grid.Column="1" Name="cmbThemeColor" 
                                              Margin="0,0,0,10">
                                        <ComboBoxItem Content="Blue (Default)" Tag="#0078D4"/>
                                        <ComboBoxItem Content="Green" Tag="#107C10"/>
                                        <ComboBoxItem Content="Red" Tag="#D13438"/>
                                        <ComboBoxItem Content="Purple" Tag="#5C2D91"/>
                                    </ComboBox>
                                    
                                    <TextBlock Grid.Row="2" Grid.Column="0" Text="Debug Mode:" 
                                               VerticalAlignment="Center"/>
                                    <CheckBox Grid.Row="2" Grid.Column="1" Name="chkDebugMode" 
                                              Content="Enable debug logging" VerticalAlignment="Center"/>
                                </Grid>
                            </StackPanel>
                        </Border>
                    </StackPanel>
                    
                    <!-- Company Settings Panel -->
                    <StackPanel Name="CompanyPanel" Visibility="Collapsed">
                        <TextBlock Text="Company Information" FontSize="24" FontWeight="SemiBold" 
                                   Margin="0,0,0,20" Foreground="#1A1A1A"/>
                        
                        <Border Background="White" CornerRadius="8" Padding="20" Margin="0,0,0,20">
                            <Border.Effect>
                                <DropShadowEffect BlurRadius="10" Opacity="0.1" ShadowDepth="2"/>
                            </Border.Effect>
                            
                            <StackPanel>
                                <TextBlock Text="Company Details" FontSize="16" FontWeight="SemiBold" 
                                           Margin="0,0,0,15" Foreground="#1A1A1A"/>
                                
                                <Grid>
                                    <Grid.ColumnDefinitions>
                                        <ColumnDefinition Width="200"/>
                                        <ColumnDefinition Width="*"/>
                                    </Grid.ColumnDefinitions>
                                    <Grid.RowDefinitions>
                                        <RowDefinition Height="Auto"/>
                                        <RowDefinition Height="Auto"/>
                                        <RowDefinition Height="Auto"/>
                                        <RowDefinition Height="Auto"/>
                                        <RowDefinition Height="Auto"/>
                                        <RowDefinition Height="Auto"/>
                                    </Grid.RowDefinitions>
                                    
                                    <TextBlock Grid.Row="0" Grid.Column="0" Text="Company Name:" 
                                               VerticalAlignment="Center" Margin="0,0,0,10"/>
                                    <TextBox Grid.Row="0" Grid.Column="1" Name="txtCompanyName" 
                                             Style="{StaticResource ModernTextBox}" Margin="0,0,0,10"/>
                                    
                                    <TextBlock Grid.Row="1" Grid.Column="0" Text="Short Name:" 
                                               VerticalAlignment="Center" Margin="0,0,0,10"/>
                                    <TextBox Grid.Row="1" Grid.Column="1" Name="txtCompanyShortName" 
                                             Style="{StaticResource ModernTextBox}" Margin="0,0,0,10"/>
                                    
                                    <TextBlock Grid.Row="2" Grid.Column="0" Text="Street Address:" 
                                               VerticalAlignment="Center" Margin="0,0,0,10"/>
                                    <TextBox Grid.Row="2" Grid.Column="1" Name="txtCompanyStreet" 
                                             Style="{StaticResource ModernTextBox}" Margin="0,0,0,10"/>
                                    
                                    <TextBlock Grid.Row="3" Grid.Column="0" Text="Postal Code:" 
                                               VerticalAlignment="Center" Margin="0,0,0,10"/>
                                    <TextBox Grid.Row="3" Grid.Column="1" Name="txtCompanyZIP" 
                                             Style="{StaticResource ModernTextBox}" Margin="0,0,0,10"/>
                                    
                                    <TextBlock Grid.Row="4" Grid.Column="0" Text="City:" 
                                               VerticalAlignment="Center" Margin="0,0,0,10"/>
                                    <TextBox Grid.Row="4" Grid.Column="1" Name="txtCompanyCity" 
                                             Style="{StaticResource ModernTextBox}" Margin="0,0,0,10"/>
                                    
                                    <TextBlock Grid.Row="5" Grid.Column="0" Text="Country:" 
                                               VerticalAlignment="Center"/>
                                    <TextBox Grid.Row="5" Grid.Column="1" Name="txtCompanyCountry" 
                                             Style="{StaticResource ModernTextBox}"/>
                                </Grid>
                            </StackPanel>
                        </Border>
                    </StackPanel>
                    
                    <!-- More panels would be added here for other categories -->
                </StackPanel>
            </ScrollViewer>
            
            <!-- Bottom Action Bar -->
            <Border Grid.Row="1" Background="White" BorderBrush="#E0E0E0" BorderThickness="0,1,0,0">
                <Grid Margin="30,15">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="Auto"/>
                        <ColumnDefinition Width="Auto"/>
                    </Grid.ColumnDefinitions>
                    
                    <TextBlock Grid.Column="0" Name="txtStatus" Text="Ready" 
                               VerticalAlignment="Center" Foreground="#666666"/>
                    
                    <Button Grid.Column="1" Name="btnCancel" Content="Cancel" 
                            Margin="0,0,10,0" Padding="20,8" Background="#E0E0E0"
                            BorderThickness="0" Cursor="Hand"/>
                    
                    <Button Grid.Column="2" Name="btnSave" Content="Save Changes" 
                            Style="{StaticResource SaveButton}"/>
                </Grid>
            </Border>
        </Grid>
    </Grid>
</Window>
'@
        
        # Load XAML
        $reader = [System.Xml.XmlReader]::Create([System.IO.StringReader]$xaml)
        $script:SettingsWindow = [Windows.Markup.XamlReader]::Load($reader)
        
        # Get controls
        $btnGeneral = $script:SettingsWindow.FindName("btnGeneral")
        $btnCompany = $script:SettingsWindow.FindName("btnCompany")
        $btnActiveDirectory = $script:SettingsWindow.FindName("btnActiveDirectory")
        $btnEmail = $script:SettingsWindow.FindName("btnEmail")
        $btnPasswords = $script:SettingsWindow.FindName("btnPasswords")
        $btnGroups = $script:SettingsWindow.FindName("btnGroups")
        $btnReporting = $script:SettingsWindow.FindName("btnReporting")
        $btnAdvanced = $script:SettingsWindow.FindName("btnAdvanced")
        
        $btnSave = $script:SettingsWindow.FindName("btnSave")
        $btnCancel = $script:SettingsWindow.FindName("btnCancel")
        
        # Set up navigation
        $categoryButtons = @($btnGeneral, $btnCompany, $btnActiveDirectory, $btnEmail, 
                           $btnPasswords, $btnGroups, $btnReporting, $btnAdvanced)
        
        foreach ($btn in $categoryButtons) {
            $btn.Add_Click({
                param($sender, $e)
                
                # Update button styles
                foreach ($b in $categoryButtons) {
                    $b.Style = $script:SettingsWindow.FindResource("CategoryButton")
                }
                $sender.Style = $script:SettingsWindow.FindResource("SelectedCategoryButton")
                
                # Show appropriate panel
                Show-SettingsPanel -PanelName $sender.Name.Replace("btn", "")
            })
        }
        
        # Load current settings
        Load-SettingsToUI
        
        # Save button
        $btnSave.Add_Click({
            Save-Settings
        })
        
        # Cancel button
        $btnCancel.Add_Click({
            if ($script:HasChanges) {
                $result = [System.Windows.MessageBox]::Show(
                    "You have unsaved changes. Do you want to discard them?",
                    "Unsaved Changes",
                    [System.Windows.MessageBoxButton]::YesNo,
                    [System.Windows.MessageBoxImage]::Warning
                )
                
                if ($result -eq [System.Windows.MessageBoxResult]::Yes) {
                    $script:SettingsWindow.Close()
                }
            }
            else {
                $script:SettingsWindow.Close()
            }
        })
        
        # Show window
        if ($ParentWindow) {
            $script:SettingsWindow.Owner = $ParentWindow
        }
        
        $script:SettingsWindow.ShowDialog()
    }
    catch {
        Write-SettingsLog "Error showing settings window: $($_.Exception.Message)" "ERROR"
        throw
    }
}

function Show-SettingsPanel {
    param([string]$PanelName)
    
    $contentPanel = $script:SettingsWindow.FindName("ContentPanel")
    
    # Hide all panels
    foreach ($child in $contentPanel.Children) {
        $child.Visibility = [System.Windows.Visibility]::Collapsed
    }
    
    # Show selected panel
    $panel = $script:SettingsWindow.FindName("${PanelName}Panel")
    if ($panel) {
        $panel.Visibility = [System.Windows.Visibility]::Visible
    }
}

function Load-SettingsToUI {
    try {
        # General settings
        $txtAppName = $script:SettingsWindow.FindName("txtAppName")
        if ($txtAppName -and $script:Config.WPFGUI.APPName) {
            $txtAppName.Text = $script:Config.WPFGUI.APPName
        }
        
        $chkDebugMode = $script:SettingsWindow.FindName("chkDebugMode")
        if ($chkDebugMode -and $script:Config.Logging.DebugMode) {
            $chkDebugMode.IsChecked = $script:Config.Logging.DebugMode -eq "1"
        }
        
        # Company settings
        $txtCompanyName = $script:SettingsWindow.FindName("txtCompanyName")
        if ($txtCompanyName -and $script:Config.Company.CompanyNameFirma) {
            $txtCompanyName.Text = $script:Config.Company.CompanyNameFirma
        }
        
        # Add more field mappings as needed
        
        # Monitor changes
        Monitor-UIChanges
    }
    catch {
        Write-SettingsLog "Error loading settings: $($_.Exception.Message)" "ERROR"
    }
}

function Monitor-UIChanges {
    # Add event handlers to detect changes
    $textBoxes = @("txtAppName", "txtCompanyName", "txtCompanyShortName", 
                   "txtCompanyStreet", "txtCompanyZIP", "txtCompanyCity", "txtCompanyCountry")
    
    foreach ($tbName in $textBoxes) {
        $tb = $script:SettingsWindow.FindName($tbName)
        if ($tb) {
            $tb.Add_TextChanged({
                $script:HasChanges = $true
                Update-StatusText "Unsaved changes"
            })
        }
    }
    
    $checkBoxes = @("chkDebugMode")
    foreach ($cbName in $checkBoxes) {
        $cb = $script:SettingsWindow.FindName($cbName)
        if ($cb) {
            $cb.Add_Checked({ 
                $script:HasChanges = $true
                Update-StatusText "Unsaved changes"
            })
            $cb.Add_Unchecked({ 
                $script:HasChanges = $true
                Update-StatusText "Unsaved changes"
            })
        }
    }
}

function Update-StatusText {
    param([string]$Text)
    
    $txtStatus = $script:SettingsWindow.FindName("txtStatus")
    if ($txtStatus) {
        $txtStatus.Text = $Text
    }
}

function Save-Settings {
    try {
        Update-StatusText "Saving..."
        
        # Update configuration from UI
        $txtAppName = $script:SettingsWindow.FindName("txtAppName")
        if ($txtAppName) {
            $script:Config.WPFGUI.APPName = $txtAppName.Text
        }
        
        $chkDebugMode = $script:SettingsWindow.FindName("chkDebugMode")
        if ($chkDebugMode) {
            $script:Config.Logging.DebugMode = if ($chkDebugMode.IsChecked) { "1" } else { "0" }
        }
        
        # Company settings
        $txtCompanyName = $script:SettingsWindow.FindName("txtCompanyName")
        if ($txtCompanyName) {
            $script:Config.Company.CompanyNameFirma = $txtCompanyName.Text
        }
        
        # Save to INI file
        Save-IniContent -Path $script:ConfigPath -Content $script:Config
        
        $script:HasChanges = $false
        Update-StatusText "Settings saved successfully"
        
        [System.Windows.MessageBox]::Show(
            "Settings have been saved successfully.",
            "Success",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Information
        )
    }
    catch {
        Write-SettingsLog "Error saving settings: $($_.Exception.Message)" "ERROR"
        Update-StatusText "Error saving settings"
        
        [System.Windows.MessageBox]::Show(
            "Error saving settings: $($_.Exception.Message)",
            "Error",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Error
        )
    }
}

function Save-IniContent {
    param(
        [string]$Path,
        [hashtable]$Content
    )
    
    $output = @()
    
    foreach ($section in $Content.Keys) {
        $output += "[$section]"
        
        foreach ($key in $Content[$section].Keys) {
            $value = $Content[$section][$key]
            $output += "$key=$value"
        }
        
        $output += ""  # Empty line between sections
    }
    
    $output | Out-File -FilePath $Path -Encoding UTF8
}

function Get-IniContent {
    param([string]$Path)
    
    $ini = [ordered]@{}
    $currentSection = ""
    
    $lines = Get-Content -Path $Path
    
    foreach ($line in $lines) {
        $line = $line.Trim()
        
        if ($line -match '^\[(.+)\]$') {
            $currentSection = $matches[1]
            if (-not $ini.Contains($currentSection)) {
                $ini[$currentSection] = [ordered]@{}
            }
        }
        elseif ($line -match '^(.+?)=(.*)$') {
            if ($currentSection) {
                $key = $matches[1].Trim()
                $value = $matches[2].Trim()
                $ini[$currentSection][$key] = $value
            }
        }
    }
    
    return $ini
}
#endregion

# Export module members
Export-ModuleMember -Function @(
    'Initialize-SettingsGUI',
    'Show-SettingsWindow'
) 
# SIG # Begin signature block
# MIIcCAYJKoZIhvcNAQcCoIIb+TCCG/UCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDXgkex9PVmHtQO
# oI8vjWGaar5n6L2pznMxnRRbvzMcmaCCFk4wggMQMIIB+KADAgECAhB3jzsyX9Cg
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
# BgkqhkiG9w0BCQQxIgQgLqYThrHsPmtIi+hmuEwPFPsqWqSbk0CvkI/PYx6buJgw
# DQYJKoZIhvcNAQEBBQAEggEAVxrynnu9fLbZ4fu5inSAXxAMfnsWW921Sysg1T4E
# /qJzzl8fsxqQwSWpKZX3a+S5vdKMU/OMyUT9fZ6MfcwlxB82pybCO70X9cCFQTT/
# foLbtI9Ebt4EOLVyltNT8CpIstU2/FEbw6HkviXCUULUXLIuPNuxdTkUV2/knAu3
# hnc1+Mwribzz0jT2D+Q2gKA5a/W2m5N701WA2BhjiO2Qw+Vx41jMGWqtWug/VMPg
# gHTDtRCD39jZiZrUrnKM3hjIGEUpB4aE4k08/T6sAJ5gKMFLniLUUbHYJuLl9EMn
# sgWSm5EQXWxxaqBV4XaIzZ8GMHslO0sv/KTmnz4mAXyUwaGCAyYwggMiBgkqhkiG
# 9w0BCQYxggMTMIIDDwIBATB9MGkxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdp
# Q2VydCwgSW5jLjFBMD8GA1UEAxM4RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3Rh
# bXBpbmcgUlNBNDA5NiBTSEEyNTYgMjAyNSBDQTECEAqA7xhLjfEFgtHEdqeVdGgw
# DQYJYIZIAWUDBAIBBQCgaTAYBgkqhkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqG
# SIb3DQEJBTEPFw0yNTA3MTExOTMyMTRaMC8GCSqGSIb3DQEJBDEiBCD/TW83vLJR
# oxsJ534D4mV+y6YZx0L8HnSqAMPHHe9XxzANBgkqhkiG9w0BAQEFAASCAgB/U4td
# EvPqYXngb6rqX+DODoLyKhpUADIJbzn9rLgsD4D7J10tbAJEpE7+ReFPBVe4mhMa
# iSidiHyr9q6+JJZHRzx0x+VEkHW0Fn8KkVKfG6QUlZxds+Jcw7Wd0GjCk/qBO8Je
# cBGKKnk1sPhLdrqzJ16GGfTgIQYtakuTpyx0QM5mQwDmIC3ugeGeQW1djvoZrKjF
# DtyylnfjcADvxWq/jMgffjE6eosOnDgzq+20sSoxbxz4eZoAqBH44LzOTeZNbcZP
# ffhf2t0bNjYwaK5Sa2II2O72I8F2bCTqj1veWsZdWJNfkEPUpUGL1lAfcwwtHoe0
# AJeIeApuCgcM+5DyY0XhNlf4sHrdKv5TONfXlTPvw4sT1bbNdSKEfhSxQM4O6Xpa
# 4PVowDFTYq0gEUTheYsIhpODkbgLIayL4oEDD71Jg4ePpCFzWqYH9+pk7wJqBnAM
# W0FVE/L1vHr+pCm+UN/nLo/23RvgI4aPmwGkQ02KVKWdg6vR5BLuvbyL7mdZBbCx
# Eb6lZ57dn2k4fY0r+CoosyJTbHoZY/mMNbLNxICqzf6ZlmAtAAs9/Qc2utz38vS/
# WrJ1NIU7ZlRo0yLKZGXMQ1XUg9MkkMbxLAWadwD53WA57mPyzDcuWaNYAIkjFcRv
# Kd8x5G3W5u6OxVsr7abxhIVHPVsZ5Qg8WD9mOw==
# SIG # End signature block
