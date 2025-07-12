# SettingsGUI2.psm1 - Additional Settings Panels
# This file extends SettingsGUI.psm1 with more setting panels

function Add-ActiveDirectoryPanel {
    param($ContentPanel)
    
    $xaml = @'
    <!-- Active Directory Settings Panel -->
    <StackPanel Name="ActiveDirectoryPanel" Visibility="Collapsed">
        <TextBlock Text="Active Directory Settings" FontSize="24" FontWeight="SemiBold" 
                   Margin="0,0,0,20" Foreground="#1A1A1A"/>
        
        <Border Background="White" CornerRadius="8" Padding="20" Margin="0,0,0,20">
            <Border.Effect>
                <DropShadowEffect BlurRadius="10" Opacity="0.1" ShadowDepth="2"/>
            </Border.Effect>
            
            <StackPanel>
                <TextBlock Text="Domain Configuration" FontSize="16" FontWeight="SemiBold" 
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
                    </Grid.RowDefinitions>
                    
                    <TextBlock Grid.Row="0" Grid.Column="0" Text="Domain:" 
                               VerticalAlignment="Center" Margin="0,0,0,10"/>
                    <TextBox Grid.Row="0" Grid.Column="1" Name="txtADDomain" 
                             Style="{StaticResource ModernTextBox}" Margin="0,0,0,10"
                             ToolTip="e.g., company.local"/>
                    
                    <TextBlock Grid.Row="1" Grid.Column="0" Text="Default OU:" 
                               VerticalAlignment="Center" Margin="0,0,0,10"/>
                    <Grid Grid.Row="1" Grid.Column="1" Margin="0,0,0,10">
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="Auto"/>
                        </Grid.ColumnDefinitions>
                        <TextBox Grid.Column="0" Name="txtDefaultOU" 
                                 Style="{StaticResource ModernTextBox}"
                                 ToolTip="Distinguished Name of default OU"/>
                        <Button Grid.Column="1" Name="btnBrowseOU" Content="Browse..." 
                                Margin="5,0,0,0" Padding="10,5"/>
                    </Grid>
                    
                    <TextBlock Grid.Row="2" Grid.Column="0" Text="Preferred DC:" 
                               VerticalAlignment="Center" Margin="0,0,0,10"/>
                    <TextBox Grid.Row="2" Grid.Column="1" Name="txtPreferredDC" 
                             Style="{StaticResource ModernTextBox}" Margin="0,0,0,10"
                             ToolTip="Leave empty for automatic selection"/>
                    
                    <TextBlock Grid.Row="3" Grid.Column="0" Text="AD Sync Server:" 
                               VerticalAlignment="Center"/>
                    <Grid Grid.Row="3" Grid.Column="1">
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="Auto"/>
                        </Grid.ColumnDefinitions>
                        <TextBox Grid.Column="0" Name="txtADSyncServer" 
                                 Style="{StaticResource ModernTextBox}"
                                 ToolTip="Server name for AD synchronization"/>
                        <Button Grid.Column="1" Name="btnTestADSync" Content="Test" 
                                Margin="5,0,0,0" Padding="10,5"/>
                    </Grid>
                </Grid>
            </StackPanel>
        </Border>
        
        <Border Background="White" CornerRadius="8" Padding="20">
            <Border.Effect>
                <DropShadowEffect BlurRadius="10" Opacity="0.1" ShadowDepth="2"/>
            </Border.Effect>
            
            <StackPanel>
                <TextBlock Text="User Defaults" FontSize="16" FontWeight="SemiBold" 
                           Margin="0,0,0,15" Foreground="#1A1A1A"/>
                
                <CheckBox Name="chkMustChangePassword" Content="User must change password at next logon" 
                          Margin="0,0,0,10"/>
                <CheckBox Name="chkPasswordNeverExpires" Content="Password never expires" 
                          Margin="0,0,0,10"/>
                <CheckBox Name="chkAccountDisabledDefault" Content="Create accounts as disabled by default" 
                          Margin="0,0,0,10"/>
                
                <TextBlock Text="Home Directory:" Margin="0,10,0,5"/>
                <Grid>
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="100"/>
                        <ColumnDefinition Width="*"/>
                    </Grid.ColumnDefinitions>
                    
                    <TextBlock Grid.Column="0" Text="Path:" VerticalAlignment="Center"/>
                    <TextBox Grid.Column="1" Name="txtHomeDirectory" 
                             Style="{StaticResource ModernTextBox}"
                             ToolTip="Use %username% as placeholder"/>
                </Grid>
            </StackPanel>
        </Border>
    </StackPanel>
'@
    
    return $xaml
}

function Add-EmailSettingsPanel {
    param($ContentPanel)
    
    $xaml = @'
    <!-- Email Settings Panel -->
    <StackPanel Name="EmailPanel" Visibility="Collapsed">
        <TextBlock Text="Email Settings" FontSize="24" FontWeight="SemiBold" 
                   Margin="0,0,0,20" Foreground="#1A1A1A"/>
        
        <Border Background="White" CornerRadius="8" Padding="20" Margin="0,0,0,20">
            <Border.Effect>
                <DropShadowEffect BlurRadius="10" Opacity="0.1" ShadowDepth="2"/>
            </Border.Effect>
            
            <StackPanel>
                <TextBlock Text="SMTP Configuration" FontSize="16" FontWeight="SemiBold" 
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
                    </Grid.RowDefinitions>
                    
                    <TextBlock Grid.Row="0" Grid.Column="0" Text="SMTP Server:" 
                               VerticalAlignment="Center" Margin="0,0,0,10"/>
                    <TextBox Grid.Row="0" Grid.Column="1" Name="txtSMTPServer" 
                             Style="{StaticResource ModernTextBox}" Margin="0,0,0,10"/>
                    
                    <TextBlock Grid.Row="1" Grid.Column="0" Text="Port:" 
                               VerticalAlignment="Center" Margin="0,0,0,10"/>
                    <TextBox Grid.Row="1" Grid.Column="1" Name="txtSMTPPort" 
                             Style="{StaticResource ModernTextBox}" Margin="0,0,0,10"
                             Text="25"/>
                    
                    <TextBlock Grid.Row="2" Grid.Column="0" Text="From Address:" 
                               VerticalAlignment="Center" Margin="0,0,0,10"/>
                    <TextBox Grid.Row="2" Grid.Column="1" Name="txtFromAddress" 
                             Style="{StaticResource ModernTextBox}" Margin="0,0,0,10"/>
                    
                    <CheckBox Grid.Row="3" Grid.Column="1" Name="chkUseSSL" 
                              Content="Use SSL/TLS" Margin="0,0,0,10"/>
                    
                    <CheckBox Grid.Row="4" Grid.Column="1" Name="chkSendWelcomeEmail" 
                              Content="Send welcome email to new users (optional)" 
                              FontWeight="SemiBold"/>
                </Grid>
            </StackPanel>
        </Border>
        
        <Border Background="White" CornerRadius="8" Padding="20">
            <Border.Effect>
                <DropShadowEffect BlurRadius="10" Opacity="0.1" ShadowDepth="2"/>
            </Border.Effect>
            
            <StackPanel>
                <TextBlock Text="Email Domains" FontSize="16" FontWeight="SemiBold" 
                           Margin="0,0,0,15" Foreground="#1A1A1A"/>
                
                <TextBlock Text="Available email domains (one per line):" Margin="0,0,0,5"/>
                <TextBox Name="txtEmailDomains" Height="100" 
                         Style="{StaticResource ModernTextBox}"
                         AcceptsReturn="True" TextWrapping="Wrap"
                         VerticalScrollBarVisibility="Auto"/>
                
                <CheckBox Name="chkSetProxyAddresses" 
                          Content="Automatically set proxy addresses for Exchange" 
                          Margin="0,10,0,0"/>
            </StackPanel>
        </Border>
    </StackPanel>
'@
    
    return $xaml
}

function Add-PasswordPolicyPanel {
    param($ContentPanel)
    
    $xaml = @'
    <!-- Password Policy Panel -->
    <StackPanel Name="PasswordsPanel" Visibility="Collapsed">
        <TextBlock Text="Password Policy" FontSize="24" FontWeight="SemiBold" 
                   Margin="0,0,0,20" Foreground="#1A1A1A"/>
        
        <Border Background="White" CornerRadius="8" Padding="20" Margin="0,0,0,20">
            <Border.Effect>
                <DropShadowEffect BlurRadius="10" Opacity="0.1" ShadowDepth="2"/>
            </Border.Effect>
            
            <StackPanel>
                <TextBlock Text="Password Generation" FontSize="16" FontWeight="SemiBold" 
                           Margin="0,0,0,15" Foreground="#1A1A1A"/>
                
                <RadioButton Name="radGeneratePassword" Content="Generate secure passwords" 
                             IsChecked="True" Margin="0,0,0,10" GroupName="PasswordMode"/>
                <RadioButton Name="radFixedPassword" Content="Use fixed password" 
                             Margin="0,0,0,10" GroupName="PasswordMode"/>
                
                <Border BorderBrush="#E0E0E0" BorderThickness="1" CornerRadius="4" 
                        Padding="15" Margin="0,10,0,0">
                    <StackPanel Name="pnlPasswordGeneration">
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
                            </Grid.RowDefinitions>
                            
                            <TextBlock Grid.Row="0" Grid.Column="0" Text="Password Length:" 
                                       VerticalAlignment="Center" Margin="0,0,0,10"/>
                            <StackPanel Grid.Row="0" Grid.Column="1" Orientation="Horizontal" Margin="0,0,0,10">
                                <Slider Name="sldPasswordLength" Minimum="8" Maximum="32" 
                                        Value="12" Width="200" VerticalAlignment="Center"/>
                                <TextBlock Name="txtPasswordLengthValue" Text="12" 
                                           Margin="10,0,0,0" VerticalAlignment="Center" FontWeight="Bold"/>
                            </StackPanel>
                            
                            <TextBlock Grid.Row="1" Grid.Column="0" Text="Minimum Uppercase:" 
                                       VerticalAlignment="Center" Margin="0,0,0,10"/>
                            <TextBox Grid.Row="1" Grid.Column="1" Name="txtMinUppercase" 
                                     Style="{StaticResource ModernTextBox}" Width="50" 
                                     HorizontalAlignment="Left" Margin="0,0,0,10" Text="2"/>
                            
                            <TextBlock Grid.Row="2" Grid.Column="0" Text="Minimum Digits:" 
                                       VerticalAlignment="Center" Margin="0,0,0,10"/>
                            <TextBox Grid.Row="2" Grid.Column="1" Name="txtMinDigits" 
                                     Style="{StaticResource ModernTextBox}" Width="50" 
                                     HorizontalAlignment="Left" Margin="0,0,0,10" Text="2"/>
                            
                            <CheckBox Grid.Row="3" Grid.Column="1" Name="chkIncludeSpecial" 
                                      Content="Include special characters" IsChecked="True" 
                                      Margin="0,0,0,10"/>
                            
                            <CheckBox Grid.Row="4" Grid.Column="1" Name="chkAvoidAmbiguous" 
                                      Content="Avoid ambiguous characters (0, O, l, I)" 
                                      IsChecked="True"/>
                        </Grid>
                    </StackPanel>
                    
                    <StackPanel Name="pnlFixedPassword" Visibility="Collapsed">
                        <TextBlock Text="Fixed Password:" Margin="0,0,0,5"/>
                        <PasswordBox Name="pwdFixedPassword" 
                                     Style="{StaticResource ModernTextBox}"/>
                        <TextBlock Text="Warning: Using a fixed password is not recommended for security reasons." 
                                   Foreground="#D13438" TextWrapping="Wrap" Margin="0,5,0,0"/>
                    </StackPanel>
                </Border>
            </StackPanel>
        </Border>
    </StackPanel>
'@
    
    return $xaml
}

function Add-GroupsAndLicensesPanel {
    param($ContentPanel)
    
    $xaml = @'
    <!-- Groups and Licenses Panel -->
    <StackPanel Name="GroupsPanel" Visibility="Collapsed">
        <TextBlock Text="Groups & Licenses" FontSize="24" FontWeight="SemiBold" 
                   Margin="0,0,0,20" Foreground="#1A1A1A"/>
        
        <Border Background="White" CornerRadius="8" Padding="20" Margin="0,0,0,20">
            <Border.Effect>
                <DropShadowEffect BlurRadius="10" Opacity="0.1" ShadowDepth="2"/>
            </Border.Effect>
            
            <StackPanel>
                <TextBlock Text="AD Groups" FontSize="16" FontWeight="SemiBold" 
                           Margin="0,0,0,15" Foreground="#1A1A1A"/>
                
                <Grid>
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="Auto"/>
                    </Grid.ColumnDefinitions>
                    
                    <DataGrid Grid.Column="0" Name="dgADGroups" Height="200" 
                              AutoGenerateColumns="False" CanUserAddRows="True"
                              CanUserDeleteRows="True">
                        <DataGrid.Columns>
                            <DataGridTextColumn Header="Display Name" Binding="{Binding DisplayName}" Width="200"/>
                            <DataGridTextColumn Header="AD Group Name" Binding="{Binding GroupName}" Width="*"/>
                        </DataGrid.Columns>
                    </DataGrid>
                    
                    <StackPanel Grid.Column="1" Margin="10,0,0,0">
                        <Button Name="btnAddGroup" Content="Add" Width="80" Margin="0,0,0,5"/>
                        <Button Name="btnRemoveGroup" Content="Remove" Width="80" Margin="0,0,0,5"/>
                        <Button Name="btnImportGroups" Content="Import..." Width="80"/>
                    </StackPanel>
                </Grid>
            </StackPanel>
        </Border>
        
        <Border Background="White" CornerRadius="8" Padding="20">
            <Border.Effect>
                <DropShadowEffect BlurRadius="10" Opacity="0.1" ShadowDepth="2"/>
            </Border.Effect>
            
            <StackPanel>
                <TextBlock Text="License Groups" FontSize="16" FontWeight="SemiBold" 
                           Margin="0,0,0,15" Foreground="#1A1A1A"/>
                
                <DataGrid Name="dgLicenses" Height="150" 
                          AutoGenerateColumns="False" CanUserAddRows="True"
                          CanUserDeleteRows="True">
                    <DataGrid.Columns>
                        <DataGridTextColumn Header="License Name" Binding="{Binding LicenseName}" Width="200"/>
                        <DataGridTextColumn Header="AD Group" Binding="{Binding ADGroup}" Width="*"/>
                    </DataGrid.Columns>
                </DataGrid>
            </StackPanel>
        </Border>
    </StackPanel>
'@
    
    return $xaml
}

function Add-AdvancedSettingsPanel {
    param($ContentPanel)
    
    $xaml = @'
    <!-- Advanced Settings Panel -->
    <StackPanel Name="AdvancedPanel" Visibility="Collapsed">
        <TextBlock Text="Advanced Settings" FontSize="24" FontWeight="SemiBold" 
                   Margin="0,0,0,20" Foreground="#1A1A1A"/>
        
        <Border Background="White" CornerRadius="8" Padding="20" Margin="0,0,0,20">
            <Border.Effect>
                <DropShadowEffect BlurRadius="10" Opacity="0.1" ShadowDepth="2"/>
            </Border.Effect>
            
            <StackPanel>
                <TextBlock Text="AD Synchronization" FontSize="16" FontWeight="SemiBold" 
                           Margin="0,0,0,15" Foreground="#1A1A1A"/>
                
                <CheckBox Name="chkEnableADSync" Content="Enable automatic AD synchronization" 
                          Margin="0,0,0,10"/>
                
                <Grid IsEnabled="{Binding ElementName=chkEnableADSync, Path=IsChecked}">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="200"/>
                        <ColumnDefinition Width="*"/>
                    </Grid.ColumnDefinitions>
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>
                    
                    <TextBlock Grid.Row="0" Grid.Column="0" Text="Sync Group:" 
                               VerticalAlignment="Center" Margin="0,0,0,10"/>
                    <TextBox Grid.Row="0" Grid.Column="1" Name="txtADSyncGroup" 
                             Style="{StaticResource ModernTextBox}" Margin="0,0,0,10"/>
                    
                    <TextBlock Grid.Row="1" Grid.Column="0" Text="Sync Command:" 
                               VerticalAlignment="Center" Margin="0,0,0,10"/>
                    <TextBox Grid.Row="1" Grid.Column="1" Name="txtSyncCommand" 
                             Style="{StaticResource ModernTextBox}" Margin="0,0,0,10"
                             Text="Start-ADSyncSyncCycle -PolicyType Delta"/>
                    
                    <StackPanel Grid.Row="2" Grid.Column="1" Orientation="Horizontal">
                        <Button Name="btnTestSync" Content="Test Sync" Padding="10,5" Margin="0,0,10,0"/>
                        <Button Name="btnRunSync" Content="Run Sync Now" Padding="10,5"/>
                    </StackPanel>
                </Grid>
            </StackPanel>
        </Border>
        
        <Border Background="White" CornerRadius="8" Padding="20">
            <Border.Effect>
                <DropShadowEffect BlurRadius="10" Opacity="0.1" ShadowDepth="2"/>
            </Border.Effect>
            
            <StackPanel>
                <TextBlock Text="System Settings" FontSize="16" FontWeight="SemiBold" 
                           Margin="0,0,0,15" Foreground="#1A1A1A"/>
                
                <CheckBox Name="chkAutoBackup" Content="Enable automatic configuration backup" 
                          Margin="0,0,0,10"/>
                <CheckBox Name="chkCheckUpdates" Content="Check for updates automatically" 
                          Margin="0,0,0,10"/>
                <CheckBox Name="chkCollectUsageStats" Content="Send anonymous usage statistics" 
                          Margin="0,0,0,10"/>
                
                <Button Name="btnExportConfig" Content="Export Configuration..." 
                        HorizontalAlignment="Left" Padding="15,8" Margin="0,10,0,0"/>
            </StackPanel>
        </Border>
    </StackPanel>
'@
    
    return $xaml
}

# Export additional functions
Export-ModuleMember -Function @(
    'Add-ActiveDirectoryPanel',
    'Add-EmailSettingsPanel', 
    'Add-PasswordPolicyPanel',
    'Add-GroupsAndLicensesPanel',
    'Add-AdvancedSettingsPanel'
) 
# SIG # Begin signature block
# MIIcCAYJKoZIhvcNAQcCoIIb+TCCG/UCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCxZIhHuw4IimMQ
# A1+VdwT1gN5Q+1d+wo2BdAHEzWFL/qCCFk4wggMQMIIB+KADAgECAhB3jzsyX9Cg
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
# BgkqhkiG9w0BCQQxIgQgPw+nIrXXXn3rIufTHwphNaqckxui0PnXuKhAW9YZ8uMw
# DQYJKoZIhvcNAQEBBQAEggEAoOqNRVKS6I8OFxHI/RKU5y7QG0vQ8WtJ6C2XBObD
# rOd9ycG3S/Hl/UMKH7ELkYoJ+eoSRLEX24v/+Pauny8kiOGhEoQtF55C+nxeViCQ
# sc2mR0C4dGdQkPSUie11Iij9YHvWCO/a8IeiP0ZLSIG3YoZa86RhMumv0kiv8k3w
# 0C7A8MawdLhxeM1zjNiVBoenssW1DmxHnIklQTpGRy+ZlIx9dLVsRl0KcMpUO3ww
# zQxAHk+w7swJ9IODYDUFg6Lqh2jUE6EToBEoshGwwmg9S7VtrLO2dpI1Y0bVqpIy
# vOCaoFyVn3SxRHT7IJbbU7HcA1O9sYfRREzXCaPuLwrneaGCAyYwggMiBgkqhkiG
# 9w0BCQYxggMTMIIDDwIBATB9MGkxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdp
# Q2VydCwgSW5jLjFBMD8GA1UEAxM4RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3Rh
# bXBpbmcgUlNBNDA5NiBTSEEyNTYgMjAyNSBDQTECEAqA7xhLjfEFgtHEdqeVdGgw
# DQYJYIZIAWUDBAIBBQCgaTAYBgkqhkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqG
# SIb3DQEJBTEPFw0yNTA3MTExOTMyMjBaMC8GCSqGSIb3DQEJBDEiBCDQ+v9z5msv
# eSo4PCZeDZ+I8QdeuCo+xdeWaJS8zWbLLjANBgkqhkiG9w0BAQEFAASCAgBPt0Hm
# wVSCIlJ2ZfAdGJmXk5FAHiveRU4MMaRFDGP5O0hpClEt7hnFjBkzUa+hbJfop1MS
# dIQFEMlFGdS+V/txV+uJjid5nzLG8ClM6C08LdJ1lfa46d024/kiLC1YQk7sOmwS
# O3s/naNqQUOpoFdNlwXHPzlu28F7v/Am/3OgbxGEt9/0MsfiplUFnpLi8BfJakTp
# OLqxgBN3SP3axzOeWA5RFWU9d7+QhAeHUDTDjyUuHNCkDKzTI1PGoZbutB6+RlUm
# Zr55sUHiW0qXcH0tJignmlvar1urIvNpbUPS4kXCSBkacJnTWyhI29/l4Pf7aqbw
# ojmaTWwyy4nvZbEJCTmGcZ+3W0ZoxDGO1OyTAdanFi1ehPfmOxAojJ78RAC07PVP
# WRRHtqX38afYYIx00StIrrR6cXYGTny0Bxm80yluoRKdtWoNGNyJeueYgWXKSuV4
# Vx2XOGtcum+uvN0puxSeOx3yQesBFJLS8rlN7fQng5xn9vJ7Zev9esTyWdT1FJdN
# 9qpp0i6cYp3AixbRQ0HVYumMubwTMbvyVsGWdpEBfqAZ9EqOQTlJ7yO/uR4Lbprd
# Qrni1U0Uo0bg2Tn7xkPpXcYyW0ULfCkfnoBsL8P7U7DAxv0vdO2GUiTiP98ZnLQn
# flaCDl75eX6Pp/Q6apw257998OuZ3kiW9RKhew==
# SIG # End signature block
