<#
.SYNOPSIS
    Core helper functions for PowerShell scripts
.DESCRIPTION
    This module provides essential helper functions used across PowerShell scripts
    including debug output, logging, and error handling.
.NOTES
    Version: 1.0
    Author: Created for easyONBOARDING
#>

# CoreHelpers.psm1 - Grundfunktionen für das easyONBOARDING Tool

#region Variables
$script:DebugEnabled = $false
$script:DebugMode = 0
$script:LogFile = $null
#endregion

#region Debug Functions
function Initialize-DebugSystem {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$false)]
        [bool]$DebugEnabled = $false,
        
        [Parameter(Mandatory=$false)]
        [int]$DebugLevel = 0,
        
        [Parameter(Mandatory=$false)]
        [string]$LogPath = $null
    )
    
    $script:DebugEnabled = $DebugEnabled
    $script:DebugMode = $DebugLevel
    
    if ($LogPath) {
        $script:LogFile = Join-Path -Path $LogPath -ChildPath "easyONB_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
        
        # Create log directory if it doesn't exist
        $logDir = Split-Path -Parent $script:LogFile
        if (-not (Test-Path -Path $logDir)) {
            New-Item -Path $logDir -ItemType Directory -Force | Out-Null
        }
        
        # Initialize log file with header
        $logHeader = "==========================================================`r`n"
        $logHeader += "easyONBOARDING Debug Log - Started $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`r`n"
        $logHeader += "Debug Level: $script:DebugMode`r`n"
        $logHeader += "==========================================================`r`n"
        
        $logHeader | Out-File -FilePath $script:LogFile -Encoding UTF8
    }
    
    # Return current settings
    return @{
        DebugEnabled = $script:DebugEnabled
        DebugLevel = $script:DebugMode
        LogFile = $script:LogFile
    }
}

function Write-DebugOutput {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true, Position=0)]
        [string]$Message,
        
        [Parameter(Mandatory=$false)]
        [ValidateSet("ERROR", "WARNING", "INFO", "DEBUG", "VERBOSE")]
        [string]$Level = "INFO"
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $caller = (Get-PSCallStack)[1].Command
    
    # Determine appropriate color
    $color = switch ($Level) {
        "ERROR"   { "Red"; break }
        "WARNING" { "Yellow"; break }
        "INFO"    { "White"; break }
        "DEBUG"   { "Cyan"; break }
        "VERBOSE" { "Gray"; break }
    }
    
    # Format message with timestamp and caller
    $formattedMessage = "[$timestamp] [$Level] [$caller] $Message"
    
    # Output to console
    Write-Host $formattedMessage -ForegroundColor $color
    
    # Log to file if enabled
    if ($script:LogFile) {
        $formattedMessage | Out-File -FilePath $script:LogFile -Encoding UTF8 -Append
    }
}

# Compatibility function for old calls
function Write-DebugMessage {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true, Position=0)]
        [string]$Message,
        
        [Parameter(Mandatory=$false)]
        [string]$LogLevel = "INFO"
    )
    
    # Map old log levels to new format
    $mappedLevel = switch ($LogLevel) {
        "ERROR" { "ERROR"; break }
        "WARNING" { "WARNING"; break }
        "INFO" { "INFO"; break }
        "DEBUG" { "DEBUG"; break }
        default { "INFO"; break }
    }
    
    Write-DebugOutput -Message $Message -Level $mappedLevel
}

function Write-LogMessage {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$Message,
        
        [Parameter(Mandatory=$false)]
        [ValidateSet("ERROR", "WARNING", "INFO", "DEBUG")]
        [string]$LogLevel = "INFO",
        
        [Parameter(Mandatory=$false)]
        [switch]$NoConsole
    )
    
    # Redirect to unified debug output
    if (-not $NoConsole) {
        Write-DebugOutput -Message $Message -Level $LogLevel
    }
    else {
        # Only log to file
        if ($script:LogFile) {
            $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            $caller = (Get-PSCallStack)[1].Command
            "[$timestamp] [$LogLevel] [$caller] $Message" | Out-File -FilePath $script:LogFile -Encoding UTF8 -Append
        }
    }
}
#endregion

#region UI Helper Functions
function Test-IsAdmin {
    [CmdletBinding()]
    param()
    
    $currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    return $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Start-ElevatedProcess {
    [CmdletBinding()]
    param()
    
    Write-DebugOutput "Restarting script with elevated privileges" -Level INFO
    
    try {
        $arguments = "-NoProfile -ExecutionPolicy Bypass -File `"$($MyInvocation.PSCommandPath)`""
        Start-Process -FilePath PowerShell.exe -ArgumentList $arguments -Verb RunAs
    }
    catch {
        Write-DebugOutput "Failed to restart with elevated privileges: $($_.Exception.Message)" -Level ERROR
        throw
    }
}

function Import-XamlWindow {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$XamlPath
    )
    
    Write-DebugOutput "Importing XAML from: $XamlPath" -Level INFO
    
    if (-not (Test-Path $XamlPath)) {
        Write-DebugOutput "XAML file not found: $XamlPath" -Level ERROR
        throw "XAML file not found: $XamlPath"
    }
    
    try {
        Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase
        
        [xml]$xaml = Get-Content -Path $XamlPath -ErrorAction Stop
        
        $reader = New-Object System.Xml.XmlNodeReader $xaml
        $window = [Windows.Markup.XamlReader]::Load($reader)
        
        if ($null -eq $window) {
            throw "Failed to load window from XAML"
        }
        
        Write-DebugOutput "XAML imported successfully" -Level INFO
        return $window
    }
    catch {
        Write-DebugOutput "Error importing XAML: $($_.Exception.Message)" -Level ERROR
        
        try {
            # Create fallback error window
            $errorWindow = New-Object System.Windows.Window
            $errorWindow.Title = "XAML Loading Error"
            $errorWindow.Width = 500
            $errorWindow.Height = 300
            
            $stackPanel = New-Object System.Windows.Controls.StackPanel
            $errorWindow.Content = $stackPanel
            
            $textBlock = New-Object System.Windows.Controls.TextBlock
            $textBlock.Text = "Error loading XAML from: $XamlPath`n$($_.Exception.Message)"
            $textBlock.Margin = New-Object System.Windows.Thickness(10)
            $textBlock.TextWrapping = [System.Windows.TextWrapping]::Wrap
            $stackPanel.Children.Add($textBlock)
            
            $closeButton = New-Object System.Windows.Controls.Button
            $closeButton.Content = "Close"
            $closeButton.Width = 100
            $closeButton.Margin = New-Object System.Windows.Thickness(10)
            $closeButton.Add_Click({ $errorWindow.Close() })
            $stackPanel.Children.Add($closeButton)
            
            Write-DebugOutput "Displaying error window" -Level INFO
            $errorWindow.ShowDialog() | Out-Null
        }
        catch {
            Write-DebugOutput "Failed to create error window: $($_.Exception.Message)" -Level ERROR
        }
        
        throw "Failed to load XAML: $($_.Exception.Message)"
    }
}

function Import-RequiredModule {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$ModuleName,
        
        [Parameter(Mandatory=$false)]
        [string[]]$Dependencies = @(),
        
        [Parameter(Mandatory=$false)]
        [switch]$Critical
    )
    
    $modulePath = Join-Path -Path $global:ModulesDir -ChildPath "$ModuleName.psm1"
    $result = $false
    
    try {
        Write-DebugOutput "Checking module: $ModuleName" -Level DEBUG
        
        if (-not (Test-Path $modulePath)) {
            if ($Critical) {
                Write-DebugOutput "Critical module not found: $ModuleName" -Level ERROR
                return $false
            }
            else {
                Write-DebugOutput "Module not found: $ModuleName" -Level WARNING
                return $false
            }
        }
        
        # First load dependencies
        foreach ($dependency in $Dependencies) {
            $depPath = Join-Path -Path $global:ModulesDir -ChildPath "$dependency.psm1"
            
            if (-not (Test-Path $depPath)) {
                Write-DebugOutput "Dependency not found: $dependency" -Level WARNING
                continue
            }
            
            if (-not (Get-Module -Name $dependency -ErrorAction SilentlyContinue)) {
                try {
                    Import-Module $depPath -Force -Global
                    Write-DebugOutput "Loaded dependency: $dependency" -Level DEBUG
                }
                catch {
                    Write-DebugOutput "Failed to load dependency $dependency: $($_.Exception.Message)" -Level WARNING
                }
            }
        }
        
        # Load the module
        Import-Module $modulePath -Force -Global
        Write-DebugOutput "Loaded module: $ModuleName" -Level INFO
        $result = $true
    }
    catch {
        if ($Critical) {
            Write-DebugOutput "Failed to load critical module $ModuleName: $($_.Exception.Message)" -Level ERROR
        }
        else {
            Write-DebugOutput "Failed to load module $ModuleName: $($_.Exception.Message)" -Level WARNING
        }
        $result = $false
    }
    
    return $result
}

function Find-UIElement {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$Name,
        
        [Parameter(Mandatory=$true)]
        [System.Windows.DependencyObject]$Parent
    )
    
    if ($null -eq $Parent) {
        Write-DebugOutput "Parent element is null when searching for $Name" -Level WARNING
        return $null
    }
    
    # First try FindName
    $element = $Parent.FindName($Name)
    if ($null -ne $element) {
        return $element
    }
    
    # Then try visual tree search
    for ($i = 0; $i -lt [System.Windows.Media.VisualTreeHelper]::GetChildrenCount($Parent); $i++) {
        $child = [System.Windows.Media.VisualTreeHelper]::GetChild($Parent, $i)
        
        # Check if child is a framework element with this name
        if ($child -is [System.Windows.FrameworkElement] -and $child.Name -eq $Name) {
            return $child
        }
        
        # Recursively search children
        $result = Find-UIElement -Name $Name -Parent $child
        if ($null -ne $result) {
            return $result
        }
    }
    
    return $null
}

function Register-GUIEvent {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [System.Windows.UIElement]$Control,
        
        [Parameter(Mandatory=$true)]
        [ScriptBlock]$EventAction,
        
        [Parameter(Mandatory=$true)]
        [string]$EventName,
        
        [Parameter(Mandatory=$false)]
        [string]$ErrorMessagePrefix = "Error in event handler"
    )
    
    if ($null -eq $Control) {
        Write-DebugOutput "Control is null when trying to register $EventName event" -Level WARNING
        return $false
    }
    
    try {
        # Create the event handler with error handling
        $safeEventAction = {
            param($sender, $e)
            
            try {
                # Invoke the original event handler
                & $EventAction $sender $e
            }
            catch {
                Write-DebugOutput "$ErrorMessagePrefix`: $($_.Exception.Message)" -Level ERROR
                [System.Windows.MessageBox]::Show(
                    "$ErrorMessagePrefix`n`n$($_.Exception.Message)",
                    "Error",
                    [System.Windows.MessageBoxButton]::OK,
                    [System.Windows.MessageBoxImage]::Error
                )
            }
        }
        
        # Add the event handler
        $eventMethod = "Add_$EventName"
        $Control.$eventMethod($safeEventAction)
        
        return $true
    }
    catch {
        Write-DebugOutput "Failed to register $EventName event: $($_.Exception.Message)" -Level ERROR
        return $false
    }
}
#endregion

# Export functions
Export-ModuleMember -Function Write-DebugOutput, Write-DebugMessage, Initialize-DebugSystem,
    Write-LogMessage, Test-IsAdmin, Start-ElevatedProcess, Import-XamlWindow,
    Import-RequiredModule, Find-UIElement, Register-GUIEvent
