# Lädt und verarbeitet GUI Event-Handling Funktionalität
# Version: 1.4.1 - Optimized

# Define Write-DebugMessage function if it doesn't exist
if (-not (Get-Command -Name Write-DebugMessage -ErrorAction SilentlyContinue)) {
    function Write-DebugMessage {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory=$true)]
            [string]$Message,
            [string]$LogLevel = "DEBUG"
        )
        Write-Verbose $Message
    }
}

Write-DebugMessage "Loading GUI event handler module."

# Function to register and handle GUI button click events with error handling
function Register-GUIEvent {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        [ValidateNotNull()]
        [System.Windows.Controls.Control]$Control,
        
        [Parameter(Mandatory=$true)]
        [ValidateNotNull()]
        [ScriptBlock]$EventAction,
        
        [Parameter(Mandatory=$false)]
        [string]$ErrorMessagePrefix = "Error executing event",
        
        [Parameter(Mandatory=$false)]
        [string]$EventName = "Click",
        
        [Parameter(Mandatory=$false)]
        [switch]$SuppressErrors
    )

    Write-DebugMessage "Registering GUI event handler ($EventName) for control: $($Control.Name)"

    if ($Control) {
        $eventMethod = "Add_$EventName"
        
        if (-not ($Control | Get-Member -Name $eventMethod -MemberType Method)) {
            $errorMsg = "Control '$($Control.Name)' does not support the '$EventName' event."
            Write-DebugMessage $errorMsg
            
            if (-not $SuppressErrors) {
                throw $errorMsg
            }
            return $false
        }
        
        $Control.$eventMethod({
            try {
                Write-DebugMessage "Executing $EventName event for $($Control.Name)"
                & $EventAction
            }
            catch {
                $errorDetails = $_.Exception.Message
                if ($_.Exception.InnerException) {
                    $errorDetails += " | Inner: $($_.Exception.InnerException.Message)"
                }
                
                $errorMsg = "$ErrorMessagePrefix`: $errorDetails"
                Write-DebugMessage "Error in event handler: $errorMsg"
                
                if (-not $SuppressErrors) {
                    [System.Windows.MessageBox]::Show($errorMsg, "Error", "OK", "Error")
                }
            }
        })
        
        return $true
    } else {
        Write-DebugMessage "Warning: Attempted to register event on null control"
        
        if (-not $SuppressErrors) {
            throw "Cannot register event: Control is null"
        }
        return $false
    }
}

# Enhanced function to register multiple events at once
function Register-GUIEvents {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        [ValidateNotNull()]
        [hashtable]$EventMappings,
        
        [Parameter(Mandatory=$false)]
        [string]$CommonErrorPrefix = "Error executing event",
        
        [Parameter(Mandatory=$false)]
        [switch]$SuppressErrors
    )
    
    Write-DebugMessage "Registering multiple GUI events: $($EventMappings.Count) handlers"
    $results = @()
    
    foreach ($controlName in $EventMappings.Keys) {
        $mapping = $EventMappings[$controlName]
        $control = $null
        
        # Allow for either direct control objects or control names with a parent window
        if ($mapping.ContainsKey('Control') -and $mapping.Control -ne $null) {
            $control = $mapping.Control
        }
        elseif ($mapping.ContainsKey('Window') -and $mapping.Window -ne $null) {
            $control = $mapping.Window.FindName($controlName)
        }
        elseif ($global:XamlWindow -ne $null) {
            $control = $global:XamlWindow.FindName($controlName)
        }
        elseif ($global:XamlElements -ne $null -and $global:XamlElements.ContainsKey($controlName)) {
            $control = $global:XamlElements[$controlName]
        }
        
        if ($control) {
            $eventName = if ($mapping.ContainsKey('EventName')) { $mapping.EventName } else { "Click" }
            $action = $mapping.Action
            $errorPrefix = if ($mapping.ContainsKey('ErrorPrefix')) { $mapping.ErrorPrefix } else { $CommonErrorPrefix }
            
            try {
                $success = Register-GUIEvent -Control $control -EventAction $action -EventName $eventName `
                            -ErrorMessagePrefix $errorPrefix -SuppressErrors:$SuppressErrors
                
                $results += [PSCustomObject]@{
                    ControlName = $controlName
                    EventName = $eventName
                    Success = $success
                }
            }
            catch {
                Write-DebugMessage "Error registering event for $controlName`: $($_.Exception.Message)"
                $results += [PSCustomObject]@{
                    ControlName = $controlName
                    EventName = $eventName
                    Success = $false
                    Error = $_.Exception.Message
                }
                
                if (-not $SuppressErrors) {
                    throw $_
                }
            }
        }
        else {
            Write-DebugMessage "Control not found: $controlName"
            $results += [PSCustomObject]@{
                ControlName = $controlName
                EventName = if ($mapping.ContainsKey('EventName')) { $mapping.EventName } else { "Click" }
                Success = $false
                Error = "Control not found"
            }
        }
    }
    
    return $results
}

# Function to find UI element by name in the WPF tree
function Find-UIElement {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        [ValidateNotNull()]
        [System.Windows.DependencyObject]$RootElement,
        
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$ElementName,
        
        [Parameter(Mandatory=$false)]
        [switch]$Recursive
    )
    
    Write-DebugMessage "Looking for element '$ElementName' in UI tree"
    
    # Try direct lookup first if root is a FrameworkElement
    if ($RootElement -is [System.Windows.FrameworkElement] -and $RootElement.Name -eq $ElementName) {
        return $RootElement
    }
    
    # Look for FindName method (usually on Window or UserControl)
    if ($RootElement | Get-Member -Name "FindName" -MemberType Method) {
        $element = $RootElement.FindName($ElementName)
        if ($element) {
            return $element
        }
    }
    
    # If not recursive, we're done
    if (-not $Recursive) {
        return $null
    }
    
    # Recursive search in the visual tree
    try {
        $childCount = [System.Windows.Media.VisualTreeHelper]::GetChildrenCount($RootElement)
        
        for ($i = 0; $i -lt $childCount; $i++) {
            $child = [System.Windows.Media.VisualTreeHelper]::GetChild($RootElement, $i)
            
            # Check if this child is the one we're looking for
            if ($child -is [System.Windows.FrameworkElement] -and $child.Name -eq $ElementName) {
                return $child
            }
            
            # Recursively search in the child's subtree
            $found = Find-UIElement -RootElement $child -ElementName $ElementName -Recursive
            if ($found) {
                return $found
            }
        }
    }
    catch {
        Write-DebugMessage "Error searching for UI element: $($_.Exception.Message)"
    }
    
    return $null
}

# Export the functions
Export-ModuleMember -Function Register-GUIEvent, Register-GUIEvents, Find-UIElement
