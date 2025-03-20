# Module für GUI Event-Handling Funktionalität
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

# Function to register and handle GUI events with improved error handling
function Register-GUIEvent {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        [ValidateNotNull()]
        [System.Windows.UIElement]$Control,
        
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
        # Überprüfen, ob der Control-Typ das angegebene Ereignis unterstützt
        $eventMethod = "Add_$EventName"
        
        if (-not ($Control | Get-Member -Name $eventMethod -MemberType Method)) {
            $errorMsg = "Control '$($Control.Name)' does not support the '$EventName' event."
            Write-DebugMessage $errorMsg -LogLevel "ERROR"
            
            if (-not $SuppressErrors) {
                throw $errorMsg
            }
            return $false
        }
        
        # Event-Handler definieren und registrieren
        $handler = {
            param($sender, $eventArgs)
            
            try {
                Write-DebugMessage "Executing $EventName event for $($Control.Name)"
                
                # Ursprüngliches Skriptblock mit Sender und EventArgs ausführen, falls sie verwendet werden
                if ($EventAction.GetNewClosure().ParameterSets.Count -gt 0 -and 
                    $EventAction.GetNewClosure().ParameterSets[0].Parameters.Count -ge 2) {
                    & $EventAction $sender $eventArgs
                } else {
                    & $EventAction
                }
            }
            catch {
                # Detailliertere Fehlerinformationen sammeln
                $errorDetails = $_.Exception.Message
                if ($_.Exception.InnerException) {
                    $errorDetails += " | Inner: $($_.Exception.InnerException.Message)"
                }
                
                $stackTrace = $_.ScriptStackTrace
                $errorMsg = "$ErrorMessagePrefix`: $errorDetails"
                
                Write-DebugMessage "Error in event handler: $errorMsg" -LogLevel "ERROR"
                Write-DebugMessage "Stack trace: $stackTrace" -LogLevel "DEBUG"
                
                if (-not $SuppressErrors) {
                    [System.Windows.MessageBox]::Show($errorMsg, "Error", "OK", "Error")
                }
            }
        }
        
        # Event registrieren mit dem korrekt definierten Handler
        $Control.$eventMethod($handler)
        
        return $true
    } else {
        Write-DebugMessage "Warning: Attempted to register event on null control" -LogLevel "WARNING"
        
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
    $results = [System.Collections.ArrayList]::new()
    
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
        elseif ($null -ne (Get-Variable -Name 'XamlWindow' -Scope 'Global' -ErrorAction SilentlyContinue)) {
            $control = $global:XamlWindow.FindName($controlName)
        }
        elseif ($null -ne (Get-Variable -Name 'window' -Scope 'Global' -ErrorAction SilentlyContinue)) {
            $control = $global:window.FindName($controlName)
        }
        elseif ($null -ne (Get-Variable -Name 'XamlElements' -Scope 'Global' -ErrorAction SilentlyContinue) -and 
                $global:XamlElements.ContainsKey($controlName)) {
            $control = $global:XamlElements[$controlName]
        }
        
        if ($control) {
            $eventName = if ($mapping.ContainsKey('EventName')) { $mapping.EventName } else { "Click" }
            $action = $mapping.Action
            $errorPrefix = if ($mapping.ContainsKey('ErrorPrefix')) { $mapping.ErrorPrefix } else { $CommonErrorPrefix }
            
            try {
                $success = Register-GUIEvent -Control $control -EventAction $action -EventName $eventName `
                            -ErrorMessagePrefix $errorPrefix -SuppressErrors:$SuppressErrors
                
                [void]$results.Add([PSCustomObject]@{
                    ControlName = $controlName
                    EventName = $eventName
                    Success = $success
                })
            }
            catch {
                Write-DebugMessage "Error registering event for $controlName`: $($_.Exception.Message)" -LogLevel "ERROR"
                [void]$results.Add([PSCustomObject]@{
                    ControlName = $controlName
                    EventName = $eventName
                    Success = $false
                    Error = $_.Exception.Message
                })
                
                if (-not $SuppressErrors) {
                    throw $_
                }
            }
        }
        else {
            Write-DebugMessage "Control not found: $controlName" -LogLevel "WARNING"
            [void]$results.Add([PSCustomObject]@{
                ControlName = $controlName
                EventName = if ($mapping.ContainsKey('EventName')) { $mapping.EventName } else { "Click" }
                Success = $false
                Error = "Control not found"
            })
        }
    }
    
    return $results
}

# Function to find UI element by name in the WPF tree - improved to search both visual and logical trees
function Find-UIElement {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true, Position=0)]
        [ValidateNotNull()]
        [System.Windows.DependencyObject]$Parent,
        
        [Parameter(Mandatory=$true, Position=1)]
        [ValidateNotNullOrEmpty()]
        [string]$Name,
        
        [Parameter(Mandatory=$false)]
        [switch]$Recursive = $true
    )
    
    Write-DebugMessage "Looking for element '$Name' in UI tree"
    
    # Try direct lookup first if parent is a FrameworkElement with FindName
    if ($Parent -is [System.Windows.FrameworkElement]) {
        # Fast path with FindName if available
        if ($Parent | Get-Member -Name "FindName" -MemberType Method) {
            $element = $Parent.FindName($Name)
            if ($element) {
                Write-DebugMessage "Found element '$Name' using FindName"
                return $element
            }
        }
        
        # Check if the parent itself is the element we're looking for
        if ($Parent.Name -eq $Name) {
            Write-DebugMessage "Parent element is the requested element: '$Name'"
            return $Parent
        }
    }
    
    # If we don't want recursive search, we're done
    if (-not $Recursive) {
        return $null
    }
    
    # TRY: Search in the logical tree first (more efficient in many cases)
    if ($Parent -is [System.Windows.FrameworkElement] -or $Parent -is [System.Windows.FrameworkContentElement]) {
        try {
            $logicalChildren = [System.Windows.LogicalTreeHelper]::GetChildren($Parent)
            foreach ($child in $logicalChildren) {
                if ($child -is [System.Windows.FrameworkElement] -and $child.Name -eq $Name) {
                    Write-DebugMessage "Found element '$Name' in logical tree"
                    return $child
                }
                
                # Recursively search in the logical child's subtree
                $found = Find-UIElement -Parent $child -Name $Name -Recursive
                if ($found) {
                    return $found
                }
            }
        }
        catch {
            Write-DebugMessage "Error searching logical tree: $($_.Exception.Message)" -LogLevel "DEBUG"
            # Continue to visual tree search if logical search fails
        }
    }
    
    # CATCH: Fallback to the visual tree search if logical search didn't find it
    try {
        $childCount = [System.Windows.Media.VisualTreeHelper]::GetChildrenCount($Parent)
        
        for ($i = 0; $i -lt $childCount; $i++) {
            $child = [System.Windows.Media.VisualTreeHelper]::GetChild($Parent, $i)
            
            # Check if this child is the one we're looking for
            if ($child -is [System.Windows.FrameworkElement] -and $child.Name -eq $Name) {
                Write-DebugMessage "Found element '$Name' in visual tree"
                return $child
            }
            
            # Recursively search in the child's subtree
            $found = Find-UIElement -Parent $child -Name $Name -Recursive
            if ($found) {
                return $found
            }
        }
    }
    catch {
        Write-DebugMessage "Error searching visual tree: $($_.Exception.Message)" -LogLevel "DEBUG"
    }
    
    # Element not found
    Write-DebugMessage "Element '$Name' not found in either logical or visual tree" -LogLevel "DEBUG"
    return $null
}

# Helper function to check if a control supports a specific event
function Test-ControlSupportsEvent {
    [CmdletBinding()]
    [OutputType([bool])]
    param (
        [Parameter(Mandatory=$true)]
        [ValidateNotNull()]
        [System.Windows.UIElement]$Control,
        
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$EventName
    )
    
    try {
        $eventMethod = "Add_$EventName"
        return ($Control | Get-Member -Name $eventMethod -MemberType Method) -ne $null
    }
    catch {
        Write-DebugMessage "Error testing event support: $($_.Exception.Message)" -LogLevel "DEBUG"
        return $false
    }
}

# Export the functions
Export-ModuleMember -Function Register-GUIEvent, Register-GUIEvents, Find-UIElement, Test-ControlSupportsEvent
