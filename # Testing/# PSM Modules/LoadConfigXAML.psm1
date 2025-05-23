# Lädt und verarbeitet die XAML UI-Definition
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

Write-DebugMessage "Loading XAML configuration."

# Funktion zur Überprüfung und zum Laden von WPF-Assemblies
function Initialize-WPFAssemblies {
    try {
        # Lade notwendige Assemblies für WPF
        Add-Type -AssemblyName PresentationFramework -ErrorAction Stop
        Add-Type -AssemblyName PresentationCore -ErrorAction Stop
        Add-Type -AssemblyName WindowsBase -ErrorAction Stop
        return $true
    }
    catch {
        Write-Error "Fehler beim Laden der WPF-Assemblies: $($_.Exception.Message)"
        return $false
    }
}

# Funktion zum Import von XAML
function Import-XAMLGUI {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$XAMLPath
    )
    
    # Stelle sicher, dass WPF-Assemblies geladen sind
    if (-not (Initialize-WPFAssemblies)) {
        throw "WPF-Assemblies konnten nicht geladen werden"
    }
    
    try {
        if (-not (Test-Path $XAMLPath)) {
            throw "XAML-Datei nicht gefunden: $XAMLPath"
        }
        
        [xml]$xaml = Get-Content -Path $XAMLPath
        
        # XML-Namespace für korrekte Verarbeitung hinzufügen
        $nsManager = New-Object System.Xml.XmlNamespaceManager($xaml.NameTable)
        $nsManager.AddNamespace("x", "http://schemas.microsoft.com/winfx/2006/xaml")
        
        # XMLReader erstellen
        $reader = New-Object System.Xml.XmlNodeReader $xaml
        
        # XAML laden
        $window = [Windows.Markup.XamlReader]::Load($reader)
        
        return $window
    }
    catch {
        Write-Error "Fehler beim Laden des XAML: $($_.Exception.Message)"
        throw
    }
}

Write-DebugMessage "Loading XAML configuration."

function Import-XamlFromFile {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$XamlPath
    )
    
    try {
        Write-DebugMessage "Reading XAML file: $XamlPath"
        if (-not (Test-Path -Path $XamlPath)) {
            throw "XAML file not found at: $XamlPath"
        }
        
        # Read the content first to check for common errors
        $xamlContent = Get-Content -Path $XamlPath -Raw -ErrorAction Stop
        
        # Check for C-style comments which are invalid in XAML
        if ($xamlContent -match '//.*?\n') {
            Write-DebugMessage "WARNING: Found C-style comments in XAML file. These are invalid and may cause loading issues."
            $xamlContent = $xamlContent -replace '//.*?\n', "`n"
        }
        
        # Validate XAML content
        if ($xamlContent -match '<.*>') {
            try {
                $reader = New-Object System.Xml.XmlTextReader (New-Object System.IO.StringReader $xamlContent)
                return [Windows.Markup.XamlReader]::Load($reader)
            } catch {
                $detailedError = "Error parsing XAML content: $($_.Exception.Message)"
                if ($_.Exception.InnerException) {
                    $detailedError += " Inner exception: $($_.Exception.InnerException.Message)"
                }
                Write-DebugMessage $detailedError
                throw $detailedError
            }
        } else {
            throw "Invalid XAML content - file does not contain valid XML markup"
        }
    }
    catch {
        Write-DebugMessage "Error loading XAML file: $($_.Exception.Message)"
        throw "Error loading XAML: $($_.Exception.Message)"
    }
}

# Improved function to get XAML paths with better error handling
function Get-XamlFilePathCandidates {
    [CmdletBinding()]
    [OutputType([System.Collections.Generic.List[string]])]
    param()
    
    $xamlPaths = New-Object System.Collections.Generic.List[string]

    # First check if the global path is set
    if ($global:XAMLPath -and (Test-Path -Path $global:XAMLPath)) {
        Write-DebugMessage "Using globally defined XAML path: $global:XAMLPath"
        $xamlPaths.Add($global:XAMLPath)
    }

    # Add additional path options
    $scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
    $possiblePaths = @(
        (Join-Path -Path $scriptDir -ChildPath "MainGUI.xaml"),
        (Join-Path -Path (Split-Path -Parent $scriptDir) -ChildPath "MainGUI.xaml"),
        (Join-Path -Path (Split-Path -Parent $PSScriptRoot) -ChildPath "MainGUI.xaml")
    )

    foreach ($path in $possiblePaths) {
        if (Test-Path -Path $path) {
            Write-DebugMessage "Found XAML file at: $path"
            $xamlPaths.Add($path)
        }
    }

    # Display diagnostic information about paths
    Write-DebugMessage "All potential XAML paths:"
    foreach ($potentialPath in $possiblePaths) {
        $exists = Test-Path -Path $potentialPath
        Write-DebugMessage "  $potentialPath - $(if($exists){'EXISTS'}else{'NOT FOUND'})"
    }
    
    return $xamlPaths
}

# Function to extract all named elements from XAML window
function Get-XamlNamedElements {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        [ValidateNotNull()]
        [object]$XamlObject
    )
    
    try {
        # Create a hashtable to store all named elements
        $namedElements = @{}
        
        # Helper function to recursively find named elements
        function Find-NamedChildren {
            param (
                [Parameter(Mandatory=$true)]
                [object]$Parent
            )
            
            # Check if the current element has a name
            if ($Parent.Name) {
                $namedElements[$Parent.Name] = $Parent
            }
            
            # Get all child elements using VisualTreeHelper for better compatibility
            try {
                $childCount = [System.Windows.Media.VisualTreeHelper]::GetChildrenCount($Parent)
                for ($i = 0; $i -lt $childCount; $i++) {
                    $child = [System.Windows.Media.VisualTreeHelper]::GetChild($Parent, $i)
                    if ($child -is [Windows.FrameworkElement]) {
                        Find-NamedChildren -Parent $child
                    }
                }
            } catch {
                # Fallback for logical tree if visual tree isn't available yet
                $childrenProperties = $Parent.GetType().GetProperties() | Where-Object {
                    try {
                        if ($_.PropertyType.Name -like "*Collection*" -and $_.PropertyType.Name -notmatch "^String") {
                            $values = $Parent.$($_.Name)
                            return $values -and $values.Count -gt 0 -and $values[0] -is [Windows.FrameworkElement]
                        } else {
                            $value = $Parent.$($_.Name)
                            return $value -is [Windows.FrameworkElement]
                        }
                    } catch {
                        return $false
                    }
                }
                
                # Process each property that contains child elements
                foreach ($prop in $childrenProperties) {
                    try {
                        $values = $Parent.$($prop.Name)
                        if ($values -is [System.Collections.IEnumerable] -and $values -isnot [string]) {
                            foreach ($child in $values) {
                                if ($child -is [Windows.FrameworkElement]) {
                                    Find-NamedChildren -Parent $child
                                }
                            }
                        } elseif ($values -is [Windows.FrameworkElement]) {
                            Find-NamedChildren -Parent $values
                        }
                    } catch {
                        # Skip this property if we can't access it
                        continue
                    }
                }
            }
        }
        
        # Start the recursive search from the root element
        Find-NamedChildren -Parent $XamlObject
        
        return $namedElements
    } catch {
        Write-DebugMessage "Error extracting named elements: $($_.Exception.Message)"
        throw "Error extracting named elements: $($_.Exception.Message)"
    }
}

# The main function that loads and processes the XAML UI
function Load-XamlInterface {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$false)]
        [string]$XamlPath
    )
    
    try {
        $paths = New-Object System.Collections.Generic.List[string]
        
        # If a specific path is provided, use it first
        if ($XamlPath) {
            $paths.Add($XamlPath)
        }
        
        # Add all other paths from get-XamlFilePathCandidates function
        $additionalPaths = Get-XamlFilePathCandidates
        foreach ($path in $additionalPaths) {
            if (-not $paths.Contains($path)) {
                $paths.Add($path)
            }
        }
        
        # Try each path until one works
        $window = $null
        $errorMessages = @()
        
        foreach ($path in $paths) {
            try {
                Write-DebugMessage "Attempting to load XAML from: $path"
                $window = Import-XamlFromFile -XamlPath $path
                if ($window) {
                    Write-DebugMessage "Successfully loaded XAML from: $path"
                    # Store the successful path for future reference
                    $global:SuccessfulXamlPath = $path
                    break
                }
            } catch {
                $errorMessages += "Failed to load XAML from '$path': $($_.Exception.Message)"
                Write-DebugMessage "Failed to load XAML from '$path': $($_.Exception.Message)"
                # Continue to the next path
            }
        }
        
        if (-not $window) {
            $errorMsg = "Failed to load XAML from any of the provided paths:`n$($errorMessages -join "`n")"
            Write-DebugMessage $errorMsg
            throw $errorMsg
        }
        
        # Extract all named elements
        $namedElements = Get-XamlNamedElements -XamlObject $window
        Write-DebugMessage "Found $($namedElements.Count) named elements in XAML"
        
        # Store results in global variables for ease of access from other modules
        $global:XamlWindow = $window
        $global:XamlElements = $namedElements
        
        # Return both the window and named elements
        return @{
            Window = $window
            Elements = $namedElements
        }
    } catch {
        Write-DebugMessage "Error in Load-XamlInterface: $($_.Exception.Message)"
        throw "Error loading XAML interface: $($_.Exception.Message)"
    }
}

# Try to load the XAML interface when the module is imported
try {
    $global:XamlResult = Load-XamlInterface
    Write-DebugMessage "XAML interface loaded automatically: $($global:XamlResult.Window.GetType().Name)"
} catch {
    Write-DebugMessage "Failed to automatically load XAML interface: $($_.Exception.Message)"
    # Not throwing an error here to allow manual loading later
}

# Export the functions
Export-ModuleMember -Function Import-XamlFromFile, Get-XamlNamedElements, Load-XamlInterface, Get-XamlFilePathCandidates, Import-XAMLGUI, Initialize-WPFAssemblies