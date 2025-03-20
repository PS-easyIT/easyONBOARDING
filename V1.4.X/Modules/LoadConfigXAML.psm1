#region [Region 05 | XAML CONFIGURATION LOADING]
# Loads and processes the XAML UI definition
Write-DebugMessage "Loading XAML configuration."

function Import-XamlFromFile {
    param(
        [Parameter(Mandatory=$true)]
        [string]$XamlPath
    )
    
    try {
        Write-DebugMessage "Reading XAML file: $XamlPath"
        if (Test-Path -Path $XamlPath) {
            [xml]$xamlContent = Get-Content -Path $XamlPath
            $reader = New-Object System.Xml.XmlNodeReader $xamlContent
            return [Windows.Markup.XamlReader]::Load($reader)
        }
        else {
            Write-DebugMessage "XAML file not found: $XamlPath"
            throw "XAML file not found: $XamlPath"
        }
    }
    catch {
        Write-DebugMessage "Error loading XAML file: $($_.Exception.Message)"
        throw "Error loading XAML file: $($_.Exception.Message)"
    }
}

# Define possible XAML file paths in order of preference
$xamlPaths = @(
    (Join-Path -Path $ScriptDir -ChildPath "GUI\easyONBOARDING.xaml"),
    (Join-Path -Path $ScriptDir -ChildPath "MainGUI.xaml")
)

# Try each path until we find a valid XAML file
$window = $null
$xamlFound = $false

foreach ($xamlPath in $xamlPaths) {
    try {
        Write-DebugMessage "Attempting to load XAML from: $xamlPath"
        if (Test-Path $xamlPath) {
            $window = Import-XamlFromFile -XamlPath $xamlPath
            Write-DebugMessage "XAML UI loaded successfully from: $xamlPath"
            if ($global:Config.Logging.DebugMode -eq "1") {
                Write-Log "XAML file successfully loaded from: $xamlPath" "DEBUG"
            }
            $xamlFound = $true
            break
        }
    } catch {
        Write-DebugMessage "Failed to load XAML from $xamlPath: $($_.Exception.Message)"
    }
}

if (-not $xamlFound) {
    $errorMsg = "Critical Error: No valid XAML file found at any of the expected locations."
    Write-LogMessage -Message $errorMsg -LogLevel "ERROR"
    throw $errorMsg
}
#endregion