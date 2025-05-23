# Function to open the INI configuration editor
Write-DebugMessage "Opening INI editor."
function Open-INIEditor {
    try {
        # Check if $global:Config is defined
        if ($global:Config) {
            if ($global:Config.Logging.DebugMode -eq "1") {
                # Check if Write-Log function exists before calling it
                if (Get-Command Write-Log -ErrorAction SilentlyContinue) {
                    Write-Log " INI Editor is starting." "DEBUG"
                } else {
                    Write-Warning "Write-Log function not found. Debug logging will be skipped."
                }
            }
        } else {
            Write-Warning "\$global:Config is not defined. Please ensure the configuration is loaded."
        }
        Write-DebugMessage "INI Editor has started."
        Write-LogMessage -Message "INI Editor has started." -Level "Info"
        [System.Windows.MessageBox]::Show("INI Editor (Settings) – Implement functionality here.", "Settings")
    }
    catch {
        Throw "Error opening INI Editor: $_"
    }
}
# Export the function to make it available outside the module
Export-ModuleMember -Function Open-INIEditor

