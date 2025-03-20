# Function to open the INI configuration editor
Write-DebugMessage "Opening INI editor."
function Open-INIEditor {
    try {
        if ($global:Config.Logging.DebugMode -eq "1") {
            Write-Log " INI Editor is starting." "DEBUG"
        }
        Write-DebugMessage "INI Editor has started."
        Write-LogMessage -Message "INI Editor has started." -Level "Info"
        [System.Windows.MessageBox]::Show("INI Editor (Settings) – Implement functionality here.", "Settings")
    }
    catch {
        Throw "Error opening INI Editor: $_"
    }
}
#endregion

