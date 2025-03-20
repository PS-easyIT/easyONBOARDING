#region [Region 04 | ACTIVE DIRECTORY MODULE]
# Loads the Active Directory PowerShell module required for user management
Write-DebugMessage "Loading Active Directory module."
try {
    Write-DebugMessage "Loading AD module"
    Import-Module ActiveDirectory -SkipEditionCheck -ErrorAction Stop
    Write-DebugMessage "Active Directory module loaded successfully."
} catch {
    Write-Log -Message "ActiveDirectory module could not be loaded: $($_.Exception.Message)" -LogLevel "ERROR"
    Throw "Critical error: ActiveDirectory module missing!"
}
#endregion