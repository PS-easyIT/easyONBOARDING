# Consolidates error handling for Active Directory commands
Write-DebugMessage "Executing AD command."
function Invoke-ADCommand {
    param (
        [Parameter(Mandatory=$true)]
        [ScriptBlock]$Command,
        
        [Parameter(Mandatory=$true)]
        [string]$ErrorContext
    )
    try {
        & $Command
    }
    catch {
        $errMsg = "Error in Invoke-ADCommand: $ErrorContext - $($_.Exception.Message)"
        Write-Error $errMsg
        Throw $errMsg
    }
}

