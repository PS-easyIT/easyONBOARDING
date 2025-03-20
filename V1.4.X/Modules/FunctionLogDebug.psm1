#region [Region 03 | FUNCTION DEFINITIONS]
# Contains all helper and core functionality functions

#region [Region 03.1 | LOGGING FUNCTIONS]
# Defines logging capabilities for different message levels
function Write-LogMessage {
    # [03.1.1 - Primary logging wrapper for consistent message formatting]
    param(
        [string]$message,
        [string]$logLevel = "INFO"
    )
    Write-Log -message $message -logLevel $logLevel
}

function Write-DebugMessage {
    # [03.1.2 - Debug-specific logging with conditional execution]
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true, ValueFromPipeline=$true)]
        [AllowEmptyString()]
        [string]$Message,
        
        [Parameter(Mandatory=$false)]
        [string]$LogLevel = "DEBUG"
    )
    process {
        # Only log debug messages if DebugMode is enabled in config
        if ($null -ne $global:Config -and 
            $null -ne $global:Config.Logging -and 
            $global:Config.Logging.DebugMode -eq "1") {
            
            # Call Write-Log without capturing output to avoid pipeline return
            Write-Log -Message $Message -LogLevel $LogLevel
            
            # Also output to console for immediate feedback during debugging
            Write-Host "[DEBUG] $Message" -ForegroundColor Cyan
        }
        
        # No return value to avoid unwanted pipeline output
    }
}
#endregion

# Log level (default is "INFO"): "WARN", "ERROR", "DEBUG".

#region [Region 03.2 | LOG FILE WRITER] 
# Core logging function that writes messages to log files
function Write-Log {
    # [03.2.1 - Low-level file logging implementation]
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true, ValueFromPipeline=$true)]
        [AllowEmptyString()]
        [string]$Message,
        
        [Parameter(Mandatory=$false)]
        [ValidateSet("INFO", "WARN", "ERROR", "DEBUG")]
        [string]$LogLevel = "INFO"
    )
    process {
        # log path determination using null-conditional and coalescing operators
        $logFilePath = $global:Config?.Logging?.ExtraFile ?? 
                      $global:Config?.Logging?.LogFile ?? 
                      (Join-Path $ScriptDir "Logs")
        
        # Ensure the log directory exists
        if (-not (Test-Path $logFilePath)) {
            try {
                # Use -Force to create parent directories if needed
                [void](New-Item -ItemType Directory -Path $logFilePath -Force -ErrorAction Stop)
                Write-Host "Created log directory: $logFilePath"
            } catch {
                # Handle directory creation errors gracefully
                Write-Warning "Error creating log directory: $($_.Exception.Message)"
                # Try writing to a fallback location if possible
                $logFilePath = $env:TEMP
            }
        }

        # Define log file paths
        $logFile = Join-Path -Path $logFilePath -ChildPath "easyOnboarding.log"
        $errorLogFile = Join-Path -Path $logFilePath -ChildPath "easyOnboarding_error.log"
        
        # Generate timestamp and log entry
        $timeStamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        $logEntry = "$timeStamp [$LogLevel] $Message"

        try {
            Add-Content -Path $logFile -Value $logEntry -ErrorAction Stop
        } catch {
            # Log additional error if writing to log file fails
            try {
                Add-Content -Path $errorLogFile -Value "$timeStamp [ERROR] Error writing to log file: $($_.Exception.Message)" -ErrorAction SilentlyContinue
            } catch {
                # Last resort: output to console
                Write-Warning "Critical error: Cannot write to any log file: $($_.Exception.Message)"
            }
        }

        # Output to console based on DebugMode setting
        if ($null -ne $global:Config -and $null -ne $global:Config.Logging) {
            # When DebugMode=1, output all messages regardless of log level
            if ($global:Config.Logging.DebugMode -eq "1") {
                # Use color coding based on log level
                switch ($LogLevel) {
                    "DEBUG" { Write-Host "$logEntry" -ForegroundColor Cyan }
                    "INFO"  { Write-Host "$logEntry" -ForegroundColor White }
                    "WARN"  { Write-Host "$logEntry" -ForegroundColor Yellow }
                    "ERROR" { Write-Host "$logEntry" -ForegroundColor Red }
                    default { Write-Host "$logEntry" }
                }
            }
            # When DebugMode=0, no console output
        }
        
        # No return value to avoid unwanted pipeline output
    }
}
#endregion
