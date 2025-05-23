<#
.SYNOPSIS
    Tests PowerShell script performance and identifies bottlenecks.

.DESCRIPTION
    This script runs performance analysis on PowerShell scripts to identify
    potential bottlenecks and suggest performance improvements.

.PARAMETER ScriptPath
    Path to the PowerShell script to be analyzed.

.PARAMETER Iterations
    Number of times to run the script for performance measurements.

.EXAMPLE
    .\Test-ScriptPerformance.ps1 -ScriptPath "C:\Scripts\MyScript.ps1" -Iterations 5
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory = $true)]
    [string]$ScriptPath,
    
    [Parameter(Mandatory = $false)]
    [int]$Iterations = 3
)

function Measure-ScriptBlock {
    param (
        [ScriptBlock]$Script,
        [int]$Iterations
    )
    
    $results = @()
    
    for ($i = 1; $i -le $Iterations; $i++) {
        $sw = [System.Diagnostics.Stopwatch]::StartNew()
        
        # Measure CPU and memory usage
        $process = Get-Process -Id $pid
        $startCPU = $process.CPU
        $startMemory = $process.WorkingSet64
        
        # Execute script
        & $Script
        
        # Get final measurements
        $process = Get-Process -Id $pid
        $endCPU = $process.CPU
        $endMemory = $process.WorkingSet64
        
        $sw.Stop()
        
        $results += [PSCustomObject]@{
            Iteration = $i
            ExecutionTime = $sw.Elapsed
            CPUUsage = ($endCPU - $startCPU)
            MemoryUsageMB = [Math]::Round(($endMemory - $startMemory) / 1MB, 2)
        }
        
        # Force garbage collection between runs
        [System.GC]::Collect()
        Start-Sleep -Milliseconds 500
    }
    
    return $results
}

function Analyze-ScriptPerformance {
    param (
        [string]$ScriptContent
    )
    
    $recommendations = @()
    
    # Check for inefficient patterns
    if ($ScriptContent -match "Get-WmiObject") {
        $recommendations += @{
            Issue = "Use of Get-WmiObject detected"
            Recommendation = "Replace Get-WmiObject with Get-CimInstance for better performance"
            Impact = "Medium"
        }
    }
    
    if ($ScriptContent -match "ForEach\s*\(\s*\$\w+\s+in\s+\$\w+\s*\)") {
        $recommendations += @{
            Issue = "ForEach loop detected"
            Recommendation = "Consider using the ForEach-Object cmdlet or pipeline for better memory efficiency with large collections"
            Impact = "Low"
        }
    }
    
    if ($ScriptContent -match "\s+\|\s+Where-Object\s+\{.*\}\s+\|\s+ForEach-Object") {
        $recommendations += @{
            Issue = "Where-Object followed by ForEach-Object pattern detected"
            Recommendation = "Consider using Where() and ForEach() methods instead for better performance with large collections"
            Impact = "Medium"
        }
    }
    
    if ($ScriptContent -match "\s+\|\s+Select-Object\s+") {
        $recommendations += @{
            Issue = "Select-Object usage in pipeline detected"
            Recommendation = "If only selecting a few properties early in a large pipeline, consider using calculated properties instead"
            Impact = "Low"
        }
    }
    
    if (($ScriptContent -match "Add-Content") -or ($ScriptContent -match "Out-File -Append")) {
        $recommendations += @{
            Issue = "Repeated file append operations detected"
            Recommendation = "For multiple append operations, consider collecting output and writing once at the end"
            Impact = "High"
        }
    }
    
    if ($ScriptContent -match "\[\w+\[\]\]\s*\$\w+\s*=\s*@\(\)") {
        $recommendations += @{
            Issue = "Fixed-size array with additions detected"
            Recommendation = "Use ArrayList or Generic List<T> when frequently adding items to collections"
            Impact = "Medium"
        }
    }
    
    return $recommendations
}

# Main script execution
try {
    if (-not (Test-Path -Path $ScriptPath)) {
        Write-Error "Script not found: $ScriptPath"
        return
    }
    
    Write-Host "Performance Testing: $ScriptPath" -ForegroundColor Cyan
    Write-Host "Running $Iterations iterations..." -ForegroundColor Cyan
    Write-Host "----------------------------------------" -ForegroundColor Cyan
    
    # Check if script can be loaded as a script block
    $scriptContent = Get-Content -Path $ScriptPath -Raw
    
    try {
        $scriptBlock = [ScriptBlock]::Create($scriptContent)
        
        # Measure performance
        $results = Measure-ScriptBlock -Script $scriptBlock -Iterations $Iterations
        
        # Display performance results
        $averageTime = $results | Measure-Object -Property ExecutionTime -Average | Select-Object -ExpandProperty Average
        $averageCPU = $results | Measure-Object -Property CPUUsage -Average | Select-Object -ExpandProperty Average
        $averageMemory = $results | Measure-Object -Property MemoryUsageMB -Average | Select-Object -ExpandProperty Average
        
        Write-Host "Performance Results:" -ForegroundColor Green
        Write-Host "  Average Execution Time: $($averageTime.ToString("hh\:mm\:ss\.fff"))" -ForegroundColor Yellow
        Write-Host "  Average CPU Usage: $([Math]::Round($averageCPU, 2))" -ForegroundColor Yellow
        Write-Host "  Average Memory Usage: $([Math]::Round($averageMemory, 2)) MB" -ForegroundColor Yellow
        
        # Analyze script for performance recommendations
        $recommendations = Analyze-ScriptPerformance -ScriptContent $scriptContent
        
        if ($recommendations.Count -gt 0) {
            Write-Host "`nPerformance Improvement Recommendations:" -ForegroundColor Green
            
            foreach ($recommendation in $recommendations) {
                Write-Host "`n[$($recommendation.Impact) Impact] $($recommendation.Issue)" -ForegroundColor Yellow
                Write-Host "Recommendation: $($recommendation.Recommendation)" -ForegroundColor Cyan
            }
        } else {
            Write-Host "`nNo specific performance improvement recommendations." -ForegroundColor Green
        }
    } catch {
        Write-Warning "Cannot execute script as a script block. Running static analysis only."
        
        # Run static analysis only
        $recommendations = Analyze-ScriptPerformance -ScriptContent $scriptContent
        
        if ($recommendations.Count -gt 0) {
            Write-Host "`nPerformance Improvement Recommendations:" -ForegroundColor Green
            
            foreach ($recommendation in $recommendations) {
                Write-Host "`n[$($recommendation.Impact) Impact] $($recommendation.Issue)" -ForegroundColor Yellow
                Write-Host "Recommendation: $($recommendation.Recommendation)" -ForegroundColor Cyan
            }
        } else {
            Write-Host "`nNo specific performance improvement recommendations." -ForegroundColor Green
        }
    }
} catch {
    Write-Error "Error analyzing script performance: $_"
}
