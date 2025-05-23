<#
.SYNOPSIS
    Scans PowerShell scripts for common errors and offers suggestions to fix them.

.DESCRIPTION
    This script analyzes PowerShell scripts to find syntax errors, best practice violations,
    and other common issues. It provides suggestions for fixing these issues.

.PARAMETER ScriptPath
    Path to the PowerShell script to be analyzed.

.PARAMETER FixErrors
    Switch to automatically apply recommended fixes where possible.

.EXAMPLE
    .\Check-ScriptErrors.ps1 -ScriptPath "C:\Scripts\MyScript.ps1"
    
.EXAMPLE
    .\Check-ScriptErrors.ps1 -ScriptPath "C:\Scripts\MyScript.ps1" -FixErrors
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory = $true)]
    [string]$ScriptPath,
    
    [Parameter(Mandatory = $false)]
    [switch]$FixErrors
)

function Test-ScriptSyntax {
    param (
        [string]$ScriptContent
    )
    
    try {
        [System.Management.Automation.PSParser]::Tokenize($ScriptContent, [ref]$null)
        return @{
            Success = $true
            Errors = $null
        }
    }
    catch {
        return @{
            Success = $false
            Errors = $_.Exception.Message
        }
    }
}

function Find-BestPracticeViolations {
    param (
        [string]$ScriptContent
    )
    
    $violations = @()
    
    # Check for Write-Host usage (prefer Write-Output)
    if ($ScriptContent -match "Write-Host ") {
        $violations += @{
            Type = "BestPractice"
            Issue = "Use of Write-Host detected"
            Recommendation = "Consider using Write-Output instead of Write-Host for better pipeline support"
            Fixable = $true
            FixFunc = { param($content) $content -replace "Write-Host ", "Write-Output " }
        }
    }
    
    # Check for cmdlet aliases
    $commonAliases = @{
        "gci" = "Get-ChildItem"
        "ls" = "Get-ChildItem"
        "dir" = "Get-ChildItem"
        "gi" = "Get-Item"
        "cat" = "Get-Content"
        "gc" = "Get-Content"
        "type" = "Get-Content"
        "cd" = "Set-Location"
        "chdir" = "Set-Location"
        "sl" = "Set-Location"
        "cls" = "Clear-Host"
        "clear" = "Clear-Host"
        "cp" = "Copy-Item"
        "copy" = "Copy-Item"
        "mv" = "Move-Item"
        "move" = "Move-Item"
        "rm" = "Remove-Item"
        "del" = "Remove-Item"
        "rmdir" = "Remove-Item"
        "echo" = "Write-Output"
        "ft" = "Format-Table"
        "fw" = "Format-Wide"
        "fl" = "Format-List"
        "foreach" = "ForEach-Object"
        "%" = "ForEach-Object"
        "where" = "Where-Object"
        "?" = "Where-Object"
    }
    
    foreach ($alias in $commonAliases.Keys) {
        if ($ScriptContent -match "\s$alias\s") {
            $violations += @{
                Type = "BestPractice"
                Issue = "Use of alias '$alias' detected"
                Recommendation = "Use full cmdlet name '$($commonAliases[$alias])' instead of alias for better readability"
                Fixable = $true
                FixFunc = { param($content) $content -replace "\s$alias\s", " $($commonAliases[$alias]) " }
            }
        }
    }
    
    # Check for unintialized variables
    $matches = [regex]::Matches($ScriptContent, '\$\w+\s')
    $declaredVars = @()
    
    foreach ($match in $matches) {
        $varName = $match.Value.Trim()
        if (($varName -ne '$_') -and ($varName -ne '$null') -and ($varName -ne '$true') -and ($varName -ne '$false') -and ($declaredVars -notcontains $varName)) {
            $declaredVars += $varName
            
            if ($ScriptContent -notmatch "\`$varName\s*=") {
                $violations += @{
                    Type = "Potential Issue"
                    Issue = "Potential use of uninitialized variable: $varName"
                    Recommendation = "Ensure $varName is initialized before use or use `$null initializer"
                    Fixable = $false
                }
            }
        }
    }
    
    return $violations
}

function Fix-ScriptIssues {
    param (
        [string]$ScriptContent,
        [array]$Violations
    )
    
    $fixedContent = $ScriptContent
    
    foreach ($violation in $Violations) {
        if ($violation.Fixable) {
            $fixedContent = & $violation.FixFunc $fixedContent
        }
    }
    
    return $fixedContent
}

# Main script execution
try {
    if (-not (Test-Path -Path $ScriptPath)) {
        Write-Error "Script not found: $ScriptPath"
        return
    }
    
    $scriptContent = Get-Content -Path $ScriptPath -Raw
    $syntaxCheck = Test-ScriptSyntax -ScriptContent $scriptContent
    
    Write-Host "Analyzing script: $ScriptPath" -ForegroundColor Cyan
    Write-Host "----------------------------------------" -ForegroundColor Cyan
    
    if (-not $syntaxCheck.Success) {
        Write-Host "Syntax Errors Detected:" -ForegroundColor Red
        Write-Host $syntaxCheck.Errors -ForegroundColor Red
    } else {
        Write-Host "Syntax Check: Passed" -ForegroundColor Green
        
        $bestPracticeViolations = Find-BestPracticeViolations -ScriptContent $scriptContent
        
        if ($bestPracticeViolations.Count -eq 0) {
            Write-Host "Best Practice Check: No issues found" -ForegroundColor Green
        } else {
            Write-Host "Issues Found: $($bestPracticeViolations.Count)" -ForegroundColor Yellow
            
            foreach ($violation in $bestPracticeViolations) {
                Write-Host "`n[$($violation.Type)]: $($violation.Issue)" -ForegroundColor Yellow
                Write-Host "Recommendation: $($violation.Recommendation)" -ForegroundColor Cyan
                if ($violation.Fixable) {
                    Write-Host "Automatically fixable: Yes" -ForegroundColor Green
                } else {
                    Write-Host "Automatically fixable: No (manual review required)" -ForegroundColor DarkYellow
                }
            }
            
            if ($FixErrors) {
                $fixedContent = Fix-ScriptIssues -ScriptContent $scriptContent -Violations $bestPracticeViolations
                $backupPath = "$ScriptPath.backup"
                
                Copy-Item -Path $ScriptPath -Destination $backupPath
                Set-Content -Path $ScriptPath -Value $fixedContent
                
                Write-Host "`nApplied fixes to script. Original backed up to: $backupPath" -ForegroundColor Green
            } else {
                Write-Host "`nRun with -FixErrors to automatically apply recommended fixes." -ForegroundColor Cyan
            }
        }
    }
} catch {
    Write-Error "Error analyzing script: $_"
}
