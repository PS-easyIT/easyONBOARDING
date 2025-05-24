# Get script directory for relative path resolution
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
$easyONBDir = "Z:\OD_PowerShell\easyONBOARDING\V1.4.1"
$ModulesDir = Join-Path -Path $easyONBDir -ChildPath 'Modules'

# Import the ModuleStatusChecker module
$moduleCheckerPath = Join-Path -Path "Z:\OD_PowerShell\easyONBOARDING\V1.4.1\Modules" -ChildPath "ModuleStatusChecker.psm1"
if (Test-Path $moduleCheckerPath) {
    Import-Module $moduleCheckerPath -Force
    Write-Host "ModuleStatusChecker module imported successfully" -ForegroundColor Green
} else {
    Write-Host "ModuleStatusChecker module not found at $moduleCheckerPath" -ForegroundColor Red
    Write-Host "Please create the module first using the provided code" -ForegroundColor Yellow
    exit
}

# Display header
Write-Host "===========================================" -ForegroundColor Cyan
Write-Host "=   easyONBOARDING Module Health Check   =" -ForegroundColor Cyan
Write-Host "===========================================" -ForegroundColor Cyan
Write-Host "Script directory: $ScriptDir"
Write-Host "easyONBOARDING directory: $easyONBDir"
Write-Host "Modules directory: $ModulesDir"
Write-Host

# Check directory exists
if (-not (Test-Path $ModulesDir)) {
    Write-Host "ERROR: Modules directory not found at: $ModulesDir" -ForegroundColor Red
    exit
}

# Run module status check
Write-Host "Checking module status..." -ForegroundColor Yellow
$moduleStatus = Get-ModuleStatus -ModulesDirectory $ModulesDir -ExportToCsv

# Analyze dependencies
Write-Host "`nAnalyzing module dependencies..." -ForegroundColor Yellow
$dependencies = Test-ModuleDependencies -ModulesDirectory $ModulesDir

# Check for common issues
Write-Host "`nChecking for common issues..." -ForegroundColor Yellow
$missingModules = $moduleStatus | Where-Object { $_.Loaded -eq $false }
if ($missingModules.Count -gt 0) {
    Write-Host "WARNING: The following modules failed to load:" -ForegroundColor Red
    foreach ($module in $missingModules) {
        Write-Host "  - $($module.ModuleName): $($module.Error)" -ForegroundColor Red
    }
    
    Write-Host "`nTROUBLESHOOTING TIPS:" -ForegroundColor Yellow
    Write-Host "1. Check that the module files exist in the Modules directory"
    Write-Host "2. Verify you have the necessary permissions to access the files"
    Write-Host "3. Look for syntax errors in the module files"
    Write-Host "4. Check if required dependencies are installed"
}

# Check XAML loading specifically (since output showed it cut off there)
$xamlModule = $moduleStatus | Where-Object { $_.ModuleName -eq "LoadConfigXAML" }
if ($xamlModule) {
    Write-Host "`nXAML Module Status:" -ForegroundColor Cyan
    if ($xamlModule.Loaded) {
        Write-Host "  The XAML configuration module loaded successfully." -ForegroundColor Green
    } else {
        Write-Host "  WARNING: The XAML configuration module failed to load!" -ForegroundColor Red
        Write-Host "  Error: $($xamlModule.Error)" -ForegroundColor Red
        Write-Host "`n  POSSIBLE SOLUTIONS:" -ForegroundColor Yellow
        Write-Host "  1. Check that the MainGUI.xaml file exists at: $easyONBDir\MainGUI.xaml"
        Write-Host "  2. Verify the XAML file is valid and doesn't contain syntax errors"
        Write-Host "  3. Check that the PresentationFramework assembly is loaded correctly"
    }
}

Write-Host "`nHealth check completed. See above for any issues that need addressing." -ForegroundColor Cyan
