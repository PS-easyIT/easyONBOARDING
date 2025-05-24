# WPF-Diagnose-Tool

$ModulesDir = Join-Path -Path $PSScriptRoot -ChildPath 'Modules'
$XamlPath = Join-Path -Path $PSScriptRoot -ChildPath 'MainGUI.xaml'

# Benötigte Assemblies laden
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase

# Debug-Modul importieren
Import-Module "$ModulesDir\WPFDebugHelper.psm1" -Force

Write-Host "Starting WPF diagnostic..." -ForegroundColor Cyan
Write-Host "PowerShell Version: $($PSVersionTable.PSVersion)" -ForegroundColor Cyan
Write-Host "Working directory: $PSScriptRoot" -ForegroundColor Cyan
Write-Host "XAML file path: $XamlPath" -ForegroundColor Cyan

# 1. WPF-Umgebung testen
$wpfTest = Test-WPFLoading
if (-not $wpfTest) {
    Write-Host "WPF environment test failed! Cannot continue with diagnostics." -ForegroundColor Red
    exit
}

# 2. XAML-Datei analysieren
Write-Host "`nAnalyzing XAML file..." -ForegroundColor Cyan
$xamlTest = Debug-XAMLFile -XamlPath $XamlPath
if (-not $xamlTest) {
    Write-Host "XAML file has errors. Attempting to create a fixed version..." -ForegroundColor Yellow
    Fix-CommonXAMLIssues -XamlPath $XamlPath -OutputPath "$XamlPath.fixed"
    Write-Host "Testing fixed XAML file..." -ForegroundColor Cyan
    Debug-XAMLFile -XamlPath "$XamlPath.fixed"
}

# 3. Minimales XAML testen
Write-Host "`nCreating and testing minimal XAML window..." -ForegroundColor Cyan
$minimalXaml = @"
<Window 
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="Minimal Test Window" 
    Width="400" 
    Height="300">
    <Grid>
        <TextBlock Text="If you can see this window, WPF is working correctly!" 
                  HorizontalAlignment="Center" 
                  VerticalAlignment="Center"
                  FontSize="16" />
    </Grid>
</Window>
"@

try {
    $reader = [System.Xml.XmlReader]::Create([System.IO.StringReader]::new($minimalXaml))
    $testWindow = [System.Windows.Markup.XamlReader]::Load($reader)
    
    Write-Host "Minimal XAML loaded successfully. Displaying test window..." -ForegroundColor Green
    $result = $testWindow.ShowDialog()
    Write-Host "Test window closed with result: $result" -ForegroundColor Cyan
}
catch {
    Write-Host "ERROR creating minimal test window: $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host "`nWPF Diagnostic completed." -ForegroundColor Cyan
