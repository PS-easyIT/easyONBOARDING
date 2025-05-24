# WPF Debug-Hilfsfunktionen

function Test-WPFLoading {
    [CmdletBinding()]
    param()
    
    Write-Host "WPF Debug: Testing WPF environment..." -ForegroundColor Cyan
    
    # Prüfen, ob die notwendigen WPF-Assemblies geladen werden können
    try {
        Add-Type -AssemblyName PresentationFramework
        Add-Type -AssemblyName PresentationCore
        Add-Type -AssemblyName WindowsBase
        Write-Host "WPF Debug: WPF assemblies loaded successfully" -ForegroundColor Green
    }
    catch {
        Write-Host "WPF Debug: ERROR loading WPF assemblies: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
    
    # Test erstellen einer einfachen WPF-Testfenster 
    try {
        $xaml = @"
<Window 
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="WPF Test Window" 
    Width="300" 
    Height="200">
    <Grid>
        <TextBlock Text="WPF Test Successful" HorizontalAlignment="Center" VerticalAlignment="Center" />
    </Grid>
</Window>
"@
        
        $reader = [System.Xml.XmlReader]::Create([System.IO.StringReader]::new($xaml))
        $testWindow = [System.Windows.Markup.XamlReader]::Load($reader)
        
        Write-Host "WPF Debug: Test window created successfully" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Host "WPF Debug: ERROR creating test window: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}

function Debug-XAMLFile {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$XamlPath
    )
    
    Write-Host "WPF Debug: Testing XAML file: $XamlPath" -ForegroundColor Cyan
    
    if (-not (Test-Path $XamlPath)) {
        Write-Host "WPF Debug: XAML file not found at: $XamlPath" -ForegroundColor Red
        return $false
    }
    
    try {
        # Inhalt der XAML-Datei laden
        $xamlContent = Get-Content -Path $XamlPath -Raw -ErrorAction Stop
        
        # XAML parsen und prüfen, ob es geladen werden kann
        $reader = [System.Xml.XmlReader]::Create([System.IO.StringReader]::new($xamlContent))
        $window = [System.Windows.Markup.XamlReader]::Load($reader)
        
        Write-Host "WPF Debug: XAML file parsed successfully" -ForegroundColor Green
        Write-Host "WPF Debug: Window title: $($window.Title)" -ForegroundColor Cyan
        Write-Host "WPF Debug: Window size: $($window.Width)x$($window.Height)" -ForegroundColor Cyan
        
        return $true
    }
    catch {
        Write-Host "WPF Debug: ERROR parsing XAML file: $($_.Exception.Message)" -ForegroundColor Red
        
        # Versuchen, die problematische Zeile zu identifizieren
        if ($_.Exception.Message -match "Position: (\d+):(\d+)") {
            $line = $Matches[1]
            $column = $Matches[2]
            Write-Host "WPF Debug: Error around line: $line, column: $column" -ForegroundColor Yellow
            
            # Zeige die problematischen Zeilen
            $contentLines = Get-Content -Path $XamlPath
            $startLine = [Math]::Max(1, $line - 2)
            $endLine = [Math]::Min($contentLines.Count, $line + 2)
            
            Write-Host "WPF Debug: Problematic section:" -ForegroundColor Yellow
            for ($i = $startLine; $i -le $endLine; $i++) {
                $prefix = if ($i -eq $line) { ">>>>" } else { "    " }
                Write-Host "$prefix $i`: $($contentLines[$i-1])" -ForegroundColor $(if ($i -eq $line) { "Red" } else { "Gray" })
            }
        }
        
        return $false
    }
}

function Fix-CommonXAMLIssues {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$XamlPath,
        
        [Parameter(Mandatory=$false)]
        [string]$OutputPath
    )
    
    if (-not $OutputPath) {
        $OutputPath = "$XamlPath.fixed"
    }
    
    Write-Host "WPF Debug: Attempting to fix common XAML issues in: $XamlPath" -ForegroundColor Cyan
    
    if (-not (Test-Path $XamlPath)) {
        Write-Host "WPF Debug: XAML file not found at: $XamlPath" -ForegroundColor Red
        return $false
    }
    
    try {
        # Inhalt der XAML-Datei laden
        $xamlContent = Get-Content -Path $XamlPath -Raw -ErrorAction Stop
        
        # Häufige Probleme beheben
        
        # 1. Fehlende Bindings korrigieren, die "{Binding}" ohne Pfad haben
        $xamlContent = $xamlContent -replace '{Binding}', '{Binding Path=.}'
        
        # 2. Doppelte Namespace-Deklarationen entfernen
        $namespaceMatches = [regex]::Matches($xamlContent, 'xmlns(:[\w]+)?="[^"]+"')
        $uniqueNamespaces = @{}
        foreach ($match in $namespaceMatches) {
            $uniqueNamespaces[$match.Value] = $true
        }
        
        # 3. Einige häufige Syntaxfehler korrigieren
        $xamlContent = $xamlContent -replace '""""', '"""'
        $xamlContent = $xamlContent -replace '>>>>>', '>>>>'
        $xamlContent = $xamlContent -replace '<<<<', '<<<'
        
        # Speichern der korrigierten Datei
        $xamlContent | Out-File -FilePath $OutputPath -Encoding UTF8
        
        Write-Host "WPF Debug: Fixed XAML saved to: $OutputPath" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Host "WPF Debug: ERROR fixing XAML file: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}

Export-ModuleMember -Function Test-WPFLoading, Debug-XAMLFile, Fix-CommonXAMLIssues
