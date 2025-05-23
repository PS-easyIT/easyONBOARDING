function Get-ModuleStatus {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $false)]
        [string]$ModulesDirectory = (Join-Path -Path (Split-Path -Parent $MyInvocation.MyCommand.Definition) -ChildPath 'Modules'),
        
        [Parameter(Mandatory = $false)]
        [switch]$IncludeLoadedModules,
        
        [Parameter(Mandatory = $false)]
        [switch]$ExportToCsv
    )
    
    Write-Host "Checking module status in: $ModulesDirectory" -ForegroundColor Cyan
    
    # Get all PSM1 files in the modules directory
    $moduleFiles = Get-ChildItem -Path $ModulesDirectory -Filter "*.psm1" -ErrorAction SilentlyContinue
    
    if (-not $moduleFiles -or $moduleFiles.Count -eq 0) {
        Write-Warning "No .psm1 modules found in $ModulesDirectory"
        return
    }
    
    $results = @()
    
    # Check each module
    foreach ($moduleFile in $moduleFiles) {
        $moduleName = $moduleFile.BaseName
        $loaded = $false
        $loadError = $null
        
        try {
            # Check if module is loaded
            $loadedModule = Get-Module -Name $moduleName -ErrorAction SilentlyContinue
            $loaded = $null -ne $loadedModule
            
            # Try to import if requested and not already loaded
            if ($IncludeLoadedModules -and -not $loaded) {
                Import-Module $moduleFile.FullName -Force -ErrorAction Stop
                $loaded = $true
            }
        }
        catch {
            $loadError = $_.Exception.Message
        }
        
        $results += [PSCustomObject]@{
            ModuleName = $moduleName
            FilePath = $moduleFile.FullName
            Loaded = $loaded
            Error = $loadError
            FileSize = $moduleFile.Length
            LastModified = $moduleFile.LastWriteTime
        }
    }
    
    # Check for common module dependencies
    $commonDependencies = @(
        @{Name = "ActiveDirectory"; Type = "Windows Feature Module"},
        @{Name = "Microsoft.PowerShell.Management"; Type = "Built-in Module"},
        @{Name = "PresentationFramework"; Type = "Assembly"}
    )
    
    foreach ($dep in $commonDependencies) {
        $depLoaded = $false
        $loadError = $null
        
        try {
            if ($dep.Type -eq "Assembly") {
                $depLoaded = [System.Reflection.Assembly]::LoadWithPartialName($dep.Name) -ne $null
            } else {
                $depLoaded = $null -ne (Get-Module -Name $dep.Name -ListAvailable -ErrorAction SilentlyContinue)
            }
        }
        catch {
            $loadError = $_.Exception.Message
        }
        
        $results += [PSCustomObject]@{
            ModuleName = "$($dep.Name) ($($dep.Type))"
            FilePath = "System"
            Loaded = $depLoaded
            Error = $loadError
            FileSize = $null
            LastModified = $null
        }
    }
    
    # Output results
    $results | Format-Table -AutoSize
    
    # Calculate stats
    $loadedCount = ($results | Where-Object { $_.Loaded -eq $true }).Count
    $totalCount = $results.Count
    $percentLoaded = [math]::Round(($loadedCount / $totalCount) * 100, 1)
    
    Write-Host "Module Status Summary:" -ForegroundColor Cyan
    Write-Host "  Total Modules: $totalCount" -ForegroundColor White
    Write-Host "  Loaded Modules: $loadedCount" -ForegroundColor $(if ($loadedCount -eq $totalCount) { "Green" } else { "Yellow" })
    Write-Host "  Loading Rate: $percentLoaded%" -ForegroundColor $(if ($percentLoaded -gt 90) { "Green" } elseif ($percentLoaded -gt 70) { "Yellow" } else { "Red" })
    
    # Export to CSV if requested
    if ($ExportToCsv) {
        $csvPath = Join-Path -Path (Split-Path -Parent $ModulesDirectory) -ChildPath "ModuleStatus_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
        $results | Export-Csv -Path $csvPath -NoTypeInformation
        Write-Host "Exported module status to: $csvPath" -ForegroundColor Green
    }
    
    return $results
}

function Test-ModuleDependencies {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $false)]
        [string]$ModulesDirectory = (Join-Path -Path (Split-Path -Parent $MyInvocation.MyCommand.Definition) -ChildPath 'Modules')
    )
    
    Write-Host "Analyzing module dependencies in: $ModulesDirectory" -ForegroundColor Cyan
    
    # Get all PSM1 files
    $moduleFiles = Get-ChildItem -Path $ModulesDirectory -Filter "*.psm1" -ErrorAction SilentlyContinue
    
    $dependencies = @{}
    
    foreach ($moduleFile in $moduleFiles) {
        $content = Get-Content -Path $moduleFile.FullName -Raw
        
        # Look for Import-Module statements
        $importMatches = [regex]::Matches($content, 'Import-Module\s+([''"](?<Path>.*?)[''"]|(?<Module>\S+))')
        
        if ($importMatches.Count -gt 0) {
            $moduleImports = @()
            
            foreach ($match in $importMatches) {
                $modulePath = $match.Groups['Path'].Value
                $moduleName = $match.Groups['Module'].Value
                
                if (-not [string]::IsNullOrEmpty($modulePath)) {
                    # Extract module name from path
                    $moduleNameFromPath = [System.IO.Path]::GetFileNameWithoutExtension($modulePath)
                    $moduleImports += $moduleNameFromPath
                }
                elseif (-not [string]::IsNullOrEmpty($moduleName)) {
                    $moduleImports += $moduleName
                }
            }
            
            if ($moduleImports.Count -gt 0) {
                $dependencies[$moduleFile.BaseName] = $moduleImports
            }
        }
    }
    
    # Display dependency tree
    Write-Host "Module Dependencies:" -ForegroundColor Yellow
    foreach ($module in $dependencies.Keys) {
        Write-Host "  $module depends on:" -ForegroundColor White
        foreach ($dependency in $dependencies[$module]) {
            $exists = Test-Path (Join-Path -Path $ModulesDirectory -ChildPath "$dependency.psm1")
            Write-Host "    - $dependency" -ForegroundColor $(if ($exists) { "Green" } else { "Red" })
        }
    }
    
    # Find circular dependencies
    $circularDeps = Find-CircularDependencies -DependencyMap $dependencies
    
    if ($circularDeps.Count -gt 0) {
        Write-Host "Warning: Circular dependencies detected:" -ForegroundColor Red
        foreach ($chain in $circularDeps) {
            Write-Host "  $($chain -join ' -> ')" -ForegroundColor Red
        }
    }
    
    return @{
        Dependencies = $dependencies
        CircularDependencies = $circularDeps
    }
}

function Find-CircularDependencies {
    param (
        [Parameter(Mandatory = $true)]
        [hashtable]$DependencyMap
    )
    
    $circularChains = @()
    
    foreach ($module in $DependencyMap.Keys) {
        $visited = @{}
        $path = @()
        
        function DFS {
            param (
                [Parameter(Mandatory = $true)]
                [string]$CurrentModule,
                
                [Parameter(Mandatory = $true)]
                [hashtable]$Visited,
                
                [Parameter(Mandatory = $true)]
                [array]$Path
            )
            
            $Visited[$CurrentModule] = $true
            $Path += $CurrentModule
            
            if ($DependencyMap.ContainsKey($CurrentModule)) {
                foreach ($dep in $DependencyMap[$CurrentModule]) {
                    if ($Visited.ContainsKey($dep)) {
                        # Circular dependency found
                        $startIndex = $Path.IndexOf($dep)
                        if ($startIndex -ge 0) {
                            $cycle = $Path[$startIndex..($Path.Count-1)] + $dep
                            $cycleStr = $cycle -join ' -> '
                            
                            # Add to results if not already present
                            if ($circularChains -notcontains $cycleStr) {
                                $circularChains += $cycleStr
                            }
                        }
                    }
                    elseif ($DependencyMap.ContainsKey($dep)) {
                        DFS -CurrentModule $dep -Visited $Visited -Path $Path
                    }
                }
            }
            
            # Backtrack
            $Visited.Remove($CurrentModule)
            $Path = $Path[0..($Path.Count-2)]
        }
        
        DFS -CurrentModule $module -Visited $visited -Path $path
    }
    
    return $circularChains
}

# Export functions
Export-ModuleMember -Function Get-ModuleStatus, Test-ModuleDependencies
