# Check if Write-DebugMessage exists, if not create a stub function
if (-not (Get-Command -Name "Write-DebugMessage" -ErrorAction SilentlyContinue)) {
    function Write-DebugMessage { 
        [CmdletBinding()]
        param(
            [Parameter(Mandatory=$true)]
            [string]$Message,
            [string]$LogLevel = "DEBUG"
        )
        Write-Verbose $Message 
    }
}

function FindVisualChild {
    [CmdletBinding(DefaultParameterSetName='ByType')]
    param(
        [Parameter(Mandatory=$true, Position=0)]
        [System.Windows.DependencyObject]$Parent,
        
        [Parameter(Mandatory=$true, ParameterSetName='ByType', Position=1)]
        [type]$ChildType,
        
        [Parameter(Mandatory=$true, ParameterSetName='ByName', Position=1)]
        [string]$Name
    )
    
    try {
        if ($PSCmdlet.ParameterSetName -eq 'ByType') {
            Write-DebugMessage "Searching for child of type $($ChildType.Name) in visual tree"
        } else {
            Write-DebugMessage "Searching for child with name '$Name' in visual tree"
        }
        
        for ($i = 0; $i -lt [System.Windows.Media.VisualTreeHelper]::GetChildrenCount($Parent); $i++) {
            $child = [System.Windows.Media.VisualTreeHelper]::GetChild($Parent, $i)
            
            if ($null -eq $child) {
                continue
            }
            
            # Check if child matches our criteria
            if ($PSCmdlet.ParameterSetName -eq 'ByType' -and $child -is $ChildType) {
                Write-DebugMessage "Found child of type $($ChildType.Name)"
                return $child
            } 
            elseif ($PSCmdlet.ParameterSetName -eq 'ByName' -and 
                    $child -is [System.Windows.FrameworkElement] -and 
                    $child.Name -eq $Name) {
                Write-DebugMessage "Found child with name '$Name'"
                return $child
            }
            
            # Recursive search
            $result = $null
            if ($PSCmdlet.ParameterSetName -eq 'ByType') {
                $result = FindVisualChild -Parent $child -ChildType $ChildType
            } else {
                $result = FindVisualChild -Parent $child -Name $Name
            }
            
            if ($null -ne $result) {
                return $result
            }
        }
        
        return $null
    }
    catch {
        Write-DebugMessage "Error searching for visual child: $($_.Exception.Message)"
        return $null
    }
}

# Export the function
Export-ModuleMember -Function FindVisualChild
