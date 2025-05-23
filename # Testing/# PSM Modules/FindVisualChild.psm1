# Check if Write-DebugMessage exists, if not create a stub function
if (-not (Get-Command -Name "Write-DebugMessage" -ErrorAction SilentlyContinue)) {
    function Write-DebugMessage { param([string]$Message) Write-Verbose $Message }
}

function FindVisualChild {
    param(
        [Parameter(Mandatory=$true)]
        [System.Windows.DependencyObject]$parent,
        
        [Parameter(Mandatory=$true)]
        [type]$childType
    )
    
    for ($i = 0; $i -lt [System.Windows.Media.VisualTreeHelper]::GetChildrenCount($parent); $i++) {
        $child = [System.Windows.Media.VisualTreeHelper]::GetChild($parent, $i)
        
        if ($child -and $child -is $childType) {
            return $child
        } else {
            $result = FindVisualChild -parent $child -childType $childType
            if ($result) {
                return $result
            }
        }
    }
    
    return $null
}

Export-ModuleMember -Function FindVisualChild
