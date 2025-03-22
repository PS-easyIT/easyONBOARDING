# Functions for executing Active Directory commands

function Get-ADUsers {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$false)]
        [string]$Filter = "*",
        
        [Parameter(Mandatory=$false)]
        [int]$ResultLimit = 100,
        
        [Parameter(Mandatory=$false)]
        [string[]]$Properties = @("DisplayName", "SamAccountName", "UserPrincipalName", "Mail", "Enabled")
    )
    
    try {
        if (Get-Command -Name "Write-DebugOutput" -ErrorAction SilentlyContinue) {
            Write-DebugOutput "Getting AD users with filter: $Filter" -Level INFO
        }
        
        return Get-ADUser -Filter $Filter -Properties $Properties -ResultSetSize $ResultLimit
    }
    catch {
        if (Get-Command -Name "Write-DebugOutput" -ErrorAction SilentlyContinue) {
            Write-DebugOutput "Error getting AD users: $($_.Exception.Message)" -Level ERROR
        }
        return $null
    }
}

function Get-ADGroups {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$false)]
        [string]$Filter = "*",
        
        [Parameter(Mandatory=$false)]
        [int]$ResultLimit = 200,
        
        [Parameter(Mandatory=$false)]
        [string[]]$Properties = @("Name", "Description", "GroupCategory", "GroupScope")
    )
    
    try {
        if (Get-Command -Name "Write-DebugOutput" -ErrorAction SilentlyContinue) {
            Write-DebugOutput "Getting AD groups with filter: $Filter" -Level INFO
        }
        
        return Get-ADGroup -Filter $Filter -Properties $Properties -ResultSetSize $ResultLimit
    }
    catch {
        if (Get-Command -Name "Write-DebugOutput" -ErrorAction SilentlyContinue) {
            Write-DebugOutput "Error getting AD groups: $($_.Exception.Message)" -Level ERROR
        }
        return $null
    }
}

function Get-ADOrganizationalUnits {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$false)]
        [string]$Filter = "*",
        
        [Parameter(Mandatory=$false)]
        [int]$ResultLimit = 100,
        
        [Parameter(Mandatory=$false)]
        [string[]]$Properties = @("Name", "DistinguishedName")
    )
    
    try {
        if (Get-Command -Name "Write-DebugOutput" -ErrorAction SilentlyContinue) {
            Write-DebugOutput "Getting AD OUs with filter: $Filter" -Level INFO
        }
        
        return Get-ADOrganizationalUnit -Filter $Filter -Properties $Properties
    }
    catch {
        if (Get-Command -Name "Write-DebugOutput" -ErrorAction SilentlyContinue) {
            Write-DebugOutput "Error getting AD OUs: $($_.Exception.Message)" -Level ERROR
        }
        return $null
    }
}

function Test-SamAccountNameAvailable {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$SamAccountName
    )
    
    try {
        if (Get-Command -Name "Write-DebugOutput" -ErrorAction SilentlyContinue) {
            Write-DebugOutput "Checking if SamAccountName is available: $SamAccountName" -Level INFO
        }
        
        $user = Get-ADUser -Filter "SamAccountName -eq '$SamAccountName'" -ErrorAction SilentlyContinue
        
        if ($null -eq $user) {
            if (Get-Command -Name "Write-DebugOutput" -ErrorAction SilentlyContinue) {
                Write-DebugOutput "SamAccountName is available: $SamAccountName" -Level INFO
            }
            return $true
        } else {
            if (Get-Command -Name "Write-DebugOutput" -ErrorAction SilentlyContinue) {
                Write-DebugOutput "SamAccountName is already in use: $SamAccountName" -Level WARNING
            }
            return $false
        }
    }
    catch {
        if (Get-Command -Name "Write-DebugOutput" -ErrorAction SilentlyContinue) {
            Write-DebugOutput "Error checking SamAccountName: $($_.Exception.Message)" -Level ERROR
        }
        return $false
    }
}

# Export all functions
Export-ModuleMember -Function Get-ADUsers, Get-ADGroups, Get-ADOrganizationalUnits, Test-SamAccountNameAvailable