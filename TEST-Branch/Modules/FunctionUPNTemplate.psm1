# Functions to replace placeholders in templates with user data
Write-DebugMessage "Resolving template placeholders."
function Resolve-TemplatePlaceholders {
    # [12.1 - Generic placeholder replacement for string templates]
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Template,

        [Parameter(Mandatory = $true)]
        [ValidateNotNull()]
        [pscustomobject]$userData
    )
    # Replaces placeholders {first} and {last} with corresponding values from $userData
    $result = $Template -replace '{first}', $userData.FirstName `
                          -replace '{last}', $userData.LastName
    return $result
}

Write-DebugMessage "Function to replace placeholders - UPN"
# In the UPN part:
function Get-UPN {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNull()]
        [pscustomobject]$userData,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$SamAccountName,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$adDomain
    )

    # Normalize the template: trim and convert to lowercase
    if (-not [string]::IsNullOrWhiteSpace($userData.UPNFormat)) {
        $upnTemplate = $userData.UPNFormat.Trim().ToLower()
        Write-DebugMessage "Invoke-Onboarding: UPN Format from userData: $upnTemplate"
        if ($upnTemplate -like "*{first}*") {
            # Dynamic replacement of placeholders from the template
            $upnBase = Resolve-TemplatePlaceholders -Template $upnTemplate -userData $userData
            $UPN = "$upnBase$adDomain"
        }
        else {
            # Fixed cases as fallback – add more cases if needed
            switch ($upnTemplate) {
                "firstname.lastname"    { $UPN = "$($userData.FirstName).$($userData.LastName)$adDomain" }
                "f.lastname"            { $UPN = "$($userData.FirstName.Substring(0,1)).$($userData.LastName)$adDomain" }
                "firstnamelastname"     { $UPN = "$($userData.FirstName)$($userData.LastName)$adDomain" }
                "flastname"             { $UPN = "$($userData.FirstName.Substring(0,1))$($userData.LastName)$adDomain" }
                Default                 { $UPN = "$SamAccountName$adDomain" }
            }
        }
    }
    else {
        Write-DebugMessage "No UPNFormat given, fallback to SamAccountName + domain"
        $UPN = "$SamAccountName$adDomain"
    }
    return $UPN
}

# Export the Get-UPN function so it can be used by other modules/scripts
Export-ModuleMember -Function Get-UPN

Write-DebugMessage "Loading AD groups."

# Functions for managing UPN templates

function Initialize-UPNTemplateModule {
    [CmdletBinding()]
    param()
    
    try {
        if (Get-Command -Name "Write-DebugOutput" -ErrorAction SilentlyContinue) {
            Write-DebugOutput "Initializing UPN Template module" -Level INFO
        }
        
        # Make sure we have the required assemblies loaded
        if (-not ('System.Windows.Controls.ComboBox' -as [type])) {
            Add-Type -AssemblyName PresentationFramework -ErrorAction Stop
            Add-Type -AssemblyName PresentationCore -ErrorAction Stop
            Add-Type -AssemblyName WindowsBase -ErrorAction Stop
        }
        
        if (Get-Command -Name "Write-DebugOutput" -ErrorAction SilentlyContinue) {
            Write-DebugOutput "UPN Template module initialized successfully" -Level INFO
        }
        
        return $true
    }
    catch {
        if (Get-Command -Name "Write-DebugOutput" -ErrorAction SilentlyContinue) {
            Write-DebugOutput "Error initializing UPN Template module: $($_.Exception.Message)" -Level ERROR
        } else {
            Write-Host "Error initializing UPN Template module: $($_.Exception.Message)" -ForegroundColor Red
        }
        return $false
    }
}

function Format-UPNFromTemplate {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$FirstName,
        
        [Parameter(Mandatory=$true)]
        [string]$LastName,
        
        [Parameter(Mandatory=$false)]
        [string]$Template = "Vorname.Nachname",
        
        [Parameter(Mandatory=$false)]
        [string]$Domain = "example.com"
    )
    
    try {
        # Replace special characters and spaces
        $FirstName = $FirstName -replace '[^a-zA-Z0-9]', ''
        $LastName = $LastName -replace '[^a-zA-Z0-9]', ''
        
        # Convert to lowercase
        $FirstName = $FirstName.ToLower()
        $LastName = $LastName.ToLower()
        
        $upn = switch ($Template) {
            "Vorname.Nachname" { "$FirstName.$LastName" }
            "NachnameV" { "$LastName$($FirstName.Substring(0,1))" }
            "Nachname.Vorname" { "$LastName.$FirstName" }
            "VornameN" { "$FirstName$($LastName.Substring(0,1))" }
            default { "$FirstName.$LastName" }
        }
        
        return "$upn@$Domain"
    }
    catch {
        if (Get-Command -Name "Write-DebugOutput" -ErrorAction SilentlyContinue) {
            Write-DebugOutput "Error formatting UPN: $($_.Exception.Message)" -Level ERROR
        }
        return "$FirstName.$LastName@$Domain"
    }
}

function Get-UPNTemplates {
    [CmdletBinding()]
    param()
    
    return @(
        "Vorname.Nachname",
        "NachnameV",
        "Nachname.Vorname",
        "VornameN"
    )
}

function Update-UPNTemplateDisplay {
    [CmdletBinding()]
    param()
    
    try {
        $cmbDisplayTemplate = $global:window.FindName("cmbDisplayTemplate")
        if ($null -eq $cmbDisplayTemplate) {
            if (Get-Command -Name "Write-DebugOutput" -ErrorAction SilentlyContinue) {
                Write-DebugOutput "cmbDisplayTemplate control not found" -Level WARNING
            }
            return $false
        }
        
        $templates = Get-UPNTemplates
        $cmbDisplayTemplate.Items.Clear()
        foreach ($template in $templates) {
            $cmbDisplayTemplate.Items.Add($template)
        }
        
        if ($cmbDisplayTemplate.Items.Count -gt 0) {
            $cmbDisplayTemplate.SelectedIndex = 0
        }
        
        if (Get-Command -Name "Write-DebugOutput" -ErrorAction SilentlyContinue) {
            Write-DebugOutput "UPN template display updated" -Level INFO
        }
        return $true
    }
    catch {
        if (Get-Command -Name "Write-DebugOutput" -ErrorAction SilentlyContinue) {
            Write-DebugOutput "Error updating UPN template display: $($_.Exception.Message)" -Level ERROR
        }
        return $false
    }
}

# Export all functions
Export-ModuleMember -Function Initialize-UPNTemplateModule, Format-UPNFromTemplate, Get-UPNTemplates, Update-UPNTemplateDisplay
