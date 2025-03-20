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
