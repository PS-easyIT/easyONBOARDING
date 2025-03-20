#region [Region 08 | UPN GENERATION]
# Creates user principal names based on templates and user data
Write-DebugMessage "Generating UPN for user."
function New-UPN {
    param(
        [pscustomobject]$userData,
        [hashtable]$Config
    )

    # Check if FirstName and LastName are provided
    if ([string]::IsNullOrWhiteSpace($userData.FirstName) -or [string]::IsNullOrWhiteSpace($userData.LastName)) {
        Throw "Error: FirstName and LastName must be set!"
    }

    # 1) Generate SamAccountName (first letter of first name + entire last name, all lowercase)
    $SamAccountName = ''
    if ($userData.FirstName.Length -gt 0) {
        $SamAccountName = ($userData.FirstName.Substring(0,1) + $userData.LastName).ToLower()
    } else {
        $SamAccountName = $userData.LastName.ToLower()
    }
    Write-DebugMessage "SamAccountName= $SamAccountName"

    # 2) If a manual UPN was entered, use it immediately
    if (-not [string]::IsNullOrWhiteSpace($userData.UPNEntered)) {
        return @{
            SamAccountName = $SamAccountName
            UPN            = $userData.UPNEntered
            CompanySection = "Company"
        }
    }

    # 3) Determine the Company section (default: "Company")
    $companySection = "Company"  # Default value
    
    # Optimized check for CompanySection with simplified null value check
    if ($null -ne $userData.CompanySection) {
        # Case 1: CompanySection is directly a string
        if ($userData.CompanySection -is [string] -and -not [string]::IsNullOrWhiteSpace($userData.CompanySection)) {
            $companySection = $userData.CompanySection
        }
        # Case 2: CompanySection is an object with Section property
        elseif ($userData.CompanySection -is [PSObject] -and 
                $null -ne ($userData.CompanySection.PSObject.Properties.Match('Section') | Select-Object -First 1) -and
                -not [string]::IsNullOrWhiteSpace($userData.CompanySection.Section)) {
            $companySection = $userData.CompanySection.Section
        }
    }
    Write-DebugMessage "Using Company section: '$companySection'"

    # 4) Check if the desired section exists in the INI with improved error handling
    if ($null -eq $Config) {
        Throw "Error: Config object is NULL! No configuration available."
    }
    
    if (-not $Config.Contains($companySection)) {
        Write-DebugMessage "Section '$companySection' not found in Config, using 'Company'"
        $companySection = "Company" # Fallback to default value
        if (-not $Config.Contains($companySection)) {
            Throw "Error: Neither the requested section nor the default section 'Company' exists in the INI!"
        }
    }
    $companyData = $Config[$companySection]
    $suffix = ($companySection -replace "\D","")

    # 5) Determine domain key and read domain - with improved null checks
    $domainKey = "CompanyActiveDirectoryDomain$suffix"
    $fallbackDomainKey = "CompanyActiveDirectoryDomain"
    
    # Check for the specific key first, then fallback - with more precise checks
    if ($null -ne $companyData -and 
        $companyData.Contains($domainKey) -and 
        -not [string]::IsNullOrWhiteSpace($companyData[$domainKey])) {
        $adDomain = "@" + $companyData[$domainKey].Trim()
    }
    elseif ($null -ne $companyData -and 
            $companyData.Contains($fallbackDomainKey) -and 
            -not [string]::IsNullOrWhiteSpace($companyData[$fallbackDomainKey])) {
        $adDomain = "@" + $companyData[$fallbackDomainKey].Trim()
        Write-DebugMessage "Using fallback domain: $adDomain"
    }
    else {
        # Extended error message with more details about the problem
        $availableKeys = $null -ne $companyData ? ($companyData.Keys -join ", ") : "No keys available"
        Throw "Error: Domain information missing in the INI! Neither '$domainKey' nor '$fallbackDomainKey' found or valid. Available keys: $availableKeys"
    }

    # 6) Always use UPN template from INI (no GUI dropdown) - with optimized NULL check
    $upnTemplate = "FIRSTNAME.LASTNAME" # Safe default value
    
    # PowerShell 7 ?:-operator for NULL check and improved readability
    $displayNameTemplates = $null -ne $Config -and $Config.Contains("DisplayNameUPNTemplates") ? 
        $Config.DisplayNameUPNTemplates : $null
    
    if ($null -ne $displayNameTemplates -and 
        $displayNameTemplates.Contains("DefaultUserPrincipalNameFormat") -and 
        -not [string]::IsNullOrWhiteSpace($displayNameTemplates["DefaultUserPrincipalNameFormat"])) {
        $upnTemplate = $displayNameTemplates["DefaultUserPrincipalNameFormat"].ToUpper()
        Write-DebugMessage "UPN template loaded from INI: $upnTemplate"
    }
    else {
        Write-DebugMessage "No valid UPN template found in INI, using default '$upnTemplate'"
    }
    
    Write-DebugMessage "UPN template (from INI): $upnTemplate"

    # 7) Generate UPN based on template - with safe substring processing and null checks
    # Make sure FirstName and LastName are not null
    $firstName = if ($null -eq $userData.FirstName) { "" } else { $userData.FirstName }
    $lastName = if ($null -eq $userData.LastName) { "" } else { $userData.LastName }
    
    $UPN = ""
    switch ($upnTemplate) {
        "FIRSTNAME.LASTNAME"    { $UPN = "$firstName.$lastName".ToLower() + $adDomain }
        "LASTNAME.FIRSTNAME"    { $UPN = "$lastName.$firstName".ToLower() + $adDomain }
        "FIRSTINITIAL.LASTNAME" { 
            if (-not [string]::IsNullOrEmpty($firstName) -and $firstName.Length -gt 0) {
                $UPN = "$($firstName.Substring(0,1)).$lastName".ToLower() + $adDomain
            } else {
                $UPN = "x.$lastName".ToLower() + $adDomain
            }
        }
        "FIRSTNAME.LASTINITIAL" { 
            if (-not [string]::IsNullOrEmpty($lastName)) {
                $UPN = "$firstName.$($lastName.Substring(0,1))".ToLower() + $adDomain
            } else {
                $UPN = "$firstName.x".ToLower() + $adDomain
            }
        }
        "FIRSTNAME_LASTNAME"    { $UPN = "$firstName_$lastName".ToLower() + $adDomain }
        "LASTNAME_FIRSTNAME"    { $UPN = "$lastName_$firstName".ToLower() + $adDomain }
        "FIRSTINITIAL_LASTNAME" { 
            if (-not [string]::IsNullOrWhiteSpace($firstName)) {
                $UPN = "$($firstName.Substring(0,1))_$lastName".ToLower() + $adDomain
            } else {
                $UPN = "x_$lastName".ToLower() + $adDomain
            }
        }
        "FIRSTNAME_LASTINITIAL" { 
            if (-not [string]::IsNullOrWhiteSpace($lastName)) {
                $UPN = "$firstName_$($lastName.Substring(0,1))".ToLower() + $adDomain
            } else {
                $UPN = "$firstName_x".ToLower() + $adDomain
            }
        }
        "FIRSTNAMELASTNAME"     { $UPN = "$firstName$lastName".ToLower() + $adDomain }
        "LASTNAMEFIRSTNAME"     { $UPN = "$lastName$firstName".ToLower() + $adDomain }
        "FIRSTINITIALLASTNAME"  { 
            if (-not [string]::IsNullOrWhiteSpace($firstName)) {
                $UPN = "$($firstName.Substring(0,1))$lastName".ToLower() + $adDomain
            } else {
                $UPN = "x$lastName".ToLower() + $adDomain
            }
        }
        "FIRSTNAMELASTINITIAL"  { 
            if (-not [string]::IsNullOrWhiteSpace($lastName)) {
                $UPN = "$firstName$($lastName.Substring(0,1))".ToLower() + $adDomain
            } else {
                $UPN = "$firstName x".ToLower() + $adDomain
            }
        }
        default                 { $UPN = "$firstName.$lastName".ToLower() + $adDomain }
    }

    # 8) Clean UPN of umlauts and special characters
    if ($UPN) {
        $UPN = $UPN -replace "ä", "ae" -replace "ö", "oe" -replace "ü", "ue" -replace "ß", "ss" -replace "Ä", "ae" -replace "Ö", "oe" -replace "Ü", "ue"
        $UPN = $UPN -replace "[^a-zA-Z0-9._@-]", ""
    } else {
        # Fallback for empty UPN
        $UPN = "user@" + $companyData[$domainKey ?? $fallbackDomainKey].Trim()
    }

    # 9) Return result
    return @{
        SamAccountName = $SamAccountName
        UPN            = $UPN
        CompanySection = $companySection
    }
}
#endregion
