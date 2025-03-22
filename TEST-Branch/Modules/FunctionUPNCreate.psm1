# Check if Write-DebugMessage exists, if not create a stub function
if (-not (Get-Command -Name "Write-DebugMessage" -ErrorAction SilentlyContinue)) {
    function Write-DebugMessage { param([string]$Message) Write-Verbose $Message }
}

# Creates user principal names based on templates and user data
Write-DebugMessage "Generating UPN for user."
function New-UPN {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [pscustomobject]$userData,
        [Parameter(Mandatory = $true)]
        [hashtable]$Config
    )

    #region Input Validation
    # Validate that required user data is provided
    if ([string]::IsNullOrWhiteSpace($userData.FirstName) -or [string]::IsNullOrWhiteSpace($userData.LastName)) {
        Throw "Error: FirstName and LastName must be set!"
    }
    #endregion

    #region SamAccountName Generation
    # 1) Generate SamAccountName (first letter of first name + entire last name, all lowercase)
    $SamAccountName = if ($userData.FirstName.Length -gt 0) {
        ($userData.FirstName.Substring(0,1) + $userData.LastName).ToLower()
    } else {
        $userData.LastName.ToLower()
    }
    Write-DebugMessage "SamAccountName= $SamAccountName"
    #endregion

    #region Manual UPN Handling
    # 2) If a manual UPN was entered, use it immediately
    if (-not [string]::IsNullOrWhiteSpace($userData.UPNEntered)) {
        Write-DebugMessage "Manual UPN provided: $($userData.UPNEntered)"
        return [pscustomobject]@{
            SamAccountName = $SamAccountName
            UPN            = $userData.UPNEntered
            CompanySection = "Company"
        }
    }
    #endregion

    #region Company Section Determination
    # 3) Determine the Company section (default: "Company")
    $companySection = "Company"  # Default value
    
    if ($userData.CompanySection) {
        if ($userData.CompanySection -is [string] -and -not [string]::IsNullOrWhiteSpace($userData.CompanySection)) {
            $companySection = $userData.CompanySection
        } elseif ($userData.CompanySection -is [PSObject] -and $userData.CompanySection.PSObject.Properties['Section'] -ne $null -and -not [string]::IsNullOrWhiteSpace($userData.CompanySection.Section)) {
            $companySection = $userData.CompanySection.Section
        }
    }
    Write-DebugMessage "Using Company section: '$companySection'"
    #endregion

    #region INI Configuration Retrieval
    # 4) Check if the desired section exists in the INI with improved error handling
    if (-not $Config) {
        Throw "Error: Config object is NULL! No configuration available."
    }
    
    if (-not $Config.ContainsKey($companySection)) {
        Write-DebugMessage "Section '$companySection' not found in Config, using 'Company' as fallback."
        $companySection = "Company" # Fallback to default value
        if (-not $Config.ContainsKey($companySection)) {
            Throw "Error: Neither the requested section nor the default section 'Company' exists in the INI!"
        }
    }
    $companyData = $Config[$companySection]
    $suffix = $companySection -replace "\D",""
    #endregion

    #region Domain Retrieval
    # 5) Determine domain key and read domain - with improved null checks
    $domainKey = "CompanyActiveDirectoryDomain$suffix"
    $fallbackDomainKey = "CompanyActiveDirectoryDomain"
    
    if ($companyData.ContainsKey($domainKey) -and -not [string]::IsNullOrWhiteSpace($companyData[$domainKey])) {
        $adDomain = "@" + $companyData[$domainKey].Trim()
    } elseif ($companyData.ContainsKey($fallbackDomainKey) -and -not [string]::IsNullOrWhiteSpace($companyData[$fallbackDomainKey])) {
        $adDomain = "@" + $companyData[$fallbackDomainKey].Trim()
        Write-DebugMessage "Using fallback domain: $adDomain"
    } else {
        $availableKeys = if ($companyData) { ($companyData.Keys -join ", ") } else { "No keys available" }
        Throw "Error: Domain information missing in the INI! Neither '$domainKey' nor '$fallbackDomainKey' found or valid. Available keys: $availableKeys"
    }
    #endregion

    #region UPN Template Retrieval
    # 6) Always use UPN template from INI (no GUI dropdown) - with optimized NULL check
    $upnTemplate = "FIRSTNAME.LASTNAME" # Safe default value
    
    $displayNameTemplates = $Config['DisplayNameUPNTemplates']
    
    if ($displayNameTemplates -and $displayNameTemplates.ContainsKey("DefaultUserPrincipalNameFormat") -and -not [string]::IsNullOrWhiteSpace($displayNameTemplates["DefaultUserPrincipalNameFormat"])) {
        $upnTemplate = $displayNameTemplates["DefaultUserPrincipalNameFormat"].ToUpper()
        Write-DebugMessage "UPN template loaded from INI: $upnTemplate"
    } else {
        Write-DebugMessage "No valid UPN template found in INI, using default '$upnTemplate'"
    }
    
    Write-DebugMessage "UPN template (from INI): $upnTemplate"
    #endregion

    #region UPN Generation Based on Template
    # 7) Generate UPN based on template - with safe substring processing and null checks
    $firstName = $userData.FirstName
    $lastName = $userData.LastName
    
    $UPN = switch ($upnTemplate) {
        "FIRSTNAME.LASTNAME"    { "$firstName.$lastName".ToLower() + $adDomain }
        "LASTNAME.FIRSTNAME"    { "$lastName.$firstName".ToLower() + $adDomain }
        "FIRSTINITIAL.LASTNAME" { 
            if ($firstName -and $firstName.Length -gt 0) {
                "$($firstName.Substring(0,1)).$lastName".ToLower() + $adDomain
            } else {
                "x.$lastName".ToLower() + $adDomain
            }
        }
        "FIRSTNAME.LASTINITIAL" { 
            if ($lastName) {
                "$firstName.$($lastName.Substring(0,1))".ToLower() + $adDomain
            } else {
                "$firstName.x".ToLower() + $adDomain
            }
        }
        "FIRSTNAME_LASTNAME"    { "$firstName_$lastName".ToLower() + $adDomain }
        "LASTNAME_FIRSTNAME"    { "$lastName_$firstName".ToLower() + $adDomain }
        "FIRSTINITIAL_LASTNAME" { 
            if ($firstName) {
                "$($firstName.Substring(0,1))_$lastName".ToLower() + $adDomain
            } else {
                "x_$lastName".ToLower() + $adDomain
            }
        }
        "FIRSTNAME_LASTINITIAL" { 
            if ($lastName) {
                "$firstName_$($lastName.Substring(0,1))".ToLower() + $adDomain
            } else {
                "$firstName_x".ToLower() + $adDomain
            }
        }
        "FIRSTNAMELASTNAME"     { "$firstName$lastName".ToLower() + $adDomain }
        "LASTNAMEFIRSTNAME"     { "$lastName$firstName".ToLower() + $adDomain }
        "FIRSTINITIALLASTNAME"  { 
            if ($firstName) {
                "$($firstName.Substring(0,1))$lastName".ToLower() + $adDomain
            } else {
                "x$lastName".ToLower() + $adDomain
            }
        }
        "FIRSTNAMELASTINITIAL"  { 
            if ($lastName) {
                "$firstName$($lastName.Substring(0,1))".ToLower() + $adDomain
            } else {
                "$firstName x".ToLower() + $adDomain
            }
        }
        default                 { "$firstName.$lastName".ToLower() + $adDomain }
    }
    #region UPN Cleaning
    # 8) Clean UPN of umlauts and special characters
    $UPN = $UPN -replace "ä", "ae" -replace "ö", "oe" -replace "ü", "ue" -replace "ß", "ss" -replace "Ä", "Ae" -replace "Ö", "Oe" -replace "Ü", "Ue"
    $UPN = $UPN -replace "[^a-zA-Z0-9._@-]", ""

    #region Fallback for Empty UPN
    if ([string]::IsNullOrWhiteSpace($UPN)) {
        # Fallback for empty UPN
        $UPN = "user@" + $companyData[$domainKey, $fallbackDomainKey].Trim()
    }

    #region Output
    # 9) Return result
    return [pscustomobject]@{
        SamAccountName = $SamAccountName
        UPN            = $UPN
        CompanySection = $companySection
    }
}

# Export the function for use in other modules or scripts
Export-ModuleMember -Function New-UPN
