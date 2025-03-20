#region [Region 09 | AD USER CREATION]
# Creates Active Directory user accounts based on input data
function New-ADUserAccount {
    param (
        [Parameter(Mandatory=$true)]
        [pscustomobject]$UserData,
        [Parameter(Mandatory=$true)]
        [hashtable]$Config
    )
    
    try {
        # 1) Generate UPN data
        $upnData = New-UPN -userData $UserData -Config $Config
        $samAccountName = $upnData.SamAccountName
        $userPrincipalName = $upnData.UPN
        $companySection = $upnData.CompanySection
        
        Write-DebugMessage "Generated UPN: $userPrincipalName"
        Write-DebugMessage "Using Company section: $companySection"
        
        # 2) Check if all required information is available
        if (-not $Config.Contains($companySection)) {
            throw "No configuration found for Company section '$companySection'."
        }
        
        $suffix = ($companySection -replace "\D","") 
        $adOUKey = "CompanyActiveDirectoryOU$suffix"
        $adPath = $Config[$companySection][$adOUKey]
        
        if (-not $adPath) {
            throw "No AD path (OU) found for Company section '$companySection'."
        }
        
        # 3) Generate DisplayName according to configured template (from INI)
        $displayNameFormat = ""
        if ($Config.Contains("DisplayNameUPNTemplates") -and 
            $Config.DisplayNameUPNTemplates.Contains("DefaultDisplayNameFormat")) {
            $displayNameFormat = $Config.DisplayNameUPNTemplates.DefaultDisplayNameFormat
        }
        
        # Default to "LastName, FirstName"
        if ([string]::IsNullOrWhiteSpace($displayNameFormat)) {
            $displayNameFormat = "LastName, FirstName"
        }
        
        $displayName = switch -Wildcard ($displayNameFormat) {
            "LastName, FirstName" { "$($UserData.LastName), $($UserData.FirstName)" }
            "FirstName LastName"  { "$($UserData.FirstName) $($UserData.LastName)" }
            "LastName FirstName"  { "$($UserData.LastName) $($UserData.FirstName)" }
            default               { "$($UserData.LastName), $($UserData.FirstName)" }
        }
        
        # 4) Check if user already exists
        $userExists = $false
        try {
            $existingUser = Get-ADUser -Identity $samAccountName
            if ($existingUser) {
                $userExists = $true
                Write-Log "User $samAccountName already exists." -LogLevel "WARN"
            }
        } catch {
            # User doesn't exist - that's good
            $userExists = $false
        }
        
        if ($userExists) {
            return @{
                Success = $false
                Message = "User $samAccountName already exists."
                SamAccountName = $samAccountName
            }
        }
        
        # 5) Generate secure password if not set
        if ([string]::IsNullOrWhiteSpace($UserData.Password)) {
            # Password length from Config or default 12 characters
            $passwordLength = 12
            if ($Config.Contains("PasswordFixGenerate") -and $Config.PasswordFixGenerate.Contains("DefaultPasswordLength")) {
                $tempLength = [int]::TryParse($Config.PasswordFixGenerate.DefaultPasswordLength, [ref]$passwordLength)
                if (-not $tempLength) { $passwordLength = 12 }
            }
            
            # Character pool for password
            $charPool = "abcdefghijkmnopqrstuvwxyzABCDEFGHJKLMNPQRSTUVWXYZ23456789!@#$%^&*_=+-"
            $securePassword = ""
            $random = New-Object System.Random
            
            # At least 1 uppercase, 1 lowercase, 1 number and 1 special character
            $securePassword += $charPool.Substring($random.Next(0, 25), 1)  # Lowercase
            $securePassword += $charPool.Substring($random.Next(26, 50), 1) # Uppercase
            $securePassword += $charPool.Substring($random.Next(51, 59), 1) # Number
            $securePassword += $charPool.Substring($random.Next(60, $charPool.Length-1), 1) # Special character
            
            # Fill up to desired length
            for ($i = 4; $i -lt $passwordLength; $i++) {
                $securePassword += $charPool.Substring($random.Next(0, $charPool.Length-1), 1)
            }
            
            # Randomize character order
            $securePasswordArray = $securePassword.ToCharArray()
            $randomizedPassword = ""
            for ($i = $securePasswordArray.Count; $i -gt 0; $i--) {
                $randomPosition = $random.Next(0, $i)
                $randomizedPassword += $securePasswordArray[$randomPosition]
                $securePasswordArray = $securePasswordArray[0..($randomPosition-1)] + $securePasswordArray[($randomPosition+1)..($securePasswordArray.Count-1)]
            }
            
            $UserData.Password = $randomizedPassword
            Write-DebugMessage "Generated password: $randomizedPassword"
        }
        
        # 6) Collect more AD attributes
        $adUserParams = @{
            SamAccountName = $samAccountName
            UserPrincipalName = $userPrincipalName
            Name = $displayName
            DisplayName = $displayName
            GivenName = $UserData.FirstName
            Surname = $UserData.LastName
            Path = $adPath
            AccountPassword = (ConvertTo-SecureString -String $UserData.Password -AsPlainText -Force)
            Enabled = $true
            PasswordNeverExpires = $false
            ChangePasswordAtLogon = $true
        }
        
        # Optional: More attributes if available
        if (-not [string]::IsNullOrWhiteSpace($UserData.Email)) {
            $adUserParams.EmailAddress = $UserData.Email
        }
        
        if (-not [string]::IsNullOrWhiteSpace($UserData.Description)) {
            $adUserParams.Description = $UserData.Description
        }
        
        if (-not [string]::IsNullOrWhiteSpace($UserData.Phone)) {
            $adUserParams.OfficePhone = $UserData.Phone
        }
        
        # 7) Create user
        Write-DebugMessage "Creating AD user: $samAccountName"
        $newUser = New-ADUser @adUserParams -PassThru
        
        # 8) Set optional attributes
        # If certain attributes need to be set separately...
        
        # 9) Set password options (if configured)
        if ($UserData.PreventPasswordChange -eq $true) {
            Write-DebugMessage "Setting 'Prevent Password Change' for $samAccountName"
            Set-CannotChangePassword -SamAccountName $samAccountName
        }
        
        Write-Log "User $samAccountName was successfully created." -LogLevel "INFO"
        return @{
            Success = $true
            Message = "User was successfully created."
            SamAccountName = $samAccountName
            UserPrincipalName = $userPrincipalName
            Password = $UserData.Password
        }
    }
    catch {
        Write-Log "Error creating user: $($_.Exception.Message)" -LogLevel "ERROR"
        return @{
            Success = $false
            Message = "Error creating user: $($_.Exception.Message)"
            SamAccountName = $samAccountName
        }
    }
}
#endregion
