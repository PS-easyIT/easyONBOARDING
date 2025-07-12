# ExtendedFunctions2.psm1 - Additional extended functionality
# This file contains additional functions that extend ExtendedFunctions.psm1

# Set user photo in AD and Exchange
function Set-UserPhoto {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Identity,
        
        [Parameter(Mandatory = $true)]
        [string]$PhotoPath,
        
        [Parameter(Mandatory = $false)]
        [ValidateSet('AD', 'Exchange', 'Both')]
        [string]$Target = 'Both',
        
        [Parameter(Mandatory = $false)]
        [int]$MaxSizeKB = 100
    )
    
    try {
        Write-ModuleLog "Setting user photo for $Identity" "INFO"
        
        # Validate photo file
        if (-not (Test-Path $PhotoPath)) {
            throw "Photo file not found: $PhotoPath"
        }
        
        # Check file size
        $photoFile = Get-Item $PhotoPath
        $fileSizeKB = [math]::Round($photoFile.Length / 1KB, 2)
        
        if ($fileSizeKB -gt $MaxSizeKB) {
            Write-ModuleLog "Photo file is too large ($fileSizeKB KB). Resizing..." "WARN"
            # Here you would implement image resizing logic
        }
        
        # Read photo data
        $photoData = [System.IO.File]::ReadAllBytes($PhotoPath)
        
        # Set AD photo
        if ($Target -eq 'AD' -or $Target -eq 'Both') {
            try {
                Set-ADUser -Identity $Identity -Replace @{thumbnailPhoto = $photoData}
                Write-ModuleLog "AD photo set successfully" "INFO"
            }
            catch {
                Write-ModuleLog "Error setting AD photo: $($_.Exception.Message)" "ERROR"
                if ($Target -eq 'AD') { throw }
            }
        }
        
        # Set Exchange photo
        if ($Target -eq 'Exchange' -or $Target -eq 'Both') {
            try {
                if (Get-Command Set-UserPhoto -ErrorAction SilentlyContinue) {
                    Set-UserPhoto -Identity $Identity -PictureData $photoData -Confirm:$false
                    Write-ModuleLog "Exchange photo set successfully" "INFO"
                }
                else {
                    Write-ModuleLog "Exchange cmdlets not available" "WARN"
                }
            }
            catch {
                Write-ModuleLog "Error setting Exchange photo: $($_.Exception.Message)" "ERROR"
                if ($Target -eq 'Exchange') { throw }
            }
        }
        
        return $true
    }
    catch {
        Write-ModuleLog "Error setting user photo: $($_.Exception.Message)" "ERROR"
        throw
    }
}

# Create home directory for user
function New-HomeDirectory {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$SamAccountName,
        
        [Parameter(Mandatory = $true)]
        [string]$HomePath,
        
        [Parameter(Mandatory = $false)]
        [string]$HomeDrive = "H:",
        
        [Parameter(Mandatory = $false)]
        [int]$QuotaMB = 5120,
        
        [Parameter(Mandatory = $false)]
        [switch]$SetInAD
    )
    
    try {
        Write-ModuleLog "Creating home directory for $SamAccountName" "INFO"
        
        # Build full path
        $fullPath = Join-Path $HomePath $SamAccountName
        
        # Create directory
        if (-not (Test-Path $fullPath)) {
            New-Item -ItemType Directory -Path $fullPath -Force | Out-Null
            Write-ModuleLog "Directory created: $fullPath" "INFO"
        }
        else {
            Write-ModuleLog "Directory already exists: $fullPath" "WARN"
        }
        
        # Set permissions
        $acl = Get-Acl $fullPath
        
        # Remove inheritance
        $acl.SetAccessRuleProtection($true, $false)
        
        # Clear existing permissions
        $acl.Access | ForEach-Object { $acl.RemoveAccessRule($_) }
        
        # Add SYSTEM full control
        $systemRule = New-Object System.Security.AccessControl.FileSystemAccessRule(
            "NT AUTHORITY\SYSTEM",
            "FullControl",
            "ContainerInherit,ObjectInherit",
            "None",
            "Allow"
        )
        $acl.AddAccessRule($systemRule)
        
        # Add Domain Admins full control
        $adminRule = New-Object System.Security.AccessControl.FileSystemAccessRule(
            "Domain Admins",
            "FullControl",
            "ContainerInherit,ObjectInherit",
            "None",
            "Allow"
        )
        $acl.AddAccessRule($adminRule)
        
        # Add user modify permissions
        $userRule = New-Object System.Security.AccessControl.FileSystemAccessRule(
            $SamAccountName,
            "Modify",
            "ContainerInherit,ObjectInherit",
            "None",
            "Allow"
        )
        $acl.AddAccessRule($userRule)
        
        # Apply ACL
        Set-Acl -Path $fullPath -AclObject $acl
        Write-ModuleLog "Permissions set successfully" "INFO"
        
        # Set quota if on NTFS
        if ($QuotaMB -gt 0) {
            try {
                # This would require FSRM or other quota management tools
                Write-ModuleLog "Quota setting requires FSRM configuration" "INFO"
            }
            catch {
                Write-ModuleLog "Could not set quota: $($_.Exception.Message)" "WARN"
            }
        }
        
        # Update AD user if requested
        if ($SetInAD) {
            try {
                Set-ADUser -Identity $SamAccountName -HomeDirectory $fullPath -HomeDrive $HomeDrive
                Write-ModuleLog "AD home directory attributes updated" "INFO"
            }
            catch {
                Write-ModuleLog "Error updating AD attributes: $($_.Exception.Message)" "ERROR"
            }
        }
        
        return $fullPath
    }
    catch {
        Write-ModuleLog "Error creating home directory: $($_.Exception.Message)" "ERROR"
        throw
    }
}

# Set user permissions
function Set-UserPermissions {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Identity,
        
        [Parameter(Mandatory = $true)]
        [hashtable]$Permissions,
        
        [Parameter(Mandatory = $false)]
        [switch]$RemoveExisting
    )
    
    try {
        Write-ModuleLog "Setting permissions for $Identity" "INFO"
        
        # Get user object
        $user = Get-ADUser -Identity $Identity -Properties memberOf
        
        if (-not $user) {
            throw "User $Identity not found"
        }
        
        # Process each permission type
        foreach ($permType in $Permissions.Keys) {
            switch ($permType) {
                'Groups' {
                    # Handle group memberships
                    $groups = $Permissions.Groups
                    
                    if ($RemoveExisting) {
                        # Remove from all current groups except Domain Users
                        $currentGroups = Get-ADPrincipalGroupMembership -Identity $Identity |
                            Where-Object { $_.Name -ne 'Domain Users' }
                        
                        foreach ($group in $currentGroups) {
                            Remove-ADGroupMember -Identity $group -Members $Identity -Confirm:$false
                        }
                    }
                    
                    # Add to new groups
                    foreach ($group in $groups) {
                        try {
                            Add-ADGroupMember -Identity $group -Members $Identity
                            Write-ModuleLog "Added to group: $group" "INFO"
                        }
                        catch {
                            Write-ModuleLog "Error adding to group $group : $($_.Exception.Message)" "ERROR"
                        }
                    }
                }
                
                'SharedFolders' {
                    # Handle shared folder permissions
                    $folders = $Permissions.SharedFolders
                    
                    foreach ($folder in $folders.Keys) {
                        if (Test-Path $folder) {
                            $acl = Get-Acl $folder
                            $permission = $folders[$folder]
                            
                            $rule = New-Object System.Security.AccessControl.FileSystemAccessRule(
                                $Identity,
                                $permission,
                                "ContainerInherit,ObjectInherit",
                                "None",
                                "Allow"
                            )
                            
                            $acl.AddAccessRule($rule)
                            Set-Acl -Path $folder -AclObject $acl
                            Write-ModuleLog "Set $permission permission on $folder" "INFO"
                        }
                    }
                }
                
                'Applications' {
                    # Handle application-specific permissions
                    # This would be customized based on your applications
                    Write-ModuleLog "Application permissions would be set here" "INFO"
                }
            }
        }
        
        return $true
    }
    catch {
        Write-ModuleLog "Error setting permissions: $($_.Exception.Message)" "ERROR"
        throw
    }
}

# Backup user data
function New-UserDataBackup {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Identity,
        
        [Parameter(Mandatory = $true)]
        [string]$BackupPath,
        
        [Parameter(Mandatory = $false)]
        [switch]$IncludeMailbox,
        
        [Parameter(Mandatory = $false)]
        [switch]$IncludeHomeDirectory,
        
        [Parameter(Mandatory = $false)]
        [switch]$Compress
    )
    
    try {
        Write-ModuleLog "Starting backup for user $Identity" "INFO"
        
        # Create backup directory
        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $backupDir = Join-Path $BackupPath "$Identity_$timestamp"
        New-Item -ItemType Directory -Path $backupDir -Force | Out-Null
        
        # Export AD data
        $adBackupFile = Join-Path $backupDir "AD_Data.xml"
        Get-ADUser -Identity $Identity -Properties * | Export-Clixml -Path $adBackupFile
        Write-ModuleLog "AD data backed up" "INFO"
        
        # Export group memberships
        $groupsFile = Join-Path $backupDir "Groups.txt"
        Get-ADPrincipalGroupMembership -Identity $Identity | 
            Select-Object -ExpandProperty Name | 
            Out-File -FilePath $groupsFile
        Write-ModuleLog "Group memberships backed up" "INFO"
        
        # Backup home directory if requested
        if ($IncludeHomeDirectory) {
            $user = Get-ADUser -Identity $Identity -Properties HomeDirectory
            if ($user.HomeDirectory -and (Test-Path $user.HomeDirectory)) {
                $homeBackup = Join-Path $backupDir "HomeDirectory"
                Copy-Item -Path $user.HomeDirectory -Destination $homeBackup -Recurse
                Write-ModuleLog "Home directory backed up" "INFO"
            }
        }
        
        # Backup mailbox if requested (requires Exchange tools)
        if ($IncludeMailbox) {
            if (Get-Command New-MailboxExportRequest -ErrorAction SilentlyContinue) {
                $pstFile = Join-Path $backupDir "$Identity.pst"
                # Note: This is a simplified example. Real implementation would need proper Exchange permissions
                Write-ModuleLog "Mailbox backup would be performed here" "INFO"
            }
        }
        
        # Compress if requested
        if ($Compress) {
            $zipFile = "$backupDir.zip"
            Compress-Archive -Path $backupDir -DestinationPath $zipFile -Force
            Remove-Item -Path $backupDir -Recurse -Force
            Write-ModuleLog "Backup compressed to $zipFile" "INFO"
            return $zipFile
        }
        
        return $backupDir
    }
    catch {
        Write-ModuleLog "Error during backup: $($_.Exception.Message)" "ERROR"
        throw
    }
}

# Generate detailed user report
function New-UserReport {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Identity,
        
        [Parameter(Mandatory = $true)]
        [ValidateSet('Summary', 'Detailed', 'Audit')]
        [string]$ReportType,
        
        [Parameter(Mandatory = $true)]
        [string]$OutputPath,
        
        [Parameter(Mandatory = $false)]
        [int]$DaysBack = 30
    )
    
    try {
        Write-ModuleLog "Generating $ReportType report for $Identity" "INFO"
        
        # Get user data
        $user = Get-ADUser -Identity $Identity -Properties *
        
        $report = switch ($ReportType) {
            'Summary' {
                @"
USER SUMMARY REPORT
==================
Generated: $(Get-Date)
User: $($user.DisplayName) ($($user.SamAccountName))

Basic Information:
- UPN: $($user.UserPrincipalName)
- Email: $($user.EmailAddress)
- Department: $($user.Department)
- Title: $($user.Title)
- Manager: $($user.Manager)
- Created: $($user.Created)
- Last Modified: $($user.Modified)
- Account Status: $(if($user.Enabled){"Enabled"}else{"Disabled"})

Group Memberships:
$(Get-ADPrincipalGroupMembership -Identity $Identity | ForEach-Object { "- $($_.Name)" } | Out-String)
"@
            }
            
            'Detailed' {
                # Create detailed HTML report
                @"
<!DOCTYPE html>
<html>
<head>
    <title>Detailed User Report - $($user.DisplayName)</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; }
        h1, h2 { color: #0078D4; }
        table { border-collapse: collapse; width: 100%; margin: 20px 0; }
        th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }
        th { background-color: #0078D4; color: white; }
        tr:nth-child(even) { background-color: #f2f2f2; }
    </style>
</head>
<body>
    <h1>Detailed User Report</h1>
    <h2>$($user.DisplayName)</h2>
    
    <h3>Account Information</h3>
    <table>
        <tr><th>Property</th><th>Value</th></tr>
        <tr><td>SAM Account Name</td><td>$($user.SamAccountName)</td></tr>
        <tr><td>User Principal Name</td><td>$($user.UserPrincipalName)</td></tr>
        <tr><td>Email Address</td><td>$($user.EmailAddress)</td></tr>
        <tr><td>Display Name</td><td>$($user.DisplayName)</td></tr>
        <tr><td>Account Status</td><td>$(if($user.Enabled){"<span style='color:green'>Enabled</span>"}else{"<span style='color:red'>Disabled</span>"})</td></tr>
        <tr><td>Created</td><td>$($user.Created)</td></tr>
        <tr><td>Last Modified</td><td>$($user.Modified)</td></tr>
        <tr><td>Last Logon</td><td>$($user.LastLogonDate)</td></tr>
    </table>
    
    <h3>Organization Information</h3>
    <table>
        <tr><th>Property</th><th>Value</th></tr>
        <tr><td>Department</td><td>$($user.Department)</td></tr>
        <tr><td>Title</td><td>$($user.Title)</td></tr>
        <tr><td>Company</td><td>$($user.Company)</td></tr>
        <tr><td>Office</td><td>$($user.Office)</td></tr>
        <tr><td>Manager</td><td>$($user.Manager)</td></tr>
    </table>
    
    <p>Report generated on: $(Get-Date)</p>
</body>
</html>
"@
            }
            
            'Audit' {
                # Generate audit report with recent activities
                Write-ModuleLog "Generating audit report for last $DaysBack days" "INFO"
                "Audit report generation would query event logs and AD changes"
            }
        }
        
        # Save report
        $report | Out-File -FilePath $OutputPath -Encoding UTF8
        Write-ModuleLog "Report saved to $OutputPath" "INFO"
        
        return $OutputPath
    }
    catch {
        Write-ModuleLog "Error generating report: $($_.Exception.Message)" "ERROR"
        throw
    }
}

# Validate user input
function Test-UserInput {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$UserData,
        
        [Parameter(Mandatory = $true)]
        [hashtable]$ValidationRules
    )
    
    $errors = @()
    
    try {
        # Validate each field based on rules
        foreach ($field in $ValidationRules.Keys) {
            $rule = $ValidationRules[$field]
            $value = $UserData[$field]
            
            # Check required fields
            if ($rule.Required -and [string]::IsNullOrWhiteSpace($value)) {
                $errors += "$field is required"
                continue
            }
            
            # Skip further validation if not required and empty
            if ([string]::IsNullOrWhiteSpace($value)) {
                continue
            }
            
            # Check pattern
            if ($rule.Pattern -and $value -notmatch $rule.Pattern) {
                $errors += "$field does not match required pattern"
            }
            
            # Check length
            if ($rule.MinLength -and $value.Length -lt $rule.MinLength) {
                $errors += "$field must be at least $($rule.MinLength) characters"
            }
            
            if ($rule.MaxLength -and $value.Length -gt $rule.MaxLength) {
                $errors += "$field must not exceed $($rule.MaxLength) characters"
            }
            
            # Custom validation
            if ($rule.CustomValidation) {
                $result = & $rule.CustomValidation $value
                if (-not $result.Valid) {
                    $errors += $result.Message
                }
            }
        }
        
        # Check for duplicate users
        if ($UserData.SamAccountName) {
            try {
                $existing = Get-ADUser -Identity $UserData.SamAccountName -ErrorAction SilentlyContinue
                if ($existing) {
                    $errors += "User with SAM Account Name '$($UserData.SamAccountName)' already exists"
                }
            }
            catch {
                # User doesn't exist, which is good
            }
        }
        
        return @{
            Valid = $errors.Count -eq 0
            Errors = $errors
        }
    }
    catch {
        Write-ModuleLog "Error during validation: $($_.Exception.Message)" "ERROR"
        throw
    }
}

# Test prerequisites
function Test-Prerequisites {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [string[]]$RequiredModules = @('ActiveDirectory'),
        
        [Parameter(Mandatory = $false)]
        [string[]]$RequiredFeatures = @(),
        
        [Parameter(Mandatory = $false)]
        [hashtable]$RequiredServices = @{},
        
        [Parameter(Mandatory = $false)]
        [switch]$Detailed
    )
    
    $results = @{
        Success = $true
        Details = @()
        Errors = @()
    }
    
    try {
        Write-ModuleLog "Testing prerequisites" "INFO"
        
        # Check PowerShell version
        $psVersion = $PSVersionTable.PSVersion
        if ($psVersion.Major -lt 5) {
            $results.Success = $false
            $results.Errors += "PowerShell 5.0 or higher required (current: $psVersion)"
        }
        else {
            $results.Details += "PowerShell version: $psVersion ✓"
        }
        
        # Check required modules
        foreach ($module in $RequiredModules) {
            if (Get-Module -ListAvailable -Name $module) {
                $results.Details += "Module '$module' available ✓"
            }
            else {
                $results.Success = $false
                $results.Errors += "Required module '$module' not found"
            }
        }
        
        # Check Windows features
        foreach ($feature in $RequiredFeatures) {
            $state = Get-WindowsFeature -Name $feature -ErrorAction SilentlyContinue
            if ($state -and $state.InstallState -eq 'Installed') {
                $results.Details += "Feature '$feature' installed ✓"
            }
            else {
                $results.Success = $false
                $results.Errors += "Required feature '$feature' not installed"
            }
        }
        
        # Check services
        foreach ($service in $RequiredServices.Keys) {
            $svc = Get-Service -Name $service -ErrorAction SilentlyContinue
            if ($svc) {
                if ($RequiredServices[$service] -eq 'Running' -and $svc.Status -ne 'Running') {
                    $results.Success = $false
                    $results.Errors += "Service '$service' is not running"
                }
                else {
                    $results.Details += "Service '$service' status: $($svc.Status) ✓"
                }
            }
            else {
                $results.Success = $false
                $results.Errors += "Required service '$service' not found"
            }
        }
        
        # Check AD connectivity
        try {
            $domain = Get-ADDomain
            $results.Details += "AD Domain: $($domain.DNSRoot) ✓"
        }
        catch {
            $results.Success = $false
            $results.Errors += "Cannot connect to Active Directory"
        }
        
        return $results
    }
    catch {
        Write-ModuleLog "Error testing prerequisites: $($_.Exception.Message)" "ERROR"
        throw
    }
}

# Export additional functions
Export-ModuleMember -Function @(
    'Set-UserPhoto',
    'New-HomeDirectory',      # Changed from Create-HomeDirectory
    'Set-UserPermissions',
    'New-UserDataBackup',     # Changed from Backup-UserData
    'New-UserReport',         # Changed from Generate-UserReport
    'Test-UserInput',         # Changed from Validate-UserInput
    'Test-Prerequisites'
) 
# SIG # Begin signature block
# MIIcCAYJKoZIhvcNAQcCoIIb+TCCG/UCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBqZSRbqeM6GrJ/
# xCI1toSUJmhx2vVfB2IFiqOf+RDwDKCCFk4wggMQMIIB+KADAgECAhB3jzsyX9Cg
# jEi+sBC2rBMTMA0GCSqGSIb3DQEBCwUAMCAxHjAcBgNVBAMMFVBoaW5JVC1QU3Nj
# cmlwdHNfU2lnbjAeFw0yNTA3MDUwODI4MTZaFw0yNzA3MDUwODM4MTZaMCAxHjAc
# BgNVBAMMFVBoaW5JVC1QU3NjcmlwdHNfU2lnbjCCASIwDQYJKoZIhvcNAQEBBQAD
# ggEPADCCAQoCggEBALmz3o//iDA5MvAndTjGX7/AvzTSACClfuUR9WYK0f6Ut2dI
# mPxn+Y9pZlLjXIpZT0H2Lvxq5aSI+aYeFtuJ8/0lULYNCVT31Bf+HxervRBKsUyi
# W9+4PH6STxo3Pl4l56UNQMcWLPNjDORWRPWHn0f99iNtjI+L4tUC/LoWSs3obzxN
# 3uTypzlaPBxis2qFSTR5SWqFdZdRkcuI5LNsJjyc/QWdTYRrfmVqp0QrvcxzCv8u
# EiVuni6jkXfiE6wz+oeI3L2iR+ywmU6CUX4tPWoS9VTtmm7AhEpasRTmrrnSg20Q
# jiBa1eH5TyLAH3TcYMxhfMbN9a2xDX5pzM65EJUCAwEAAaNGMEQwDgYDVR0PAQH/
# BAQDAgeAMBMGA1UdJQQMMAoGCCsGAQUFBwMDMB0GA1UdDgQWBBQO7XOqiE/EYi+n
# IaR6YO5M2MUuVTANBgkqhkiG9w0BAQsFAAOCAQEAjYOKIwBu1pfbdvEFFaR/uY88
# peKPk0NnvNEc3dpGdOv+Fsgbz27JPvItITFd6AKMoN1W48YjQLaU22M2jdhjGN5i
# FSobznP5KgQCDkRsuoDKiIOTiKAAknjhoBaCCEZGw8SZgKJtWzbST36Thsdd/won
# ihLsuoLxfcFnmBfrXh3rTIvTwvfujob68s0Sf5derHP/F+nphTymlg+y4VTEAijk
# g2dhy8RAsbS2JYZT7K5aEJpPXMiOLBqd7oTGfM7y5sLk2LIM4cT8hzgz3v5yPMkF
# H2MdR//K403e1EKH9MsGuGAJZddVN8ppaiESoPLoXrgnw2SY5KCmhYw1xRFdjTCC
# BY0wggR1oAMCAQICEA6bGI750C3n79tQ4ghAGFowDQYJKoZIhvcNAQEMBQAwZTEL
# MAkGA1UEBhMCVVMxFTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3
# LmRpZ2ljZXJ0LmNvbTEkMCIGA1UEAxMbRGlnaUNlcnQgQXNzdXJlZCBJRCBSb290
# IENBMB4XDTIyMDgwMTAwMDAwMFoXDTMxMTEwOTIzNTk1OVowYjELMAkGA1UEBhMC
# VVMxFTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0
# LmNvbTEhMB8GA1UEAxMYRGlnaUNlcnQgVHJ1c3RlZCBSb290IEc0MIICIjANBgkq
# hkiG9w0BAQEFAAOCAg8AMIICCgKCAgEAv+aQc2jeu+RdSjwwIjBpM+zCpyUuySE9
# 8orYWcLhKac9WKt2ms2uexuEDcQwH/MbpDgW61bGl20dq7J58soR0uRf1gU8Ug9S
# H8aeFaV+vp+pVxZZVXKvaJNwwrK6dZlqczKU0RBEEC7fgvMHhOZ0O21x4i0MG+4g
# 1ckgHWMpLc7sXk7Ik/ghYZs06wXGXuxbGrzryc/NrDRAX7F6Zu53yEioZldXn1RY
# jgwrt0+nMNlW7sp7XeOtyU9e5TXnMcvak17cjo+A2raRmECQecN4x7axxLVqGDgD
# EI3Y1DekLgV9iPWCPhCRcKtVgkEy19sEcypukQF8IUzUvK4bA3VdeGbZOjFEmjNA
# vwjXWkmkwuapoGfdpCe8oU85tRFYF/ckXEaPZPfBaYh2mHY9WV1CdoeJl2l6SPDg
# ohIbZpp0yt5LHucOY67m1O+SkjqePdwA5EUlibaaRBkrfsCUtNJhbesz2cXfSwQA
# zH0clcOP9yGyshG3u3/y1YxwLEFgqrFjGESVGnZifvaAsPvoZKYz0YkH4b235kOk
# GLimdwHhD5QMIR2yVCkliWzlDlJRR3S+Jqy2QXXeeqxfjT/JvNNBERJb5RBQ6zHF
# ynIWIgnffEx1P2PsIV/EIFFrb7GrhotPwtZFX50g/KEexcCPorF+CiaZ9eRpL5gd
# LfXZqbId5RsCAwEAAaOCATowggE2MA8GA1UdEwEB/wQFMAMBAf8wHQYDVR0OBBYE
# FOzX44LScV1kTN8uZz/nupiuHA9PMB8GA1UdIwQYMBaAFEXroq/0ksuCMS1Ri6en
# IZ3zbcgPMA4GA1UdDwEB/wQEAwIBhjB5BggrBgEFBQcBAQRtMGswJAYIKwYBBQUH
# MAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBDBggrBgEFBQcwAoY3aHR0cDov
# L2NhY2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEUm9vdENBLmNy
# dDBFBgNVHR8EPjA8MDqgOKA2hjRodHRwOi8vY3JsMy5kaWdpY2VydC5jb20vRGln
# aUNlcnRBc3N1cmVkSURSb290Q0EuY3JsMBEGA1UdIAQKMAgwBgYEVR0gADANBgkq
# hkiG9w0BAQwFAAOCAQEAcKC/Q1xV5zhfoKN0Gz22Ftf3v1cHvZqsoYcs7IVeqRq7
# IviHGmlUIu2kiHdtvRoU9BNKei8ttzjv9P+Aufih9/Jy3iS8UgPITtAq3votVs/5
# 9PesMHqai7Je1M/RQ0SbQyHrlnKhSLSZy51PpwYDE3cnRNTnf+hZqPC/Lwum6fI0
# POz3A8eHqNJMQBk1RmppVLC4oVaO7KTVPeix3P0c2PR3WlxUjG/voVA9/HYJaISf
# b8rbII01YBwCA8sgsKxYoA5AY8WYIsGyWfVVa88nq2x2zm8jLfR+cWojayL/ErhU
# LSd+2DrZ8LaHlv1b0VysGMNNn3O3AamfV6peKOK5lDCCBrQwggScoAMCAQICEA3H
# rFcF/yGZLkBDIgw6SYYwDQYJKoZIhvcNAQELBQAwYjELMAkGA1UEBhMCVVMxFTAT
# BgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNvbTEh
# MB8GA1UEAxMYRGlnaUNlcnQgVHJ1c3RlZCBSb290IEc0MB4XDTI1MDUwNzAwMDAw
# MFoXDTM4MDExNDIzNTk1OVowaTELMAkGA1UEBhMCVVMxFzAVBgNVBAoTDkRpZ2lD
# ZXJ0LCBJbmMuMUEwPwYDVQQDEzhEaWdpQ2VydCBUcnVzdGVkIEc0IFRpbWVTdGFt
# cGluZyBSU0E0MDk2IFNIQTI1NiAyMDI1IENBMTCCAiIwDQYJKoZIhvcNAQEBBQAD
# ggIPADCCAgoCggIBALR4MdMKmEFyvjxGwBysddujRmh0tFEXnU2tjQ2UtZmWgyxU
# 7UNqEY81FzJsQqr5G7A6c+Gh/qm8Xi4aPCOo2N8S9SLrC6Kbltqn7SWCWgzbNfiR
# +2fkHUiljNOqnIVD/gG3SYDEAd4dg2dDGpeZGKe+42DFUF0mR/vtLa4+gKPsYfwE
# u7EEbkC9+0F2w4QJLVSTEG8yAR2CQWIM1iI5PHg62IVwxKSpO0XaF9DPfNBKS7Za
# zch8NF5vp7eaZ2CVNxpqumzTCNSOxm+SAWSuIr21Qomb+zzQWKhxKTVVgtmUPAW3
# 5xUUFREmDrMxSNlr/NsJyUXzdtFUUt4aS4CEeIY8y9IaaGBpPNXKFifinT7zL2gd
# FpBP9qh8SdLnEut/GcalNeJQ55IuwnKCgs+nrpuQNfVmUB5KlCX3ZA4x5HHKS+rq
# BvKWxdCyQEEGcbLe1b8Aw4wJkhU1JrPsFfxW1gaou30yZ46t4Y9F20HHfIY4/6vH
# espYMQmUiote8ladjS/nJ0+k6MvqzfpzPDOy5y6gqztiT96Fv/9bH7mQyogxG9QE
# PHrPV6/7umw052AkyiLA6tQbZl1KhBtTasySkuJDpsZGKdlsjg4u70EwgWbVRSX1
# Wd4+zoFpp4Ra+MlKM2baoD6x0VR4RjSpWM8o5a6D8bpfm4CLKczsG7ZrIGNTAgMB
# AAGjggFdMIIBWTASBgNVHRMBAf8ECDAGAQH/AgEAMB0GA1UdDgQWBBTvb1NK6eQG
# fHrK4pBW9i/USezLTjAfBgNVHSMEGDAWgBTs1+OC0nFdZEzfLmc/57qYrhwPTzAO
# BgNVHQ8BAf8EBAMCAYYwEwYDVR0lBAwwCgYIKwYBBQUHAwgwdwYIKwYBBQUHAQEE
# azBpMCQGCCsGAQUFBzABhhhodHRwOi8vb2NzcC5kaWdpY2VydC5jb20wQQYIKwYB
# BQUHMAKGNWh0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0
# ZWRSb290RzQuY3J0MEMGA1UdHwQ8MDowOKA2oDSGMmh0dHA6Ly9jcmwzLmRpZ2lj
# ZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRSb290RzQuY3JsMCAGA1UdIAQZMBcwCAYG
# Z4EMAQQCMAsGCWCGSAGG/WwHATANBgkqhkiG9w0BAQsFAAOCAgEAF877FoAc/gc9
# EXZxML2+C8i1NKZ/zdCHxYgaMH9Pw5tcBnPw6O6FTGNpoV2V4wzSUGvI9NAzaoQk
# 97frPBtIj+ZLzdp+yXdhOP4hCFATuNT+ReOPK0mCefSG+tXqGpYZ3essBS3q8nL2
# UwM+NMvEuBd/2vmdYxDCvwzJv2sRUoKEfJ+nN57mQfQXwcAEGCvRR2qKtntujB71
# WPYAgwPyWLKu6RnaID/B0ba2H3LUiwDRAXx1Neq9ydOal95CHfmTnM4I+ZI2rVQf
# jXQA1WSjjf4J2a7jLzWGNqNX+DF0SQzHU0pTi4dBwp9nEC8EAqoxW6q17r0z0noD
# js6+BFo+z7bKSBwZXTRNivYuve3L2oiKNqetRHdqfMTCW/NmKLJ9M+MtucVGyOxi
# Df06VXxyKkOirv6o02OoXN4bFzK0vlNMsvhlqgF2puE6FndlENSmE+9JGYxOGLS/
# D284NHNboDGcmWXfwXRy4kbu4QFhOm0xJuF2EZAOk5eCkhSxZON3rGlHqhpB/8Ml
# uDezooIs8CVnrpHMiD2wL40mm53+/j7tFaxYKIqL0Q4ssd8xHZnIn/7GELH3IdvG
# 2XlM9q7WP/UwgOkw/HQtyRN62JK4S1C8uw3PdBunvAZapsiI5YKdvlarEvf8EA+8
# hcpSM9LHJmyrxaFtoza2zNaQ9k+5t1wwggbtMIIE1aADAgECAhAKgO8YS43xBYLR
# xHanlXRoMA0GCSqGSIb3DQEBCwUAMGkxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5E
# aWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1l
# U3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYgMjAyNSBDQTEwHhcNMjUwNjA0MDAwMDAw
# WhcNMzYwOTAzMjM1OTU5WjBjMQswCQYDVQQGEwJVUzEXMBUGA1UEChMORGlnaUNl
# cnQsIEluYy4xOzA5BgNVBAMTMkRpZ2lDZXJ0IFNIQTI1NiBSU0E0MDk2IFRpbWVz
# dGFtcCBSZXNwb25kZXIgMjAyNSAxMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIIC
# CgKCAgEA0EasLRLGntDqrmBWsytXum9R/4ZwCgHfyjfMGUIwYzKomd8U1nH7C8Dr
# 0cVMF3BsfAFI54um8+dnxk36+jx0Tb+k+87H9WPxNyFPJIDZHhAqlUPt281mHrBb
# ZHqRK71Em3/hCGC5KyyneqiZ7syvFXJ9A72wzHpkBaMUNg7MOLxI6E9RaUueHTQK
# WXymOtRwJXcrcTTPPT2V1D/+cFllESviH8YjoPFvZSjKs3SKO1QNUdFd2adw44wD
# cKgH+JRJE5Qg0NP3yiSyi5MxgU6cehGHr7zou1znOM8odbkqoK+lJ25LCHBSai25
# CFyD23DZgPfDrJJJK77epTwMP6eKA0kWa3osAe8fcpK40uhktzUd/Yk0xUvhDU6l
# vJukx7jphx40DQt82yepyekl4i0r8OEps/FNO4ahfvAk12hE5FVs9HVVWcO5J4dV
# mVzix4A77p3awLbr89A90/nWGjXMGn7FQhmSlIUDy9Z2hSgctaepZTd0ILIUbWuh
# KuAeNIeWrzHKYueMJtItnj2Q+aTyLLKLM0MheP/9w6CtjuuVHJOVoIJ/DtpJRE7C
# e7vMRHoRon4CWIvuiNN1Lk9Y+xZ66lazs2kKFSTnnkrT3pXWETTJkhd76CIDBbTR
# ofOsNyEhzZtCGmnQigpFHti58CSmvEyJcAlDVcKacJ+A9/z7eacCAwEAAaOCAZUw
# ggGRMAwGA1UdEwEB/wQCMAAwHQYDVR0OBBYEFOQ7/PIx7f391/ORcWMZUEPPYYzo
# MB8GA1UdIwQYMBaAFO9vU0rp5AZ8esrikFb2L9RJ7MtOMA4GA1UdDwEB/wQEAwIH
# gDAWBgNVHSUBAf8EDDAKBggrBgEFBQcDCDCBlQYIKwYBBQUHAQEEgYgwgYUwJAYI
# KwYBBQUHMAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBdBggrBgEFBQcwAoZR
# aHR0cDovL2NhY2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0VHJ1c3RlZEc0VGlt
# ZVN0YW1waW5nUlNBNDA5NlNIQTI1NjIwMjVDQTEuY3J0MF8GA1UdHwRYMFYwVKBS
# oFCGTmh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRHNFRp
# bWVTdGFtcGluZ1JTQTQwOTZTSEEyNTYyMDI1Q0ExLmNybDAgBgNVHSAEGTAXMAgG
# BmeBDAEEAjALBglghkgBhv1sBwEwDQYJKoZIhvcNAQELBQADggIBAGUqrfEcJwS5
# rmBB7NEIRJ5jQHIh+OT2Ik/bNYulCrVvhREafBYF0RkP2AGr181o2YWPoSHz9iZE
# N/FPsLSTwVQWo2H62yGBvg7ouCODwrx6ULj6hYKqdT8wv2UV+Kbz/3ImZlJ7YXwB
# D9R0oU62PtgxOao872bOySCILdBghQ/ZLcdC8cbUUO75ZSpbh1oipOhcUT8lD8QA
# GB9lctZTTOJM3pHfKBAEcxQFoHlt2s9sXoxFizTeHihsQyfFg5fxUFEp7W42fNBV
# N4ueLaceRf9Cq9ec1v5iQMWTFQa0xNqItH3CPFTG7aEQJmmrJTV3Qhtfparz+BW6
# 0OiMEgV5GWoBy4RVPRwqxv7Mk0Sy4QHs7v9y69NBqycz0BZwhB9WOfOu/CIJnzkQ
# TwtSSpGGhLdjnQ4eBpjtP+XB3pQCtv4E5UCSDag6+iX8MmB10nfldPF9SVD7weCC
# 3yXZi/uuhqdwkgVxuiMFzGVFwYbQsiGnoa9F5AaAyBjFBtXVLcKtapnMG3VH3EmA
# p/jsJ3FVF3+d1SVDTmjFjLbNFZUWMXuZyvgLfgyPehwJVxwC+UpX2MSey2ueIu9T
# HFVkT+um1vshETaWyQo8gmBto/m3acaP9QsuLj3FNwFlTxq25+T4QwX9xa6ILs84
# ZPvmpovq90K8eWyG2N01c4IhSOxqt81nMYIFEDCCBQwCAQEwNDAgMR4wHAYDVQQD
# DBVQaGluSVQtUFNzY3JpcHRzX1NpZ24CEHePOzJf0KCMSL6wELasExMwDQYJYIZI
# AWUDBAIBBQCggYQwGAYKKwYBBAGCNwIBDDEKMAigAoAAoQKAADAZBgkqhkiG9w0B
# CQMxDAYKKwYBBAGCNwIBBDAcBgorBgEEAYI3AgELMQ4wDAYKKwYBBAGCNwIBFTAv
# BgkqhkiG9w0BCQQxIgQgnq1Wz8cKPxFICNKcVl6wGfIOXHiOvMUV03e4oitVO4Iw
# DQYJKoZIhvcNAQEBBQAEggEAsmBjBwHNSltOk8KrLCuWZhIFCMxcNl1lNm5Jef7b
# O9FMygeOqNinv+EXtIXKY3st0+ATQh5Zz4l/AVwXC8Qjc1sB05jybspl3BFpEQme
# Kv8nGcv4Pu1K68GOnmmsNYvxWYDj2jpDT1i7mnHtrMFpkzFvdyz6FcIvpBsTS2Sq
# dpeOwmvb1xoXqZjqI2+tVEtoeYBLmxcQfi3ob+3zRTvomHonwkhANVlhX+hGvbDr
# 4m2TH7ajHDLGROWLY9S8f1lYOItxcjNxcbdVfwP4SW7AUiwDpZxI0kjQoaJH1QlJ
# 3SkwlyEozEiWYNKJbdAtFqjfm7HsFeFRaXRgoYcQHUdkgqGCAyYwggMiBgkqhkiG
# 9w0BCQYxggMTMIIDDwIBATB9MGkxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdp
# Q2VydCwgSW5jLjFBMD8GA1UEAxM4RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3Rh
# bXBpbmcgUlNBNDA5NiBTSEEyNTYgMjAyNSBDQTECEAqA7xhLjfEFgtHEdqeVdGgw
# DQYJYIZIAWUDBAIBBQCgaTAYBgkqhkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqG
# SIb3DQEJBTEPFw0yNTA3MTExOTMxNDFaMC8GCSqGSIb3DQEJBDEiBCBY4Xu91E3l
# M7rTuYA24Eqxx2/agTvd7JaPSNNLsrIWGTANBgkqhkiG9w0BAQEFAASCAgBVh3DE
# Q3o5Vhy8Z+xaGsv189+En+HazcFgp0wkdk0jQK9hH0INSlzVylxKZurnfz0PUIoL
# Kmkj/McSqR20jZJsaWD2b5ypJgetjtvmaRIiOgh6T+qQkDWlpWgg5RXkxaAAO+2n
# 7YsIgim8UwuhYHCmakjbAaMr2O7K5maBMMwlpuyW/jc4QIJ0PgplBhX9Bn17SSle
# 2h5oOk3d0cX+tjMDwKhrAEzEjDOOgYRbbFNCl3IWfFt15N9te1yWiGOifB0DLBj2
# iVq775+VmsJsLQaAJZmobf0PrYRsyk4rSzaKI0b/1HvHyhF2w7UmXgSYGFmRhLlg
# LELlNBgPquW1cKnku4OrBodObXz0gRhL9LIN9pwQw9kOkHVNAcldE9kKFZHesNE1
# FGd900OhqP9E9QiikioA/VaFOeZW97nZREvCSCJiiWoAnPEMuUEdSNZOJMroEJZ5
# hpkMuLMrJqeYkyTEXvqdIZLqp0Ry47yzg7EKNBw+558b8jLmBDKe7zpdb2Vn62t9
# iIOSKu+025qn0EcqMo5aMWnezVg6Zge7DYXZX7RZy/3IUk+2Kcu56f360TWWkXHb
# nSgR77ksOOeDL6hSi9xjOoct6iSsNughAZ89FYyzpv1KGKGUprC/RuSKs06uiUL/
# 4FMvBYkdzfxf+EK6+Vo+tiSAeR/tOQ5AKZDShQ==
# SIG # End signature block
