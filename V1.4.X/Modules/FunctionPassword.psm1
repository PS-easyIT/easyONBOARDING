# Check if required functions exist, if not create stub functions
if (-not (Get-Command -Name "Write-DebugMessage" -ErrorAction SilentlyContinue)) {
    function Write-DebugMessage { param([string]$Message) Write-Verbose $Message }
}
if (-not (Get-Command -Name "Write-Log" -ErrorAction SilentlyContinue)) {
    function Write-Log { param([string]$Message, [string]$LogLevel) Write-Verbose $Message }
}
# Sets ACL restrictions to prevent users from changing their passwords
Write-DebugMessage "Defining Set-CannotChangePassword function."
function Set-CannotChangePassword {
    # [07.1.1 - Modifies ACL settings to deny password change permissions]
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$SamAccountName
    )

    try {
        $adUser = Get-ADUser -Identity $SamAccountName -Properties DistinguishedName -ErrorAction Stop
        
        $user = [ADSI]"LDAP://$($adUser.DistinguishedName)"
        $acl = $user.psbase.ObjectSecurity

        Write-DebugMessage "Set-CannotChangePassword: Defining AccessRule"
        # Define AccessRule: SELF is not allowed to 'Change Password'
        $denyRule = New-Object System.DirectoryServices.ActiveDirectoryAccessRule(
            [System.Security.Principal.NTAccount]"NT AUTHORITY\\SELF",
            [System.DirectoryServices.ActiveDirectoryRights]::WriteProperty,
            [System.Security.AccessControl.AccessControlType]::Deny,
            [GUID]"ab721a53-1e2f-11d0-9819-00aa0040529b"  # GUID for 'User-Change-Password'
        )
        $acl.AddAccessRule($denyRule)
        $user.psbase.ObjectSecurity = $acl
        $user.psbase.CommitChanges()
        Write-Log "Prevent Password Change has been set for $($PSBoundParameters.SamAccountName)." "DEBUG"
    }
    catch {
        Write-Warning "Error setting password change restriction for $($PSBoundParameters.SamAccountName): $($_.Exception.Message)"
    }
}
#endregion

Write-DebugMessage "Defining Remove-CannotChangePassword function."

# Removes ACL restrictions to allow users to change their passwords
function Remove-CannotChangePassword {
    # [07.2.1 - Removes deny rules from user ACL for password change permission]
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$SamAccountName
    )

    try {
        $adUser = Get-ADUser -Identity $SamAccountName -Properties DistinguishedName -ErrorAction Stop

        $user = [ADSI]"LDAP://$($adUser.DistinguishedName)"
        $acl = $user.psbase.ObjectSecurity

        Write-DebugMessage "Remove-CannotChangePassword: Removing all deny rules"
        # Remove all deny rules that affect SELF and have GUID ab721a53-1e2f-11d0-9819-00aa0040529b
        $rulesToRemove = $acl.Access | Where-Object {
            $_.IdentityReference -eq "NT AUTHORITY\\SELF" -and
            $_.AccessControlType -eq 'Deny' -and
            $_.ObjectType -eq "ab721a53-1e2f-11d0-9819-00aa0040529b"
        }
        foreach ($rule in $rulesToRemove) {
            $acl.RemoveAccessRule($rule) | Out-Null
        }
        $user.psbase.ObjectSecurity = $acl
        $user.psbase.CommitChanges()
        Write-Log "Prevent Password Change has been removed for $($PSBoundParameters.SamAccountName)." "DEBUG"
    }
    catch {
        Write-Warning "Error removing password change restriction for $($PSBoundParameters.SamAccountName): $($_.Exception.Message)"
    }
}
# Export the functions
Export-ModuleMember -Function Set-CannotChangePassword, Remove-CannotChangePassword
