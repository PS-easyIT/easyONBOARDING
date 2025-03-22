Write-DebugMessage "Set-Logo"

# Function to handle logo uploads and management for reports
Write-DebugMessage "Setting logo."
function Set-Logo {
    # [15.1.1 - Handles file selection and saving of logo images]
    param (
        [Parameter(Mandatory=$true)]
        [hashtable]$brandingConfig
    )
    try {
        Write-DebugMessage "GUI Opening file selection dialog for logo upload"
        $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
        $openFileDialog.Filter = "Image files (*.jpg;*.png;*.bmp)|*.jpg;*.jpeg;*.png;*.bmp"
        $openFileDialog.Title = "Select a logo for the onboarding document"
        if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $selectedFilePath = $openFileDialog.FileName
            Write-Log "Selected file: $selectedFilePath" "DEBUG"

            if (-not $brandingConfig.Contains("TemplateLogo") -or [string]::IsNullOrWhiteSpace($brandingConfig["TemplateLogo"])) {
                Throw "No 'TemplateLogo' defined in the Report section."
            }

            $TemplateLogo = $brandingConfig["TemplateLogo"]
            if (-not (Test-Path $TemplateLogo)) {
                try {
                    New-Item -ItemType Directory -Path $TemplateLogo -Force -ErrorAction Stop | Out-Null
                    Write-Log "Created directory: $TemplateLogo" "DEBUG"
                } catch {
                    Throw "Could not create target directory for logo: $($_.Exception.Message)"
                }
            }

            $targetLogoTemplate = Join-Path -Path $TemplateLogo -ChildPath "ReportLogo.png"
            Copy-Item -Path $selectedFilePath -Destination $targetLogoTemplate -Force -ErrorAction Stop
            Write-Log "Logo successfully saved at: $targetLogoTemplate" "DEBUG"
            [System.Windows.MessageBox]::Show("The logo was successfully saved!`nLocation: $targetLogoTemplate", "Success", "OK", "Information")
        }
    } catch {
        Write-Log "Error in Set-Logo: $($_.Exception.Message)" "ERROR"
        [System.Windows.MessageBox]::Show("Error uploading logo: $($_.Exception.Message)", "Error", "OK", "Error")
    }
}
#endregion

Write-DebugMessage "Defining advanced password generation function."

# Advanced password generation with security requirements
Write-DebugMessage "Generating advanced password."
function New-AdvancedPassword {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$false)]
        [ValidateRange(8, 100)]
        [int]$Length = 12,

        [Parameter(Mandatory=$false)]
        [ValidateRange(1, 10)]
        [int]$MinUpperCase = 2,

        [Parameter(Mandatory=$false)]
        [ValidateRange(1, 10)]
        [int]$MinDigits = 2,

        [Parameter(Mandatory=$false)]
        [bool]$AvoidAmbiguous = $false,

        [Parameter(Mandatory=$false)]
        [bool]$IncludeSpecial = $true,

        [Parameter(Mandatory=$false)]
        [ValidateRange(1,10)]
        [int]$MinNonAlpha = 2
    )

    # Validate that the password length is sufficient for the requirements
    $minimumRequiredLength = $MinUpperCase + $MinDigits + $MinNonAlpha
    if ($Length -lt $minimumRequiredLength) {
        Throw "Error: The password length ($Length) is too short for the required minimum values (MinUpperCase + MinDigits + MinNonAlpha = $minimumRequiredLength)."
    }

    Write-DebugMessage "New-AdvancedPassword: Defining character pools"
    $lower = 'abcdefghijklmnopqrstuvwxyz'
    $upper = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
    $digits = '0123456789'
    $special = '!@#$%^&*()'
    
    Write-DebugMessage "New-AdvancedPassword: Removing ambiguous characters if desired"
    if ($AvoidAmbiguous) {
        $ambiguous = 'Il1O0'
        $upper = -join ($upper.ToCharArray() | Where-Object { $ambiguous -notcontains $_ })
        $lower = -join ($lower.ToCharArray() | Where-Object { $ambiguous -notcontains $_ })
        $digits = -join ($digits.ToCharArray() | Where-Object { $ambiguous -notcontains $_ })
        $special = -join ($special.ToCharArray() | Where-Object { $ambiguous -notcontains $_ })
    }
    
    # Ensure that character pools are never empty
    if ([string]::IsNullOrEmpty($upper)) { $upper = 'ABCDEFGHJKLMNPQRSTUVWXYZ' }
    if ([string]::IsNullOrEmpty($lower)) { $lower = 'abcdefghijkmnopqrstuvwxyz' }
    if ([string]::IsNullOrEmpty($digits)) { $digits = '23456789' }
    if ([string]::IsNullOrEmpty($special)) { $special = '!@#$%^&*()' }
    
    # Recalculate 'all' after ensuring pools are not empty
    $all = $lower + $upper + $digits
    if ($IncludeSpecial) { $all += $special }
    
    do {
        Write-DebugMessage "New-AdvancedPassword: Starting password generation"
        # Initialize as a simple string array
        $passwordChars = [System.Collections.ArrayList]::new()
        
        Write-DebugMessage "New-AdvancedPassword: Adding minimum number of uppercase letters"
        for ($i = 0; $i -lt $MinUpperCase; $i++) {
            [void]$passwordChars.Add($upper[(Get-Random -Minimum 0 -Maximum $upper.Length)].ToString())
        }
        
        Write-DebugMessage "New-AdvancedPassword: Adding minimum number of digits"
        for ($i = 0; $i -lt $MinDigits; $i++) {
            [void]$passwordChars.Add($digits[(Get-Random -Minimum 0 -Maximum $digits.Length)].ToString())
        }
        
        # Add special characters if needed to meet MinNonAlpha requirement
        if ($IncludeSpecial -and ($MinNonAlpha -gt $MinDigits)) {
            $specialCharsNeeded = $MinNonAlpha - $MinDigits
            for ($i = 0; $i -lt $specialCharsNeeded; $i++) {
                [void]$passwordChars.Add($special[(Get-Random -Minimum 0 -Maximum $special.Length)].ToString())
            }
        }
        
        Write-DebugMessage "New-AdvancedPassword: Filling up to desired length"
        while ($passwordChars.Count -lt $Length) {
            [void]$passwordChars.Add($all[(Get-Random -Minimum 0 -Maximum $all.Length)].ToString())
        }
        
        Write-DebugMessage "New-AdvancedPassword: Randomizing order"
        # Get array of strings, then join them at the end
        $shuffledChars = $passwordChars | Get-Random -Count $passwordChars.Count
        $generatedPassword = -join $shuffledChars
        
        Write-DebugMessage "New-AdvancedPassword: Checking minimum number of non-alphabetic characters"
        # Count characters that don't match letters
        $nonAlphaCount = ($generatedPassword.ToCharArray() | Where-Object { $_ -notmatch '[a-zA-Z]' }).Count
        
    } while ($nonAlphaCount -lt $MinNonAlpha)
    
    Write-DebugMessage "Advanced password generated successfully."
    return $generatedPassword
}

# Export the module members
Export-ModuleMember -Function Set-Logo, New-AdvancedPassword
