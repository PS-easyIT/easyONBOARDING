#region [Region 01 | MODULE INFORMATION]
<#
.SYNOPSIS
    PDF Generator module for easyONBOARDING
.DESCRIPTION
    Provides functions to convert HTML files to PDF
.NOTES
    Version: 1.0
    Author: easyIT Systems
#>
#endregion

#region [Region 02 | PREREQUISITES]
# Check if running in PowerShell 5.1 or later
if ($PSVersionTable.PSVersion.Major -lt 5) {
    Write-Error "This module requires PowerShell 5.1 or later."
    return
}

# Add assembly references if needed
try {
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
} catch {
    Write-Warning "Failed to load required assemblies: $($_.Exception.Message)"
}
#endregion

#region [Region 03 | PRIMARY CONVERSION FUNCTION]
function Convert-HTMLToPDF {
    <#
    .SYNOPSIS
        Converts an HTML file to PDF using wkhtmltopdf
    .DESCRIPTION
        Uses the wkhtmltopdf utility to convert HTML to PDF with configurable options
    .PARAMETER HtmlFile
        Path to the HTML file to convert
    .PARAMETER PdfFile
        Path where the PDF file should be saved
    .PARAMETER WkhtmltopdfPath
        Optional path to the wkhtmltopdf executable
    .PARAMETER Options
        Optional parameters to pass to wkhtmltopdf
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$HtmlFile,
        
        [Parameter(Mandatory = $true)]
        [string]$PdfFile,
        
        [Parameter(Mandatory = $false)]
        [string]$WkhtmltopdfPath,
        
        [Parameter(Mandatory = $false)]
        [string]$Options = "--print-media-type --no-background"
    )
    
    try {
        # Verify HTML file exists
        if (-not (Test-Path $HtmlFile)) {
            throw "HTML file not found: $HtmlFile"
        }
        
        # Ensure output directory exists
        $pdfDir = Split-Path -Path $PdfFile -Parent
        if (-not (Test-Path $pdfDir)) {
            New-Item -Path $pdfDir -ItemType Directory -Force | Out-Null
        }
        
        # Find wkhtmltopdf if not specified
        if (-not $WkhtmltopdfPath -or -not (Test-Path $WkhtmltopdfPath)) {
            # Look in common locations
            $possiblePaths = @(
                "${env:ProgramFiles}\wkhtmltopdf\bin\wkhtmltopdf.exe",
                "${env:ProgramFiles(x86)}\wkhtmltopdf\bin\wkhtmltopdf.exe",
                "C:\wkhtmltopdf\bin\wkhtmltopdf.exe"
            )
            
            # Check if we have a global config variable available
            if ($global:Config -and $global:Config.ContainsKey("Report") -and $global:Config["Report"].ContainsKey("wkhtmltopdfPath")) {
                $possiblePaths += $global:Config["Report"]["wkhtmltopdfPath"]
            }
            
            foreach ($path in $possiblePaths) {
                if (Test-Path $path) {
                    $WkhtmltopdfPath = $path
                    break
                }
            }
            
            if (-not $WkhtmltopdfPath -or -not (Test-Path $WkhtmltopdfPath)) {
                throw "wkhtmltopdf executable not found. Please specify the path using the WkhtmltopdfPath parameter."
            }
        }
        
        # Build the arguments
        $optionsArray = $Options -split ' '
        $arguments = @()
        $arguments += $optionsArray
        $arguments += "`"$HtmlFile`""
        $arguments += "`"$PdfFile`""
        
        # Execute the command
        Write-Verbose "Executing: $WkhtmltopdfPath $($arguments -join ' ')"
        $processInfo = New-Object System.Diagnostics.ProcessStartInfo
        $processInfo.FileName = $WkhtmltopdfPath
        $processInfo.RedirectStandardError = $true
        $processInfo.RedirectStandardOutput = $true
        $processInfo.UseShellExecute = $false
        $processInfo.Arguments = $arguments -join ' '
        
        $process = New-Object System.Diagnostics.Process
        $process.StartInfo = $processInfo
        $process.Start() | Out-Null
        $process.WaitForExit()
        
        $stdout = $process.StandardOutput.ReadToEnd()
        $stderr = $process.StandardError.ReadToEnd()
        
        if ($process.ExitCode -ne 0) {
            throw "wkhtmltopdf exited with code $($process.ExitCode). Error: $stderr"
        }
        
        if (-not (Test-Path $PdfFile)) {
            throw "PDF file was not created. Process completed but output file is missing."
        }
        
        return $true
    }
    catch {
        Write-Error "Failed to convert HTML to PDF: $($_.Exception.Message)"
        return $false
    }
}
#endregion

#region [Region 04 | HTML HELPER FUNCTIONS]
function Export-UserToHTML {
    <#
    .SYNOPSIS
        Exports user data to an HTML report
    .DESCRIPTION
        Creates an HTML document with user information for onboarding
    .PARAMETER userData
        User data object with properties
    .PARAMETER result
        Result object from the onboarding process
    .PARAMETER htmlPath
        Path where the HTML report should be saved
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [PSObject]$userData,
        
        [Parameter(Mandatory = $true)]
        [PSObject]$result,
        
        [Parameter(Mandatory = $true)]
        [string]$htmlPath
    )
    
    try {
        # Create HTML template with styling
        $htmlTemplate = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>User Onboarding Report</title>
    <style>
        body { font-family: Arial, sans-serif; line-height: 1.6; margin: 0; padding: 20px; }
        .container { max-width: 800px; margin: 0 auto; border: 1px solid #ddd; padding: 20px; }
        .header { text-align: center; margin-bottom: 20px; }
        .logo { max-width: 200px; height: auto; }
        h1 { color: #2c3e50; }
        .section { margin-bottom: 20px; }
        .section h2 { color: #3498db; border-bottom: 1px solid #eee; padding-bottom: 5px; }
        table { width: 100%; border-collapse: collapse; }
        th, td { padding: 8px; text-align: left; border-bottom: 1px solid #ddd; }
        th { background-color: #f2f2f2; }
        .footer { margin-top: 30px; text-align: center; font-size: 0.8em; color: #7f8c8d; }
        @media print {
            body { padding: 0; }
            .container { border: none; }
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>User Onboarding Report</h1>
            {LOGO_PLACEHOLDER}
        </div>

        <div class="section">
            <h2>User Information</h2>
            <table>
                <tr><th>Full Name</th><td>{DISPLAY_NAME}</td></tr>
                <tr><th>Username</th><td>{USERNAME}</td></tr>
                <tr><th>UPN</th><td>{UPN}</td></tr>
                <tr><th>Email Address</th><td>{EMAIL}</td></tr>
                <tr><th>Department</th><td>{DEPARTMENT}</td></tr>
                <tr><th>Position</th><td>{POSITION}</td></tr>
            </table>
        </div>

        <div class="section">
            <h2>Contact Information</h2>
            <table>
                <tr><th>Office Phone</th><td>{PHONE}</td></tr>
                <tr><th>Mobile Phone</th><td>{MOBILE}</td></tr>
                <tr><th>Office Location</th><td>{OFFICE}</td></tr>
            </table>
        </div>

        <div class="section">
            <h2>Account Information</h2>
            <table>
                <tr><th>Account Status</th><td>{ACCOUNT_STATUS}</td></tr>
                <tr><th>Password Settings</th><td>{PASSWORD_SETTINGS}</td></tr>
                <tr><th>Groups</th><td>{GROUPS}</td></tr>
            </table>
        </div>

        <div class="section">
            <h2>Company Information</h2>
            <table>
                <tr><th>Company Name</th><td>{COMPANY_NAME}</td></tr>
                <tr><th>Company Address</th><td>{COMPANY_ADDRESS}</td></tr>
                <tr><th>Company Phone</th><td>{COMPANY_PHONE}</td></tr>
            </table>
        </div>

        <div class="footer">
            <p>Generated on {CURRENT_DATE} by easyONBOARDING</p>
        </div>
    </div>
</body>
</html>
"@

        # Company logo handling
        $logoHtml = ""
        if ($global:Config -and $global:Config.ContainsKey("Report") -and $global:Config["Report"].ContainsKey("CompanyLogo")) {
            $logoPath = $global:Config["Report"]["CompanyLogo"]
            if (Test-Path $logoPath) {
                $logoHtml = "<img src='$logoPath' alt='Company Logo' class='logo'>"
            }
        }
        
        # Replace placeholders with actual values
        $htmlContent = $htmlTemplate -replace "{LOGO_PLACEHOLDER}", $logoHtml
        $htmlContent = $htmlContent -replace "{DISPLAY_NAME}", $userData.DisplayName
        $htmlContent = $htmlContent -replace "{USERNAME}", $result.SamAccountName
        $htmlContent = $htmlContent -replace "{UPN}", $result.UPN
        $htmlContent = $htmlContent -replace "{EMAIL}", $userData.EmailAddress
        $htmlContent = $htmlContent -replace "{DEPARTMENT}", $userData.DepartmentField
        $htmlContent = $htmlContent -replace "{POSITION}", $userData.Position
        $htmlContent = $htmlContent -replace "{PHONE}", $userData.PhoneNumber
        $htmlContent = $htmlContent -replace "{MOBILE}", $userData.MobileNumber
        $htmlContent = $htmlContent -replace "{OFFICE}", $userData.OfficeRoom
        
        $accountStatus = if ($userData.AccountDisabled) { "Disabled" } else { "Enabled" }
        $passwordSettings = "Must Change: $($userData.MustChangePassword), Never Expires: $($userData.PasswordNeverExpires)"
        
        $htmlContent = $htmlContent -replace "{ACCOUNT_STATUS}", $accountStatus
        $htmlContent = $htmlContent -replace "{PASSWORD_SETTINGS}", $passwordSettings
        
        # Format groups as list
        $groupsList = if ($userData.ADGroupsSelected -eq "NONE") { 
            "No groups assigned" 
        } else { 
            $groups = $userData.ADGroupsSelected -split ','
            "<ul>" + ($groups | ForEach-Object { "<li>$_</li>" }) + "</ul>"
        }
        $htmlContent = $htmlContent -replace "{GROUPS}", $groupsList
        
        # Company info
        $htmlContent = $htmlContent -replace "{COMPANY_NAME}", $userData.CompanyName
        $companyAddress = "$($userData.CompanyStrasse), $($userData.CompanyPLZ) $($userData.CompanyOrt), $($userData.CompanyCountry)"
        $htmlContent = $htmlContent -replace "{COMPANY_ADDRESS}", $companyAddress
        $htmlContent = $htmlContent -replace "{COMPANY_PHONE}", $userData.CompanyTelefon
        
        # Date
        $htmlContent = $htmlContent -replace "{CURRENT_DATE}", (Get-Date -Format "yyyy-MM-dd HH:mm")
        
        # Ensure the directory exists
        $directory = Split-Path -Path $htmlPath -Parent
        if (-not (Test-Path $directory)) {
            New-Item -Path $directory -ItemType Directory -Force | Out-Null
        }
        
        # Save the HTML file
        $htmlContent | Out-File -FilePath $htmlPath -Encoding UTF8
        
        if (Test-Path $htmlPath) {
            return $true
        } else {
            throw "Failed to create HTML file"
        }
    }
    catch {
        Write-Error "Failed to export user data to HTML: $($_.Exception.Message)"
        return $false
    }
}
#endregion

# Export the module functions
Export-ModuleMember -Function Convert-HTMLToPDF, Export-UserToHTML
