# easyONBOARDING Web Application Deployment Guide

This document provides instructions for deploying the easyONBOARDING web application on an IIS web server.

## Prerequisites

- Windows Server with IIS installed (version 8.0 or later)
- .NET Framework 4.8 installed
- ASP.NET registered with IIS
- For PowerShell integration: PowerShell 5.1 or later installed on the server
- IIS URL Rewrite Module installed

## Deployment Steps

1. **Create an Application Pool**
   - Open IIS Manager
   - Right-click on "Application Pools" and select "Add Application Pool"
   - Name: `easyONBOARDING`
   - .NET CLR Version: `.NET CLR v4.0.30319`
   - Managed pipeline mode: `Integrated`
   - Click OK

2. **Create a Web Application**
   - In IIS Manager, right-click on "Sites" and select "Add Website" (or add as an application under an existing site)
   - Site name: `easyONBOARDING`
   - Application pool: Select the `easyONBOARDING` pool created earlier
   - Physical path: Path to the application folder
   - Binding: Configure as needed (e.g., hostname, port)
   - Click OK

3. **Install Required IIS Modules**
   - Ensure URL Rewrite Module is installed 
   - If needed, download from: https://www.iis.net/downloads/microsoft/url-rewrite

4. **Configure Folder Permissions**
   - Grant the application pool identity (typically `IIS AppPool\easyONBOARDING`) read/write access to the application folders, especially:
     - The entire web application directory
     - The `App_Data` subdirectory

5. **Configure Authentication**
   - In IIS Manager, select the application, then open "Authentication"
   - Enable "Windows Authentication" (for Active Directory integration)
   - Disable "Anonymous Authentication" if you want to require login

6. **Directory Structure**
   Ensure your deployment has the following structure:
   ```
   /easyONBOARDING/
   ├── App_Data/           # Data storage (will be created automatically)
   │   ├── Backups/        # CSV backups
   │   └── Logs/           # Application logs
   ├── Scripts/            # PowerShell scripts
   │   └── easyONB_HR-AL_V0.1.1.ps1  # PowerShell backend script
   ├── api-handler.ashx     # API handler
   ├── easyONBOARDING.html  # Main application file
   ├── Global.asax          # Application lifecycle events
   ├── web.config           # IIS configuration
   ├── Web.Debug.config     # Debug configuration
   └── [other static files]
   ```

7. **Test the Deployment**
   - Open a web browser and navigate to your configured URL
   - You should see the login screen
   - Test the authentication and basic functionality

## Troubleshooting

1. **Application Not Starting**
   - Check the application pool is running
   - Verify .NET Framework 4.8 is installed
   - Check Application Event Log for errors

2. **API Errors**
   - Check IIS logs in `%SystemDrive%\inetpub\logs\LogFiles`
   - Check application logs in the `App_Data\Logs` folder
   - Ensure the ASP.NET Core Module is installed if using that feature

3. **Authentication Issues**
   - Verify Windows Authentication is enabled in IIS
   - Check that the application pool identity has appropriate permissions

4. **PowerShell Integration Issues**
   - Ensure PowerShell execution policy allows script execution
   - Verify the application pool identity has rights to execute PowerShell scripts

## Security Considerations

1. **Data Protection**
   - The application stores sensitive employee data in CSV files
   - These files are protected with basic encryption
   - For production use, consider implementing stronger encryption

2. **Authentication**
   - The application uses Windows Authentication by default
   - For internet-facing deployments, consider adding SSL/TLS certificates

3. **Authorization**
   - Role-based access control is implemented
   - Users are assigned roles based on their AD group membership

## Maintenance

1. **Backups**
   - CSV data is automatically backed up before changes
   - Backup files are stored in `App_Data\Backups`
   - Consider scheduling regular filesystem backups of the entire `App_Data` directory

2. **Logging**
   - Application logs are written to `App_Data\Logs`
   - Consider implementing a log rotation policy for long-term use
