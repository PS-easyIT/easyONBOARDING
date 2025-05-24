<%@ Application Language="C#" %>
<%@ Import Namespace="System.IO" %>
<%@ Import Namespace="System.Web.Routing" %>

<script runat="server">
    void Application_Start(object sender, EventArgs e)
    {
        // Initialize application
        string dataFolder = Path.Combine(HttpRuntime.AppDomainAppPath, "App_Data");
        if (!Directory.Exists(dataFolder))
        {
            Directory.CreateDirectory(dataFolder);
        }
        
        // Create logs folder
        string logFolder = Path.Combine(dataFolder, "Logs");
        if (!Directory.Exists(logFolder))
        {
            Directory.CreateDirectory(logFolder);
        }
        
        // Create backups folder
        string backupFolder = Path.Combine(dataFolder, "Backups");
        if (!Directory.Exists(backupFolder))
        {
            Directory.CreateDirectory(backupFolder);
        }
        
        // Create scripts folder and copy PowerShell scripts if needed
        string scriptsFolder = Path.Combine(HttpRuntime.AppDomainAppPath, "Scripts");
        if (!Directory.Exists(scriptsFolder))
        {
            Directory.CreateDirectory(scriptsFolder);
        }
        
        // Log application start
        string logFile = Path.Combine(logFolder, $"Application_{DateTime.Now:yyyyMMdd}.log");
        using (StreamWriter writer = new StreamWriter(logFile, true))
        {
            writer.WriteLine($"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - INFO - Application started");
        }
    }
    
    void Application_End(object sender, EventArgs e)
    {
        // Log application end
        string logFolder = Path.Combine(HttpRuntime.AppDomainAppPath, "App_Data", "Logs");
        if (Directory.Exists(logFolder))
        {
            string logFile = Path.Combine(logFolder, $"Application_{DateTime.Now:yyyyMMdd}.log");
            using (StreamWriter writer = new StreamWriter(logFile, true))
            {
                writer.WriteLine($"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - INFO - Application stopped");
            }
        }
    }
    
    void Application_Error(object sender, EventArgs e)
    {
        // Log unhandled exceptions
        Exception ex = Server.GetLastError();
        if (ex != null)
        {
            string logFolder = Path.Combine(HttpRuntime.AppDomainAppPath, "App_Data", "Logs");
            if (Directory.Exists(logFolder))
            {
                string logFile = Path.Combine(logFolder, $"Error_{DateTime.Now:yyyyMMdd}.log");
                using (StreamWriter writer = new StreamWriter(logFile, true))
                {
                    writer.WriteLine($"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - ERROR - Unhandled exception: {ex.Message}");
                    writer.WriteLine($"Stack trace: {ex.StackTrace}");
                    if (ex.InnerException != null)
                    {
                        writer.WriteLine($"Inner exception: {ex.InnerException.Message}");
                    }
                }
            }
        }
    }
</script>
