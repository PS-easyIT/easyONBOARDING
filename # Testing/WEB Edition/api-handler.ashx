<%@ WebHandler Language="C#" Class="EasyOnboarding.ApiHandler" %>

using System;
using System.IO;
using System.Web;
using System.Text;
using System.Linq;
using System.Xml;
using System.Net;
using System.Collections.Generic;
using System.Security.Principal;
using System.DirectoryServices.AccountManagement;
using System.Web.Script.Serialization;
using System.Security.Cryptography;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using System.Configuration;

namespace EasyOnboarding
{
    /// <summary>
    /// API handler for easyONBOARDING requests
    /// </summary>
    public class ApiHandler : IHttpHandler
    {
        // Data storage paths 
        private readonly string dataFolder = Path.Combine(HttpRuntime.AppDomainAppPath, "App_Data");
        private readonly string csvFile;
        private readonly string auditLogFile;
        private readonly string backupFolder;
        private readonly string powershellScriptPath;

        // Simple in-memory token storage (would use a more robust solution in production)
        private static Dictionary<string, UserSession> activeSessions = new Dictionary<string, UserSession>();
        private const int TOKEN_EXPIRATION_MINUTES = 60;

        public ApiHandler()
        {
            // Create data folder if needed
            if (!Directory.Exists(dataFolder))
            {
                Directory.CreateDirectory(dataFolder);
            }

            csvFile = Path.Combine(dataFolder, "HROnboardingData.csv");
            auditLogFile = Path.Combine(dataFolder, "AuditLog.csv");
            backupFolder = Path.Combine(dataFolder, "Backups");
            powershellScriptPath = Path.Combine(HttpRuntime.AppDomainAppPath, "Scripts", "easyONB_HR-AL_V0.1.1.ps1");

            if (!Directory.Exists(backupFolder))
            {
                Directory.CreateDirectory(backupFolder);
            }
        }

        public void ProcessRequest(HttpContext context)
        {
            try
            {
                // Set appropriate content type
                context.Response.ContentType = "application/json";
                
                // Handle OPTIONS requests for CORS
                if (context.Request.HttpMethod == "OPTIONS")
                {
                    HandleCorsOptions(context);
                    return;
                }

                string path = context.Request.QueryString["path"];
                string[] pathSegments = (path ?? "").Split(new char[] { '/' }, StringSplitOptions.RemoveEmptyEntries);

                // Check for authentication exception paths
                if (IsAuthExceptionPath(path))
                {
                    ProcessApiRequest(context, path, pathSegments);
                    return;
                }

                // Validate authentication for all other paths
                if (!IsAuthenticated(context))
                {
                    RespondWithError(context, "Unauthorized", 401);
                    return;
                }

                // Process the request based on the path
                ProcessApiRequest(context, path, pathSegments);
            }
            catch (Exception ex)
            {
                LogError("API handler error: " + ex.Message);
                RespondWithError(context, "Server error: " + ex.Message, 500);
            }
        }

        private bool IsAuthenticated(HttpContext context)
        {
            // Check for auth token in headers
            string authHeader = context.Request.Headers["Authorization"];
            if (string.IsNullOrEmpty(authHeader) || !authHeader.StartsWith("Bearer "))
            {
                return false;
            }

            string token = authHeader.Substring("Bearer ".Length).Trim();
            if (string.IsNullOrEmpty(token))
            {
                return false;
            }

            // Validate token
            lock (activeSessions)
            {
                if (activeSessions.ContainsKey(token))
                {
                    var session = activeSessions[token];
                    if (session.Expiration > DateTime.Now)
                    {
                        // Extend token if still valid
                        session.Expiration = DateTime.Now.AddMinutes(TOKEN_EXPIRATION_MINUTES);
                        return true;
                    }
                    else
                    {
                        // Token expired
                        activeSessions.Remove(token);
                    }
                }
            }

            return false;
        }

        private bool IsAuthExceptionPath(string path)
        {
            // Paths that don't require authentication
            return path != null && (
                path.StartsWith("login", StringComparison.OrdinalIgnoreCase) ||
                path.StartsWith("validateToken", StringComparison.OrdinalIgnoreCase)
            );
        }

        private void ProcessApiRequest(HttpContext context, string path, string[] pathSegments)
        {
            if (string.IsNullOrEmpty(path))
            {
                RespondWithError(context, "Invalid API path", 400);
                return;
            }

            // Route the request based on path
            if (path.StartsWith("login", StringComparison.OrdinalIgnoreCase))
            {
                HandleLogin(context);
            }
            else if (path.StartsWith("validateToken", StringComparison.OrdinalIgnoreCase))
            {
                HandleValidateToken(context);
            }
            else if (path.StartsWith("records", StringComparison.OrdinalIgnoreCase))
            {
                if (pathSegments.Length > 1)
                {
                    // Single record operations: /records/{id}
                    string recordId = pathSegments[1];
                    
                    if (context.Request.HttpMethod == "GET")
                    {
                        GetRecordById(context, recordId);
                    }
                    else if (context.Request.HttpMethod == "PUT")
                    {
                        UpdateRecord(context, recordId);
                    }
                    else if (context.Request.HttpMethod == "DELETE")
                    {
                        DeleteRecord(context, recordId);
                    }
                    else
                    {
                        RespondWithError(context, "Method not allowed", 405);
                    }
                }
                else
                {
                    // Collection operations: /records
                    if (context.Request.HttpMethod == "GET")
                    {
                        GetRecords(context);
                    }
                    else if (context.Request.HttpMethod == "POST")
                    {
                        CreateRecord(context);
                    }
                    else
                    {
                        RespondWithError(context, "Method not allowed", 405);
                    }
                }
            }
            else if (path.StartsWith("managers", StringComparison.OrdinalIgnoreCase))
            {
                GetManagers(context);
            }
            else if (path.StartsWith("auditlog", StringComparison.OrdinalIgnoreCase))
            {
                GetAuditLog(context);
            }
            else
            {
                RespondWithError(context, "Unknown API endpoint", 404);
            }
        }

        #region Authentication Handlers

        private void HandleLogin(HttpContext context)
        {
            if (context.Request.HttpMethod != "POST")
            {
                RespondWithError(context, "Method not allowed", 405);
                return;
            }

            try
            {
                // Read the request body
                string requestBody;
                using (StreamReader reader = new StreamReader(context.Request.InputStream, context.Request.ContentEncoding))
                {
                    requestBody = reader.ReadToEnd();
                }

                // Parse the JSON
                JavaScriptSerializer serializer = new JavaScriptSerializer();
                dynamic loginData = serializer.Deserialize<Dictionary<string, object>>(requestBody);

                string username = (string)loginData["username"];
                string password = (string)loginData["password"];
                bool rememberMe = loginData.ContainsKey("rememberMe") ? (bool)loginData["rememberMe"] : false;

                // For demo: Check credentials against Windows authentication
                bool isAuthenticated = false;
                string userRole = "User"; // Default role

                try
                {
                    using (PrincipalContext pc = new PrincipalContext(ContextType.Domain))
                    {
                        isAuthenticated = pc.ValidateCredentials(username, password);
                        
                        if (isAuthenticated)
                        {
                            // Determine user role based on group membership
                            UserPrincipal user = UserPrincipal.FindByIdentity(pc, username);
                            if (user != null)
                            {
                                if (IsMemberOfGroup(user, "HR Department"))
                                {
                                    userRole = "HR";
                                }
                                else if (IsMemberOfGroup(user, "IT Support"))
                                {
                                    userRole = "IT";
                                }
                                else if (IsMemberOfGroup(user, "Department Managers"))
                                {
                                    userRole = "Manager";
                                }
                                else if (IsMemberOfGroup(user, "System Administrators"))
                                {
                                    userRole = "Admin";
                                }
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    LogError("Authentication error: " + ex.Message);
                    
                    // Fallback to a simple test user check for development/testing
                    if (IsTestUser(username, password))
                    {
                        isAuthenticated = true;
                        userRole = GetTestUserRole(username);
                    }
                }

                if (isAuthenticated)
                {
                    // Generate a token
                    string token = GenerateToken();
                    
                    // Store in session
                    lock (activeSessions)
                    {
                        activeSessions[token] = new UserSession
                        {
                            Username = username,
                            Role = userRole,
                            Expiration = DateTime.Now.AddMinutes(rememberMe ? TOKEN_EXPIRATION_MINUTES * 24 : TOKEN_EXPIRATION_MINUTES)
                        };
                    }

                    // Log successful login
                    LogAudit("Login", username, "User login successful");

                    // Return the token
                    var response = new
                    {
                        token = token,
                        username = username,
                        role = userRole
                    };

                    RespondWithJson(context, serializer.Serialize(response));
                }
                else
                {
                    // Log failed login attempt
                    LogAudit("LoginFailed", username, "Invalid login attempt");
                    RespondWithError(context, "Invalid username or password", 401);
                }
            }
            catch (Exception ex)
            {
                LogError("Login error: " + ex.Message);
                RespondWithError(context, "Login failed due to server error", 500);
            }
        }

        private void HandleValidateToken(HttpContext context)
        {
            if (context.Request.HttpMethod != "GET")
            {
                RespondWithError(context, "Method not allowed", 405);
                return;
            }

            try
            {
                string authHeader = context.Request.Headers["Authorization"];
                if (string.IsNullOrEmpty(authHeader) || !authHeader.StartsWith("Bearer "))
                {
                    RespondWithError(context, "Invalid authorization header", 401);
                    return;
                }

                string token = authHeader.Substring("Bearer ".Length).Trim();
                
                UserSession session = null;
                lock (activeSessions)
                {
                    if (activeSessions.ContainsKey(token) && activeSessions[token].Expiration > DateTime.Now)
                    {
                        session = activeSessions[token];
                        session.Expiration = DateTime.Now.AddMinutes(TOKEN_EXPIRATION_MINUTES); // Extend token
                    }
                }

                if (session != null)
                {
                    var response = new
                    {
                        username = session.Username,
                        role = session.Role,
                        valid = true
                    };

                    JavaScriptSerializer serializer = new JavaScriptSerializer();
                    RespondWithJson(context, serializer.Serialize(response));
                }
                else
                {
                    RespondWithError(context, "Invalid or expired token", 401);
                }
            }
            catch (Exception ex)
            {
                LogError("Token validation error: " + ex.Message);
                RespondWithError(context, "Token validation failed", 500);
            }
        }

        #endregion

        #region Record Management Handlers

        private void GetRecords(HttpContext context)
        {
            try
            {
                // Get query string parameters for filtering
                string state = context.Request.QueryString["state"];
                string search = context.Request.QueryString["search"];
                string department = context.Request.QueryString["department"];
                DateTime? fromDate = null;
                DateTime? toDate = null;

                if (!string.IsNullOrEmpty(context.Request.QueryString["fromDate"]))
                {
                    DateTime parsedDate;
                    if (DateTime.TryParse(context.Request.QueryString["fromDate"], out parsedDate))
                    {
                        fromDate = parsedDate;
                    }
                }

                if (!string.IsNullOrEmpty(context.Request.QueryString["toDate"]))
                {
                    DateTime parsedDate;
                    if (DateTime.TryParse(context.Request.QueryString["toDate"], out parsedDate))
                    {
                        toDate = parsedDate;
                    }
                }

                // Get current user info from token
                string token = context.Request.Headers["Authorization"].Substring("Bearer ".Length).Trim();
                UserSession session;
                lock (activeSessions)
                {
                    session = activeSessions[token];
                }

                List<Dictionary<string, object>> records = GetFilteredRecords(session.Username, session.Role, state, search, department, fromDate, toDate);
                
                JavaScriptSerializer serializer = new JavaScriptSerializer();
                string json = serializer.Serialize(records);
                
                RespondWithJson(context, json);
            }
            catch (Exception ex)
            {
                LogError("Error getting records: " + ex.Message);
                RespondWithError(context, "Failed to retrieve records: " + ex.Message, 500);
            }
        }

        private void GetRecordById(HttpContext context, string recordId)
        {
            try
            {
                // Check if the record exists
                List<Dictionary<string, object>> allRecords = LoadRecordsFromCsv();
                Dictionary<string, object> record = allRecords.FirstOrDefault(r => r["RecordID"].ToString() == recordId);

                if (record == null)
                {
                    RespondWithError(context, "Record not found", 404);
                    return;
                }

                // Get current user info from token
                string token = context.Request.Headers["Authorization"].Substring("Bearer ".Length).Trim();
                UserSession session;
                lock (activeSessions)
                {
                    session = activeSessions[token];
                }

                // Check permission to view this record
                if (!CanAccessRecord(session.Username, session.Role, record))
                {
                    RespondWithError(context, "You don't have permission to view this record", 403);
                    return;
                }

                JavaScriptSerializer serializer = new JavaScriptSerializer();
                string json = serializer.Serialize(record);
                
                RespondWithJson(context, json);
            }
            catch (Exception ex)
            {
                LogError("Error getting record: " + ex.Message);
                RespondWithError(context, "Failed to retrieve record: " + ex.Message, 500);
            }
        }

        private void CreateRecord(HttpContext context)
        {
            try
            {
                // Get current user info from token
                string token = context.Request.Headers["Authorization"].Substring("Bearer ".Length).Trim();
                UserSession session;
                lock (activeSessions)
                {
                    session = activeSessions[token];
                }

                // Check if user has permission to create records
                if (session.Role != "HR" && session.Role != "Admin")
                {
                    RespondWithError(context, "You don't have permission to create records", 403);
                    return;
                }

                // Read the request body
                string requestBody;
                using (StreamReader reader = new StreamReader(context.Request.InputStream, context.Request.ContentEncoding))
                {
                    requestBody = reader.ReadToEnd();
                }

                // Parse the JSON
                JavaScriptSerializer serializer = new JavaScriptSerializer();
                Dictionary<string, object> newRecord = serializer.Deserialize<Dictionary<string, object>>(requestBody);

                // Add required fields
                newRecord["RecordID"] = Guid.NewGuid().ToString();
                newRecord["CreatedBy"] = session.Username;
                newRecord["CreatedDate"] = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                newRecord["LastUpdated"] = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                newRecord["LastUpdatedBy"] = session.Username;
                newRecord["WorkflowState"] = 1; // PendingManagerInput

                // Save the record
                List<Dictionary<string, object>> allRecords = LoadRecordsFromCsv();
                allRecords.Add(newRecord);
                SaveRecordsToCsv(allRecords);

                // Log the creation
                LogAudit("CreateRecord", session.Username, "Created record for " + newRecord["FirstName"] + " " + newRecord["LastName"]);

                // Return the created record
                RespondWithJson(context, serializer.Serialize(newRecord));
            }
            catch (Exception ex)
            {
                LogError("Error creating record: " + ex.Message);
                RespondWithError(context, "Failed to create record: " + ex.Message, 500);
            }
        }

        private void UpdateRecord(HttpContext context, string recordId)
        {
            try
            {
                // Get current user info from token
                string token = context.Request.Headers["Authorization"].Substring("Bearer ".Length).Trim();
                UserSession session;
                lock (activeSessions)
                {
                    session = activeSessions[token];
                }

                // Check if the record exists
                List<Dictionary<string, object>> allRecords = LoadRecordsFromCsv();
                Dictionary<string, object> existingRecord = allRecords.FirstOrDefault(r => r["RecordID"].ToString() == recordId);

                if (existingRecord == null)
                {
                    RespondWithError(context, "Record not found", 404);
                    return;
                }

                // Check permission to update this record
                if (!CanModifyRecord(session.Username, session.Role, existingRecord))
                {
                    RespondWithError(context, "You don't have permission to update this record", 403);
                    return;
                }

                // Read the request body
                string requestBody;
                using (StreamReader reader = new StreamReader(context.Request.InputStream, context.Request.ContentEncoding))
                {
                    requestBody = reader.ReadToEnd();
                }

                // Parse the JSON
                JavaScriptSerializer serializer = new JavaScriptSerializer();
                Dictionary<string, object> updatedData = serializer.Deserialize<Dictionary<string, object>>(requestBody);

                // Get the action to determine workflow changes
                string action = updatedData.ContainsKey("action") ? updatedData["action"].ToString() : "";
                
                // Store original state for audit
                string originalState = existingRecord["WorkflowState"].ToString();

                // Update fields based on role and current workflow state
                UpdateRecordBasedOnRoleAndState(session.Username, session.Role, existingRecord, updatedData, action);

                // Always update these fields
                existingRecord["LastUpdated"] = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                existingRecord["LastUpdatedBy"] = session.Username;

                // Save the updated records
                SaveRecordsToCsv(allRecords);

                // Log the update
                LogAudit("UpdateRecord", session.Username, 
                    string.Format("Updated record for {0} {1}. State changed from {2} to {3}", 
                    existingRecord["FirstName"], existingRecord["LastName"], 
                    originalState, existingRecord["WorkflowState"]));

                // Return the updated record
                RespondWithJson(context, serializer.Serialize(existingRecord));
            }
            catch (Exception ex)
            {
                LogError("Error updating record: " + ex.Message);
                RespondWithError(context, "Failed to update record: " + ex.Message, 500);
            }
        }

        private void DeleteRecord(HttpContext context, string recordId)
        {
            try
            {
                // Get current user info from token
                string token = context.Request.Headers["Authorization"].Substring("Bearer ".Length).Trim();
                UserSession session;
                lock (activeSessions)
                {
                    session = activeSessions[token];
                }

                // Only admins can delete records
                if (session.Role != "Admin")
                {
                    RespondWithError(context, "Only administrators can delete records", 403);
                    return;
                }

                // Check if the record exists
                List<Dictionary<string, object>> allRecords = LoadRecordsFromCsv();
                Dictionary<string, object> recordToDelete = allRecords.FirstOrDefault(r => r["RecordID"].ToString() == recordId);

                if (recordToDelete == null)
                {
                    RespondWithError(context, "Record not found", 404);
                    return;
                }

                // Remove the record
                allRecords.Remove(recordToDelete);
                SaveRecordsToCsv(allRecords);

                // Log the deletion
                LogAudit("DeleteRecord", session.Username, 
                    string.Format("Deleted record for {0} {1}", 
                    recordToDelete["FirstName"], recordToDelete["LastName"]));

                // Return success
                var response = new { success = true, message = "Record deleted successfully" };
                JavaScriptSerializer serializer = new JavaScriptSerializer();
                RespondWithJson(context, serializer.Serialize(response));
            }
            catch (Exception ex)
            {
                LogError("Error deleting record: " + ex.Message);
                RespondWithError(context, "Failed to delete record: " + ex.Message, 500);
            }
        }

        #endregion

        #region Other API Handlers

        private void GetManagers(HttpContext context)
        {
            try
            {
                // In a real system, you would query AD for managers
                // For demo purposes, we'll return a sample list
                var managers = new List<Dictionary<string, object>>
                {
                    new Dictionary<string, object> { { "username", "Manager1" }, { "displayName", "Michael Müller" } },
                    new Dictionary<string, object> { { "username", "Manager2" }, { "displayName", "Sarah Schmidt" } },
                    new Dictionary<string, object> { { "username", "Manager3" }, { "displayName", "Thomas Weber" } }
                };

                JavaScriptSerializer serializer = new JavaScriptSerializer();
                RespondWithJson(context, serializer.Serialize(managers));
            }
            catch (Exception ex)
            {
                LogError("Error getting managers: " + ex.Message);
                RespondWithError(context, "Failed to retrieve managers: " + ex.Message, 500);
            }
        }

        private void GetAuditLog(HttpContext context)
        {
            try
            {
                // Get current user info from token
                string token = context.Request.Headers["Authorization"].Substring("Bearer ".Length).Trim();
                UserSession session;
                lock (activeSessions)
                {
                    session = activeSessions[token];
                }

                // Only admins can view the audit log
                if (session.Role != "Admin")
                {
                    RespondWithError(context, "Only administrators can view the audit log", 403);
                    return;
                }

                List<Dictionary<string, object>> auditEntries = new List<Dictionary<string, object>>();
                if (File.Exists(auditLogFile))
                {
                    using (StreamReader reader = new StreamReader(auditLogFile))
                    {
                        string header = reader.ReadLine(); // Skip CSV header
                        string[] headerFields = header.Split(',');

                        string line;
                        while ((line = reader.ReadLine()) != null)
                        {
                            string[] fields = line.Split(',');
                            Dictionary<string, object> entry = new Dictionary<string, object>();
                            
                            for (int i = 0; i < headerFields.Length && i < fields.Length; i++)
                            {
                                entry[headerFields[i]] = fields[i];
                            }
                            
                            auditEntries.Add(entry);
                        }
                    }
                }

                JavaScriptSerializer serializer = new JavaScriptSerializer();
                RespondWithJson(context, serializer.Serialize(auditEntries));
            }
            catch (Exception ex)
            {
                LogError("Error getting audit log: " + ex.Message);
                RespondWithError(context, "Failed to retrieve audit log: " + ex.Message, 500);
            }
        }

        #endregion

        #region Helper Methods

        private List<Dictionary<string, object>> LoadRecordsFromCsv()
        {
            List<Dictionary<string, object>> records = new List<Dictionary<string, object>>();
            
            if (!File.Exists(csvFile))
            {
                return records;
            }

            using (StreamReader reader = new StreamReader(csvFile))
            {
                string header = reader.ReadLine(); // CSV header
                if (string.IsNullOrEmpty(header))
                {
                    return records;
                }

                string[] headerFields = header.Split(',');

                string line;
                while ((line = reader.ReadLine()) != null)
                {
                    string[] fields = line.Split(',');
                    Dictionary<string, object> record = new Dictionary<string, object>();
                    
                    for (int i = 0; i < headerFields.Length && i < fields.Length; i++)
                    {
                        // Convert numeric and boolean values
                        if (int.TryParse(fields[i], out int intValue))
                        {
                            record[headerFields[i]] = intValue;
                        }
                        else if (bool.TryParse(fields[i], out bool boolValue))
                        {
                            record[headerFields[i]] = boolValue;
                        }
                        else
                        {
                            record[headerFields[i]] = fields[i];
                        }
                    }
                    
                    // Decrypt sensitive fields if needed
                    if (record.ContainsKey("PhoneNumber")) record["PhoneNumber"] = DecryptData(record["PhoneNumber"].ToString());
                    if (record.ContainsKey("MobileNumber")) record["MobileNumber"] = DecryptData(record["MobileNumber"].ToString());
                    if (record.ContainsKey("EmailAddress")) record["EmailAddress"] = DecryptData(record["EmailAddress"].ToString());
                    if (record.ContainsKey("PersonalNumber")) record["PersonalNumber"] = DecryptData(record["PersonalNumber"].ToString());
                    if (record.ContainsKey("ManagerNotes")) record["ManagerNotes"] = DecryptData(record["ManagerNotes"].ToString());
                    if (record.ContainsKey("ITNotes")) record["ITNotes"] = DecryptData(record["ITNotes"].ToString());
                    
                    records.Add(record);
                }
            }

            return records;
        }

        private void SaveRecordsToCsv(List<Dictionary<string, object>> records)
        {
            // Create a backup of the existing file
            if (File.Exists(csvFile))
            {
                string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                string backupFile = Path.Combine(backupFolder, $"HROnboardingData_{timestamp}.bak");
                File.Copy(csvFile, backupFile, true);
            }

            using (StreamWriter writer = new StreamWriter(csvFile, false))
            {
                // Determine all possible columns by combining keys from all records
                HashSet<string> allColumns = new HashSet<string>();
                foreach (var record in records)
                {
                    foreach (var key in record.Keys)
                    {
                        allColumns.Add(key);
                    }
                }

                // Write header
                writer.WriteLine(string.Join(",", allColumns));

                // Write records
                foreach (var record in records)
                {
                    List<string> fields = new List<string>();
                    foreach (string column in allColumns)
                    {
                        if (record.ContainsKey(column))
                        {
                            // Encrypt sensitive fields
                            string value = record[column]?.ToString() ?? "";
                            if (column == "PhoneNumber" || column == "MobileNumber" || column == "EmailAddress" ||
                                column == "PersonalNumber" || column == "ManagerNotes" || column == "ITNotes")
                            {
                                value = EncryptData(value);
                            }
                            fields.Add(value);
                        }
                        else
                        {
                            fields.Add("");
                        }
                    }
                    writer.WriteLine(string.Join(",", fields));
                }
            }
        }

        private string EncryptData(string data)
        {
            if (string.IsNullOrEmpty(data))
            {
                return data;
            }

            try
            {
                // Simple XOR encryption with a fixed key (just for demo)
                string key = "easyOnboardingSecureKey2023";
                byte[] dataBytes = Encoding.UTF8.GetBytes(data);
                byte[] keyBytes = Encoding.UTF8.GetBytes(key);

                byte[] result = new byte[dataBytes.Length];
                for (int i = 0; i < dataBytes.Length; i++)
                {
                    result[i] = (byte)(dataBytes[i] ^ keyBytes[i % keyBytes.Length]);
                }

                return Convert.ToBase64String(result);
            }
            catch
            {
                return data; // Return original on error
            }
        }

        private string DecryptData(string encryptedData)
        {
            if (string.IsNullOrEmpty(encryptedData))
            {
                return encryptedData;
            }

            try
            {
                // Check if the data is base64 encoded
                byte[] encryptedBytes;
                try
                {
                    encryptedBytes = Convert.FromBase64String(encryptedData);
                }
                catch
                {
                    return encryptedData; // Not encrypted, return as is
                }

                // Simple XOR decryption with a fixed key (just for demo)
                string key = "easyOnboardingSecureKey2023";
                byte[] keyBytes = Encoding.UTF8.GetBytes(key);

                byte[] result = new byte[encryptedBytes.Length];
                for (int i = 0; i < encryptedBytes.Length; i++)
                {
                    result[i] = (byte)(encryptedBytes[i] ^ keyBytes[i % keyBytes.Length]);
                }

                return Encoding.UTF8.GetString(result);
            }
            catch
            {
                return encryptedData; // Return encrypted on error
            }
        }

        private void LogAudit(string action, string user, string details)
        {
            try
            {
                bool fileExists = File.Exists(auditLogFile);
                
                using (StreamWriter writer = new StreamWriter(auditLogFile, true))
                {
                    if (!fileExists)
                    {
                        writer.WriteLine("Timestamp,User,Action,RecordID,Details");
                    }
                    
                    writer.WriteLine($"{DateTime.Now:yyyy-MM-dd HH:mm:ss},{user},{action},,{details.Replace(",", ";")}");
                }
            }
            catch (Exception ex)
            {
                LogError("Error writing to audit log: " + ex.Message);
            }
        }

        private void LogError(string message)
        {
            try
            {
                string logFolder = Path.Combine(dataFolder, "Logs");
                if (!Directory.Exists(logFolder))
                {
                    Directory.CreateDirectory(logFolder);
                }

                string logFile = Path.Combine(logFolder, $"Error_{DateTime.Now:yyyyMMdd}.log");
                using (StreamWriter writer = new StreamWriter(logFile, true))
                {
                    writer.WriteLine($"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - ERROR - {message}");
                }
            }
            catch
            {
                // Can't do much if logging itself fails
            }
        }

        private string GenerateToken()
        {
            using (RandomNumberGenerator rng = RandomNumberGenerator.Create())
            {
                byte[] tokenData = new byte[32]; // 256 bits
                rng.GetBytes(tokenData);
                return Convert.ToBase64String(tokenData);
            }
        }

        private bool IsMemberOfGroup(UserPrincipal user, string groupName)
        {
            try
            {
                PrincipalContext context = user.Context;
                GroupPrincipal group = GroupPrincipal.FindByIdentity(context, groupName);
                
                if (group == null)
                {
                    return false;
                }
                
                return user.IsMemberOf(group);
            }
            catch
            {
                return false;
            }
        }

        private bool IsTestUser(string username, string password)
        {
            // For testing/demo purposes only
            Dictionary<string, string> testUsers = new Dictionary<string, string>
            {
                { "hradmin", "hr123" },
                { "itadmin", "it123" },
                { "manager1", "mgr123" },
                { "sysadmin", "admin123" }
            };

            return testUsers.ContainsKey(username.ToLower()) && testUsers[username.ToLower()] == password;
        }

        private string GetTestUserRole(string username)
        {
            // Map test users to roles
            switch (username.ToLower())
            {
                case "hradmin": return "HR";
                case "itadmin": return "IT";
                case "manager1": return "Manager";
                case "sysadmin": return "Admin";
                default: return "User";
            }
        }

        private bool CanAccessRecord(string username, string role, Dictionary<string, object> record)
        {
            // Admin can access all records
            if (role == "Admin" || role == "HR")
            {
                return true;
            }

            // IT can access records that are ready for IT or completed
            if (role == "IT")
            {
                int workflowState = Convert.ToInt32(record["WorkflowState"]);
                return workflowState == 3 || workflowState == 4; // ReadyForIT or Completed
            }

            // Manager can only access records assigned to them
            if (role == "Manager")
            {
                string assignedManager = record["AssignedManager"]?.ToString() ?? "";
                return assignedManager.Equals(username, StringComparison.OrdinalIgnoreCase);
            }

            return false;
        }

        private bool CanModifyRecord(string username, string role, Dictionary<string, object> record)
        {
            // Admin can modify all records
            if (role == "Admin")
            {
                return true;
            }

            int workflowState = Convert.ToInt32(record["WorkflowState"]);

            // HR can modify new records and those pending HR verification
            if (role == "HR")
            {
                return workflowState == 0 || workflowState == 2; // New or PendingHRVerification
            }

            // IT can modify records that are ready for IT
            if (role == "IT")
            {
                return workflowState == 3; // ReadyForIT
            }

            // Manager can modify records assigned to them and in pending manager input state
            if (role == "Manager")
            {
                string assignedManager = record["AssignedManager"]?.ToString() ?? "";
                return assignedManager.Equals(username, StringComparison.OrdinalIgnoreCase) && workflowState == 1; // PendingManagerInput
            }

            return false;
        }

        private void UpdateRecordBasedOnRoleAndState(string username, string role, Dictionary<string, object> existingRecord, Dictionary<string, object> updatedData, string action)
        {
            int currentState = Convert.ToInt32(existingRecord["WorkflowState"]);

            if (role == "Admin")
            {
                // Admin can update any field including workflow state
                foreach (var key in updatedData.Keys.Where(k => k != "action"))
                {
                    existingRecord[key] = updatedData[key];
                }
                return;
            }

            if (role == "HR")
            {
                if (currentState == 0) // New
                {
                    // HR creating/editing a new record
                    if (action == "HRSubmit")
                    {
                        UpdateFieldsIfPresent(existingRecord, updatedData, new[]
                        {
                            "FirstName", "LastName", "Description", "OfficeRoom", "PhoneNumber",
                            "MobileNumber", "EmailAddress", "External", "ExternalCompany",
                            "StartWorkDate", "AssignedManager", "HRNotes", "WorkflowState"
                        });
                    }
                }
                else if (currentState == 2) // PendingHRVerification
                {
                    // HR verifying a record
                    if (action == "HRVerify")
                    {
                        UpdateFieldsIfPresent(existingRecord, updatedData, new[]
                        {
                            "HRVerified", "VerificationNotes", "WorkflowState"
                        });
                    }
                }
            }
            else if (role == "Manager" && currentState == 1) // PendingManagerInput
            {
                // Manager providing input
                if (action == "ManagerSubmit")
                {
                    UpdateFieldsIfPresent(existingRecord, updatedData, new[]
                    {
                        "Position", "DepartmentField", "PersonalNumber", "Ablaufdatum",
                        "TL", "AL", "ManagerNotes", "SoftwareSage", "SoftwareGenesis",
                        "ZugangLizenzmanager", "ZugangMS365", "Zugriffe", "WorkflowState"
                    });
                }
            }
            else if (role == "IT" && currentState == 3) // ReadyForIT
            {
                // IT completing tasks
                if (action == "ITComplete")
                {
                    UpdateFieldsIfPresent(existingRecord, updatedData, new[]
                    {
                        "AccountCreated", "EquipmentReady", "ITNotes", "WorkflowState"
                    });
                }
            }
        }

        private void UpdateFieldsIfPresent(Dictionary<string, object> target, Dictionary<string, object> source, string[] allowedFields)
        {
            foreach (var field in allowedFields)
            {
                if (source.ContainsKey(field))
                {
                    target[field] = source[field];
                }
            }
        }

        private List<Dictionary<string, object>> GetFilteredRecords(string username, string role, string state, string search, string department, DateTime? fromDate, DateTime? toDate)
        {
            List<Dictionary<string, object>> allRecords = LoadRecordsFromCsv();
            List<Dictionary<string, object>> filteredRecords = new List<Dictionary<string, object>>();

            // Filter based on user role
            if (role == "Admin" || role == "HR")
            {
                // Admin and HR see all records
                filteredRecords = allRecords;
            }
            else if (role == "IT")
            {
                // IT sees records that are ReadyForIT or Completed
                filteredRecords = allRecords.Where(r => 
                    Convert.ToInt32(r["WorkflowState"]) == 3 || // ReadyForIT
                    Convert.ToInt32(r["WorkflowState"]) == 4    // Completed
                ).ToList();
            }
            else if (role == "Manager")
            {
                // Manager sees only records assigned to them
                filteredRecords = allRecords.Where(r => 
                    r["AssignedManager"]?.ToString().Equals(username, StringComparison.OrdinalIgnoreCase) ?? false
                ).ToList();
            }

            // Apply additional filters
            // Filter by workflow state if specified
            if (!string.IsNullOrEmpty(state))
            {
                int stateValue = -1;
                switch (state.ToLower())
                {
                    case "new": stateValue = 0; break;
                    case "pendingmanagerinput": stateValue = 1; break;
                    case "pendinghrverification": stateValue = 2; break;
                    case "readyforit": stateValue = 3; break;
                    case "completed": stateValue = 4; break;
                }

                if (stateValue >= 0)
                {
                    filteredRecords = filteredRecords.Where(r => Convert.ToInt32(r["WorkflowState"]) == stateValue).ToList();
                }
            }

            // Filter by search text
            if (!string.IsNullOrEmpty(search))
            {
                search = search.ToLower();
                filteredRecords = filteredRecords.Where(r =>
                    (r["FirstName"]?.ToString().ToLower().Contains(search) ?? false) ||
                    (r["LastName"]?.ToString().ToLower().Contains(search) ?? false) ||
                    (r["Description"]?.ToString().ToLower().Contains(search) ?? false) ||
                    (r["Position"]?.ToString().ToLower().Contains(search) ?? false)
                ).ToList();
            }

            // Filter by department
            if (!string.IsNullOrEmpty(department))
            {
                filteredRecords = filteredRecords.Where(r =>
                    (r["DepartmentField"]?.ToString().Equals(department, StringComparison.OrdinalIgnoreCase) ?? false)
                ).ToList();
            }

            // Filter by date range
            if (fromDate.HasValue || toDate.HasValue)
            {
                filteredRecords = filteredRecords.Where(r =>
                {
                    if (r.ContainsKey("CreatedDate") && DateTime.TryParse(r["CreatedDate"]?.ToString(), out DateTime recordDate))
                    {
                        return (!fromDate.HasValue || recordDate >= fromDate.Value) &&
                               (!toDate.HasValue || recordDate <= toDate.Value);
                    }
                    return true; // Include records with invalid dates
                }).ToList();
            }

            return filteredRecords;
        }

        private void RespondWithJson(HttpContext context, string json)
        {
            context.Response.ContentType = "application/json";
            context.Response.Write(json);
        }

        private void RespondWithError(HttpContext context, string message, int statusCode)
        {
            context.Response.ContentType = "application/json";
            context.Response.StatusCode = statusCode;
            
            var error = new { error = true, message = message };
            JavaScriptSerializer serializer = new JavaScriptSerializer();
            context.Response.Write(serializer.Serialize(error));
        }

        private void HandleCorsOptions(HttpContext context)
        {
            // Set CORS headers for preflight OPTIONS requests
            context.Response.StatusCode = 200;
            context.Response.End();
        }

        #endregion

        public bool IsReusable
        {
            get { return false; }
        }
    }

    public class UserSession
    {
        public string Username { get; set; }
        public string Role { get; set; }
        public DateTime Expiration { get; set; }
    }
}
