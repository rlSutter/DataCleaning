# Data Cleaning Service - Customization and Implementation Guide

This comprehensive guide provides step-by-step instructions for implementing the Data Cleaning Service in your environment, including configuration, deployment, and customization options.

## Table of Contents

1. [Prerequisites and System Requirements](#prerequisites-and-system-requirements)
2. [Environment Setup](#environment-setup)
3. [Database Configuration](#database-configuration)
4. [MelissaData Integration](#melissadata-integration)
5. [Third-Party Service Configuration](#third-party-service-configuration)
6. [Web Service Deployment](#web-service-deployment)
7. [Configuration Customization](#configuration-customization)
8. [Security Configuration](#security-configuration)
9. [Performance Tuning](#performance-tuning)
10. [Testing and Validation](#testing-and-validation)
11. [Monitoring and Maintenance](#monitoring-and-maintenance)
12. [Troubleshooting](#troubleshooting)

## Prerequisites and System Requirements

### Hardware Requirements
- **CPU**: Minimum 4 cores, recommended 8+ cores
- **RAM**: Minimum 8GB, recommended 16GB+
- **Storage**: 50GB free space for MelissaData files and logs
- **Network**: Stable internet connection for geocoding services

### Software Requirements
- **Operating System**: Windows Server 2012 R2 or higher
- **Web Server**: IIS 8.0 or higher
- **Framework**: .NET Framework 4.0 or higher
- **Database**: SQL Server 2012 or higher
- **Visual Studio**: 2010 or higher (for development)

### Required Licenses
- **MelissaData DQT License**: For data quality and standardization
- **MelissaData MatchUP License**: For fuzzy matching and deduplication
- **Bing Maps API Key**: For geocoding services
- **Geocodio API Key**: For enhanced geocoding accuracy
- **Certegrity Cloud Services**: For authentication (if applicable)

## Environment Setup

### 1. Windows Server Configuration

#### Install Required Features
```powershell
# Install IIS with required features
Enable-WindowsOptionalFeature -Online -FeatureName IIS-WebServerRole
Enable-WindowsOptionalFeature -Online -FeatureName IIS-WebServer
Enable-WindowsOptionalFeature -Online -FeatureName IIS-CommonHttpFeatures
Enable-WindowsOptionalFeature -Online -FeatureName IIS-HttpErrors
Enable-WindowsOptionalFeature -Online -FeatureName IIS-HttpLogging
Enable-WindowsOptionalFeature -Online -FeatureName IIS-RequestFiltering
Enable-WindowsOptionalFeature -Online -FeatureName IIS-StaticContent
Enable-WindowsOptionalFeature -Online -FeatureName IIS-DefaultDocument
Enable-WindowsOptionalFeature -Online -FeatureName IIS-DirectoryBrowsing
Enable-WindowsOptionalFeature -Online -FeatureName IIS-ASPNET45
```

#### Configure Application Pool
1. Open IIS Manager
2. Create new Application Pool:
   - **Name**: DataCleaningService
   - **.NET CLR Version**: v4.0
   - **Managed Pipeline Mode**: Integrated
   - **Identity**: ApplicationPoolIdentity (or custom service account)

#### Set Application Pool Settings
```xml
<!-- In applicationHost.config -->
<add name="DataCleaningService">
    <processModel identityType="ApplicationPoolIdentity" />
    <recycling>
        <periodicRestart time="00:00:00" />
        <periodicRestart requests="0" />
    </recycling>
    <cpu limit="0" />
    <memory limit="0" />
</add>
```

### 2. .NET Framework Configuration

#### Install .NET Framework 4.0+
```powershell
# Download and install .NET Framework 4.8
# Ensure ASP.NET 4.0 is registered
C:\Windows\Microsoft.NET\Framework64\v4.0.30319\aspnet_regiis.exe -i
```

#### Configure Machine.config
```xml
<!-- Add to machine.config if needed -->
<system.web>
    <httpRuntime maxUrlLength="8192" maxQueryStringLength="8192" />
    <compilation debug="false" targetFramework="4.0" />
</system.web>
```

## Database Configuration

### 1. SQL Server Setup

#### Create Database
```sql
-- Create the main database
CREATE DATABASE [DataCleaningDB]
ON (NAME = 'DataCleaningDB', FILENAME = 'C:\Data\DataCleaningDB.mdf')
LOG ON (NAME = 'DataCleaningDB_Log', FILENAME = 'C:\Data\DataCleaningDB_Log.ldf');
GO

-- Set database options
ALTER DATABASE [DataCleaningDB] SET RECOVERY_SIMPLE;
ALTER DATABASE [DataCleaningDB] SET AUTO_SHRINK OFF;
ALTER DATABASE [DataCleaningDB] SET AUTO_CREATE_STATISTICS ON;
ALTER DATABASE [DataCleaningDB] SET AUTO_UPDATE_STATISTICS ON;
GO
```

#### Create Service Account
```sql
-- Create database user for the service
CREATE LOGIN [DataCleaningService] WITH PASSWORD = 'YourSecurePassword123!';
GO

USE [DataCleaningDB];
CREATE USER [DataCleaningService] FOR LOGIN [DataCleaningService];
GO

-- Grant necessary permissions
ALTER ROLE [db_datareader] ADD MEMBER [DataCleaningService];
ALTER ROLE [db_datawriter] ADD MEMBER [DataCleaningService];
ALTER ROLE [db_ddladmin] ADD MEMBER [DataCleaningService];
GO
```

#### Run Database Schema Script
```bash
# Execute the provided database_schema.sql
sqlcmd -S YourServer -d DataCleaningDB -i database_schema.sql
```

### 2. Connection String Configuration

#### Update web.config Connection Strings
```xml
<connectionStrings>
    <!-- Production Database -->
    <add name="hcidb" 
         connectionString="server=YourServer\Instance;uid=DataCleaningService;pwd=YourSecurePassword123!;database=DataCleaningDB;Min Pool Size=5;Max Pool Size=100;Connect Timeout=30;Command Timeout=60;" 
         providerName="System.Data.SqlClient"/>
    
    <!-- Read-Only Database (for reporting) -->
    <add name="hcidb_ro" 
         connectionString="server=YourServer\Instance;uid=DataCleaningService;pwd=YourSecurePassword123!;ApplicationIntent=ReadOnly;database=DataCleaningDB;Min Pool Size=3;Max Pool Size=50;Connect Timeout=30;" 
         providerName="System.Data.SqlClient"/>
    
    <!-- Reports Database -->
    <add name="reports" 
         connectionString="server=YourServer\Instance;uid=DataCleaningService;pwd=YourSecurePassword123!;database=DataCleaningDB;Min Pool Size=3;Max Pool Size=50;" 
         providerName="System.Data.SqlClient"/>
</connectionStrings>
```

## MelissaData Integration

### 1. MelissaData Installation

#### Download and Install
1. Download MelissaData DQT from [MelissaData website](https://www.melissadata.com/)
2. Install with administrative privileges
3. Download and install MatchUP separately
4. Ensure data files are accessible

#### Default Installation Paths
```
DQT Data Path: C:\Program Files\Melissa DATA\DQT\Data
MatchUP Data Path: C:\ProgramData\Melissa DATA\MatchUP
```

### 2. License Configuration

#### Obtain License Keys
1. Contact MelissaData for license keys
2. Ensure licenses are valid and not expired
3. Test license activation

#### Configure web.config
```xml
<appSettings>
    <!-- MelissaData DQT License -->
    <add key="MD_Key" value="YOUR_ACTUAL_MELISSADATA_DQT_LICENSE_KEY"/>
    
    <!-- MelissaData MatchUP License -->
    <add key="MD_MU_Key" value="YOUR_ACTUAL_MELISSADATA_MATCHUP_LICENSE_KEY"/>
    
    <!-- Data File Paths -->
    <add key="MD_DataPath" value="C:\Program Files\Melissa DATA\DQT\Data"/>
    <add key="MD_MU_DataPath" value="C:\ProgramData\Melissa DATA\MatchUP"/>
    
    <!-- License Verification Settings -->
    <add key="MDVerifiedDaysAgo" value="1"/>
</appSettings>
```

### 3. Data File Verification

#### Test Data File Access
```vb
' Add this test method to verify MelissaData setup
Public Function TestMelissaDataSetup() As String
    Try
        Dim dPath As String = ConfigurationManager.AppSettings("MD_DataPath")
        Dim dLICENSE As String = ConfigurationManager.AppSettings("MD_Key")
        
        Dim nameObj As New MelissaData.NameObject()
        Dim result As Integer = nameObj.InitializeDataFiles(dLICENSE, dPath)
        
        If result = 0 Then
            Return "MelissaData DQT initialized successfully. License expires: " & nameObj.GetLicenseExpirationDate()
        Else
            Return "MelissaData DQT initialization failed. Error code: " & result
        End If
    Catch ex As Exception
        Return "MelissaData test failed: " & ex.Message
    End Try
End Function
```

## Third-Party Service Configuration

### 1. Bing Maps API Setup

#### Create Bing Maps Account
1. Go to [Bing Maps Portal](https://www.bingmapsportal.com/)
2. Create account and sign in
3. Create new application
4. Generate API key

#### Configure API Key
```xml
<appSettings>
    <add key="bing_key" value="YOUR_ACTUAL_BING_MAPS_API_KEY"/>
</appSettings>
```

#### Test Bing Maps Integration
```vb
Public Function TestBingMapsAPI() As String
    Try
        Dim apiKey As String = ConfigurationManager.AppSettings("bing_key")
        Dim testUrl As String = "http://dev.virtualearth.net/REST/v1/Locations?query=1 Microsoft Way, Redmond, WA&key=" & apiKey
        
        Dim client As New WebClient()
        Dim response As String = client.DownloadString(testUrl)
        
        Return "Bing Maps API test successful"
    Catch ex As Exception
        Return "Bing Maps API test failed: " & ex.Message
    End Try
End Function
```

### 2. Geocodio API Setup

#### Create Geocodio Account
1. Go to [Geocodio website](https://www.geocod.io/)
2. Sign up for account
3. Generate API key from dashboard

#### Configure API Key
```xml
<appSettings>
    <add key="geocodio_key" value="YOUR_ACTUAL_GEOCODIO_API_KEY"/>
    <add key="GeocodeUrl" value="https://api.geocod.io/v1/geocode?"/>
</appSettings>
```

### 3. Certegrity Cloud Services (Optional)

#### Configure Authentication Service
```xml
<appSettings>
    <add key="com.certegrity.cloudsvc.service" value="https://cloudsvc.certegrity.com/basic/service.asmx"/>
    <add key="com.certegrity.cloudsvc.basic.service" value="https://cloudsvc.certegrity.com/basic/service.asmx"/>
</appSettings>
```

## Web Service Deployment

### 1. File Deployment

#### Copy Application Files
```powershell
# Create application directory
New-Item -ItemType Directory -Path "C:\inetpub\wwwroot\DataCleaningService"

# Copy application files
Copy-Item -Path ".\*" -Destination "C:\inetpub\wwwroot\DataCleaningService" -Recurse
```

#### Set File Permissions
```powershell
# Grant IIS_IUSRS read permissions
icacls "C:\inetpub\wwwroot\DataCleaningService" /grant "IIS_IUSRS:(OI)(CI)R"

# Grant application pool identity permissions
icacls "C:\inetpub\wwwroot\DataCleaningService" /grant "IIS AppPool\DataCleaningService:(OI)(CI)F"
```

### 2. IIS Configuration

#### Create Application
1. Open IIS Manager
2. Right-click on Default Web Site
3. Add Application:
   - **Alias**: DataCleaningService
   - **Application Pool**: DataCleaningService
   - **Physical Path**: C:\inetpub\wwwroot\DataCleaningService

#### Configure Application Settings
```xml
<!-- In web.config -->
<system.web>
    <httpRuntime maxUrlLength="8192" maxQueryStringLength="8192" requestValidationMode="2.0"/>
    <compilation debug="false" strict="false" explicit="true" targetFramework="4.0">
        <assemblies>
            <add assembly="mscorlib, Version=4.0.0.0, Culture=neutral, PublicKeyToken=B77A5C561934E089"/>
        </assemblies>
    </compilation>
    <pages controlRenderingCompatibilityVersion="4.0" clientIDMode="AutoID" validateRequest="false">
        <namespaces>
            <clear/>
            <add namespace="System"/>
            <add namespace="System.Collections"/>
            <add namespace="System.Collections.Specialized"/>
            <add namespace="System.Configuration"/>
            <add namespace="System.Text"/>
            <add namespace="System.Text.RegularExpressions"/>
            <add namespace="System.Web"/>
            <add namespace="System.Web.Caching"/>
            <add namespace="System.Web.SessionState"/>
            <add namespace="System.Web.Security"/>
            <add namespace="System.Web.Profile"/>
            <add namespace="System.Web.UI"/>
            <add namespace="System.Web.UI.WebControls"/>
            <add namespace="System.Web.UI.WebControls.WebParts"/>
            <add namespace="System.Web.UI.HtmlControls"/>
        </namespaces>
    </pages>
    <webServices>
        <protocols>
            <add name="HttpGet"/>
            <add name="HttpPost"/>
        </protocols>
    </webServices>
</system.web>
```

### 3. SSL Configuration (Recommended)

#### Install SSL Certificate
```powershell
# Import SSL certificate
Import-Certificate -FilePath "C:\Certificates\DataCleaningService.pfx" -CertStoreLocation Cert:\LocalMachine\My
```

#### Configure HTTPS Binding
1. Open IIS Manager
2. Select DataCleaningService application
3. Bindings → Add:
   - **Type**: https
   - **Port**: 443
   - **SSL Certificate**: Your certificate

## Configuration Customization

### 1. Debug and Logging Configuration

#### Enable/Disable Debug Modes
```xml
<appSettings>
    <!-- Debug Settings - Set to "N" for production -->
    <add key="CleanOrganization_debug" value="N"/>
    <add key="CleanContact_debug" value="N"/>
    <add key="CleanAddress_debug" value="N"/>
    <add key="CleanAddress_detailed_debug" value="N"/>
    <add key="StandardizeOrganization_debug" value="N"/>
    <add key="StandardizeContact_debug" value="N"/>
    <add key="StandardizeAddress_debug" value="N"/>
</appSettings>
```

#### Configure Logging
```xml
<log4net>
    <!-- Remote Syslog Configuration -->
    <appender name="RemoteSyslogAppender" type="log4net.Appender.RemoteSyslogAppender">
        <identity value="DataCleaningService"/>
        <layout type="log4net.Layout.PatternLayout" value="%message"/>
        <remoteAddress value="YOUR_SYSLOG_SERVER_IP"/>
        <filter type="log4net.Filter.LevelRangeFilter">
            <levelMin value="INFO"/>
            <levelMax value="FATAL"/>
        </filter>
    </appender>
    
    <!-- File Logging Configuration -->
    <appender name="ServiceLogFileAppender" type="log4net.Appender.RollingFileAppender">
        <file type="log4net.Util.PatternString" value="C:\Logs\DataCleaningService\Service.log"/>
        <lockingModel type="log4net.Appender.RollingFileAppender+MinimalLock"/>
        <appendToFile value="true"/>
        <rollingStyle value="Date"/>
        <datePattern value="yyyyMMdd"/>
        <maxSizeRollBackups value="30"/>
        <maximumFileSize value="10MB"/>
        <staticLogFileName value="false"/>
        <immediateFlush value="true"/>
        <layout type="log4net.Layout.PatternLayout">
            <conversionPattern value="%date [%thread] %-5level %logger - %message%newline"/>
        </layout>
    </appender>
    
    <!-- Logger Configuration -->
    <logger name="EventLog">
        <level value="INFO"/>
        <appender-ref ref="RemoteSyslogAppender"/>
    </logger>
    
    <root>
        <level value="INFO"/>
        <appender-ref ref="ServiceLogFileAppender"/>
    </root>
</log4net>
```

### 2. Performance Configuration

#### Connection Pool Settings
```xml
<connectionStrings>
    <add name="hcidb" 
         connectionString="server=YourServer;uid=YourUser;pwd=YourPassword;database=YourDB;Min Pool Size=10;Max Pool Size=200;Connect Timeout=30;Command Timeout=120;Connection Lifetime=300;" 
         providerName="System.Data.SqlClient"/>
</connectionStrings>
```

#### Application Pool Optimization
```xml
<!-- In applicationHost.config -->
<add name="DataCleaningService">
    <processModel identityType="ApplicationPoolIdentity" idleTimeout="00:20:00" />
    <recycling>
        <periodicRestart time="00:00:00" />
        <periodicRestart requests="0" />
        <periodicRestart memory="0" />
    </recycling>
    <cpu limit="0" />
    <memory limit="0" />
    <queueLength value="1000" />
</add>
```

### 3. Custom Business Rules

#### Implement Custom Matching Logic
```vb
' Add to Service.vb for custom business rules
Private Function ApplyCustomBusinessRules(ByVal contactData As Object) As Object
    ' Example: Custom name standardization
    If Not String.IsNullOrEmpty(contactData.FirstName) Then
        contactData.FirstName = contactData.FirstName.Trim().ToTitleCase()
    End If
    
    ' Example: Custom validation rules
    If String.IsNullOrEmpty(contactData.LastName) Then
        Throw New ArgumentException("Last name is required")
    End If
    
    Return contactData
End Function
```

## Security Configuration

### 1. Authentication and Authorization

#### Implement API Key Authentication
```vb
' Add to Service.vb
Private Function ValidateApiKey(ByVal apiKey As String) As Boolean
    Dim validKeys As String() = {"YOUR_API_KEY_1", "YOUR_API_KEY_2"}
    Return validKeys.Contains(apiKey)
End Function

<WebMethod(Description:="Standardizes contact with API key authentication")> _
Public Function StandardizeContactSecure(ByVal sXML As String, ByVal apiKey As String) As XmlDocument
    If Not ValidateApiKey(apiKey) Then
        Throw New UnauthorizedAccessException("Invalid API key")
    End If
    
    Return StandardizeContact(sXML)
End Function
```

#### Configure IP Restrictions
```xml
<!-- In web.config -->
<system.webServer>
    <security>
        <ipSecurity allowUnlisted="false">
            <add ipAddress="192.168.1.0" subnetMask="255.255.255.0" allowed="true"/>
            <add ipAddress="10.0.0.0" subnetMask="255.0.0.0" allowed="true"/>
        </ipSecurity>
    </security>
</system.webServer>
```

### 2. Data Encryption

#### Encrypt Sensitive Configuration
```xml
<!-- Use aspnet_regiis to encrypt sections -->
<!-- aspnet_regiis -pef "connectionStrings" "C:\inetpub\wwwroot\DataCleaningService" -->
<connectionStrings configProtectionProvider="RsaProtectedConfigurationProvider">
    <EncryptedData Type="http://www.w3.org/2001/04/xmlenc#Element"
                   xmlns="http://www.w3.org/2001/04/xmlenc#">
        <!-- Encrypted connection string data -->
    </EncryptedData>
</connectionStrings>
```

### 3. Input Validation

#### Implement Input Sanitization
```vb
Private Function SanitizeInput(ByVal input As String) As String
    If String.IsNullOrEmpty(input) Then Return String.Empty
    
    ' Remove potentially dangerous characters
    input = input.Replace("'", "''")
    input = input.Replace("--", "")
    input = input.Replace("/*", "")
    input = input.Replace("*/", "")
    
    ' Limit length
    If input.Length > 1000 Then
        input = input.Substring(0, 1000)
    End If
    
    Return input.Trim()
End Function
```

## Performance Tuning

### 1. Database Optimization

#### Create Additional Indexes
```sql
-- Performance indexes for common queries
CREATE NONCLUSTERED INDEX [IX_S_CONTACT_FULL_NAME] 
ON [dbo].[S_CONTACT] ([FST_NAME], [LAST_NAME], [MID_NAME])
INCLUDE ([ROW_ID], [X_MATCH_CD]);

CREATE NONCLUSTERED INDEX [IX_S_ORG_EXT_NAME_LOC] 
ON [dbo].[S_ORG_EXT] ([NAME], [LOC])
INCLUDE ([ROW_ID], [DEDUP_TOKEN]);

CREATE NONCLUSTERED INDEX [IX_S_ADDR_PER_FULL_ADDRESS] 
ON [dbo].[S_ADDR_PER] ([ADDR], [CITY], [STATE], [ZIPCODE])
INCLUDE ([ROW_ID], [X_MATCH_CD], [X_LAT], [X_LONG]);
```

#### Configure Database Settings
```sql
-- Optimize database for the workload
ALTER DATABASE [DataCleaningDB] SET AUTO_CREATE_STATISTICS ON;
ALTER DATABASE [DataCleaningDB] SET AUTO_UPDATE_STATISTICS ON;
ALTER DATABASE [DataCleaningDB] SET AUTO_UPDATE_STATISTICS_ASYNC ON;

-- Set appropriate compatibility level
ALTER DATABASE [DataCleaningDB] SET COMPATIBILITY_LEVEL = 130;
```

### 2. Application Performance

#### Implement Caching
```vb
' Add caching for MelissaData objects
Private Shared ReadOnly MelissaDataCache As New Dictionary(Of String, Object)

Private Function GetCachedMelissaDataObject(ByVal key As String) As Object
    If MelissaDataCache.ContainsKey(key) Then
        Return MelissaDataCache(key)
    End If
    Return Nothing
End Function

Private Sub CacheMelissaDataObject(ByVal key As String, ByVal obj As Object)
    MelissaDataCache(key) = obj
End Sub
```

#### Optimize Memory Usage
```xml
<!-- In web.config -->
<system.web>
    <httpRuntime maxUrlLength="8192" 
                 maxQueryStringLength="8192" 
                 maxRequestLength="4096"
                 executionTimeout="300"
                 minFreeThreads="8"
                 minLocalRequestFreeThreads="4"/>
</system.web>
```

## Testing and Validation

### 1. Unit Testing

#### Create Test Methods
```vb
<WebMethod(Description:="Test method for validation")> _
Public Function TestService() As String
    Dim results As New StringBuilder()
    
    ' Test database connection
    Try
        Dim con As SqlConnection
        Dim cmd As SqlCommand
        Dim result As String = OpenDBConnection(ConfigurationManager.ConnectionStrings("hcidb").ConnectionString, con, cmd)
        If result = "" Then
            results.AppendLine("✓ Database connection: SUCCESS")
            con.Close()
        Else
            results.AppendLine("✗ Database connection: FAILED - " & result)
        End If
    Catch ex As Exception
        results.AppendLine("✗ Database connection: ERROR - " & ex.Message)
    End Try
    
    ' Test MelissaData
    Try
        Dim testResult As String = TestMelissaDataSetup()
        results.AppendLine("✓ MelissaData: " & testResult)
    Catch ex As Exception
        results.AppendLine("✗ MelissaData: ERROR - " & ex.Message)
    End Try
    
    ' Test geocoding services
    Try
        Dim geoResult As String = TestBingMapsAPI()
        results.AppendLine("✓ Bing Maps API: " & geoResult)
    Catch ex As Exception
        results.AppendLine("✗ Bing Maps API: ERROR - " & ex.Message)
    End Try
    
    Return results.ToString()
End Function
```

### 2. Integration Testing

#### Test Contact Standardization
```xml
<!-- Test XML for contact standardization -->
<Contacts>
    <Contact>
        <Debug>Y</Debug>
        <Database>T</Database>
        <FirstName>John</FirstName>
        <LastName>Smith</LastName>
        <Gender>M</Gender>
        <FullName>John Smith</FullName>
    </Contact>
</Contacts>
```

#### Test Organization Standardization
```xml
<!-- Test XML for organization standardization -->
<Organizations>
    <Organization>
        <Debug>Y</Debug>
        <Database>T</Database>
        <Name>Acme Corporation</Name>
        <Location>New York, NY</Location>
    </Organization>
</Organizations>
```

### 3. Load Testing

#### Create Load Test Script
```powershell
# PowerShell script for load testing
$serviceUrl = "http://localhost/DataCleaningService/Service.asmx"
$testXml = @"
<Contacts>
    <Contact>
        <Debug>N</Debug>
        <Database>T</Database>
        <FirstName>Test</FirstName>
        <LastName>User</LastName>
        <Gender>M</Gender>
        <FullName>Test User</FullName>
    </Contact>
</Contacts>
"@

# Simulate 100 concurrent requests
1..100 | ForEach-Object -Parallel {
    $response = Invoke-WebRequest -Uri $serviceUrl -Method POST -Body $testXml -ContentType "text/xml"
    Write-Host "Request $_ completed with status: $($response.StatusCode)"
} -ThrottleLimit 10
```

## Monitoring and Maintenance

### 1. Health Monitoring

#### Create Health Check Endpoint
```vb
<WebMethod(Description:="Health check endpoint")> _
Public Function HealthCheck() As String
    Dim health As New Dictionary(Of String, String)
    
    ' Check database connectivity
    Try
        Dim con As SqlConnection
        Dim cmd As SqlCommand
        Dim result As String = OpenDBConnection(ConfigurationManager.ConnectionStrings("hcidb").ConnectionString, con, cmd)
        health("Database") = If(result = "", "Healthy", "Unhealthy: " & result)
        If con IsNot Nothing Then con.Close()
    Catch ex As Exception
        health("Database") = "Unhealthy: " & ex.Message
    End Try
    
    ' Check MelissaData
    Try
        Dim dPath As String = ConfigurationManager.AppSettings("MD_DataPath")
        Dim dLICENSE As String = ConfigurationManager.AppSettings("MD_Key")
        Dim nameObj As New MelissaData.NameObject()
        Dim initResult As Integer = nameObj.InitializeDataFiles(dLICENSE, dPath)
        health("MelissaData") = If(initResult = 0, "Healthy", "Unhealthy: Error " & initResult)
    Catch ex As Exception
        health("MelissaData") = "Unhealthy: " & ex.Message
    End Try
    
    ' Check external services
    Try
        Dim apiKey As String = ConfigurationManager.AppSettings("bing_key")
        If Not String.IsNullOrEmpty(apiKey) Then
            health("BingMaps") = "Configured"
        Else
            health("BingMaps") = "Not Configured"
        End If
    Catch ex As Exception
        health("BingMaps") = "Error: " & ex.Message
    End Try
    
    Return String.Join("; ", health.Select(Function(kvp) kvp.Key & "=" & kvp.Value))
End Function
```

### 2. Log Monitoring

#### Set Up Log Monitoring
```powershell
# PowerShell script for log monitoring
$logPath = "C:\Logs\DataCleaningService"
$errorPattern = "ERROR|FATAL|Exception"

Get-ChildItem -Path $logPath -Filter "*.log" | ForEach-Object {
    $content = Get-Content $_.FullName -Tail 100
    $errors = $content | Select-String -Pattern $errorPattern
    if ($errors) {
        Write-Host "Errors found in $($_.Name):"
        $errors | ForEach-Object { Write-Host "  $($_.Line)" }
    }
}
```

### 3. Performance Monitoring

#### Create Performance Metrics
```vb
<WebMethod(Description:="Performance metrics endpoint")> _
Public Function GetPerformanceMetrics() As String
    Dim metrics As New Dictionary(Of String, String)
    
    ' Get process information
    Dim process As Process = Process.GetCurrentProcess()
    metrics("MemoryUsage") = (process.WorkingSet64 / 1024 / 1024).ToString("F2") & " MB"
    metrics("ThreadCount") = process.Threads.Count.ToString()
    metrics("HandleCount") = process.HandleCount.ToString()
    
    ' Get database connection pool info
    Try
        Dim con As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings("hcidb").ConnectionString)
        con.Open()
        metrics("DatabaseConnections") = "Connected"
        con.Close()
    Catch ex As Exception
        metrics("DatabaseConnections") = "Error: " & ex.Message
    End Try
    
    Return String.Join("; ", metrics.Select(Function(kvp) kvp.Key & "=" & kvp.Value))
End Function
```

## Troubleshooting

### 1. Common Issues

#### MelissaData License Issues
**Problem**: "MelissaData Data License Expired" error
**Solution**:
1. Check license expiration date
2. Renew license with MelissaData
3. Update license key in web.config
4. Restart application pool

#### Database Connection Issues
**Problem**: "Login failed" or connection timeout errors
**Solution**:
1. Verify connection string parameters
2. Check SQL Server service status
3. Verify firewall settings
4. Test connection with SQL Server Management Studio

#### Geocoding Service Failures
**Problem**: Geocoding requests failing
**Solution**:
1. Verify API keys are valid
2. Check API quota limits
3. Test internet connectivity
4. Review proxy settings if applicable

### 2. Debug Mode

#### Enable Comprehensive Debugging
```xml
<appSettings>
    <add key="CleanOrganization_debug" value="Y"/>
    <add key="CleanContact_debug" value="Y"/>
    <add key="CleanAddress_debug" value="Y"/>
    <add key="CleanAddress_detailed_debug" value="Y"/>
    <add key="StandardizeOrganization_debug" value="Y"/>
    <add key="StandardizeContact_debug" value="Y"/>
    <add key="StandardizeAddress_debug" value="Y"/>
</appSettings>
```

#### Review Debug Logs
```powershell
# Monitor debug logs in real-time
Get-Content "C:\Logs\DataCleaningService\Service.log" -Wait -Tail 50
```

### 3. Performance Issues

#### Database Performance
```sql
-- Check for blocking queries
SELECT 
    session_id,
    blocking_session_id,
    wait_type,
    wait_time,
    command,
    text
FROM sys.dm_exec_requests r
CROSS APPLY sys.dm_exec_sql_text(r.sql_handle)
WHERE blocking_session_id > 0;

-- Check index usage
SELECT 
    i.name AS IndexName,
    s.user_seeks,
    s.user_scans,
    s.user_lookups,
    s.user_updates
FROM sys.indexes i
LEFT JOIN sys.dm_db_index_usage_stats s ON i.object_id = s.object_id AND i.index_id = s.index_id
WHERE i.object_id = OBJECT_ID('S_CONTACT');
```

#### Application Performance
```vb
' Add performance timing to methods
Private Function TimeMethod(Of T)(ByVal method As Func(Of T)) As T
    Dim stopwatch As Stopwatch = Stopwatch.StartNew()
    Try
        Return method()
    Finally
        stopwatch.Stop()
        LogPerformance("Method execution time: " & stopwatch.ElapsedMilliseconds & "ms")
    End Try
End Function
```

## Conclusion

This customization guide provides comprehensive instructions for implementing the Data Cleaning Service in your environment. Follow the steps in order, and ensure all prerequisites are met before proceeding to the next section.

### Key Success Factors:
1. **Proper Planning**: Ensure all licenses and prerequisites are in place
2. **Security First**: Implement proper authentication and authorization
3. **Performance Monitoring**: Set up comprehensive monitoring from day one
4. **Regular Maintenance**: Schedule regular updates and maintenance windows
5. **Documentation**: Keep detailed records of all customizations and configurations

### Support Resources:
- MelissaData Support: [support.melissadata.com](https://support.melissadata.com)
- Microsoft IIS Documentation: [docs.microsoft.com/iis](https://docs.microsoft.com/iis)
- SQL Server Documentation: [docs.microsoft.com/sql](https://docs.microsoft.com/sql)

For additional support or customization requirements, consult with your development team or contact the service maintainers.
