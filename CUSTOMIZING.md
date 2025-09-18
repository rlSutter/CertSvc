# Certificate Service Application - Customization Guide

## Table of Contents
1. [Overview](#overview)
2. [Configuration Customization](#configuration-customization)
3. [Database Customization](#database-customization)
4. [Report Template Customization](#report-template-customization)
5. [Web Service Customization](#web-service-customization)
6. [UI and Branding Customization](#ui-and-branding-customization)
7. [Integration Customization](#integration-customization)
8. [Security Customization](#security-customization)
9. [Performance Customization](#performance-customization)
10. [Deployment Customization](#deployment-customization)

## Overview

This guide provides comprehensive instructions for customizing the Certificate Service Application to meet specific organizational requirements. The application is designed with flexibility in mind, allowing customization at multiple levels from configuration to core business logic.

## Configuration Customization

### Web.config Modifications

#### Database Connection Customization
```xml
<connectionStrings>
    <!-- Customize database connections for your environment -->
    <add name="siebeldb" 
         connectionString="server=YOUR_SERVER\INSTANCE;uid=YOUR_USER;pwd=YOUR_PASSWORD;database=siebeldb;Min Pool Size=5;Max Pool Size=20" 
         providerName="System.Data.SqlClient"/>
    
    <!-- Add custom connection strings for additional databases -->
    <add name="customdb" 
         connectionString="server=YOUR_SERVER;database=CustomDB;Integrated Security=true" 
         providerName="System.Data.SqlClient"/>
</connectionStrings>
```

#### Application Settings Customization
```xml
<appSettings>
    <!-- Customize file paths -->
    <add key="basepath" value="C:\YourAppPath\"/>
    <add key="reports" value="C:\YourAppPath\Reports"/>
    <add key="temp" value="C:\YourAppPath\Temp"/>
    
    <!-- Customize external service endpoints -->
    <add key="com.yourcompany.service" value="http://your-service.com/api"/>
    <add key="custom.integration.endpoint" value="https://api.yourcompany.com/v1"/>
    
    <!-- Customize debug settings -->
    <add key="GenCertProd_debug" value="Y"/>
    <add key="CustomModule_debug" value="Y"/>
    
    <!-- Add custom application settings -->
    <add key="CustomFeature_Enabled" value="true"/>
    <add key="CustomTimeout_Seconds" value="300"/>
</appSettings>
```

#### Logging Configuration Customization
```xml
<log4net>
    <!-- Customize log file locations -->
    <appender name="CustomLogAppender" type="log4net.Appender.RollingFileAppender">
        <file value="C:\Logs\CustomApp.log"/>
        <appendToFile value="true"/>
        <rollingStyle value="Date"/>
        <datePattern value="yyyyMMdd"/>
        <layout type="log4net.Layout.PatternLayout">
            <conversionPattern value="%date [%thread] %-5level %logger - %message%newline"/>
        </layout>
    </appender>
    
    <!-- Add custom logger -->
    <logger name="CustomModule">
        <level value="DEBUG"/>
        <appender-ref ref="CustomLogAppender"/>
    </logger>
</log4net>
```

## Database Customization

### Adding Custom Tables

#### 1. Create Custom Tables
```sql
-- Example: Custom certificate metadata table
CREATE TABLE [dbo].[CX_CUSTOM_CERT_METADATA] (
    [ROW_ID] NVARCHAR(15) NOT NULL PRIMARY KEY,
    [CERT_ID] NVARCHAR(15) NULL,
    [CUSTOM_FIELD_1] NVARCHAR(100) NULL,
    [CUSTOM_FIELD_2] NVARCHAR(100) NULL,
    [CUSTOM_DATE] DATETIME NULL,
    [CUSTOM_FLAG] NVARCHAR(1) NULL,
    [CREATED] DATETIME NULL,
    [CREATED_BY] NVARCHAR(15) NULL
);

-- Example: Custom organization settings
CREATE TABLE [dbo].[CX_CUSTOM_ORG_SETTINGS] (
    [ROW_ID] NVARCHAR(15) NOT NULL PRIMARY KEY,
    [ORG_ID] NVARCHAR(15) NULL,
    [SETTING_NAME] NVARCHAR(50) NULL,
    [SETTING_VALUE] NVARCHAR(255) NULL,
    [ACTIVE_FLG] NVARCHAR(1) NULL
);
```

#### 2. Modify Existing Tables
```sql
-- Add custom columns to existing tables
ALTER TABLE [dbo].[CX_CERT_PROD] 
ADD [CUSTOM_TEMPLATE_PATH] NVARCHAR(255) NULL,
    [CUSTOM_BRANDING_ID] NVARCHAR(15) NULL,
    [CUSTOM_VALIDATION_RULES] NVARCHAR(500) NULL;

-- Add indexes for performance
CREATE INDEX [IX_CX_CERT_PROD_CUSTOM_BRANDING] 
ON [dbo].[CX_CERT_PROD] ([CUSTOM_BRANDING_ID]);
```

### Custom Stored Procedures
```sql
-- Example: Custom certificate validation procedure
CREATE PROCEDURE [dbo].[sp_ValidateCustomCertificate]
    @CertId NVARCHAR(15),
    @ValidationRules NVARCHAR(500)
AS
BEGIN
    -- Custom validation logic
    DECLARE @IsValid BIT = 1;
    
    -- Add your custom validation rules here
    IF EXISTS (SELECT 1 FROM CX_CUSTOM_CERT_METADATA 
               WHERE CERT_ID = @CertId AND CUSTOM_FLAG = 'N')
    BEGIN
        SET @IsValid = 0;
    END
    
    SELECT @IsValid AS IsValid;
END;
```

## Report Template Customization

### Creating Custom Crystal Reports

#### 1. Report Structure
```
reports/
├── custom/
│   ├── CustomCertificate.rpt
│   ├── CustomCertificate.ttx
│   ├── CustomCard.rpt
│   └── CustomCard.ttx
├── templates/
│   ├── CustomTemplate.rpt
│   └── CustomTemplate.ttx
└── images/
    ├── custom-logo.png
    └── custom-watermark.png
```

#### 2. Custom Report Template Example
```vb
' In Service.vb - Add custom report processing method
Public Function GenerateCustomCertificate(ByVal ConId As String, 
                                        ByVal CustomTemplate As String, 
                                        ByVal CustomData As String) As String
    Try
        Dim reportPath As String = ConfigurationManager.AppSettings("reports") & "\custom\" & CustomTemplate & ".rpt"
        
        ' Load custom report
        Dim report As New ReportDocument()
        report.Load(reportPath)
        
        ' Set custom parameters
        report.SetParameterValue("CustomData", CustomData)
        report.SetParameterValue("GeneratedDate", DateTime.Now)
        
        ' Process custom data
        Dim customDataSet As DataSet = GetCustomData(ConId)
        report.SetDataSource(customDataSet)
        
        ' Generate output
        Dim outputPath As String = GenerateOutput(report, "PDF")
        
        Return outputPath
    Catch ex As Exception
        LogError("Custom certificate generation failed: " & ex.Message)
        Throw
    End Try
End Function
```

#### 3. Custom Data Source Integration
```vb
Private Function GetCustomData(ByVal ConId As String) As DataSet
    Dim ds As New DataSet()
    Dim dt As New DataTable("CustomData")
    
    ' Add custom columns
    dt.Columns.Add("CustomField1", GetType(String))
    dt.Columns.Add("CustomField2", GetType(String))
    dt.Columns.Add("CustomDate", GetType(DateTime))
    
    ' Query custom data
    Dim sql As String = "SELECT * FROM CX_CUSTOM_CERT_METADATA WHERE CERT_ID = @CertId"
    ' Execute query and populate DataTable
    
    ds.Tables.Add(dt)
    Return ds
End Function
```

## Web Service Customization

### Adding Custom Web Methods

#### 1. Custom Service Methods
```vb
<WebMethod()> _
Public Function GetCustomCertificateData(ByVal ConId As String, 
                                        ByVal CustomFilters As String) As String
    Try
        ' Custom business logic
        Dim customData As String = ProcessCustomRequest(ConId, CustomFilters)
        Return customData
    Catch ex As Exception
        LogError("Custom service method failed: " & ex.Message)
        Return "ERROR: " & ex.Message
    End Try
End Function

<WebMethod()> _
Public Function ValidateCustomCertificate(ByVal CertId As String, 
                                         ByVal ValidationType As String) As Boolean
    Try
        ' Custom validation logic
        Return PerformCustomValidation(CertId, ValidationType)
    Catch ex As Exception
        LogError("Custom validation failed: " & ex.Message)
        Return False
    End Try
End Function
```

#### 2. Custom Data Processing
```vb
Private Function ProcessCustomRequest(ByVal ConId As String, 
                                     ByVal Filters As String) As String
    Dim result As New StringBuilder()
    
    ' Parse custom filters
    Dim filterArray As String() = Filters.Split("|"c)
    
    ' Process each filter
    For Each filter As String In filterArray
        Dim processedData As String = ProcessFilter(ConId, filter)
        result.Append(processedData & ";")
    Next
    
    Return result.ToString()
End Function
```

### Custom API Endpoints
```vb
' Add custom REST-like endpoints
<WebMethod()> _
Public Function CustomAPI(ByVal Action As String, 
                         ByVal Parameters As String) As String
    Try
        Select Case Action.ToUpper()
            Case "GETCUSTOMDATA"
                Return GetCustomDataAPI(Parameters)
            Case "UPDATECUSTOMDATA"
                Return UpdateCustomDataAPI(Parameters)
            Case "DELETECUSTOMDATA"
                Return DeleteCustomDataAPI(Parameters)
            Case Else
                Return "ERROR: Unknown action"
        End Select
    Catch ex As Exception
        Return "ERROR: " & ex.Message
    End Try
End Function
```

## UI and Branding Customization

### Custom Styling and Branding

#### 1. Custom CSS and JavaScript
```html
<!-- Add to help.aspx or create custom UI pages -->
<style>
.custom-header {
    background-color: #your-brand-color;
    color: white;
    padding: 10px;
    font-family: 'Your Brand Font', Arial, sans-serif;
}

.custom-logo {
    max-height: 50px;
    margin-right: 10px;
}

.custom-button {
    background-color: #your-accent-color;
    border: none;
    padding: 8px 16px;
    border-radius: 4px;
    color: white;
    cursor: pointer;
}
</style>

<script>
function customFunction() {
    // Custom JavaScript functionality
    console.log('Custom function executed');
}

// Custom API integration
function callCustomAPI(action, parameters) {
    // Implementation for custom API calls
}
</script>
```

#### 2. Custom Help Pages
```html
<!-- Create custom help pages -->
<div class="custom-header">
    <img src="images/custom-logo.png" class="custom-logo" alt="Your Company Logo">
    <h1>Certificate Service - Custom Help</h1>
</div>

<div class="custom-content">
    <h2>Custom Features</h2>
    <ul>
        <li>Custom certificate templates</li>
        <li>Branded output formats</li>
        <li>Custom validation rules</li>
    </ul>
</div>
```

## Integration Customization

### External System Integration

#### 1. Custom Web Service Integration
```vb
' Add custom web service client
Public Class CustomServiceClient
    Private serviceUrl As String
    
    Public Sub New(url As String)
        serviceUrl = url
    End Sub
    
    Public Function CallCustomService(ByVal method As String, 
                                     ByVal parameters As String) As String
        Try
            ' Custom web service call implementation
            Dim client As New WebClient()
            client.Headers.Add("Content-Type", "application/json")
            
            Dim requestData As String = CreateRequestData(method, parameters)
            Dim response As String = client.UploadString(serviceUrl, requestData)
            
            Return response
        Catch ex As Exception
            LogError("Custom service call failed: " & ex.Message)
            Return "ERROR: " & ex.Message
        End Try
    End Function
End Class
```

#### 2. Custom Email Integration
```vb
' Custom email service
Public Class CustomEmailService
    Public Function SendCustomEmail(ByVal toAddress As String, 
                                   ByVal subject As String, 
                                   ByVal body As String, 
                                   ByVal attachments As List(Of String)) As Boolean
        Try
            Dim message As New MailMessage()
            message.To.Add(toAddress)
            message.Subject = subject
            message.Body = body
            message.IsBodyHtml = True
            
            ' Add custom branding
            message.Body = ApplyCustomBranding(message.Body)
            
            ' Add attachments
            For Each attachment As String In attachments
                message.Attachments.Add(New Attachment(attachment))
            Next
            
            ' Send email using custom SMTP settings
            Dim smtp As New SmtpClient(GetCustomSMTPServer())
            smtp.Send(message)
            
            Return True
        Catch ex As Exception
            LogError("Custom email failed: " & ex.Message)
            Return False
        End Try
    End Function
End Class
```

## Security Customization

### Custom Authentication and Authorization

#### 1. Custom Authentication Provider
```vb
' Custom authentication logic
Public Class CustomAuthenticationProvider
    Public Function ValidateUser(ByVal username As String, 
                                ByVal password As String) As Boolean
        Try
            ' Custom authentication logic
            Dim isValid As Boolean = CheckCustomCredentials(username, password)
            
            If isValid Then
                LogAuthentication(username, "SUCCESS")
            Else
                LogAuthentication(username, "FAILED")
            End If
            
            Return isValid
        Catch ex As Exception
            LogError("Authentication error: " & ex.Message)
            Return False
        End Try
    End Function
    
    Private Function CheckCustomCredentials(ByVal username As String, 
                                          ByVal password As String) As Boolean
        ' Implement custom credential validation
        ' This could integrate with Active Directory, custom database, etc.
        Return True ' Placeholder
    End Function
End Class
```

#### 2. Custom Authorization Rules
```vb
' Custom authorization logic
Public Class CustomAuthorizationProvider
    Public Function CheckPermission(ByVal userId As String, 
                                   ByVal resource As String, 
                                   ByVal action As String) As Boolean
        Try
            ' Custom authorization logic
            Dim hasPermission As Boolean = QueryCustomPermissions(userId, resource, action)
            Return hasPermission
        Catch ex As Exception
            LogError("Authorization error: " & ex.Message)
            Return False
        End Try
    End Function
End Class
```

## Performance Customization

### Caching and Optimization

#### 1. Custom Caching Implementation
```vb
' Custom caching service
Public Class CustomCacheService
    Private cache As New Dictionary(Of String, Object)
    Private cacheExpiry As New Dictionary(Of String, DateTime)
    
    Public Function GetCachedData(ByVal key As String) As Object
        If cache.ContainsKey(key) AndAlso cacheExpiry(key) > DateTime.Now Then
            Return cache(key)
        End If
        Return Nothing
    End Function
    
    Public Sub SetCachedData(ByVal key As String, 
                            ByVal data As Object, 
                            ByVal expiryMinutes As Integer)
        cache(key) = data
        cacheExpiry(key) = DateTime.Now.AddMinutes(expiryMinutes)
    End Sub
End Class
```

#### 2. Custom Performance Monitoring
```vb
' Custom performance monitoring
Public Class CustomPerformanceMonitor
    Public Function MeasureExecutionTime(ByVal operation As String, 
                                        ByVal action As Action) As TimeSpan
        Dim stopwatch As New Stopwatch()
        stopwatch.Start()
        
        Try
            action.Invoke()
        Finally
            stopwatch.Stop()
            LogPerformance(operation, stopwatch.Elapsed)
        End Try
        
        Return stopwatch.Elapsed
    End Function
End Class
```

## Deployment Customization

### Custom Deployment Scripts

#### 1. Custom Build Script
```batch
@echo off
REM Custom deployment script
echo Starting custom deployment...

REM Backup existing files
xcopy "C:\Inetpub\CertSvc" "C:\Backup\CertSvc_%date%" /E /I

REM Deploy new files
xcopy ".\Deploy\*" "C:\Inetpub\CertSvc\" /E /Y

REM Update configuration
powershell -ExecutionPolicy Bypass -File ".\Scripts\UpdateConfig.ps1"

REM Restart application pool
%windir%\system32\inetsrv\appcmd recycle apppool "CertSvcAppPool"

echo Deployment completed successfully!
```

#### 2. Custom Configuration Updates
```powershell
# UpdateConfig.ps1 - Custom configuration update script
param(
    [string]$Environment = "Production"
)

$configPath = "C:\Inetpub\CertSvc\web.config"
$config = [xml](Get-Content $configPath)

# Update connection strings based on environment
switch ($Environment) {
    "Production" {
        $config.configuration.connectionStrings.add[0].connectionString = "server=PROD_SERVER;database=siebeldb;..."
    }
    "Staging" {
        $config.configuration.connectionStrings.add[0].connectionString = "server=STAGE_SERVER;database=siebeldb;..."
    }
    "Development" {
        $config.configuration.connectionStrings.add[0].connectionString = "server=DEV_SERVER;database=siebeldb;..."
    }
}

# Save updated configuration
$config.Save($configPath)
Write-Host "Configuration updated for $Environment environment"
```

## Testing Customizations

### Custom Test Framework

#### 1. Unit Testing
```vb
' Custom unit test example
<TestClass()>
Public Class CustomCertificateServiceTests
    <TestMethod()>
    Public Sub TestCustomCertificateGeneration()
        ' Arrange
        Dim service As New CertificateService()
        Dim conId As String = "TEST_CON_001"
        Dim customTemplate As String = "CustomTemplate"
        
        ' Act
        Dim result As String = service.GenerateCustomCertificate(conId, customTemplate, "")
        
        ' Assert
        Assert.IsNotNull(result)
        Assert.IsTrue(result.Contains("SUCCESS"))
    End Sub
End Class
```

#### 2. Integration Testing
```vb
' Custom integration test
<TestClass()>
Public Class CustomIntegrationTests
    <TestMethod()>
    Public Sub TestCustomDatabaseIntegration()
        ' Test custom database operations
        Dim connectionString As String = GetTestConnectionString()
        Dim result As Boolean = TestCustomDatabaseOperations(connectionString)
        
        Assert.IsTrue(result)
    End Sub
End Class
```

## Best Practices for Customization

### 1. Code Organization
- Keep custom code in separate modules/classes
- Use consistent naming conventions
- Document all custom modifications
- Implement proper error handling

### 2. Configuration Management
- Use environment-specific configuration files
- Implement configuration validation
- Document all configuration changes
- Use secure storage for sensitive data

### 3. Testing Strategy
- Implement comprehensive unit tests
- Create integration tests for custom features
- Use automated testing where possible
- Maintain test data and environments

### 4. Documentation
- Document all custom modifications
- Maintain change logs
- Create user guides for custom features
- Keep technical documentation updated

### 5. Version Control
- Use proper version control practices
- Tag releases and deployments
- Maintain branching strategies
- Document merge procedures

## Troubleshooting Customizations

### Common Issues and Solutions

#### 1. Configuration Issues
```vb
' Add configuration validation
Public Function ValidateCustomConfiguration() As Boolean
    Try
        ' Validate custom settings
        Dim customSetting As String = ConfigurationManager.AppSettings("CustomFeature_Enabled")
        If String.IsNullOrEmpty(customSetting) Then
            LogError("CustomFeature_Enabled setting is missing")
            Return False
        End If
        
        Return True
    Catch ex As Exception
        LogError("Configuration validation failed: " & ex.Message)
        Return False
    End Try
End Function
```

#### 2. Database Connection Issues
```vb
' Add database connection testing
Public Function TestCustomDatabaseConnection() As Boolean
    Try
        Using connection As New SqlConnection(GetCustomConnectionString())
            connection.Open()
            Return True
        End Using
    Catch ex As Exception
        LogError("Custom database connection failed: " & ex.Message)
        Return False
    End Try
End Function
```

This customization guide provides a comprehensive framework for extending and modifying the Certificate Service Application to meet specific organizational needs while maintaining system stability and performance.
