# Certificate Service Application

## Overview

The Certificate Service Application is a VB.NET web service that provides on-demand generation of customized certificates and forms using Crystal Reports. The application supports multiple output formats (PDF, images) and delivery methods (web portal, email) for various types of training and certification programs.

## Architecture

### Technology Stack
- **Framework**: .NET Framework 4.0
- **Language**: VB.NET
- **Reporting Engine**: Crystal Reports 14.0
- **Database**: SQL Server
- **Web Service**: ASMX Web Services
- **Logging**: log4net
- **PDF Processing**: Custom PDF libraries (verywrite.dll, pdf2image.dll)

### Key Components

1. **Web Service (`Service.asmx`)**: Main entry point for certificate generation requests
2. **Service Logic (`App_Code/Service.vb`)**: Core business logic and report processing
3. **Crystal Reports**: Template-based report generation system
4. **Database Layer**: Multi-database architecture supporting different data sources
5. **Document Management System (DMS)**: File storage and metadata management

## Database Schema

The application uses a multi-database architecture:

### Primary Databases

#### `siebeldb` - Main Application Database
- **CX_CERT_PROD_QUEUE**: Certificate production queue management
- **CX_CERT_PROD_RESULTS**: Generated certificate results and metadata
- **CX_CERT_PROD_ID_POOL**: Certificate ID pool for unique identifier management
- **CX_CERT_PROD**: Certificate product definitions and templates
- **CX_CERT_PROD_CRSE**: Course-to-certificate mappings
- **S_CONTACT**: Contact and participant information
- **S_CRSE**: Course definitions and metadata
- **S_ORG_EXT**: Organization information

#### `DMS` - Document Management System
- **Documents**: Document metadata and file information
- **Document_Versions**: File versions and binary data storage
- **Document_Categories**: Document categorization system
- **Document_Keywords**: Searchable keyword associations
- **Document_Associations**: Document-to-entity relationships

#### `reports` - Reporting Database
- **TEST_CERT_PROD**: Test data for development and testing

#### `scanner` - Email and Communication
- **EmailLog**: Email delivery tracking and status

## Key Features

### 1. Certificate Generation
- **On-demand Processing**: Real-time certificate generation from queue
- **Multiple Formats**: PDF, JPEG, TIFF output support
- **Template-based**: Crystal Reports templates for consistent formatting
- **Batch Processing**: Support for bulk certificate generation

### 2. Document Management
- **Version Control**: Document versioning and history tracking
- **Metadata Management**: Rich metadata for search and organization
- **Association Tracking**: Links between documents and entities (contacts, courses, organizations)
- **Category System**: Hierarchical document categorization

### 3. Integration Capabilities
- **Web Service API**: SOAP-based web service for external integration
- **Email Delivery**: Automated email delivery of generated certificates
- **Web Portal Integration**: Direct web portal delivery support
- **External Service Integration**: Integration with cloud services and external systems

### 4. Reporting and Analytics
- **Crystal Reports Integration**: Advanced reporting capabilities
- **SAP BusinessObjects Integration**: Enterprise reporting platform support
- **Custom Report Templates**: Multiple certificate and form templates
- **Data Export**: Support for various data export formats

## Configuration

### Web.config Settings

#### Application Settings
```xml
<appSettings>
    <add key="basepath" value="C:\Inetpub\CertSvc\"/>
    <add key="reports" value="C:\Inetpub\CertSvc\Reports"/>
    <add key="cms" value="@SAP-BOBI"/>
    <add key="GenCertProd_RASFolderID" value="10687"/>
</appSettings>
```

#### Database Connections
```xml
<connectionStrings>
    <add name="siebeldb" connectionString="server=DBSERVER\DBINSTANCE;uid=dbuser;pwd=dbpassword;database=siebeldb;Min Pool Size=3;Max Pool Size=5" providerName="System.Data.SqlClient"/>
    <add name="dms" connectionString="server=DBSERVER\DBINSTANCE;uid=dmsuser;pwd=dmspassword;Min Pool Size=3;Max Pool Size=5;Connect Timeout=10;database=DMS" providerName="System.Data.SqlClient"/>
    <add name="email" connectionString="server=DBSERVER\DBINSTANCE;uid=dbuser;pwd=dbpassword;database=scanner;Min Pool Size=3;Max Pool Size=5" providerName="System.Data.SqlClient"/>
</connectionStrings>
```

### Logging Configuration
The application uses log4net for comprehensive logging:
- **Remote Syslog**: Centralized logging to syslog server
- **File Logging**: Local file-based logging with rotation
- **Multiple Loggers**: Separate loggers for different components
- **Debug Support**: Configurable debug logging levels

## API Reference

### Web Service Methods

#### GenerateCertificate
Generates a single certificate based on provided parameters.

**Parameters:**
- `ConId`: Contact ID
- `CrseId`: Course ID
- `ProdId`: Product ID
- `OutputDest`: Output destination (web/email)
- `JurisId`: Jurisdiction ID
- `OrgId`: Organization ID

**Returns:** Certificate generation status and metadata

#### GenerateCurrentCards
Generates current certification cards for participants.

**Parameters:**
- `ConId`: Contact ID
- `OutputDest`: Output destination

**Returns:** Card generation results

#### ScheduledGenerateCertificate
Scheduled certificate generation for batch processing.

**Parameters:**
- `QueueId`: Queue identifier
- `Debug`: Debug mode flag

**Returns:** Processing status

## File Structure

```
CertSvc/
├── App_Code/
│   └── Service.vb                 # Main service logic
├── App_Data/
│   └── PublishProfiles/          # Deployment profiles
├── App_WebReferences/            # Web service references
├── Bin/                          # Compiled assemblies
├── reports/                      # Crystal Reports templates
│   ├── *.rpt                     # Report templates
│   ├── *.ttx                     # Report data definitions
│   └── *.xsd                     # XML schema definitions
├── temp/                         # Temporary files
├── web.config                    # Application configuration
├── Service.asmx                  # Web service entry point
└── help.aspx                     # Service help page
```

## Report Templates

The application includes numerous Crystal Reports templates for different certificate types:

### Certificate Types
- **Completion Certificates**: Course completion certificates
- **Training Cards**: Training completion cards
- **Exam Certificates**: Examination completion certificates
- **Part Cards**: Component-specific certificates
- **Session Rosters**: Training session participant lists

### Template Features
- **Dynamic Data Binding**: Real-time data integration
- **Multi-format Output**: PDF, image, and other formats
- **Customizable Layouts**: Flexible template design
- **Jurisdiction Support**: State/region-specific formatting

## Installation and Setup

### Prerequisites
- Windows Server with IIS
- .NET Framework 4.0
- SQL Server (2008 or later)
- Crystal Reports Runtime 14.0
- SAP BusinessObjects (optional)

### Installation Steps

1. **Database Setup**
   ```sql
   -- Run the provided database_schema.sql script
   -- Create required databases and tables
   -- Set up user permissions
   ```

2. **Application Deployment**
   - Deploy application files to IIS
   - Configure application pool for .NET Framework 4.0
   - Set appropriate permissions for file system access

3. **Configuration**
   - Update web.config with environment-specific settings
   - Configure database connection strings
   - Set up logging destinations
   - Configure external service endpoints

4. **Crystal Reports Setup**
   - Install Crystal Reports runtime
   - Configure report server connections
   - Set up report folder permissions

### Environment Configuration

#### Development Environment
- Use test database connections
- Enable debug logging
- Configure local file paths

#### Production Environment
- Use production database connections
- Configure centralized logging
- Set up monitoring and alerting
- Implement backup procedures

## Security Considerations

### Authentication and Authorization
- Windows Authentication for web service access
- Database-level user permissions
- File system access controls

### Data Protection
- Encrypted database connections
- Secure file storage
- Audit logging for sensitive operations

### Network Security
- HTTPS for web service communication
- Firewall configuration for database access
- Secure external service integration

## Monitoring and Maintenance

### Logging
- Application logs via log4net
- Database transaction logs
- System performance monitoring

### Backup Procedures
- Regular database backups
- Report template backups
- Configuration file backups

### Performance Optimization
- Database indexing strategy
- Connection pooling configuration
- Report caching mechanisms

## Troubleshooting

### Common Issues

#### Certificate Generation Failures
- Check database connectivity
- Verify report template availability
- Review file system permissions
- Check Crystal Reports runtime status

#### Performance Issues
- Monitor database performance
- Check connection pool settings
- Review report processing times
- Analyze system resource usage

#### Integration Problems
- Verify external service connectivity
- Check authentication credentials
- Review network configuration
- Validate data format compatibility

### Debug Mode
Enable debug mode in configuration:
```xml
<add key="GenCertProd_debug" value="Y"/>
<add key="SchedGenCertProd_debug" value="Y"/>
<add key="GenCurrentCards_debug" value="Y"/>
```

## Development Guidelines

### Code Structure
- Follow VB.NET coding standards
- Implement proper error handling
- Use consistent naming conventions
- Document complex business logic

### Database Design
- Use appropriate data types
- Implement proper indexing
- Follow normalization principles
- Maintain referential integrity

### Testing
- Unit testing for business logic
- Integration testing for database operations
- Performance testing for report generation
- User acceptance testing for end-to-end workflows

## Support and Maintenance

### Regular Maintenance Tasks
- Database maintenance and optimization
- Log file cleanup and rotation
- Performance monitoring and tuning
- Security updates and patches

### Version Control
- Track configuration changes
- Maintain deployment documentation
- Document custom modifications
- Version control for report templates

## License and Compliance

This application is designed for internal use and may be subject to various licensing requirements for:
- Crystal Reports runtime
- SAP BusinessObjects components
- Third-party PDF processing libraries
- Database licensing

Ensure compliance with all applicable licenses and regulations in your deployment environment.

---

*This documentation provides a comprehensive overview of the Certificate Service Application. For specific implementation details or troubleshooting assistance, refer to the source code and configuration files.*
