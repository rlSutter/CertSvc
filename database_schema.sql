-- =============================================
-- Certificate Service Database Schema
-- Generated from VB.NET Web Service Analysis
-- =============================================

-- Create databases
CREATE DATABASE [siebeldb];
GO

CREATE DATABASE [DMS];
GO

CREATE DATABASE [reports];
GO

CREATE DATABASE [scanner];
GO

-- =============================================
-- SIEBELDB DATABASE SCHEMA
-- =============================================

USE [siebeldb];
GO

-- Certificate Production Queue Table
CREATE TABLE [dbo].[CX_CERT_PROD_QUEUE] (
    [ROW_ID] NVARCHAR(15) NOT NULL PRIMARY KEY,
    [CONFLICT_ID] NVARCHAR(15) NULL,
    [CREATED] DATETIME NULL,
    [CREATED_BY] NVARCHAR(15) NULL,
    [LAST_UPD] DATETIME NULL,
    [LAST_UPD_BY] NVARCHAR(15) NULL,
    [MODIFICATION_NUM] INT NULL,
    [PROD_TYPE] NVARCHAR(1) NULL,
    [TRAIN_TYPE] NVARCHAR(10) NULL,
    [CRSE_ID] NVARCHAR(15) NULL,
    [SRC_ID] NVARCHAR(15) NULL,
    [IDENT_START] NVARCHAR(15) NULL,
    [IDENT_END] NVARCHAR(15) NULL,
    [NOTIFY_FLG] NVARCHAR(1) NULL,
    [ATTACH_FLG] NVARCHAR(1) NULL,
    [MULTI_OUT_FLG] NVARCHAR(1) NULL,
    [FORMAT] NVARCHAR(10) NULL,
    [DOMAIN] NVARCHAR(50) NULL,
    [JURIS_ID] NVARCHAR(15) NULL,
    [CON_ID] NVARCHAR(15) NULL,
    [OU_ID] NVARCHAR(15) NULL,
    [DEST_CODE] NVARCHAR(10) NULL,
    [PROD_ID] NVARCHAR(15) NULL,
    [ACCESS_FLG] NVARCHAR(1) NULL,
    [REQD_FLG] NVARCHAR(1) NULL,
    [SPECIAL_MSG] NVARCHAR(255) NULL,
    [EMP_ID] NVARCHAR(15) NULL,
    [DOC_ID] NVARCHAR(15) NULL,
    [EXECUTED] DATETIME NULL
);

-- Certificate Production Results Table
CREATE TABLE [dbo].[CX_CERT_PROD_RESULTS] (
    [ROW_ID] NVARCHAR(15) NOT NULL PRIMARY KEY,
    [CERT_POOL_ID] NVARCHAR(15) NULL,
    [PROD_QUEUE_ID] NVARCHAR(15) NULL,
    [CON_ID] NVARCHAR(15) NULL,
    [CRSE_ID] NVARCHAR(15) NULL,
    [REG_ID] NVARCHAR(15) NULL,
    [GENERATED] DATETIME NULL,
    [DESTINATION] NVARCHAR(50) NULL,
    [DOC_ID] NVARCHAR(15) NULL,
    [PROD_ID] NVARCHAR(15) NULL,
    [CERT_CRSE_ID] NVARCHAR(15) NULL
);

-- Certificate Production ID Pool Table
CREATE TABLE [dbo].[CX_CERT_PROD_ID_POOL] (
    [ROW_ID] NVARCHAR(15) NOT NULL PRIMARY KEY,
    [CONFLICT_ID] NVARCHAR(15) NULL,
    [CREATED] DATETIME NULL,
    [CREATED_BY] NVARCHAR(15) NULL,
    [LAST_UPD] DATETIME NULL,
    [LAST_UPD_BY] NVARCHAR(15) NULL,
    [MODIFICATION_NUM] INT NULL,
    [JURIS_ID] NVARCHAR(15) NULL,
    [PROD_QUEUE_ID] NVARCHAR(15) NULL,
    [PROD_RESULT_ID] NVARCHAR(15) NULL,
    [JURIS_CERT_ID] NVARCHAR(50) NULL
);

-- Certificate Production Table
CREATE TABLE [dbo].[CX_CERT_PROD] (
    [ROW_ID] NVARCHAR(15) NOT NULL PRIMARY KEY,
    [PROD_TYPE] NVARCHAR(1) NULL,
    [SPECIAL_NOTICE] NVARCHAR(255) NULL,
    [RES_X] INT NULL,
    [RES_Y] INT NULL,
    [WIDTH] INT NULL,
    [HEIGHT] INT NULL
);

-- Certificate Production Course Table
CREATE TABLE [dbo].[CX_CERT_PROD_CRSE] (
    [ROW_ID] NVARCHAR(15) NOT NULL PRIMARY KEY,
    [CERT_ID] NVARCHAR(15) NULL,
    [CRSE_ID] NVARCHAR(15) NULL,
    [JURIS_ID] NVARCHAR(15) NULL,
    [JURIS_CERT_ID_FLG] NVARCHAR(1) NULL,
    [JURIS_CERT_EMAIL] NVARCHAR(1) NULL,
    [TEMP_PROD_ID] NVARCHAR(15) NULL,
    [PRIMARY_FLG] NVARCHAR(1) NULL
);

-- Participant Current Claim Table
CREATE TABLE [dbo].[CX_PART_CURRCLM] (
    [ROW_ID] NVARCHAR(15) NOT NULL PRIMARY KEY,
    [PART_ID] NVARCHAR(15) NULL,
    [CURRENT_SPART_ID] NVARCHAR(15) NULL
);

-- Participant Extended Table
CREATE TABLE [dbo].[CX_PARTICIPANT_X] (
    [ROW_ID] NVARCHAR(15) NOT NULL PRIMARY KEY,
    [JURIS_ID] NVARCHAR(15) NULL
);

-- Contact Table
CREATE TABLE [dbo].[S_CONTACT] (
    [ROW_ID] NVARCHAR(15) NOT NULL PRIMARY KEY,
    [X_PART_ID] NVARCHAR(15) NULL,
    [PR_PER_ADDR_ID] NVARCHAR(15) NULL,
    [PR_OU_ADDR_ID] NVARCHAR(15) NULL
);

-- Session Participant Extended Table
CREATE TABLE [dbo].[CX_SESS_PART_X] (
    [ROW_ID] NVARCHAR(15) NOT NULL PRIMARY KEY,
    [JURIS_ID] NVARCHAR(15) NULL,
    [CRSE_TST_ID] NVARCHAR(15) NULL,
    [SESS_ID] NVARCHAR(15) NULL
);

-- Session Registration Table
CREATE TABLE [dbo].[CX_SESS_REG] (
    [ROW_ID] NVARCHAR(15) NOT NULL PRIMARY KEY,
    [CONTACT_ID] NVARCHAR(15) NULL,
    [SESS_PART_ID] NVARCHAR(15) NULL,
    [OU_ID] NVARCHAR(15) NULL,
    [CRSE_ID] NVARCHAR(15) NULL,
    [TRAIN_OFFR_ID] NVARCHAR(15) NULL
);

-- Current Claim Per Table
CREATE TABLE [dbo].[S_CURRCLM_PER] (
    [ROW_ID] NVARCHAR(15) NOT NULL PRIMARY KEY,
    [PERSON_ID] NVARCHAR(15) NULL,
    [CURRCLM_ID] NVARCHAR(15) NULL,
    [GRANT_DT] DATETIME NULL,
    [X_CRSE_TSTRUN_ID] NVARCHAR(15) NULL
);

-- Current Claim Table
CREATE TABLE [dbo].[S_CURRCLM] (
    [ROW_ID] NVARCHAR(15) NOT NULL PRIMARY KEY
);

-- Contact Current Claim Table
CREATE TABLE [dbo].[CX_CONTACT_CURRCLM] (
    [ROW_ID] NVARCHAR(15) NOT NULL PRIMARY KEY,
    [CURRENT_CERT_ID] NVARCHAR(15) NULL,
    [CURRENT_EXP_DT] DATETIME NULL
);

-- Course Test Run Table
CREATE TABLE [dbo].[S_CRSE_TSTRUN] (
    [ROW_ID] NVARCHAR(15) NOT NULL PRIMARY KEY,
    [CRSE_TST_ID] NVARCHAR(15) NULL,
    [X_PART_ID] NVARCHAR(15) NULL
);

-- Course Test Table
CREATE TABLE [dbo].[S_CRSE_TST] (
    [ROW_ID] NVARCHAR(15) NOT NULL PRIMARY KEY,
    [CRSE_ID] NVARCHAR(15) NULL
);

-- Course Table
CREATE TABLE [dbo].[S_CRSE] (
    [ROW_ID] NVARCHAR(15) NOT NULL PRIMARY KEY,
    [X_SUMMARY_CD] NVARCHAR(10) NULL,
    [SKILL_LEVEL_CD] NVARCHAR(10) NULL
);

-- Organization Extended Table
CREATE TABLE [dbo].[S_ORG_EXT] (
    [ROW_ID] NVARCHAR(15) NOT NULL PRIMARY KEY,
    [NAME] NVARCHAR(100) NULL,
    [LOC] NVARCHAR(100) NULL
);

-- Jurisdiction Extended Table
CREATE TABLE [dbo].[CX_JURISDICTION_X] (
    [ROW_ID] NVARCHAR(15) NOT NULL PRIMARY KEY,
    [NAME] NVARCHAR(100) NULL
);

-- Address Person Table
CREATE TABLE [dbo].[S_ADDR_PER] (
    [ROW_ID] NVARCHAR(15) NOT NULL PRIMARY KEY
);

-- Address Organization Table
CREATE TABLE [dbo].[S_ADDR_ORG] (
    [ROW_ID] NVARCHAR(15) NOT NULL PRIMARY KEY
);

-- =============================================
-- DMS DATABASE SCHEMA
-- =============================================

USE [DMS];
GO

-- Documents Table
CREATE TABLE [dbo].[Documents] (
    [row_id] INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
    [ext_id] NVARCHAR(50) NULL,
    [data_type_id] INT NULL,
    [dfilename] NVARCHAR(255) NULL,
    [name] NVARCHAR(255) NULL,
    [created_by] INT NULL,
    [last_upd_by] INT NULL,
    [description] NVARCHAR(500) NULL,
    [deleted] DATETIME NULL,
    [last_version_id] INT NULL
);

-- Document Versions Table
CREATE TABLE [dbo].[Document_Versions] (
    [row_id] INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
    [doc_id] INT NULL,
    [dimage] VARBINARY(MAX) NULL,
    [dsize] INT NULL,
    [created_by] INT NULL,
    [last_upd_by] INT NULL,
    [backed_up] DATETIME NULL
);

-- Document Categories Table
CREATE TABLE [dbo].[Document_Categories] (
    [row_id] INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
    [doc_id] INT NULL,
    [cat_id] INT NULL,
    [created_by] INT NULL,
    [last_upd_by] INT NULL,
    [pr_flag] NVARCHAR(1) NULL
);

-- Document Keywords Table
CREATE TABLE [dbo].[Document_Keywords] (
    [row_id] INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
    [doc_id] INT NULL,
    [key_id] INT NULL,
    [created_by] INT NULL,
    [last_upd_by] INT NULL,
    [val] NVARCHAR(255) NULL,
    [pr_flag] NVARCHAR(1) NULL
);

-- Document Associations Table
CREATE TABLE [dbo].[Document_Associations] (
    [row_id] INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
    [created_by] INT NULL,
    [last_upd_by] INT NULL,
    [association_id] INT NULL,
    [doc_id] INT NULL,
    [fkey] NVARCHAR(50) NULL,
    [pr_flag] NVARCHAR(1) NULL,
    [access_flag] NVARCHAR(1) NULL,
    [reqd_flag] NVARCHAR(1) NULL
);

-- =============================================
-- REPORTS DATABASE SCHEMA
-- =============================================

USE [reports];
GO

-- Test Certificate Production Table
CREATE TABLE [dbo].[TEST_CERT_PROD] (
    [ROW_ID] NVARCHAR(15) NOT NULL PRIMARY KEY,
    [PROD_TYPE] NVARCHAR(1) NULL,
    [CRSE_ID] NVARCHAR(15) NULL,
    [CON_ID] NVARCHAR(15) NULL,
    [JURIS_ID] NVARCHAR(15) NULL,
    [OU_ID] NVARCHAR(15) NULL,
    [DEST_CODE] NVARCHAR(10) NULL,
    [PROD_ID] NVARCHAR(15) NULL,
    [ACCESS_FLG] NVARCHAR(1) NULL,
    [REQD_FLG] NVARCHAR(1) NULL,
    [SPECIAL_MSG] NVARCHAR(255) NULL,
    [EMP_ID] NVARCHAR(15) NULL,
    [DOC_ID] NVARCHAR(15) NULL
);

-- =============================================
-- SCANNER DATABASE SCHEMA
-- =============================================

USE [scanner];
GO

-- Email tracking table (referenced in connection strings)
-- Structure inferred from usage patterns
CREATE TABLE [dbo].[EmailLog] (
    [ID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
    [EmailAddress] NVARCHAR(255) NULL,
    [Subject] NVARCHAR(500) NULL,
    [SentDate] DATETIME NULL,
    [Status] NVARCHAR(50) NULL,
    [ErrorMessage] NVARCHAR(1000) NULL
);

-- =============================================
-- INDEXES
-- =============================================

USE [siebeldb];
GO

-- Indexes for performance optimization
CREATE INDEX [IX_CX_CERT_PROD_QUEUE_CRSE_SRC] ON [dbo].[CX_CERT_PROD_QUEUE] ([CRSE_ID], [SRC_ID]);
CREATE INDEX [IX_CX_CERT_PROD_QUEUE_EXECUTED] ON [dbo].[CX_CERT_PROD_QUEUE] ([EXECUTED]);
CREATE INDEX [IX_CX_CERT_PROD_RESULTS_QUEUE] ON [dbo].[CX_CERT_PROD_RESULTS] ([PROD_QUEUE_ID]);
CREATE INDEX [IX_CX_CERT_PROD_RESULTS_CON_CRSE] ON [dbo].[CX_CERT_PROD_RESULTS] ([CON_ID], [CERT_CRSE_ID]);
CREATE INDEX [IX_CX_CERT_PROD_ID_POOL_JURIS] ON [dbo].[CX_CERT_PROD_ID_POOL] ([JURIS_ID]);
CREATE INDEX [IX_CX_CERT_PROD_ID_POOL_QUEUE] ON [dbo].[CX_CERT_PROD_ID_POOL] ([PROD_QUEUE_ID]);
CREATE INDEX [IX_S_CONTACT_X_PART_ID] ON [dbo].[S_CONTACT] ([X_PART_ID]);
CREATE INDEX [IX_CX_SESS_PART_X_CRSE] ON [dbo].[CX_SESS_PART_X] ([CRSE_TST_ID]);
CREATE INDEX [IX_S_CURRCLM_PER_PERSON] ON [dbo].[S_CURRCLM_PER] ([PERSON_ID]);
CREATE INDEX [IX_S_CURRCLM_PER_GRANT_DT] ON [dbo].[S_CURRCLM_PER] ([GRANT_DT]);

USE [DMS];
GO

CREATE INDEX [IX_Documents_ext_id] ON [dbo].[Documents] ([ext_id]);
CREATE INDEX [IX_Documents_data_type] ON [dbo].[Documents] ([data_type_id]);
CREATE INDEX [IX_Document_Versions_doc_id] ON [dbo].[Document_Versions] ([doc_id]);
CREATE INDEX [IX_Document_Categories_doc_id] ON [dbo].[Document_Categories] ([doc_id]);
CREATE INDEX [IX_Document_Keywords_doc_id] ON [dbo].[Document_Keywords] ([doc_id]);
CREATE INDEX [IX_Document_Associations_doc_id] ON [dbo].[Document_Associations] ([doc_id]);
CREATE INDEX [IX_Document_Associations_association] ON [dbo].[Document_Associations] ([association_id]);

-- =============================================
-- STORED PROCEDURES
-- =============================================

USE [reports];
GO

-- Stored procedure for opening HCI keys
CREATE PROCEDURE [dbo].[OpenHCIKeys]
AS
BEGIN
    -- Implementation would depend on specific business logic
    -- This is a placeholder based on the SQL usage in the code
    SELECT 1;
END;
GO

-- =============================================
-- SAMPLE DATA INSERTION (Optional)
-- =============================================

-- Insert sample categories for document management
USE [DMS];
GO

INSERT INTO [dbo].[Document_Categories] ([doc_id], [cat_id], [created_by], [last_upd_by], [pr_flag])
VALUES 
(1, 118, 1, 1, 'N'), -- Images category
(1, 1, 1, 1, 'Y');   -- Default category

-- =============================================
-- SECURITY AND PERMISSIONS
-- =============================================

-- Create database users (adjust as needed for your environment)
USE [siebeldb];
GO

CREATE USER [SIEBEL] FOR LOGIN [SIEBEL];
ALTER ROLE [db_datareader] ADD MEMBER [SIEBEL];
ALTER ROLE [db_datawriter] ADD MEMBER [SIEBEL];

USE [DMS];
GO

CREATE USER [DMS] FOR LOGIN [DMS];
ALTER ROLE [db_datareader] ADD MEMBER [DMS];
ALTER ROLE [db_datawriter] ADD MEMBER [DMS];

USE [reports];
GO

CREATE USER [SIEBEL] FOR LOGIN [SIEBEL];
ALTER ROLE [db_datareader] ADD MEMBER [SIEBEL];
ALTER ROLE [db_datawriter] ADD MEMBER [SIEBEL];

USE [scanner];
GO

CREATE USER [SIEBEL] FOR LOGIN [SIEBEL];
ALTER ROLE [db_datareader] ADD MEMBER [SIEBEL];
ALTER ROLE [db_datawriter] ADD MEMBER [SIEBEL];

PRINT 'Database schema creation completed successfully.';
