-- SQL Server DDL Script for Data Cleaning Service Database Schema
-- Generated from ASP.NET VB Web Services Data Cleaning Library
-- This script creates the database tables used by the fuzzy-matching and weighted data-cleaning library

-- =============================================
-- Database: siebeldb
-- =============================================

USE [siebeldb]
GO

-- =============================================
-- Table: S_CONTACT
-- Description: Main contact records table
-- =============================================
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[S_CONTACT]') AND type in (N'U'))
BEGIN
    CREATE TABLE [dbo].[S_CONTACT](
        [ROW_ID] [nvarchar](15) NOT NULL,
        [CREATED] [datetime] NOT NULL,
        [CREATED_BY] [nvarchar](15) NOT NULL,
        [LAST_UPD] [datetime] NOT NULL,
        [LAST_UPD_BY] [nvarchar](15) NOT NULL,
        [MODIFICATION_NUM] [int] NOT NULL,
        [CONFLICT_ID] [int] NOT NULL,
        [BU_ID] [nvarchar](15) NOT NULL,
        [FST_NAME] [nvarchar](50) NULL,
        [LAST_NAME] [nvarchar](50) NULL,
        [MID_NAME] [nvarchar](50) NULL,
        [SEX_MF] [nvarchar](1) NULL,
        [X_MATCH_CD] [nvarchar](50) NULL,
        [X_MATCH_DT] [datetime] NULL,
        [LOGIN] [nvarchar](50) NULL,
        CONSTRAINT [PK_S_CONTACT] PRIMARY KEY CLUSTERED ([ROW_ID] ASC)
    )
END
GO

-- =============================================
-- Table: S_CONTACT_X
-- Description: Contact extension table
-- =============================================
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[S_CONTACT_X]') AND type in (N'U'))
BEGIN
    CREATE TABLE [dbo].[S_CONTACT_X](
        [ROW_ID] [nvarchar](15) NOT NULL,
        [CREATED] [datetime] NOT NULL,
        [CREATED_BY] [nvarchar](15) NOT NULL,
        [LAST_UPD] [datetime] NOT NULL,
        [LAST_UPD_BY] [nvarchar](15) NOT NULL,
        [DCKING_NUM] [int] NOT NULL,
        [MODIFICATION_NUM] [int] NOT NULL,
        [CONFLICT_ID] [int] NOT NULL,
        [PAR_ROW_ID] [nvarchar](15) NOT NULL,
        CONSTRAINT [PK_S_CONTACT_X] PRIMARY KEY CLUSTERED ([ROW_ID] ASC)
    )
END
GO

-- =============================================
-- Table: S_POSTN_CON
-- Description: Contact position records
-- =============================================
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[S_POSTN_CON]') AND type in (N'U'))
BEGIN
    CREATE TABLE [dbo].[S_POSTN_CON](
        [ROW_ID] [nvarchar](15) NOT NULL,
        [CREATED] [datetime] NOT NULL,
        [CREATED_BY] [nvarchar](15) NOT NULL,
        [LAST_UPD] [datetime] NOT NULL,
        [LAST_UPD_BY] [nvarchar](15) NOT NULL,
        [MODIFICATION_NUM] [int] NOT NULL,
        [CONFLICT_ID] [int] NOT NULL,
        [CON_FST_NAME] [nvarchar](50) NULL,
        [CON_ID] [nvarchar](15) NOT NULL,
        [CON_LAST_NAME] [nvarchar](50) NULL,
        [POSTN_ID] [nvarchar](15) NOT NULL,
        [ROW_STATUS] [nvarchar](1) NULL,
        [ASGN_DNRM_FLG] [nvarchar](1) NULL,
        [ASGN_MANL_FLG] [nvarchar](1) NULL,
        [ASGN_SYS_FLG] [nvarchar](1) NULL,
        [STATUS] [nvarchar](20) NULL,
        CONSTRAINT [PK_S_POSTN_CON] PRIMARY KEY CLUSTERED ([ROW_ID] ASC)
    )
END
GO

-- =============================================
-- Table: S_ORG_EXT
-- Description: Organization/Account records
-- =============================================
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[S_ORG_EXT]') AND type in (N'U'))
BEGIN
    CREATE TABLE [dbo].[S_ORG_EXT](
        [ROW_ID] [nvarchar](15) NOT NULL,
        [CREATED] [datetime] NOT NULL,
        [CREATED_BY] [nvarchar](15) NOT NULL,
        [LAST_UPD] [datetime] NOT NULL,
        [LAST_UPD_BY] [nvarchar](15) NOT NULL,
        [DCKING_NUM] [int] NOT NULL,
        [MODIFICATION_NUM] [int] NOT NULL,
        [CONFLICT_ID] [int] NOT NULL,
        [BU_ID] [nvarchar](15) NOT NULL,
        [DISA_CLEANSE_FLG] [nvarchar](1) NULL,
        [NAME] [nvarchar](100) NULL,
        [LOC] [nvarchar](100) NULL,
        [PROSPECT_FLG] [nvarchar](1) NULL,
        [PRTNR_FLG] [nvarchar](1) NULL,
        [ENTERPRISE_FLAG] [nvarchar](1) NULL,
        [LANG_ID] [nvarchar](15) NULL,
        [BASE_CURCY_CD] [nvarchar](3) NULL,
        [CREATOR_LOGIN] [nvarchar](50) NULL,
        [CUST_STAT_CD] [nvarchar](15) NULL,
        [DESC_TEXT] [nvarchar](255) NULL,
        [DISA_ALL_MAILS_FLG] [nvarchar](1) NULL,
        [FRGHT_TERMS_CD] [nvarchar](15) NULL,
        [MAIN_FAX_PH_NUM] [nvarchar](20) NULL,
        [MAIN_PH_NUM] [nvarchar](20) NULL,
        [DEDUP_TOKEN] [nvarchar](50) NULL,
        [X_MATCH_DT] [datetime] NULL,
        CONSTRAINT [PK_S_ORG_EXT] PRIMARY KEY CLUSTERED ([ROW_ID] ASC)
    )
END
GO

-- =============================================
-- Table: S_ACCNT_POSTN
-- Description: Account position records
-- =============================================
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[S_ACCNT_POSTN]') AND type in (N'U'))
BEGIN
    CREATE TABLE [dbo].[S_ACCNT_POSTN](
        [ROW_ID] [nvarchar](15) NOT NULL,
        [CREATED] [datetime] NOT NULL,
        [CREATED_BY] [nvarchar](15) NOT NULL,
        [LAST_UPD] [datetime] NOT NULL,
        [LAST_UPD_BY] [nvarchar](15) NOT NULL,
        [MODIFICATION_NUM] [int] NOT NULL,
        [CONFLICT_ID] [int] NOT NULL,
        [ACCNT_NAME] [nvarchar](100) NULL,
        [OU_EXT_ID] [nvarchar](15) NOT NULL,
        [POSITION_ID] [nvarchar](15) NOT NULL,
        [ROW_STATUS] [nvarchar](1) NULL,
        CONSTRAINT [PK_S_ACCNT_POSTN] PRIMARY KEY CLUSTERED ([ROW_ID] ASC)
    )
END
GO

-- =============================================
-- Table: S_ADDR_PER
-- Description: Personal address records
-- =============================================
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[S_ADDR_PER]') AND type in (N'U'))
BEGIN
    CREATE TABLE [dbo].[S_ADDR_PER](
        [ROW_ID] [nvarchar](15) NOT NULL,
        [CREATED] [datetime] NOT NULL,
        [CREATED_BY] [nvarchar](15) NOT NULL,
        [LAST_UPD] [datetime] NOT NULL,
        [LAST_UPD_BY] [nvarchar](15) NOT NULL,
        [DCKING_NUM] [int] NOT NULL,
        [MODIFICATION_NUM] [int] NOT NULL,
        [CONFLICT_ID] [int] NOT NULL,
        [DISA_CLEANSE_FLG] [nvarchar](1) NULL,
        [PER_ID] [nvarchar](15) NOT NULL,
        [ADDR] [nvarchar](100) NULL,
        [CITY] [nvarchar](50) NULL,
        [COMMENTS] [nvarchar](255) NULL,
        [COUNTY] [nvarchar](50) NULL,
        [COUNTRY] [nvarchar](3) NULL,
        [STATE] [nvarchar](20) NULL,
        [ZIPCODE] [nvarchar](20) NULL,
        [X_MATCH_CD] [nvarchar](50) NULL,
        [X_MATCH_DT] [datetime] NULL,
        [X_LAT] [decimal](10, 6) NULL,
        [X_LONG] [decimal](10, 6) NULL,
        [X_CASS_CHECKED] [datetime] NULL,
        [X_CASS_CODE] [nvarchar](10) NULL,
        [JURIS_ID] [nvarchar](15) NULL,
        CONSTRAINT [PK_S_ADDR_PER] PRIMARY KEY CLUSTERED ([ROW_ID] ASC)
    )
END
GO

-- =============================================
-- Table: S_ADDR_PER_X
-- Description: Personal address extension records
-- =============================================
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[S_ADDR_PER_X]') AND type in (N'U'))
BEGIN
    CREATE TABLE [dbo].[S_ADDR_PER_X](
        [ROW_ID] [nvarchar](15) NOT NULL,
        [CREATED] [datetime] NOT NULL,
        [CREATED_BY] [nvarchar](15) NOT NULL,
        [LAST_UPD] [datetime] NOT NULL,
        [LAST_UPD_BY] [nvarchar](15) NOT NULL,
        [MODIFICATION_NUM] [int] NOT NULL,
        [CONFLICT_ID] [int] NOT NULL,
        [PAR_ROW_ID] [nvarchar](15) NOT NULL,
        [ATTRIB_03] [nvarchar](50) NULL,
        [ATTRIB_04] [nvarchar](50) NULL,
        [ATTRIB_34] [nvarchar](50) NULL,
        CONSTRAINT [PK_S_ADDR_PER_X] PRIMARY KEY CLUSTERED ([ROW_ID] ASC)
    )
END
GO

-- =============================================
-- Table: S_ADDR_ORG
-- Description: Organization address records
-- =============================================
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[S_ADDR_ORG]') AND type in (N'U'))
BEGIN
    CREATE TABLE [dbo].[S_ADDR_ORG](
        [ROW_ID] [nvarchar](15) NOT NULL,
        [CREATED] [datetime] NOT NULL,
        [LAST_UPD] [datetime] NOT NULL,
        [LAST_UPD_BY] [nvarchar](15) NOT NULL,
        [DCKING_NUM] [int] NOT NULL,
        [MODIFICATION_NUM] [int] NOT NULL,
        [CONFLICT_ID] [int] NOT NULL,
        [DISA_CLEANSE_FLG] [nvarchar](1) NULL,
        [OU_ID] [nvarchar](15) NOT NULL,
        [ADDR] [nvarchar](100) NULL,
        [CITY] [nvarchar](50) NULL,
        [COMMENTS] [nvarchar](255) NULL,
        [COUNTY] [nvarchar](50) NULL,
        [COUNTRY] [nvarchar](3) NULL,
        [STATE] [nvarchar](20) NULL,
        [ZIPCODE] [nvarchar](20) NULL,
        [X_MATCH_CD] [nvarchar](50) NULL,
        [X_MATCH_DT] [datetime] NULL,
        [X_LAT] [decimal](10, 6) NULL,
        [X_LONG] [decimal](10, 6) NULL,
        [X_CASS_CHECKED] [datetime] NULL,
        [X_CASS_CODE] [nvarchar](10) NULL,
        [JURIS_ID] [nvarchar](15) NULL,
        CONSTRAINT [PK_S_ADDR_ORG] PRIMARY KEY CLUSTERED ([ROW_ID] ASC)
    )
END
GO

-- =============================================
-- Table: S_ADDR_ORG_X
-- Description: Organization address extension records
-- =============================================
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[S_ADDR_ORG_X]') AND type in (N'U'))
BEGIN
    CREATE TABLE [dbo].[S_ADDR_ORG_X](
        [ROW_ID] [nvarchar](15) NOT NULL,
        [CREATED] [datetime] NOT NULL,
        [CREATED_BY] [nvarchar](15) NOT NULL,
        [LAST_UPD] [datetime] NOT NULL,
        [LAST_UPD_BY] [nvarchar](15) NOT NULL,
        [MODIFICATION_NUM] [int] NOT NULL,
        [CONFLICT_ID] [int] NOT NULL,
        [PAR_ROW_ID] [nvarchar](15) NOT NULL,
        [ATTRIB_03] [nvarchar](50) NULL,
        [ATTRIB_04] [nvarchar](50) NULL,
        [ATTRIB_34] [nvarchar](50) NULL,
        CONSTRAINT [PK_S_ADDR_ORG_X] PRIMARY KEY CLUSTERED ([ROW_ID] ASC)
    )
END
GO

-- =============================================
-- Table: S_LST_OF_VAL
-- Description: List of values/lookup table
-- =============================================
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[S_LST_OF_VAL]') AND type in (N'U'))
BEGIN
    CREATE TABLE [dbo].[S_LST_OF_VAL](
        [ROW_ID] [nvarchar](15) NOT NULL,
        [TYPE] [nvarchar](30) NOT NULL,
        [CODE] [nvarchar](30) NOT NULL,
        [VAL] [nvarchar](100) NOT NULL,
        [ACTIVE_FLG] [nvarchar](1) NULL,
        [CREATED] [datetime] NULL,
        [CREATED_BY] [nvarchar](15) NULL,
        [LAST_UPD] [datetime] NULL,
        [LAST_UPD_BY] [nvarchar](15) NULL,
        CONSTRAINT [PK_S_LST_OF_VAL] PRIMARY KEY CLUSTERED ([ROW_ID] ASC)
    )
END
GO

-- =============================================
-- Indexes for Performance
-- =============================================

-- Index on S_CONTACT for name matching
IF NOT EXISTS (SELECT * FROM sys.indexes WHERE object_id = OBJECT_ID(N'[dbo].[S_CONTACT]') AND name = N'IX_S_CONTACT_NAMES')
BEGIN
    CREATE NONCLUSTERED INDEX [IX_S_CONTACT_NAMES] ON [dbo].[S_CONTACT]
    (
        [FST_NAME] ASC,
        [LAST_NAME] ASC,
        [MID_NAME] ASC
    )
END
GO

-- Index on S_CONTACT for match code
IF NOT EXISTS (SELECT * FROM sys.indexes WHERE object_id = OBJECT_ID(N'[dbo].[S_CONTACT]') AND name = N'IX_S_CONTACT_MATCH_CD')
BEGIN
    CREATE NONCLUSTERED INDEX [IX_S_CONTACT_MATCH_CD] ON [dbo].[S_CONTACT]
    (
        [X_MATCH_CD] ASC
    )
END
GO

-- Index on S_ORG_EXT for name matching
IF NOT EXISTS (SELECT * FROM sys.indexes WHERE object_id = OBJECT_ID(N'[dbo].[S_ORG_EXT]') AND name = N'IX_S_ORG_EXT_NAME')
BEGIN
    CREATE NONCLUSTERED INDEX [IX_S_ORG_EXT_NAME] ON [dbo].[S_ORG_EXT]
    (
        [NAME] ASC
    )
END
GO

-- Index on S_ORG_EXT for match code
IF NOT EXISTS (SELECT * FROM sys.indexes WHERE object_id = OBJECT_ID(N'[dbo].[S_ORG_EXT]') AND name = N'IX_S_ORG_EXT_DEDUP_TOKEN')
BEGIN
    CREATE NONCLUSTERED INDEX [IX_S_ORG_EXT_DEDUP_TOKEN] ON [dbo].[S_ORG_EXT]
    (
        [DEDUP_TOKEN] ASC
    )
END
GO

-- Index on S_ADDR_PER for address matching
IF NOT EXISTS (SELECT * FROM sys.indexes WHERE object_id = OBJECT_ID(N'[dbo].[S_ADDR_PER]') AND name = N'IX_S_ADDR_PER_ADDRESS')
BEGIN
    CREATE NONCLUSTERED INDEX [IX_S_ADDR_PER_ADDRESS] ON [dbo].[S_ADDR_PER]
    (
        [ADDR] ASC,
        [CITY] ASC,
        [STATE] ASC,
        [ZIPCODE] ASC
    )
END
GO

-- Index on S_ADDR_ORG for address matching
IF NOT EXISTS (SELECT * FROM sys.indexes WHERE object_id = OBJECT_ID(N'[dbo].[S_ADDR_ORG]') AND name = N'IX_S_ADDR_ORG_ADDRESS')
BEGIN
    CREATE NONCLUSTERED INDEX [IX_S_ADDR_ORG_ADDRESS] ON [dbo].[S_ADDR_ORG]
    (
        [ADDR] ASC,
        [CITY] ASC,
        [STATE] ASC,
        [ZIPCODE] ASC
    )
END
GO

-- Index on S_LST_OF_VAL for lookups
IF NOT EXISTS (SELECT * FROM sys.indexes WHERE object_id = OBJECT_ID(N'[dbo].[S_LST_OF_VAL]') AND name = N'IX_S_LST_OF_VAL_TYPE_CODE')
BEGIN
    CREATE NONCLUSTERED INDEX [IX_S_LST_OF_VAL_TYPE_CODE] ON [dbo].[S_LST_OF_VAL]
    (
        [TYPE] ASC,
        [CODE] ASC
    )
END
GO

-- =============================================
-- Sample Data for S_LST_OF_VAL (Country Codes)
-- =============================================
IF NOT EXISTS (SELECT * FROM [dbo].[S_LST_OF_VAL] WHERE TYPE = 'COUNTRY_CODE' AND CODE = 'US')
BEGIN
    INSERT INTO [dbo].[S_LST_OF_VAL] ([ROW_ID], [TYPE], [CODE], [VAL], [ACTIVE_FLG], [CREATED], [CREATED_BY], [LAST_UPD], [LAST_UPD_BY])
    VALUES 
    ('0-US', 'COUNTRY_CODE', 'US', 'USA', 'Y', GETDATE(), 'SYSTEM', GETDATE(), 'SYSTEM'),
    ('0-CA', 'COUNTRY_CODE', 'CA', 'CAN', 'Y', GETDATE(), 'SYSTEM', GETDATE(), 'SYSTEM'),
    ('0-MX', 'COUNTRY_CODE', 'MX', 'MEX', 'Y', GETDATE(), 'SYSTEM', GETDATE(), 'SYSTEM'),
    ('0-GB', 'COUNTRY_CODE', 'GB', 'GBR', 'Y', GETDATE(), 'SYSTEM', GETDATE(), 'SYSTEM'),
    ('0-DE', 'COUNTRY_CODE', 'DE', 'DEU', 'Y', GETDATE(), 'SYSTEM', GETDATE(), 'SYSTEM'),
    ('0-FR', 'COUNTRY_CODE', 'FR', 'FRA', 'Y', GETDATE(), 'SYSTEM', GETDATE(), 'SYSTEM'),
    ('0-JP', 'COUNTRY_CODE', 'JP', 'JPN', 'Y', GETDATE(), 'SYSTEM', GETDATE(), 'SYSTEM'),
    ('0-AU', 'COUNTRY_CODE', 'AU', 'AUS', 'Y', GETDATE(), 'SYSTEM', GETDATE(), 'SYSTEM')
END
GO

-- =============================================
-- Foreign Key Constraints
-- =============================================

-- S_CONTACT_X references S_CONTACT
IF NOT EXISTS (SELECT * FROM sys.foreign_keys WHERE object_id = OBJECT_ID(N'[dbo].[FK_S_CONTACT_X_S_CONTACT]') AND parent_object_id = OBJECT_ID(N'[dbo].[S_CONTACT_X]'))
BEGIN
    ALTER TABLE [dbo].[S_CONTACT_X] 
    ADD CONSTRAINT [FK_S_CONTACT_X_S_CONTACT] 
    FOREIGN KEY([PAR_ROW_ID]) REFERENCES [dbo].[S_CONTACT] ([ROW_ID])
END
GO

-- S_POSTN_CON references S_CONTACT
IF NOT EXISTS (SELECT * FROM sys.foreign_keys WHERE object_id = OBJECT_ID(N'[dbo].[FK_S_POSTN_CON_S_CONTACT]') AND parent_object_id = OBJECT_ID(N'[dbo].[S_POSTN_CON]'))
BEGIN
    ALTER TABLE [dbo].[S_POSTN_CON] 
    ADD CONSTRAINT [FK_S_POSTN_CON_S_CONTACT] 
    FOREIGN KEY([CON_ID]) REFERENCES [dbo].[S_CONTACT] ([ROW_ID])
END
GO

-- S_ACCNT_POSTN references S_ORG_EXT
IF NOT EXISTS (SELECT * FROM sys.foreign_keys WHERE object_id = OBJECT_ID(N'[dbo].[FK_S_ACCNT_POSTN_S_ORG_EXT]') AND parent_object_id = OBJECT_ID(N'[dbo].[S_ACCNT_POSTN]'))
BEGIN
    ALTER TABLE [dbo].[S_ACCNT_POSTN] 
    ADD CONSTRAINT [FK_S_ACCNT_POSTN_S_ORG_EXT] 
    FOREIGN KEY([OU_EXT_ID]) REFERENCES [dbo].[S_ORG_EXT] ([ROW_ID])
END
GO

-- S_ADDR_PER_X references S_ADDR_PER
IF NOT EXISTS (SELECT * FROM sys.foreign_keys WHERE object_id = OBJECT_ID(N'[dbo].[FK_S_ADDR_PER_X_S_ADDR_PER]') AND parent_object_id = OBJECT_ID(N'[dbo].[S_ADDR_PER_X]'))
BEGIN
    ALTER TABLE [dbo].[S_ADDR_PER_X] 
    ADD CONSTRAINT [FK_S_ADDR_PER_X_S_ADDR_PER] 
    FOREIGN KEY([PAR_ROW_ID]) REFERENCES [dbo].[S_ADDR_PER] ([ROW_ID])
END
GO

-- S_ADDR_ORG_X references S_ADDR_ORG
IF NOT EXISTS (SELECT * FROM sys.foreign_keys WHERE object_id = OBJECT_ID(N'[dbo].[FK_S_ADDR_ORG_X_S_ADDR_ORG]') AND parent_object_id = OBJECT_ID(N'[dbo].[S_ADDR_ORG_X]'))
BEGIN
    ALTER TABLE [dbo].[S_ADDR_ORG_X] 
    ADD CONSTRAINT [FK_S_ADDR_ORG_X_S_ADDR_ORG] 
    FOREIGN KEY([PAR_ROW_ID]) REFERENCES [dbo].[S_ADDR_ORG] ([ROW_ID])
END
GO

PRINT 'Database schema creation completed successfully.'
PRINT 'Tables created: S_CONTACT, S_CONTACT_X, S_POSTN_CON, S_ORG_EXT, S_ACCNT_POSTN, S_ADDR_PER, S_ADDR_PER_X, S_ADDR_ORG, S_ADDR_ORG_X, S_LST_OF_VAL'
PRINT 'Indexes and foreign key constraints have been applied.'
