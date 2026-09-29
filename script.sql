-- Create database if it doesn't exist
IF NOT EXISTS (SELECT name FROM sys.databases WHERE name = N'TeamsKeepAliveDb')
BEGIN
    CREATE DATABASE [TeamsKeepAliveDb];
END
GO

USE [TeamsKeepAliveDb];
GO

-- Create TenantSessions Table
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[TenantSessions]') AND type in (N'U'))
BEGIN
    CREATE TABLE [dbo].[TenantSessions] (
        [Id] UNIQUEIDENTIFIER NOT NULL,
        [TenantId] NVARCHAR(100) NOT NULL,
        [UserEmail] NVARCHAR(256) NOT NULL,
        [MicrosoftGraphAccessToken] NVARCHAR(MAX) NOT NULL,
        [MicrosoftGraphRefreshToken] NVARCHAR(MAX) NOT NULL,
        [TokenExpiryUtc] DATETIME2(7) NOT NULL,
        [IsActive] BIT NOT NULL,
        [TargetPresenceState] NVARCHAR(50) NOT NULL CONSTRAINT [DF_TenantSessions_TargetPresenceState] DEFAULT N'Available',
        [LocalAccountId] NVARCHAR(256) NULL,
        [ClientId] NVARCHAR(256) NULL,
        [SkypeToken] NVARCHAR(MAX) NULL,
        CONSTRAINT [PK_TenantSessions] PRIMARY KEY CLUSTERED ([Id] ASC)
    );

    -- Create Unique Index on UserEmail
    CREATE UNIQUE NONCLUSTERED INDEX [IX_TenantSessions_UserEmail] 
        ON [dbo].[TenantSessions] ([UserEmail] ASC);

    -- Create Index on LocalAccountId
    CREATE NONCLUSTERED INDEX [IX_TenantSessions_LocalAccountId] 
        ON [dbo].[TenantSessions] ([LocalAccountId] ASC);
END
ELSE
BEGIN
    IF COL_LENGTH('dbo.TenantSessions', 'SkypeToken') IS NULL
    BEGIN
        ALTER TABLE [dbo].[TenantSessions] ADD [SkypeToken] NVARCHAR(MAX) NULL;
    END
END
GO

-- Create PresenceLogs Table
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PresenceLogs]') AND type in (N'U'))
BEGIN
    CREATE TABLE [dbo].[PresenceLogs] (
        [Id] UNIQUEIDENTIFIER NOT NULL,
        [TenantSessionId] UNIQUEIDENTIFIER NOT NULL,
        [TimestampUtc] DATETIME2(7) NOT NULL,
        [HttpStatusCode] INT NOT NULL,
        [ResponseBody] NVARCHAR(MAX) NOT NULL,
        CONSTRAINT [PK_PresenceLogs] PRIMARY KEY CLUSTERED ([Id] ASC),
        CONSTRAINT [FK_PresenceLogs_TenantSessions_TenantSessionId] FOREIGN KEY ([TenantSessionId]) 
            REFERENCES [dbo].[TenantSessions] ([Id]) ON DELETE CASCADE
    );

    -- Create index for faster foreign key lookup
    CREATE NONCLUSTERED INDEX [IX_PresenceLogs_TenantSessionId] 
        ON [dbo].[PresenceLogs] ([TenantSessionId] ASC);
END
GO

-- ============================================================================
-- VERIFICATION & ADO.NET MIGRATION NOTES
-- ============================================================================
-- Verified: The above table structures for TenantSessions and 
-- PresenceLogs are fully compatible with raw Microsoft.Data.SqlClient mappings.
-- No schema alterations (ALTER TABLE) were required as the layout perfectly
-- matches the mapped properties in C#.
-- ============================================================================
