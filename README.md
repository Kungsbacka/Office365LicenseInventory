# Office365LicenseInventory
Takes an inventory of all assigned licenses in Office 365.

## Dependencies
* AzureAD PowerShell module
* SQL Server database

## Database tables
```SQL
CREATE TABLE [dbo].[Office365AssignedLicense](
	[id] [int] IDENTITY(1,1) NOT NULL,
	[objectId] [uniqueidentifier] NOT NULL,
	[skuId] [uniqueidentifier] NOT NULL,
    CONSTRAINT [PK_Office365AssignedLicense] PRIMARY KEY CLUSTERED ([id])
)

CREATE TABLE [dbo].[Office365User](
	[objectId] [uniqueidentifier] NOT NULL,
	[userPrincipalName] [nvarchar](1024) NOT NULL,
    CONSTRAINT [PK_Office365User] PRIMARY KEY CLUSTERED ([objectId])
)
```
