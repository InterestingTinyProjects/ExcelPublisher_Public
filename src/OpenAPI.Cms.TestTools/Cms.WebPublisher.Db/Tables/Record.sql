CREATE TABLE [dbo].[Record]
(
	[row_id] BIGINT NOT NULL PRIMARY KEY IDENTITY, 
    [sheetName] NVARCHAR(150) NULL, 
    [publishTime] DATETIME2 NULL, 
    [rows] INT NULL, 
    [columns] INT NULL, 
    [data] VARBINARY(MAX) NULL
)

GO

CREATE INDEX [IX_Record_CategoryPublishTime] ON [dbo].[Record] ([sheetName], [publishTime] DESC)


GO

CREATE INDEX [IX_SheetName] ON [dbo].[Record] ([sheetName])
