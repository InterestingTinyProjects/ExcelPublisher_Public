
-- =============================================
-- Author:		<Author,,Name>
-- Create date: <Create Date,,>
-- Description:	<Description,,>
-- =============================================
CREATE PROCEDURE dbo.PublishPositions
	@publishTime DATETIME2,
	@sheetName NVARCHAR(150),
	@rows INT,
	@columns INT,
	@data VARBINARY(MAX)
AS
BEGIN

DECLARE @count INT;

-- remove redundent rows
SELECT @count = COUNT(row_id)
FROM [dbo].[Record]

IF @count > 18000
BEGIN
	DELETE [dbo].[Record]
	WHERE [sheetName] = @sheetName
END


INSERT INTO [dbo].[Record]
(
	[sheetName],
	[publishTime],
	[rows],
	[columns],
	[data]
)
VALUES
(
	@sheetName,
	@publishTime,
	@rows,
	@columns,
	@data
)

-- return
SELECT @@ROWCOUNT

END
GO
