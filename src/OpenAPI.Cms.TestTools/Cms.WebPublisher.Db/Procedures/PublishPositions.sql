
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
	@maxStoredRecords BIGINT = NULL,
	@data VARBINARY(MAX)
AS
BEGIN

	DECLARE @count INT;

-- remove redundent rows
	IF @maxStoredRecords IS NOT NULL 
		AND @maxStoredRecords > 0
	BEGIN
		SELECT @count = COUNT(row_id)
		FROM [dbo].[Record]
		WHERE [sheetName] = @sheetName

		IF @count > @maxStoredRecords
		BEGIN
			DELETE [dbo].[Record]
			WHERE [row_id] IN 
			(
				SELECT TOP(@count - @maxStoredRecords + 1) row_id
				FROM [dbo].[Record]
				WHERE [sheetName] = @sheetName
				ORDER BY [publishTime] ASC
			)
		END
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

