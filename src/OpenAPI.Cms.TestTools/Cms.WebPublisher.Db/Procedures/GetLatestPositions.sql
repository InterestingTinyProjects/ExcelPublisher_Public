
-- =============================================
-- Author:		<Author,,Name>
-- Create date: <Create Date,,>
-- Description:	<Description,,>
-- =============================================
CREATE PROCEDURE dbo.GetLatestPositions
	@sheetName NVARCHAR(150)
AS
BEGIN

SELECT TOP 1 [publishTime],
		r.[sheetName],
		[rows],
		[columns],
		[data]
FROM [dbo].[Record] r
WHERE sheetName = @sheetName
ORDER BY publishTime DESC

END
GO