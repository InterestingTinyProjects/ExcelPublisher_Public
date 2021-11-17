
-- =============================================
-- Author:		<Author,,Name>
-- Create date: <Create Date,,>
-- Description:	<Description,,>
-- =============================================


-- =============================================
-- Author:		<Author,,Name>
-- Create date: <Create Date,,>
-- Description:	<Description,,>
-- =============================================
CREATE PROCEDURE [dbo].[GetLatestPositions]
	@sheetName NVARCHAR(150),
	@topCount INT = 1,
	@isPercentage BIT = 0
AS
BEGIN

IF @topCount < 0 
BEGIN 
	SET @topCount = 0
END

IF @isPercentage = 0 
BEGIN
	SELECT TOP (@topCount) [publishTime],
			r.[sheetName],
			[rows],
			[columns],
			[data]
	FROM [dbo].[Record] r
	WHERE sheetName = @sheetName
	ORDER BY publishTime DESC
END
ELSE
BEGIN
	SELECT TOP (@topCount) PERCENT [publishTime],
			r.[sheetName],
			[rows],
			[columns],
			[data]
	FROM [dbo].[Record] r
	WHERE sheetName = @sheetName
	ORDER BY publishTime DESC
END

END
GO



