--**************************************
-- Name: Month Range Table
-- Description:Build a table and populate it with dates between a specified range. I often use this method to create a set of months available, then use a left outer join to add any existing data within an actual table.
-- By: Lewis Moten
--
--
-- Inputs:None
--
-- Returns:None
--
--Assumes:None
--
--Side Effects:None
--This code is copyrighted and has limited warranties.
--**************************************

DECLARE @Counter INT
DECLARE @Start SMALLDATETIME
DECLARE @End SMALLDATETIME
DECLARE @Months INT
SET @Start = '5/28/1975'
SET @End = GetDate()
SET @Months = DateDIFF(m, @Start, @End)
SET @Counter = 0
DECLARE @Dates TABLE ([Date] datetime)
WHILE @Counter < @Months
	BEGIN
		INSERT INTO @Dates VALUES (DATEADD(m, @Counter, @Start))
		SET @Counter = @Counter + 1
	END
SELECT [Date] FROM @Dates
		
