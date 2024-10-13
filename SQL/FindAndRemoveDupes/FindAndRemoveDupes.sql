--**************************************
-- Name: Find and remove Dupes
-- Description:Find and remove duplicate values within a table.
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

DECLARE @MyValue VARCHAR(50)
-- Define cursor
DECLARE Dupes CURSOR FOR
	SELECT [MyField] FROM [MyTable] GROUP BY [MyField]
	HAVING COUNT(0) > 1 ORDER BY [MyField] ASC
-- Open connection to records
OPEN Dupes
-- Grab first record
FETCH NEXT FROM Dupes INTO @MyValue
-- While we have a record
WHILE @@FETCH_STATUS = 0
	BEGIN
		-- Delete extra dupes
		/*
		We find 1st values id, then delete all duplicate
		values that do not have same ID
		*/
		DELETE FROM [MyTable] WHERE [MyRecordID] IN(SELECT [MyRecordID]
		FROM [MyTable] WHERE [MyField] = @MyValue AND [MyRecorID] NOT
		IN(SELECT TOP 1 [MyRecordID] FROM [MyTable] WHERE [MyField] =
		@MyValue))
		-- move to next record
		FETCH NEXT FROM Dupes INTO @MyValue
	END
-- close the connection
CLOSE Dupes
-- remove cursor
DEALLOCATE Dupes