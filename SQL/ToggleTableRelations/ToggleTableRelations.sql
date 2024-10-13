--**************************************
-- Name: Toggle Table Relations
-- Description:Useful for disabling constraints during database replication. When copying data to another database, use this procedure on the target database first. When finnished, run it again to enable the constraints. I used this because I was working on a database with tables referencing themselves. I couldn't add records when previouse records did not exist. This solved the problem.
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

CREATE PROCEDURE dbo.sprCheckRelations
(
	@Check BIT = 1
)
AS
	DECLARE @Table VARCHAR(128)
	DECLARE @Constraint VARCHAR(128)
	DECLARE @SQL VARCHAR(1000)
	DECLARE @True BIT
	DECLARE @False BIT
	
	SET @True = 1
	SET @False = 0
	
	-- Grab a list of all relational constriaints in database
	DECLARE Constraints CURSOR FOR 
		SELECT TABLE_NAME, CONSTRAINT_NAME 
		FROM INFORMATION_SCHEMA.TABLE_CONSTRAINTS
	
	OPEN Constraints
	
	-- Loop through each constraint
	FETCH NEXT FROM Constraints INTO @Table, @Constraint
	
	WHILE @@FETCH_STATUS = 0
	BEGIN
		-- Determine if we are to check constraints
		IF @Check = @True
			SET @SQL = 'ALTER TABLE ' + @Table + ' CHECK CONSTRAINT ' + @Constraint
		ELSE
			SET @SQL = 'ALTER TABLE ' + @Table + ' NOCHECK CONSTRAINT ' + @Constraint
	
		-- Add/Remove constraint
		EXEC(@SQL)
		-- Move to next constraint
		FETCH NEXT FROM Constraints INTO @Table, @Constraint
	END
	
	CLOSE Constraints
	
	DEALLOCATE Constraints