--**************************************
-- Name: sprSearchTable
-- Description:This procedure searches all columns within the table specified. All columns are casted as varchar(1000). If any of the columns contain the query, then all columns are returned. (SQL2K)
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

Create Procedure dbo.SearchTable
(
	@TableName VARCHAR(128),	-- Table to be searched
	@Query VARCHAR(128)			-- Query to search against
)
As
	/*
	This procedure searches all columns within the table specified.
	All columns are casted as varchar(1000)
	If any of the columns contain the query, then all columns are returned.
	*/
	DECLARE @Column	VARCHAR(128)	-- Name of column to search
	DECLARE @SQL	VARCHAR(1000)	-- SQL Query built to query against table columns
	DECLARE @Count	TINYINT			-- Loop counter
	
	-- Default SQL to empty string
	SET @SQL = ''
	
	-- Default loop counter as the first pass
	SET @Count = 1
	
	-- Prevent SQL Injection Attack
	SET @Query = Replace(@Query, '''', '''''')
	
	-- Verify that the table exists
	IF NOT EXISTS(SELECT 0 FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = @TableName)
	BEGIN
		-- Raise an error if the table is invalid
		RaisError('Table does not exist.', 16, 1)
		
		-- Exit procedure
		RETURN
	END
	
	-- Query schema for all column names within table
	DECLARE Columns CURSOR FOR
	SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = @TableName
	-- Open the cursor
	OPEN Columns
	
	-- Grab first column name
	FETCH NEXT FROM Columns INTO @Column
	
	-- Process columns while the cursor is not at the end
	WHILE @@FETCH_STATUS = 0
	BEGIN
		-- If we are not on the first pass, append an "OR" operator
		IF NOT @Count = 1 SET @SQL = @SQL + ' OR'
		
		-- Append SQL to cast column as a varchar and compare against the query
		SET @SQL = @SQL + ' CAST(' + @Column + ' AS VARCHAR(1000)) LIKE ''%' + @Query + '%'''
		
		-- Grab the next column name
		FETCH NEXT FROM Columns INTO @Column
		-- Incriment the counter
		SET @Count = @Count + 1
	END
	
	-- Close the cursor
	CLOSE Columns
	
	-- Remove cursor from memory
	DEALLOCATE Columns
	
	/*
	Execute the compiled query against the specified table to
	return all rows.
	*/
	EXEC('SELECT * FROM [' + @TableName + '] WHERE' + @SQL)
	
	return