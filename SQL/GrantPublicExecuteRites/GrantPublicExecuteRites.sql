--**************************************
-- Name: Grant Public Execute Rites
-- Description:Grants execute rites on all stored procedures to the public. can be modified to grant to other users.
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

-- grant public to execute all procedures within
-- current database.
DECLARE @Name as sysname
DECLARE [Procedures] CURSOR FOR
	SELECT [Name] FROM sysobjects
	WHERE [Type] = 'P' AND [Category] = 0
OPEN [Procedures]
FETCH NEXT FROM [Procedures] INTO @Name
WHILE @@FETCH_STATUS = 0
	BEGIN
		EXECUTE('GRANT EXECUTE ON ' + @Name + ' TO PUBLIC')
		FETCH NEXT FROM [Procedures] INTO @Name
	END
CLOSE [Procedures]
DEALLOCATE [Procedures]