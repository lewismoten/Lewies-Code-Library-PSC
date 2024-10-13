--**************************************
-- Name: GrantUserExecuteAll
-- Description:Allows you to specify a username and grant access to execute all SQL stored procedures within the current database. Very useful for developers that create databases with one account, but have programs accessing them through a seperate user account. This is an update to an old piece of code I had on PSC that limited to the public account.
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

CREATE PROCEDURE dbo.GrantUserExecuteAll
(
	@UserName VARCHAR(64)
)
AS
	SET NOCOUNT ON
	
	DECLARE @Name SYSNAME
	
	DECLARE [Procedures] CURSOR FOR
		SELECT [Name] FROM sysobjects
		WHERE [Type] = 'P' AND [Category] = 0
		
	OPEN [Procedures]
	
	FETCH NEXT FROM [Procedures] INTO @Name
	WHILE @@FETCH_STATUS = 0
	
	BEGIN
		EXECUTE('GRANT EXECUTE ON ' + @Name + ' TO ' + @UserName)
		FETCH NEXT FROM [Procedures] INTO @Name
	END
	
	CLOSE [Procedures]
	DEALLOCATE [Procedures]
	RETURN