--**************************************
-- Name: Batch Change Object Owner
-- Description:demonstrates how to use a cursor to change all db objects to DBO as the owner.
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

-- change object owner of stored procedures to DBO
DECLARE @Name as sysname
DECLARE [Procedures] CURSOR FOR
	SELECT [Name] FROM sysobjects
	WHERE [Type] = 'P' AND [Category] = 0 AND NOT [UID] = 1
OPEN [Procedures]
FETCH NEXT FROM [Procedures] INTO @Name
WHILE @@FETCH_STATUS = 0
	BEGIN
		EXECUTE sp_ChangeObjectOwner @Name, dbo
		FETCH NEXT FROM [Procedures] INTO @Name
	END
CLOSE [Procedures]
DEALLOCATE [Procedures]