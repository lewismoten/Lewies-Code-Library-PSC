--**************************************
-- Name: Log Record Changes
-- Description:Make a copy of records within a table each time it is updated or a new record is created. This is very basic stuff. I'll be working on an update to handle primary keys and such.
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

/*
-- Basics for setting up a log file.
-- Replace the [MyTable] with the table you wish to log
*/
-- Copy Schema
SELECT * INTO [dbo].[Log_MyTable] FROM [MyTable] 
WHERE 1=2
GO
-- Create trigger
CREATE Trigger dbo.[MyTable_log] ON dbo.[MyTable]
AFTER INSERT, UPDATE AS
INSERT INTO Log_MyTable SELECT * FROM INSERTED
GO