--**************************************
-- Name: Int2GUID
-- Description:Datatype conversion to change integers to unique identifier datatypes. This
-- method is most useful when upgrading database schemas from using integers to
-- using guids. You will still be able to "read" values by eye to match old Integers with the new GUIDs.
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

CREATE FUNCTION [dbo].[Int2GUID] (@INT INT)
RETURNS UNIQUEIDENTIFIER AS 


    BEGIN 
    	DECLARE @GUID VARCHAR(36)
    	SET @GUID = CAST(@INT AS VARCHAR(12))
    	SET @GUID = RIGHT('000000000000' + @GUID, 12)
    	SET @GUID = '00000000-0000-0000-0000-' + @GUID
    	return(@GUID)
END