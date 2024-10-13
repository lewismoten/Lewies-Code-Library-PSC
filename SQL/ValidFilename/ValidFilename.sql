--**************************************
-- Name: ValidFilename
-- Description:Validates if the text proveded is a valid file name for windows. Useful in T-SQL statements validation before adding records.
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

CREATE FUNCTION [dbo].[validFilename] 
(
	@FileName varchar(255)
)
RETURNS BIT 
AS 


    BEGIN
    	IF @FileName LIKE '%/%' RETURN(0)
    	IF @FileName LIKE '%\%' RETURN(0)
    	IF @FileName LIKE '%:%' RETURN(0)
    	IF @FileName LIKE '%*%' RETURN(0)
    	IF @FileName LIKE '%?%' RETURN(0)
    	IF @FileName LIKE '%"%' RETURN(0)
    	IF @FileName LIKE '%<%' RETURN(0)
    	IF @FileName LIKE '%>%' RETURN(0)
    	IF @FileName LIKE '%|%' RETURN(0)
    	RETURN(1)
END