--**************************************
-- Name: Month/Year function
-- Description:Returns a date formated as "Mmm YYYY" (Ex. Oct 2002). Not anything ground breaking ...
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

CREATE FUNCTION [dbo].[Mmm_YYYY] (@Date smalldatetime) 
RETURNS char(8)
AS 
    BEGIN 
    	DECLARE @ME CHAR(8)
    	SET @ME = Left(DateName(mm, @Date), 3) + ' ' + Cast(DatePart(yy, @Date) AS VARCHAR(4))
    	Return(@ME)
END