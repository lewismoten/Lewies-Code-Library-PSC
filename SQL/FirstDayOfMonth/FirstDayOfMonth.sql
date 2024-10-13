--**************************************
-- Name: First day of month
-- Description:Pass a date to the function and it will return the first day of the month. Example - 
-- print dbo.firstday('9/2/2002 3:57:31 PM')
-- will print
--     'Sep 1 2002 12:00 AM'
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

CREATE FUNCTION [dbo].[FirstDay] (@Date smalldatetime) 
RETURNS smalldatetime
AS 


    BEGIN 
    	SET @Date = convert(varchar, @Date, 101)
    	SET @Date = DATEADD(d, -(DATEPART(dd, @Date) - 1), @Date)
    	Return(@Date)
END

		
