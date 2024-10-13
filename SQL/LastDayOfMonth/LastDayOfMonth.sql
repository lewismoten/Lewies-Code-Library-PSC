--**************************************
-- Name: Last Day of month
-- Description:Returns last day, hour, minute, second of date passed to it. Example - print dbo.lastday('9/2/2002 3:57:31 PM') would return 'Sep 30 2002 11:59PM'. This function depends on dbo.FirstDay() located on PSC under First day of month
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

CREATE FUNCTION [dbo].[LastDay] (@Date smalldatetime) 
RETURNS smalldatetime
AS 


    BEGIN 
    RETURN( DATEADD(nn,-1,DATEADD(mm,1,DATEADD(d,-(DATEPART(dd,@Date) - 1), CONVERT(VARCHAR,@Date,101))) )
END