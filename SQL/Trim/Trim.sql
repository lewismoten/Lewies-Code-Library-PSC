--**************************************
-- Name: Trim
-- Description:Allows you to Trim your strings when querying tables. I'm tired of using the Trim function in my ASP pages and Visual Basic. I created this little feature to trim away within SQL Server itself. This is a User Defined function rather then a stored procedure, so you can use it rite inside your select statements. "SELECT Trim([FirstName]) FROM [Users]". Sometimes you may find that you need to prefix the function with the owner - "SELECT dbo.Trim([FirstName]) FROM [Users]". 
-- Within the SQL Server Enterprise manager you will need to expand the "Databases" node. Locate your database and expand the node. Next, find the User Defined Functions node. Right-Mouse-Click and choose "New User Defined Function ...". Paste my code into it.
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

CREATE FUNCTION [dbo].[Trim] (@String varchar(255)) 
RETURNS varchar(255) 
AS 


    BEGIN 
    	/*
    	When white space is trailing within a string, SQL Server does not seem to count it.
    	Whith this in mind, some of the logic here may begin to make sense.
    	*/
    	
    	-- Append string terminator (can be anything you want)
    	Select @String = @String + '.'
    	
    	-- remove prefixed white space 1 at a time
    	While Left(@String, 1) = ' ' SELECT @String = Right(@String, Len(@String) - 1)
    	
    	-- remove string terminator
    	Select @String = Left(@String, Len(@String) - 1)
    	
    	-- remove trailing whitespace
    	Select @String = Left(@String, Len(@String))
    	
    	Return(@String) 
END