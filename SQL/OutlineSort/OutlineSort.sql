--**************************************
-- Name: OutlineSort '3.2.10.2'
-- Description:This function pads each number found within each
-- decimal to allow sorting on outline numbers. Fixes problems where 2 comes after 10 in string sorts. An
-- example of how this function is used would appear
-- like this:
-- SELECT OutlineNumber FROM Outline 
-- ORDER BY dbo.OutlineSort(OutlineNumber)
-- By: Lewis Moten
--
--
-- Inputs:String '1.2.3.1'
--
-- Returns:String '00001.00002.00003.00001'
--
--Assumes:Assumes numbers are not greater then or equal to 100000. Assumes there is a depth of 42 numbers or less.
--
--Side Effects:None
--This code is copyrighted and has limited warranties.
--**************************************

/*
These statements parse an outline number and print
out the depth, and index of the number. This can
be converted to a function to help with sorting
outlines.
This function pads each number found within each
decimal to allow sorting on outline numbers. An
example of how this function is used would appear
like this:
SELECT OutlineNumber FROM Outline 
ORDER BY dbo.OutlineSort(OutlineNumber)
For a quick demonstration of what is returned,
you may execute the following statement after
the function as been created in your database.
PRINT dbo.OutlineSort('1.4.32.27')
*/
CREATE FUNCTION [dbo].[OutlineSort] (@OutlineNumber VARCHAR(255))
RETURNS VARCHAR(255)
AS 


    BEGIN 
    	DECLARE @Position INT
    	DECLARE @ListIndex VARCHAR(30)
    	DECLARE @FinalNumber VARCHAR(255)
    	DECLARE @MaxDigits INT
    	
    	-- determine the depth of the last number
    	SET @Position = 1
    	SET @ListIndex = ''
    	SET @FinalNumber = ''
    	SET @MaxDigits = 5
    	
    	SET @Position = CHARINDEX('.', @OutlineNumber, 1)
    	WHILE NOT @Position = 0
    	BEGIN
    		SET @ListIndex = SUBSTRING(@OutlineNumber, 1, @Position - 1)
    		SET @OutlineNumber = SUBSTRING(@OutlineNumber, @Position + 1, LEN(@OutlineNumber))
    		SET @FinalNumber = @FinalNumber + REPLICATE('0', @MaxDigits - LEN(@ListIndex)) + @ListIndex + '.'
    		SET @Position = CHARINDEX('.', @OutlineNumber, 1)
    	END
    	SET @FinalNumber = @FinalNumber + REPLICATE('0', @MaxDigits - LEN(@OutlineNumber)) + @OutlineNumber
    	
    	Return(@FinalNumber)
END