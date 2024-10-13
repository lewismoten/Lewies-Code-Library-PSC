--**************************************
-- Name: Count Matches
-- Description:Allows you to count how many times a string appears within another string. example - count carriage returns within text. Handles up to 8K characters. Tricks can get around 8K limit for large text fields.
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
CREATE FUNCTION [dbo].[funCountMatch] (@Text varchar(8000), @Pattern varchar(20))
RETURNS INT
AS 


    BEGIN 
    	DECLARE @Matches INT
    	DECLARE @Index INT
    	DECLARE @MaxIndex INT
    	DECLARE @PatternLen INT
    	
    	SET @PatternLen = Len(@Pattern + '.') - 1
    	SET @MaxIndex = Len(@Text) - (@PatternLen - 1)
    	SET @Index = 1
    	SET @Matches = 0
    	WHILE @Index < @MaxIndex
    		BEGIN
    			IF SUBSTRING(@Text, @Index, @PatternLen) = @Pattern SET @Matches = @Matches + 1
    			SET @Index = @Index + 1
    		END
    	RETURN(@Matches)
    
END

		
