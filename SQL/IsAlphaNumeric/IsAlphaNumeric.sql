--**************************************
-- Name: IsAlphaNumeric
-- Description:Validates if the string provided contains only letters and numbers.
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

CREATE FUNCTION [dbo].[IsAlphaNumeric] 
(
	@Text varchar(255)
)
RETURNS BIT 
AS 


    BEGIN
    	DECLARE @Letters VARCHAR(255)
    	DECLARE @Index INT
    	DECLARE @Max INT
    	DECLARE @Pass BIT
    	SET @Letters = ' abcdefghijklmnopqrstuvwxyz01234567890'
    	SET @Max = LEN(@Text)
    	SET @Index = 0
    	SET @Pass = 1
    	WHILE @Index < @Max AND @Pass = 1
    	BEGIN
    		SET @Index = @Index + 1
    		IF NOT @Letters LIKE '%' + SUBSTRING(@Text, @Index, 1) + '%' SET @Pass = 0
    	END
    	RETURN(@Pass)
END