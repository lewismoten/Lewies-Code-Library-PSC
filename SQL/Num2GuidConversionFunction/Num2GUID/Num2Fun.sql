CREATE FUNCTION [dbo].[Num2Fun] (@Number BIGINT)
RETURNS VARCHAR(64) AS
BEGIN
	/*

	Converts a base-10 number to a funny base-14 number.
	
	This demonstration is for fun.  It shows you how you can also come up with
	other base conversions that are not standard.  In this case, i used my webites
	domain name.  Since the letter "e" appears twice, I changed the first one to
	upper-case.  Although the function will not cause any errors if they were both
	lowercase, preventing duplicate characters ensures the possibility of converting
	back to a decimal number with the proper algorithm.

	*/
	DECLARE @Result VARCHAR(16)
	DECLARE @BaseChars VARCHAR(14)
	DECLARE @Base TINYINT

	-- setup base characters in appropriate positions
	SET @BaseChars = 'LEwisMoten.com'

	-- determine base conversion based on length of characters defined
	SET @Base = LEN(@BaseChars)

	-- Initialize guid
	SET @Result = ''

	WHILE NOT @Number = 0
	BEGIN
		-- insert next hex character
		SET @Result = SUBSTRING(@BaseChars, (@Number - ((@Number / @Base) * @Base)) + 1, 1) + @Result

		-- Reduce number
		SET @Number = @Number / @Base

	END

	return(@Result)

END
