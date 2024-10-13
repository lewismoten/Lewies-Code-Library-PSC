CREATE FUNCTION [dbo].[Num2Oct] (@Number BIGINT)
RETURNS VARCHAR(16) AS
BEGIN
	/*
	Converts a base-10 number to a base-8 Octet number.
	(String representation of bits)
	*/
	DECLARE @Result VARCHAR(16)
	DECLARE @BaseChars VARCHAR(8)
	DECLARE @Base TINYINT

	-- setup base characters in appropriate positions
	SET @BaseChars = '01234567'

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
