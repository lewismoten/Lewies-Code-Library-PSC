CREATE FUNCTION [dbo].[Num2BaseX] (@Number BIGINT, @BaseChars VARCHAR(255))
RETURNS VARCHAR(256) AS
BEGIN
	/*

	Converts a base-10 number to a user-defined base.  The user simply passes
	the base characters along with the number.

	Examples:

	-- Bit String
	PRINT dbo.Num2BaseX(1975, '01')

	-- Hexidecimal
	PRINT dbo.Num2BaseX(1975, '0123456789ABCDEF')

	*/
	DECLARE @Result VARCHAR(16)
	DECLARE @Base TINYINT

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
