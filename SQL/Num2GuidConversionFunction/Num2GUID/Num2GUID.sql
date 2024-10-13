CREATE FUNCTION [dbo].[Num2GUID] (@Number BIGINT)
RETURNS UNIQUEIDENTIFIER AS
BEGIN
	/*
	Converts a base-10 number to a base-16 hexidecimal number.
	The number is returned in GUID format used with unique identifier
	datatypes.
	*/
	DECLARE @Result VARCHAR(36)
	DECLARE @BaseChars VARCHAR(16)
	DECLARE @Base TINYINT

	-- setup base characters in appropriate positions
	SET @BaseChars = '0123456789ABCDEF'

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

		-- insert seperator depending on length of hex
		SET @Result = CASE LEN(@Result)
			WHEN 12 THEN '-' + @Result
			WHEN 17 THEN '-' + @Result
			WHEN 22 THEN '-' + @Result
			WHEN 27 THEN '-' + @Result
			ELSE @Result
			END
	END

	-- insert formatted padding
	SET @Result = LEFT('00000000-0000-0000-0000-000000000000', 36 - LEN(@Result)) + @Result

	return(@Result)

END
