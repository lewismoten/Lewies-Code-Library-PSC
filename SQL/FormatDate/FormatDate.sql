CREATE FUNCTION [dbo].[FormatDate] (@Date datetime, @Format varchar(256))
RETURNS varchar(256)
AS  
BEGIN 

	IF @Date IS NULL RETURN('')

/*
	d	The one-digit of the two-digit day.
	dd	The two-digit day.  Single-digit day values are preceded by a zero.
	ddd	The three-character weekday abbreviation.
	dddd	The full weekday name.
	h	The one-digit or the two-digit hour in 12-hour format.
	hh	The two-digit hour in 12-hour format.  Single-digit values are preceded 
		by a zero.
	HHH 	The one-digit or the two-digit hour in 24-hour format. 
	HHHH 	The two-digit hour in 24-hour format. Single-digit values are preceded 
		by a zero. 
	n 	The one-digit or the two-digit minute. 
	nn 	The two-digit minute. Single-digit values are preceded by a zero. 
	M 	The one-digit or the two-digit month number. 
	MM 	The two-digit month number. Single-digit values are preceded by a zero. 
	MMM 	The three-character month abbreviation. 
	MMMM 	The full month name. 
	t 	The one-letter A.M. and P.M. abbreviation (that is, AM is displayed as A). 
	tt 	The two-letter A.M. and P.M. abbreviation (that is, AM is displayed as AM). 
	y 	The year is displayed as the last two digits, but with no leading zero for any 
		year that is less than 10. 
	yy 	The last two digits of the year. For example, 1998 would be displayed as 
		98. 
	yyy 	The full year. For example, 1998 would be displayed as 1998. 
	
*/

	DECLARE @d varchar(2)	-- day
	DECLARE @dd char(2)
	DECLARE @ddd char(3)
	DECLARE @dddd varchar(10)

	DECLARE @hour int
	DECLARE @h varchar(2)	-- 12 hour
	DECLARE @hh char(2)

	DECLARE @HHH varchar(2)	-- 24 hour
	DECLARE @HHHH char(2)

	DECLARE @n varchar(2)	-- minutes
	DECLARE @nn char(2)

	DECLARE @M varchar(2)	-- month
	DECLARE @MM varchar(2)
	DECLARE @MMM varchar(3)
	DECLARE @MMMM varchar(10)

	DECLARE @t char(1)		-- am/pm
	DECLARE @tt char(2)

	DECLARE @y varchar(2)	-- year
	DECLARE @yy char(2)
	DECLARE @yyy char(4)
	DECLARE @year int

	-- Retrieve values
	SET @d = CAST(DATEPART(d, @Date) AS VARCHAR(2))
	SET @dd = RIGHT('0' + @d, 2)
	SET @dddd = DATENAME(dw, @Date)
	SET @ddd = LEFT(@dddd, 3)

	SET @Hour = DATEPART(hh, @Date)
	SET @HHH = CAST(@Hour AS VARCHAR(2))
	SET @HHHH = RIGHT('0' + @HHH, 2)

	IF @Hour < 12 SET @tt = 'AM' ELSE SET @tt = 'PM'
	SET @t = LEFT(@tt, 1)
	IF @Hour > 12 SET @h = CAST(@Hour - 12 AS VARCHAR(2)) ELSE SET @h = CAST(@Hour AS VARCHAR(2))
	IF @Hour = 0 OR @Hour = 12 SET @h = '12'
	SET @hh = RIGHT('0' + @h, 2)
	
	SET @n = CAST(DATEPART(n, @Date) AS VARCHAR(2))
	SET @nn = RIGHT('0' + @n, 2)

	SET @M = CAST(DATEPART(m, @Date) AS VARCHAR(2))
	SET @MM = RIGHT('0' + @M, 2)
	SET @MMMM = DATENAME(m, @Date) 
	SET @MMM = LEFT(@MMMM, 3)

	SET @year = DATEPART(yyyy, @Date)
	SET @y = RIGHT(CAST(@year AS VARCHAR(4)), 2)
	SET @yy = RIGHT('0' + CAST(@year AS VARCHAR(4)), 2)
	SET @yyy = RIGHT('000' + CAST(@year AS VARCHAR(4)), 4)

	-- Parse format
	SET @Format = REPLACE(@Format, 'dddd', '\\\\1\\\\')
	SET @Format = REPLACE(@Format, 'ddd', '\\\\2\\\\')
	SET @Format = REPLACE(@Format, 'dd', '\\\\3\\\\')
	SET @Format = REPLACE(@Format, 'd', '\\\\4\\\\')
	SET @Format = REPLACE(@Format, 'hhhh', '\\\\5\\\\')
	SET @Format = REPLACE(@Format, 'hhh', '\\\\6\\\\')
	SET @Format = REPLACE(@Format, 'hh', '\\\\7\\\\')
	SET @Format = REPLACE(@Format, 'h', '\\\\8\\\\')
	SET @Format = REPLACE(@Format, 'nn', '\\\\9\\\\')
	SET @Format = REPLACE(@Format, 'n', '\\\\10\\\\')
	SET @Format = REPLACE(@Format, 'mmmm', '\\\\11\\\\')
	SET @Format = REPLACE(@Format, 'mmm', '\\\\12\\\\')
	SET @Format = REPLACE(@Format, 'mm', '\\\\13\\\\')
	SET @Format = REPLACE(@Format, 'm', '\\\\14\\\\')
	SET @Format = REPLACE(@Format, 'tt', '\\\\15\\\\')
	SET @Format = REPLACE(@Format, 't', '\\\\16\\\\')
	SET @Format = REPLACE(@Format, 'yyy', '\\\\17\\\\')
	SET @Format = REPLACE(@Format, 'yy', '\\\\18\\\\')
	SET @Format = REPLACE(@Format, 'y', '\\\\19\\\\')

	-- Populate with values
	SET @Format = REPLACE(@Format, '\\\\1\\\\', @dddd)
	SET @Format = REPLACE(@Format, '\\\\2\\\\', @ddd)
	SET @Format = REPLACE(@Format, '\\\\3\\\\', @dd)
	SET @Format = REPLACE(@Format, '\\\\4\\\\', @d)
	SET @Format = REPLACE(@Format, '\\\\5\\\\', @hhhh)
	SET @Format = REPLACE(@Format, '\\\\6\\\\', @hhh)
	SET @Format = REPLACE(@Format, '\\\\7\\\\', @hh)
	SET @Format = REPLACE(@Format, '\\\\8\\\\', @h)
	SET @Format = REPLACE(@Format, '\\\\9\\\\', @nn)
	SET @Format = REPLACE(@Format, '\\\\10\\\\', @n)
	SET @Format = REPLACE(@Format, '\\\\11\\\\', @mmmm)
	SET @Format = REPLACE(@Format, '\\\\12\\\\', @mmm)
	SET @Format = REPLACE(@Format, '\\\\13\\\\', @mm)
	SET @Format = REPLACE(@Format, '\\\\14\\\\', @m)
	SET @Format = REPLACE(@Format, '\\\\15\\\\', @tt)
	SET @Format = REPLACE(@Format, '\\\\16\\\\', @t)
	SET @Format = REPLACE(@Format, '\\\\17\\\\', @yyy)
	SET @Format = REPLACE(@Format, '\\\\18\\\\', @yy)
	SET @Format = REPLACE(@Format, '\\\\19\\\\', @y)

	return(@Format)

END

