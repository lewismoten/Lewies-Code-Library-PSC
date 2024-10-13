--**************************************
-- Name: Newbie Line Seperators
-- Description:Simple way to insert a carriage return + line feed into a string. Really beginner stuff ... I keep forgetting sometimes. Rather spend time looking it up in the docs or old code, I'm putting it here to access it quicker. Hopefully someone else may also find it helpful.
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

DECLARE @CRLF CHAR(2)
SET @CRLF = CHAR(13) + CHAR(10)
print 'line 1' + @CRLF + 'line 2'