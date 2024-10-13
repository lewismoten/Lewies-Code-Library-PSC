
--**************************************
-- for :Random Pronouncable Password
--**************************************
-- Copyright (c) 2001, Lewis Edward Moten III, All rights reserved.

--**************************************
-- Name: Random Pronouncable Password
-- Description:Procedure made in order to create a random password that is somewhat pronouncable and easy to remember.
-- By: Lewis Moten
--
--
-- Inputs:None
--
-- Returns:None
--
--Assumes:tested only with SQL Server 2000
--
--Side Effects:None
--This code is copyrighted and has limited warranties.
--**************************************

CREATE PROCEDURE [dbo].[AccountUsers_RandomPassword]
(
	@MinLength smallint = 5,
	@MaxLength smallint = 8
)
AS
/*
Description:
Procedure made in order to create a random password that is somewhat pronouncable and easy to remember.
Assumptions:
This procedure assumes your are expecting a password data type of nvarchar(16) 
Details:
Password length can be defined by suggesting the minimum and maximum length. For fixed
size passwords, you may suggest the same number for both minimum and maximum
(i.e. - EXEC AccountUsers_RandomPassword 8, 8)
Author:
Created by Lewis Edward Moten III on Saturday, July 07, 2001
email: lewis@moten.com
url: http://www.lewismoten.com
Copyright (c) 2001, Lewis Edward Moten III, All rights reserved.
You may alter the text for your own database only to esure proper working conditions against your
database (i.e. - Changing the location of the table where user names reside and the field name)
If you redistribute the code, you must include the original content including copyright and contact
information.
*/
declare @Password nvarchar(16)
declare @Length smallint
declare @WriteConsonant bit
declare @Position smallint
declare @Vowels nchar(5)
declare @Consonants nchar(16)
declare @DoubleConsonants nchar(11)
declare @Random smallint
declare @LetterPosition int
declare @Letter nchar(1)
select @Vowels = 'aeiou'
select @Consonants = 'bcdfghjklmnprstv'
select @DoubleConsonants = 'cdfglmnprst'
select @Password = ''
select @WriteConsonant = 1
-- determine length
Select @Length = @MinLength + Round(Rand() * (@MaxLength - @MinLength), 0, 0)
Select @Position = 1
while @Position <= @Length


    begin
    	select @Random = Round(Rand() * 100, 0, 0) + 1
    	-- 10% double consonants
    	if not @Password = '' and @WriteConsonant = 1 and @Random < 10
    	begin
    		select @LetterPosition = Convert(int, Round((Rand() * (Len(@DoubleConsonants)-1)), 0, 0), 1) +1
    		select @Letter = SubString(@DoubleConsonants, @LetterPosition, 1)
    		select @Password = @Password + Replicate(@Letter, 2)
    		select @Position = @Position + 1 -- because we just took up 2 positions
    		select @WriteConsonant = 0
    	end
    	else
    	begin
    		-- 80% consonants
    		if @WriteConsonant = 1 and @Random < 90
    		begin
    			select @LetterPosition = Convert(int, Round((Rand() * (Len(@Consonants)-1)), 0, 0), 1) + 1
    			select @Letter = SubString(@Consonants, @LetterPosition, 1)
    			select @Password = @Password + @Letter
    			select @WriteConsonant = 0
    		end
    		-- 10% vowels
    		else
    		begin
    			select @LetterPosition = Convert(int, Round((Rand() * (Len(@Vowels)-1)), 0, 0), 1) + 1
    			select @Letter = SubString(@Vowels, @LetterPosition, 1)
    			select @Password = @Password + @Letter
    			select @WriteConsonant = 1
    		end
    	end
    	select @Position = @Position + 1 -- incriment counter
end
-- just in case that last character was a double consonant
select @password = Left(@Password, @length)
-- return password
SELECT @password As [Password]
GO