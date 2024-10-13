
--**************************************
-- for :Suggest Available Usernames
--**************************************
-- Copyright (c) 2001, Lewis Edward Moten III, All rights reserved.
--**************************************
-- Name: Suggest Available Usernames
-- Description:Procedure created in order to determine if a duplicate user name exists 
-- during a sign up process. If a user name does not exist, an empty
-- recordset is returned. Otherwise, any digits found at the end of the username
-- are trimmed, and new usernames that do not yet exist are suggested with
-- a suffix of numbers appended to the end.
-- By: Lewis Moten
--
--
-- Inputs:None
--
-- Returns:None
--
--Assumes:Tested only on SQL Sever 2000
--
--Side Effects:None
--This code is copyrighted and has limited warranties.
--**************************************

CREATE PROCEDURE [dbo].[AccountUsers_Exists]
(
	@UserName nvarchar(16),
	@Suggest int = 10
)
AS
/*
Description:
Procedure created in order to determine if a duplicate user name exists 
during a sign up process. If a user name does not exist, an empty
recordset is returned. Otherwise, any digits found at the end of the username
are trimmed, and new usernames that do not yet exist are suggested with
a suffix of numbers appended to the end.
Assumptions:
This procedure assumes your users exist in a table called [AccountUsers]. The user names
are stored in a field called [UserName] with a datatype of nvarchar(16)
Details:
The suffix appended to user names begins at 1 and continues until the username is
16 characters in length or 10 user names (@Suggest) have been found
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
-- need bigint because number can range from 1 to 9,999,999,999,999,999 
	-- (max char length of number = max length of username)
Declare @Suffix bigint
Declare @MaxSuffix bigint
-- make sure someone submitted a valid user name
if @UserName = '' or @UserName is null


    begin
    	raiserror('User name is required', 0, 1)
    	return
end
-- if user name does not exist, return empty recordset
if not exists(select [userid] from [accountusers] where [username] = @username)


    begin
    	return
end
-- Trim any current suffix of numbers from username (i.e. - 'lewie104' becomes 'lewie')
while Right(@userName, 1) like '[0-9]'


    begin
    	select @username = left(@username, len(@username) - 1)
    	if len(@username) = 1
    	begin
    		break
    	end
    	else
    	begin
    		continue
    	end
end
-- if the entire username is a digit
if @Username like '[0-9]'


    begin
    	-- set user name to nothing
    	select @username = ''
end
-- create temporary table to hold new user names
create table #TempUsers([UserName] nvarchar(16))
-- see if trimmed version of username exists
if not @username = ''


    begin
    	if not exists(select [username] from [accountusers] where [username] = @username)
    	begin
    		insert into #TempUsers ([UserName]) values (@UserName)
    	end
end
-- initialize suffix of numbers
-- if you want numbers to begin at 100 (or another number), change the suffix to reflect this
select @Suffix = 1
-- determine largest number to append as suffix
select @MaxSuffix = Convert(bigint, Replicate('9', (16 - len(@UserName))))
-- search for available users names while we have less then 10 (@Suggest) in the temporary table
while (select count([UserName]) from #TempUsers) < @Suggest and @Suffix <= @MaxSuffix


    begin
    	-- if new user name does not exist
    	if not exists(select [UserName] from [AccountUsers] Where [UserName] = @UserName + Convert(nVarChar(16), @Suffix))
    	begin
    		-- insert username into suggested user names
    		insert into #TempUsers ([UserName]) values (@UserName + Convert(nVarChar(16), @Suffix))
    	end
    	-- incirment suffix
    	Select @Suffix = @Suffix + 1
end
-- return list of user names
select [UserName] From #TempUsers
GO