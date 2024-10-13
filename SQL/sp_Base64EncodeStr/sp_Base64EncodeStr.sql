--**************************************
-- Name: sp_Base64EncodeStr
-- Description:This procedure demonstrates how to instantiate an OLE object.
-- I have used an ActiveX dll called
-- AspConv&#8482; (provided in AspLib&#8482;) to encode text with Base64 encoding in this example.
-- The object must be registered on THE DATABASE SERVER (not to be
-- confused with the web server/client/ or application server)
-- By: Lewis Moten
--
--
-- Inputs:--
-- @ClearText 
-- Text to be encoded
-- @EncodedText
-- Variable to store encoded text
--
-- Returns:--
-- @EncodedText
--Text that has been Base64 Encoded
--
--Assumes:--
-- The object must be registered on
-- THE DATABASE SERVER (not to be
-- confused with the web server).
-- More information about this object
-- and others created by ServerObjects
-- Inc. can be found at:
--
-- serverobjects.com
--
--Side Effects:--
-- I have had some instances where
-- this procedure wouldn't work due
-- to permission issues. More
-- information about this can be
-- found in the SQL Server Online
-- Books under sp_OACreate in the
-- index.
--This code is copyrighted and has limited warranties.
--**************************************

Create Procedure sp_OLE_AspConv_Base64EncodeStr
	(
		@ClearText VarChar(255) = Null,
		@EncodedText VarChar(255) = Null OutPut
	)
As
	Set Nocount On
	Declare @HResult Int
	Declare @AspConv Int	-- Pointer to Object Reference
	
	-- Error Variables
	Declare @Err_Source VarChar(255)
	Declare @Err_Description VarChar(255)
	Declare @Err_HelpFile VarChar(255)
	Declare @Err_HelpID Int
	
	-- Make sure user sent something to encode
	If @ClearText Is Null
		Begin
			-- Note: Might want to raise an error here
			Return
		End
	If @ClearText = ''
		Begin
			-- Note: Might want to raise an error here
			Return
		End
	
	/*
	Create an instance of the OLE object similar to stating:
	
	Set @AspConv = CreateObject('AspConv.Expert')
	*/
	Exec @HResult = sp_OACreate 'AspConv.Expert', @AspConv Out
	
	
	If @HResult != 0
		Begin
			Exec sp_OAGetErrorInfo
				@AspConv,
				@Err_Source OUT,
				@Err_Description OUT,
				@Err_HelpFile OUT,
				@Err_HelpID OUT
			
			RaisError(@Err_Description, 0, 1)
				
		Return
		End
	
	/*
	
	Call a method of the OLE object similar to stating:
	
		@EncodedText = @AspConv.Base64EncodeStr(@ClearText)
	
	*/
		
	Exec @HResult = sp_OAMethod @AspConv, Base64EncodeStr, 
		@EncodedText OutPut, @ClearText
		
	
	-- Check for any errors	
	If @HResult != 0
		Begin
			Exec sp_OAGetErrorInfo
				@AspConv,
				@Err_Source OUT,
				@Err_Description OUT,
				@Err_HelpFile OUT,
				@Err_HelpID OUT
			
			RaisError(@Err_Description, 0, 1)
			
			Return
		End
	-- Destroy the OLE object.
	Exec @HResult = sp_OADestroy @AspConv
	-- Check for any errors	
	If @HResult != 0
		Begin
			Exec sp_OAGetErrorInfo
				@AspConv,
				@Err_Source OUT,
				@Err_Description OUT,
				@Err_HelpFile OUT,
				@Err_HelpID OUT
			
			RaisError(@Err_Description, 0, 1)
			
		Return
		End
	Return