--**************************************
-- Name: sp_ZipCode_Distance
-- Description:This code looks up 2 zip codes latitude and longitude and determines the distance between the two in miles.
-- By: Lewis Moten
--
--
-- Inputs:-- @ZipCode1 -- first zip code to look for
-- @ZipCode2 -- second zip code to look for
--
-- Returns:-- @Miles -- Distance between the two zip codes
--
--Assumes:-- Assumes that you have a Table called
--
-- ZipCodeList
--.Latitude
--.Longitude
--.ZipCode
--
-- The table is assumed to be populated
-- with accurate information. Various
-- places on the internet offer zip
-- code data. ZipInfo dot com offers
-- a short demo list of data you can
-- test with. You may also want to try
-- internet sites for the census beuro
-- or the post office.
--
--Side Effects:-- Distance may be off by about 36 feet.
--This code is copyrighted and has limited warranties.
--**************************************

Create Procedure sp_ZipCode_Distance
	(
		@ZipCode1 Numeric(5, 0) = Null,
		@ZipCode2 Numeric(5, 0) = Null
	)
As
	set nocount on
	Declare @Latitude1 Numeric(6,4)
	Declare @Longitude1 Numeric(7,4)
	Declare @Latitude2 Numeric(6,4)
	Declare @Longitude2 Numeric(7,4)
	
	Declare @Miles Float
	-- No data? GET OUT!
	If @ZipCode1 Is Null Or @ZipCode2 Is Null
	Begin
		Select -1 As Miles
		Return
	End
	
	-- Lookup longitude, latitude for zip codes
	Select @Latitude1 = Latitude, @Longitude1 = Longitude From ZipCodeList Where ZipCode = @ZipCode1
	Select @Latitude2 = Latitude, @Longitude2 = Longitude From ZipCodeList Where ZipCode = @ZipCode2
	
	-- No data? GET OUT!
	If @Latitude1 Is Null Or @Longitude1 Is Null Or @Latitude2 Is Null Or @Longitude2 Is Null
	Begin
		Select -1 As Miles
		Return
	End
	-- Do your magic ...
	Select @Miles = 3958.75 * ACos(Sin(@Latitude1/57.2958) *
Sin(@Latitude2/57.2958) + 
Cos(@Latitude1/57.2958) * 
Cos(@Latitude2/57.2958) * 
Cos(@Longitude2/57.2958 - @Longitude1/57.2958))
	-- Return the results 	
	Select @Miles As Miles

	return