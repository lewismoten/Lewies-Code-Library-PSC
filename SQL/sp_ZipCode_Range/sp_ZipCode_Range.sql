--**************************************
-- Name: sp_ZipCode_Range
-- Description:Returns all zip codes within a specified range of a specific zip code.
-- By: Lewis Moten
--
--
-- Inputs:-- 
-- @ZipCode
--the zip code to search from
-- @Miles
--the maximum distance to
--search for other zip codes
--
-- Returns:--
-- ZipCode
--Zipcode within @Miles
--of @ZipCode
-- Miles
--Distance from @ZipCode
--
--Assumes::-- Assumes that you have a Ta
-- ble called
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
--
-- a short demo list of data you can
-- test with. You may also want to try
-- internet sites for the census beuro
-- or the post office.
--
--Side Effects:--
-- Distance may be off by about
-- 36 feet.
--This code is copyrighted and has limited warranties.
--**************************************

/*
Returns zip codes within specified range.
*/
Create Procedure sp_ZipCode_Range
	(
		@ZipCode Numeric(5, 0) = Null,
		@Miles Float
	)
As
	set nocount on
	
	Declare @Latitude Numeric(6,4)
	Declare @Longitude Numeric(7,4)
	-- Lookup longitude, latitude for zip codes
	Select @Latitude = Latitude, @Longitude = Longitude From ZipCodeList Where ZipCode = @ZipCode
	
	Select
		ZipCode,
		3958.75 * ACos(Sin(@Latitude/57.2958) *
		Sin(Latitude/57.2958) + 
		Cos(@Latitude/57.2958) * 
		Cos(Latitude/57.2958) * 
		Cos(Longitude/57.2958 - @Longitude/57.2958)) As Miles
	From
		ZipCodeList
	Where
		Longitude Is Not Null
		And Latitude Is Not Null
		And @Miles >= 
			(
			3958.75 * ACos(Sin(@Latitude/57.2958) *
			Sin(Latitude/57.2958) + 
			Cos(@Latitude/57.2958) * 
			Cos(Latitude/57.2958) * 
			Cos(Longitude/57.2958 - @Longitude/57.2958))
			)
--		And ZipCode != @ZipCode
	return