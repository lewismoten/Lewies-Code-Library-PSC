--**************************************
-- Name: US Government Holidays
-- Description:Provide a date to this function and it will move it to the next US Government Business day - taking holidays into effect. US Government holidays are New Years Day, Birthday of Martin Luther King, Jr., Washington's Birthday, Memorial Day, Independence Day, Labor Day, Columbus Day, Veterans Day, Thanksgiving, and Christmas. Functions provided may be used for other special days as well.
-- By: Lewis Moten
--
--
-- Inputs:select dbo.USGovBusinessDay('5/24/2003') -- Saturday before memorial day will return as the 27th (Tuesday)
--
-- Returns:None
--
--Assumes:None
--
--Side Effects:None
--This code is copyrighted and has limited warranties.
--**************************************

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[NthDayOfMonth]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[NthDayOfMonth]
GO
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[GovHolidayAdjustment]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[GovHolidayAdjustment]
GO
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[NthDayEndOfMonth]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[NthDayEndOfMonth]
GO
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USGovBusinessDay]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[USGovBusinessDay]
GO
SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO
CREATE FUNCTION dbo.NthDayOfMonth
(
	@Year INT,
	@Month INT,
	@N INT,
	@DayOfWeek INT
)
RETURNS SMALLDATETIME
AS


    BEGIN
    	-- Primarily used for determining holidays
    	-- falling on the Nth day of the month
    	
    	-- Thanksgiving - 4th Thursday of November
    	-- Martin Luther King's Birthday - 3rd Monday of January
    	DECLARE @Date SMALLDATETIME
    	-- Determine first day of month
    	SET @Date = CAST
    		(
    			CAST(@Month AS VARCHAR(2)) + '/' + 
    			CAST(1 AS VARCHAR(2)) + '/' + 
    			CAST(@Year AS VARCHAR(4)
    		) AS SMALLDATETIME)
    	
    	-- Adjust to lay on the first week day needed
    	WHILE NOT @DayOfWeek = DatePart(dw, @Date)
    	BEGIN
    		SET @Date = DateAdd(d, 1, @Date)
    	END
    	-- Move to the Nth Week
    	SET @Date = DateAdd(ww, @N -1, @Date)	
    	
    	RETURN @Date
    	
END
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO
SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO
CREATE FUNCTION dbo.GovHolidayAdjustment
	(
	@Date SMALLDATETIME
	)
RETURNS SMALLDATETIME
AS
	BEGIN
	DECLARE @Sunday INT
	DECLARE @Monday INT
	DECLARE @Tuesday INT
	DECLARE @Wednesday INT
	DECLARE @Thursday INT
	DECLARE @Friday INT
	DECLARE @Saturday INT
	
	DECLARE @January INT
	DECLARE @February INT
	DECLARE @March INT
	DECLARE @April INT
	DECLARE @May INT
	DECLARE @June INT
	DECLARE @July INT
	DECLARE @August INT
	DECLARE @September INT
	DECLARE @October INT
	DECLARE @November INT
	DECLARE @December INT
	
	SET @Sunday = 1
	SET @Monday = 2
	SET @Tuesday = 3
	SET @Wednesday = 4
	SET @Thursday = 5
	SET @Friday = 6
	SET @Saturday = 7
	
	SET @January = 1
	SET @February = 2
	SET @March = 3
	SET @April = 4
	SET @May = 5
	SET @June = 6
	SET @July = 7
	SET @August = 8
	SET @September = 9
	SET @October = 10
	SET @November = 11
	SET @December = 12
	/*
	Federal law (5 U.S.C. 6103) establishes the 
	following public holidays for Federal employees. 
	Please note that most Federal employees work on 
	a Monday through Friday schedule. For these 
	employees, when a holiday falls on a nonworkday
	-- Saturday or Sunday -- the holiday usually is 
	observed on Monday (if the holiday falls on Sunday) 
	or Friday (if the holiday falls on Saturday).
	*/
	
	-- New Years day (Jan 1st)
	IF DATEPART(m, @Date) = @January AND DATEPART(d, @Date) = 1
	BEGIN
		SET @Date = 
		CASE DATEPART(dw, @Date) 
		WHEN @Sunday THEN DateAdd(d, 2, @Date)
		WHEN @Saturday THEN DateAdd(d, 3, @Date)
		WHEN @Friday THEN DateAdd(d, 3, @Date)
		ELSE DateAdd(d, 1, @Date)
		END
	END
	
	-- Birthday of Martin Luther King, Jr. (3rd Monday of Jan)
	IF @Date = dbo.NthDayOfMonth(DATEPART(yyyy, @Date), @January, 3, @Monday)
	BEGIN
		SET @Date = 
		CASE DATEPART(dw, @Date) 
		WHEN @Sunday THEN DateAdd(d, 2, @Date)
		WHEN @Saturday THEN DateAdd(d, 3, @Date)
		WHEN @Friday THEN DateAdd(d, 3, @Date)
		ELSE DateAdd(d, 1, @Date)
		END
	END
		
	-- Washington's Birthday (3rd Monday in February)
	IF @Date = dbo.NthDayOfMonth(DATEPART(yyyy, @Date), @February, 3, @Monday)
	BEGIN
		SET @Date = 
		CASE DATEPART(dw, @Date) 
		WHEN @Sunday THEN DateAdd(d, 2, @Date)
		WHEN @Saturday THEN DateAdd(d, 3, @Date)
		WHEN @Friday THEN DateAdd(d, 3, @Date)
		ELSE DateAdd(d, 1, @Date)
		END
	END
		
	-- Memorial Day (Last Monday in May)
	IF @Date = dbo.NthDayEndOfMonth(DATEPART(yyyy, @Date), @May, 1, @Monday)
	BEGIN
		SET @Date = 
		CASE DATEPART(dw, @Date) 
		WHEN @Sunday THEN DateAdd(d, 2, @Date)
		WHEN @Saturday THEN DateAdd(d, 3, @Date)
		WHEN @Friday THEN DateAdd(d, 3, @Date)
		ELSE DateAdd(d, 1, @Date)
		END
	END 
		
	-- Independence Day (July 4)
	IF DATEPART(m, @Date) = @July AND DATEPART(d, @Date) = 4
	BEGIN
		SET @Date = 
		CASE DATEPART(dw, @Date) 
		WHEN @Sunday THEN DateAdd(d, 2, @Date)
		WHEN @Saturday THEN DateAdd(d, 3, @Date)
		WHEN @Friday THEN DateAdd(d, 3, @Date)
		ELSE DateAdd(d, 1, @Date)
		END
	END
		
	-- Labor Day (1st Monday in September)
	IF @Date = dbo.NthDayOfMonth(DATEPART(yyyy, @Date), @September, 1, @Monday)
	BEGIN
		SET @Date = 
		CASE DATEPART(dw, @Date) 
		WHEN @Sunday THEN DateAdd(d, 2, @Date)
		WHEN @Saturday THEN DateAdd(d, 3, @Date)
		WHEN @Friday THEN DateAdd(d, 3, @Date)
		ELSE DateAdd(d, 1, @Date)
		END
	END 
		
	-- Columbus Day (2nd Monday in October)
	IF @Date = dbo.NthDayOfMonth(DATEPART(yyyy, @Date), @October, 2, @Monday)
	BEGIN
		SET @Date = 
		CASE DATEPART(dw, @Date) 
		WHEN @Sunday THEN DateAdd(d, 2, @Date)
		WHEN @Saturday THEN DateAdd(d, 3, @Date)
		WHEN @Friday THEN DateAdd(d, 3, @Date)
		ELSE DateAdd(d, 1, @Date)
		END
	END 
		
	-- Veterans Day (Nov 11)
	IF DATEPART(m, @Date) = @November AND DATEPART(d, @Date) = 11
	BEGIN
		SET @Date = 
		CASE DATEPART(dw, @Date) 
		WHEN @Sunday THEN DateAdd(d, 2, @Date)
		WHEN @Saturday THEN DateAdd(d, 3, @Date)
		WHEN @Friday THEN DateAdd(d, 3, @Date)
		ELSE DateAdd(d, 1, @Date)
		END
	END 
	
	-- Thanksgiving Day (4th Thursday of Nov)
	IF @Date = dbo.NthDayOfMonth(DATEPART(yyyy, @Date), @November, 4, @Thursday)
	BEGIN
		SET @Date = 
		CASE DATEPART(dw, @Date) 
		WHEN @Sunday THEN DateAdd(d, 2, @Date)
		WHEN @Saturday THEN DateAdd(d, 3, @Date)
		WHEN @Friday THEN DateAdd(d, 3, @Date)
		ELSE DateAdd(d, 1, @Date)
		END
	END 
		
	-- Christmas Day (Dec 25)
	IF DATEPART(m, @Date) = @December AND DATEPART(d, @Date) = 25
	BEGIN
		SET @Date = 
		CASE DATEPART(dw, @Date) 
		WHEN @Sunday THEN DateAdd(d, 2, @Date)
		WHEN @Saturday THEN DateAdd(d, 3, @Date)
		WHEN @Friday THEN DateAdd(d, 3, @Date)
		ELSE DateAdd(d, 1, @Date)
		END
	END 
		
	RETURN @Date
	END
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO
SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO
CREATE FUNCTION dbo.NthDayEndOfMonth
(
	@Year INT,
	@Month INT,
	@N INT,
	@DayOfWeek INT
)
RETURNS SMALLDATETIME
AS


    BEGIN
    	-- Primarily used for determining holidays
    	-- falling on the Nth day before the end
    	-- of the month
    	
    	-- Memorial day, 1st Monday before the end of May
    	DECLARE @Date SMALLDATETIME
    	-- Determine first day of month
    	SET @Date = CAST
    		(
    			CAST(@Month AS VARCHAR(2)) + '/' + 
    			CAST(1 AS VARCHAR(2)) + '/' + 
    			CAST(@Year AS VARCHAR(4)
    		) AS SMALLDATETIME)
    	
    	-- Adjust to lay on the first week day needed
    	WHILE NOT @DayOfWeek = DatePart(dw, @Date)
    	BEGIN
    		SET @Date = DateAdd(d, 1, @Date)
    	END
    	-- Count Number of times day appears in month
    	DECLARE @Appearance INT
    	SET @Appearance = 1
    	WHILE DATEPART(m, DATEADD(d, 7 * @Appearance, @Date)) = @Month
    	BEGIN
    		SET @Appearance = @Appearance + 1	
    	END
    	-- Move to the Nth Week
    	SET @Appearance = @Appearance - @N
    	SET @Date = DateAdd(ww, @Appearance, @Date)	
    	
    	RETURN @Date
    	
END
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO
SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO
CREATE FUNCTION dbo.USGovBusinessDay
	(
	@Date SMALLDATETIME
	)
RETURNS SMALLDATETIME
AS
	BEGIN
	DECLARE @Sunday INT
	DECLARE @Monday INT
	DECLARE @Tuesday INT
	DECLARE @Wednesday INT
	DECLARE @Thursday INT
	DECLARE @Friday INT
	DECLARE @Saturday INT
	
	SET @Sunday = 1
	SET @Monday = 2
	SET @Tuesday = 3
	SET @Wednesday = 4
	SET @Thursday = 5
	SET @Friday = 6
	SET @Saturday = 7
	
	-- Bump business day if it falls on a
	-- government holiday
	SET @Date = dbo.GovHolidayAdjustment(@Date)
	
	-- Change Sunday to Monday
	IF DATEPART(dw, @Date) = @Sunday
	SET @Date = DateAdd(d, 1, @Date)
	-- Change Saturday to Monday
	IF DATEPART(dw, @Date) = @Saturday
	SET @Date = DateAdd(d, 2, @Date)
	-- Moving from sunday or saturday may have
	-- changed the date to a government holiday.
	-- check for the holidays again.
	SET @Date = dbo.GovHolidayAdjustment(@Date)
	RETURN @Date
	END
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO