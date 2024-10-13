-- drop spr_CacheEntry procedure
IF EXISTS
	(
	SELECT * 
	FROM 
		[dbo].[sysobjects]
	WHERE
		[id] = object_id(N'[dbo].[spr_CacheEntry]') 
		AND OBJECTPROPERTY([id], N'IsProcedure') = 1
	)
	DROP PROCEDURE [dbo].[spr_CacheEntry]

GO

-- drop entries table
IF EXISTS
	(
	SELECT * 
	FROM [dbo].[sysobjects]
	WHERE
		[id] = object_id(N'[dbo].[Entries]') 
		AND OBJECTPROPERTY([id], N'IsUserTable') = 1
	)
	DROP TABLE [dbo].[Entries]

GO

-- add entries table
IF NOT EXISTS
	(
	SELECT * 
	FROM [dbo].[sysobjects]
	WHERE
		[id] = object_id(N'[dbo].[Entries]') 
		AND OBJECTPROPERTY([id], N'IsUserTable') = 1
	)
	BEGIN
		CREATE TABLE 
			[dbo].[Entries] 
		(
			[Logged] SMALLDATETIME NOT NULL,
			[Date] SMALLDATETIME NOT NULL,
			[Path] VARCHAR(256) NOT NULL,
			[Version] INT NOT NULL,
			[Action] VARCHAR(128) NOT NULL,
			[User] VARCHAR(32) NOT NULL,
			[Event] VARCHAR(16) NOT NULL,
			[Comment] TEXT NULL 
		) 
		ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
	END

GO

-- add new key combination
ALTER TABLE 
	[dbo].[Entries] 
WITH NOCHECK ADD CONSTRAINT 
	[PK_Entries] PRIMARY KEY CLUSTERED
	(
		[Path],
		[Version],
		[Date],
		[Action]
	) ON [PRIMARY]

GO

-- add new default date
ALTER TABLE 
	[dbo].[Entries] 
ADD CONSTRAINT 
	[DF_Entries_Logged] DEFAULT 
	(
		getdate()
	) FOR [Logged]

GO

-- add procedure spr_CacheEntry

CREATE PROCEDURE [dbo].[spr_CacheEntry]
(
	@Logged SMALLDATETIME,
	@Date SMALLDATETIME,
	@Path VARCHAR(256),
	@Version INT,
	@Action VARCHAR(128),
	@User VARCHAR(32),
	@Event VARCHAR(16),
	@Comment TEXT
)
AS
	SET NOCOUNT ON
	
	IF EXISTS(
		SELECT 
			0 
		FROM 
			[Entries]
		WHERE 
			[Date] = @Date 
			AND [Path] = @Path 
			AND [Version] = @Version 
			AND [Action] = @Action
		)
		RETURN
	
	INSERT INTO [Entries]
	(
		[Logged],
		[Date],
		[Path],
		[Version],
		[Action],
		[User],
		[Event],
		[Comment]
	)
	VALUES
	(
		@Logged,
		@Date,
		@Path,
		@Version,
		@Action,
		@User,
		@Event,
		@Comment
	)

	RETURN