if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[spr_CacheEntry]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[spr_CacheEntry]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Entries]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Entries]
GO

if not exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Entries]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
 BEGIN
CREATE TABLE [dbo].[Entries] (
	[Logged] [smalldatetime] NOT NULL ,
	[Date] [smalldatetime] NOT NULL ,
	[Path] [varchar] (256) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Version] [int] NOT NULL ,
	[Action] [varchar] (128) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[User] [varchar] (32) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Event] [varchar] (16) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Comment] [text] COLLATE SQL_Latin1_General_CP1_CI_AS NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
END

GO

ALTER TABLE [dbo].[Entries] WITH NOCHECK ADD 
	CONSTRAINT [PK_Entries] PRIMARY KEY  CLUSTERED 
	(
		[Path],
		[Version],
		[Date],
		[Action]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[Entries] ADD 
	CONSTRAINT [DF_Entries_Logged] DEFAULT (getdate()) FOR [Logged]
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE dbo.spr_CacheEntry
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
			Entries 
		WHERE 
			[Date] = @Date 
			AND Path = @Path 
			AND Version = @Version 
			AND Action = @Action
		)
		RETURN
	
	INSERT INTO Entries
	(
		[Logged],
		[Date],
		Path,
		Version,
		Action,
		[User],
		Event,
		Comment
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

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

