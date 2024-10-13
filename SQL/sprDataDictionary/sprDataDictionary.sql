/*
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

'Program: Lewies Data Dictionary Report Generator
'Purpose: Return schema information about database tables.
'Author: Lewis Moten
'URL: http://www.lewismoten.com
'Email: lewis@moten.com

' See Readme.txt if you are having problems

'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
*/
CREATE PROCEDURE dbo.sprDataDictionary AS

DECLARE @Table VARCHAR(128)

-- Create temp table
DECLARE @Fields TABLE
(
	TABLE_NAME			VARCHAR(128),
	COLUMN_NAME			VARCHAR(128),
	DATA_TYPE			VARCHAR(128),
	COLUMN_DEFAULT		VARCHAR(255),
	IS_NULLABLE			VARCHAR(3),
	PK					BIT,
	FKey_NAME			VARCHAR(128),
	PKey_NAME			VARCHAR(128),
	PKEY_TABLE			VARCHAR(128),
	PKEY_COLUMN			VARCHAR(128),
	Description			VARCHAR(255),
	TABLE_DESCRIPTION	VARCHAR(255),
	IsIdentity			VARCHAR(3),
	IsRowGuidCol		VARCHAR(3),
	IsFullTextIndexed	VARCHAR(3),
	CHECK_CONSTRAINT	VARCHAR(255),
	INDEX_CONSTRAINT	VARCHAR(255)
)

/*
	grab all columns with information regarding
		Table name they are found within
		data type
		default value
		if the column can contain null values
		name of a foreign key that field relates to
		if the column represents the identity of a record in its table
		if the column represents the guid identity of a record in its table
		if it is full-text indexed
*/
INSERT INTO 
	@Fields
SELECT 
	[Cols].[TABLE_NAME], 
	[Cols].[COLUMN_NAME],
	[Cols].[DATA_TYPE] + ISNULL('(' + CAST([Cols].[CHARACTER_MAXIMUM_LENGTH] AS VARCHAR(255)) + ')', ''),
	[Cols].[COLUMN_DEFAULT],
	[Cols].[IS_NULLABLE],
	0,
	[Usage].[CONSTRAINT_NAME],
	NULL,
	NULL,
	NULL,
	NULL,
	NULL,
	COLUMNPROPERTY(OBJECT_ID([Cols].[TABLE_NAME]), [Cols].[COLUMN_NAME], 'IsIdentity'),
	COLUMNPROPERTY(OBJECT_ID([Cols].[TABLE_NAME]), [Cols].[COLUMN_NAME], 'IsRowGuidCol'),
	COLUMNPROPERTY(OBJECT_ID([Cols].[TABLE_NAME]), [Cols].[COLUMN_NAME], 'IsFulltextIndexed'),
	NULL,
	NULL
FROM 
	[INFORMATION_SCHEMA].[COLUMNS] AS [Cols]
	INNER JOIN [INFORMATION_SCHEMA].[tables] AS [Tabs]
		ON [Cols].[TABLE_NAME] = [Tabs].[TABLE_NAME]
		AND [TABLE_TYPE] = 'BASE TABLE'
	LEFT OUTER JOIN [INFORMATION_SCHEMA].[KEY_COLUMN_USAGE] AS [Usage]
		ON [Cols].[TABLE_NAME] = [Usage].[TABLE_NAME]
		AND [Cols].[COLUMN_NAME] = [Usage].[COLUMN_NAME]
		AND [Usage].[CONSTRAINT_NAME] LIKE 'FK_%'
WHERE
	[Cols].[TABLE_NAME] NOT IN('dtproperties')--, 'sysconstraints', 'syssegments')
ORDER BY 
	[Cols].[TABLE_NAME] ASC,
	[Cols].[ORDINAL_POSITION] ASC

/*
	Query databse for the constraint definition for any field
	that has a constraint applied to it.
*/
UPDATE @Fields
	SET [CHECK_CONSTRAINT] = [B].[CHECK_CLAUSE]
FROM
	@Fields AS [Fields]
	INNER JOIN [INFORMATION_SCHEMA].[CONSTRAINT_COLUMN_USAGE] AS [A]
		ON [Fields].[TABLE_NAME] = [A].[TABLE_NAME]
		AND [Fields].[COLUMN_NAME] = [A].[COLUMN_NAME]
	LEFT OUTER JOIN [INFORMATION_SCHEMA].[CHECK_CONSTRAINTS] AS [B]
		ON [A].[CONSTRAINT_NAME] = [B].[CONSTRAINT_NAME]
/*
	Query database for the name of an index for the current field
	determine if a field is indexed
*/

UPDATE @Fields
	SET [INDEX_CONSTRAINT] = [A].[CONSTRAINT_NAME]
FROM
	@Fields AS [Fields]
	INNER JOIN [INFORMATION_SCHEMA].[CONSTRAINT_COLUMN_USAGE] AS [A]
		ON [A].[TABLE_NAME] = [Fields].[TABLE_NAME]
		AND [A].[COLUMN_NAME] = [Fields].[COLUMN_NAME]
		AND [A].[CONSTRAINT_NAME] LIKE 'IX_%'

/*
	Determine what fields are primary keys
*/
UPDATE @Fields SET
	[PK] = -1
FROM
	@Fields AS [Fields]
	INNER JOIN [INFORMATION_SCHEMA].[KEY_COLUMN_USAGE] AS [Keys]
		ON [Fields].[TABLE_NAME] = [Keys].[TABLE_NAME]
		AND [Fields].[COLUMN_NAME] = [Keys].[COLUMN_NAME]
		AND [Keys].[CONSTRAINT_NAME] LIKE 'PK_%'

/*
	Determine the name of the primary key that the
	field relates to (used to determine if field is a foreign key)
*/
UPDATE @Fields SET
	[PKey_Name] = [Match].[UNIQUE_CONSTRAINT_NAME]
FROM
	@Fields AS [Fields]
	INNER JOIN [INFORMATION_SCHEMA].[REFERENTIAL_CONSTRAINTS] AS [Match]
		ON [Match].[CONSTRAINT_NAME] = [Fields].[FKey_NAME]
/*
	Determine the table and column that the foreign key relates to
*/
UPDATE @Fields SET
	[PKEY_TABLE] = [Match].[TABLE_NAME],
	[PKEY_COLUMN] = [Match].[COLUMN_NAME]
FROM
	@Fields AS [Fields]
	INNER JOIN [INFORMATION_SCHEMA].[CONSTRAINT_COLUMN_USAGE] AS [Match]
		ON [Match].[CONSTRAINT_NAME] = [Fields].[PKEY_NAME]

/*
	Query extended properties for the description of each table
*/
UPDATE @Fields SET
	[TABLE_DESCRIPTION] = CAST([N].[Value] AS VARCHAR(1000))
FROM
	::fn_listextendedproperty(N'MS_Description', N'user', 'dbo', 'TABLE', default, default, default) AS [N]
WHERE
	[N].[ObjName] = [TABLE_NAME]

-- >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
-- Query extended properties for each table
-- >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

DECLARE [Tables] CURSOR FOR
	SELECT DISTINCT [TABLE_NAME] FROM @Fields

OPEN [Tables]

FETCH NEXT FROM [Tables] INTO @Table

WHILE @@FETCH_STATUS = 0

	BEGIN

		-- Update field descriptons within current table
		UPDATE @Fields SET
			[Description] = CAST([N].[Value] AS VARCHAR(1000))
		FROM
			::fn_listextendedproperty('MS_Description', 'user', 'dbo', N'TABLE', @Table, 'COLUMN', default) as [n]
		WHERE
			[n].[ObjName] = [COLUMN_NAME]
			AND [TABLE_NAME] = @Table

		FETCH NEXT FROM [Tables] INTO @Table

	END

CLOSE [Tables]

DEALLOCATE [Tables]

-- >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
-- Return the results
-- >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

SELECT * FROM @Fields
