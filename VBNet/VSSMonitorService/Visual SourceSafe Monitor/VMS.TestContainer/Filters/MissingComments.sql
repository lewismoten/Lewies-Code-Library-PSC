SELECT 
	getdate() AS Date,
	'' AS Path,
	0 AS Version,
	'' AS [Action],
	[User],
	Event,
	CAST(COUNT(0) AS VARCHAR(32)) + ' ' + Event + 
	' events without comments.' AS Comment
FROM 
	Entries 
WHERE 
	Path LIKE '$/Visual SourceSafe Monitor/%'
	AND DATALENGTH(Comment) < 1
	AND Logged > @LastAlerted
GROUP BY 
	[User], 
	Event