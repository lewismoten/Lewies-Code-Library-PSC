SELECT 
	* 
FROM 
	Entries 
WHERE 
	Path LIKE '$/Visual SourceSafe Monitor/%'
	AND Event IN
	(
		'deleted',
		'destroyed',
		'purged',
		'added', 
		'moved', 
		'restored'
	)
	AND Logged > @LastAlerted