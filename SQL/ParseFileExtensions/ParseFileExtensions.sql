--**************************************
-- Name: Parse File Extensions
-- Description:Parse file extensions with SQL
-- By: Lewis Moten
--This code is copyrighted and has limited warranties.
--**************************************

SELECT DISTINCT
SUBSTRING([filename], CHARINDEX('.', [filename], LEN([filename]) - 4)+1, 99) 
FROM [MyTableOfFiles] 
WHERE [filename] LIKE '%.%'
ORDER BY
SUBSTRING([filename], CHARINDEX('.', [filename], LEN([filename]) - 4)+1, 99) 
