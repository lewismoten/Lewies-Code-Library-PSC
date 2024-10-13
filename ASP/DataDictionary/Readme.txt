'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

'Program: Lewies Data Dictionary Report Generator
'Purpose: Return schema information about database tables.
'Author: Lewis Moten
'URL: http://www.lewismoten.com
'Email: lewis@moten.com

'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

Installation:

In C:\inetpub\wwwroot, create a subfolder
called "DataDictionary"

Copy the included files to this directory:
* Default.htm
* default.js
* Readme.txt
* Report.asp
* sprDataDictionary.sql
* Style.css

Visit the URL:

http://localhost/DataDictionary/Default.htm

This program will go out to a SQL Server 2000
database and query it for information about
its schema.

If the stored procedure that it uses does not
exist, it will attempt to install it for you.

If you wish to log into the database with
a trusted connection, then you must go into
IIS and setup the directory security that you
installed the data dictionary report generator
in to only allow Integrated Windows
authentication.  No other option must be 
selected!

By default, Anonymouse access and Integrated
Windows authentication are both selected.

Notice - these scripts are ment to work with
ASP running on Microsofts Internet information
server.  The communicate with a Microsoft SQL 
Server 2000 database.