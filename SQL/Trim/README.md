# [Lewie's Code Library PSC](../../README.md)

Open source projects that I had published to Planet Source Code.

## [SQL](../README.md)

### Trim

*10/15/2001 5:51:45 PM*

Allows you to Trim your strings when querying tables. I'm tired of using the Trim function in my ASP pages and Visual Basic. I created this little feature to trim away within SQL Server itself. This is a User Defined function rather then a stored procedure, so you can use it rite inside your select statements. "SELECT Trim([FirstName]) FROM [Users]". Sometimes you may find that you need to prefix the function with the owner - "SELECT dbo.Trim([FirstName]) FROM [Users]". Within the SQL Server Enterprise manager you will need to expand the "Databases" node. Locate your database and expand the node. Next, find the User Defined Functions node. Right-Mouse-Click and choose "New User Defined Function ...". Paste my code into it.


