<HTML>
<HEAD>
	<TITLE>Lewis Motens Data Dictionary Report Results</TITLE>
	<LINK rel="stylesheet" type="text/css" href="Style.css">
</HEAD>
<BODY>
<CENTER>
	<FONT size="+1">
		<B>Data Dictionary Report</B><BR>
	</FONT>
	Server: <%=Server.HTMLEncode(Request.Form("Server"))%><BR>
	Database: <%=Server.HTMLEncode(Request.Form("Catalog"))%><BR>
	<BR>
	<%=FormatDateTime(Date, vbLongDate)%><BR>
</CENTER>
<%
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

' Persist login information in cookies
Response.Cookies("Trusted") = Request.Form("Trusted")
Response.Cookies("Server") = Request.Form("Server")
Response.Cookies("Catalog") = Request.Form("Catalog")
Response.Cookies("UserName") = Request.Form("UserName")

' ADODB Constants used in this script
Const adOpenForwardOnly = 0
Const adLockOptimistic = 3
Const adCmdStoredProc = 4
Const adCmdText = 1
Const adStateOpen = 1

' ADODB objects
Dim Conn	' Connection
Dim Rs		' Recordset
Dim Cmd		' Command

' variables to store data from each record
Dim Table_Name				' Name of the current table
Dim AltRow					' are we working on an even or odd row?
Dim Column_Name				' Name of the current column
Dim Data_Type				' Data-Type of column (varchar, int, bit, etc.)
Dim Column_Default			' Default value of the column
Dim Is_Nullable				' Can the column contain NULL values?
Dim PK						' Is the column part of the primary key in the table?
Dim FKey_Name				' Name of the foreign key that this column refers to
Dim PKey_Name				' Name of the primary key that this columns foreign key refers to
Dim PKey_Table				' Table that foreign key refers to for current column
Dim PKey_Column				' Column that foreign key refers to for current column
Dim Description				' Description of the column
Dim Table_Description		' Description of the table
Dim IsIdentity				' Is the current column marked as the identity column for the table?
Dim IsRowGuidCol			' Is the current column marked as the RowGUID column for the table?
Dim IsFullTextIndexed		' is the current column indexed for full-text searches?
Dim Check_Constraint		' the check-constraint applied to the column
Dim Index_Constraint		' the index-constraint applied to the column

' Instantiate ADODB objects
Set Conn = Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Cmd = Server.CreateObject("ADODB.Command")

' Database connection or interaction may cause errors.
' Let's handle them ourselves by telling the server to
' ignore them.
On Error Resume Next

' Open the database
Conn.Open ConnectionString()
If Err Then WriteError(Err.Description)

' Execute Stored Procedure
RunProcedure

If Err Then

	' Stored procedure may not exist.
	InstallProcedure

	' Try to run stored procedure again	
	RunProcedure

	' Still having problems.  Might be a permission problem?
	WriteError(Err.Description)

End If

' Ok, we got past the hard part.
' Stop ignoring errors
On Error Goto 0

' Open our table
Response.Write "<TABLE width=""100%"">"

' Loop through each record
While Not Rs.EOF
	
	' if we are looking at a new table ...
	If Not Table_Name = Rs.Fields("TABLE_NAME").Value & "" Then
		
		' Reset alternate row flag
		AltRow = False
		
		' Cache current table
		Table_Name = Rs.Fields("TABLE_NAME").Value & ""

		' Write Table Name and Description
		Response.Write "<TR>"
		Response.Write "<TD class=""None"" valign=""top""><BR><BR>"
		Response.Write "<B>"
		Response.Write Server.HTMLEncode(Rs.Fields("TABLE_NAME").Value & "")
		Response.Write "</B>"
		Response.Write "</TD>"
		Response.Write "<TD class=""None"" colspan=""2""><BR><BR>"
		Response.Write Server.HTMLEncode(Rs.Fields("TABLE_DESCRIPTION").Value & "")
		Response.Write "</TD>"
		Response.Write "</TR>"
			
		' Write Column headers
		Response.Write "<TR>"
		Response.Write "<TD bgColor=""#cccccc"" valign=""top"">"
		Response.Write "<B>Column</B>"
		Response.Write "</TD>"
		Response.Write "<TD bgColor=""#cccccc"" colspan=""2"">"
		Response.Write "<B>Properties</B>"
		Response.Write "</TD>"
		Response.Write "</TR>"
			
	End If
		
	' Grab current records data and store in local variables
	' HTML encode to ensure proper display
	Data_Type = Server.HTMLEncode(Rs.Fields("DATA_TYPE").Value & "")
	Description = Server.HTMLEncode(Rs.Fields("DESCRIPTION").Value & "")
	Column_Default = Server.HTMLEncode(Rs.Fields("Column_Default").Value & "")
	Check_Constraint = Server.HTMLEncode(Rs.Fields("Check_Constraint").Value & "")
	Index_Constraint = Server.HTMLEncode(Rs.Fields("Index_Constraint").Value & "")
	IsIdentity = Server.HTMLEncode(Rs.Fields("ISIDENTITY").Value & "")
	IsRowGuidCol = Server.HTMLEncode(Rs.Fields("ISROWGUIDCOL").Value & "")
	IsFullTextIndexed = Server.HTMLEncode(Rs.Fields("ISFULLTEXTINDEXED").Value & "")
	PKey_Table = Server.HTMLEncode(Rs.Fields("PKey_Table").Value & "")
	PKey_Column = Server.HTMLEncode(Rs.Fields("PKey_Column").Value & "")
	PK = Server.HTMLEncode(Rs.Fields("PK").Value & "")
		
	' Determine row span
	RowSpan = 0
	If Not Data_Type = "" Then RowSpan = RowSpan + 1
	If PK = "True" Then RowSpan = RowSpan + 1
	If Not Description = "" Then RowSpan = RowSpan + 1
	If Not Column_Default = "" Then RowSpan = RowSpan + 1
	If Not Check_Constraint = "" Then RowSpan = RowSpan + 1
	If Not Index_Constraint = "" Then RowSpan = RowSpan + 1
	If Not IsIdentity = "0" Then RowSpan = RowSpan + 1
	If Not IsRowGuidCol = "0" Then RowSpan = RowSpan + 1
	If Not IsFullTextIndexed = "0" Then RowSpan = RowSpan + 1
	If Not PKey_Table = "" Then RowSpan = RowSpan + 1

	' Field Name
	Response.Write "<TR>"
	Response.Write "<TD valign=""top"" rowspan=""" & RowSpan & """ bgColor=""" & GetColor(AltRow) & """>"
	Response.Write Rs.Fields("COLUMN_NAME").Value & ""
	Response.Write "</TD>"
	Response.Write "<TD bgColor=""" & GetColor(AltRow) & """>Data Type</TD>"
	Response.Write "<TD bgColor=""" & GetColor(AltRow) & """>"
	Response.Write Data_Type
	Response.Write "</TD>"
	Response.Write "</TR>"

	' Write Each property
	If PK = "True" Then
		Response.Write "<TR>"
		Response.Write "<TD colspan=""2"" bgColor=""" & GetColor(AltRow) & """>Primary Key</TD>"
		Response.Write "</TR>"
	End If

	If Not Description = "" Then
		Response.Write "<TR>"
		Response.Write "<TD bgColor=""" & GetColor(AltRow) & """>Description</TD>"
		Response.Write "<TD bgColor=""" & GetColor(AltRow) & """>"
		Response.Write Description
		Response.Write "</TD>"
		Response.Write "</TR>"
	End If	
	If Not Column_Default = "" Then
		Response.Write "<TR>"
		Response.Write "<TD bgColor=""" & GetColor(AltRow) & """>Default</TD>"
		Response.Write "<TD bgColor=""" & GetColor(AltRow) & """>"
		Response.Write Column_Default
		Response.Write "</TD>"
		Response.Write "</TR>"
	End If	
	If Not Check_Constraint = "" Then
		Response.Write "<TR>"
		Response.Write "<TD bgColor=""" & GetColor(AltRow) & """>Check Constraint</TD>"
		Response.Write "<TD bgColor=""" & GetColor(AltRow) & """>"
		Response.Write Check_Constraint
		Response.Write "</TD>"
		Response.Write "</TR>"
	End If	
	If Not Index_Constraint = "" Then
		Response.Write "<TR>"
		Response.Write "<TD bgColor=""" & GetColor(AltRow) & """>Index Constraint</TD>"
		Response.Write "<TD bgColor=""" & GetColor(AltRow) & """>"
		Response.Write Index_Constraint
		Response.Write "</TD>"
		Response.Write "</TR>"
	End If	
	If Not IsIdentity = "0" Then
		Response.Write "<TR>"
		Response.Write "<TD colspan=""2"" bgColor=""" & GetColor(AltRow) & """>IDENTITY</TD>"
		Response.Write "</TR>"
	End If	
	If Not IsRowGuidCol = "0" Then
		Response.Write "<TR>"
		Response.Write "<TD colspan=""2"" bgColor=""" & GetColor(AltRow) & """>GUID Column for each Row</TD>"
		Response.Write "</TR>"
	End If	
	If Not IsFullTextIndexed = "0" Then
		Response.Write "<TR>"
		Response.Write "<TD colspan=""2"" bgColor=""" & GetColor(AltRow) & """>Full Text Indexed</TD>"
		Response.Write "</TR>"
	End If	
	If Not PKey_Table = "" Then
		Response.Write "<TR>"
		Response.Write "<TD bgColor=""" & GetColor(AltRow) & """ valign=""top"">Foreign Key</TD>"
		Response.Write "<TD bgColor=""" & GetColor(AltRow) & """>"
		Response.Write "Table: " & PKey_Table & "<BR>" & "Column: " & PKey_Column
		Response.Write "</TD>"
		Response.Write "</TR>"
	End If	
		
	' Toggle alternate row
	AltRow = Not AltRow

	' Move to next record
	Rs.MoveNext
Wend

' Close table
Response.Write "</TABLE>"

' All done!
WriteFooter

' ------------------------------------------------------------------------------
Function GetColor(IsAlt)
	' Choose background color based on the flag if the row is an alternate row.
	If IsAlt Then GetColor = "#ffffff" Else GetColor = "#eeeeee"
End Function
' ------------------------------------------------------------------------------
Function WriteFooter

	' Throw in transparent mirror
	Response.Write "<HR>"
	Response.Write "<CENTER>"
	Response.Write "Database Generation Script Created By "
	Response.Write "<A href=""http://www.lewismoten.com/"">"
	Response.Write "Lewis Moten</A>."
	Response.Write "</BODY>"
	Response.Write "</HTML>"

	' Garbage Clean Up
	If Not Rs Is Nothing Then
		If Rs.State = adStateOpen Then Rs.Close
		Set Rs = Nothing
	End If
	If Not Cmd Is Nothing Then
		If Cmd.State = adStateOpen Then Cmd.Cancel
		Set Cmd = Nothing
	End If
	If Not Conn Is Nothing Then
		If Conn.State = adStateOpen Then Conn.Close
		Set Conn = Nothing
	End If
	
	' Finish sending output to the client
	Response.Flush
	Response.End

End Function
' ------------------------------------------------------------------------------
Function ConnectionString
	' Return connection string based on values posted in
	' form data.
	
	' Trusted	- use trusted connection
	' Server	- server where database resides
	' Catalog	- database/catalog of data to be accessed
	' UserName	- Login user name
	' Password	- Login password
	
	If Request.Form("Trusted") = "True" Then
		ConnectionString = _
		"Driver={SQL Server};Server=" & Request.Form("Server") & _
		";Trusted_Connection=yes" & _
		";Database=" & Request.Form("Catalog") & _
		";"
	Else
		ConnectionString = _
		"Driver={SQL Server};Server=" & Request.Form("Server") & _
		";Uid=" & Request.Form("UserName") & _
		";Pwd=" & Request.Form("Password") & _
		";Database=" & Request.Form("Catalog") & _
		";"
	End If	
End Function
' ------------------------------------------------------------------------------
Sub WriteError(ByVal Message)
	Response.Write "<BR><BR><BR><CENTER><FONT color=""red"">"
	Response.Write Message
	Response.Write "</FONT></CENTER>"
	WriteFooter
End Sub
' ------------------------------------------------------------------------------
Sub RunProcedure
	Cmd.ActiveConnection = Conn
	Cmd.CommandType = adCmdStoredProc
	Cmd.CommandText = "dbo.sprDataDictionary"
	Rs.Open cmd
End Sub
' ------------------------------------------------------------------------------
Public Sub InstallProcedure

	Err.Clear

	Dim FSO
	Dim Text
	Dim FileName

	' Load procedure from file
	FileName = Server.MapPath("sprDataDictionary.sql")
	Set FSO = Server.CreateObject("Scripting.FileSystemObject")
	If Not FSO.FileExists(FileName) Then
		Set FSO = Nothing
		WriteError("dbp.sprDataDictionary not found in database and file ""<I>" & FileName & "</I>"" does not exist.")
	End If

	Text = FSO.OpenTextFile(FileName).ReadAll
	Set FSO = Nothing

	' add procedure to database
	Cmd.CommandType = adCmdText
	Cmd.CommandText = Text
	Cmd.Execute

	WriteError(Err.Description)

End Sub
' ------------------------------------------------------------------------------
%>