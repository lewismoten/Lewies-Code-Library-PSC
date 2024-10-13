<!--#INCLUDE FILE="clsDatabase.asp"-->
<LINK rel="stylesheet" type="text/css" href="PickFolder.css">
<%
Dim PATH
Dim Folder
Dim strSQL
Dim objDB
Dim varExists
Dim NewPath
Dim NewFolder

PATH = Request.QueryString("PATH")
Folder = Request.Form("Folder")

Set objDB = New clsDatabase


NewPATH = objDB.FormatString(PATH & "/" & Folder, 255)
NewFolder = objDB.FormatString(Folder, 63)


' See if path exists ...
strSQL = "SELECT 1 FROM Folders WHERE Path = " & NewPath

Call objDB.SetData(strSQL, varExists)

If IsArray(varExists) Then
	
	%>
	<BR>
	<B>Create New Folder - Error</B>
	<P>
		Cannot create <I><%=Folder%></I>: A sub folder with the name
		you specified already exists.
		<A href="CreateNewFolder.asp?<%=Request.QueryString%>">Specify a different folder name</A>.
	</P>
	
	<A href="SubFolders.asp?<%=Request.QueryString%>">Go back to list</A>
	<%

	Response.End
		
End If

strSQL = "INSERT INTO Folders(Folder, Path) VALUES(" & NewFolder & ", " & NewPath & ")"
Call objDB.ExecuteSQL(strSQL)

Response.Redirect("SubFolders.asp?Path=" & Server.URLEncode(Path & "/" & Folder))
%>