<!--#INCLUDE FILE="clsDatabase.asp"-->
<LINK rel="stylesheet" type="text/css" href="PickFolder.css">
<%

Dim PATH
Dim objDB

PATH = Request.QueryString("Path")


Set objDB = New clsDatabase

Call CheckExistance()
Call WriteUpOneLevel()
Call WriteSubFolders()
Call WriteCreateFolder()

Set objDB = Nothing

Sub CheckExistance()
	Dim lstrSQL
	Dim lvarExists
	lstrSQL = "SELECT 1 FROM Folders WHERE Path = '" & Replace(Path, "'", "''") & "'"
	Call objDB.SetData(lstrSQL, lvarExists)
	If Not IsArray(lvarExists) Then
		Response.Write "<BR>"
		Response.Write "<B>Folder Error</B>"
		Response.Write "<P>"
		Response.Write "The folder that you requested does not exist."
		Response.Write "</P>"
		
		Response.Write Replace(Path, "/", " / ")
		
		Response.End
	End If
End Sub
Sub WriteUpOneLevel()

	Dim lstrUpOneLevel
	
	If PATH = "" Then Exit Sub
	
	lstrUpOneLevel = Mid(PATH, 1, InStrRev(PATH, "/") - 1)
	lstrUpOneLevel = Server.URLEncode(lstrUpOneLevel)
	
	Response.Write "<DIV>"
	Response.Write "<A href=""SubFolders.asp?Path=" & lstrUpOneLevel & """>"
	Call Image("icnUpOneLevel.gif", "Up One Level")
	Response.Write "Up One Level"
	Response.Write "</A>"
	Response.Write "</DIV>"

End Sub

Sub WriteSubFolders()

	Dim lstrSQL
	Dim lvarFolderAry
	Dim llngIndex
	Dim llngMaxIndex
	Dim lstrPath
	Dim lstrFolder
	
	lstrSQL = "SELECT Path, Folder FROM Folders WHERE Path LIKE '" & Replace(Path, "'", "''") & "/%' AND NOT Path LIKE '" & Replace(Path, "'", "''") & "/%/%'"

	Call objDB.SetData(lstrSQL, lvarFolderAry)

	If IsArray(lvarFolderAry) Then

		llngMaxIndex = UBound(lvarFolderAry, 2)

		For llngIndex = 0 To llngMaxIndex

			lstrPath = lvarFolderAry(0, llngIndex) & ""
			lstrFolder = lvarFolderAry(1, llngIndex) & ""
			
			lstrPath = Server.URLEncode(lstrPath)
			lstrFolder = Server.HTMLEncode(lstrFolder)
			
			Response.Write "<DIV>"
			Response.Write "<A href=""SubFolders.asp?Path=" & lstrPath & """>"
			Call Image("icnFolder.gif", lstrFolder)
			Response.Write lstrFolder
			Response.Write "</A>"
			Response.Write "</DIV>"

		Next

	End If

End Sub

Sub WriteCreateFolder()
	Response.Write "<DIV>"
	Response.Write "<A href=""CreateNewFolder.asp?Path=" & Server.URLEncode(PATH) & """>"
	Call Image("icnCreateNewFolder.gif", "Create New Folder")
	Response.Write "Create New Folder"
	Response.Write "</A>"
	Response.Write "</DIV>"
End Sub

Sub Image(ByRef pstrPath, ByRef pstrAlt)
	Response.Write "<IMG src=""" & pstrPath & """ "
	Response.Write "width=""16"" height=""16"" "
	Response.Write "alt=""" & pstrAlt & """ "
	Response.Write "border=""0"" align=""absmiddle"" "
	Response.Write "vspace=""3"" hspace=""5"" "
	Response.Write ">"
End Sub

Function ToJavaString(ByVal pstrString)
	pstrString = Replace(pstrString, "\", "\\")
	pstrString = Replace(pstrString, vbCr, "\n")
	pstrString = Replace(pstrString, vbLf, "\l")
	pstrString = Replace(pstrString, vbTab, "\t")
	ToJavaString = pstrString
End Function
%>
<SCRIPT>
	window.top.returnValue = '<%=ToJavaString(PATH)%>';
	window.top.FullPath.location = 'Path.asp?path=' + escape(window.top.returnValue)
</SCRIPT>
