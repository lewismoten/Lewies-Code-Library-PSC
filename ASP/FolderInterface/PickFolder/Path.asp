<LINK rel="stylesheet" type="text/css" href="PickFolder.css">
<BODY bgColor="#CCCCCC">
<%
Dim PATH
Dim llngIndex
Dim llngMaxIndex
Dim lstrPath
Dim lstrFolder
Dim lstrFolderAry

PATH = Request.QueryString("Path")


Response.Write "<B>Path:</B> "

lstrFolderAry = Split(PATH, "/")
llngMaxIndex = UBound(lstrFolderAry, 1)
lstrPath = ""

Response.Write "<A href=""SubFolders.asp?Path="" target=""SubFolders"">"
Call Image("icnHomeFolder.gif", "Home Folder")
Response.Write "Home"
Response.Write "</A>"


For llngIndex = 1 To llngMaxIndex

		lstrPath = lstrPath & Server.URLEncode("/" & lstrFolderAry(llngIndex))
		lstrFolder = Server.HTMLEncode(lstrFolderAry(llngIndex))
			
		Response.Write " / "
		Response.Write "<A href=""SubFolders.asp?Path=" & lstrPath & """ target=""SubFolders"">"
		Call Image("icnFolder.gif", lstrFolder)
		Response.Write lstrFolder
		Response.Write "</A>"

Next

Sub Image(ByRef pstrPath, ByRef pstrAlt)
	Response.Write "<IMG src=""" & pstrPath & """ "
	Response.Write "width=""16"" height=""16"" "
	Response.Write "alt=""" & pstrAlt & """ "
	Response.Write "border=""0"" align=""absmiddle"" "
	Response.Write "vspace=""3"" hspace=""5"" "
	Response.Write ">"
End Sub
%>