<!--#INCLUDE FILE="clsDatabase.asp"-->
<%
If Not Session("IsAdmin") Then Response.Redirect("admin.Login.asp")

Dim objDB
Dim strSQL
Dim lngImageID

lngImageID = Request.Form("ImageID")

Set objDB = New clsDatabase

Call objDB.ExecuteSQL(strSQL)

strSQL = _
	"DELETE FROM " & _
		"[Images] " & _
	"WHERE " & _
		"[ImageID] = " & lngImageID
				
Call objDB.ExecuteSQL(strSQL)

Set objDB = Nothing

Response.Redirect("admin.asp")
%>
