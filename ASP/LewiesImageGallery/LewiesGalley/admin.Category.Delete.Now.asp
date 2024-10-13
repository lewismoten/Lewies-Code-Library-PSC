<!--#INCLUDE FILE="clsDatabase.asp"-->
<%
If Not Session("IsAdmin") Then Response.Redirect("admin.Login.asp")

Dim objDB
Dim strSQL
Dim lngCategoryID

lngCategoryID = Request.Form("CategoryID")

Set objDB = New clsDatabase

strSQL = _
	"UPDATE " & _
		"[Images] " & _
	"SET " & _
		"[CategoryID] = NULL " & _
	"WHERE " & _
		"[CategoryID] = " & lngCategoryID
	

Call objDB.ExecuteSQL(strSQL)

strSQL = _
	"DELETE FROM " & _
		"[Categories] " & _
	"WHERE " & _
		"[CategoryID] = " & lngCategoryID
				
Call objDB.ExecuteSQL(strSQL)

Set objDB = Nothing

Response.Redirect("admin.asp")
%>
