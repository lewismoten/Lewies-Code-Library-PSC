<!--#INCLUDE FILE="clsDatabase.asp"-->
<%
If Not Session("IsAdmin") Then Response.Redirect("admin.Login.asp")

Dim objDB
Dim strSQL
Dim strCategoryName

strCategoryName = Request.Form("CategoryName")

Set objDB = New clsDatabase

strCategoryName = objDB.FormatString(strCategoryName, -1)

strSQL = _
	"INSERT INTO " & _
		"[Categories] " & _
	"(" & _
		"[CategoryName]" & _
	") VALUES (" & _
		strCategoryName & _
	")"
		
Call objDB.ExecuteSQL(strSQL)

Set objDB = Nothing

Response.Redirect("admin.asp")
%>
