<!--#INCLUDE FILE="clsDatabase.asp"-->
<%
If Not Session("IsAdmin") Then Response.Redirect("admin.Login.asp")

Dim objDB
Dim strSQL
Dim lngCategoryID
Dim varCategoryAry
Dim strCategoryName

lngCategoryID = Request.Form("CategoryID")
strCategoryName = Request.Form("CategoryName")

Set objDB = New clsDatabase

strCategoryName = objDB.FormatString(strCategoryName, -1)

strSQL = _
	"UPDATE " & _
		"[Categories] " & _
	"SET " & _
		"[CategoryName] = " & strCategoryName & " " & _
	"WHERE " & _
		"[CategoryID] = " & lngCategoryID
		
Call objDB.ExecuteSQL(strSQL)

Set objDB = Nothing

Response.Redirect("admin.asp")
%>
