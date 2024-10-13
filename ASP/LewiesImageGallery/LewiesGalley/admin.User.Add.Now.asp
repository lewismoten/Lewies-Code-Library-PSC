<!--#INCLUDE FILE="clsDatabase.asp"-->
<%
If Not Session("IsAdmin") Then Response.Redirect("admin.Login.asp")

Dim objDB
Dim strSQL
Dim strLoginName
Dim strLoginPassword

strLoginName = Request.Form("LoginName")
strLoginPassword = Request.Form("LoginPassword")

Set objDB = New clsDatabase

strLoginName = objDB.FormatString(strLoginName, -1)
strLoginPassword = objDB.FormatString(strLoginPassword, -1)

strSQL = _
	"INSERT INTO " & _
		"[Users] " & _
	"(" & _
		"[LoginName], " & _
		"[LoginPassword]" & _
	") VALUES (" & _
		strLoginName & ", " & _
		strLoginPassword & _
	")"
		
Call objDB.ExecuteSQL(strSQL)

Set objDB = Nothing

Response.Redirect("admin.asp")
%>
