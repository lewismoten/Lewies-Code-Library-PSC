<!--#INCLUDE FILE="clsDatabase.asp"-->
<%
If Not Session("IsAdmin") Then Response.Redirect("admin.Login.asp")

Dim objDB
Dim strSQL
Dim lngUserID
Dim varUserAry
Dim strLoginName
Dim strLoginPassword

lngUserID = Request.Form("UserID")
strLoginName = Request.Form("LoginName")
strLoginPassword = Request.Form("LoginPassword")

Set objDB = New clsDatabase

strLoginName = objDB.FormatString(strLoginName, -1)
strLoginPassword = objDB.FormatString(strLoginPassword, -1)

strSQL = _
	"UPDATE " & _
		"[Users] " & _
	"SET " & _
		"[LoginName] = " & strLoginName & ", " & _
		"[LoginPassword] = " & strLoginPassword & " " & _
	"WHERE " & _
		"[UserID] = " & lngUserID
		
Call objDB.ExecuteSQL(strSQL)

Set objDB = Nothing

Response.Redirect("admin.asp")
%>
