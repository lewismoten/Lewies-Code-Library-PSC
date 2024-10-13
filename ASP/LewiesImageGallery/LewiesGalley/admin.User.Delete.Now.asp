<!--#INCLUDE FILE="clsDatabase.asp"-->
<%
If Not Session("IsAdmin") Then Response.Redirect("admin.Login.asp")

Dim objDB
Dim strSQL
Dim lngUserID
Dim varUserAry

lngUserID = Request.Form("UserID")

Set objDB = New clsDatabase

strSQL = _
	"SELECT " & _
		"COUNT([UserID]) " & _
	"FROM " & _
		"[Users]"

Call objDB.SetData(strSQL, varUserAry)

If IsArray(varUserAry) Then

	If varUserAry(0, 0) > 1 Then
		strSQL = _
			"DELETE FROM " & _
				"[Users] " & _
			"WHERE " & _
				"[UserID] = " & lngUserID
				
		Call objDB.ExecuteSQL(strSQL)
	End If
End If

Set objDB = Nothing

Response.Redirect("admin.asp")
%>
