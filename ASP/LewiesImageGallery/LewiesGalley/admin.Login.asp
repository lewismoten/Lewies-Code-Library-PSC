<!--#INCLUDE FILE="clsDatabase.asp"-->
<CENTER><H3><A href="http://www.lewismoten.com/">Lewie's</A> Photo Gallery Administration</H3><H5>Login</H5></CENTER>

<%
Dim strLoginName
Dim strLoginPassword
Dim strSQL
Dim objDB
Dim varUserAry

If Request.Form = "" Then
	Call LoginForm()
Else
	strLoginName = Request.Form("LoginName")
	strLoginPassword = Request.Form("LoginPassword")
	
	If strLoginName = "" Or strLoginPassword = "" Then
		Call LoginForm()
	Else
		Set objDB = New clsDatabase
		
		Response.Write objDB.Errors
		
		strLoginName = objDB.FormatString(strLoginName, -1)
		strLoginPassword = objDB.FormatString(strLoginPassword, -1)
		
		strSQL = _
			"SELECT " & _
				"[UserID] " & _
			"FROM " & _
				"[Users] " & _
			"WHERE " & _
				"[LoginName] = " & strLoginName & " " & _
				"AND [LoginPassword] = " & strLoginPassword
		
		Call objDB.SetData(strSQL, varUserAry)
		
		Set objDB = Nothing
		
		If IsArray(varUserAry) Then
			Session("IsAdmin") = True
			Response.Redirect("admin.asp")
		Else
			Call LoginForm()
		End If
	End If
End If
%>

<%Sub LoginForm()%>
	<FORM method="post">
		<TABLE>
			<TR>
				<TD>Login Name</TD>
				<TD><INPUT name="LoginName"></TD>
			</TR>
			<TR>
				<TD>Password</TD>
				<TD><INPUT name="LoginPassword"></TD>
			</TR>
			<TR>
				<TD>&nbsp;</TD>
				<TD><INPUT type="submit" value="Login"></TD>
			</TR>
		</TABLE>
	</FORM>
<%End Sub%>