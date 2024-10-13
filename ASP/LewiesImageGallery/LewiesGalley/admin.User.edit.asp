<!--#INCLUDE FILE="clsDatabase.asp"-->
<%If Not Session("IsAdmin") Then Response.Redirect("admin.Login.asp")%>

<%
Dim objDB
Dim strSQL
Dim lngUserID
Dim varUserAry
Dim strLoginName
Dim strLoginPassword

lngUserID = Request.Form("UserID")

Set objDB = New clsDatabase

' Grab Users
strSQL = _
	"SELECT " & _
		"[LoginName], " & _
		"[LoginPassword] " & _
	"FROM " & _
		"[Users] " & _
	"WHERE " & _
		"[UserID] = " & lngUserID
		
Call objDB.SetData(strSQL, varUserAry)

If IsArray(varUserAry) Then
	strLoginName = varUserAry(0, 0)
	strLoginPassword = varUserAry(1, 0)
End If

Set objDB = Nothing
%>
<CENTER><H3><A href="http://www.lewismoten.com/">Lewie's</A> Photo Gallery Administration</H3><H5><A href="admin.asp">Home</A>: Edit User</H5></CENTER>

<FORM method="post" action="admin.User.Edit.Now.asp">
	<INPUT type="hidden" name="UserID" value="<%=lngUserID%>">
	<TABLE>
		<TR>
			<TD>Login Name</TD>
			<TD><INPUT name="LoginName" value="<%=Server.HTMLEncode(strLoginName)%>"></TD>
		</TR>
		<TR>
			<TD>Password</TD>
			<TD><INPUT name="LoginPassword" value="<%=Server.HTMLEncode(strLoginPassword)%>"></TD>
		</TR>
		<TR>
			<TD>&nbsp;</TD>
			<TD><INPUT type="Submit" value="Save"></TD>
		</TR>
	</TABLE>
</FORM>
<HR>
Press this button to delete this user.  This can not be undone.
<FORM method="post" action="admin.User.Delete.Now.asp">
	<INPUT type="hidden" name="UserID" value="<%=lngUserID%>">
	<INPUT type="Submit" value="Delete" id=Submit1 name=Submit1>
</FORM>