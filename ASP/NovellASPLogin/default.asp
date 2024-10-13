<!--#INCLUDE FILE="template.asp"-->
<%
Response.Buffer = True
Response.Expires = 0

Dim sCookie_UserName
Dim sCookie_Password
Dim bRememberMe
Dim sCookie_Server
Dim bInSession

' Grab any old "Rember Me" data from cookies
sCookie_UserName = Request.Cookies("UserName")
sCookie_Password = Request.Cookies("Password")
sCookie_Server = Request.Cookies("Server")

If Not Request.Form = "" Then
	sCookie_Server = Request.Form("Server")
End If

If Not Session("UserName") = "" Then
	sCookie_UserName = Session("UserName")
	sCookie_UserName = Session("UserName")
	sCookie_UserName = Session("Server")
	bInSession = True
End If

' Determine if user information was rememberd
bRememberMe = Not(sCookie_UserName = "" Or sCookie_Password = "")

%>
<HTML>
	<HEAD>
		<TITLE>Novell Login</TITLE>
		<META HTTP-EQUIV="Expires" CONTENT="Thu, 01 Dec 1994 16:00:00 GMT">
		<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
		<SCRIPT language="vbScript" src="getUserName.vbs"></SCRIPT>
		<SCRIPT language="JavaScript" src="validation.js"></SCRIPT>
	</HEAD>
	<%Call TemplateHeader("login.asp")%>
		<INPUT type="hidden" name="InSession" id="InSession" value="<%=bInSession%>">
		<TABLE width="100%" cellpadding="0" cellspacing="0" border="0">
			<%If Request.QueryString("Login") = "Failed" Then%>
			<!-- Login Failed -->
			<TR>
				<TD valign="top" width="1%" nowrap><FONT face="Helvetica,Arial,Sans-Serif" size="2" color="#880000"><B>Login Failed:</B></FONT></TD>
				<TD width="1%" nowrap><FONT face="Helvetica,Arial,Sans-Serif" size="2">&nbsp;</FONT></TD>
				<TD valign="top"><FONT face="Helvetica,Arial,Sans-Serif" size="2">Please try again.  If you have forgotten your login information, please contact your network administrator.<BR><BR></FONT></TD>
			</TR>
			<%End If%>
			<!-- User Name -->
			<TR>
				<TD width="1%" nowrap><FONT face="Helvetica,Arial,Sans-Serif" size="2"><LABEL FOR="UserName" ACCESSKEY="U"><U>U</U>sername:</LABLE>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</FONT></TD>
				<TD rowspan="3" width="1%" nowrap><FONT face="Helvetica,Arial,Sans-Serif" size="2">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</FONT></TD>
				<TD><INPUT type="text" name="UserName" id="UserName" style="width:100%;" size="41" value="<%=sCookie_UserName%>" onFocus="document.NovellClient.UserName.select"></TD>
			</TR>
			<!-- Password -->
			<TR>
				<TD width="1%" nowrap><FONT face="Helvetica,Arial,Sans-Serif" size="2"><LABEL FOR="Password" ACCESSKEY="P"><U>P</U>assword:</LABLE></FONT></TD>
				<TD><INPUT type="password" name="Password" id="Password" style="width:100%;" size="41" value="<%=sCookie_Password%>" onFocus="document.NovellClient.Password.select"></TD>
			</TR>
			<!-- Server -->
			<TR>
				<TD width="1%" nowrap><FONT face="Helvetica,Arial,Sans-Serif" size="2"><LABEL FOR="Server" ACCESSKEY="P"><U>S</U>ite Location:</LABLE></FONT></TD>
				<TD>
					<SELECT name="Server" id="Server" style="width:100%;">
						<%
						Set loNWSess = Server.CreateObject("NWSessLib.NWSessCtrl.1")

						For Each loServer In loNWSess.ServerNames
							Response.Write("<OPTION")
							If sCookie_Server = loServer.FullName Then Response.Write(" selected")
							Response.Write(">" & loServer.FullName & "</OPTION>")
						Next

						Set loNWSess = Nothing
						%>
					</SELECT>
				</TD>
			</TR>
		</TABLE>
		<TABLE width="100%" cellpadding="0" cellspacing="10" border="0">
			<!-- Remember Me -->
			<TR>
				<TD align="right" colspan="2">
					<INPUT type="checkbox" name="RememberMe" id="RememberMe" value="True" onClick="RememberMeWarning()"<%If bRememberMe Then Response.Write(" checked")%>>
					<FONT face="Helvetica,Arial,Sans-Serif" size="2"><LABEL FOR="RememberMe" ACCESSKEY="R"><U>R</U>emember Me</LABLE></FONT>
					<BR>
				</TD>
			</TR>
			<!-- Buttons -->
			<TR>
				<TD align="right">
					<INPUT type="button" name="OK" id="OK" value="    OK    " style="width:75px;height:23px;font-size:8pt;" onClick="OK_onClick()">
				</TD>
				<TD align="right" width="1%">
					<INPUT type="button" name="Cancel" id="Cancel" value="  Cancel  " style="width:75px;height:23px;font-size:8pt;" onClick="Cancel_onClick()">
				</TD>
			</TR>
		</TABLE>
	<%Call TemplateFooter()%>
</HTML>