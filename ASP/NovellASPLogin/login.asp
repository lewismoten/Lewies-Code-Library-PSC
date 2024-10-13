<%

' Prepare for redirects after text may have been written.
Response.Buffer = True

Dim sHTTP_REFERER ' Original (protected) URL that user had attempted to view.
Const sLOGIN_PATH = "/login/" ' path that login scripts reside in.

' Grab Referer from posted data
sHTTP_REFERER = Request.Form("HTTP_REFERER")

' If referer is empty, then grab it from the querystring
If sHTTP_REFERER = "" Then sHTTP_REFERER = Request.QueryString("HTTP_REFERER")

' If referer is still empty, then assume that the referer is the current page we are viewing
If sHTTP_REFERER = "" Then sHTTP_REFERER = Request.ServerVariables("HTTP_REFERER")

' If the referer is within the login directory...
If Not InStr(1, LCase(sHTTP_REFERER), sLOGIN_PATH) = 0 Then
	' set the referer to nothing
	sHTTP_REFERER = ""
End If ' Not InStr(1, LCase(sHTTP_REFERER), "/login/") = 0

%>
<!--#INCLUDE FILE="template.asp"-->
<HTML>
	<HEAD>
		<TITLE>Novell Login</TITLE>
		<META HTTP-EQUIV="Expires" CONTENT="Thu, 01 Dec 1994 16:00:00 GMT">
		<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
	</HEAD>
	<%

	Call TemplateHeader("")

	If Request.Form = "" Then
		Response.Redirect("default.asp?HTTP_REFERER=" & Server.URLEncode(sHTTP_REFERER))
	Else
		If ValidateUser() Then

			Response.Write("<P>Login Successful</P>")
			
			If Not sHTTP_REFERER = "" Then
				' Response.Redirect(sHTTP_REFERER)
				%>
				<P>Click the following link to return to the original page that you requested.</P>
				<A href="<%=sHTTP_REFERER%>"><%=sHTTP_REFERER%></A>
				<%
			End If
		
		' Else couldn't validate user
		Else
		
			Response.Write "Not validated<BR>"
			
'			Response.Redirect("default.asp?Login=Failed&HTTP_REFERER=" & Server.URLEncode(sHTTP_REFERER))
			
		End If ' ValidateUser()
	End If

	Call TemplateFooter()
	%>
</HTML>
<%
'-------------------------------------------------------------------------------
Function ValidateUser()

	Dim lsUserName				' Users UserName
	Dim lsPassword				' Users Password
	Dim loNwSess				' Novell Session object
	Dim lbRememberMe			' Does the user want us to remember our login information
	Dim lsServer				' Server that user is loging into

	Const lbSHOW_UI = False		' Show user interface (Can't because this is a server execution)

	' Grab form data
	lsUserName = Request.Form("UserName")
	lsPassword = Request.Form("Password")
	lsServer = Request.Form("Server")
	lbRememberMe = (Request.Form("RememberMe") = "True")

'	On Error Resume Next

	Set loNwSess = Server.CreateObject("NWSessLib.NWSessCtrl.1")
	
	If Err Then

		' Most likely to be problems because something wasn't installed.  Lets give user/webmaster
		' a detailed message about what they may need.
		
		%>
		<FONT color="#880000"><B>Error</B></FONT><BR>
		<P>
			<%=Err.Description%>
		</P>
		<HR>
		<P>
			The object that is required to validate you against the Novell servers can not be instantiated.
			This is because it is missing, or the web servers security settings does not give it permission
			to create the object.
		</P>
		<P>
			Please contact the web administrator to fix this problem.  If you are the administrator, and you
			have found that the controls are not yet installed, then you may visit the Novell Site to download
			the required files.
		</P>
		<P>
			The <A href="http://developer.novell.com">Developers Novell Site</A> contains a page of information
			and software with <A href="http://developer.novell.com/ndk/ocx.htm">Novell Controls for ActiveX</A>
			in the Novell Developer Kit.
		</P>
		<BR>
		<%

		Err.Clear
		
		ValidateUser = False

		Exit Function

	End If ' Err
	
	' Initialize some default actions
	loNWSess.Bindery = True
	loNWSess.RunScripts = False
	loNWSess.DisplayResults = False
	
	' log out any current user
	Call loNWSess.Logout(lsServer)
	
	' Attempt to log the user into the novell server
	ValidateUser = loNWSess.Login(lsServer, lsUserName, lsPassword, lbSHOW_UI)

	' If the user successfully loged in to Novell
	If ValidateUser Then

		' Save username/password within users session
		Session("UserName") = lsUserName
		Session("Password") = lsPassword
		
		' Logout of Novell Session Appliction
		Call loNWSess.Logout(lsServer)
		
		' If the user wants us to remember there login information
		If lbRememberMe Then
		
			' Write login info to cookies
			
			' UserName
			Response.Cookies("UserName") = lsUserName
			Response.Cookies("UserName").Expires = DateAdd("d", Now(), 30)
			
			' Password
			Response.Cookies("Password") = lsPassword
			Response.Cookies("Password").Expires = DateAdd("d", Now(), 30)
			
		Else
			
			' Destroy any login cookies if they exist
			
			' Username
			If Not Request.Cookies("UserName") = "" Then
				Response.Cookies("UserName") = ""
				Response.Cookies("UserName").Expires = DateAdd("y", Now(), -1)
			End If ' Not Request.Cookies("UserName") = ""
			
			' Password
			If Not Request.Cookies("Password") = "" Then
				Response.Cookies("Password") = ""
				Response.Cookies("Password").Expires = DateAdd("y", Now(), -1)
			End If ' Not Request.Cookies("Password") = ""
			
		End If ' lbRememberMe

	Else
	
		' Destroy any login cookies if they exist
		
		' Username
		If Not Request.Cookies("UserName") = "" Then
			Response.Cookies("UserName") = ""
			Response.Cookies("UserName").Expires = DateAdd("y", Now(), -1)
		End If

		' Password
		If Not Request.Cookies("Password") = "" Then
			Response.Cookies("Password") = ""
			Response.Cookies("Password").Expires = DateAdd("y", Now(), -1)
		End If

		' Logout the user from the session
		Session.Abandon

	End If ' ValidateUser

	' Release objects from memory
	Set loNWSess = Nothing

End Function ' ValidateUser
'-------------------------------------------------------------------------------
%>