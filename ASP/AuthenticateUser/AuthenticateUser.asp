<%
' FileName: AuthenticateUser.asp
'
' Proposed Location: http://www.YourDomain.com/Inc/
'
' Use: 
'
'	Insert the following code at the top of any page
'	that requires a user to login.
'
'	<!--#INCLUDE VIRTUAL="/Inc/Authenticateuser.asp"-->
'
' Description:
'
'	Authenticates a user to make sure if they have
'	previously logged into the site.
'
'		Requires Session("UserID") to be populated.
'		This usually represents the Users ID within
'		a data base. (Users.UserID)
'
'	If a user is not loged in, they are redirected to a
'	page to attempt a login.  This is useful when the
'	ability to "Auto-Login" has been enabled to use 
'	previously saved login information in the users 
'	cookies.
'
'	When a user is redirected to the login page, The URL
'	they were attempting to view is passed along in the 
'	Query String along with the reason why they need to 
'	login.
'
'	If the user was posting data to the protected page
'	(perhaps a session timed out), then the previous page
'	they were posting from is sent as the URL that the 
'	user is redirected to after they have successfully 
'	logged in.  This is done to help reduce errors when
'	visiting a page that expected posted form data.
'
' Warning:
'	This script makes use of the Response.Redirect routine
'	It is important that data is not sent to the browser
'	before this routine is called or an error may result.
'	You can either set Response.Buffer to True (the
'	default setting on Windows NT 2000), or reference
'	this file before any thing is sent to the browser.

' ---------------------------------------------------------

' authenticate user...
'	The process is stored within a sub routine so that
'	variable names have less of a chance conflicting
'	with those located on the original page that calls it.
Call AuthenticateUser()

Sub AuthenticateUser()

	Dim lsURL		' URL that user will be returned to
					' after they have logged into the
					' site.
	
	Dim lsRedirect	' Page that user is redirected to if
					' they must login first.

	Dim lsReason	' Reason why a user must login to
					' view the requested page

	' Setup the reason why a user must login
	lsReason = "You have requested a page that " & _
		"requires authentication."

'	If the UserID is stored within a cookie,
'	this can easily be changed to:
'	If Request.Cookies("UserID") = "" Then

	' If the user is not currently logged into the site
	If IsEmpty(Session("UserID")) Then

		' Setup url to redirect users to Login Page.	
		lsRedirect = "/Authentication/LoginNow.asp?"
		
		' Give reason why the user is presented with a
		' login page.
		lsRedirect = lsRedirect & "Reason=" & _
			Server.URLEncode(lsReason)

		' Compile a URL that the user is redirected to
		' after they have logged in successfully.
		
		' If form data was not posted to this page
		If Request.Form = "" Then
		
			' compile the URL of current page
			lsURL = Request.ServerVariables("SCRIPT_NAME")
			
			' if a query string was passed to current page
			If Not Request.QueryString = "" Then
				
				' append query string to URL
				' (Query String is still encoded)
				lsURL = lsURL & "?" & Request.QueryString
				
			End If ' Not Request.QueryString = ""
		
		' Else form data was posted to this page
		Else
			
			' Let user come back to page that they
			' submitted information from.
			' (helps prevent server errrors)
			lsURL = Request.ServerVariables("HTTP_REFERER")
			
		End If ' Request.Form = ""
		
		' Encode the URL to be passed to the Login Page
		lsURL = Server.URLEncode(lsURL)

		' If we have a URL to redirect the user to once
		' they login
		If Not lsURL = "" Then
			
			' Append the URL to the login redirection page
			lsRedirect = lsRedirect & "&URL=" & lsURL
		
		End If ' Not lsURL = ""
		
		' Redirect users to AutoLogin Page
		Response.Redirect(lsRedirect)

	End If ' IsEmpty(Session("UserID"))

End Sub ' AuthenticateUser()
%>
