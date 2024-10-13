<%@ Language=VBScript %>
<%
' FileName: LoginNow.asp
'
' Proposed Location: http://www.YourDomain.com/Authentication/
'
' Use: 
'
'	If a page requires a user to be authenticated, Redirect 
'	anonymouse users to this page.
'
'	If IsEmpty(Session("UserID")) Then
'		Response.Redirect("/Authentication/LoginNow.asp")
'	End If
'
' About:
'	
'	This script does not "run out of the box" primarily because
'	the database connection will not be setup correctly for you.
'	It is only the process of logging in users and not the login form
'	itself.  Please keep a lookout for an upcomming article
'	that explains the full process of logging users into a website
'	against a database.  The main intent of this script is to
'	get you started with the most basic login processes that
'	most websites have standardized on.
'
' Description:
'
'	Attempts to log a user into the web site through either
'	cookies or posted form data by validating UserName and
'	password against a database.
'
'	Inputs:
'
'		URL: Original URL that user attempted to view
'
'		Username: Username to login to website
'
'		Password: Password to login to website
'
'		Reason: Reason why user is being requested to login
'
'		RememberMe: Determines weather or not to save login
'					data into cookies.
'
'
' Warning:
'	This script accesses a database.  You will need to change
'	the default settings with those that you would normally use
'	to login to your database with.  (DSN, UserID, Password).
'	Take care not confuse the database connection UserID and
'	Password with those that the user passes to login to the
'	website.  The script is dependent upon a table with the
'	following properties.
'
'	Table: Users
'	Fields:
'		UserID		Int/Number
'		Username	Char(?)/VarChar(?)/Text ("string")
'		Password	Char(?)/VarChar(?)/Text ("string")
'
'	If your database is setup differently, then you will need to
'	change the SQL that is passed to the connection object.
'
'	This script DOES NOT come with a database.  It is assumed
'	that you have a database, or the tools to create one.
'
'	The UserID is stored within a session variable Session("UserID")
'	so that protected pages can verify that the user is logged in.
'	If your site requires different methods to prove authentication,
'	then please make those changes below.
'
'	This script does not come with the Login Form that posts the
'	login data to this page.  This is mearly a process wich handles
'	that data.  This page expects the following posted form data.
'
'	Username (text)
'		User name to login with
'
'	Password (password)
'		Password to login with
'
'	RememberMe (checkbox value=True)
'		Determines if we are to store Login data in cookies.
'	URL (hidden)
'		original URL that user was attempting to visit before
'		being directed to this login page.
'
'	It is also assumed that your login page is within the same
'	directory and called "Login.asp".  If it is not, then you
'	will need to change the redirection scripts provided in this
'	script.

Option Explicit

Dim sUsername			' Username used to login to site
Dim sPassword			' Password used to login to site
Dim sURL				' Original URL user was trying to view
Dim bSuccess			' Determines if the user successfully loged into web site.

' High Level Flow of Process ...

' Grab Login Information from either:
'	(a) Posted Form Data
'	(b) Cookie Data
Call GrabLoginData()

' Verify that Login information was recieved, or
' Redirect user to Login form.
Call VerifyLoginData()

' Attempt to log the user into the website
Call LoginUser()

' If the user did not successfully login then redirect them to the login page
If Not bSuccess Then Call RedirectToLogin()

' Setup Cookie Login information for future "Auto-Login" attempts.
Call SetupCookies()

' Redirect the user to either
'	(a) the original protected page they tried to access
'	(b) the home page or there personal profile / settings
Call RedirectToProtected()


' ---------------------------------------------------------
Sub GrabLoginData()

	' If form data was not posted
	If Request.Form = "" Then
		
		' The user was probably redirected to this page for
		' an auto-login attempt.
		
		' Grab login data from cookies.
		sUserName = Request.Cookies("UserName")
		sPassword = Request.Cookies("Password")
		
		' Grab the protected URL that the user was attempting to view
		sURL = Request.QueryString("URL")

	' Else form data was posted
	Else

		' The user is comming from a Login Page

		' Grab posted login data
		sUsername	= Request.Form("Username")
		sPassword	= Request.Form("Password")

		' Grab the protected URL that the user was originally attempting to view
		sURL = Request.Form("URL")

	End If ' Request.Form = ""

End Sub ' GrabLoginData()
' ---------------------------------------------------------
Sub VerifyLoginData()

	Dim lsRedirect		' Login Page to redirect user to if
						' Login Data was not received.
	
	Dim lsReason		' Reason why user was not loged in
						' or why user must login

	' If the username or password was not submitted / empty
	If sUserName = "" Or sPassword = "" Then
		
		' Build the URL to redirect the user to
		lsRedirect = "Login.asp?URL=" & Server.URLEncode(sURL)
		
		' If user was login in through cookies
		If Request.Form = "" Then
			
			' Grab any reasons in the query string why the user is
			' requested to login.
			lsReason = Request.QueryString("Reason")
			
			' If a reason isn't given
			If lsReason = "" Then
			
				' Give the user a reason why they must login.			
				lsReason = "Please submit your Username and Password in order to login to the site."

			End If ' lsReason = ""
		
		' The user was logging in through a login form.
		Else
			
			' Notify the user that they must give a user name and a password in order to login.
			lsReason = "A Username and Password combination is required in order to login to the site."

		End If ' Request.Form = ""
		
		' Add the reason to the redirection URL
		lsRedirect = lsRedirect & "&Reason=" & Server.URLEncode(lsReason)
		
		' Redirect the user to the login page.
		Response.Redirect(lsRedirect)
		
	End If ' sUserName = "" Or sPassword = ""

End Sub ' VerifyLoginData()
' ---------------------------------------------------------
Sub LoginUser()
	
	Dim lsUsername	' SQL Formatted Username
	Dim lsPassword	' SQL Formatted Password
	Dim loConn		' Database Connection
	Dim loRs		' Database Recordset
	Dim lsSQL		' Structured Query Language
	
	' Format Username/Password into SQL strings
	lsUserName = "'" & Replace(sUserName, "'", "''") & "'"
	lsPassword = "'" & Replace(sPassword, "'", "''") & "'"

	' Build SQL statement to retrieve the Users ID	
	
	' WARNING:
	'	This is a sample of how to validate if a user exists in a database.
	'	If you use a different name for your table and/or fields, then
	'	you will need to change the following SQL statement to reflect
	'	those differences.

	lsSQL = _
		"SELECT " & _
			"UserID " & _
		"FROM " & _
			"Users " & _
		"WHERE " & _
			"Username = " & lsUserName & " " & _
			"AND Password = " & lsPassword
	
	' Open a database connection

	' WARNING:
	'	The connection Open method needs to be changed so that it will
	'	Connect to your own database.  This is just a sample.

	Set loConn = Server.CreateObject("ADODB.Connection")
	loConn.Open "MyDatabase", "Connection_Username", "Connection_Password"
	
	' Retrieve the Users ID
	Set loRs = Server.CreateObject("ADODB.Recordset")
	Set loRs = loConn.Execute(lsSQL)
	
	' If the username and password matched records in the database
	If Not loRs.EOF Then
		
		' Raise the success flag
		bSuccess = True
		
		' Setup the users session so that we can tell that the user
		' has logged into the website.
		Session("UserID") = loRs.Fields("UserID").Value

	End If ' Not loRs.EOF

	' Garbage Cleanup
	'	(a) Close connections to database
	'	(b) Release objects from memory

	loRs.Close
	Set loRs = Nothing
	
	loConn.Close
	Set loConn = Nothing
		
End Sub ' Login User
' ---------------------------------------------------------
Sub RedirectToLogin()

		
	Dim lsRedirect		' Login Page to redirect user to if
						' Login Data was not received.
	
	Dim lsReason		' Reason why user was not loged in
						' or why user must login

	' If an Auto-Login attempt was made (cookies)
	If Request.Form = "" Then
			
		' Grab any reasons in the query string why the user is
		' requested to login.
		lsReason = Request.QueryString("Reason")
			
		' If a reason isn't given
		If lsReason = "" Then
			
			' Give the user a reason why they must login.			
			lsReason = "Please submit your Username and Password in order to login to the site."

		End If ' lsReason = ""

	' Else user was loging in through a form
	Else
			
		' Grab reason from database results
		lsReason = vLoginAry(1, nLoginIndex)
		
	End If ' Request.Form = ""
		
		 
	' Build a URL to redirect the user to the Login Form.
	lsRedirect = "Login.asp?"
		
	' Append the Original URL that the user requested
	lsRedirect = lsRedirect & "URL=" & Server.URLEncode(sURL)

	' Append the reason why the user (must login/Failed)
	lsRedirect = lsRedirect & "&Reason=" & Server.URLEncode(lsReason)

	' Redirect the user to the login form.
	Response.Redirect(lsRedirect)
		

End Sub ' RedirectToLogin()	
' ---------------------------------------------------------
Sub SetupCookies()

	' If the user was able to login through cookies
	If Request.Form = "" Then

		' Refresh Expiration on cookie login data
		Response.Cookies("UserName").Expires = DateAdd("d", 30, Now())
		Response.Cookies("Password").Expires = DateAdd("d", 30, Now())

	' Else the user loged in through posted form data
	Else

		' If the user had asked us to remember them
		If Request.Form("RememberMe") = "True" Then
		
			' If the Username has prevously changed (initially blank)
			If Not Request.Cookies("UserName") = sUserName Then
				
				' Setup new user name
				Response.Cookies("UserName") = sUserName
				
			End If ' Not Request.Cookies("UserName") = sUserName
			
			' If the Password has prevously changed (initially blank)
			If Not Request.Cookies("Password") = sPassword Then
				
				' Setup new Password
				Response.Cookies("Password") = sPassword
				
			End If ' Not Request.Cookies("Password") = sPassword

			' Refresh Cookie Expiration to expire in 1 month from now
			Response.Cookies("UserName").Expires = DateAdd("m", 1, Now())
			Response.Cookies("Password").Expires = DateAdd("m", 1, Now())
		
		' Else the user does not want to be remembered
		Else
			
			' If the user name exists
			If Not Request.Cookies("UserName") = "" Then

				' Erase Username
				Response.Cookies("UserName") = ""
				Response.Cookies("UserName").Expires = DateAdd("m", -3, Now())

			End If ' Not Request.Cookies("UserName") = ""
			
			' If the password exists
			If Not Request.Cookies("Password") = "" Then
				
				' Erase the password
				Response.Cookies("Password") = ""
				Response.Cookies("Password").Expires = DateAdd("m", -3, Now())
			
			End If ' Not Request.Cookies("Password") = ""
			
		End If ' Request.Form("RememberMe") = "True"
	
	End If ' Request.Form = ""
	
End Sub ' SetupCookies()
' ---------------------------------------------------------
Sub RedirectToProtected()

	' If the user didn't originally request a protected page
	If sURL = "" Then
		
		' Redirect them to the home page.
		'	May want to change this later to redirect user to
		'	personal settings/profile page
		Response.Redirect("/")
		
	' Else the user originally requested a protected page
	Else
		
		' Redirect to portected page
		Response.Redirect(sURL)
		
	End If ' sURL = ""

End Sub ' RedirectToProtected()
' ---------------------------------------------------------
%>