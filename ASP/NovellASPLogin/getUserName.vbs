' Target is IE users (ActiveX objects and vbScript)

Function GetCookie(ByRef asName)

	Dim lsCookiesAry
	Dim lsCrumbsAry
	Dim lnIndex
  
	lsCookiesAry = Split(document.cookie, "; ")

  	If Not IsArray(lsCookiesAry) Then Exit Function
  	
	For lnIndex = 0 To UBound(lsCookiesAry, 1)
  		lsCrumbsAry = Split(lsCookiesAry(lnIndex), "=")
  		If IsArray(lsCrumbsAry) Then
  			If UBound(lsCrumbsAry, 1) = 1 Then
  				If LCase(asName) = LCase(lsCrumbsAry(0)) Then
  					GetCookie = lsCrumbsAry(1)
  					Exit Function
				End If
			End If
		End If
	Next

End Function

Sub window_onLoad()

	' Attempt to load the users UserName into the form.

	Dim loNWSess			' Novell Session Object
	Dim lsUserName			' The users UserName
	Dim lsCookieUserName		' UserName stored in cookie

	' if user is currently in session, don't auto populate.
	If document.NovellClient.InSession.value = "True" Then Exit Sub
	
	lsCookieUserName = GetCookie("UserName")

	' Prepare for errors
	'	user may not have activeXobject installed
	'	user may not be logged into novell session on there computer
	On Error Resume Next

	' Create reference to Novell Session
	Set loNWSess = CreateObject("NWSessLib.NWSessCtrl.1")

	' If error occured (while creating object)
	If Err Then
		' Object can't be created/missing on users computer

		' Clear Error
		Err.Clear

		' Release Novell Session Object
		Set loNWSess = Nothing

		' Exit Procedure
		Exit Sub

	End If ' Err

	' Attempt to retrieve users UserName
	lsUserName = loNWSess.LoginName("NDS:\\AL01")

	' If error occured
	If Err Then
		' User isn't logged into local machine

		' Clear Error
		Err.Clear

		' Release Novell Session Object
		Set loNWSess = Nothing

		' Exit Procedure
		Exit Sub

	End If ' Err

	' If UserName was found
	If Not lsUserName  = "" Then

		' Populate form with UserName
		document.NovellClient.UserName.value = lsUserName

		' If UserName is the same as the username found in users cookies
		' Else User had old cookies stored from previous user
		If Not LCase(lsUserName) = LCase(lsCookieUserName) Then

			' Populate form with Users real UserName
			document.NovellClient.UserName.value = lsUserName

			' clear the password
			document.NovellClient.Password.value = ""

			' make sure the rememberme box is unchecked
			document.NovellClient.RememberMe.checked = False

		End If ' Not LCase(lsUserName) = LCase(lsCookieUserName)

	End If ' Not lsUserName = ""

	' Release object from memory
	Set loNWSess = Nothing

End Sub
