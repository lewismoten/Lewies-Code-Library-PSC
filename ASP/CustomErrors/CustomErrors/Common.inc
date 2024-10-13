<%
Sub Authenticate(ByRef pObjSite)
	
	Dim lLngInstanceID
	Dim lStrMetabasePath
	Dim lBlnContinue
	Dim lBlnLoginFailure
	
	On Error Resume Next

	lLngInstanceID = Request.ServerVariables("INSTANCE_ID")
	
	' Programmers Notes ...
	'
	' Metabase Path											Key Type 
	' /LM/W3SVC												IIsWebService 
	' /LM/W3SVC/N											IIsWebServer 
	' /LM/W3SVC/N/ROOT										IIsWebVirtualDir 
	' /LM/W3SVC/N/ROOT/WebVirtualDir						IIsWebVirtualDir 
	' /LM/W3SVC/N/ROOT/WebVirtualDir/WebDirectory			IIsWebDirectory 
	' /LM/W3SVC/N/ROOT/WebVirtualDir/WebDirectory/WebFile	IIsWebFile 
	'
	' N = lLngInstanceID
	'
	'
	
	lStrMetabasePath = Request.ServerVariables("APPL_MD_PATH")
	lStrMetabasePath = Replace(lStrMetabasePath, "/LM/", "IIS://LOCALHOST/", 1, vbTextCompare)
	
	'
	'
	'

	Set pObjSite = GetObject(lStrMetabasePath)

	If Err = &H800401E4 Or Err = 70 Then
		Response.Status = "401 access denied"
		BlnContinue = False
		BlnLoginFailure = True
	Else
		If Err = 0 Then
			lBlnContinue = True
		Else
			lBlnContinue = False
			lBlnLoginFailure = False
		End If
	End If
		
	If lBlnLoginFailure Then
		Response.Write "Login Failure.<BR>"
		Response.End
	End If
	If Not lBlnContinue Then
		Response.Write "Can not continue.<BR>"
		Response.End
	End If

End Sub
%>