<%
Function BreadCrumbs()
	Dim lstrPathAry
	Dim llngMaxIndex
	Dim lstrPath
	Dim lstrHTML

	lstrPathAry = Split(Request.ServerVariables("PATH_INFO"), "/")
	
	llngMaxIndex = UBound(lstrPathAry) - 1
	
	lstrHTML = "<B>You are here:</B> "

	lstrPath = "/"
	lstrHTML = lstrHTML & "<A href=""" & lstrPath & """>Home</A> "
	
	For llngIndex = 1 To llngMaxIndex
		lstrPath = lstrPath & Server.URLPathEncode(lstrPathAry(llngIndex)) & "/"
		lstrHTML = lstrHTML & "&gt; <A href=""" & lstrPath & """>" & Server.HTMLEncode(lstrPathAry(llngIndex)) & "</A>"
	Next
	
	BreadCrumbs = lstrHTML
	
End Function
%>