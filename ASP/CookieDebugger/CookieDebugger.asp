<%
    '**************************************
    ' for :Cookie Debugger
    '**************************************
    'Copyright (c) 2001, Lewis Edward Moten III. All rights reserved.
    '**************************************
    ' Name: Cookie Debugger
    ' Description:Creates a list of all cook
    '     ies and there crumbs along with the valu
    '     es assigned to each one.
    ' By: Lewis Moten
    '
    '
    ' Inputs:None
    '
    ' Returns:Returns an orderd list of name
    '     s and values of cookies and crumbs.
    '
    'Assumes:None
    '
    'Side Effects:None
    'This code is copyrighted and has limite
    '     d warranties.
    '**************************************
    
    Response.Write CookieData()
    Function CookieData()
    	Dim llngMaxCookieIndex
    	Dim llngCookieIndex
    	Dim llngMaxCrumbIndex
    	Dim llngCrumbIndex
    	Dim lstrDebug
    	' Count Cookies
    	llngMaxCookieIndex = Request.Cookies.Count
    	
    	' Let user know if cookies do not exist
    	If llngMaxCookieIndex = 0 Then
    		CookieData = "cookie data is empty."
    		Exit Function
    	End If
    	
    	' Begin building a list of all cookies
    	lstrDebug = "<OL>"
    	
    	' Loop through each cookie
    	For llngCookieIndex = 1 To llngMaxCookieIndex
    		lstrDebug = lstrDebug & "<LI>" & Server.HTMLEncode(Request.Cookies.Key(llngCookieIndex))
    		
    		' Count the crumbs
    		llngMaxCrumbIndex = Request.Cookies(llngCookieIndex).Count
    		
    		' If the cookie doesn't have crumbs ...
    		If llngMaxCrumbIndex = 0 Then
    			lstrDebug = lstrDebug & " = "
    			lstrDebug = lstrDebug & Server.HTMLEncode(Request.Cookies.Item(llngCookieIndex))
    		' Else loop through each crumb
    		Else
    			lstrDebug = lstrDebug & "<OL>"
    			For llngCrumbIndex = 1 to llngMaxCrumbIndex
    				lstrDebug = lstrDebug & "<LI>"
    				lstrDebug = lstrDebug & Server.HTMLEncode(Request.Cookies(llngCookieIndex).Key(llngCrumbIndex))
    				lstrDebug = lstrDebug & " = "
    				lstrDebug = lstrDebug & Server.HTMLEncode(Request.Cookies(llngCookieIndex)(llngCrumbIndex))
    				lstrDebug = lstrDebug & "</LI>"
    			Next
    			lstrDebug = lstrDebug & "</OL>"
    		End If
    		lstrDebug = lstrDebug & "</LI>"
    	Next
    	lstrDebug = lstrDebug & "</OL>"
    	' Return the data
    	CookieData = lstrDebug
    	
    End Function

%>