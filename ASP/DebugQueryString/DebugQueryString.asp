<%
    '**************************************
    ' for :Debug QueryString
    '**************************************
    'Copyright (c) 2001, Lewis Moten. All rights reserved.
    '**************************************
    ' Name: Debug QueryString
    ' Description:Just used for debugging qu
    '     erystring data. Creates an orderd list o
    '     f field names and the values assigned to
    '     each one.
    ' By: Lewis Moten
    '
    '
    ' Inputs:None
    '
    ' Returns:None
    '
    'Assumes:None
    '
    'Side Effects:None
    'This code is copyrighted and has limite
    '     d warranties.
    '**************************************
    
    Response.Write QueryStringData()
    Function QueryStringData()
    	Dim llngMaxFieldIndex
    	Dim llngFieldIndex
    	Dim llngMaxValueIndex
    	Dim llngValueIndex
    	Dim lstrDebug
    	' Count QueryString
    	llngMaxFieldIndex = Request.QueryString.Count
    	
    	' Let user know if QueryString do not exist
    	If llngMaxFieldIndex = 0 Then
    		QueryStringData = "QueryString data is empty."
    		Exit Function
    	End If
    	
    	' Begin building a list of all QueryString
    	lstrDebug = "<OL>"
    	
    	' Loop through each QueryString
    	For llngFieldIndex = 1 To llngMaxFieldIndex
    		lstrDebug = lstrDebug & "<LI>" & Server.HTMLEncode(Request.QueryString.Key(llngFieldIndex))
    		
    		' Count the values
    		llngMaxValueIndex = Request.QueryString(llngFieldIndex).Count
    		
    		' If the Field doesn't have multiple values ...
    		If llngMaxValueIndex = 1 Then
    			lstrDebug = lstrDebug & " = "
    			lstrDebug = lstrDebug & Server.HTMLEncode(Request.QueryString.Item(llngFieldIndex))
    		' Else loop through each value
    		Else
    			lstrDebug = lstrDebug & "<OL>"
    			For llngValueIndex = 1 to llngMaxValueIndex
    				lstrDebug = lstrDebug & "<LI>"
    				lstrDebug = lstrDebug & Server.HTMLEncode(Request.QueryString(llngFieldIndex)(llngValueIndex))
    				lstrDebug = lstrDebug & "</LI>"
    			Next
    			lstrDebug = lstrDebug & "</OL>"
    		End If
    		lstrDebug = lstrDebug & "</LI>"
    	Next
    	lstrDebug = lstrDebug & "</OL>"
    	' Return the data
    	QueryStringData = lstrDebug
    	
    End Function

%>