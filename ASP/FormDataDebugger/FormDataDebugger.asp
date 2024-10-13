<%
    '**************************************
    ' for :Form Data Debugger
    '**************************************
    'Copyright (c) 2001, Lewis Moten. All rights reserved.
    '**************************************
    ' Name: Form Data Debugger
    ' Description:Reads posted form data and
    '     returns an HTML formatted display of fie
    '     ld names and the values assigned to them
    '     .
    ' By: Lewis Moten
    '
    '
    ' Inputs:None
    '
    ' Returns:HTML formatted Ordered List of
    '     field names and values.
    '
    'Assumes:None
    '
    'Side Effects:None
    'This code is copyrighted and has limite
    '     d warranties.
    '**************************************
    
    Response.Write FormData()
    Function FormData()
    	Dim llngMaxFieldIndex
    	Dim llngFieldIndex
    	Dim llngMaxValueIndex
    	Dim llngValueIndex
    	Dim lstrDebug
    	' Count Form
    	llngMaxFieldIndex = Request.Form.Count
    	
    	' Let user know if Form do not exist
    	If llngMaxFieldIndex = 0 Then
    		FormData = "Form data is empty."
    		Exit Function
    	End If
    	
    	' Begin building a list of all Form
    	lstrDebug = "<OL>"
    	
    	' Loop through each Form
    	For llngFieldIndex = 1 To llngMaxFieldIndex
    		lstrDebug = lstrDebug & "<LI>" & Server.HTMLEncode(Request.Form.Key(llngFieldIndex))
    		
    		' Count the values
    		llngMaxValueIndex = Request.Form(llngFieldIndex).Count
    		
    		' If the Field doesn't have multiple values ...
    		If llngMaxValueIndex = 1 Then
    			lstrDebug = lstrDebug & " = "
    			lstrDebug = lstrDebug & Server.HTMLEncode(Request.Form.Item(llngFieldIndex))
    		' Else loop through each value
    		Else
    			lstrDebug = lstrDebug & "<OL>"
    			For llngValueIndex = 1 to llngMaxValueIndex
    				lstrDebug = lstrDebug & "<LI>"
    				lstrDebug = lstrDebug & Server.HTMLEncode(Request.Form(llngFieldIndex)(llngValueIndex))
    				lstrDebug = lstrDebug & "</LI>"
    			Next
    			lstrDebug = lstrDebug & "</OL>"
    		End If
    		lstrDebug = lstrDebug & "</LI>"
    	Next
    	lstrDebug = lstrDebug & "</OL>"
    	' Return the data
    	FormData = lstrDebug
    	
    End Function

%>